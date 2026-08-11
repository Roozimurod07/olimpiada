import asyncio
import logging
import sys
import random
import os
import re
import json
from datetime import datetime
from aiogram import Bot, Dispatcher, F, Router
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import Message, CallbackQuery, InlineKeyboardMarkup, InlineKeyboardButton, ReplyKeyboardMarkup, KeyboardButton, BufferedInputFile
from sqlalchemy import Column, Integer, String, Text, DateTime, Boolean, ForeignKey, Float, select, update, delete, func, text as sa_text, event
from sqlalchemy.ext.asyncio import create_async_engine, async_sessionmaker
from sqlalchemy.orm import declarative_base, relationship

import pypdf
import docx
import pandas as pd
import io

# PDF Sertifikat uchun kutubxona
from reportlab.lib.pagesizes import landscape, letter
from reportlab.pdfgen import canvas
from reportlab.lib import colors

# Google Sheets kutubxonalari
import gspread
from google.oauth2.service_account import Credentials

# --- RAILWAY VA GOOGLE SHEETS SOZLAMALARI ---
BOT_TOKEN = os.getenv("BOT_TOKEN")
GOOGLE_CREDS_JSON = os.getenv("GOOGLE_CREDS_JSON")
SHEET_NAME = "Olimpiada"

SUPER_ADMIN_IDS = [8317043750]

Base = declarative_base()

class Student(Base):
    __tablename__ = "students"
    id = Column(Integer, primary_key=True, autoincrement=True)
    student_id = Column(String(50), unique=True, index=True, nullable=False)
    first_name = Column(String(100), nullable=False)
    last_name = Column(String(100), nullable=False)
    school = Column(String(150), nullable=True)
    grade = Column(String(20), nullable=True)
    telegram_id = Column(Integer, unique=True, nullable=True)
    is_active = Column(Boolean, default=True)

    test_sessions = relationship("TestSession", back_populates="student", cascade="all, delete-orphan")
    appeals = relationship("Appeal", back_populates="student", cascade="all, delete-orphan")

class Admin(Base):
    __tablename__ = "admins"
    id = Column(Integer, primary_key=True, autoincrement=True)
    telegram_id = Column(Integer, unique=True, nullable=False)
    role = Column(String(50), default="moderator")

class Test(Base):
    __tablename__ = "tests"
    id = Column(Integer, primary_key=True, autoincrement=True)
    title = Column(String(200), nullable=False)
    subject = Column(String(100), nullable=False)
    grade_level = Column(String(20), nullable=False)
    max_attempts = Column(Integer, default=1)
    mode = Column(String(20), default="question_timer")
    duration_minutes = Column(Integer, default=30)
    duration_seconds_per_question = Column(Integer, default=15)
    start_time = Column(DateTime, nullable=True)
    end_time = Column(DateTime, nullable=True)
    is_active = Column(Boolean, default=True)
    is_finished = Column(Boolean, default=False) # Admin tomonidan test yakunlanganligi
    
    questions = relationship("Question", back_populates="test", cascade="all, delete-orphan")
    test_sessions = relationship("TestSession", back_populates="test", cascade="all, delete-orphan")

class Question(Base):
    __tablename__ = "questions"
    id = Column(Integer, primary_key=True, autoincrement=True)
    test_id = Column(Integer, ForeignKey("tests.id", ondelete="CASCADE"), nullable=False)
    question_text = Column(Text, nullable=False)
    photo_file_id = Column(String(200), nullable=True)
    option_a = Column(Text, nullable=False)
    option_b = Column(Text, nullable=False)
    option_c = Column(Text, nullable=True)
    option_d = Column(Text, nullable=True)
    correct_option = Column(String(5), nullable=False)
    points = Column(Float, default=1.0)
    
    test = relationship("Test", back_populates="questions")
    answers = relationship("Answer", back_populates="question", cascade="all, delete-orphan")

class TestSession(Base):
    __tablename__ = "test_sessions"
    id = Column(Integer, primary_key=True, autoincrement=True)
    student_id = Column(Integer, ForeignKey("students.id", ondelete="CASCADE"), nullable=False)
    test_id = Column(Integer, ForeignKey("tests.id", ondelete="CASCADE"), nullable=False)
    status = Column(String(30), default="IN_PROGRESS")
    started_at = Column(DateTime, default=datetime.utcnow)
    finished_at = Column(DateTime, nullable=True)
    score = Column(Float, default=0.0)
    score_percentage = Column(Float, default=0.0)
    correct_answers = Column(Integer, default=0)
    wrong_answers = Column(Integer, default=0)
    unanswered = Column(Integer, default=0)

    student = relationship("Student", back_populates="test_sessions")
    test = relationship("Test", back_populates="test_sessions")
    answers = relationship("Answer", back_populates="session", cascade="all, delete-orphan")

class Answer(Base):
    __tablename__ = "answers"
    id = Column(Integer, primary_key=True, autoincrement=True)
    session_id = Column(Integer, ForeignKey("test_sessions.id", ondelete="CASCADE"), nullable=False)
    question_id = Column(Integer, ForeignKey("questions.id", ondelete="CASCADE"), nullable=False)
    selected_option = Column(String(5), nullable=True)

    session = relationship("TestSession", back_populates="answers")
    question = relationship("Question", back_populates="answers")

class Appeal(Base):
    __tablename__ = "appeals"
    id = Column(Integer, primary_key=True, autoincrement=True)
    student_id = Column(Integer, ForeignKey("students.id", ondelete="CASCADE"), nullable=False)
    test_session_id = Column(Integer, ForeignKey("test_sessions.id", ondelete="CASCADE"), nullable=False)
    question_id = Column(Integer, ForeignKey("questions.id", ondelete="CASCADE"), nullable=False)
    message_text = Column(Text, nullable=False)
    status = Column(String(30), default="PENDING")
    
    student = relationship("Student", back_populates="appeals")

engine = create_async_engine("sqlite+aiosqlite:///professional_olimpiada.db", echo=False)

@event.listens_for(engine.sync_engine, "connect")
def set_sqlite_pragma(dbapi_connection, connection_record):
    cursor = dbapi_connection.cursor()
    cursor.execute("PRAGMA foreign_keys=ON")
    cursor.close()

async_session = async_sessionmaker(engine, expire_on_commit=False)

async def init_db():
    async with engine.begin() as conn:
        await conn.execute(sa_text("PRAGMA foreign_keys = ON;"))
        await conn.run_sync(Base.metadata.create_all)

async def is_admin(user_id: int) -> bool:
    if user_id in SUPER_ADMIN_IDS:
        return True
    async with async_session() as session:
        adm = (await session.execute(select(Admin).where(Admin.telegram_id == user_id))).scalar_one_or_none()
        return adm is not None

# --- GOOGLE SHEETS BILAN ISHLASH ---
def get_gspread_sheet():
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    if GOOGLE_CREDS_JSON:
        creds_dict = json.loads(GOOGLE_CREDS_JSON)
        creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    else:
        creds = Credentials.from_service_account_file('credentials.json', scopes=scope)
    client = gspread.authorize(creds)
    return client.open(SHEET_NAME).sheet1

def save_result_to_sheet(student_id, full_name, school, grade, test_title, subject, score, percentage, correct, wrong):
    try:
        sheet = get_gspread_sheet()
        row_data = [
            str(student_id), str(full_name), str(school), str(grade),
            str(subject), str(test_title), str(score), f"{percentage}%",
            str(correct), str(wrong), datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ]
        sheet.append_row(row_data)
    except Exception as e:
        print(f"❌ Google Sheets qo'shish xatosi: {e}")

def update_result_in_sheet(student_id, test_title, score, percentage, correct):
    try:
        sheet = get_gspread_sheet()
        records = sheet.get_all_records()
        for idx, row in enumerate(records, start=2):
            # ID va Test nomini to'g'ri solishtirish uchun stringga o'tkazamiz
            if str(row.get("ID")) == str(student_id) and str(row.get("Test")) == str(test_title):
                sheet.update_cell(idx, 7, str(score))          # Ball ustuni (7-ustun)
                sheet.update_cell(idx, 8, f"{percentage}%")    # Foiz ustuni (8-ustun)
                sheet.update_cell(idx, 9, str(correct))        # To'g'ri javoblar ustuni (9-ustun)
                break
    except Exception as e:
        print(f"❌ Google Sheets yangilash xatosi: {e}")

# --- SERTIFIKAT GENERATSIYA QILISH ---
def generate_certificate_pdf(student_name, test_title, subject, score_pct):
    pdf_buffer = io.BytesIO()
    c = canvas.Canvas(pdf_buffer, pagesize=landscape(letter))
    width, height = 792, 612
    
    c.setStrokeColor(colors.HexColor("#1A365D"))
    c.setLineWidth(5)
    c.rect(20, 20, width - 40, height - 40)
    
    c.setStrokeColor(colors.HexColor("#D69E2E"))
    c.setLineWidth(2)
    c.rect(28, 28, width - 56, height - 56)
    
    c.setFillColor(colors.HexColor("#1A365D"))
    c.setFont("Helvetica-Bold", 28)
    c.drawCentredString(width / 2, height - 100, "SERTIFIKAT")
    
    c.setFont("Helvetica", 14)
    c.setFillColor(colors.HexColor("#4A5568"))
    c.drawCentredString(width / 2, height - 140, "Ushbu sertifikat quyidagi o'quvchiga beriladi:")
    
    c.setFont("Helvetica-Bold", 32)
    c.setFillColor(colors.HexColor("#2B6CB0"))
    c.drawCentredString(width / 2, height - 200, student_name)
    
    c.setFont("Helvetica", 14)
    c.setFillColor(colors.HexColor("#4A5568"))
    c.drawCentredString(width / 2, height - 250, f"{subject} fani bo'yicha o'tkazilgan '{test_title}'")
    c.drawCentredString(width / 2, height - 275, f"olimpiadasida muvaffaqiyatli qatnashib, {score_pct}% natija ko'rsatgani uchun taqdirlanadi.")
    
    c.setFont("Helvetica-Oblique", 12)
    c.drawString(60, 80, f"Sana: {datetime.now().strftime('%Y-%m-%d')}")
    c.drawRightString(width - 60, 80, "Tizim rahbarligi: Professional Olimpiada")
    
    c.save()
    pdf_buffer.seek(0)
    return pdf_buffer.read()

class RegState(StatesGroup):
    waiting_for_id = State()

class AdminAddStudent(StatesGroup):
    waiting_for_data = State()
    waiting_for_excel = State()

class AdminAddTest(StatesGroup):
    waiting_for_title = State()
    waiting_for_subject = State()
    waiting_for_grade = State()
    waiting_for_mode = State()
    waiting_for_attempts = State()
    waiting_for_duration = State()
    waiting_for_questions = State()
    waiting_for_answers = State()

class AdminBroadcast(StatesGroup):
    waiting_for_message = State()

class AdminManageAdmins(StatesGroup):
    waiting_for_id = State()

class AppealState(StatesGroup):
    waiting_for_text = State()

class TestProcessState(StatesGroup):
    in_test = State()

router = Router()

def get_main_menu():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="📝 Testni boshlash"), KeyboardButton(text="👤 Profilim")],
            [KeyboardButton(text="📊 Mening urinishlarim"), KeyboardButton(text="⚖️ Apellyatsiya")],
            [KeyboardButton(text="🏆 Reyting"), KeyboardButton(text="ℹ️ Olimpiada haqida")]
        ],
        resize_keyboard=True
    )

def get_admin_menu():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="➕ ID qo'shish"), KeyboardButton(text="📂 Excel orqali ID'lar")],
            [KeyboardButton(text="📂 Test yuklash"), KeyboardButton(text="⚙️ Testlarni boshqarish")],
            [KeyboardButton(text="📊 Jonli statistika"), KeyboardButton(text="📥 Excel natijalar")],
            [KeyboardButton(text="⚖️ Apellyatsiyalar"), KeyboardButton(text="👥 Adminlar")],
            [KeyboardButton(text="🧹 Bazani tozalash"), KeyboardButton(text="📢 Xabar yuborish")],
            [KeyboardButton(text="⬅️ Bosh menyu")]
        ],
        resize_keyboard=True
    )

def get_finish_test_keyboard():
    return ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text="✅ Javobni yuklash / Testni saqlash")]],
        resize_keyboard=True
    )

@router.message(Command("start"))
async def cmd_start(message: Message, state: FSMContext):
    await state.clear()
    if await is_admin(message.from_user.id):
        await message.answer("🛠 <b>Xush kelibsiz, Admin!</b>", reply_markup=get_admin_menu())
        return

    async with async_session() as session:
        result = await session.execute(select(Student).where(Student.telegram_id == message.from_user.id))
        student = result.scalar_one_or_none()
        
        if student:
            if not student.is_active:
                await message.answer("❌ Sizning profilingiz administrator tomonidan bloklangan.")
                return
            await message.answer(f"Xush kelibsiz, <b>{student.first_name} {student.last_name}</b>!\nSinfingiz: <b>{student.grade or 'Nomaʼlum'}</b>", reply_markup=get_main_menu())
            return

    await state.set_state(RegState.waiting_for_id)
    await message.answer("🎓 <b>Olimpiada tizimiga xush kelibsiz!</b>\n\nIltimos, ID raqamingizni kiriting:")

@router.message(RegState.waiting_for_id)
async def process_student_id(message: Message, state: FSMContext):
    entered_id = message.text.strip()
    async with async_session() as session:
        result = await session.execute(select(Student).where(Student.student_id == entered_id))
        student = result.scalar_one_or_none()
        
        if not student:
            await message.answer("❌ ID raqami topilmadi. Qayta kiriting:")
            return
        if student.telegram_id is not None and student.telegram_id != message.from_user.id:
            await message.answer("⚠️ Bu ID boshqa akkauntga ulangan.")
            return
        
        await session.execute(update(Student).where(Student.id == student.id).values(telegram_id=message.from_user.id))
        await session.commit()
        
    await state.clear()
    await message.answer(f"✅ Muvaffaqiyatli ro'yxatdan o'tdingiz, {student.first_name}!\nSinfingiz: <b>{student.grade or '-'}</b>", reply_markup=get_main_menu())

@router.message(F.text == "👤 Profilim")
async def profile_handler(message: Message, state: FSMContext):
    if await state.get_state() == TestProcessState.in_test.state: return
    async with async_session() as session:
        student = (await session.execute(select(Student).where(Student.telegram_id == message.from_user.id))).scalar_one_or_none()
        if not student:
            await message.answer("Siz ro'yxatdan o'tmagansiz. /start ni bosing.")
            return
        await message.answer(f"👤 <b>Profil:</b>\n\nID: <code>{student.student_id}</code>\nIsm: {student.first_name} {student.last_name}\nMaktab: {student.school or '-'}\nSinf: {student.grade or '-'}")

@router.message(F.text == "📊 Mening urinishlarim")
async def my_attempts_handler(message: Message, state: FSMContext):
    if await state.get_state() == TestProcessState.in_test.state: return
    async with async_session() as session:
        student = (await session.execute(select(Student).where(Student.telegram_id == message.from_user.id))).scalar_one_or_none()
        if not student:
            await message.answer("Ro'yxatdan o'tmagansiz.")
            return
        
        sessions = (await session.execute(
            select(TestSession, Test)
            .join(Test, TestSession.test_id == Test.id)
            .where(TestSession.student_id == student.id, TestSession.status == "COMPLETED")
            .order_by(TestSession.finished_at.desc())
        )).all()
        
        if not sessions:
            await message.answer("⚠️ Siz hali birorta testni yakunlamagansiz.")
            return
            
        text = "📊 <b>Sizning ishlagan testlaringiz natijalari:</b>\n\n"
        for ts, t in sessions:
            date_str = ts.finished_at.strftime("%Y-%m-%d %H:%M") if ts.finished_at else ""
            status_text = "🟢 Test yakunlangan (Natijalar ochiq)" if t.is_finished else "🟡 Test hali davom etmoqda (Tahlil va apellyatsiya yopilgan)"
            text += f"📚 <b>{t.subject}</b> ({t.title})\n⭐ Ball: {ts.score} ({ts.score_percentage}%)\n📅 Sana: {date_str}\nStatus: {status_text}\n----------------------------------\n"
            
        await message.answer(text)

# --- O'QUVCHILAR UCHUN ALOHIDA APELLYATSIYA BO'LIMI ---
@router.message(F.text == "⚖️ Apellyatsiya")
async def student_appeal_menu(message: Message, state: FSMContext):
    if await state.get_state() == TestProcessState.in_test.state: return
    async with async_session() as session:
        student = (await session.execute(select(Student).where(Student.telegram_id == message.from_user.id))).scalar_one_or_none()
        if not student:
            await message.answer("Ro'yxatdan o'tmagansiz.")
            return
        
        # Faqat admin yakunlagan (is_finished=True) testlar bo'yicha tahlil va apellyatsiya ko'rsatiladi
        sessions = (await session.execute(
            select(TestSession, Test)
            .join(Test, TestSession.test_id == Test.id)
            .where(TestSession.student_id == student.id, TestSession.status == "COMPLETED", Test.is_finished == True)
            .order_by(TestSession.finished_at.desc())
        )).all()
        
        if not sessions:
            await message.answer("⚠️ Hozirda apellyatsiya berish uchun yakunlangan va natijalari e'lon qilingan testlar mavjud emas.\n(Admin testni to'xtatib, yakunlamaguncha savollar va apellyatsiya yopiq bo'ladi).")
            return
            
        keyboard_buttons = []
        for ts, t in sessions:
            keyboard_buttons.append([InlineKeyboardButton(
                text=f"📚 {t.subject} | {ts.score} ball ({ts.score_percentage}%)",
                callback_data=f"student_appeal_test_{ts.id}"
            )])
            
        markup = InlineKeyboardMarkup(inline_keyboard=keyboard_buttons)
        await message.answer("⚖️ <b>Apellyatsiya bo'limi:</b>\n\nSavollar tahlili va apellyatsiya berish uchun testni tanlang:", reply_markup=markup)

@router.callback_query(F.data.startswith("student_appeal_test_"))
async def show_test_analysis_for_appeal(callback: CallbackQuery):
    session_id = int(callback.data.split("_")[3])
    async with async_session() as session:
        ts = await session.get(TestSession, session_id)
        test = await session.get(Test, ts.test_id)
        
        if not test.is_finished:
            await callback.answer("Bu test uchun hali tahlil ochiq emas!", show_alert=True)
            return

        questions = (await session.execute(select(Question).where(Question.test_id == test.id))).scalars().all()
        answers = {a.question_id: a.selected_option for a in (await session.execute(select(Answer).where(Answer.session_id == ts.id))).scalars().all()}
        
        text = f"📋 <b>Test tahlili: {test.subject} ({test.title})</b>\n" \
               f"⭐ Ball: {ts.score} ({ts.score_percentage}%)\n" \
               f"✅ To'g'ri: {ts.correct_answers} | ❌ Noto'g'ri: {ts.wrong_answers} | ⭕ Javobsiz: {ts.unanswered}\n\n"
               
        keyboard = []
        for idx, q in enumerate(questions, 1):
            sel = answers.get(q.id, "Javob berilmagan")
            status = "✅" if sel == q.correct_option else "❌"
            text += f"<b>{idx}. {q.question_text}</b>\nSizning javob: <b>{sel}</b> {status} | To'g'ri: <b>{q.correct_option}</b>\n\n"
            keyboard.append([InlineKeyboardButton(text=f"⚖️ {idx}-savolga apellyatsiya", callback_data=f"appeal_q_{ts.id}_{q.id}")])
            
        keyboard.append([InlineKeyboardButton(text="🎓 Sertifikatni yuklab olish", callback_data=f"get_cert_{ts.id}")])
        
        if len(text) > 4000:
            text = text[:3900] + "\n... (matn qisqartirildi)"
            
        await callback.message.edit_text(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=keyboard))
        await callback.answer()

@router.callback_query(F.data.startswith("get_cert_"))
async def download_certificate(callback: CallbackQuery, bot: Bot):
    session_id = int(callback.data.split("_")[3]) # updated index due to split
    async with async_session() as session:
        ts = await session.get(TestSession, session_id)
        test = await session.get(Test, ts.test_id)
        student = await session.get(Student, ts.student_id)
        
        pdf_bytes = generate_certificate_pdf(
            f"{student.first_name} {student.last_name}",
            test.title,
            test.subject,
            ts.score_percentage
        )
        file_doc = BufferedInputFile(pdf_bytes, filename=f"sertifikat_{student.student_id}.pdf")
        await bot.send_document(chat_id=callback.message.chat.id, document=file_doc, caption="🎓 Sizning rasmiy sertifikatingiz!")
        await callback.answer()

@router.callback_query(F.data.startswith("appeal_q_"))
async def start_appeal(callback: CallbackQuery, state: FSMContext):
    parts = callback.data.split("_")
    session_id, question_id = int(parts[2]), int(parts[3])
    async with async_session() as session:
        ts = await session.get(TestSession, session_id)
        test = await session.get(Test, ts.test_id)
        if not test.is_finished:
            await callback.answer("Test yakunlanmagani uchun apellyatsiya berib bo'lmaydi!", show_alert=True)
            return

    await state.update_data(appeal_session_id=session_id, appeal_question_id=question_id)
    await state.set_state(AppealState.waiting_for_text)
    await callback.message.answer("✍️ Ushbu savol bo'yicha o'z e'tirozingiz yoki apellyatsiya sababingizni yozib yuboring:")
    await callback.answer()

@router.message(AppealState.waiting_for_text)
async def process_appeal_text(message: Message, state: FSMContext):
    data = await state.get_data()
    async with async_session() as session:
        student = (await session.execute(select(Student).where(Student.telegram_id == message.from_user.id))).scalar_one_or_none()
        appeal = Appeal(
            student_id=student.id,
            test_session_id=data["appeal_session_id"],
            question_id=data["appeal_question_id"],
            message_text=message.text
        )
        session.add(appeal)
        await session.commit()
    
    await state.clear()
    await message.answer("✅ Apellyatsiyangiz adminga yuborildi. Tez orada ko'rib chiqiladi!", reply_markup=get_main_menu())

user_next_question_flags = {}

@router.message(F.text == "📝 Testni boshlash")
async def start_test_prompt(message: Message, state: FSMContext):
    if await state.get_state() == TestProcessState.in_test.state: return
    async with async_session() as session:
        student = (await session.execute(select(Student).where(Student.telegram_id == message.from_user.id))).scalar_one_or_none()
        if not student or not student.grade:
            await message.answer("❌ Profilingizda sinf ko'rsatilmagan yoki ro'yxatdan o'tmagansiz.")
            return

        # Faqat aktiv va admin yakunlamagan testlar ko'rsatiladi
        tests = (await session.execute(
            select(Test).where(Test.is_active == True, Test.is_finished == False, Test.grade_level == student.grade)
        )).scalars().all()
        
        if not tests:
            await message.answer(f"⚠️ Hozirda <b>{student.grade}</b> uchun faol testlar mavjud emas.")
            return

        keyboard_buttons = [[InlineKeyboardButton(text=f"📚 {t.subject} — {t.title}", callback_data=f"start_test_{t.id}")] for t in tests]
        await message.answer("📝 <b>Mavjud testlar:</b>\n\nFanni tanlang:", reply_markup=InlineKeyboardMarkup(inline_keyboard=keyboard_buttons))

@router.callback_query(F.data.startswith("start_test_"))
async def begin_test_session(callback: CallbackQuery, state: FSMContext, bot: Bot):
    test_id = int(callback.data.split("_")[2])
    
    async with async_session() as session:
        student = (await session.execute(select(Student).where(Student.telegram_id == callback.from_user.id))).scalar_one_or_none()
        test = (await session.execute(select(Test).where(Test.id == test_id))).scalar_one_or_none()
        
        if not test or not test.is_active or test.is_finished:
            await callback.answer("Bu test topilmadi yoki yopilgan!", show_alert=True)
            return

        now = datetime.utcnow()
        if test.start_time and now < test.start_time:
            await callback.answer(f"⏳ Test hali boshlanmagan!", show_alert=True)
            return
        if test.end_time and now > test.end_time:
            await callback.answer("⏰ Bu testning vaqti tugagan!", show_alert=True)
            return

        if test.max_attempts > 0:
            attempts_count = await session.scalar(
                select(func.count(TestSession.id)).where(
                    TestSession.student_id == student.id,
                    TestSession.test_id == test_id,
                    TestSession.status == "COMPLETED"
                )
            )
            if attempts_count >= test.max_attempts:
                await callback.answer(f"❌ Urinishlar soni tugagan: {test.max_attempts}", show_alert=True)
                return

        questions = (await session.execute(select(Question).where(Question.test_id == test_id))).scalars().all()
        if not questions:
            await callback.answer("Bu testda savollar yo'q!", show_alert=True)
            return
            
        random.shuffle(questions)
        
        test_session = TestSession(student_id=student.id, test_id=test_id, status="IN_PROGRESS")
        session.add(test_session)
        await session.commit()
        await session.refresh(test_session)

        await callback.message.edit_text(f"🚀 <b>{test.subject}</b> testi boshlandi!")
        user_id = callback.from_user.id
        await state.set_state(TestProcessState.in_test)
        
        for index, q in enumerate(questions):
            # Test davomida admin testni to'xtatgan bo'lsa, testni to'xtatamiz
            current_test_check = await session.get(Test, test_id)
            if not current_test_check.is_active or current_test_check.is_finished:
                break

            options = [("A", q.option_a), ("B", q.option_b)]
            if q.option_c: options.append(("C", q.option_c))
            if q.option_d: options.append(("D", q.option_d))
            random.shuffle(options)
            
            keyboard_buttons = []
            row = []
            for new_key, (orig_key, text_val) in zip(["A", "B", "C", "D"][:len(options)], options):
                row.append(InlineKeyboardButton(text=f"{new_key}) {text_val}", callback_data=f"ans_{test_session.id}_{q.id}_{orig_key}"))
                if len(row) == 2:
                    keyboard_buttons.append(row)
                    row = []
            if row: keyboard_buttons.append(row)
            keyboard_buttons.append([InlineKeyboardButton(text="➡️ Keyingi savol", callback_data=f"next_q_{test_session.id}")])
            markup = InlineKeyboardMarkup(inline_keyboard=keyboard_buttons)
            
            duration_per_q = test.duration_seconds_per_question if test.duration_seconds_per_question else 15
            remaining_time = duration_per_q
            user_next_question_flags[user_id] = False
            
            if q.photo_file_id:
                q_msg = await bot.send_photo(chat_id=user_id, photo=q.photo_file_id, caption=f"<b>Savol {index + 1} / {len(questions)}</b>\n\n{q.question_text}", reply_markup=markup)
            else:
                q_msg = await bot.send_message(chat_id=user_id, text=f"<b>Savol {index + 1} / {len(questions)}</b> (⏱ {remaining_time}s)\n\n{q.question_text}", reply_markup=markup)
            
            for _ in range(duration_per_q):
                await asyncio.sleep(1)
                if user_next_question_flags.get(user_id, False): break
                remaining_time -= 1
                try:
                    if not q.photo_file_id:
                        await bot.edit_message_text(chat_id=user_id, message_id=q_msg.message_id, text=f"<b>Savol {index + 1} / {len(questions)}</b> (⏱ {remaining_time}s)\n\n{q.question_text}", reply_markup=markup)
                except Exception: pass
            
            try: await bot.delete_message(chat_id=user_id, message_id=q_msg.message_id)
            except Exception: pass

        async with async_session() as final_session:
            sess = await final_session.get(TestSession, test_session.id)
            if sess and sess.status == "IN_PROGRESS":
                sess.status = "COMPLETED"
                sess.finished_at = datetime.utcnow()
                await calculate_and_save_results(final_session, sess)
                await final_session.commit()
                
                student_obj = await final_session.get(Student, sess.student_id)
                test_obj = await final_session.get(Test, sess.test_id)
                if student_obj and test_obj:
                    save_result_to_sheet(student_obj.student_id, f"{student_obj.first_name} {student_obj.last_name}", student_obj.school or "-", student_obj.grade or "-", test_obj.title, test_obj.subject, sess.score, sess.score_percentage, sess.correct_answers, sess.wrong_answers)
                
                await state.clear()
                await bot.send_message(chat_id=user_id, text=f"🏆 <b>TEST YAKUNLANDI!</b>\n\nNatijangiz saqlandi. Admin testni yakunlagach, batafsil tahlil va apellyatsiya bo'limi ochiladi.", reply_markup=get_main_menu())

@router.callback_query(F.data.startswith("ans_"))
async def save_answer(callback: CallbackQuery, bot: Bot):
    parts = callback.data.split("_")
    session_id, question_id, selected = int(parts[1]), int(parts[2]), parts[3]
    async with async_session() as session:
        existing = (await session.execute(select(Answer).where(Answer.session_id == session_id, Answer.question_id == question_id))).scalar_one_or_none()
        if existing: existing.selected_option = selected
        else: session.add(Answer(session_id=session_id, question_id=question_id, selected_option=selected))
        await session.commit()
    user_next_question_flags[callback.from_user.id] = True
    await callback.answer(f"Tanlandi: {selected}")
    try: await bot.delete_message(chat_id=callback.message.chat.id, message_id=callback.message.message_id)
    except Exception: pass

@router.callback_query(F.data.startswith("next_q_"))
async def next_question_callback(callback: CallbackQuery, bot: Bot):
    user_next_question_flags[callback.from_user.id] = True
    await callback.answer("Keyingi savol...")
    try: await bot.delete_message(chat_id=callback.message.chat.id, message_id=callback.message.message_id)
    except Exception: pass

async def calculate_and_save_results(session, sess: TestSession):
    questions = (await session.execute(select(Question).where(Question.test_id == sess.test_id))).scalars().all()
    answers = {a.question_id: a.selected_option for a in (await session.execute(select(Answer).where(Answer.session_id == sess.id))).scalars().all()}
    
    correct, wrong, unanswered, total_score = 0, 0, 0, 0.0
    for q in questions:
        sel = answers.get(q.id)
        if not sel: unanswered += 1
        elif sel == q.correct_option:
            correct += 1
            total_score += q.points
        else: wrong += 1
            
    total_q = len(questions)
    sess.correct_answers = correct
    sess.wrong_answers = wrong
    sess.unanswered = unanswered
    sess.score = total_score
    sess.score_percentage = round((correct / total_q) * 100, 2) if total_q > 0 else 0.0

@router.message(F.text == "🏆 Reyting")
async def rating_menu_prompt(message: Message, state: FSMContext):
    if await state.get_state() == TestProcessState.in_test.state: return
    async with async_session() as session:
        tests = (await session.execute(select(Test))).scalars().all()
        if not tests:
            await message.answer("🏆 Hozircha testlar yo'q.")
            return
        keyboard = [[InlineKeyboardButton(text=f"📊 [{t.grade_level}] {t.subject} — {t.title}", callback_data=f"show_rating_{t.id}")] for t in tests]
        await message.answer("🏆 Reyting uchun testni tanlang:", reply_markup=InlineKeyboardMarkup(inline_keyboard=keyboard))

@router.callback_query(F.data.startswith("show_rating_"))
async def show_specific_test_rating(callback: CallbackQuery):
    test_id = int(callback.data.split("_")[2])
    async with async_session() as session:
        test = await session.get(Test, test_id)
        rows = (await session.execute(
            select(Student, TestSession).join(TestSession, Student.id == TestSession.student_id)
            .where(TestSession.test_id == test_id, TestSession.status == "COMPLETED")
            .order_by(TestSession.score.desc()).limit(15)
        )).all()
        
        if not rows:
            await callback.message.edit_text("Natijalar mavjud emas.")
            return
            
        text = f"🏆 <b>REYTING: [{test.grade_level}] {test.subject}</b>\n\n"
        for idx, (s, ts) in enumerate(rows, 1):
            medal = "🥇" if idx == 1 else ("🥈" if idx == 2 else ("🥉" if idx == 3 else f"{idx}."))
            text += f"{medal} {s.first_name} {s.last_name} ({s.grade}) — <b>{ts.score} ball</b> ({ts.score_percentage}%)\n"
        await callback.message.edit_text(text)

@router.message(F.text == "ℹ️ Olimpiada haqida")
async def about_handler(message: Message, state: FSMContext):
    if await state.get_state() == TestProcessState.in_test.state: return
    await message.answer("ℹ️ Professional Olimpiada Tizimi v2.1 — Xavfsizlik va Google Sheets sinxronizatsiyasi bilan.")

@router.message(Command("admin"))
async def admin_panel(message: Message, state: FSMContext):
    if not await is_admin(message.from_user.id): return
    await state.clear()
    await message.answer("🛠 <b>Admin Panel</b>", reply_markup=get_admin_menu())

@router.message(F.text == "➕ ID qo'shish")
async def admin_add_student_prompt(message: Message, state: FSMContext):
    if not await is_admin(message.from_user.id): return
    await state.set_state(AdminAddStudent.waiting_for_data)
    await message.answer("📝 Ma'lumotlarni yuboring:\n<code>Ism, Familiya, Sinf, Maktab</code>")

@router.message(AdminAddStudent.waiting_for_data)
async def admin_save_student(message: Message, state: FSMContext):
    if not await is_admin(message.from_user.id): return
    parts = [p.strip() for p in message.text.split(",")]
    if len(parts) < 4:
        await message.answer("❌ Format xato!")
        return
    unique_id = f"OLM-2026-{''.join(random.choices('ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789', k=6))}"
    async with async_session() as session:
        session.add(Student(student_id=unique_id, first_name=parts[0], last_name=parts[1], grade=parts[2], school=parts[3]))
        await session.commit()
    await state.clear()
    await message.answer(f"✅ O'quvchi qo'shildi!\nID: <code>{unique_id}</code>", reply_markup=get_admin_menu())

@router.message(F.text == "📂 Excel orqali ID'lar")
async def admin_excel_students_prompt(message: Message, state: FSMContext):
    if not await is_admin(message.from_user.id): return
    await state.set_state(AdminAddStudent.waiting_for_excel)
    await message.answer("📂 O'quvchilar ro'yxati bor Excel faylni (`.xlsx`) yuboring.")

@router.message(AdminAddStudent.waiting_for_excel, F.document)
async def admin_process_excel_students(message: Message, state: FSMContext, bot: Bot):
    if not await is_admin(message.from_user.id): return
    document = message.document
    file_info = await bot.get_file(document.file_id)
    downloaded = await bot.download_file(file_info.file_path)
    try:
        df = pd.read_excel(io.BytesIO(downloaded.read()))
        added_count = 0
        async with async_session() as session:
            for _, row in df.iterrows():
                session.add(Student(
                    student_id=f"OLM-2026-{''.join(random.choices('ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789', k=6))}",
                    first_name=str(row.iloc[0]), last_name=str(row.iloc[1]), grade=str(row.iloc[2]), school=str(row.iloc[3])
                ))
                added_count += 1
            await session.commit()
        await state.clear()
        await message.answer(f"✅ {added_count} ta o'quvchi qo'shildi!", reply_markup=get_admin_menu())
    except Exception as e:
        await message.answer(f"❌ Xatolik: {e}")

@router.message(F.text == "📂 Test yuklash")
async def admin_add_test_start(message: Message, state: FSMContext):
    if not await is_admin(message.from_user.id): return
    await state.set_state(AdminAddTest.waiting_for_title)
    await message.answer("📂 Test sarlavhasini kiriting:")

@router.message(AdminAddTest.waiting_for_title)
async def admin_get_title(message: Message, state: FSMContext):
    await state.update_data(title=message.text.strip())
    await state.set_state(AdminAddTest.waiting_for_subject)
    await message.answer("Fan nomini kiriting:")

@router.message(AdminAddTest.waiting_for_subject)
async def admin_get_subject(message: Message, state: FSMContext):
    await state.update_data(subject=message.text.strip())
    await state.set_state(AdminAddTest.waiting_for_grade)
    await message.answer("Sinfni kiriting (masalan: `11-sinf`):")

@router.message(AdminAddTest.waiting_for_grade)
async def admin_get_grade(message: Message, state: FSMContext):
    await state.update_data(grade=message.text.strip())
    await state.set_state(AdminAddTest.waiting_for_attempts)
    await message.answer("Maksimal urinishlar sonini kiriting (masalan: 1):")

@router.message(AdminAddTest.waiting_for_attempts)
async def admin_get_attempts(message: Message, state: FSMContext):
    try: att = int(message.text.strip())
    except: att = 1
    await state.update_data(max_attempts=att, questions=[])
    await state.set_state(AdminAddTest.waiting_for_questions)
    await message.answer("📂 Savollarni matn yoki Word/PDF fayl ko'rinishida yuboring.", reply_markup=get_finish_test_keyboard())

@router.message(AdminAddTest.waiting_for_questions, F.document)
async def admin_handle_document(message: Message, state: FSMContext, bot: Bot):
    if not await is_admin(message.from_user.id): return
    document = message.document
    file_info = await bot.get_file(document.file_id)
    downloaded = await bot.download_file(file_info.file_path)
    file_path = f"temp_{document.file_name}"
    with open(file_path, "wb") as f: f.write(downloaded.read())
    
    extracted_text = ""
    try:
        if document.file_name.endswith('.pdf'):
            reader = pypdf.PdfReader(file_path)
            for page in reader.pages: extracted_text += page.extract_text() + "\n"
        elif document.file_name.endswith('.docx'):
            doc = docx.Document(file_path)
            for para in doc.paragraphs: extracted_text += para.text + "\n"
    except Exception as e:
        await message.answer(f"Xatolik: {e}")
        return
    finally:
        if os.path.exists(file_path): os.remove(file_path)
        
    added = await parse_and_add_questions(extracted_text, state)
    data = await state.get_data()
    await message.answer(f"✅ Fayldan {added} ta savol qo'shildi! Jami: {len(data.get('questions', []))}")

@router.message(AdminAddTest.waiting_for_questions, F.text == "✅ Javobni yuklash / Testni saqlash")
async def admin_ask_for_answers(message: Message, state: FSMContext):
    if not await is_admin(message.from_user.id): return
    data = await state.get_data()
    if not data.get("questions"):
        await message.answer("❌ Savollar mavjud emas!")
        return
    await state.set_state(AdminAddTest.waiting_for_answers)
    await message.answer("🔑 To'g'ri javoblarni ketma-ketlikda yuboring (masalan: `A B C D A...`):", reply_markup=get_admin_menu())

@router.message(AdminAddTest.waiting_for_questions, F.text)
async def admin_add_bulk_questions_text(message: Message, state: FSMContext):
    if not await is_admin(message.from_user.id): return
    added = await parse_and_add_questions(message.text, state)
    data = await state.get_data()
    await message.answer(f"✅ {added} ta savol qo'shildi! Jami: {len(data.get('questions', []))}")

async def parse_and_add_questions(text: str, state: FSMContext) -> int:
    data = await state.get_data()
    questions_list = data.get("questions", [])
    lines = [line.strip() for line in text.split("\n") if line.strip()]
    added = 0
    i = 0
    while i < len(lines):
        a_idx, b_idx, c_idx, d_idx = -1, -1, -1, -1
        for j in range(i + 1, min(i + 6, len(lines))):
            l = lines[j].lower()
            if l.startswith("a)") or l.startswith("a."): a_idx = j
            elif l.startswith("b)") or l.startswith("b."): b_idx = j
            elif l.startswith("c)") or l.startswith("c."): c_idx = j
            elif l.startswith("d)") or l.startswith("d."): d_idx = j

        if a_idx != -1 and b_idx != -1:
            q_text = " ".join(lines[i:a_idx])
            questions_list.append({
                "text": q_text, 
                "a": re.sub(r'^[aA][\.\)]\s*', '', lines[a_idx]), 
                "b": re.sub(r'^[bB][\.\)]\s*', '', lines[b_idx]),
                "c": re.sub(r'^[cC][\.\)]\s*', '', lines[c_idx]) if c_idx != -1 else "C",
                "d": re.sub(r'^[dD][\.\)]\s*', '', lines[d_idx]) if d_idx != -1 else "D",
                "correct": "A"
            })
            added += 1
            i = max(a_idx, b_idx, c_idx, d_idx) + 1
        else: i += 1
    await state.update_data(questions=questions_list)
    return added

@router.message(AdminAddTest.waiting_for_answers)
async def admin_save_answers_and_test(message: Message, state: FSMContext):
    if not await is_admin(message.from_user.id): return
    tokens = re.findall(r'[A-D]', message.text.upper())
    data = await state.get_data()
    questions_list = data.get("questions", [])
    
    for idx, q in enumerate(questions_list):
        if idx < len(tokens): q["correct"] = tokens[idx]

    async with async_session() as session:
        new_test = Test(
            title=data["title"], subject=data["subject"], grade_level=data["grade"],
            max_attempts=data["max_attempts"], is_active=True, is_finished=False
        )
        session.add(new_test)
        await session.flush()
        
        for q in questions_list:
            session.add(Question(
                test_id=new_test.id, question_text=q["text"],
                option_a=q["a"], option_b=q["b"], option_c=q["c"], option_d=q["d"],
                correct_option=q["correct"]
            ))
        await session.commit()
    await state.clear()
    await message.answer("✅ Test saqlandi!", reply_markup=get_admin_menu())

@router.message(F.text == "⚙️ Testlarni boshqarish")
async def manage_tests_admin(message: Message):
    if not await is_admin(message.from_user.id): return
    async with async_session() as session:
        tests = (await session.execute(select(Test))).scalars().all()
        if not tests:
            await message.answer("Testlar yo'q.")
            return
        keyboard = []
        for t in tests:
            status_emoji = '🟢 Aktiv' if t.is_active and not t.is_finished else ('🏁 Yakunlangan' if t.is_finished else '🔴 To\'xtatilgan')
            keyboard.append([
                InlineKeyboardButton(text=f"[{t.grade_level}] {t.subject} ({status_emoji})", callback_data="none"),
            ])
            keyboard.append([
                InlineKeyboardButton(text="🔄 Holatni o'zgartirish", callback_data=f"toggle_test_{t.id}"),
                InlineKeyboardButton(text="🏁 Testni yakunlash", callback_data=f"finish_test_{t.id}"),
                InlineKeyboardButton(text="🗑 O'chirish", callback_data=f"delete_test_{t.id}")
            ])
        await message.answer("⚙️ Testlarni boshqarish:", reply_markup=InlineKeyboardMarkup(inline_keyboard=keyboard))

@router.callback_query(F.data.startswith("toggle_test_"))
async def toggle_test(callback: CallbackQuery):
    async with async_session() as session:
        test = await session.get(Test, int(callback.data.split("_")[2]))
        if test:
            test.is_active = not test.is_active
            await session.commit()
            await callback.answer(f"Test holati o'zgardi! Aktiv: {test.is_active}")

@router.callback_query(F.data.startswith("finish_test_"))
async def finish_test_by_admin(callback: CallbackQuery):
    async with async_session() as session:
        test = await session.get(Test, int(callback.data.split("_")[2]))
        if test:
            test.is_active = False
            test.is_finished = True # Shu orqali o'quvchilarga savollar, tahlil va apellyatsiya ochiladi
            await session.commit()
            await callback.answer("Test to'liq yakunlandi! Endi o'quvchilar tahlil va apellyatsiya ko'rishlari mumkin.", show_alert=True)

@router.callback_query(F.data.startswith("delete_test_"))
async def delete_test(callback: CallbackQuery):
    async with async_session() as session:
        test = await session.get(Test, int(callback.data.split("_")[2]))
        if test:
            await session.delete(test)
            await session.commit()
            await callback.answer("Test o'chirildi!")

@router.message(F.text == "⚖️ Apellyatsiyalar")
async def admin_appeals_handler(message: Message):
    if not await is_admin(message.from_user.id): return
    async with async_session() as session:
        appeals = (await session.execute(select(Appeal).where(Appeal.status == "PENDING"))).scalars().all()
        if not appeals:
            await message.answer("Ko'rib chiqilishi kerak bo'lgan apellyatsiyalar yo'q.")
            return
        keyboard = [[InlineKeyboardButton(text=f"Apellyatsiya #{a.id}", callback_data=f"view_appeal_{a.id}")] for a in appeals]
        await message.answer("⚖️ Apellyatsiyalar ro'yxati:", reply_markup=InlineKeyboardMarkup(inline_keyboard=keyboard))

@router.callback_query(F.data.startswith("view_appeal_"))
async def view_appeal(callback: CallbackQuery):
    appeal_id = int(callback.data.split("_")[2])
    async with async_session() as session:
        appeal = await session.get(Appeal, appeal_id)
        if not appeal:
            await callback.answer("Topilmadi!", show_alert=True)
            return
        student = await session.get(Student, appeal.student_id)
        question = await session.get(Question, appeal.question_id)
        test_session = await session.get(TestSession, appeal.test_session_id)
        test = await session.get(Test, test_session.test_id)
        
        text = f"⚖️ <b>Apellyatsiya #{appeal.id}</b>\n\n" \
               f"👤 O'quvchi: <b>{student.first_name} {student.last_name}</b> ({student.student_id})\n" \
               f"📚 Test: {test.subject} — {test.title}\n" \
               f"❓ Savol: {question.question_text}\n" \
               f"✍️ E'tiroz: <i>{appeal.message_text}</i>\n" \
               f"📌 Holati: <b>{appeal.status}</b>"
               
        keyboard = [
            [
                InlineKeyboardButton(text="✅ Tasdiqlash (+1 ball)", callback_data=f"app_acc_{appeal.id}"),
                InlineKeyboardButton(text="❌ Rad etish", callback_data=f"app_rej_{appeal.id}")
            ]
        ]
        await callback.message.edit_text(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=keyboard))
        await callback.answer()

@router.callback_query(F.data.startswith("app_acc_"))
async def accept_appeal(callback: CallbackQuery, bot: Bot):
    appeal_id = int(callback.data.split("_")[2])
    async with async_session() as session:
        appeal = await session.get(Appeal, appeal_id)
        if not appeal: return
        appeal.status = "APPROVED"
        
        ts = await session.get(TestSession, appeal.test_session_id)
        ts.score += 1.0
        ts.correct_answers += 1
        
        questions_count = await session.scalar(select(func.count(Question.id)).where(Question.test_id == ts.test_id))
        if questions_count > 0:
            ts.score_percentage = round((ts.score / questions_count) * 100, 2)
            
        await session.commit()
        
        student = await session.get(Student, appeal.student_id)
        test = await session.get(Test, ts.test_id)
        
        # Google Sheets jadvalidagi natijani to'g'ri yangilash
        update_result_in_sheet(student.student_id, test.title, ts.score, ts.score_percentage, ts.correct_answers)
        
        if student and student.telegram_id:
            try:
                await bot.send_message(
                    chat_id=student.telegram_id,
                    text=f"✅ <b>Apellyatsiyangiz ma'qullandi!</b>\n\nBaligacha 1 ball qo'shildi. Hozirgi balingiz: <b>{ts.score}</b>"
                )
            except Exception: pass
            
    await callback.message.edit_text("✅ Apellyatsiya tasdiqlandi, ball qo'shildi va Google Sheets jadvali yangilandi!")

@router.callback_query(F.data.startswith("app_rej_"))
async def reject_appeal(callback: CallbackQuery, bot: Bot):
    appeal_id = int(callback.data.split("_")[2])
    async with async_session() as session:
        appeal = await session.get(Appeal, appeal_id)
        if not appeal: return
        appeal.status = "REJECTED"
        await session.commit()
        
        student = await session.get(Student, appeal.student_id)
        if student and student.telegram_id:
            try:
                await bot.send_message(
                    chat_id=student.telegram_id,
                    text=f"❌ <b>Apellyatsiyangiz rad etildi.</b>"
                )
            except Exception: pass
            
    await callback.message.edit_text("❌ Apellyatsiya rad etildi.")

@router.message(F.text == "👥 Adminlar")
async def admins_management(message: Message, state: FSMContext):
    if message.from_user.id not in SUPER_ADMIN_IDS: return
    async with async_session() as session:
        admins = (await session.execute(select(Admin))).scalars().all()
        text = "👥 <b>Moderator adminlar:</b>\n\n" + "\n".join([str(a.telegram_id) for a in admins]) if admins else "Adminlar yo'q."
        keyboard = [[InlineKeyboardButton(text="➕ Admin qo'shish", callback_data="add_admin")]]
        await message.answer(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=keyboard))

@router.callback_query(F.data == "add_admin")
async def add_admin_prompt(callback: CallbackQuery, state: FSMContext):
    await state.set_state(AdminManageAdmins.waiting_for_id)
    await callback.message.answer("Yangi adminning Telegram ID raqamini kiriting:")
    await callback.answer()

@router.message(AdminManageAdmins.waiting_for_id)
async def save_new_admin(message: Message, state: FSMContext):
    try:
        tg_id = int(message.text.strip())
        async with async_session() as session:
            session.add(Admin(telegram_id=tg_id, role="moderator"))
            await session.commit()
        await state.clear()
        await message.answer("✅ Yangi admin qo'shildi!", reply_markup=get_admin_menu())
    except Exception as e:
        await message.answer(f"Xatolik: {e}")

@router.message(F.text == "📊 Jonli statistika")
async def live_statistics(message: Message):
    if not await is_admin(message.from_user.id): return
    async with async_session() as session:
        total = await session.scalar(select(func.count(Student.id)))
        completed = await session.scalar(select(func.count(TestSession.id)).where(TestSession.status == "COMPLETED"))
        await message.answer(f"📊 <b>Statistika:</b>\n\nJami o'quvchilar: {total}\nTest topshirganlar: {completed or 0}")

@router.message(F.text == "📥 Excel natijalar")
async def export_excel_results(message: Message, bot: Bot):
    if not await is_admin(message.from_user.id): return
    async with async_session() as session:
        rows = (await session.execute(select(Student, TestSession, Test).join(TestSession, Student.id == TestSession.student_id).join(Test, TestSession.test_id == Test.id))).all()
        if not rows:
            await message.answer("Natijalar yo'q.")
            return
        df = pd.DataFrame([{
            "ID": s.student_id, "Ism": s.first_name, "Familiya": s.last_name, "Maktab": s.school,
            "Sinf": s.grade, "Fan": t.subject, "Ball": ts.score, "Foiz": ts.score_percentage, "Sana": ts.finished_at
        } for s, ts, t in rows])
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer: df.to_excel(writer, index=False)
        output.seek(0)
        await message.answer_document(BufferedInputFile(output.read(), filename="natijalar.xlsx"))

@router.message(F.text == "🧹 Bazani tozalash")
async def reset_db_prompt(message: Message):
    if message.from_user.id not in SUPER_ADMIN_IDS: return
    keyboard = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="Tasdiqlash ⚠️", callback_data="reset_db_confirm")]])
    await message.answer("Barcha sessiyalarni tozalashni tasdiqlaysizmi?", reply_markup=keyboard)

@router.callback_query(F.data == "reset_db_confirm")
async def confirm_reset(callback: CallbackQuery):
    async with async_session() as session:
        await session.execute(delete(TestSession))
        await session.commit()
    await callback.message.edit_text("✅ Sessiyalar tozalandi.")

@router.message(F.text == "📢 Xabar yuborish")
async def broadcast_prompt(message: Message, state: FSMContext):
    if not await is_admin(message.from_user.id): return
    await state.set_state(AdminBroadcast.waiting_for_message)
    await message.answer("Yubormoqchi bo'lgan xabaringizni kiriting:")

@router.message(AdminBroadcast.waiting_for_message)
async def send_broadcast(message: Message, state: FSMContext, bot: Bot):
    if not await is_admin(message.from_user.id): return
    async with async_session() as session:
        students = (await session.execute(select(Student).where(Student.telegram_id.is_not(None)))).scalars().all()
    success = 0
    for s in students:
        try:
            await message.send_copy(chat_id=s.telegram_id)
            success += 1
            await asyncio.sleep(0.05)
        except: pass
    await state.clear()
    await message.answer(f"✅ Xabar {success} ta o'quvchiga yuborildi!", reply_markup=get_admin_menu())

@router.message(F.text == "⬅️ Bosh menyu")
async def back_to_menu(message: Message, state: FSMContext):
    if await state.get_state() == TestProcessState.in_test.state: return
    await state.clear()
    if await is_admin(message.from_user.id):
        await message.answer("Admin menyu:", reply_markup=get_admin_menu())
    else:
        await message.answer("Asosiy menyu:", reply_markup=get_main_menu())

async def main():
    logging.basicConfig(level=logging.INFO, stream=sys.stdout)
    await init_db()
    bot = Bot(token=BOT_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    dp = Dispatcher(storage=MemoryStorage())
    dp.include_router(router)
    await bot.delete_webhook(drop_pending_updates=True)
    logging.info("🚀 Bot ishga tushdi!")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
