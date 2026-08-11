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
REQUIRED_CHANNEL = "@olimpiada01111"

SUPER_ADMIN_IDS = [8317043750]

Base = declarative_base()

class Student(Base):
    __tablename__ = "students"
    id = Column(Integer, primary_key=True, autoincrement=True)
    student_id = Column(String(50), unique=True, index=True, nullable=False)
    first_name = Column(String(100), nullable=False)
    last_name = Column(String(100), nullable=False)
    age = Column(String(20), nullable=True)
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
    mode = Column(String(20), default="dtm_block") # DTM blok testi rejimi uchun
    duration_minutes = Column(Integer, default=180) # 3 soatlik umumiy vaqt
    start_time = Column(DateTime, nullable=True)
    end_time = Column(DateTime, nullable=True)
    is_active = Column(Boolean, default=True)
    is_finished = Column(Boolean, default=False)
    
    questions = relationship("Question", back_populates="test", cascade="all, delete-orphan")
    test_sessions = relationship("TestSession", back_populates="test", cascade="all, delete-orphan")

class Question(Base):
    __tablename__ = "questions"
    id = Column(Integer, primary_key=True, autoincrement=True)
    test_id = Column(Integer, ForeignKey("tests.id", ondelete="CASCADE"), nullable=False)
    subject_type = Column(String(50), default="main_1") # DTM bloklari uchun fan turi: history, native_lang, math, main_1, main_2
    question_text = Column(Text, nullable=False)
    photo_file_id = Column(String(200), nullable=True)
    option_a = Column(Text, nullable=False)
    option_b = Column(Text, nullable=False)
    option_c = Column(Text, nullable=True)
    option_d = Column(Text, nullable=True)
    correct_option = Column(String(5), nullable=False)
    points = Column(Float, default=1.1)
    
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

# --- KANALGA OBUNANI TEKSHIRISH ---
async def check_subscription(user_id: int, bot: Bot) -> bool:
    try:
        member = await bot.get_chat_member(chat_id=REQUIRED_CHANNEL, user_id=user_id)
        if member.status in ["member", "administrator", "creator"]:
            return True
    except Exception:
        pass
    return False

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

def save_result_to_sheet(student_id, full_name, age, school, grade, test_title, subject, score, percentage, correct, wrong):
    try:
        sheet = get_gspread_sheet()
        row_data = [
            str(student_id), str(full_name), str(age), str(school), str(grade),
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
            if str(row.get("ID")) == str(student_id) and str(row.get("Test")) == str(test_title):
                sheet.update_cell(idx, 8, str(score))
                sheet.update_cell(idx, 9, f"{percentage}%")
                sheet.update_cell(idx, 10, str(correct))
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

# --- STATES ---
class SelfRegState(StatesGroup):
    waiting_for_fullname = State()
    waiting_for_age = State()
    waiting_for_grade = State()
    waiting_for_school = State()

class AdminAddStudent(StatesGroup):
    waiting_for_data = State()
    waiting_for_excel = State()

class AdminAddTest(StatesGroup):
    waiting_for_title = State()
    waiting_for_subject_1 = State()
    waiting_for_subject_2 = State()
    waiting_for_grade = State()
    waiting_for_attempts = State()
    waiting_for_questions = State()

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
            [KeyboardButton(text="📚 Blok test ishlash"), KeyboardButton(text="👤 Profilim")],
            [KeyboardButton(text="📊 Mening urinishlarim"), KeyboardButton(text="⚖️ Apellyatsiya")],
            [KeyboardButton(text="🏆 Reyting"), KeyboardButton(text="ℹ️ Olimpiada haqida")]
        ],
        resize_keyboard=True
    )

def get_admin_menu():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="➕ ID qo'shish"), KeyboardButton(text="📂 Excel orqali ID'lar")],
            [KeyboardButton(text="➕ Blok test qo'shish (Admin)"), KeyboardButton(text="⚙️ Testlarni boshqarish")],
            [KeyboardButton(text="📊 Jonli statistika"), KeyboardButton(text="📥 Excel natijalar")],
            [KeyboardButton(text="⚖️ Apellyatsiyalar"), KeyboardButton(text="👥 Adminlar")],
            [KeyboardButton(text="🧹 Bazani tozalash"), KeyboardButton(text="📢 Xabar yuborish")],
            [KeyboardButton(text="⬅️ Bosh menyu")]
        ],
        resize_keyboard=True
    )

@router.message(Command("start"))
async def cmd_start(message: Message, state: FSMContext, bot: Bot):
    await state.clear()
    if await is_admin(message.from_user.id):
        await message.answer("🛠 <b>Xush kelibsiz, Admin!</b>", reply_markup=get_admin_menu())
        return

    is_subscribed = await check_subscription(message.from_user.id, bot)
    if not is_subscribed:
        keyboard = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="📢 Kanalga obuna bo'lish", url=f"https://t.me/{REQUIRED_CHANNEL.replace('@', '')}")],
            [InlineKeyboardButton(text="✅ Obunani tekshirish", callback_data="check_sub")]
        ])
        await message.answer(f"⚠️ Botdan foydalanish uchun avval quyidagi kanalga obuna bo'lishingiz kerak:\n\n{REQUIRED_CHANNEL}", reply_markup=keyboard)
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

    await state.set_state(SelfRegState.waiting_for_fullname)
    await message.answer("🎓 <b>Olimpiada tizimiga xush kelibsiz!</b>\n\nIltimos, to'liq <b>Ism va Familiyangizni</b> kiriting:")

@router.callback_query(F.data == "check_sub")
async def check_sub_callback(callback: CallbackQuery, state: FSMContext, bot: Bot):
    is_subscribed = await check_subscription(callback.from_user.id, bot)
    if not is_subscribed:
        await callback.answer("❌ Hali kanalga obuna bo'lmagansiz!", show_alert=True)
        return
    
    await callback.message.delete()
    async with async_session() as session:
        result = await session.execute(select(Student).where(Student.telegram_id == callback.from_user.id))
        student = result.scalar_one_or_none()
        
        if student:
            await callback.message.answer(f"Xush kelibsiz, <b>{student.first_name} {student.last_name}</b>!", reply_markup=get_main_menu())
            return

    await state.set_state(SelfRegState.waiting_for_fullname)
    await callback.message.answer("🎓 Muvaffaqiyatli obuna bo'ldingiz!\n\nIltimos, to'liq <b>Ism va Familiyangizni</b> kiriting:")

@router.message(SelfRegState.waiting_for_fullname)
async def process_self_fullname(message: Message, state: FSMContext):
    await state.update_data(fullname=message.text.strip())
    await state.set_state(SelfRegState.waiting_for_age)
    await message.answer("Rahmat! Endi yoshingizni kiriting (masalan: 16):")

@router.message(SelfRegState.waiting_for_age)
async def process_self_age(message: Message, state: FSMContext):
    await state.update_data(age=message.text.strip())
    await state.set_state(SelfRegState.waiting_for_grade)
    await message.answer("Sinfingizni kiriting (masalan: 11-sinf):")

@router.message(SelfRegState.waiting_for_grade)
async def process_self_grade(message: Message, state: FSMContext):
    await state.update_data(grade=message.text.strip())
    await state.set_state(SelfRegState.waiting_for_school)
    await message.answer("Maktabingiz raqami yoki nomini kiriting:")

@router.message(SelfRegState.waiting_for_school)
async def process_self_school(message: Message, state: FSMContext):
    data = await state.get_data()
    fullname_parts = data["fullname"].split(" ", 1)
    first_name = fullname_parts[0]
    last_name = fullname_parts[1] if len(fullname_parts) > 1 else "-"
    
    unique_id = f"OLM-2026-{''.join(random.choices('ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789', k=6))}"
    
    async with async_session() as session:
        student = Student(
            student_id=unique_id,
            first_name=first_name,
            last_name=last_name,
            age=data["age"],
            grade=data["grade"],
            school=message.text.strip(),
            telegram_id=message.from_user.id
        )
        session.add(student)
        await session.commit()
        
    await state.clear()
    await message.answer(f"✅ Muvaffaqiyatli ro'yxatdan o'tdingiz!\nID raqamingiz: <code>{unique_id}</code>", reply_markup=get_main_menu())

@router.message(F.text == "👤 Profilim")
async def profile_handler(message: Message, state: FSMContext):
    if await state.get_state() == TestProcessState.in_test.state: return
    async with async_session() as session:
        student = (await session.execute(select(Student).where(Student.telegram_id == message.from_user.id))).scalar_one_or_none()
        if not student:
            await message.answer("Siz ro'yxatdan o'tmagansiz. /start ni bosing.")
            return
        await message.answer(f"👤 <b>Profil:</b>\n\nID: <code>{student.student_id}</code>\nIsm: {student.first_name} {student.last_name}\nYosh: {student.age or '-'}\nMaktab: {student.school or '-'}\nSinf: {student.grade or '-'}")

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
            status_text = "🟢 Test yakunlangan (Natijalar ochiq)" if t.is_finished else "🟡 Test hali davom etmoqda"
            text += f"📚 <b>{t.subject}</b> ({t.title})\n⭐ Ball: {ts.score} ({ts.score_percentage}%)\n📅 Sana: {date_str}\nStatus: {status_text}\n----------------------------------\n"
            
        await message.answer(text)

@router.message(F.text == "⚖️ Apellyatsiya")
async def student_appeal_menu(message: Message, state: FSMContext):
    if await state.get_state() == TestProcessState.in_test.state: return
    async with async_session() as session:
        student = (await session.execute(select(Student).where(Student.telegram_id == message.from_user.id))).scalar_one_or_none()
        if not student:
            await message.answer("Ro'yxatdan o'tmagansiz.")
            return
        
        sessions = (await session.execute(
            select(TestSession, Test)
            .join(Test, TestSession.test_id == Test.id)
            .where(TestSession.student_id == student.id, TestSession.status == "COMPLETED", Test.is_finished == True)
            .order_by(TestSession.finished_at.desc())
        )).all()
        
        if not sessions:
            await message.answer("⚠️ Hozirda apellyatsiya berish uchun yakunlangan va natijalari e'lon qilingan testlar mavjud emas.")
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
            text += f"<b>{idx}. [{q.subject_type}] {q.question_text}</b>\nSizning javob: <b>{sel}</b> {status} | To'g'ri: <b>{q.correct_option}</b>\n\n"
            keyboard.append([InlineKeyboardButton(text=f"⚖️ {idx}-savolga apellyatsiya", callback_data=f"appeal_q_{ts.id}_{q.id}")])
            
        keyboard.append([InlineKeyboardButton(text="🎓 Sertifikatni yuklab olish", callback_data=f"get_cert_{ts.id}")])
        
        if len(text) > 4000:
            text = text[:3900] + "\n... (matn qisqartirildi)"
            
        await callback.message.edit_text(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=keyboard))
        await callback.answer()

@router.callback_query(F.data.startswith("get_cert_"))
async def download_certificate(callback: CallbackQuery, bot: Bot):
    session_id = int(callback.data.split("_")[3])
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

# --- DTM BLOK TEST ISHLASH & GLOBAL TAYMER (3 SOAT) ---
@router.message(F.text == "📚 Blok test ishlash")
async def list_block_tests(message: Message, state: FSMContext):
    if await state.get_state() == TestProcessState.in_test.state: return
    async with async_session() as session:
        student = (await session.execute(select(Student).where(Student.telegram_id == message.from_user.id))).scalar_one_or_none()
        if not student or not student.grade:
            await message.answer("❌ Profilingizda sinf ko'rsatilmagan yoki ro'yxatdan o'tmagansiz.")
            return

        tests = (await session.execute(
            select(Test).where(Test.is_active == True, Test.is_finished == False, Test.grade_level == student.grade)
        )).scalars().all()
        
        if not tests:
            await message.answer(f"⚠️ Hozirda <b>{student.grade}</b> uchun faol blok testlar mavjud emas.")
            return

        kb = []
        for t in tests:
            kb.append([InlineKeyboardButton(text=f"📌 {t.title} ({t.subject})", callback_data=f"start_btest_{t.id}")])
            
        await message.answer("Mavjud blok testlardan birini tanlang:", reply_markup=InlineKeyboardMarkup(inline_keyboard=kb))

@router.callback_query(F.data.startswith("start_btest_"))
async def start_user_block_test(callback: CallbackQuery, state: FSMContext, bot: Bot):
    test_id = int(callback.data.split("_")[2])
    
    async with async_session() as session:
        student = (await session.execute(select(Student).where(Student.telegram_id == callback.from_user.id))).scalar_one_or_none()
        test = await session.get(Test, test_id)
        
        if not test or not test.is_active or test.is_finished:
            await callback.answer("Bu test topilmadi yoki yopilgan!", show_alert=True)
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

        rows = (await session.execute(select(Question).where(Question.test_id == test_id))).scalars().all()
        if not rows:
            await callback.answer("Bu testda savollar topilmadi!", show_alert=True)
            return
            
        questions = []
        for r in rows:
            questions.append({
                'id': r.id, 'subject': r.subject_type, 'text': r.question_text, 'photo': r.photo_file_id,
                'options': [r.option_a, r.option_b, r.option_c, r.option_d]
            })
            
        test_session = TestSession(student_id=student.id, test_id=test_id, status="IN_PROGRESS")
        session.add(test_session)
        await session.commit()
        await session.refresh(test_session)

        await state.update_data(
            test_session_id=test_session.id,
            test_id=test_id,
            questions=questions,
            current_index=0,
            user_answers={},
            active_session=True
        )
        await state.set_state(TestProcessState.in_test)
        
        await callback.message.answer(
            "⏱ **Blok test boshlandi!**\n"
            "Sizga umumiy **3 soat** vaqt berildi. Savollar ketma-ket ko'rsatiladi:", parse_mode="Markdown"
        )
        
        # 3 soatlik global taymer fonda ishga tushadi (3 * 3600 soniya)
        asyncio.create_task(block_test_timer(bot, callback.from_user.id, test_id, test_session.id, state))
        
        await send_block_question(callback.message, state, bot)

async def block_test_timer(bot: Bot, user_id: int, test_id: int, session_id: int, state: FSMContext):
    await asyncio.sleep(3 * 60 * 60) # 3 soat
    data = await state.get_data()
    if data.get('active_session') and data.get('test_id') == test_id:
        async with async_session() as session:
            sess = await session.get(TestSession, session_id)
            if sess and sess.status == "IN_PROGRESS":
                sess.status = "COMPLETED"
                sess.finished_at = datetime.utcnow()
                await calculate_dtm_results(session, sess)
                await session.commit()
                
                student_obj = await session.get(Student, sess.student_id)
                test_obj = await session.get(Test, sess.test_id)
                if student_obj and test_obj:
                    save_result_to_sheet(student_obj.student_id, f"{student_obj.first_name} {student_obj.last_name}", student_obj.age or "-", student_obj.school or "-", student_obj.grade or "-", test_obj.title, test_obj.subject, sess.score, sess.score_percentage, sess.correct_answers, sess.wrong_answers)
                
                await state.clear()
                await bot.send_message(user_id, "⏰ **Vaqt tugadi!** 3 soatlik vaqt o'z nihoyasiga yetdi. Test avtomatik yakunlandi.", reply_markup=get_main_menu())

async def send_block_question(message: Message, state: FSMContext, bot: Bot):
    data = await state.get_data()
    questions = data['questions']
    index = data['current_index']
    
    if index >= len(questions):
        await finish_block_test(bot, message.chat.id, state)
        return
        
    q = questions[index]
    
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [
            InlineKeyboardButton(text="A", callback_data=f"b_ans_{q['id']}_A"),
            InlineKeyboardButton(text="B", callback_data=f"b_ans_{q['id']}_B"),
            InlineKeyboardButton(text="C", callback_data=f"b_ans_{q['id']}_C"),
            InlineKeyboardButton(text="D", callback_data=f"b_ans_{q['id']}_D"),
        ]
    ])
    
    text = f"<b>{index + 1}-savol (Fan bloki: {q['subject']}):</b>\n\n{q['text']}\n\nA) {q['options'][0]}\nB) {q['options'][1]}\nC) {q['options'][2]}\nD) {q['options'][3]}"
    
    if q['photo']:
        await bot.send_photo(chat_id=message.chat.id, photo=q['photo'], caption=text, reply_markup=kb, parse_mode="HTML")
    else:
        await bot.send_message(chat_id=message.chat.id, text=text, reply_markup=kb, parse_mode="HTML")

@router.callback_query(F.data.startswith("b_ans_"), TestProcessState.in_test)
async def process_user_answer(callback: CallbackQuery, state: FSMContext, bot: Bot):
    data = callback.data.split("_")
    q_id = int(data[2])
    ans = data[3]
    
    st_data = await state.get_data()
    session_id = st_data['test_session_id']
    user_answers = st_data['user_answers']
    user_answers[q_id] = ans
    
    async with async_session() as session:
        existing = (await session.execute(select(Answer).where(Answer.session_id == session_id, Answer.question_id == q_id))).scalar_one_or_none()
        if existing: existing.selected_option = ans
        else: session.add(Answer(session_id=session_id, question_id=q_id, selected_option=ans))
        await session.commit()
    
    await state.update_data(user_answers=user_answers, current_index=st_data['current_index'] + 1)
    
    try: await callback.message.delete()
    except Exception: pass
    
    await callback.answer(f"Tanlandi: {ans}")
    await send_block_question(callback.message, state, bot)

async def calculate_dtm_results(session, sess: TestSession):
    questions = (await session.execute(select(Question).where(Question.test_id == sess.test_id))).scalars().all()
    answers = {a.question_id: a.selected_option for a in (await session.execute(select(Answer).where(Answer.session_id == sess.id))).scalars().all()}
    
    # DTM Ballarni hisoblash qoidasi:
    # Majburiy fanlar (history, native_lang, math) - har biri 1.1 balldan.
    # 1-asosiy fan (main_1) - 3.1 balldan.
    # 2-asosiy fan (main_2) - 2.1 balldan.
    weights = {
        'history': 1.1,
        'native_lang': 1.1,
        'math': 1.1,
        'main_1': 3.1,
        'main_2': 2.1
    }
    
    correct, wrong, unanswered, total_score = 0, 0, 0, 0.0
    for q in questions:
        sel = answers.get(q.id)
        if not sel: 
            unanswered += 1
        elif sel == q.correct_option:
            correct += 1
            total_score += weights.get(q.subject_type, 1.1)
        else: 
            wrong += 1
            
    sess.correct_answers = correct
    sess.wrong_answers = wrong
    sess.unanswered = unanswered
    sess.score = round(total_score, 2)
    sess.score_percentage = round((correct / len(questions)) * 100, 2) if len(questions) > 0 else 0.0

async def finish_block_test(bot: Bot, user_id: int, state: FSMContext):
    data = await state.get_data()
    if not data.get('active_session'):
        return
        
    session_id = data['test_session_id']
    test_id = data['test_id']
    
    async with async_session() as session:
        sess = await session.get(TestSession, session_id)
        if sess and sess.status == "IN_PROGRESS":
            sess.status = "COMPLETED"
            sess.finished_at = datetime.utcnow()
            await calculate_dtm_results(session, sess)
            await session.commit()
            
            student_obj = await session.get(Student, sess.student_id)
            test_obj = await session.get(Test, sess.test_id)
            if student_obj and test_obj:
                save_result_to_sheet(student_obj.student_id, f"{student_obj.first_name} {student_obj.last_name}", student_obj.age or "-", student_obj.school or "-", student_obj.grade or "-", test_obj.title, test_obj.subject, sess.score, sess.score_percentage, sess.correct_answers, sess.wrong_answers)
            
            await state.clear()
            
            await bot.send_message(
                user_id,
                f"📊 **Blok test yakunlandi!**\n\n"
                f"✅ To'g'ri javoblar: {sess.correct_answers} ta\n"
                f"🎯 To'plangan umumiy ball: **{sess.score}** ball",
                parse_mode="Markdown",
                reply_markup=get_main_menu()
            )

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
    await message.answer("ℹ️ Professional DTM Blok Test Tizimi v2.5 — 3 soatlik global taymer va DTM ballar tizimi bilan.")

@router.message(Command("admin"))
async def admin_panel(message: Message, state: FSMContext):
    if not await is_admin(message.from_user.id): return
    await state.clear()
    await message.answer("🛠 <b>Admin Panel</b>", reply_markup=get_admin_menu())

@router.message(F.text == "➕ ID qo'shish")
async def admin_add_student_prompt(message: Message, state: FSMContext):
    if not await is_admin(message.from_user.id): return
    await state.set_state(AdminAddStudent.waiting_for_data)
    await message.answer("📝 Ma'lumotlarni yuboring:\n<code>Ism, Familiya, Yosh, Sinf, Maktab</code>")

@router.message(AdminAddStudent.waiting_for_data)
async def admin_save_student(message: Message, state: FSMContext):
    if not await is_admin(message.from_user.id): return
    parts = [p.strip() for p in message.text.split(",")]
    if len(parts) < 5:
        await message.answer("❌ Format xato! 5 ta ma'lumot kiritilishi kerak.")
        return
    unique_id = f"OLM-2026-{''.join(random.choices('ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789', k=6))}"
    async with async_session() as session:
        session.add(Student(student_id=unique_id, first_name=parts[0], last_name=parts[1], age=parts[2], grade=parts[3], school=parts[4]))
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
                    first_name=str(row.iloc[0]), last_name=str(row.iloc[1]), age=str(row.iloc[2]), grade=str(row.iloc[3]), school=str(row.iloc[4])
                ))
                added_count += 1
            await session.commit()
        await state.clear()
        await message.answer(f"✅ {added_count} ta o'quvchi qo'shildi!", reply_markup=get_admin_menu())
    except Exception as e:
        await message.answer(f"❌ Xatolik: {e}")

# --- ADMIN: DTM BLOK TEST QO'SHISH ---
@router.message(F.text == "➕ Blok test qo'shish (Admin)")
async def admin_add_test(message: Message, state: FSMContext):
    if not await is_admin(message.from_user.id): return
    await message.answer("Blok test nomini kiriting (masalan: '1-sonli DTM imtihoni'):")
    await state.set_state(AdminAddTest.waiting_for_title)

@router.message(AdminAddTest.waiting_for_title)
async def process_test_title(message: Message, state: FSMContext):
    await state.update_data(title=message.text.strip())
    await message.answer("1-asosiy fan nomini kiriting (masalan: Fizika):")
    await state.set_state(AdminAddTest.waiting_for_subject_1)

@router.message(AdminAddTest.waiting_for_subject_1)
async def process_sub_1(message: Message, state: FSMContext):
    await state.update_data(subject_1=message.text.strip())
    await message.answer("2-asosiy fan nomini kiriting (masalan: Ingliz tili):")
    await state.set_state(AdminAddTest.waiting_for_subject_2)

@router.message(AdminAddTest.waiting_for_subject_2)
async def process_sub_2(message: Message, state: FSMContext):
    await state.update_data(subject_2=message.text.strip())
    await message.answer("Sinfni kiriting (masalan: `11-sinf`):")
    await state.set_state(AdminAddTest.waiting_for_grade)

@router.message(AdminAddTest.waiting_for_grade)
async def process_test_grade(message: Message, state: FSMContext):
    await state.update_data(grade=message.text.strip())
    await message.answer("Maksimal urinishlar sonini kiriting (masalan: 1):")
    await state.set_state(AdminAddTest.waiting_for_attempts)

@router.message(AdminAddTest.waiting_for_attempts)
async def process_test_attempts(message: Message, state: FSMContext):
    try: att = int(message.text.strip())
    except: att = 1
    
    data = await state.get_data()
    async with async_session() as session:
        new_test = Test(
            title=data["title"], 
            subject=f"{data['subject_1']} + {data['subject_2']}", 
            grade_level=data["grade"],
            max_attempts=att, 
            is_active=True, 
            is_finished=False
        )
        session.add(new_test)
        await session.flush()
        await session.commit()
        test_id = new_test.id
        
    sub_1_name = data['subject_1']
    sub_2_name = data['subject_2']
    
    await state.update_data(test_id=test_id, current_subject='history', q_count=0, sub_1_name=sub_1_name, sub_2_name=sub_2_name)
    await message.answer(
        f"✅ Blok test yaratildi!\n"
        f"Endi savollarni kiritishni boshlaymiz.\n\n"
        f"1️⃣ **Tarix** fanidan 1-savolni yuboring (Format: Savol matni | A) ... | B) ... | C) ... | D) ... | To'g'ri javob [A/B/C/D])\n"
        f"*(Agar rasm bo'lsa, rasmli xabar sifatida yuborishingiz mumkin)*"
    )
    await state.set_state(AdminAddTest.waiting_for_questions)

@router.message(AdminAddTest.waiting_for_questions)
async def process_question_adding(message: Message, state: FSMContext):
    if not await is_admin(message.from_user.id): return
    data = await state.get_data()
    test_id = data['test_id']
    subject = data['current_subject']
    q_count = data['q_count'] + 1
    
    text = message.text or message.caption or ""
    photo = message.photo[-1].file_id if message.photo else None
    
    try:
        parts = text.split("|")
        q_text = parts[0].strip()
        opt_a = parts[1].strip()
        opt_b = parts[2].strip()
        opt_c = parts[3].strip()
        opt_d = parts[4].strip()
        correct = parts[5].strip().upper()
    except Exception:
        await message.answer("❌ Xato format! Qaytadan quyidagicha yuboring:\n`Savol matni | A) ... | B) ... | C) ... | D) ... | A`", parse_mode="Markdown")
        return

    # Ballarni to'g'ri belgilash
    weights = {'history': 1.1, 'native_lang': 1.1, 'math': 1.1, 'main_1': 3.1, 'main_2': 2.1}
    pts = weights.get(subject, 1.1)

    async with async_session() as session:
        session.add(Question(
            test_id=test_id, subject_type=subject, question_text=q_text, photo_file_id=photo,
            option_a=opt_a, option_b=opt_b, option_c=opt_c, option_d=opt_d, correct_option=correct, points=pts
        ))
        await session.commit()
    
    limits = {'history': 10, 'native_lang': 10, 'math': 10, 'main_1': 30, 'main_2': 30}
    
    if q_count >= limits[subject]:
        if subject == 'history':
            next_sub, sub_name = 'native_lang', "Ona tili"
        elif subject == 'native_lang':
            next_sub, sub_name = 'math', "Matematika"
        elif subject == 'math':
            next_sub, sub_name = 'main_1', data['sub_1_name']
        elif subject == 'main_1':
            next_sub, sub_name = 'main_2', data['sub_2_name']
        else:
            await message.answer("🎉 Tabriklayman! Barcha 90 ta savol muvaffaqiyatli yuklandi.", reply_markup=get_admin_menu())
            await state.clear()
            return
            
        await state.update_data(current_subject=next_sub, q_count=0)
        await message.answer(f"✅ {subject} fanidan savollar tugadi.\n\nEndi **{sub_name}** fanidan 1-savolni yuboring:")
    else:
        await state.update_data(q_count=q_count)
        await message.answer(f"✅ Savol saqlandi ({q_count}/{limits[subject]}). Keyingisini yuboring:")

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
            test.is_finished = True
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
            ts.score_percentage = round((ts.score / (questions_count * 3.1)) * 100, 2)
            
        await session.commit()
        
        student = await session.get(Student, appeal.student_id)
        test = await session.get(Test, ts.test_id)
        
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
            "ID": s.student_id, "Ism": s.first_name, "Familiya": s.last_name, "Yosh": s.age, "Maktab": s.school,
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
    logging.info("🚀 DTM Blok Test & Olimpiada Bot ishga tushdi!")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
