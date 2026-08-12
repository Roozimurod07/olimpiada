import asyncio
import logging
import sys
import random
import os
import re
import json
from datetime import datetime, timedelta, timezone
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

from reportlab.lib.pagesizes import landscape, letter
from reportlab.pdfgen import canvas
from reportlab.lib import colors

import gspread
from google.oauth2.service_account import Credentials

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

class Setting(Base):
    __tablename__ = "settings"
    key = Column(String(50), primary_key=True)
    value = Column(String(50), nullable=False)

class Test(Base):
    __tablename__ = "tests"
    id = Column(Integer, primary_key=True, autoincrement=True)
    title = Column(String(200), nullable=False)
    subject = Column(String(100), nullable=False)
    grade_level = Column(String(20), nullable=False)
    max_attempts = Column(Integer, default=1)
    mode = Column(String(20), default="global_timer")
    duration_minutes = Column(Integer, default=180)
    question_time_seconds = Column(Integer, default=60)
    is_block_test = Column(Boolean, default=False)
    block_subjects = Column(Text, nullable=True)
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
    section_name = Column(String(100), nullable=True)
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
    started_at = Column(DateTime, default=lambda: datetime.now(timezone.utc))
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

class Reminder(Base):
    __tablename__ = "reminders"
    id = Column(Integer, primary_key=True, autoincrement=True)
    test_id = Column(Integer, ForeignKey("tests.id", ondelete="CASCADE"), nullable=False)
    student_id = Column(Integer, ForeignKey("students.id", ondelete="CASCADE"), nullable=False)
    reminded = Column(Boolean, default=False)

class Appeal(Base):
    __tablename__ = "appeals"
    id = Column(Integer, primary_key=True, autoincrement=True)
    student_id = Column(Integer, ForeignKey("students.id", ondelete="CASCADE"), nullable=False)
    test_session_id = Column(Integer, ForeignKey("test_sessions.id", ondelete="CASCADE"), nullable=False)
    question_id = Column(Integer, ForeignKey("questions.id", ondelete="CASCADE"), nullable=False)
    message_text = Column(Text, nullable=False)
    status = Column(String(30), default="PENDING")
    
    student = relationship("Student", back_populates="appeals")

DB_DIR = os.getenv("DB_DIR", ".")
DB_PATH = os.path.join(DB_DIR, "professional_olimpiada.db")
engine = create_async_engine(f"sqlite+aiosqlite:///{DB_PATH}", echo=False)

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
        
        try:
            await conn.execute(sa_text("SELECT question_time_seconds FROM tests LIMIT 1"))
        except Exception:
            try:
                await conn.execute(sa_text("ALTER TABLE tests ADD COLUMN question_time_seconds INTEGER DEFAULT 60;"))
            except Exception:
                pass
        
    async with async_session() as session:
        setting = await session.get(Setting, "blok_test_status")
        if not setting:
            session.add(Setting(key="blok_test_status", value="0"))
            await session.commit()

async def is_admin(user_id: int) -> bool:
    if user_id in SUPER_ADMIN_IDS:
        return True
    async with async_session() as session:
        adm = (await session.execute(select(Admin).where(Admin.telegram_id == user_id))).scalar_one_or_none()
        return adm is not None

async def get_blok_test_status() -> str:
    async with async_session() as session:
        setting = await session.get(Setting, "blok_test_status")
        return setting.value if setting else "0"

async def check_subscription(user_id: int, bot: Bot) -> bool:
    try:
        member = await bot.get_chat_member(chat_id=REQUIRED_CHANNEL, user_id=user_id)
        if member.status in ["member", "administrator", "creator"]:
            return True
    except Exception:
        pass
    return False

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
            str(correct), str(wrong), datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
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
    c.drawString(60, 80, f"Sana: {datetime.now(timezone.utc).strftime('%Y-%m-%d')}")
    c.drawRightString(width - 60, 80, "Tizim rahbarligi: Professional Olimpiada")
    
    c.save()
    pdf_buffer.seek(0)
    return pdf_buffer.read()

class SelfRegState(StatesGroup):
    waiting_for_fullname = State()
    waiting_for_age = State()
    waiting_for_grade = State()
    waiting_for_school = State()

class AdminAddStudent(StatesGroup):
    waiting_for_data = State()

class AdminAddTest(StatesGroup):
    waiting_for_title = State()
    waiting_for_subject = State()
    waiting_for_grade = State()
    waiting_for_is_block = State()
    waiting_for_block_sub1 = State()
    waiting_for_block_sub2 = State()
    waiting_for_question_time = State()
    waiting_for_attempts = State()
    waiting_for_start_time = State()
    waiting_for_questions = State()
    waiting_for_answers = State()

class AdminBroadcast(StatesGroup):
    waiting_for_message = State()

class AdminManageAdmins(StatesGroup):
    waiting_for_id = State()

class AdminSearchStudentState(StatesGroup):
    waiting_for_query = State()

class AppealState(StatesGroup):
    waiting_for_text = State()

class TestProcessState(StatesGroup):
    in_test = State()

router = Router()

async def get_main_menu_keyboard():
    keyboard = [
        [KeyboardButton(text="📝 Testni boshlash")]
    ]
    if await get_blok_test_status() == "1":
        keyboard.append([KeyboardButton(text="🗂 Blok testlar")])
        
    keyboard.extend([
        [KeyboardButton(text="👤 Profilim"), KeyboardButton(text="📊 Mening urinishlarim")],
        [KeyboardButton(text="⚖️ Apellyatsiya"), KeyboardButton(text="🏆 Reyting")],
        [KeyboardButton(text="ℹ️ Olimpiada haqida")]
    ])
    return ReplyKeyboardMarkup(keyboard=keyboard, resize_keyboard=True)

def get_admin_menu():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="➕ ID qo'shish")],
            [KeyboardButton(text="📂 Test yuklash"), KeyboardButton(text="🧩 Blok test yuklash")],
            [KeyboardButton(text="⚙️ Blok test holati"), KeyboardButton(text="⚙️ Testlarni boshqarish")],
            [KeyboardButton(text="🏆 Admin reyting"), KeyboardButton(text="🔍 O'quvchini qidirish")],
            [KeyboardButton(text="📊 Jonli statistika"), KeyboardButton(text="📥 Excel natijalar")],
            [KeyboardButton(text="⚖️ Apellyatsiyalar"), KeyboardButton(text="👥 Adminlar")],
            [KeyboardButton(text="📢 Xabar yuborish"), KeyboardButton(text="🧹 Bazani tozalash")],
            [KeyboardButton(text="⬅️ Bosh menyu")]
        ],
        resize_keyboard=True
    )

def get_finish_test_keyboard():
    return ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text="✅ Testni yakunlash va saqlash")]],
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
            main_menu = await get_main_menu_keyboard()
            await message.answer(f"Xush kelibsiz, <b>{student.first_name} {student.last_name}</b>!\nSinfingiz: <b>{student.grade or 'Nomaʼlum'}</b>", reply_markup=main_menu)
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
            main_menu = await get_main_menu_keyboard()
            await callback.message.answer(f"Xush kelibsiz, <b>{student.first_name} {student.last_name}</b>!", reply_markup=main_menu)
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
    main_menu = await get_main_menu_keyboard()
    await message.answer(f"✅ Muvaffaqiyatli ro'yxatdan o'tdingiz!\nID raqamingiz: <code>{unique_id}</code>", reply_markup=main_menu)

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
            status_text = "🟢 Test yakunlangan" if t.is_finished else "🟡 Test hali davom etmoqda"
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
            sec = f"[{q.section_name}] " if q.section_name else ""
            text += f"<b>{idx}. {sec}{q.question_text}</b>\nSizning javob: <b>{sel}</b> {status} | To'g'ri: <b>{q.correct_option}</b>\n\n"
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
    main_menu = await get_main_menu_keyboard()
    await message.answer("✅ Apellyatsiyangiz adminga yuborildi. Tez orada ko'rib chiqiladi!", reply_markup=main_menu)

@router.message(F.text == "📝 Testni boshlash")
async def start_test_prompt(message: Message, state: FSMContext):
    if await state.get_state() == TestProcessState.in_test.state: return
    async with async_session() as session:
        student = (await session.execute(select(Student).where(Student.telegram_id == message.from_user.id))).scalar_one_or_none()
        if not student or not student.grade:
            await message.answer("❌ Profilingizda sinf ko'rsatilmagan yoki ro'yxatdan o'tmagansiz.")
            return

        tests = (await session.execute(
            select(Test).where(
                Test.is_active == True, 
                Test.is_finished == False, 
                Test.grade_level == student.grade,
                Test.is_block_test == False
            )
        )).scalars().all()
        
        if not tests:
            await message.answer(f"⚠️ Hozirda <b>{student.grade}</b> uchun faol oddiy testlar mavjud emas.")
            return

        keyboard_buttons = [[InlineKeyboardButton(text=f"📚 {t.subject} — {t.title}", callback_data=f"start_test_{t.id}")] for t in tests]
        await message.answer("📝 <b>Mavjud testlar:</b>\n\nTestni tanlang:", reply_markup=InlineKeyboardMarkup(inline_keyboard=keyboard_buttons))

@router.message(F.text == "🗂 Blok testlar")
async def start_block_test_prompt(message: Message, state: FSMContext):
    if await state.get_state() == TestProcessState.in_test.state: return
    
    if await get_blok_test_status() != "1":
        await message.answer("⚠️ Hozirda blok test bo'limi yopiq.")
        return

    async with async_session() as session:
        student = (await session.execute(select(Student).where(Student.telegram_id == message.from_user.id))).scalar_one_or_none()
        if not student or not student.grade:
            await message.answer("❌ Profilingizda sinf ko'rsatilmagan yoki ro'yxatdan o'tmagansiz.")
            return

        tests = (await session.execute(
            select(Test).where(
                Test.is_active == True, 
                Test.is_finished == False, 
                Test.grade_level == student.grade,
                Test.is_block_test == True
            )
        )).scalars().all()
        
        if not tests:
            await message.answer(f"⚠️ Hozirda <b>{student.grade}</b> uchun faol blok testlar mavjud emas.")
            return

        keyboard_buttons = [[InlineKeyboardButton(text=f"🧩 {t.subject} — {t.title}", callback_data=f"start_test_{t.id}")] for t in tests]
        await message.answer("🗂 <b>Mavjud blok testlar:</b>\n\nTestni tanlang:", reply_markup=InlineKeyboardMarkup(inline_keyboard=keyboard_buttons))

@router.callback_query(F.data.startswith("start_test_"))
async def begin_test_session(callback: CallbackQuery, state: FSMContext, bot: Bot):
    test_id = int(callback.data.split("_")[2])
    
    async with async_session() as session:
        student = (await session.execute(select(Student).where(Student.telegram_id == callback.from_user.id))).scalar_one_or_none()
        test = (await session.execute(select(Test).where(Test.id == test_id))).scalar_one_or_none()
        
        if not test or not test.is_active or test.is_finished:
            await callback.answer("Bu test topilmadi yoki yopilgan!", show_alert=True)
            return

        now = datetime.now(timezone.utc)
        if test.start_time:
            start_t = test.start_time.replace(tzinfo=timezone.utc) if test.start_time.tzinfo is None else test.start_time
            if now < start_t:
                await callback.answer(f"⏳ Test hali boshlanmagan! Boshlanish vaqti: {start_t.strftime('%Y-%m-%d %H:%M')}", show_alert=True)
                return
        if test.end_time:
            end_t = test.end_time.replace(tzinfo=timezone.utc) if test.end_time.tzinfo is None else test.end_time
            if now > end_t:
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
            
        test_session = TestSession(student_id=student.id, test_id=test_id, status="IN_PROGRESS")
        session.add(test_session)
        await session.commit()
        await session.refresh(test_session)

        await callback.message.edit_text(f"🚀 <b>{test.subject} ({test.title})</b> testi boshlandi!")
        user_id = callback.from_user.id
        await state.set_state(TestProcessState.in_test)
        
        total_duration_sec = test.duration_minutes * 60
        start_timestamp = datetime.now(timezone.utc)
        
        for index, q in enumerate(questions):
            current_test_check = await session.get(Test, test_id)
            if not current_test_check.is_active or current_test_check.is_finished:
                break

            elapsed = (datetime.now(timezone.utc) - start_timestamp).total_seconds()
            if elapsed >= total_duration_sec:
                await bot.send_message(chat_id=user_id, text="⏰ Test uchun ajratilgan umumiy vaqt tugadi!")
                break

            options = [("A", q.option_a), ("B", q.option_b)]
            if q.option_c: options.append(("C", q.option_c))
            if q.option_d: options.append(("D", q.option_d))
            
            keyboard_buttons = []
            row = []
            for opt_key, text_val in options:
                row.append(InlineKeyboardButton(text=f"{opt_key}) {text_val}", callback_data=f"ans_{test_session.id}_{q.id}_{opt_key}"))
                if len(row) == 2:
                    keyboard_buttons.append(row)
                    row = []
            if row: keyboard_buttons.append(row)
            
            markup = InlineKeyboardMarkup(inline_keyboard=keyboard_buttons)
            
            sec_title = f"<b>Fan / Bo'lim: {q.section_name}</b>\n" if q.section_name else ""
            text_content = f"{sec_title}<b>Savol {index + 1} / {len(questions)}</b> (Ball: {q.points})\n\n{q.question_text}"
            
            if q.photo_file_id:
                q_msg = await bot.send_photo(chat_id=user_id, photo=q.photo_file_id, caption=text_content, reply_markup=markup)
            else:
                q_msg = await bot.send_message(chat_id=user_id, text=text_content, reply_markup=markup)

            q_start_time = datetime.now(timezone.utc)
            question_seconds = max(1, int(test.question_time_seconds))
            timer_msg = await bot.send_message(
                chat_id=user_id,
                text=f"⏳ <b>Qolgan vaqt: {question_seconds} soniya</b>"
            )
            last_second = question_seconds

            user_next_question_flags.pop(user_id, None)

            while True:
                await asyncio.sleep(0.2)
                now = datetime.now(timezone.utc)
                total_elapsed = (now - start_timestamp).total_seconds()
                q_elapsed = (now - q_start_time).total_seconds()

                if total_elapsed >= total_duration_sec:
                    break

                if user_id in user_next_question_flags:
                    user_next_question_flags.pop(user_id, None)
                    break

                seconds_left = max(0, int(question_seconds - q_elapsed + 0.999))
                if seconds_left != last_second:
                    last_second = seconds_left
                    if seconds_left <= 0:
                        break
                    try:
                        await bot.edit_message_text(
                            chat_id=user_id,
                            message_id=timer_msg.message_id,
                            text=f"⏳ <b>Qolgan vaqt: {seconds_left} soniya</b>"
                        )
                    except Exception:
                        pass

                if q_elapsed >= question_seconds:
                    break

            try:
                await bot.delete_message(chat_id=user_id, message_id=timer_msg.message_id)
            except Exception:
                pass
            try:
                await bot.delete_message(chat_id=user_id, message_id=q_msg.message_id)
            except Exception:
                pass
            
            if (datetime.now(timezone.utc) - start_timestamp).total_seconds() >= total_duration_sec:
                break

        async with async_session() as final_session:
            sess = await final_session.get(TestSession, test_session.id)
            if sess and sess.status == "IN_PROGRESS":
                sess.status = "COMPLETED"
                sess.finished_at = datetime.now(timezone.utc)
                await calculate_and_save_results(final_session, sess)
                await final_session.commit()
                
                student_obj = await final_session.get(Student, sess.student_id)
                test_obj = await final_session.get(Test, sess.test_id)
                if student_obj and test_obj:
                    save_result_to_sheet(student_obj.student_id, f"{student_obj.first_name} {student_obj.last_name}", student_obj.age or "-", student_obj.school or "-", student_obj.grade or "-", test_obj.title, test_obj.subject, sess.score, sess.score_percentage, sess.correct_answers, sess.wrong_answers)
                
                await state.clear()
                main_menu = await get_main_menu_keyboard()
                await bot.send_message(chat_id=user_id, text=f"🏆 <b>TEST YAKUNLANDI!</b>\n\nNatijangiz: {sess.score} ball ({sess.score_percentage}%).", reply_markup=main_menu)

user_next_question_flags = {}

@router.callback_query(F.data.startswith("ans_"))
async def save_answer(callback: CallbackQuery, bot: Bot):
    parts = callback.data.split("_")
    session_id, question_id, selected = int(parts[1]), int(parts[2]), parts[3]
    async with async_session() as session:
        existing = (await session.execute(select(Answer).where(Answer.session_id == session_id, Answer.question_id == question_id))).scalar_one_or_none()
        if existing: existing.selected_option = selected
        else: session.add(Answer(session_id=session_id, question_id=question_id, selected_option=selected))
        await session.commit()

    user_next_question_flags[callback.from_user.id] = {"target_index": None}
    await callback.answer(f"Tanlandi: {selected}")

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
    sess.score = round(total_score, 2)
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

@router.message(F.text == "🏆 Admin reyting")
async def admin_rating_menu(message: Message):
    if not await is_admin(message.from_user.id): return
    markup = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🌍 Umumiy reyting", callback_data="adm_rate_global")],
        [InlineKeyboardButton(text="📚 Fanlar bo'yicha reyting", callback_data="adm_rate_subjects")],
        [InlineKeyboardButton(text="🏫 Sinflar bo'yicha reyting", callback_data="adm_rate_grades")]
    ])
    await message.answer("🏆 <b>Admin reyting bo'limi:</b>\n\nReyting turini tanlang:", reply_markup=markup)

@router.callback_query(F.data == "adm_rate_global")
async def admin_global_rating(callback: CallbackQuery):
    async with async_session() as session:
        rows = (await session.execute(
            select(Student, TestSession, Test)
            .join(TestSession, Student.id == TestSession.student_id)
            .join(Test, TestSession.test_id == Test.id)
            .where(TestSession.status == "COMPLETED")
            .order_by(TestSession.score.desc())
            .limit(20)
        )).all()
        
        if not rows:
            await callback.message.edit_text("⚠️ Hozircha yakunlangan natijalar mavjud emas.")
            return
            
        text = "🌍 <b>Umumiy eng yaxshi natijalar reytingi:</b>\n\n"
        for idx, (s, ts, t) in enumerate(rows, 1):
            medal = "🥇" if idx == 1 else ("🥈" if idx == 2 else ("🥉" if idx == 3 else f"{idx}."))
            text += f"{medal} <b>{s.first_name} {s.last_name}</b> ({s.grade})\n📚 {t.subject} | ⭐ Ball: <b>{ts.score}</b> ({ts.score_percentage}%)\n----------------------------------\n"
            
        await callback.message.edit_text(text)
        await callback.answer()

@router.callback_query(F.data == "adm_rate_subjects")
async def admin_subjects_rating_list(callback: CallbackQuery):
    async with async_session() as session:
        subjects = (await session.execute(select(Test.subject).distinct())).scalars().all()
        if not subjects:
            await callback.message.edit_text("⚠️ Fanlar mavjud emas.")
            return
        keyboard = [[InlineKeyboardButton(text=f"📚 {sub}", callback_data=f"adm_subj_rate_{sub}")] for sub in subjects]
        await callback.message.edit_text("📚 Fan bo'yicha reyting ko'rish uchun fanni tanlang:", reply_markup=InlineKeyboardMarkup(inline_keyboard=keyboard))
        await callback.answer()

@router.callback_query(F.data.startswith("adm_subj_rate_"))
async def admin_show_subject_rating(callback: CallbackQuery):
    subject = callback.data.replace("adm_subj_rate_", "")
    async with async_session() as session:
        rows = (await session.execute(
            select(Student, TestSession, Test)
            .join(TestSession, Student.id == TestSession.student_id)
            .join(Test, TestSession.test_id == Test.id)
            .where(Test.subject == subject, TestSession.status == "COMPLETED")
            .order_by(TestSession.score.desc())
            .limit(15)
        )).all()
        
        if not rows:
            await callback.message.edit_text(f"⚠️ {subject} fani bo'yicha natijalar topilmadi.")
            return
            
        text = f"📚 <b>Fan bo'yicha reyting: {subject}</b>\n\n"
        for idx, (s, ts, t) in enumerate(rows, 1):
            medal = "🥇" if idx == 1 else ("🥈" if idx == 2 else ("🥉" if idx == 3 else f"{idx}."))
            text += f"{medal} <b>{s.first_name} {s.last_name}</b> ({s.grade}) — <b>{ts.score} ball</b> ({ts.score_percentage}%)\n"
            
        await callback.message.edit_text(text)
        await callback.answer()

@router.callback_query(F.data == "adm_rate_grades")
async def admin_grades_rating_list(callback: CallbackQuery):
    async with async_session() as session:
        grades = (await session.execute(select(Student.grade).distinct())).scalars().all()
        grades = [g for g in grades if g]
        if not grades:
            await callback.message.edit_text("⚠️ Sinflar mavjud emas.")
            return
        keyboard = [[InlineKeyboardButton(text=f"🏫 {grade}", callback_data=f"adm_grade_rate_{grade}")] for grade in grades]
        await callback.message.edit_text("🏫 Sinf bo'yicha reyting ko'rish uchun sinfni tanlang:", reply_markup=InlineKeyboardMarkup(inline_keyboard=keyboard))
        await callback.answer()

@router.callback_query(F.data.startswith("adm_grade_rate_"))
async def admin_show_grade_rating(callback: CallbackQuery):
    grade = callback.data.replace("adm_grade_rate_", "")
    async with async_session() as session:
        rows = (await session.execute(
            select(Student, TestSession, Test)
            .join(TestSession, Student.id == TestSession.student_id)
            .join(Test, TestSession.test_id == Test.id)
            .where(Student.grade == grade, TestSession.status == "COMPLETED")
            .order_by(TestSession.score.desc())
            .limit(15)
        )).all()
        
        if not rows:
            await callback.message.edit_text(f"⚠️ {grade} sinfi bo'yicha natijalar topilmadi.")
            return
            
        text = f"🏫 <b>Sinf bo'yicha reyting: {grade}</b>\n\n"
        for idx, (s, ts, t) in enumerate(rows, 1):
            medal = "🥇" if idx == 1 else ("🥈" if idx == 2 else ("🥉" if idx == 3 else f"{idx}."))
            text += f"{medal} <b>{s.first_name} {s.last_name}</b> ({t.subject}) — <b>{ts.score} ball</b> ({ts.score_percentage}%)\n"
            
        await callback.message.edit_text(text)
        await callback.answer()

@router.message(F.text == "🔍 O'quvchini qidirish")
async def admin_search_student_prompt(message: Message, state: FSMContext):
    if not await is_admin(message.from_user.id): return
    await state.set_state(AdminSearchStudentState.waiting_for_query)
    await message.answer("🔍 Qidirilayotgan o'quvchining <b>Ism, Familiyasi</b> yoki <b>ID raqamini</b> yuboring:")

@router.message(AdminSearchStudentState.waiting_for_query)
async def admin_perform_student_search(message: Message, state: FSMContext):
    if not await is_admin(message.from_user.id): return
    query = message.text.strip()
    await state.clear()
    
    async with async_session() as session:
        students = (await session.execute(
            select(Student).where(
                (Student.student_id.ilike(f"%{query}%")) |
                (Student.first_name.ilike(f"%{query}%")) |
                (Student.last_name.ilike(f"%{query}%")) |
                ((Student.first_name + " " + Student.last_name).ilike(f"%{query}%"))
            )
        )).scalars().all()
        
        if not students:
            await message.answer("❌ Bunday o'quvchi topilmadi.")
            return
            
        keyboard = []
        for st in students:
            keyboard.append([InlineKeyboardButton(text=f"👤 {st.first_name} {st.last_name} ({st.student_id})", callback_data=f"adm_st_detail_{st.id}")])
            
        await message.answer(f"🔍 Topilgan o'quvchilar ({len(students)} ta):", reply_markup=InlineKeyboardMarkup(inline_keyboard=keyboard))

@router.callback_query(F.data.startswith("adm_st_detail_"))
async def admin_student_details(callback: CallbackQuery):
    student_id = int(callback.data.replace("adm_st_detail_", ""))
    async with async_session() as session:
        student = await session.get(Student, student_id)
        sessions = (await session.execute(
            select(TestSession, Test)
            .join(Test, TestSession.test_id == Test.id)
            .where(TestSession.student_id == student.id, TestSession.status == "COMPLETED")
        )).all()
        
        text = f"👤 <b>O'quvchi ma'lumotlari:</b>\n\n" \
               f"F.I.Sh: <b>{student.first_name} {student.last_name}</b>\n" \
               f"ID: <code>{student.student_id}</code>\n" \
               f"Sinf: {student.grade or '-'}\nMaktab: {student.school or '-'}\n\n" \
               f"📊 <b>Ishlagan testlari:</b>"
               
        if not sessions:
            text += "\nHali test ishlamagan."
            await callback.message.edit_text(text)
            await callback.answer()
            return
            
        keyboard = []
        for ts, t in sessions:
            keyboard.append([InlineKeyboardButton(text=f"📚 {t.subject} ({ts.score} ball - {ts.score_percentage}%)", callback_data=f"adm_sess_analysis_{ts.id}")])
            
        await callback.message.edit_text(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=keyboard))
        await callback.answer()

@router.callback_query(F.data.startswith("adm_sess_analysis_"))
async def admin_session_question_analysis(callback: CallbackQuery):
    session_id = int(callback.data.replace("adm_sess_analysis_", ""))
    async with async_session() as session:
        ts = await session.get(TestSession, session_id)
        test = await session.get(Test, ts.test_id)
        student = await session.get(Student, ts.student_id)
        
        questions = (await session.execute(select(Question).where(Question.test_id == test.id))).scalars().all()
        answers = {a.question_id: a.selected_option for a in (await session.execute(select(Answer).where(Answer.session_id == ts.id))).scalars().all()}
        
        text = f"📋 <b>{student.first_name} {student.last_name} ning test natijasi</b>\n" \
               f"📚 {test.subject} ({test.title})\n" \
               f"⭐ Ball: {ts.score} ({ts.score_percentage}%)\n" \
               f"✅ To'g'ri: {ts.correct_answers} | ❌ Noto'g'ri: {ts.wrong_answers} | ⭕ Javobsiz: {ts.unanswered}\n\n"
               
        for idx, q in enumerate(questions, 1):
            sel = answers.get(q.id, "Javob berilmagan")
            status = "✅" if sel == q.correct_option else "❌"
            sec = f"[{q.section_name}] " if q.section_name else ""
            text += f"<b>{idx}. {sec}{q.question_text}</b>\nUning javobi: <b>{sel}</b> {status} | To'g'ri javob: <b>{q.correct_option}</b>\n\n"
            
        if len(text) > 4000:
            text = text[:3900] + "\n... (matn qisqartirildi)"
            
        await callback.message.edit_text(text)
        await callback.answer()

@router.message(F.text == "ℹ️ Olimpiada haqida")
async def about_handler(message: Message, state: FSMContext):
    if await state.get_state() == TestProcessState.in_test.state: return
    await message.answer("ℹ️ Professional Olimpiada Tizimi v3.1 — Har bir savol uchun vaqt sozlamasi va DTM blok testlar.")

@router.message(Command("admin"))
async def admin_panel(message: Message, state: FSMContext):
    if not await is_admin(message.from_user.id): return
    await state.clear()
    await message.answer("🛠 <b>Admin Panel</b>", reply_markup=get_admin_menu())

@router.message(F.text == "⚙️ Blok test holati")
async def admin_blok_test_settings(message: Message):
    if not await is_admin(message.from_user.id): return
    status = await get_blok_test_status()
    status_text = "🟢 Yoqilgan" if status == "1" else "🔴 O'chirilgan"
    
    markup = InlineKeyboardMarkup(inline_keyboard=[
        [
            InlineKeyboardButton(text="🟢 Yoqish", callback_data="set_blok_1"),
            InlineKeyboardButton(text="🔴 O'chirish", callback_data="set_blok_0")
        ]
    ])
    await message.answer(f"⚙️ Blok test bo'limi hozirgi holati: <b>{status_text}</b>\n\nO'zgartirish uchun tugmani bosing:", reply_markup=markup)

@router.callback_query(F.data.in_(["set_blok_1", "set_blok_0"]))
async def update_blok_test_status(callback: CallbackQuery):
    if not await is_admin(callback.from_user.id): return
    new_val = "1" if callback.data == "set_blok_1" else "0"
    
    async with async_session() as session:
        setting = await session.get(Setting, "blok_test_status")
        if setting:
            setting.value = new_val
        else:
            session.add(Setting(key="blok_test_status", value=new_val))
        await session.commit()
        
    status_text = "🟢 Yoqildi" if new_val == "1" else "🔴 O'chirildi"
    await callback.message.edit_text(f"✅ Blok test holati o'zgartirildi: <b>{status_text}</b>")
    await callback.answer()

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

@router.message(F.text == "📂 Test yuklash")
async def admin_add_test_start(message: Message, state: FSMContext):
    if not await is_admin(message.from_user.id): return
    await state.update_data(is_block=False, duration_minutes=60)
    await state.set_state(AdminAddTest.waiting_for_title)
    await message.answer("📂 Test sarlavhasini kiriting:")

@router.message(F.text == "🧩 Blok test yuklash")
async def admin_add_block_test_start(message: Message, state: FSMContext):
    if not await is_admin(message.from_user.id): return
    await state.update_data(is_block=True, duration_minutes=180)
    await state.set_state(AdminAddTest.waiting_for_title)
    await message.answer("🧩 DTM Blok test sarlavhasini kiriting:")

@router.message(AdminAddTest.waiting_for_title)
async def admin_get_title(message: Message, state: FSMContext):
    await state.update_data(title=message.text.strip())
    data = await state.get_data()
    if data.get("is_block"):
        await state.set_state(AdminAddTest.waiting_for_grade)
        await message.answer("Sinfni kiriting (masalan: `11-sinf`):")
    else:
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
    data = await state.get_data()
    if data.get("is_block"):
        await state.set_state(AdminAddTest.waiting_for_block_sub1)
        await message.answer("1-asosiy fanning nomini kiriting (masalan: Matematika yoki Fizika):")
    else:
        await state.set_state(AdminAddTest.waiting_for_question_time)
        await message.answer("⏱ Har bir test (savol) uchun o'quvchiga beriladigan vaqtni **soniyada** kiriting (masalan: `45` yoki `60`):")

@router.message(AdminAddTest.waiting_for_block_sub1)
async def admin_get_block_sub1(message: Message, state: FSMContext):
    await state.update_data(block_sub1=message.text.strip())
    await state.set_state(AdminAddTest.waiting_for_block_sub2)
    await message.answer("2-asosiy fanning nomini kiriting (masalan: Ingliz tili yoki Kimyo):")

@router.message(AdminAddTest.waiting_for_block_sub2)
async def admin_get_block_sub2(message: Message, state: FSMContext):
    await state.update_data(block_sub2=message.text.strip())
    await state.set_state(AdminAddTest.waiting_for_question_time)
    await message.answer("⏱ Har bir test (savol) uchun o'quvchiga beriladigan vaqtni **soniyada** kiriting (masalan: `60`):")

@router.message(AdminAddTest.waiting_for_question_time)
async def admin_get_question_time(message: Message, state: FSMContext):
    try:
        q_time = int(message.text.strip())
        if q_time <= 0: q_time = 60
    except:
        q_time = 60
    await state.update_data(question_time_seconds=q_time)
    await state.set_state(AdminAddTest.waiting_for_attempts)
    await message.answer("Maksimal urinishlar sonini kiriting (masalan: 1):")

@router.message(AdminAddTest.waiting_for_attempts)
async def admin_get_attempts(message: Message, state: FSMContext):
    try: att = int(message.text.strip())
    except: att = 1
    await state.update_data(max_attempts=att)
    await state.set_state(AdminAddTest.waiting_for_start_time)
    await message.answer("Test boshlanish vaqtini kiriting (Format: `YYYY-MM-DD HH:MM`, masalan: `2026-06-01 10:00` yoki `-`):")

@router.message(AdminAddTest.waiting_for_start_time)
async def admin_get_start_time(message: Message, state: FSMContext):
    txt = message.text.strip()
    start_dt = None
    if txt != "-":
        try:
            start_dt = datetime.strptime(txt, "%Y-%m-%d %H:%M")
        except:
            pass
    await state.update_data(start_time=start_dt, questions=[])
    await state.set_state(AdminAddTest.waiting_for_questions)
    
    data = await state.get_data()
    if data.get("is_block"):
        await message.answer(
            "🧩 <b>Blok test uchun savollarni yuboring:</b>\n"
            "Tartib bo'yicha: \n1) Tarix (10 ta)\n2) Ona tili (10 ta)\n3) Matematika (10 ta)\n4) 1-asosiy fan (30 ta)\n5) 2-asosiy fan (30 ta)\n\n"
            "Matn yoki Word/PDF fayl ko'rinishida yuborishingiz mumkin.", reply_markup=get_finish_test_keyboard()
        )
    else:
        await message.answer("📂 Savollarni matn yoki Word/PDF fayl ko'rinishida yuborishingiz mumkin.", reply_markup=get_finish_test_keyboard())

@router.message(AdminAddTest.waiting_for_questions, F.photo)
async def admin_handle_photo_question(message: Message, state: FSMContext):
    if not await is_admin(message.from_user.id): return
    photo_id = message.photo[-1].file_id
    caption = message.caption or ""
    data = await state.get_data()
    q_list = data.get("questions", [])
    
    parsed = parse_single_question_text(caption)
    if parsed:
        parsed["photo_file_id"] = photo_id
        q_list.append(parsed)
        await state.update_data(questions=q_list)
        await message.answer(f"📸 Rasmli savol qo'shildi! Jami savollar: {len(q_list)}")
    else:
        await message.answer("❌ Rasm tagiga savol va variantlarni (A, B, C, D) to'g'ri formatda yozing!")

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
            for page in reader.pages: extracted_text += (page.extract_text() or "") + "\n"
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

@router.message(AdminAddTest.waiting_for_questions, F.text == "✅ Testni yakunlash va saqlash")
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

def parse_single_question_text(text: str) -> dict:
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    a_idx, b_idx, c_idx, d_idx = -1, -1, -1, -1
    for j in range(len(lines)):
        l = lines[j].lower()
        if l.startswith("a)") or l.startswith("a."): a_idx = j
        elif l.startswith("b)") or l.startswith("b."): b_idx = j
        elif l.startswith("c)") or l.startswith("c."): c_idx = j
        elif l.startswith("d)") or l.startswith("d."): d_idx = j

    if a_idx != -1 and b_idx != -1:
        q_text = " ".join(lines[:a_idx])
        return {
            "text": q_text, 
            "a": re.sub(r'^[aA][\.\)]\s*', '', lines[a_idx]), 
            "b": re.sub(r'^[bB][\.\)]\s*', '', lines[b_idx]),
            "c": re.sub(r'^[cC][\.\)]\s*', '', lines[c_idx]) if c_idx != -1 else "C",
            "d": re.sub(r'^[dD][\.\)]\s*', '', lines[d_idx]) if d_idx != -1 else "D",
            "correct": "A",
            "photo_file_id": None
        }
    return None

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
                "correct": "A",
                "photo_file_id": None
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

    is_block = data.get("is_block", False)
    
    async with async_session() as session:
        new_test = Test(
            title=data["title"], 
            subject="Blok Test" if is_block else data["subject"], 
            grade_level=data["grade"],
            max_attempts=data["max_attempts"], 
            mode="global_timer",
            duration_minutes=180 if is_block else data.get("duration_minutes", 60),
            question_time_seconds=data.get("question_time_seconds", 60),
            is_block_test=is_block,
            block_subjects=json.dumps({"sub1": data.get("block_sub1"), "sub2": data.get("block_sub2")}) if is_block else None,
            start_time=data.get("start_time"),
            end_time=(data.get("start_time") + timedelta(hours=5)) if data.get("start_time") else None,
            is_active=True, 
            is_finished=False
        )
        session.add(new_test)
        await session.flush()
        
        for idx, q in enumerate(questions_list):
            sec_name = None
            points = 1.0
            if is_block:
                if idx < 10:
                    sec_name = "Tarix"
                    points = 1.1
                elif idx < 20:
                    sec_name = "Ona tili"
                    points = 1.1
                elif idx < 30:
                    sec_name = "Matematika"
                    points = 1.1
                elif idx < 60:
                    sec_name = data.get("block_sub1")
                    points = 3.1
                else:
                    sec_name = data.get("block_sub2")
                    points = 2.1
            
            session.add(Question(
                test_id=new_test.id, 
                section_name=sec_name,
                question_text=q["text"],
                photo_file_id=q.get("photo_file_id"),
                option_a=q["a"], option_b=q["b"], option_c=q["c"], option_d=q["d"],
                correct_option=q["correct"],
                points=points
            ))
            
        if data.get("start_time"):
            students = (await session.execute(select(Student).where(Student.grade == data["grade"], Student.telegram_id.is_not(None)))).scalars().all()
            for st in students:
                session.add(Reminder(test_id=new_test.id, student_id=st.id, reminded=False))
                
        await session.commit()
    await state.clear()
    await message.answer("✅ Test muvaffaqiyatli saqlandi va har bir savol uchun vaqt belgilandi!", reply_markup=get_admin_menu())

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

@router.callback_query(F.data == "none")
async def ignore_callback(callback: CallbackQuery):
    await callback.answer()

@router.callback_query(F.data.startswith("toggle_test_"))
async def toggle_test(callback: CallbackQuery):
    async with async_session() as session:
        test = await session.get(Test, int(callback.data.split("_")[2]))
        if test:
            test.is_active = not test.is_active
            await session.commit()
            await callback.answer(f"Test holati o'zgardi! Aktiv: {test.is_active}")
            try:
                tests = (await session.execute(select(Test))).scalars().all()
                keyboard = []
                for t in tests:
                    status_emoji = '🟢 Aktiv' if t.is_active and not t.is_finished else ('🏁 Yakunlangan' if t.is_finished else '🔴 To\'xtatilgan')
                    keyboard.append([InlineKeyboardButton(text=f"[{t.grade_level}] {t.subject} ({status_emoji})", callback_data="none")])
                    keyboard.append([
                        InlineKeyboardButton(text="🔄 Holatni o'zgartirish", callback_data=f"toggle_test_{t.id}"),
                        InlineKeyboardButton(text="🏁 Testni yakunlash", callback_data=f"finish_test_{t.id}"),
                        InlineKeyboardButton(text="🗑 O'chirish", callback_data=f"delete_test_{t.id}")
                    ])
                await callback.message.edit_reply_markup(reply_markup=InlineKeyboardMarkup(inline_keyboard=keyboard))
            except Exception:
                pass

@router.callback_query(F.data.startswith("finish_test_"))
async def finish_test_by_admin(callback: CallbackQuery):
    async with async_session() as session:
        test = await session.get(Test, int(callback.data.split("_")[2]))
        if test:
            test.is_active = False
            test.is_finished = True
            await session.commit()
            await callback.answer("Test to'liq yakunlandi!", show_alert=True)
            try:
                tests = (await session.execute(select(Test))).scalars().all()
                keyboard = []
                for t in tests:
                    status_emoji = '🟢 Aktiv' if t.is_active and not t.is_finished else ('🏁 Yakunlangan' if t.is_finished else '🔴 To\'xtatilgan')
                    keyboard.append([InlineKeyboardButton(text=f"[{t.grade_level}] {t.subject} ({status_emoji})", callback_data="none")])
                    keyboard.append([
                        InlineKeyboardButton(text="🔄 Holatni o'zgartirish", callback_data=f"toggle_test_{t.id}"),
                        InlineKeyboardButton(text="🏁 Testni yakunlash", callback_data=f"finish_test_{t.id}"),
                        InlineKeyboardButton(text="🗑 O'chirish", callback_data=f"delete_test_{t.id}")
                    ])
                await callback.message.edit_reply_markup(reply_markup=InlineKeyboardMarkup(inline_keyboard=keyboard))
            except Exception:
                pass

@router.callback_query(F.data.startswith("delete_test_"))
async def delete_test(callback: CallbackQuery):
    async with async_session() as session:
        test = await session.get(Test, int(callback.data.split("_")[2]))
        if test:
            await session.delete(test)
            await session.commit()
            await callback.answer("Test o'chirildi!")
            try:
                tests = (await session.execute(select(Test))).scalars().all()
                if not tests:
                    await callback.message.edit_text("Testlar yo'q.")
                    return
                keyboard = []
                for t in tests:
                    status_emoji = '🟢 Aktiv' if t.is_active and not t.is_finished else ('🏁 Yakunlangan' if t.is_finished else '🔴 To\'xtatilgan')
                    keyboard.append([InlineKeyboardButton(text=f"[{t.grade_level}] {t.subject} ({status_emoji})", callback_data="none")])
                    keyboard.append([
                        InlineKeyboardButton(text="🔄 Holatni o'zgartirish", callback_data=f"toggle_test_{t.id}"),
                        InlineKeyboardButton(text="🏁 Testni yakunlash", callback_data=f"finish_test_{t.id}"),
                        InlineKeyboardButton(text="🗑 O'chirish", callback_data=f"delete_test_{t.id}")
                    ])
                await callback.message.edit_reply_markup(reply_markup=InlineKeyboardMarkup(inline_keyboard=keyboard))
            except Exception:
                pass

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
        
        update_result_in_sheet(student.student_id, test.title, ts.score, ts.score_percentage, ts.correct_answers)
        
        if student and student.telegram_id:
            try:
                await bot.send_message(
                    chat_id=student.telegram_id,
                    text=f"✅ <b>Apellyatsiyangiz ma'qullandi!</b>\n\nBaligacha 1 ball qo'shildi. Hozirgi balingiz: <b>{ts.score}</b>"
                )
            except Exception: pass
            
    await callback.message.edit_text("✅ Apellyatsiya tasdiqlandi!")

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
        
        keyboard = []
        if admins:
            for adm in admins:
                keyboard.append([
                    InlineKeyboardButton(text=f"👤 {adm.telegram_id} ({adm.role})", callback_data="none"),
                    InlineKeyboardButton(text="❌ O'chirish", callback_data=f"del_admin_{adm.id}")
                ])
        
        keyboard.append([InlineKeyboardButton(text="➕ Admin qo'shish", callback_data="add_admin")])
        
        text = "👥 <b>Moderator adminlar ro'yxati:</b>" if admins else "👥 <b>Moderator adminlar yo'q.</b>"
        await message.answer(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=keyboard))

@router.callback_query(F.data.startswith("del_admin_"))
async def delete_admin(callback: CallbackQuery):
    if callback.from_user.id not in SUPER_ADMIN_IDS:
        await callback.answer("Sizda bu huquq yo'q!", show_alert=True)
        return
    
    admin_db_id = int(callback.data.split("_")[2])
    async with async_session() as session:
        admin_obj = await session.get(Admin, admin_db_id)
        if admin_obj:
            await session.delete(admin_obj)
            await session.commit()
            await callback.answer("✅ Admin muvaffaqiyatli olib tashlandi!", show_alert=True)
        else:
            await callback.answer("❌ Admin topilmadi!", show_alert=True)
            
        admins = (await session.execute(select(Admin))).scalars().all()
        keyboard = []
        if admins:
            for adm in admins:
                keyboard.append([
                    InlineKeyboardButton(text=f"👤 {adm.telegram_id} ({adm.role})", callback_data="none"),
                    InlineKeyboardButton(text="❌ O'chirish", callback_data=f"del_admin_{adm.id}")
                ])
        keyboard.append([InlineKeyboardButton(text="➕ Admin qo'shish", callback_data="add_admin")])
        
        text = "👥 <b>Moderator adminlar ro'yxati:</b>" if admins else "👥 <b>Moderator adminlar yo'q.</b>"
        try:
            await callback.message.edit_text(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=keyboard))
        except Exception:
            pass

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
        result_query = await session.execute(select(Student, TestSession, Test).join(TestSession, Student.id == TestSession.student_id).join(Test, TestSession.test_id == Test.id))
        rows = result_query.all()
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
    await state.clear()
    if await is_admin(message.from_user.id):
        await message.answer("Admin menyu:", reply_markup=get_admin_menu())
    else:
        main_menu = await get_main_menu_keyboard()
        await message.answer("Asosiy menyu:", reply_markup=main_menu)

async def reminder_scheduler(bot: Bot):
    while True:
        await asyncio.sleep(60)
        try:
            now = datetime.now(timezone.utc)
            reminder_target_time = now + timedelta(minutes=15)
            async with async_session() as session:
                reminders = (await session.execute(
                    select(Reminder, Test, Student)
                    .join(Test, Reminder.test_id == Test.id)
                    .join(Student, Reminder.student_id == Student.id)
                    .where(Reminder.reminded == False)
                )).all()
                
                for rem, t, st in reminders:
                    if t.start_time:
                        start_t = t.start_time.replace(tzinfo=timezone.utc) if t.start_time.tzinfo is None else t.start_time
                        if now < start_t <= reminder_target_time:
                            if st.telegram_id:
                                try:
                                    await bot.send_message(
                                        chat_id=st.telegram_id,
                                        text=f"🔔 <b>Eslatma!</b>\n\nSiz ro'yxatdan o'tgan <b>{t.subject} ({t.title})</b> testi 15 daqiqadan so'ng boshlanadi."
                                    )
                                except: pass
                            rem.reminded = True
                await session.commit()
        except Exception as e:
            logging.error(f"Reminder error: {e}")

async def main():
    logging.basicConfig(level=logging.INFO, stream=sys.stdout)
    await init_db()
    bot = Bot(token=BOT_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    dp = Dispatcher(storage=MemoryStorage())
    dp.include_router(router)
    await bot.delete_webhook(drop_pending_updates=True)
    
    asyncio.create_task(reminder_scheduler(bot))
    
    logging.info("🚀 Bot ishga tushdi!")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
