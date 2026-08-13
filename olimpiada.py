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
REQUIRED_CHANNEL = "@diamir_edu"

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
    payments = relationship("Payment", back_populates="student", cascade="all, delete-orphan")

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

class Appeal(Base):
    __tablename__ = "appeals"
    id = Column(Integer, primary_key=True, autoincrement=True)
    student_id = Column(Integer, ForeignKey("students.id", ondelete="CASCADE"), nullable=False)
    test_session_id = Column(Integer, ForeignKey("test_sessions.id", ondelete="CASCADE"), nullable=False)
    question_id = Column(Integer, ForeignKey("questions.id", ondelete="CASCADE"), nullable=False)
    message_text = Column(Text, nullable=False)
    status = Column(String(30), default="PENDING")
    
    student = relationship("Student", back_populates="appeals")
    question = relationship("Question", foreign_keys=[question_id])

class Payment(Base):
    __tablename__ = "payments"
    id = Column(Integer, primary_key=True, autoincrement=True)
    student_id = Column(Integer, ForeignKey("students.id", ondelete="CASCADE"), nullable=False)
    receipt_file_id = Column(String(200), nullable=False)
    status = Column(String(30), default="PENDING") # PENDING, APPROVED, REJECTED
    created_at = Column(DateTime, default=lambda: datetime.now(timezone.utc))

    student = relationship("Student", back_populates="payments")

DB_DIR = os.getenv("DB_DIR", ".")
DB_PATH = os.path.join(DB_DIR, "professional_olimpiada.db")

engine = create_async_engine(
    f"sqlite+aiosqlite:///{DB_PATH}",
    echo=False,
    pool_size=50,
    max_overflow=100,
    pool_timeout=60,
    pool_recycle=1800,
    connect_args={"timeout": 30}
)

@event.listens_for(engine.sync_engine, "connect")
def set_sqlite_pragma(dbapi_connection, connection_record):
    cursor = dbapi_connection.cursor()
    cursor.execute("PRAGMA foreign_keys=ON")
    cursor.execute("PRAGMA journal_mode=WAL")
    cursor.execute("PRAGMA busy_timeout=30000")
    cursor.execute("PRAGMA synchronous=NORMAL")
    cursor.execute("PRAGMA cache_size=-64000")
    cursor.close()

async_session = async_sessionmaker(engine, expire_on_commit=False)

async def init_db():
    async with engine.begin() as conn:
        await conn.execute(sa_text("PRAGMA foreign_keys = ON;"))
        await conn.run_sync(Base.metadata.create_all)
        
    async with async_session() as session:
        for key, val in [("blok_test_status", "0"), ("paid_test_status", "0"), ("card_number", "8600 0000 0000 0000"), ("test_price", "20000 som")]:
            setting = await session.get(Setting, key)
            if not setting:
                session.add(Setting(key=key, value=val))
                await session.commit()

async def is_admin(user_id: int) -> bool:
    if user_id in SUPER_ADMIN_IDS:
        return True
    async with async_session() as session:
        adm = (await session.execute(select(Admin).where(Admin.telegram_id == user_id))).scalar_one_or_none()
        return adm is not None

async def get_setting(key: str) -> str:
    async with async_session() as session:
        setting = await session.get(Setting, key)
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

class EditProfileState(StatesGroup):
    waiting_for_fullname = State()
    waiting_for_age = State()
    waiting_for_grade = State()
    waiting_for_school = State()

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

class AppealState(StatesGroup):
    waiting_for_text = State()

class PaymentState(StatesGroup):
    waiting_for_receipt = State()

class TestProcessState(StatesGroup):
    in_test = State()

router = Router()

async def get_main_menu_keyboard(user_telegram_id: int):
    keyboard = [
        [KeyboardButton(text="📝 Testni boshlash")]
    ]
    if await get_setting("blok_test_status") == "1":
        keyboard.append([KeyboardButton(text="🗂 Blok testlar")])
        
    # Agar pullik test rejimi yoqilgan bo'lsa va o'quvchi to'lov qilmagan bo'lsa
    if await get_setting("paid_test_status") == "1":
        async with async_session() as session:
            student = (await session.execute(select(Student).where(Student.telegram_id == user_telegram_id))).scalar_one_or_none()
            if student:
                payment = (await session.execute(select(Payment).where(Payment.student_id == student.id, Payment.status == "APPROVED"))).scalar_one_or_none()
                if not payment:
                    keyboard.append([KeyboardButton(text="💳 To'lov qilish")])

    keyboard.extend([
        [KeyboardButton(text="👤 Profilim"), KeyboardButton(text="📊 Mening urinishlarim")],
        [KeyboardButton(text="⚖️ Apellyatsiya"), KeyboardButton(text="🏆 Reyting")],
        [KeyboardButton(text="ℹ️ Olimpiada haqida")],
        [KeyboardButton(text="🏠 Asosiy menyu")]
    ])
    return ReplyKeyboardMarkup(keyboard=keyboard, resize_keyboard=True)

def get_admin_menu():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="📂 Test yuklash"), KeyboardButton(text="🧩 Blok test yuklash")],
            [KeyboardButton(text="⚙️ Blok test holati"), KeyboardButton(text="💳 Pullik test sozlamalari")],
            [KeyboardButton(text="⚙️ Testlarni boshqarish"), KeyboardButton(text="🏆 Admin reyting")],
            [KeyboardButton(text="🔍 O'quvchini qidirish"), KeyboardButton(text="📊 Jonli statistika")],
            [KeyboardButton(text="📥 Excel natijalar"), KeyboardButton(text="⚖️ Apellyatsiyalar")],
            [KeyboardButton(text="💰 To'lovlarni tasdiqlash"), KeyboardButton(text="👥 Adminlar")],
            [KeyboardButton(text="📢 Xabar yuborish"), KeyboardButton(text="🧹 Bazani tozalash")],
            [KeyboardButton(text="🏠 Asosiy menyu")]
        ],
        resize_keyboard=True
    )

def get_finish_test_keyboard():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="✅ Testni yakunlash va saqlash")],
            [KeyboardButton(text="🏠 Asosiy menyu")]
        ],
        resize_keyboard=True
    )

@router.message(F.text == "🏠 Asosiy menyu")
async def cmd_main_menu(message: Message, state: FSMContext, bot: Bot):
    await state.clear()
    if await is_admin(message.from_user.id):
        await message.answer("🛠 <b>Admin Asosiy menyusi:</b>", reply_markup=get_admin_menu())
        return
    
    async with async_session() as session:
        student = (await session.execute(select(Student).where(Student.telegram_id == message.from_user.id))).scalar_one_or_none()
        if not student:
            await message.answer("Siz ro'yxatdan o'tmagansiz. /start ni bosing.")
            return
        main_menu = await get_main_menu_keyboard(message.from_user.id)
        await message.answer(f"🏠 <b>Asosiy menyu</b>", reply_markup=main_menu)

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
            main_menu = await get_main_menu_keyboard(message.from_user.id)
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
            main_menu = await get_main_menu_keyboard(callback.from_user.id)
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
    main_menu = await get_main_menu_keyboard(message.from_user.id)
    await message.answer(f"✅ Muvaffaqiyatli ro'yxatdan o'tdingiz!\nID raqamingiz: <code>{unique_id}</code>", reply_markup=main_menu)

# --- PROFIL VA TAHRIRLASH ---
@router.message(F.text == "👤 Profilim")
async def profile_handler(message: Message, state: FSMContext):
    if await state.get_state() == TestProcessState.in_test.state: return
    async with async_session() as session:
        student = (await session.execute(select(Student).where(Student.telegram_id == message.from_user.id))).scalar_one_or_none()
        if not student:
            await message.answer("Siz ro'yxatdan o'tmagansiz. /start ni bosing.")
            return
        
        markup = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="✏️ Profilni tahrirlash", callback_data="edit_profile")]
        ])
        await message.answer(
            f"👤 <b>Profil:</b>\n\n"
            f"ID: <code>{student.student_id}</code>\n"
            f"Ism: {student.first_name} {student.last_name}\n"
            f"Yosh: {student.age or '-'}\n"
            f"Maktab: {student.school or '-'}\n"
            f"Sinf: {student.grade or '-'}",
            reply_markup=markup
        )

@router.callback_query(F.data == "edit_profile")
async def edit_profile_start(callback: CallbackQuery, state: FSMContext):
    await state.set_state(EditProfileState.waiting_for_fullname)
    await callback.message.answer("✏️ Yangi <b>Ism va Familiyangizni</b> kiriting (masalan: Alisherbek Toshmatov):")
    await callback.answer()

@router.message(EditProfileState.waiting_for_fullname)
async def edit_profile_fullname(message: Message, state: FSMContext):
    await state.update_data(fullname=message.text.strip())
    await state.set_state(EditProfileState.waiting_for_age)
    await message.answer("Yoshingizni kiriting:")

@router.message(EditProfileState.waiting_for_age)
async def edit_profile_age(message: Message, state: FSMContext):
    await state.update_data(age=message.text.strip())
    await state.set_state(EditProfileState.waiting_for_grade)
    await message.answer("Sinfingizni kiriting (masalan: 11-sinf):")

@router.message(EditProfileState.waiting_for_grade)
async def edit_profile_grade(message: Message, state: FSMContext):
    await state.update_data(grade=message.text.strip())
    await state.set_state(EditProfileState.waiting_for_school)
    await message.answer("Maktabingizni kiriting:")

@router.message(EditProfileState.waiting_for_school)
async def edit_profile_save(message: Message, state: FSMContext):
    data = await state.get_data()
    fullname_parts = data["fullname"].split(" ", 1)
    first_name = fullname_parts[0]
    last_name = fullname_parts[1] if len(fullname_parts) > 1 else "-"
    
    async with async_session() as session:
        student = (await session.execute(select(Student).where(Student.telegram_id == message.from_user.id))).scalar_one_or_none()
        if student:
            student.first_name = first_name
            student.last_name = last_name
            student.age = data["age"]
            student.grade = data["grade"]
            student.school = message.text.strip()
            await session.commit()
            
    await state.clear()
    main_menu = await get_main_menu_keyboard(message.from_user.id)
    await message.answer("✅ Profilingiz muvaffaqiyatli yangilandi!", reply_markup=main_menu)

# --- PULLIK TEST TO'LOV TIZIMI ---
@router.message(F.text == "💳 To'lov qilish")
async def payment_prompt(message: Message, state: FSMContext):
    card = await get_setting("card_number")
    price = await get_setting("test_price")
    markup = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="✅ To'lov qildim", callback_data="did_payment")]
    ])
    await message.answer(
        f"💳 <b>Pullik testlar uchun to'lov bo'limi</b>\n\n"
        f"💳 Karta raqami: <code>{card}</code>\n"
        f"💰 Narxi: <b>{price}</b>\n\n"
        f"Iltimos, yuqoridagi kartaga to'lovni amalga oshiring va pastdagi tugmani bosing.",
        reply_markup=markup
    )

@router.callback_query(F.data == "did_payment")
async def did_payment_callback(callback: CallbackQuery, state: FSMContext):
    await state.set_state(PaymentState.waiting_for_receipt)
    await callback.message.answer("📸 Iltimos, to'lov chekingizni (rasm yoki PDF formatda) yuboring:")
    await callback.answer()

@router.message(PaymentState.waiting_for_receipt, F.photo | F.document)
async def receive_receipt(message: Message, state: FSMContext, bot: Bot):
    file_id = message.photo[-1].file_id if message.photo else message.document.file_id
    async with async_session() as session:
        student = (await session.execute(select(Student).where(Student.telegram_id == message.from_user.id))).scalar_one_or_none()
        if not student:
            await message.answer("Ro'yxatdan o'tmagansiz.")
            return
        
        payment = Payment(student_id=student.id, receipt_file_id=file_id, status="PENDING")
        session.add(payment)
        await session.commit()
        await session.refresh(payment)
        
        # Adminga xabar yuborish
        admin_markup = InlineKeyboardMarkup(inline_keyboard=[
            [
                InlineKeyboardButton(text="✅ Tasdiqlash", callback_data=f"pay_approve_{payment.id}"),
                InlineKeyboardButton(text="❌ Rad etish", callback_data=f"pay_reject_{payment.id}")
            ]
        ])
        
        for admin_id in SUPER_ADMIN_IDS:
            try:
                await bot.send_photo(
                    chat_id=admin_id,
                    photo=file_id,
                    caption=f"💰 <b>Yangi to'lov cheki!</b>\n\nO'quvchi: {student.first_name} {student.last_name} ({student.grade})\nID: <code>{student.student_id}</code>",
                    reply_markup=admin_markup
                )
            except Exception:
                pass

    await state.clear()
    main_menu = await get_main_menu_keyboard(message.from_user.id)
    await message.answer("✅ Chekingiz adminga yuborildi! Tez orada ko'rib chiqilib xabar beriladi.", reply_markup=main_menu)

@router.callback_query(F.data.startswith("pay_approve_"))
async def approve_payment(callback: CallbackQuery, bot: Bot):
    pay_id = int(callback.data.split("_")[2])
    async with async_session() as session:
        payment = await session.get(Payment, pay_id)
        if payment:
            payment.status = "APPROVED"
            await session.commit()
            student = await session.get(Student, payment.student_id)
            if student and student.telegram_id:
                try:
                    main_menu = await get_main_menu_keyboard(student.telegram_id)
                    await bot.send_message(chat_id=student.telegram_id, text="✅ <b>To'lovingiz tasdiqlandi!</b> Barcha test bo'limlaridan foydalanishingiz mumkin.", reply_markup=main_menu)
                except Exception:
                    pass
    await callback.message.edit_caption(caption=callback.message.caption + "\n\n✅ <b>TASDIQLANDI</b>")
    await callback.answer("To'lov tasdiqlandi!")

@router.callback_query(F.data.startswith("pay_reject_"))
async def reject_payment(callback: CallbackQuery, bot: Bot):
    pay_id = int(callback.data.split("_")[2])
    async with async_session() as session:
        payment = await session.get(Payment, pay_id)
        if payment:
            payment.status = "REJECTED"
            await session.commit()
            student = await session.get(Student, payment.student_id)
            if student and student.telegram_id:
                try:
                    main_menu = await get_main_menu_keyboard(student.telegram_id)
                    await bot.send_message(chat_id=student.telegram_id, text="❌ <b>To'lov qilinmadi!</b> Chekingiz yaroqsiz yoki noto'g'ri. Qaytadan urinib ko'ring.", reply_markup=main_menu)
                except Exception:
                    pass
    await callback.message.edit_caption(caption=callback.message.caption + "\n\n❌ <b>RAD ETILDI</b>")
    await callback.answer("To'lov rad etildi!")

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

# --- APELLYATSIYA (BARCHA TESTLAR VA 90 TA SAVOL UCHUN) ---
@router.message(F.text == "⚖️ Apellyatsiya")
async def student_appeal_menu(message: Message, state: FSMContext):
    if await state.get_state() == TestProcessState.in_test.state: return
    async with async_session() as session:
        student = (await session.execute(select(Student).where(Student.telegram_id == message.from_user.id))).scalar_one_or_none()
        if not student:
            await message.answer("Ro'yxatdan o'tmagansiz.")
            return
        
        # Barcha yakunlangan test sessiyalari (oddiy va blok testlar)
        sessions = (await session.execute(
            select(TestSession, Test)
            .join(Test, TestSession.test_id == Test.id)
            .where(TestSession.student_id == student.id, TestSession.status == "COMPLETED")
            .order_by(TestSession.finished_at.desc())
        )).all()
        
        if not sessions:
            await message.answer("⚠️ Hozirda apellyatsiya berish uchun yakunlangan testlar mavjud emas.")
            return
            
        keyboard_buttons = []
        for ts, t in sessions:
            keyboard_buttons.append([InlineKeyboardButton(
                text=f"📚 {t.subject} | {ts.score} ball ({ts.score_percentage}%)",
                callback_data=f"student_appeal_test_{ts.id}"
            )])
            
        markup = InlineKeyboardMarkup(inline_keyboard=keyboard_buttons)
        await message.answer("⚖️ <b>Apellyatsiya bo'limi:</b>\n\nSavollar tahlili va istalgan savolga apellyatsiya berish uchun testni tanlang:", reply_markup=markup)

@router.callback_query(F.data.startswith("student_appeal_test_"))
async def show_test_analysis_for_appeal(callback: CallbackQuery):
    session_id = int(callback.data.split("_")[3])
    async with async_session() as session:
        ts = await session.get(TestSession, session_id)
        test = await session.get(Test, ts.test_id)

        questions = (await session.execute(select(Question).where(Question.test_id == test.id).order_by(Question.id))).scalars().all()
        answers = {a.question_id: a.selected_option for a in (await session.execute(select(Answer).where(Answer.session_id == ts.id))).scalars().all()}
        
        keyboard = []
        for idx, q in enumerate(questions, 1):
            sel = answers.get(q.id, "Javob berilmagan")
            status = "✅" if sel == q.correct_option else "❌"
            sec = f"[{q.section_name}] " if q.section_name else ""
            keyboard.append([InlineKeyboardButton(text=f"⚖️ {idx}-savolga ({sec}) apellyatsiya", callback_data=f"appeal_q_{ts.id}_{q.id}")])
            
        keyboard.append([InlineKeyboardButton(text="🎓 Sertifikatni yuklab olish", callback_data=f"get_cert_{ts.id}")])
        
        await callback.message.answer(
            f"📋 <b>Test tahlili: {test.subject} ({test.title})</b>\n"
            f"⭐ Ball: {ts.score} ({ts.score_percentage}%)\n"
            f"✅ To'g'ri: {ts.correct_answers} | ❌ Noto'g'ri: {ts.wrong_answers} | ⭕ Javobsiz: {ts.unanswered}\n\n"
            f"Quyidagi tugmalar orqali istalgan savolga apellyatsiya yuborishingiz mumkin:",
            reply_markup=InlineKeyboardMarkup(inline_keyboard=keyboard)
        )
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
    main_menu = await get_main_menu_keyboard(message.from_user.id)
    await message.answer("✅ Apellyatsiyangiz adminga yuborildi. Tez orada ko'rib chiqiladi!", reply_markup=main_menu)

# --- TESTLARNI BOSHLASH VA ISHLASH ---
@router.message(F.text == "📝 Testni boshlash")
async def start_test_prompt(message: Message, state: FSMContext):
    if await state.get_state() == TestProcessState.in_test.state: return
    
    # Pullik test tekshiruvi
    if await get_setting("paid_test_status") == "1":
        async with async_session() as session:
            student = (await session.execute(select(Student).where(Student.telegram_id == message.from_user.id))).scalar_one_or_none()
            if student:
                payment = (await session.execute(select(Payment).where(Payment.payment_id == student.id if hasattr(Payment, 'payment_id') else Payment.student_id == student.id, Payment.status == "APPROVED"))).scalar_one_or_none()
                if not payment:
                    await message.answer("⚠️ Bu testlar pullik qilib belgilangan. Iltimos, avval '💳 To'lov qilish' bo'limi orqali to'lovni amalga oshiring.")
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
    
    if await get_setting("blok_test_status") != "1":
        await message.answer("⚠️ Hozirda blok test bo'limi yopiq.")
        return

    # Pullik test tekshiruvi
    if await get_setting("paid_test_status") == "1":
        async with async_session() as session:
            student = (await session.execute(select(Student).where(Student.telegram_id == message.from_user.id))).scalar_one_or_none()
            if student:
                payment = (await session.execute(select(Payment).where(Payment.student_id == student.id, Payment.status == "APPROVED"))).scalar_one_or_none()
                if not payment:
                    await message.answer("⚠️ Bu blok testlar pullik qilib belgilangan. Iltimos, avval '💳 To'lov qilish' bo'limi orqali to'lovni amalga oshiring.")
                    return

    async with async_session() as session:
        student = (await session.execute(select(Student).where(Student.telegram_id == message.from_user.id))).scalar_one_or_none()
        if not student:
            await message.answer("❌ Ro'yxatdan o'tmagansiz. /start ni bosing.")
            return

        tests = (await session.execute(
            select(Test).where(
                Test.is_active == True, 
                Test.is_finished == False, 
                Test.is_block_test == True
            )
        )).scalars().all()
        
        if not tests:
            await message.answer("⚠️ Hozirda faol blok testlar mavjud emas.")
            return

        keyboard_buttons = [[InlineKeyboardButton(text=f"🧩 {t.subject} — {t.title}", callback_data=f"start_block_{t.id}")] for t in tests]
        await message.answer("🗂 <b>Mavjud blok testlar:</b>\n\nTestni tanlang:", reply_markup=InlineKeyboardMarkup(inline_keyboard=keyboard_buttons))

BLOCK_SECTION_KEYS = ["Tarix", "Ona_tili", "Matematika", "sub1", "sub2"]
BLOCK_SECTION_DISPLAY = {
    "Tarix": "🏛 Tarix",
    "Ona_tili": "🇺🇿 Ona tili",
    "Matematika": "🔢 Matematika",
}

def _section_variants(section: str, block_subs: dict = None):
    variants = [section]
    if section == "Ona_tili":
        variants.extend(["Ona tili", "Ona_tili"])
    elif section in ("sub1", "sub2") and block_subs:
        real = block_subs.get(section)
        if real:
            variants.append(real)
    return list(dict.fromkeys(variants))

async def _load_section_questions(session, test_id: int, section: str):
    test = await session.get(Test, test_id)
    block_subs = {}
    if test and test.block_subjects:
        try:
            block_subs = json.loads(test.block_subjects)
        except Exception:
            pass
    variants = _section_variants(section, block_subs)
    questions = list((await session.execute(
        select(Question).where(
            Question.test_id == test_id,
            Question.section_name.in_(variants)
        ).order_by(Question.id)
    )).scalars().all())
    return questions

async def get_block_section_progress(session, test_id: int, ts_id: int, section: str):
    questions = await _load_section_questions(session, test_id, section)
    if not questions:
        return 0, 0, []
    answered_ids = set(
        a.question_id for a in (await session.execute(
            select(Answer).where(Answer.session_id == ts_id, Answer.selected_option.is_not(None))
        )).scalars().all()
    )
    answered_count = sum(1 for q in questions if q.id in answered_ids)
    return len(questions), answered_count, questions

async def build_block_subjects_keyboard(test_id: int, block_subs: dict, progress: dict, remaining_text: str, test_title: str):
    def btn_text(key, default_name, count_label):
        done = progress.get(key, (0, 0))
        total, ans = done[0], done[1]
        mark = " ✅" if total > 0 and ans >= total else (f" ({ans}/{total})" if total > 0 else "")
        name = BLOCK_SECTION_DISPLAY.get(key, default_name)
        return f"{name} {count_label}{mark}"

    keyboard = [
        [InlineKeyboardButton(text=btn_text("Tarix", "Tarix", "(10 ta · 1.1)"), callback_data=f"run_block_{test_id}_Tarix")],
        [InlineKeyboardButton(text=btn_text("Ona_tili", "Ona tili", "(10 ta · 1.1)"), callback_data=f"run_block_{test_id}_Ona_tili")],
        [InlineKeyboardButton(text=btn_text("Matematika", "Matematika", "(10 ta · 1.1)"), callback_data=f"run_block_{test_id}_Matematika")],
        [InlineKeyboardButton(text=btn_text("sub1", block_subs.get("sub1", "Asosiy 1"), "(30 ta · 3.1)"), callback_data=f"run_block_{test_id}_sub1")],
        [InlineKeyboardButton(text=btn_text("sub2", block_subs.get("sub2", "Asosiy 2"), "(30 ta · 2.1)"), callback_data=f"run_block_{test_id}_sub2")],
        [InlineKeyboardButton(text="🚀 Testni yakunlash va topshirish", callback_data=f"finish_block_session_{test_id}")]
    ]
    text = (
        f"🧩 <b>{test_title}</b>\n\n"
        f"⏱ Umumiy vaqt: <b>180 daqiqa</b>\n"
        f"⏳ Qolgan vaqt: <b>{remaining_text}</b>\n\n"
        "Istalgan fandan boshlang. Fan tugagach ✅ belgilanadi."
    )
    return text, InlineKeyboardMarkup(inline_keyboard=keyboard)

async def send_next_block_question(bot: Bot, chat_id: int, ts_id: int, test_id: int, section: str, state: FSMContext):
    async with async_session() as session:
        ts = await session.get(TestSession, ts_id)
        if not ts or ts.status != "IN_PROGRESS":
            return True

        started = ts.started_at
        if started.tzinfo is None:
            started = started.replace(tzinfo=timezone.utc)
        if (datetime.now(timezone.utc) - started).total_seconds() >= 180 * 60:
            return "TIME_UP"

        questions = await _load_section_questions(session, test_id, section)
        if not questions:
            return True

        answered = {
            a.question_id: a.selected_option
            for a in (await session.execute(select(Answer).where(Answer.session_id == ts_id))).scalars().all()
            if a.selected_option
        }

        next_q = None
        next_idx = 0
        for i, q in enumerate(questions):
            if q.id not in answered:
                next_q = q
                next_idx = i + 1
                break

        if not next_q:
            return True

        display_name = BLOCK_SECTION_DISPLAY.get(section, section)
        if section in ("sub1", "sub2"):
            test = await session.get(Test, test_id)
            block_subs = json.loads(test.block_subjects) if test and test.block_subjects else {}
            display_name = f"📘 {block_subs.get(section, section)}" if section == "sub1" else f"📙 {block_subs.get(section, section)}"

        options = [("A", next_q.option_a), ("B", next_q.option_b)]
        if next_q.option_c: options.append(("C", next_q.option_c))
        if next_q.option_d: options.append(("D", next_q.option_d))

        keyboard_buttons = []
        row = []
        for opt_key, text_val in options:
            row.append(InlineKeyboardButton(text=f"{opt_key}) {text_val}", callback_data=f"b_ans_{ts_id}_{next_q.id}_{opt_key}_{section}_{test_id}"))
            if len(row) == 2:
                keyboard_buttons.append(row)
                row = []
        if row: keyboard_buttons.append(row)
        keyboard_buttons.append([InlineKeyboardButton(text="⬅️ Fanlar menyusiga qaytish", callback_data=f"back_to_block_menu_{test_id}")])
        markup = InlineKeyboardMarkup(inline_keyboard=keyboard_buttons)

        text_content = (
            f"📚 <b>Fan: {display_name}</b>\n"
            f"<b>Savol {next_idx} / {len(questions)}</b> (Ball: {next_q.points})\n\n"
            f"{next_q.question_text}"
        )

        await state.set_state(TestProcessState.in_test)
        try:
            if next_q.photo_file_id:
                await bot.send_photo(chat_id=chat_id, photo=next_q.photo_file_id, caption=text_content, reply_markup=markup)
            else:
                await bot.send_message(chat_id=chat_id, text=text_content, reply_markup=markup)
        except Exception:
            await bot.send_message(chat_id=chat_id, text=text_content, reply_markup=markup)
        return False

@router.callback_query(F.data.startswith("start_block_"))
async def choose_block_subject_menu(callback: CallbackQuery, state: FSMContext):
    test_id = int(callback.data.split("_")[2])
    async with async_session() as session:
        student = (await session.execute(select(Student).where(Student.telegram_id == callback.from_user.id))).scalar_one_or_none()
        test = await session.get(Test, test_id)
        if not test or not test.is_active or test.is_finished:
            await callback.answer("Bu test topilmadi yoki yopilgan!", show_alert=True)
            return

        ts = (await session.execute(select(TestSession).where(
            TestSession.student_id == student.id,
            TestSession.test_id == test_id,
            TestSession.status == "IN_PROGRESS"
        ))).scalar_one_or_none()
        if not ts:
            ts = TestSession(student_id=student.id, test_id=test_id, status="IN_PROGRESS")
            session.add(ts)
            await session.commit()
            await session.refresh(ts)

        started = ts.started_at
        if started.tzinfo is None:
            started = started.replace(tzinfo=timezone.utc)
        elapsed_min = (datetime.now(timezone.utc) - started).total_seconds() / 60
        remaining_min = max(0, 180 - elapsed_min)

        block_subs = json.loads(test.block_subjects) if test.block_subjects else {"sub1": "1-Asosiy fan", "sub2": "2-Asosiy fan"}
        progress = {}
        for key in BLOCK_SECTION_KEYS:
            total, ans, _ = await get_block_section_progress(session, test_id, ts.id, key)
            progress[key] = (total, ans)

        remaining_text = f"{int(remaining_min)} daqiqa"
        text, markup = await build_block_subjects_keyboard(test_id, block_subs, progress, remaining_text, test.title)

    await callback.message.edit_text(text, reply_markup=markup)
    await callback.answer()

@router.callback_query(F.data.startswith("run_block_"))
async def run_block_subject_questions(callback: CallbackQuery, state: FSMContext, bot: Bot):
    parts = callback.data.split("_")
    test_id = int(parts[2])
    section = "_".join(parts[3:])

    async with async_session() as session:
        student = (await session.execute(select(Student).where(Student.telegram_id == callback.from_user.id))).scalar_one_or_none()
        ts = (await session.execute(select(TestSession).where(
            TestSession.student_id == student.id,
            TestSession.test_id == test_id,
            TestSession.status == "IN_PROGRESS"
        ))).scalar_one_or_none()

        if not ts:
            ts = TestSession(student_id=student.id, test_id=test_id, status="IN_PROGRESS")
            session.add(ts)
            await session.commit()
            await session.refresh(ts)

        total, ans, questions = await get_block_section_progress(session, test_id, ts.id, section)
        if total == 0:
            await callback.answer("Bu fanda savollar topilmadi!", show_alert=True)
            return
        if ans >= total:
            await callback.answer("Bu fan allaqachon to'liq bajarilgan! ✅", show_alert=True)
            return
        ts_id = ts.id

    try:
        await callback.message.delete()
    except Exception:
        pass

    display = BLOCK_SECTION_DISPLAY.get(section, section)
    await bot.send_message(callback.from_user.id, f"📚 <b>{display}</b> fani boshlandi.")
    await send_next_block_question(bot, callback.from_user.id, ts_id, test_id, section, state)
    await callback.answer()

@router.callback_query(F.data.startswith("back_to_block_menu_"))
async def return_to_block_menu(callback: CallbackQuery, state: FSMContext):
    test_id = int(callback.data.split("_")[4])
    async with async_session() as session:
        student = (await session.execute(select(Student).where(Student.telegram_id == callback.from_user.id))).scalar_one_or_none()
        test = await session.get(Test, test_id)
        block_subs = json.loads(test.block_subjects) if test.block_subjects else {"sub1": "1-Asosiy fan", "sub2": "2-Asosiy fan"}

        ts = (await session.execute(select(TestSession).where(
            TestSession.student_id == student.id,
            TestSession.test_id == test_id,
            TestSession.status == "IN_PROGRESS"
        ))).scalar_one_or_none()

        remaining_text = "180 daqiqa"
        progress = {}
        if ts:
            started = ts.started_at
            if started.tzinfo is None:
                started = started.replace(tzinfo=timezone.utc)
            elapsed_min = (datetime.now(timezone.utc) - started).total_seconds() / 60
            remaining_min = max(0, 180 - elapsed_min)
            remaining_text = f"{int(remaining_min)} daqiqa"
            for key in BLOCK_SECTION_KEYS:
                total, ans, _ = await get_block_section_progress(session, test_id, ts.id, key)
                progress[key] = (total, ans)

        text, markup = await build_block_subjects_keyboard(test_id, block_subs, progress, remaining_text, test.title)

    await state.clear()
    try:
        await callback.message.delete()
    except Exception:
        pass
    await callback.message.answer(text, reply_markup=markup)
    await callback.answer()

@router.callback_query(F.data.startswith("b_ans_"))
async def save_block_answer(callback: CallbackQuery, state: FSMContext, bot: Bot):
    parts = callback.data.split("_")
    session_id = int(parts[2])
    question_id = int(parts[3])
    selected = parts[4]
    rest = parts[5:]
    test_id = int(rest[-1])
    section = "_".join(rest[:-1])

    async with async_session() as session:
        ts = await session.get(TestSession, session_id)
        if not ts or ts.status != "IN_PROGRESS":
            await callback.answer("Test yakunlangan!", show_alert=True)
            return
        existing = (await session.execute(
            select(Answer).where(Answer.session_id == session_id, Answer.question_id == question_id)
        )).scalar_one_or_none()
        if existing:
            existing.selected_option = selected
        else:
            session.add(Answer(session_id=session_id, question_id=question_id, selected_option=selected))
        await session.commit()

    await callback.answer(f"✅ {selected}")
    try:
        await callback.message.delete()
    except Exception:
        pass

    result = await send_next_block_question(bot, callback.from_user.id, session_id, test_id, section, state)
    if result is True:
        await bot.send_message(
            callback.from_user.id,
            f"✅ Fani to'liq bajarildi!",
            reply_markup=InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text="📋 Fanlar menyusiga qaytish", callback_data=f"back_to_block_menu_{test_id}")]
            ])
        )
        await state.clear()

@router.callback_query(F.data.startswith("finish_block_session_"))
async def finish_block_test_by_student(callback: CallbackQuery, state: FSMContext, bot: Bot):
    test_id = int(callback.data.split("_")[3])
    async with async_session() as session:
        student = (await session.execute(select(Student).where(Student.telegram_id == callback.from_user.id))).scalar_one_or_none()
        ts = (await session.execute(select(TestSession).where(
            TestSession.student_id == student.id,
            TestSession.test_id == test_id,
            TestSession.status == "IN_PROGRESS"
        ))).scalar_one_or_none()
        
        if ts:
            ts.status = "COMPLETED"
            ts.finished_at = datetime.now(timezone.utc)
            await calculate_and_save_results(session, ts)
            await session.commit()
            
            test_obj = await session.get(Test, test_id)
            save_result_to_sheet(student.student_id, f"{student.first_name} {student.last_name}", student.age or "-", student.school or "-", student.grade or "-", test_obj.title, test_obj.subject, ts.score, ts.score_percentage, ts.correct_answers, ts.wrong_answers)
            
    await state.clear()
    main_menu = await get_main_menu_keyboard(callback.from_user.id)
    await callback.message.answer(f"🏆 <b>BLOK TEST YAKUNLANDI!</b>\n\nSiz to'plagan ball: <b>{ts.score}</b> ({ts.score_percentage}%).", reply_markup=main_menu)
    await callback.answer()

@router.callback_query(F.data.startswith("start_test_"))
async def begin_test_session(callback: CallbackQuery, state: FSMContext, bot: Bot):
    test_id = int(callback.data.split("_")[2])
    
    async with async_session() as session:
        student = (await session.execute(select(Student).where(Student.telegram_id == callback.from_user.id))).scalar_one_or_none()
        test = (await session.execute(select(Test).where(Test.id == test_id))).scalar_one_or_none()
        
        if not test or not test.is_active or test.is_finished:
            await callback.answer("Bu test topilmadi yoki yopilgan!", show_alert=True)
            return

        questions = list((await session.execute(select(Question).where(Question.test_id == test_id))).scalars().all())
        if not questions:
            await callback.answer("Bu testda savollar yo'q!", show_alert=True)
            return
        
        random.shuffle(questions)
            
        test_session = TestSession(student_id=student.id, test_id=test_id, status="IN_PROGRESS")
        session.add(test_session)
        await session.commit()
        await session.refresh(test_session)

        await callback.message.edit_text(f"🚀 <b>{test.subject} ({test.title})</b> testi boshlandi!")
        user_id = callback.from_user.id
        await state.set_state(TestProcessState.in_test)
        
        question_seconds = max(5, int(test.question_time_seconds))
        total_duration_sec = len(questions) * question_seconds + 60
        start_timestamp = datetime.now(timezone.utc)
        
        for index, q in enumerate(questions):
            elapsed = (datetime.now(timezone.utc) - start_timestamp).total_seconds()
            if elapsed >= total_duration_sec:
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
            text_content = f"<b>Savol {index + 1} / {len(questions)}</b> (Ball: {q.points})\n\n{q.question_text}"
            
            if q.photo_file_id:
                q_msg = await bot.send_photo(chat_id=user_id, photo=q.photo_file_id, caption=text_content, reply_markup=markup)
            else:
                q_msg = await bot.send_message(chat_id=user_id, text=text_content, reply_markup=markup)

            q_start_time = datetime.now(timezone.utc)
            timer_msg = await bot.send_message(chat_id=user_id, text=f"⏳ <b>Qolgan vaqt: {question_seconds} soniya</b>")
            last_second = question_seconds

            user_next_question_flags.pop(user_id, None)

            while True:
                await asyncio.sleep(0.2)
                now = datetime.now(timezone.utc)
                q_elapsed = (now - q_start_time).total_seconds()

                if user_id in user_next_question_flags:
                    user_next_question_flags.pop(user_id, None)
                    break

                seconds_left = max(0, int(question_seconds - q_elapsed + 0.999))
                if seconds_left != last_second:
                    last_second = seconds_left
                    if seconds_left <= 0:
                        break
                    try:
                        await bot.edit_message_text(chat_id=user_id, message_id=timer_msg.message_id, text=f"⏳ <b>Qolgan vaqt: {seconds_left} soniya</b>")
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
                main_menu = await get_main_menu_keyboard(user_id)
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

    user_next_question_flags[callback.from_user.id] = True
    await callback.answer(f"Tanlandi: {selected}")

async def calculate_and_save_results(session, sess: TestSession):
    questions = (await session.execute(select(Question).where(Question.test_id == sess.test_id))).scalars().all()
    answers = {a.question_id: a.selected_option for a in (await session.execute(select(Answer).where(Answer.session_id == sess.id))).scalars().all()}
    
    correct, wrong, unanswered, total_score, max_possible = 0, 0, 0, 0.0, 0.0
    for q in questions:
        max_possible += q.points
        sel = answers.get(q.id)
        if not sel: unanswered += 1
        elif sel == q.correct_option:
            correct += 1
            total_score += q.points
        else: wrong += 1
            
    sess.correct_answers = correct
    sess.wrong_answers = wrong
    sess.unanswered = unanswered
    sess.score = round(total_score, 2)
    sess.score_percentage = round((total_score / max_possible) * 100, 2) if max_possible > 0 else 0.0

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
    await message.answer("ℹ️ Professional Olimpiada Tizimi v3.3 — DTM blok testlar va pullik/bepul imtihon rejimi.")

@router.message(Command("admin"))
async def admin_panel(message: Message, state: FSMContext):
    if not await is_admin(message.from_user.id): return
    await state.clear()
    await message.answer("🛠 <b>Admin Panel</b>", reply_markup=get_admin_menu())

@router.message(F.text == "⚙️ Blok test holati")
async def admin_blok_test_settings(message: Message):
    if not await is_admin(message.from_user.id): return
    status = await get_setting("blok_test_status")
    status_text = "🟢 Yoqilgan" if status == "1" else "🔴 O'chirilgan"
    markup = InlineKeyboardMarkup(inline_keyboard=[[
        InlineKeyboardButton(text="🟢 Yoqish", callback_data="set_blok_1"),
        InlineKeyboardButton(text="🔴 O'chirish", callback_data="set_blok_0")
    ]])
    await message.answer(f"⚙️ Blok test holati: <b>{status_text}</b>", reply_markup=markup)

@router.callback_query(F.data.in_(["set_blok_1", "set_blok_0"]))
async def update_blok_test_status(callback: CallbackQuery):
    if not await is_admin(callback.from_user.id): return
    new_val = "1" if callback.data == "set_blok_1" else "0"
    async with async_session() as session:
        setting = await session.get(Setting, "blok_test_status")
        if setting: setting.value = new_val
        else: session.add(Setting(key="blok_test_status", value=new_val))
        await session.commit()
    await callback.message.edit_text(f"✅ Blok test holati: <b>{'Yoqildi' if new_val == '1' else 'O'chirildi'}</b>")
    await callback.answer()

@router.message(F.text == "💳 Pullik test sozlamalari")
async def admin_paid_test_settings(message: Message):
    if not await is_admin(message.from_user.id): return
    status = await get_setting("paid_test_status")
    card = await get_setting("card_number")
    price = await get_setting("test_price")
    status_text = "🟢 Yoqilgan (Pullik)" if status == "1" else "🔴 O'chirilgan (Bepul)"
    
    markup = InlineKeyboardMarkup(inline_keyboard=[
        [
            InlineKeyboardButton(text="🟢 Pullik qilish", callback_data="set_paid_1"),
            InlineKeyboardButton(text="🔴 Bepul qilish", callback_data="set_paid_0")
        ]
    ])
    await message.answer(f"💳 <b>Pullik test rejimi:</b> {status_text}\n💳 Karta: <code>{card}</code>\n💰 Narx: <b>{price}</b>", reply_markup=markup)

@router.callback_query(F.data.in_(["set_paid_1", "set_paid_0"]))
async def update_paid_test_status(callback: CallbackQuery):
    if not await is_admin(callback.from_user.id): return
    new_val = "1" if callback.data == "set_paid_1" else "0"
    async with async_session() as session:
        setting = await session.get(Setting, "paid_test_status")
        if setting: setting.value = new_val
        else: session.add(Setting(key="paid_test_status", value=new_val))
        await session.commit()
    await callback.message.edit_text(f"✅ Pullik test rejimi o'zgardi: <b>{'Pullik' if new_val == '1' else 'Bepul'}</b>")
    await callback.answer()

@router.message(F.text == "💰 To'lovlarni tasdiqlash")
async def admin_view_pending_payments(message: Message):
    if not await is_admin(message.from_user.id): return
    async with async_session() as session:
        payments = (await session.execute(select(Payment, Student).join(Student, Payment.student_id == Student.id).where(Payment.status == "PENDING"))).all()
        if not payments:
            await message.answer("⚠️ Tasdiqlanmagan to'lovlar yo'q.")
            return
        
        for pay, st in payments:
            markup = InlineKeyboardMarkup(inline_keyboard=[
                [
                    InlineKeyboardButton(text="✅ Tasdiqlash", callback_data=f"pay_approve_{pay.id}"),
                    InlineKeyboardButton(text="❌ Rad etish", callback_data=f"pay_reject_{pay.id}")
                ]
            ])
            await message.bot.send_photo(
                chat_id=message.from_user.id,
                photo=pay.receipt_file_id,
                caption=f"👤 O'quvchi: {st.first_name} {st.last_name} ({st.grade})\nID: <code>{st.student_id}</code>",
                reply_markup=markup
            )

# --- TEST YUKLASH VA RASMLI/SONLAR MOSLIGINI TEKSHIRISH ---
@router.message(F.text == "📂 Test yuklash")
async def admin_add_test_start(message: Message, state: FSMContext):
    if not await is_admin(message.from_user.id): return
    await state.update_data(is_block=False, duration_minutes=60, questions=[])
    await state.set_state(AdminAddTest.waiting_for_title)
    await message.answer("📂 Test sarlavhasini kiriting:")

@router.message(F.text == "🧩 Blok test yuklash")
async def admin_add_block_test_start(message: Message, state: FSMContext):
    if not await is_admin(message.from_user.id): return
    await state.update_data(is_block=True, duration_minutes=180, questions=[])
    await state.set_state(AdminAddTest.waiting_for_title)
    await message.answer("🧩 DTM Blok test sarlavhasini kiriting:")

@router.message(AdminAddTest.waiting_for_title)
async def admin_get_title(message: Message, state: FSMContext):
    await state.update_data(title=message.text.strip())
    data = await state.get_data()
    if data.get("is_block"):
        await state.update_data(grade="Barcha")
        await state.set_state(AdminAddTest.waiting_for_block_sub1)
        await message.answer("1-asosiy fanning nomini kiriting:")
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
        await message.answer("1-asosiy fanning nomini kiriting:")
    else:
        await state.set_state(AdminAddTest.waiting_for_question_time)
        await message.answer("⏱ Har bir savol uchun vaqtni **soniyada** kiriting (masalan: `60`):")

@router.message(AdminAddTest.waiting_for_question_time)
async def admin_get_question_time(message: Message, state: FSMContext):
    try: q_time = int(message.text.strip())
    except: q_time = 60
    await state.update_data(question_time_seconds=q_time)
    await state.set_state(AdminAddTest.waiting_for_attempts)
    await message.answer("Maksimal urinishlar sonini kiriting (masalan: 1):")

@router.message(AdminAddTest.waiting_for_block_sub1)
async def admin_get_block_sub1(message: Message, state: FSMContext):
    await state.update_data(block_sub1=message.text.strip())
    await state.set_state(AdminAddTest.waiting_for_block_sub2)
    await message.answer("2-asosiy fanning nomini kiriting:")

@router.message(AdminAddTest.waiting_for_block_sub2)
async def admin_get_block_sub2(message: Message, state: FSMContext):
    await state.update_data(block_sub2=message.text.strip())
    await state.set_state(AdminAddTest.waiting_for_attempts)
    await message.answer("Maksimal urinishlar sonini kiriting:")

@router.message(AdminAddTest.waiting_for_attempts)
async def admin_get_attempts(message: Message, state: FSMContext):
    try: att = int(message.text.strip())
    except: att = 1
    await state.update_data(max_attempts=att, questions=[])
    await state.set_state(AdminAddTest.waiting_for_questions)
    await message.answer("📂 Savollarni (rasmli yoki matnli) yuboring. Tugatgach '✅ Testni yakunlash va saqlash' tugmasini bosing.", reply_markup=get_finish_test_keyboard())

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

@router.message(AdminAddTest.waiting_for_questions, F.text == "✅ Testni yakunlash va saqlash")
async def admin_ask_for_answers(message: Message, state: FSMContext):
    if not await is_admin(message.from_user.id): return
    data = await state.get_data()
    if not data.get("questions"):
        await message.answer("❌ Savollar mavjud emas!")
        return
    await state.set_state(AdminAddTest.waiting_for_answers)
    await message.answer(f"🔑 Jami savollar soni: <b>{len(data.get('questions'))}</b> ta.\n\nTo'g'ri javoblarni huddi shuncha miqdorda yuboring (masalan: `A B C D A...`):", reply_markup=get_admin_menu())

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
    
    # Savollar soni bilan javoblar soni to'g'ri kelishini tekshirish
    if len(tokens) != len(questions_list):
        await message.answer(f"❌ Xatolik: Savollar soni ({len(questions_list)} ta) bilan kiritilgan javoblar soni ({len(tokens)} ta) mos kelmadi!\nIltimos, javoblarni qaytadan yuboring:")
        return

    for idx, q in enumerate(questions_list):
        q["correct"] = tokens[idx]

    is_block = data.get("is_block", False)
    
    async with async_session() as session:
        new_test = Test(
            title=data["title"], 
            subject="Blok Test" if is_block else data["subject"], 
            grade_level=data["grade"],
            max_attempts=data["max_attempts"], 
            mode="global_timer",
            duration_minutes=180 if is_block else 60,
            question_time_seconds=data.get("question_time_seconds", 60),
            is_block_test=is_block,
            block_subjects=json.dumps({"sub1": data.get("block_sub1"), "sub2": data.get("block_sub2")}) if is_block else None,
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
                    sec_name = "Ona_tili"
                    points = 1.1
                elif idx < 30:
                    sec_name = "Matematika"
                    points = 1.1
                elif idx < 60:
                    sec_name = "sub1"
                    points = 3.1
                else:
                    sec_name = "sub2"
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
            
        await session.commit()
    await state.clear()
    await message.answer("✅ Test muvaffaqiyatli saqlandi!", reply_markup=get_admin_menu())

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
            keyboard.append([InlineKeyboardButton(text=f"[{t.grade_level}] {t.subject} ({status_emoji})", callback_data="none")])
            keyboard.append([
                InlineKeyboardButton(text="🔄 Holat", callback_data=f"toggle_test_{t.id}"),
                InlineKeyboardButton(text="🏁 Yakunlash", callback_data=f"finish_test_{t.id}"),
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

@router.callback_query(F.data.startswith("finish_test_"))
async def finish_test_by_admin(callback: CallbackQuery):
    async with async_session() as session:
        test = await session.get(Test, int(callback.data.split("_")[2]))
        if test:
            test.is_active = False
            test.is_finished = True
            await session.commit()
            await callback.answer("Test to'liq yakunlandi!", show_alert=True)

@router.callback_query(F.data.startswith("delete_test_"))
async def delete_test(callback: CallbackQuery):
    async with async_session() as session:
        test = await session.get(Test, int(callback.data.split("_")[2]))
        if test:
            await session.delete(test)
            await session.commit()
            await callback.answer("Test o'chirildi!")

async def main():
    logging.basicConfig(level=logging.INFO, stream=sys.stdout)
    bot = Bot(token=BOT_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    storage = MemoryStorage()
    dp = Dispatcher(storage=storage)
    
    dp.include_router(router)
    await init_db()
    
    await bot.delete_webhook(drop_pending_updates=True)
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
