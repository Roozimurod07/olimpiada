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
    # payment_status: none | pending | approved | rejected
    payment_status = Column(String(20), default="none")

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
    value = Column(Text, nullable=False)


class Payment(Base):
    __tablename__ = "payments"
    id = Column(Integer, primary_key=True, autoincrement=True)
    student_id = Column(Integer, ForeignKey("students.id", ondelete="CASCADE"), nullable=False)
    test_id = Column(Integer, ForeignKey("tests.id", ondelete="CASCADE"), nullable=True)  # har bir blok test uchun alohida
    receipt_file_id = Column(String(300), nullable=True)
    receipt_type = Column(String(20), nullable=True)  # photo | document
    status = Column(String(20), default="PENDING")  # PENDING | APPROVED | REJECTED
    created_at = Column(DateTime, default=lambda: datetime.now(timezone.utc))
    reviewed_at = Column(DateTime, nullable=True)
    admin_note = Column(Text, nullable=True)

    student = relationship("Student", back_populates="payments")

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
# 1000 concurrent foydalanuvchi uchun: WAL + katta pool + timeout
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
    cursor.execute("PRAGMA cache_size=-64000")  # 64MB cache
    cursor.close()

async_session = async_sessionmaker(engine, expire_on_commit=False)

async def init_db():
    async with engine.begin() as conn:
        await conn.execute(sa_text("PRAGMA foreign_keys = ON;"))
        await conn.run_sync(Base.metadata.create_all)
        # Migrate: add payment_status column if missing (existing DBs)
        try:
            await conn.execute(sa_text(
                "ALTER TABLE students ADD COLUMN payment_status VARCHAR(20) DEFAULT 'none'"
            ))
        except Exception:
            pass
        # Migrate: add test_id to payments (har bir blok test uchun alohida to'lov)
        try:
            await conn.execute(sa_text(
                "ALTER TABLE payments ADD COLUMN test_id INTEGER REFERENCES tests(id) ON DELETE CASCADE"
            ))
        except Exception:
            pass

    async with async_session() as session:
        defaults = {
            "blok_test_status": "0",
            "paid_mode": "0",
            "payment_card": "8600 XXXX XXXX XXXX",
            "payment_price": "50000",
        }
        for key, val in defaults.items():
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

async def get_setting(key: str, default: str = "0") -> str:
    async with async_session() as session:
        setting = await session.get(Setting, key)
        return setting.value if setting else default


async def set_setting(key: str, value: str):
    async with async_session() as session:
        setting = await session.get(Setting, key)
        if setting:
            setting.value = value
        else:
            session.add(Setting(key=key, value=value))
        await session.commit()


async def get_blok_test_status() -> str:
    return await get_setting("blok_test_status", "0")


async def is_paid_mode() -> bool:
    return (await get_setting("paid_mode", "0")) == "1"


async def student_has_paid_for_test(student_id: int, test_id: int) -> bool:
    """Pullik rejim o'chirilgan bo'lsa True. Yoqilgan bo'lsa — shu test uchun APPROVED to'lov bor-yo'qligi."""
    if not await is_paid_mode():
        return True
    async with async_session() as session:
        payment = (await session.execute(
            select(Payment).where(
                Payment.student_id == student_id,
                Payment.test_id == test_id,
                Payment.status == "APPROVED"
            )
        )).scalar_one_or_none()
        return payment is not None


async def student_payment_status_for_test(student_id: int, test_id: int) -> str:
    """none | pending | approved | rejected — shu test uchun oxirgi to'lov holati."""
    if not await is_paid_mode():
        return "approved"
    async with async_session() as session:
        payment = (await session.execute(
            select(Payment).where(
                Payment.student_id == student_id,
                Payment.test_id == test_id
            ).order_by(Payment.created_at.desc())
        )).scalars().first()
        if not payment:
            return "none"
        return (payment.status or "PENDING").lower()

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

class ProfileEditState(StatesGroup):
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

class PaymentState(StatesGroup):
    waiting_for_test = State()
    waiting_for_receipt = State()

class AdminPaymentSettings(StatesGroup):
    waiting_for_card = State()
    waiting_for_price = State()

router = Router()

user_next_question_flags = {}
user_abort_test_flags = {}  # user_id -> True if left to main menu during test

async def get_main_menu_keyboard():
    keyboard = [
        [KeyboardButton(text="📝 Testni boshlash")]
    ]
    if await get_blok_test_status() == "1":
        keyboard.append([KeyboardButton(text="🗂 Blok testlar")])

    if await is_paid_mode():
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
            [KeyboardButton(text="⚙️ Blok test holati"), KeyboardButton(text="⚙️ Testlarni boshqarish")],
            [KeyboardButton(text="💰 Pullik rejim"), KeyboardButton(text="💳 To'lov sozlamalari")],
            [KeyboardButton(text="📋 To'lov so'rovlari")],
            [KeyboardButton(text="🏆 Admin reyting"), KeyboardButton(text="🔍 O'quvchini qidirish")],
            [KeyboardButton(text="📊 Jonli statistika"), KeyboardButton(text="📥 Excel natijalar")],
            [KeyboardButton(text="⚖️ Apellyatsiyalar"), KeyboardButton(text="👥 Adminlar")],
            [KeyboardButton(text="📢 Xabar yuborish"), KeyboardButton(text="🧹 Bazani tozalash")],
            [KeyboardButton(text="🏠 Asosiy menyu")]
        ],
        resize_keyboard=True
    )

def get_cancel_to_menu_keyboard():
    return ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text="🏠 Asosiy menyu")]],
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
    if await state.get_state() == TestProcessState.in_test.state:
        return
    async with async_session() as session:
        student = (await session.execute(select(Student).where(Student.telegram_id == message.from_user.id))).scalar_one_or_none()
        if not student:
            await message.answer("Siz ro'yxatdan o'tmagansiz. /start ni bosing.")
            return
        pay_info = ""
        if await is_paid_mode():
            payments = (await session.execute(
                select(Payment, Test)
                .outerjoin(Test, Payment.test_id == Test.id)
                .where(Payment.student_id == student.id)
                .order_by(Payment.created_at.desc())
            )).all()
            if payments:
                status_map = {
                    "PENDING": "⏳",
                    "APPROVED": "✅",
                    "REJECTED": "❌",
                }
                lines = []
                seen_tests = set()
                for p, t in payments:
                    key = p.test_id or 0
                    if key in seen_tests:
                        continue
                    seen_tests.add(key)
                    tname = t.title if t else "Noma'lum test"
                    lines.append(f"  {status_map.get(p.status, '?')} {tname}")
                pay_info = "\n💳 To'lovlar (blok testlar):\n" + "\n".join(lines[:8])
            else:
                pay_info = "\n💳 To'lov: hali hech qaysi blok test uchun yo'q"
        markup = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="✏️ Profilni tahrirlash", callback_data="edit_profile")]
        ])
        await message.answer(
            f"👤 <b>Profil:</b>\n\n"
            f"ID: <code>{student.student_id}</code>\n"
            f"Ism: {student.first_name} {student.last_name}\n"
            f"Yosh: {student.age or '-'}\n"
            f"Maktab: {student.school or '-'}\n"
            f"Sinf: {student.grade or '-'}{pay_info}",
            reply_markup=markup
        )


@router.callback_query(F.data == "edit_profile")
async def start_profile_edit(callback: CallbackQuery, state: FSMContext):
    await state.set_state(ProfileEditState.waiting_for_fullname)
    await callback.message.answer(
        "✏️ <b>Profilni tahrirlash</b>\n\nYangi <b>Ism va Familiyangizni</b> kiriting:\n\n"
        "(Bekor qilish: 🏠 Asosiy menyu)",
        reply_markup=get_cancel_to_menu_keyboard()
    )
    await callback.answer()


@router.message(ProfileEditState.waiting_for_fullname)
async def profile_edit_fullname(message: Message, state: FSMContext):
    if message.text in ("🏠 Asosiy menyu", "⬅️ Bosh menyu", "🚀 Start"):
        await back_to_menu(message, state)
        return
    await state.update_data(fullname=message.text.strip())
    await state.set_state(ProfileEditState.waiting_for_age)
    await message.answer("Yoshingizni kiriting (masalan: 16):")


@router.message(ProfileEditState.waiting_for_age)
async def profile_edit_age(message: Message, state: FSMContext):
    if message.text in ("🏠 Asosiy menyu", "⬅️ Bosh menyu", "🚀 Start"):
        await back_to_menu(message, state)
        return
    await state.update_data(age=message.text.strip())
    await state.set_state(ProfileEditState.waiting_for_grade)
    await message.answer("Sinfingizni kiriting (masalan: 11-sinf):")


@router.message(ProfileEditState.waiting_for_grade)
async def profile_edit_grade(message: Message, state: FSMContext):
    if message.text in ("🏠 Asosiy menyu", "⬅️ Bosh menyu", "🚀 Start"):
        await back_to_menu(message, state)
        return
    await state.update_data(grade=message.text.strip())
    await state.set_state(ProfileEditState.waiting_for_school)
    await message.answer("Maktabingiz raqami yoki nomini kiriting:")


@router.message(ProfileEditState.waiting_for_school)
async def profile_edit_school(message: Message, state: FSMContext):
    if message.text in ("🏠 Asosiy menyu", "⬅️ Bosh menyu", "🚀 Start"):
        await back_to_menu(message, state)
        return
    data = await state.get_data()
    fullname_parts = data["fullname"].split(" ", 1)
    first_name = fullname_parts[0]
    last_name = fullname_parts[1] if len(fullname_parts) > 1 else "-"
    async with async_session() as session:
        student = (await session.execute(
            select(Student).where(Student.telegram_id == message.from_user.id)
        )).scalar_one_or_none()
        if student:
            student.first_name = first_name
            student.last_name = last_name
            student.age = data["age"]
            student.grade = data["grade"]
            student.school = message.text.strip()
            await session.commit()
    await state.clear()
    main_menu = await get_main_menu_keyboard()
    await message.answer(
        f"✅ Profil muvaffaqiyatli yangilandi!\n\n"
        f"Ism: <b>{first_name} {last_name}</b>\n"
        f"Yosh: {data['age']}\n"
        f"Sinf: {data['grade']}\n"
        f"Maktab: {message.text.strip()}",
        reply_markup=main_menu
    )

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
    """Apellyatsiya: oddiy va blok testlar uchun (sahifalash bilan)."""
    parts = callback.data.split("_")
    session_id = int(parts[3])
    page = int(parts[4]) if len(parts) > 4 else 0
    PER_PAGE = 10

    async with async_session() as session:
        ts = await session.get(TestSession, session_id)
        if not ts:
            await callback.answer("Sessiya topilmadi!", show_alert=True)
            return
        test = await session.get(Test, ts.test_id)

        if not test.is_finished:
            await callback.answer("Bu test uchun hali tahlil ochiq emas!", show_alert=True)
            return

        questions = list((await session.execute(
            select(Question).where(Question.test_id == test.id).order_by(Question.id)
        )).scalars().all())
        answers = {
            a.question_id: a.selected_option
            for a in (await session.execute(select(Answer).where(Answer.session_id == ts.id))).scalars().all()
        }

        total_q = len(questions)
        total_pages = max(1, (total_q + PER_PAGE - 1) // PER_PAGE)
        page = max(0, min(page, total_pages - 1))
        start = page * PER_PAGE
        end = min(start + PER_PAGE, total_q)
        page_questions = questions[start:end]

        text = (
            f"📋 <b>Test tahlili: {test.subject} ({test.title})</b>\n"
            f"{'🧩 Blok test' if test.is_block_test else '📝 Oddiy test'}\n"
            f"⭐ Ball: {ts.score} ({ts.score_percentage}%)\n"
            f"✅ To'g'ri: {ts.correct_answers} | ❌ Noto'g'ri: {ts.wrong_answers} | ⭕ Javobsiz: {ts.unanswered}\n"
            f"📄 Sahifa {page + 1}/{total_pages} (jami {total_q} ta savol)\n\n"
        )

        keyboard = []
        for idx, q in enumerate(page_questions, start=start + 1):
            sel = answers.get(q.id) or "Javob berilmagan"
            status = "✅" if sel == q.correct_option else "❌"
            sec = f"[{q.section_name}] " if q.section_name else ""
            q_preview = (q.question_text[:80] + "…") if len(q.question_text) > 80 else q.question_text
            text += (
                f"<b>{idx}. {sec}{q_preview}</b>\n"
                f"Sizning javob: <b>{sel}</b> {status} | To'g'ri: <b>{q.correct_option}</b>\n\n"
            )
            keyboard.append([InlineKeyboardButton(
                text=f"⚖️ {idx}-savolga apellyatsiya",
                callback_data=f"appeal_q_{ts.id}_{q.id}"
            )])

        nav = []
        if page > 0:
            nav.append(InlineKeyboardButton(text="⬅️ Oldingi", callback_data=f"student_appeal_test_{session_id}_{page - 1}"))
        if page < total_pages - 1:
            nav.append(InlineKeyboardButton(text="Keyingi ➡️", callback_data=f"student_appeal_test_{session_id}_{page + 1}"))
        if nav:
            keyboard.append(nav)

        keyboard.append([InlineKeyboardButton(text="🎓 Sertifikatni yuklab olish", callback_data=f"get_cert_{ts.id}")])

        if len(text) > 4000:
            text = text[:3900] + "\n... (matn qisqartirildi)"

        try:
            await callback.message.edit_text(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=keyboard))
        except Exception:
            await callback.message.answer(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=keyboard))
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
    if await state.get_state() == TestProcessState.in_test.state:
        return
    async with async_session() as session:
        student = (await session.execute(select(Student).where(Student.telegram_id == message.from_user.id))).scalar_one_or_none()
        if not student or not student.grade:
            await message.answer("❌ Profilingizda sinf ko'rsatilmagan yoki ro'yxatdan o'tmagansiz.")
            return

        # Oddiy testlar bepul (to'lov faqat blok testlar uchun)
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
    if await state.get_state() == TestProcessState.in_test.state:
        return

    if await get_blok_test_status() != "1":
        await message.answer("⚠️ Hozirda blok test bo'limi yopiq.")
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

        paid_mode = await is_paid_mode()
        keyboard_buttons = []
        for t in tests:
            mark = ""
            if paid_mode:
                st = await student_payment_status_for_test(student.id, t.id)
                if st == "approved":
                    mark = " ✅"
                elif st == "pending":
                    mark = " ⏳"
                else:
                    mark = " 🔒"
            keyboard_buttons.append([InlineKeyboardButton(
                text=f"🧩 {t.subject} — {t.title}{mark}",
                callback_data=f"start_block_{t.id}"
            )])
        note = ""
        if paid_mode:
            note = "\n\n🔒 — to'lov kerak | ⏳ — tekshiruvda | ✅ — to'langan\nHar bir blok test uchun alohida to'lov qilinadi."
        await message.answer(
            f"🗂 <b>Mavjud blok testlar:</b>\n\nTestni tanlang:{note}",
            reply_markup=InlineKeyboardMarkup(inline_keyboard=keyboard_buttons)
        )

# Fan kalitlari va ko'rinadigan nomlar
BLOCK_SECTION_KEYS = ["Tarix", "Ona_tili", "Matematika", "sub1", "sub2"]
BLOCK_SECTION_DISPLAY = {
    "Tarix": "🏛 Tarix",
    "Ona_tili": "🇺🇿 Ona tili",
    "Matematika": "🔢 Matematika",
}

def _section_variants(section: str, block_subs: dict = None):
    """Eski va yangi section_name variantlari (Ona tili / Ona_tili va h.k.)"""
    variants = [section]
    if section == "Ona_tili":
        variants.extend(["Ona tili", "Ona_tili"])
    elif section == "Tarix":
        variants.append("Tarix")
    elif section == "Matematika":
        variants.append("Matematika")
    elif section in ("sub1", "sub2") and block_subs:
        real = block_subs.get(section)
        if real:
            variants.append(real)
    return list(dict.fromkeys(variants))

async def _load_section_questions(session, test_id: int, section: str):
    """Fan savollarini eski/yangi format bilan yuklash"""
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
    """Fan bo'yicha jami / javob berilgan savollar sonini qaytaradi"""
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
    """Fanlar menyusi — bajarilgan fanlar ✅ bilan"""
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
        "Istalgan fandan boshlang. Javob bergan savol keyingisiga o'tadi.\n"
        "Fan tugagach ✅ belgilanadi.\n"
        "Jami maksimal ball: <b>189</b>"
    )
    return text, InlineKeyboardMarkup(inline_keyboard=keyboard)

async def send_next_block_question(bot: Bot, chat_id: int, ts_id: int, test_id: int, section: str, state: FSMContext):
    """Fan ichidagi keyingi javob berilmagan savolni yuboradi. Hammasi tugasa True qaytaradi."""
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
            return True  # fan tugadi

        display_name = BLOCK_SECTION_DISPLAY.get(section, section)
        if section in ("sub1", "sub2"):
            test = await session.get(Test, test_id)
            block_subs = json.loads(test.block_subjects) if test and test.block_subjects else {}
            display_name = f"📘 {block_subs.get(section, section)}" if section == "sub1" else f"📙 {block_subs.get(section, section)}"

        options = [("A", next_q.option_a), ("B", next_q.option_b)]
        if next_q.option_c:
            options.append(("C", next_q.option_c))
        if next_q.option_d:
            options.append(("D", next_q.option_d))

        keyboard_buttons = []
        row = []
        for opt_key, text_val in options:
            row.append(InlineKeyboardButton(
                text=f"{opt_key}) {text_val}",
                callback_data=f"b_ans_{ts_id}_{next_q.id}_{opt_key}_{section}_{test_id}"
            ))
            if len(row) == 2:
                keyboard_buttons.append(row)
                row = []
        if row:
            keyboard_buttons.append(row)
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

        # Har bir blok test uchun alohida to'lov (pullik rejim yoqilgan bo'lsa)
        if student and not await student_has_paid_for_test(student.id, test_id):
            st = await student_payment_status_for_test(student.id, test_id)
            if st == "pending":
                await callback.answer("⏳ Bu test uchun to'lovingiz hali tekshiruvda. Admin javobini kuting.", show_alert=True)
                return
            await callback.answer("🔒 Bu blok test uchun to'lov qilinmagan!", show_alert=True)
            card = await get_setting("payment_card", "8600 XXXX XXXX XXXX")
            price = await get_setting("payment_price", "50000")
            markup = InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text="✅ Shu test uchun to'lov qildim", callback_data=f"pay_for_test_{test_id}")]
            ])
            await callback.message.answer(
                f"🔒 <b>Pullik rejim yoqilgan</b>\n\n"
                f"🧩 Test: <b>{test.title}</b>\n"
                f"Karta: <code>{card}</code>\n"
                f"Narx: <b>{price}</b> so'm\n\n"
                f"Har bir blok test uchun alohida to'lov qilinadi.\n"
                f"To'lovni amalga oshirgach, pastdagi tugmani bosing va chek yuboring.",
                reply_markup=markup
            )
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

        if remaining_min <= 0:
            ts.status = "COMPLETED"
            ts.finished_at = datetime.now(timezone.utc)
            await calculate_and_save_results(session, ts)
            await session.commit()
            await callback.message.edit_text("⏰ Blok test uchun 180 daqiqa vaqt tugadi! Natija avtomatik saqlandi.")
            await callback.answer()
            return

        block_subs = json.loads(test.block_subjects) if test.block_subjects else {"sub1": "1-Asosiy fan", "sub2": "2-Asosiy fan"}
        progress = {}
        for key in BLOCK_SECTION_KEYS:
            total, ans, _ = await get_block_section_progress(session, test_id, ts.id, key)
            progress[key] = (total, ans)

        remaining_text = f"{int(remaining_min)} daqiqa {int((remaining_min % 1)*60)} soniya"
        text, markup = await build_block_subjects_keyboard(test_id, block_subs, progress, remaining_text, test.title)

    await callback.message.edit_text(text, reply_markup=markup)
    await callback.answer()


@router.callback_query(F.data.startswith("run_block_"))
async def run_block_subject_questions(callback: CallbackQuery, state: FSMContext, bot: Bot):
    parts = callback.data.split("_")
    test_id = int(parts[2])
    section = "_".join(parts[3:])  # Tarix | Ona_tili | Matematika | sub1 | sub2

    async with async_session() as session:
        student = (await session.execute(select(Student).where(Student.telegram_id == callback.from_user.id))).scalar_one_or_none()
        test = await session.get(Test, test_id)

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
        if (datetime.now(timezone.utc) - started).total_seconds() >= 180 * 60:
            ts.status = "COMPLETED"
            ts.finished_at = datetime.now(timezone.utc)
            await calculate_and_save_results(session, ts)
            await session.commit()
            await callback.answer("⏰ 180 daqiqa vaqt tugadi!", show_alert=True)
            main_menu = await get_main_menu_keyboard()
            await callback.message.answer(f"🏆 Blok test avtomatik yakunlandi.\nBall: {ts.score} ({ts.score_percentage}%)", reply_markup=main_menu)
            return

        total, ans, questions = await get_block_section_progress(session, test_id, ts.id, section)
        if total == 0:
            await callback.answer("Bu fanda savollar topilmadi! Admin yangi blok test yuklashi kerak (Ona_tili kaliti).", show_alert=True)
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
    await bot.send_message(callback.from_user.id, f"📚 <b>{display}</b> fani boshlandi. Savollar ketma-ket keladi.")
    done = await send_next_block_question(bot, callback.from_user.id, ts_id, test_id, section, state)
    if done is True:
        await bot.send_message(callback.from_user.id, f"✅ <b>{display}</b> fani allaqachon bajarilgan!")
    elif done == "TIME_UP":
        await bot.send_message(callback.from_user.id, "⏰ Vaqt tugadi!")
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
            remaining_text = f"{int(remaining_min)} daqiqa {int((remaining_min % 1)*60)} soniya"
            if remaining_min <= 0:
                ts.status = "COMPLETED"
                ts.finished_at = datetime.now(timezone.utc)
                await calculate_and_save_results(session, ts)
                await session.commit()
                await state.clear()
                main_menu = await get_main_menu_keyboard()
                await callback.message.answer(f"⏰ Vaqt tugadi! Natija: {ts.score} ball ({ts.score_percentage}%)", reply_markup=main_menu)
                await callback.answer()
                return
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
    # Format: b_ans_{session_id}_{question_id}_{option}_{section}_{test_id}
    parts = callback.data.split("_")
    if len(parts) < 6:
        await callback.answer("Xato format!", show_alert=True)
        return
    session_id = int(parts[2])
    question_id = int(parts[3])
    selected = parts[4]
    # section va test_id: parts[5:] — section bitta yoki Ona_tili kabi
    # qayta parse: b_ans | sid | qid | opt | SECTION... | test_id
    # SECTION: Tarix | Ona_tili | Matematika | sub1 | sub2
    rest = parts[5:]
    test_id = int(rest[-1])
    section = "_".join(rest[:-1])

    async with async_session() as session:
        ts = await session.get(TestSession, session_id)
        if not ts or ts.status != "IN_PROGRESS":
            await callback.answer("Test yakunlangan!", show_alert=True)
            return
        started = ts.started_at
        if started.tzinfo is None:
            started = started.replace(tzinfo=timezone.utc)
        if (datetime.now(timezone.utc) - started).total_seconds() >= 180 * 60:
            await callback.answer("⏰ Vaqt tugagan! Javob qabul qilinmadi.", show_alert=True)
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

    # Javob berilgan savol xabarini o'chirish
    try:
        await callback.message.delete()
    except Exception:
        pass

    # Keyingi savol yoki fan tugadi
    result = await send_next_block_question(bot, callback.from_user.id, session_id, test_id, section, state)

    if result is True:
        display = BLOCK_SECTION_DISPLAY.get(section, section)
        if section in ("sub1", "sub2"):
            async with async_session() as session:
                test = await session.get(Test, test_id)
                block_subs = json.loads(test.block_subjects) if test and test.block_subjects else {}
                display = block_subs.get(section, section)
        await bot.send_message(
            callback.from_user.id,
            f"✅ <b>{display}</b> fani to'liq bajarildi!\n\nBoshqa fanga o'tishingiz yoki testni yakunlashingiz mumkin.",
            reply_markup=InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text="📋 Fanlar menyusiga qaytish", callback_data=f"back_to_block_menu_{test_id}")]
            ])
        )
        await state.clear()
    elif result == "TIME_UP":
        async with async_session() as session:
            ts = await session.get(TestSession, session_id)
            if ts and ts.status == "IN_PROGRESS":
                ts.status = "COMPLETED"
                ts.finished_at = datetime.now(timezone.utc)
                await calculate_and_save_results(session, ts)
                await session.commit()
                main_menu = await get_main_menu_keyboard()
                await bot.send_message(
                    callback.from_user.id,
                    f"⏰ Vaqt tugadi!\n🏆 Natija: {ts.score} ball ({ts.score_percentage}%)",
                    reply_markup=main_menu
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
    main_menu = await get_main_menu_keyboard()
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

        questions = list((await session.execute(select(Question).where(Question.test_id == test_id))).scalars().all())
        if not questions:
            await callback.answer("Bu testda savollar yo'q!", show_alert=True)
            return
        
        # Har bir o'quvchi uchun savollar tartibini random qilish
        random.shuffle(questions)
            
        test_session = TestSession(student_id=student.id, test_id=test_id, status="IN_PROGRESS")
        session.add(test_session)
        await session.commit()
        await session.refresh(test_session)

        await callback.message.edit_text(f"🚀 <b>{test.subject} ({test.title})</b> testi boshlandi!\n⏱ Har bir savol: <b>{test.question_time_seconds}</b> soniya")
        user_id = callback.from_user.id
        await state.set_state(TestProcessState.in_test)
        
        # Umumiy xavfsizlik vaqti: savollar soni * har bir savol vaqti + 60s buffer
        question_seconds = max(5, int(test.question_time_seconds))
        total_duration_sec = len(questions) * question_seconds + 60
        start_timestamp = datetime.now(timezone.utc)
        
        user_abort_test_flags.pop(user_id, None)

        for index, q in enumerate(questions):
            if user_abort_test_flags.pop(user_id, None):
                break

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
            timer_msg = await bot.send_message(
                chat_id=user_id,
                text=f"⏳ <b>Qolgan vaqt: {question_seconds} soniya</b>"
            )
            last_second = question_seconds

            user_next_question_flags.pop(user_id, None)

            aborted = False
            while True:
                await asyncio.sleep(0.2)
                if user_abort_test_flags.get(user_id):
                    aborted = True
                    break
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

            if aborted or user_abort_test_flags.pop(user_id, None):
                break

            if (datetime.now(timezone.utc) - start_timestamp).total_seconds() >= total_duration_sec:
                break

        # Agar foydalanuvchi asosiy menyuga chiqqan bo'lsa — sessiyani qayta yakunlamaslik
        if user_abort_test_flags.pop(user_id, None):
            return

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
                    save_result_to_sheet(
                        student_obj.student_id,
                        f"{student_obj.first_name} {student_obj.last_name}",
                        student_obj.age or "-",
                        student_obj.school or "-",
                        student_obj.grade or "-",
                        test_obj.title,
                        test_obj.subject,
                        sess.score,
                        sess.score_percentage,
                        sess.correct_answers,
                        sess.wrong_answers
                    )

                await state.clear()
                main_menu = await get_main_menu_keyboard()
                await bot.send_message(
                    chat_id=user_id,
                    text=f"🏆 <b>TEST YAKUNLANDI!</b>\n\nNatijangiz: {sess.score} ball ({sess.score_percentage}%).",
                    reply_markup=main_menu
                )

@router.callback_query(F.data.startswith("ans_"))
async def save_answer(callback: CallbackQuery, bot: Bot):
    parts = callback.data.split("_")
    session_id, question_id, selected = int(parts[1]), int(parts[2]), parts[3]
    async with async_session() as session:
        existing = (await session.execute(select(Answer).where(Answer.session_id == session_id, Answer.question_id == question_id))).scalar_one_or_none()
        if existing:
            existing.selected_option = selected
        else:
            session.add(Answer(session_id=session_id, question_id=question_id, selected_option=selected))
        await session.commit()

    user_next_question_flags[callback.from_user.id] = {"target_index": None}
    await callback.answer(f"Tanlandi: {selected}")

async def calculate_and_save_results(session, sess: TestSession):
    questions = (await session.execute(select(Question).where(Question.test_id == sess.test_id))).scalars().all()
    answers = {a.question_id: a.selected_option for a in (await session.execute(select(Answer).where(Answer.session_id == sess.id))).scalars().all()}
    
    correct, wrong, unanswered, total_score = 0, 0, 0, 0.0
    max_possible = 0.0
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
    # Blok test uchun max 189 ball asosida foiz, oddiy test uchun max ball asosida
    if max_possible > 0:
        sess.score_percentage = round((total_score / max_possible) * 100, 2)
    else:
        sess.score_percentage = 0.0

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
    if await state.get_state() == TestProcessState.in_test.state:
        return
    await message.answer("ℹ️ Professional Olimpiada Tizimi v3.3 — DTM blok testlar, pullik rejim va 180 daqiqalik umumiy imtihon.")


# ==================== TO'LOV (STUDENT) ====================
# Har bir blok test uchun alohida to'lov. Admin yangi blok test yuklasa — yana to'lov kerak.

@router.message(F.text == "💳 To'lov qilish")
async def student_payment_menu(message: Message, state: FSMContext):
    if await state.get_state() == TestProcessState.in_test.state:
        return
    if not await is_paid_mode():
        await message.answer("ℹ️ Hozirda pullik rejim o'chirilgan. Testlar bepul.")
        return

    async with async_session() as session:
        student = (await session.execute(
            select(Student).where(Student.telegram_id == message.from_user.id)
        )).scalar_one_or_none()
        if not student:
            await message.answer("Ro'yxatdan o'tmagansiz. /start ni bosing.")
            return

        tests = (await session.execute(
            select(Test).where(
                Test.is_active == True,
                Test.is_finished == False,
                Test.is_block_test == True
            )
        )).scalars().all()

        if not tests:
            await message.answer("⚠️ Hozirda faol blok testlar yo'q. To'lov qilish uchun test bo'lishi kerak.")
            return

        keyboard = []
        for t in tests:
            st = await student_payment_status_for_test(student.id, t.id)
            if st == "approved":
                mark = " ✅ to'langan"
            elif st == "pending":
                mark = " ⏳ tekshiruvda"
            else:
                mark = " 🔒 to'lov kerak"
            keyboard.append([InlineKeyboardButton(
                text=f"🧩 {t.title}{mark}",
                callback_data=f"pay_for_test_{t.id}"
            )])

    card = await get_setting("payment_card", "8600 XXXX XXXX XXXX")
    price = await get_setting("payment_price", "50000")
    await message.answer(
        f"💳 <b>To'lov — har bir blok test alohida</b>\n\n"
        f"Karta: <code>{card}</code>\n"
        f"Narx (har bir test): <b>{price}</b> so'm\n\n"
        f"Qaysi blok test uchun to'lov qilmoqchisiz?",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=keyboard)
    )


@router.callback_query(F.data.startswith("pay_for_test_"))
async def pay_for_specific_test(callback: CallbackQuery, state: FSMContext):
    if not await is_paid_mode():
        await callback.answer("Pullik rejim o'chirilgan!", show_alert=True)
        return
    test_id = int(callback.data.split("_")[3])
    async with async_session() as session:
        student = (await session.execute(
            select(Student).where(Student.telegram_id == callback.from_user.id)
        )).scalar_one_or_none()
        test = await session.get(Test, test_id)
        if not student or not test:
            await callback.answer("Xato!", show_alert=True)
            return
        st = await student_payment_status_for_test(student.id, test_id)
        if st == "approved":
            await callback.answer("✅ Bu test uchun to'lov allaqachon tasdiqlangan!", show_alert=True)
            return
        if st == "pending":
            await callback.answer("⏳ Bu test uchun to'lov tekshiruvda. Admin javobini kuting.", show_alert=True)
            return

    await state.update_data(pay_test_id=test_id)
    await state.set_state(PaymentState.waiting_for_receipt)
    card = await get_setting("payment_card", "8600 XXXX XXXX XXXX")
    price = await get_setting("payment_price", "50000")
    await callback.message.answer(
        f"📎 <b>{test.title}</b> uchun to'lov chekini yuboring (rasm yoki PDF).\n\n"
        f"Karta: <code>{card}</code>\n"
        f"Narx: <b>{price}</b> so'm\n\n"
        f"(Bekor qilish: 🏠 Asosiy menyu)",
        reply_markup=get_cancel_to_menu_keyboard()
    )
    await callback.answer()


@router.message(PaymentState.waiting_for_receipt, F.photo | F.document)
async def receive_payment_receipt(message: Message, state: FSMContext, bot: Bot):
    if message.text and message.text in ("🏠 Asosiy menyu", "⬅️ Bosh menyu", "🚀 Start"):
        await back_to_menu(message, state)
        return

    data = await state.get_data()
    test_id = data.get("pay_test_id")
    if not test_id:
        await message.answer("❌ Avval qaysi test uchun to'lov qilishni tanlang (💳 To'lov qilish).")
        await state.clear()
        return

    file_id = None
    receipt_type = None
    if message.photo:
        file_id = message.photo[-1].file_id
        receipt_type = "photo"
    elif message.document:
        file_id = message.document.file_id
        receipt_type = "document"
    else:
        await message.answer("❌ Rasm yoki hujjat yuboring.")
        return

    async with async_session() as session:
        student = (await session.execute(
            select(Student).where(Student.telegram_id == message.from_user.id)
        )).scalar_one_or_none()
        if not student:
            await message.answer("Ro'yxatdan o'tmagansiz.")
            await state.clear()
            return

        test = await session.get(Test, test_id)
        payment = Payment(
            student_id=student.id,
            test_id=test_id,
            receipt_file_id=file_id,
            receipt_type=receipt_type,
            status="PENDING"
        )
        session.add(payment)
        await session.commit()
        await session.refresh(payment)

        admins = (await session.execute(select(Admin))).scalars().all()
        admin_ids = set(SUPER_ADMIN_IDS) | {a.telegram_id for a in admins}

        test_title = test.title if test else f"Test #{test_id}"
        caption = (
            f"💳 <b>Yangi to'lov so'rovi #{payment.id}</b>\n\n"
            f"🧩 Test: <b>{test_title}</b>\n"
            f"👤 {student.first_name} {student.last_name}\n"
            f"ID: <code>{student.student_id}</code>\n"
            f"Sinf: {student.grade or '-'}\n"
            f"TG: <code>{student.telegram_id}</code>"
        )
        kb = InlineKeyboardMarkup(inline_keyboard=[[
            InlineKeyboardButton(text="✅ Tasdiqlash", callback_data=f"pay_acc_{payment.id}"),
            InlineKeyboardButton(text="❌ To'lov qilinmadi", callback_data=f"pay_rej_{payment.id}")
        ]])

        for aid in admin_ids:
            try:
                if receipt_type == "photo":
                    await bot.send_photo(chat_id=aid, photo=file_id, caption=caption, reply_markup=kb)
                else:
                    await bot.send_document(chat_id=aid, document=file_id, caption=caption, reply_markup=kb)
            except Exception:
                try:
                    await bot.send_message(chat_id=aid, text=caption + "\n\n(Chek yuborilmadi)", reply_markup=kb)
                except Exception:
                    pass

    await state.clear()
    main_menu = await get_main_menu_keyboard()
    await message.answer(
        f"✅ <b>{test_title}</b> uchun chekingiz adminga yuborildi. Tasdiqlanishini kuting.",
        reply_markup=main_menu
    )


@router.message(PaymentState.waiting_for_receipt)
async def payment_receipt_invalid(message: Message, state: FSMContext):
    if message.text in ("🏠 Asosiy menyu", "⬅️ Bosh menyu", "🚀 Start"):
        await back_to_menu(message, state)
        return
    await message.answer("❌ Iltimos, chekni rasm yoki PDF/hujjat sifatida yuboring.")


# ==================== TO'LOV (ADMIN) ====================

@router.message(F.text == "💰 Pullik rejim")
async def admin_paid_mode_menu(message: Message):
    if not await is_admin(message.from_user.id):
        return
    status = await get_setting("paid_mode", "0")
    status_text = "🟢 Yoqilgan" if status == "1" else "🔴 O'chirilgan"
    markup = InlineKeyboardMarkup(inline_keyboard=[[
        InlineKeyboardButton(text="🟢 Yoqish", callback_data="set_paid_1"),
        InlineKeyboardButton(text="🔴 O'chirish", callback_data="set_paid_0")
    ]])
    await message.answer(
        f"💰 <b>Pullik rejim</b>\n\nHozirgi holat: <b>{status_text}</b>\n\n"
        f"🟢 Yoqilganda: har bir <b>blok test</b> uchun alohida to'lov kerak.\n"
        f"   (Masalan: Matematika+Fizika va Matematika+Ingliz — ikki xil to'lov)\n"
        f"🔴 O'chirilganda: barcha testlar bepul, to'lov so'ralmaydi.",
        reply_markup=markup
    )


@router.callback_query(F.data.in_(["set_paid_1", "set_paid_0"]))
async def update_paid_mode(callback: CallbackQuery):
    if not await is_admin(callback.from_user.id):
        return
    new_val = "1" if callback.data == "set_paid_1" else "0"
    await set_setting("paid_mode", new_val)
    status_text = "🟢 Yoqildi" if new_val == "1" else "🔴 O'chirildi"
    await callback.message.edit_text(f"✅ Pullik rejim: <b>{status_text}</b>")
    await callback.answer()


@router.message(F.text == "💳 To'lov sozlamalari")
async def admin_payment_settings_prompt(message: Message, state: FSMContext):
    if not await is_admin(message.from_user.id):
        return
    card = await get_setting("payment_card", "8600 XXXX XXXX XXXX")
    price = await get_setting("payment_price", "50000")
    await state.set_state(AdminPaymentSettings.waiting_for_card)
    await message.answer(
        f"💳 Hozirgi karta: <code>{card}</code>\n"
        f"Hozirgi narx: <b>{price}</b> so'm\n\n"
        f"Yangi <b>karta raqamini</b> yuboring (yoki o'zgartirmaslik uchun `-`):",
        reply_markup=get_cancel_to_menu_keyboard()
    )


@router.message(AdminPaymentSettings.waiting_for_card)
async def admin_set_card(message: Message, state: FSMContext):
    if message.text in ("🏠 Asosiy menyu", "⬅️ Bosh menyu", "🚀 Start"):
        await back_to_menu(message, state)
        return
    txt = message.text.strip()
    if txt != "-":
        await set_setting("payment_card", txt)
    await state.set_state(AdminPaymentSettings.waiting_for_price)
    await message.answer("Test ishlash <b>narxini</b> so'mda kiriting (masalan: 50000) yoki `-`:")


@router.message(AdminPaymentSettings.waiting_for_price)
async def admin_set_price(message: Message, state: FSMContext):
    if message.text in ("🏠 Asosiy menyu", "⬅️ Bosh menyu", "🚀 Start"):
        await back_to_menu(message, state)
        return
    txt = message.text.strip()
    if txt != "-":
        await set_setting("payment_price", txt)
    await state.clear()
    card = await get_setting("payment_card")
    price = await get_setting("payment_price")
    await message.answer(
        f"✅ To'lov sozlamalari saqlandi!\n\nKarta: <code>{card}</code>\nNarx: <b>{price}</b> so'm",
        reply_markup=get_admin_menu()
    )


@router.message(F.text == "📋 To'lov so'rovlari")
async def admin_payment_requests(message: Message):
    if not await is_admin(message.from_user.id):
        return
    async with async_session() as session:
        payments = (await session.execute(
            select(Payment).where(Payment.status == "PENDING").order_by(Payment.created_at.desc())
        )).scalars().all()
        if not payments:
            await message.answer("⏳ Kutilayotgan to'lov so'rovlari yo'q.")
            return
        for p in payments:
            student = await session.get(Student, p.student_id)
            test = await session.get(Test, p.test_id) if p.test_id else None
            test_line = f"🧩 Test: <b>{test.title}</b>\n" if test else ""
            caption = (
                f"💳 <b>To'lov #{p.id}</b>\n"
                f"{test_line}"
                f"👤 {student.first_name} {student.last_name} ({student.student_id})\n"
                f"Sinf: {student.grade or '-'}"
            )
            kb = InlineKeyboardMarkup(inline_keyboard=[[
                InlineKeyboardButton(text="✅ Tasdiqlash", callback_data=f"pay_acc_{p.id}"),
                InlineKeyboardButton(text="❌ To'lov qilinmadi", callback_data=f"pay_rej_{p.id}")
            ]])
            try:
                if p.receipt_type == "photo" and p.receipt_file_id:
                    await message.answer_photo(photo=p.receipt_file_id, caption=caption, reply_markup=kb)
                elif p.receipt_file_id:
                    await message.answer_document(document=p.receipt_file_id, caption=caption, reply_markup=kb)
                else:
                    await message.answer(caption, reply_markup=kb)
            except Exception:
                await message.answer(caption + "\n(Chek ochilmadi)", reply_markup=kb)


@router.callback_query(F.data.startswith("pay_acc_"))
async def approve_payment(callback: CallbackQuery, bot: Bot):
    if not await is_admin(callback.from_user.id):
        return
    payment_id = int(callback.data.split("_")[2])
    async with async_session() as session:
        payment = await session.get(Payment, payment_id)
        if not payment or payment.status != "PENDING":
            await callback.answer("So'rov topilmadi yoki allaqachon ko'rib chiqilgan!", show_alert=True)
            return
        payment.status = "APPROVED"
        payment.reviewed_at = datetime.now(timezone.utc)
        student = await session.get(Student, payment.student_id)
        test = await session.get(Test, payment.test_id) if payment.test_id else None
        await session.commit()

        test_title = test.title if test else "blok test"
        if student and student.telegram_id:
            try:
                main_menu = await get_main_menu_keyboard()
                await bot.send_message(
                    chat_id=student.telegram_id,
                    text=f"✅ <b>To'lovingiz tasdiqlandi!</b>\n\n"
                         f"🧩 Test: <b>{test_title}</b>\n"
                         f"Endi shu blok testdan foydalanishingiz mumkin.",
                    reply_markup=main_menu
                )
            except Exception:
                pass

    try:
        await callback.message.edit_caption(
            caption=(callback.message.caption or "") + "\n\n✅ <b>TASDIQLANDI</b>"
        )
    except Exception:
        try:
            await callback.message.edit_text(
                (callback.message.text or "") + "\n\n✅ <b>TASDIQLANDI</b>"
            )
        except Exception:
            pass
    await callback.answer("✅ To'lov tasdiqlandi!")


@router.callback_query(F.data.startswith("pay_rej_"))
async def reject_payment(callback: CallbackQuery, bot: Bot):
    if not await is_admin(callback.from_user.id):
        return
    payment_id = int(callback.data.split("_")[2])
    async with async_session() as session:
        payment = await session.get(Payment, payment_id)
        if not payment or payment.status != "PENDING":
            await callback.answer("So'rov topilmadi yoki allaqachon ko'rib chiqilgan!", show_alert=True)
            return
        payment.status = "REJECTED"
        payment.reviewed_at = datetime.now(timezone.utc)
        student = await session.get(Student, payment.student_id)
        test = await session.get(Test, payment.test_id) if payment.test_id else None
        await session.commit()

        test_title = test.title if test else "blok test"
        if student and student.telegram_id:
            try:
                await bot.send_message(
                    chat_id=student.telegram_id,
                    text=f"❌ <b>To'lov qilinmadi</b> deb belgiladi admin.\n\n"
                         f"🧩 Test: <b>{test_title}</b>\n"
                         f"Agar xato deb hisoblasangiz, qayta to'lov qilib chek yuboring."
                )
            except Exception:
                pass

    try:
        await callback.message.edit_caption(
            caption=(callback.message.caption or "") + "\n\n❌ <b>RAD ETILDI</b>"
        )
    except Exception:
        try:
            await callback.message.edit_text(
                (callback.message.text or "") + "\n\n❌ <b>RAD ETILDI</b>"
            )
        except Exception:
            pass
    await callback.answer("❌ To'lov rad etildi!")


@router.message(Command("admin"))
async def admin_panel(message: Message, state: FSMContext):
    if not await is_admin(message.from_user.id):
        return
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
        # Blok test hamma sinflar uchun — sinf so'ralmaydi
        await state.update_data(grade="Barcha")
        await state.set_state(AdminAddTest.waiting_for_block_sub1)
        await message.answer("1-asosiy fanning nomini kiriting (masalan: Fizika yoki Biologiya):")
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
        await message.answer("1-asosiy fanning nomini kiriting (masalan: Fizika yoki Biologiya):")
    else:
        await state.set_state(AdminAddTest.waiting_for_question_time)
        await message.answer("⏱ Har bir savol uchun vaqtni **soniyada** kiriting (masalan: `15` yoki `60`):")

@router.message(AdminAddTest.waiting_for_question_time)
async def admin_get_question_time(message: Message, state: FSMContext):
    try:
        q_time = int(message.text.strip())
        if q_time < 5:
            q_time = 5
        if q_time > 600:
            q_time = 600
    except:
        q_time = 60
    await state.update_data(question_time_seconds=q_time)
    await state.set_state(AdminAddTest.waiting_for_attempts)
    await message.answer(f"✅ Har bir savol uchun {q_time} soniya berildi.\n\nMaksimal urinishlar sonini kiriting (masalan: 1):")

@router.message(AdminAddTest.waiting_for_block_sub1)
async def admin_get_block_sub1(message: Message, state: FSMContext):
    await state.update_data(block_sub1=message.text.strip())
    await state.set_state(AdminAddTest.waiting_for_block_sub2)
    await message.answer("2-asosiy fanning nomini kiriting (masalan: Ingliz tili yoki Kimyo):")

@router.message(AdminAddTest.waiting_for_block_sub2)
async def admin_get_block_sub2(message: Message, state: FSMContext):
    await state.update_data(block_sub2=message.text.strip())
    await state.set_state(AdminAddTest.waiting_for_attempts)
    await message.answer("Maksimal urinishlar sonini kiriting (masalan: 1):")

@router.message(AdminAddTest.waiting_for_attempts)
async def admin_get_attempts(message: Message, state: FSMContext):
    try: att = int(message.text.strip())
    except: att = 1
    await state.update_data(max_attempts=att)
    await state.set_state(AdminAddTest.waiting_for_start_time)
    await message.answer("Test boshlanish vaqtini kiriting (Format: `YYYY-MM-DD HH:MM` yoki `-`):")

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
            "🧩 <b>Blok test savollarini yuboring:</b>\n"
            "Tartib bo'yicha:\n1) Tarix (10 ta)\n2) Ona tili (10 ta)\n3) Matematika (10 ta)\n4) 1-asosiy fan (30 ta)\n5) 2-asosiy fan (30 ta)\n\n"
            "Matn yoki Word/PDF fayl ko'rinishida yuborishingiz mumkin.", reply_markup=get_finish_test_keyboard()
        )
    else:
        await message.answer("📂 Savollarni matn yoki Word/PDF fayl ko'rinishida yuborishingiz mumkin.", reply_markup=get_finish_test_keyboard())

@router.message(AdminAddTest.waiting_for_questions, F.photo)
async def admin_handle_photo_question(message: Message, state: FSMContext):
    if not await is_admin(message.from_user.id):
        return
    photo_id = message.photo[-1].file_id
    caption = (message.caption or "").strip()
    data = await state.get_data()
    q_list = data.get("questions", [])

    parsed = parse_single_question_text(caption) if caption else None
    if parsed:
        parsed["photo_file_id"] = photo_id
        q_list.append(parsed)
        await state.update_data(questions=q_list)
        await message.answer(f"📸 Rasmli savol qo'shildi! Jami savollar: {len(q_list)}")
    elif caption:
        # Caption bor lekin A/B topilmadi — savol matni sifatida qabul qilamiz, variantlar default
        q_list.append({
            "text": caption,
            "a": "A",
            "b": "B",
            "c": "C",
            "d": "D",
            "correct": "A",
            "photo_file_id": photo_id
        })
        await state.update_data(questions=q_list)
        await message.answer(
            f"📸 Rasmli savol qo'shildi (variantlar A–D default).\n"
            f"Jami: {len(q_list)}\n"
            f"⚠️ Keyin to'g'ri javoblarni alohida yuborasiz."
        )
    else:
        # Faqat rasm — keyingi matn bilan bog'lash uchun vaqtincha saqlash
        await state.update_data(pending_photo_id=photo_id)
        await message.answer(
            "📸 Rasm qabul qilindi.\n"
            "Endi shu rasm uchun savol matni va variantlarni yuboring:\n\n"
            "<code>Savol matni...\n"
            "A) variant\n"
            "B) variant\n"
            "C) variant\n"
            "D) variant</code>"
        )

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
    if not await is_admin(message.from_user.id):
        return
    if message.text in ("🏠 Asosiy menyu", "⬅️ Bosh menyu", "🚀 Start"):
        await back_to_menu(message, state)
        return
    if message.text == "✅ Testni yakunlash va saqlash":
        return

    data = await state.get_data()
    pending_photo = data.get("pending_photo_id")

    # Oldingi rasm bilan bog'langan bitta savol
    if pending_photo:
        parsed = parse_single_question_text(message.text)
        q_list = data.get("questions", [])
        if parsed:
            parsed["photo_file_id"] = pending_photo
            q_list.append(parsed)
        else:
            q_list.append({
                "text": message.text.strip(),
                "a": "A", "b": "B", "c": "C", "d": "D",
                "correct": "A",
                "photo_file_id": pending_photo
            })
        await state.update_data(questions=q_list, pending_photo_id=None)
        await message.answer(f"📸 Rasmli savol qo'shildi! Jami: {len(q_list)}")
        return

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
    if not await is_admin(message.from_user.id):
        return
    if message.text in ("🏠 Asosiy menyu", "⬅️ Bosh menyu", "🚀 Start"):
        await back_to_menu(message, state)
        return
    tokens = re.findall(r'[A-D]', message.text.upper())
    data = await state.get_data()
    questions_list = data.get("questions", [])
    q_count = len(questions_list)
    a_count = len(tokens)

    if a_count != q_count:
        await message.answer(
            f"❌ <b>Javoblar soni mos kelmadi!</b>\n\n"
            f"Savollar: <b>{q_count}</b> ta\n"
            f"Siz yuborgan javoblar: <b>{a_count}</b> ta\n\n"
            f"Iltimos, to'g'ri javoblarni qayta yuboring "
            f"(masalan: <code>A B C D A ...</code> — jami {q_count} ta harf):"
        )
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
                # Majburiy fanlar uchun doimiy kalitlar (callback bilan mos)
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
            
        if data.get("start_time"):
            students = (await session.execute(select(Student).where(Student.grade == data["grade"], Student.telegram_id.is_not(None)))).scalars().all()
            for st in students:
                session.add(Reminder(test_id=new_test.id, student_id=st.id, reminded=False))
                
        await session.commit()
    await state.clear()
    await message.answer("✅ Blok test muvaffaqiyatli saqlandi! Ballar va fanlar talabga moslab taqsimlandi.", reply_markup=get_admin_menu())

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
        question = await session.get(Question, appeal.question_id)
        add_points = question.points if question else 1.0
        
        ts.score = round(ts.score + add_points, 2)
        ts.correct_answers += 1
        if ts.wrong_answers > 0:
            ts.wrong_answers -= 1
        
        # Max ball asosida foizni qayta hisoblash
        max_possible = await session.scalar(select(func.sum(Question.points)).where(Question.test_id == ts.test_id)) or 1
        ts.score_percentage = round((ts.score / max_possible) * 100, 2)
            
        await session.commit()
        
        student = await session.get(Student, appeal.student_id)
        test = await session.get(Test, ts.test_id)
        
        update_result_in_sheet(student.student_id, test.title, ts.score, ts.score_percentage, ts.correct_answers)
        
        if student and student.telegram_id:
            try:
                await bot.send_message(
                    chat_id=student.telegram_id,
                    text=f"✅ <b>Apellyatsiyangiz ma'qullandi!</b>\n\nBalingizga <b>+{add_points}</b> ball qo'shildi.\nHozirgi balingiz: <b>{ts.score}</b> ({ts.score_percentage}%)"
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
    if message.from_user.id not in SUPER_ADMIN_IDS: 
        await message.answer("⚠️ Bu bo'limga faqat Super Admin kira oladi!")
        return
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
    if callback.from_user.id not in SUPER_ADMIN_IDS:
        await callback.answer("Sizda bu huquq yo'q!", show_alert=True)
        return
    await state.set_state(AdminManageAdmins.waiting_for_id)
    await callback.message.answer("Yangi adminning Telegram ID raqamini kiriting:")
    await callback.answer()

@router.message(AdminManageAdmins.waiting_for_id)
async def save_new_admin(message: Message, state: FSMContext):
    if message.from_user.id not in SUPER_ADMIN_IDS: return
    try:
        tg_id = int(message.text.strip())
        async with async_session() as session:
            existing = (await session.execute(select(Admin).where(Admin.telegram_id == tg_id))).scalar_one_or_none()
            if existing:
                await message.answer("⚠️ Bu foydalanuvchi allaqachon admin ro'yxatida mavjud!", reply_markup=get_admin_menu())
                await state.clear()
                return

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

@router.message(F.text.in_(["🏠 Asosiy menyu", "⬅️ Bosh menyu", "🚀 Start"]))
async def back_to_menu(message: Message, state: FSMContext, bot: Bot = None):
    """🚀 Start yoki Asosiy menyu — /start kabi qayta ishga tushirish."""
    was_in_test = await state.get_state() == TestProcessState.in_test.state
    await state.clear()
    user_next_question_flags.pop(message.from_user.id, None)
    if was_in_test:
        user_abort_test_flags[message.from_user.id] = True
        async with async_session() as session:
            student = (await session.execute(
                select(Student).where(Student.telegram_id == message.from_user.id)
            )).scalar_one_or_none()
            if student:
                active = (await session.execute(
                    select(TestSession).where(
                        TestSession.student_id == student.id,
                        TestSession.status == "IN_PROGRESS"
                    )
                )).scalars().all()
                for ts in active:
                    ts.status = "COMPLETED"
                    ts.finished_at = datetime.now(timezone.utc)
                    await calculate_and_save_results(session, ts)
                await session.commit()

    # Admin
    if await is_admin(message.from_user.id):
        await message.answer("🛠 <b>Xush kelibsiz, Admin!</b>", reply_markup=get_admin_menu())
        return

    # Obuna tekshiruvi (bot bo'lsa)
    if bot is not None:
        is_subscribed = await check_subscription(message.from_user.id, bot)
        if not is_subscribed:
            keyboard = InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text="📢 Kanalga obuna bo'lish", url=f"https://t.me/{REQUIRED_CHANNEL.replace('@', '')}")],
                [InlineKeyboardButton(text="✅ Obunani tekshirish", callback_data="check_sub")]
            ])
            await message.answer(
                f"⚠️ Botdan foydalanish uchun avval quyidagi kanalga obuna bo'lishingiz kerak:\n\n{REQUIRED_CHANNEL}",
                reply_markup=keyboard
            )
            return

    async with async_session() as session:
        student = (await session.execute(
            select(Student).where(Student.telegram_id == message.from_user.id)
        )).scalar_one_or_none()
    if student:
        if not student.is_active:
            await message.answer("❌ Sizning profilingiz administrator tomonidan bloklangan.")
            return
        main_menu = await get_main_menu_keyboard()
        note = "\n\n⚠️ Test jarayoni to'xtatildi va natija saqlandi." if was_in_test else ""
        await message.answer(
            f"🏠 <b>Asosiy menyu</b>\n\nXush kelibsiz, <b>{student.first_name} {student.last_name}</b>!\n"
            f"Sinfingiz: <b>{student.grade or 'Nomaʼlum'}</b>{note}",
            reply_markup=main_menu
        )
    else:
        await state.set_state(SelfRegState.waiting_for_fullname)
        await message.answer("🎓 <b>Olimpiada tizimiga xush kelibsiz!</b>\n\nIltimos, to'liq <b>Ism va Familiyangizni</b> kiriting:")

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
