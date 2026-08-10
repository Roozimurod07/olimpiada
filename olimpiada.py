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

# Google Sheets kutubxonalari
import gspread
from google.oauth2.service_account import Credentials

# --- RAILWAY VA GOOGLE SHEETS SOZLAMALARI ---
BOT_TOKEN = os.getenv("BOT_TOKEN")
GOOGLE_CREDS_JSON = os.getenv("GOOGLE_CREDS_JSON")
SHEET_NAME = "test"

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

class Test(Base):
    __tablename__ = "tests"
    id = Column(Integer, primary_key=True, autoincrement=True)
    title = Column(String(200), nullable=False)
    subject = Column(String(100), nullable=False)
    grade_level = Column(String(20), nullable=False)
    max_attempts = Column(Integer, default=1)
    duration_seconds_per_question = Column(Integer, default=15)
    is_active = Column(Boolean, default=True)
    
    questions = relationship("Question", back_populates="test", cascade="all, delete-orphan")
    test_sessions = relationship("TestSession", back_populates="test", cascade="all, delete-orphan")

class Question(Base):
    __tablename__ = "questions"
    id = Column(Integer, primary_key=True, autoincrement=True)
    test_id = Column(Integer, ForeignKey("tests.id", ondelete="CASCADE"), nullable=False)
    question_text = Column(Text, nullable=False)
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
        try:
            await conn.execute(sa_text("ALTER TABLE tests ADD COLUMN grade_level TEXT DEFAULT '11-sinf';"))
            await conn.execute(sa_text("ALTER TABLE tests ADD COLUMN max_attempts INTEGER DEFAULT 1;"))
        except Exception:
            pass

# --- GOOGLE SHEETS GA YOZISH FUNKSIYASI ---
def save_result_to_sheet(student_id, full_name, school, grade, test_title, subject, score, percentage, correct, wrong):
    try:
        scope = ["https://www.googleapis.com/auth/spreadsheets"]
        if GOOGLE_CREDS_JSON:
            creds_dict = json.loads(GOOGLE_CREDS_JSON)
            creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
        else:
            creds = Credentials.from_service_account_file('credentials.json', scopes=scope)
            
        client = gspread.authorize(creds)
        spreadsheet = client.open(SHEET_NAME)
        sheet = spreadsheet.sheet1
        
        # Jadvalga qator qo'shish
        row_data = [
            str(student_id),
            str(full_name),
            str(school),
            str(grade),
            str(subject),
            str(test_title),
            str(score),
            f"{percentage}%",
            str(correct),
            str(wrong),
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ]
        sheet.append_row(row_data)
        print("✅ Google Sheets'ga natija muvaffaqiyatli yozildi!")
    except Exception as e:
        print(f"❌ Google Sheets xatosi: {e}")

class RegState(StatesGroup):
    waiting_for_id = State()

class AdminAddStudent(StatesGroup):
    waiting_for_data = State()
    waiting_for_excel = State()

class AdminAddTest(StatesGroup):
    waiting_for_title = State()
    waiting_for_subject = State()
    waiting_for_grade = State()
    waiting_for_attempts = State()
    waiting_for_duration = State()
    waiting_for_questions = State()
    waiting_for_answers = State()

class AdminBroadcast(StatesGroup):
    waiting_for_message = State()

class TestProcessState(StatesGroup):
    in_test = State()

router = Router()

def get_main_menu():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="📝 Testni boshlash"), KeyboardButton(text="👤 Profilim")],
            [KeyboardButton(text="📊 Mening urinishlarim"), KeyboardButton(text="🏆 Reyting")],
            [KeyboardButton(text="ℹ️ Olimpiada haqida")]
        ],
        resize_keyboard=True
    )

def get_admin_menu():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="➕ ID qo'shish"), KeyboardButton(text="📂 Excel orqali ID'lar")],
            [KeyboardButton(text="📂 Test yuklash"), KeyboardButton(text="⚙️ Testlarni boshqarish")],
            [KeyboardButton(text="📊 Jonli statistika"), KeyboardButton(text="📥 Excel natijalar")],
            [KeyboardButton(text="🧹 Bazani tozalash"), KeyboardButton(text="📢 Xabar yuborish")],
            [KeyboardButton(text="⬅️ Bosh menyu")]
        ],
        resize_keyboard=True
    )

def get_finish_test_keyboard():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="✅ Javobni yuklash / Testni saqlash")]
        ],
        resize_keyboard=True
    )

@router.message(Command("start"))
async def cmd_start(message: Message, state: FSMContext):
    await state.clear()
    if message.from_user.id in SUPER_ADMIN_IDS:
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
    current_state = await state.get_state()
    if current_state == TestProcessState.in_test.state:
        await message.answer("⚠️ Test jarayonida boshqa menyuga o'tib bo'lmaydi!")
        return
    async with async_session() as session:
        student = (await session.execute(select(Student).where(Student.telegram_id == message.from_user.id))).scalar_one_or_none()
        if not student:
            await message.answer("Siz ro'yxatdan o'tmagansiz. /start ni bosing.")
            return
        await message.answer(f"👤 <b>Profil:</b>\n\nID: <code>{student.student_id}</code>\nIsm: {student.first_name} {student.last_name}\nMaktab: {student.school or '-'}\nSinf: {student.grade or '-'}")

@router.message(F.text == "📊 Mening urinishlarim")
async def my_attempts_handler(message: Message, state: FSMContext):
    current_state = await state.get_state()
    if current_state == TestProcessState.in_test.state:
        await message.answer("⚠️ Test jarayonida boshqa menyuga o'tib bo'lmaydi!")
        return
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
            
        keyboard_buttons = []
        for ts, t in sessions:
            date_str = ts.finished_at.strftime("%Y-%m-%d %H:%M") if ts.finished_at else ""
            keyboard_buttons.append([InlineKeyboardButton(
                text=f"📚 {t.subject} | {ts.score} ball ({ts.score_percentage}%) [{date_str}]",
                callback_data=f"attempt_detail_{ts.id}"
            )])
            
        markup = InlineKeyboardMarkup(inline_keyboard=keyboard_buttons)
        await message.answer("📊 <b>Sizning ishlagan testlaringiz tarixi:</b>\nTafsilot va xatolar tahlilini ko'rish uchun testni tanlang:", reply_markup=markup)

@router.callback_query(F.data.startswith("attempt_detail_"))
async def show_attempt_detail(callback: CallbackQuery):
    session_id = int(callback.data.split("_")[2])
    async with async_session() as session:
        ts = await session.get(TestSession, session_id)
        test = await session.get(Test, ts.test_id)
        questions = (await session.execute(select(Question).where(Question.test_id == test.id))).scalars().all()
        answers = {a.question_id: a.selected_option for a in (await session.execute(select(Answer).where(Answer.session_id == ts.id))).scalars().all()}
        
        text = f"📋 <b>Test tahlili: {test.subject} ({test.title})</b>\n" \
               f"⭐ Ball: {ts.score} ({ts.score_percentage}%)\n" \
               f"✅ To'g'ri: {ts.correct_answers} | ❌ Noto'g'ri: {ts.wrong_answers} | ⭕ Javobsiz: {ts.unanswered}\n\n"
               
        for idx, q in enumerate(questions, 1):
            sel = answers.get(q.id, "Javob berilmagan")
            status = "✅" if sel == q.correct_option else "❌"
            text += f"<b>{idx}. {q.question_text}</b>\n" \
                    f"Sizning javob: <b>{sel}</b> {status} | To'g'ri javob: <b>{q.correct_option}</b>\n\n"
                    
        if len(text) > 4096:
            text = text[:4000] + "\n... (matn uzunlik chegarasidan oshdi)"
            
        await callback.message.edit_text(text)
        await callback.answer()

user_next_question_flags = {}

@router.message(F.text == "📝 Testni boshlash")
async def start_test_prompt(message: Message, state: FSMContext):
    current_state = await state.get_state()
    if current_state == TestProcessState.in_test.state:
        await message.answer("⚠️ Siz hozir test ishlayapsiz!")
        return
    async with async_session() as session:
        student = (await session.execute(select(Student).where(Student.telegram_id == message.from_user.id))).scalar_one_or_none()
        if not student:
            await message.answer("Iltimos, oldin /start orqali ro'yxatdan o'ting.")
            return

        if not student.grade:
            await message.answer("❌ Sizning profilingizda sinfingiz ko'rsatilmagan. Administratsiyaga murojaat qiling.")
            return

        tests = (await session.execute(
            select(Test).where(Test.is_active == True, Test.grade_level == student.grade)
        )).scalars().all()
        
        if not tests:
            await message.answer(f"⚠️ Hozirda <b>{student.grade}</b> uchun faol testlar mavjud emas.")
            return

        keyboard_buttons = []
        for t in tests:
            keyboard_buttons.append([InlineKeyboardButton(text=f"📚 {t.subject} — {t.title}", callback_data=f"start_test_{t.id}")])
        
        markup = InlineKeyboardMarkup(inline_keyboard=keyboard_buttons)
        await message.answer(f"📝 <b>Sizning sinfingiz ({student.grade}) uchun mavjud testlar:</b>\n\nIltimos, fanni tanlang:", reply_markup=markup)

@router.callback_query(F.data.startswith("start_test_"))
async def begin_test_session(callback: CallbackQuery, state: FSMContext, bot: Bot):
    test_id = int(callback.data.split("_")[2])
    
    async with async_session() as session:
        student = (await session.execute(select(Student).where(Student.telegram_id == callback.from_user.id))).scalar_one_or_none()
        test = (await session.execute(select(Test).where(Test.id == test_id))).scalar_one_or_none()
        
        if not test or not test.is_active:
            await callback.answer("Bu test topilmadi yoki faol emas!", show_alert=True)
            return

        if test.grade_level != student.grade:
            await callback.answer("❌ Bu test sizning sinfingizga mos kelmaydi!", show_alert=True)
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
                await callback.answer(f"❌ Siz bu testni allaqachon topshirgansiz! Ruxsat etilgan urinishlar soni: {test.max_attempts} ta.", show_alert=True)
                return

        questions = (await session.execute(select(Question).where(Question.test_id == test_id))).scalars().all()
        if not questions:
            await callback.answer("Bu testda savollar mavjud emas!", show_alert=True)
            return
            
        random.shuffle(questions)
        
        test_session = TestSession(student_id=student.id, test_id=test_id, status="IN_PROGRESS")
        session.add(test_session)
        await session.commit()
        await session.refresh(test_session)

        duration_per_q = test.duration_seconds_per_question if test.duration_seconds_per_question else 15
        await callback.message.edit_text(f"🚀 <b>{test.subject}</b> testi boshlandi!\nHar bir savolga {duration_per_q} soniya beriladi.")
        
        user_id = callback.from_user.id
        await state.set_state(TestProcessState.in_test)
        
        for index, q in enumerate(questions):
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
            if row:
                keyboard_buttons.append(row)
            
            keyboard_buttons.append([InlineKeyboardButton(text="➡️ Keyingi savol", callback_data=f"next_q_{test_session.id}")])
            markup = InlineKeyboardMarkup(inline_keyboard=keyboard_buttons)
            
            remaining_time = duration_per_q
            user_next_question_flags[user_id] = False
            
            q_msg = await bot.send_message(
                chat_id=user_id,
                text=f"<b>Savol {index + 1} / {len(questions)}</b> (⏱ Qolgan vaqt: {remaining_time}s)\n\n{q.question_text}",
                reply_markup=markup
            )
            
            for _ in range(duration_per_q):
                await asyncio.sleep(1)
                if user_next_question_flags.get(user_id, False):
                    break
                
                remaining_time -= 1
                try:
                    await bot.edit_message_text(
                        chat_id=user_id,
                        message_id=q_msg.message_id,
                        text=f"<b>Savol {index + 1} / {len(questions)}</b> (⏱ Qolgan vaqt: {remaining_time}s)\n\n{q.question_text}",
                        reply_markup=markup
                    )
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
                sess.finished_at = datetime.utcnow()
                await calculate_and_save_results(final_session, sess)
                await final_session.commit()
                
                # Google Sheets'ga natijani yuborish uchun ma'lumotlarni yig'amiz
                student_obj = await final_session.get(Student, sess.student_id)
                test_obj = await final_session.get(Test, sess.test_id)
                if student_obj and test_obj:
                    save_result_to_sheet(
                        student_id=student_obj.student_id,
                        full_name=f"{student_obj.first_name} {student_obj.last_name}",
                        school=student_obj.school or "-",
                        grade=student_obj.grade or "-",
                        test_title=test_obj.title,
                        subject=test_obj.subject,
                        score=sess.score,
                        percentage=sess.score_percentage,
                        correct=sess.correct_answers,
                        wrong=sess.wrong_answers
                    )
                
                await state.clear()
                await bot.send_message(
                    chat_id=user_id,
                    text=f"🏆 <b>TEST YAKUNLANDI!</b>\n\n"
                         f"✅ To'g'ri: {sess.correct_answers}\n"
                         f"❌ Noto'g'ri: {sess.wrong_answers}\n"
                         f"⭕ Javobsiz: {sess.unanswered}\n"
                         f"📊 Foiz: {sess.score_percentage}%\n"
                         f"⭐ Ball: {sess.score}",
                    reply_markup=get_main_menu()
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
        
    user_next_question_flags[callback.from_user.id] = True
    await callback.answer(f"Tanlandi: {selected}")
    
    try:
        await bot.delete_message(chat_id=callback.message.chat.id, message_id=callback.message.message_id)
    except Exception:
        pass

@router.callback_query(F.data.startswith("next_q_"))
async def next_question_callback(callback: CallbackQuery, bot: Bot):
    user_next_question_flags[callback.from_user.id] = True
    await callback.answer("Keyingi savolga o'tilmoqda...")
    try:
        await bot.delete_message(chat_id=callback.message.chat.id, message_id=callback.message.message_id)
    except Exception:
        pass

async def calculate_and_save_results(session, sess: TestSession):
    questions = (await session.execute(select(Question).where(Question.test_id == sess.test_id))).scalars().all()
    answers = {a.question_id: a.selected_option for a in (await session.execute(select(Answer).where(Answer.session_id == sess.id))).scalars().all()}
    
    correct, wrong, unanswered, total_score = 0, 0, 0, 0.0
    for q in questions:
        sel = answers.get(q.id)
        if not sel:
            unanswered += 1
        elif sel == q.correct_option:
            correct += 1
            total_score += q.points
        else:
            wrong += 1
            
    total_q = len(questions)
    sess.correct_answers = correct
    sess.wrong_answers = wrong
    sess.unanswered = unanswered
    sess.score = total_score
    sess.score_percentage = round((correct / total_q) * 100, 2) if total_q > 0 else 0.0

@router.message(F.text == "🏆 Reyting")
async def rating_menu_prompt(message: Message, state: FSMContext):
    current_state = await state.get_state()
    if current_state == TestProcessState.in_test.state:
        await message.answer("⚠️ Test paytida reytingni ko'rib bo'lmaydi!")
        return
    async with async_session() as session:
        tests = (await session.execute(select(Test))).scalars().all()
        if not tests:
            await message.answer("🏆 Hozircha testlar va reyting mavjud emas.")
            return
        
        keyboard_buttons = []
        for t in tests:
            keyboard_buttons.append([InlineKeyboardButton(text=f"📊 [{t.grade_level}] {t.subject} — {t.title}", callback_data=f"show_rating_{t.id}")])
        
        markup = InlineKeyboardMarkup(inline_keyboard=keyboard_buttons)
        await message.answer("🏆 <b>Qaysi testning reytingini ko'rmoqchisiz?</b>\n\nIltimos, testni tanlang:", reply_markup=markup)

@router.callback_query(F.data.startswith("show_rating_"))
async def show_specific_test_rating(callback: CallbackQuery):
    test_id = int(callback.data.split("_")[2])
    
    async with async_session() as session:
        test = await session.get(Test, test_id)
        if not test:
            await callback.answer("Test topilmadi!", show_alert=True)
            return
            
        rows = (await session.execute(
            select(Student, TestSession)
            .join(TestSession, Student.id == TestSession.student_id)
            .where(TestSession.test_id == test_id, TestSession.status == "COMPLETED")
            .order_by(TestSession.score.desc())
            .limit(15)
        )).all()
        
        if not rows:
            await callback.message.edit_text(f"🏆 <b>[{test.grade_level}] {test.subject} ({test.title})</b>\n\nBu test bo'yicha hali natijalar mavjud emas.")
            return
            
        text = f"🏆 <b>REYTING: [{test.grade_level}] {test.subject} ({test.title})</b>\n\n"
        for idx, (s, ts) in enumerate(rows, 1):
            medal = "🥇" if idx == 1 else ("🥈" if idx == 2 else ("🥉" if idx == 3 else f"{idx}."))
            text += f"{medal} {s.first_name} {s.last_name} ({s.grade}) — <b>{ts.score} ball</b> ({ts.score_percentage}%)\n"
            
        await callback.message.edit_text(text)

@router.message(F.text == "ℹ️ Olimpiada haqida")
async def about_handler(message: Message, state: FSMContext):
    current_state = await state.get_state()
    if current_state == TestProcessState.in_test.state:
        return
    await message.answer("ℹ️ Professional Olimpiada Tizimi. Sinf kesimidagi testlar va qat'iy nazorat platformasi.")

@router.message(Command("admin"))
async def admin_panel(message: Message, state: FSMContext):
    if message.from_user.id not in SUPER_ADMIN_IDS:
        return
    await state.clear()
    await message.answer("🛠 <b>Admin Panel</b>", reply_markup=get_admin_menu())

@router.message(F.text == "➕ ID qo'shish")
async def admin_add_student_prompt(message: Message, state: FSMContext):
    if message.from_user.id not in SUPER_ADMIN_IDS:
        return
    await state.set_state(AdminAddStudent.waiting_for_data)
    await message.answer("📝 Ma'lumotlarni yuboring:\n<code>Ism, Familiya, Sinf, Maktab</code>\n(Masalan: <i>Alisher, Valiyev, 11-sinf, 12-maktab</i>)")

@router.message(AdminAddStudent.waiting_for_data)
async def admin_save_student(message: Message, state: FSMContext):
    if message.from_user.id not in SUPER_ADMIN_IDS:
        return
    parts = [p.strip() for p in message.text.split(",")]
    if len(parts) < 4:
        await message.answer("❌ Format xato! Masalan: <code>Alisher, Valiyev, 11-sinf, 12-maktab</code>")
        return
    unique_id = f"OLM-2026-{''.join(random.choices('ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789', k=6))}"
    async with async_session() as session:
        session.add(Student(student_id=unique_id, first_name=parts[0], last_name=parts[1], grade=parts[2], school=parts[3]))
        await session.commit()
    await state.clear()
    await message.answer(f"✅ O'quvchi qo'shildi!\nID: <code>{unique_id}</code>", reply_markup=get_admin_menu())

@router.message(F.text == "📂 Excel orqali ID'lar")
async def admin_excel_students_prompt(message: Message, state: FSMContext):
    if message.from_user.id not in SUPER_ADMIN_IDS:
        return
    await state.set_state(AdminAddStudent.waiting_for_excel)
    await message.answer(
        "📂 O'quvchilar ro'yxati bor Excel faylni (`.xlsx`) yuboring.\n"
        "Fayl ustunlari tartibi quyidagicha bo'lishi kerak:\n"
        "<code>Ism | Familiya | Sinf (masalan: 11-sinf) | Maktab</code>"
    )

@router.message(AdminAddStudent.waiting_for_excel, F.document)
async def admin_process_excel_students(message: Message, state: FSMContext, bot: Bot):
    if message.from_user.id not in SUPER_ADMIN_IDS:
        return
    document = message.document
    file_info = await bot.get_file(document.file_id)
    downloaded = await bot.download_file(file_info.file_path)
    
    try:
        df = pd.read_excel(io.BytesIO(downloaded.read()))
        added_count = 0
        async with async_session() as session:
            for _, row in df.iterrows():
                name, surname, grade, school = str(row.iloc[0]), str(row.iloc[1]), str(row.iloc[2]), str(row.iloc[3])
                unique_id = f"OLM-2026-{''.join(random.choices('ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789', k=6))}"
                session.add(Student(student_id=unique_id, first_name=name, last_name=surname, grade=grade, school=school))
                added_count += 1
            await session.commit()
        await state.clear()
        await message.answer(f"✅ Exceldan {added_count} ta o'quvchi muvaffaqiyatli qo'shildi!", reply_markup=get_admin_menu())
    except Exception as e:
        await message.answer(f"❌ Xatolik yuz berdi: {e}")

@router.message(F.text == "📂 Test yuklash")
async def admin_add_test_start(message: Message, state: FSMContext):
    if message.from_user.id not in SUPER_ADMIN_IDS:
        return
    await state.set_state(AdminAddTest.waiting_for_title)
    await message.answer("📂 <b>1-qadam:</b> Test sarlavhasini kiriting (masalan: <i>Respublika Olimpiadasi</i>):")

@router.message(AdminAddTest.waiting_for_title)
async def admin_get_title(message: Message, state: FSMContext):
    await state.update_data(title=message.text.strip())
    await state.set_state(AdminAddTest.waiting_for_subject)
    await message.answer("📂 <b>2-qadam:</b> Fan nomini kiriting (masalan: <i>Matematika</i>):")

@router.message(AdminAddTest.waiting_for_subject)
async def admin_get_subject(message: Message, state: FSMContext):
    await state.update_data(subject=message.text.strip())
    await state.set_state(AdminAddTest.waiting_for_grade)
    await message.answer("📂 <b>3-qadam:</b> Qaysi sinf uchunligini kiriting (Masalan: <code>11-sinf</code>, <code>9-sinf</code>, <code>8-sinf</code>, <code>5-sinf</code>):")

@router.message(AdminAddTest.waiting_for_grade)
async def admin_get_grade(message: Message, state: FSMContext):
    await state.update_data(grade=message.text.strip())
    await state.set_state(AdminAddTest.waiting_for_attempts)
    await message.answer(
        "📂 <b>4-qadam:</b> O'quvchi bu testni necha marta ishlashi mumkinligini kiriting:\n"
        "• <code>1</code> — Faqat 1 marta\n"
        "• <code>2</code> yoki ko'proq — Aniq bir necha marta\n"
        "• <code>0</code> — Cheksiz marta"
    )

@router.message(AdminAddTest.waiting_for_attempts)
async def admin_get_attempts(message: Message, state: FSMContext):
    try:
        attempts = int(message.text.strip())
    except ValueError:
        attempts = 1
    await state.update_data(max_attempts=attempts)
    await state.set_state(AdminAddTest.waiting_for_duration)
    await message.answer("📂 <b>5-qadam:</b> Har bir savolga beriladigan vaqtni sekundlarda kiriting (masalan: <code>15</code>):")

@router.message(AdminAddTest.waiting_for_duration)
async def admin_get_duration(message: Message, state: FSMContext):
    try:
        dur = int(message.text.strip())
    except ValueError:
        dur = 15
    await state.update_data(duration=dur, questions=[])
    await state.set_state(AdminAddTest.waiting_for_questions)
    await message.answer(
        "📂 <b>6-qadam:</b> Savollarni **bittada** nusxalab tashlang yoki **PDF/Word** fayl yuboring.\n\n"
        "Tugatgach, pastdagi **✅ Javobni yuklash / Testni saqlash** tugmasini bosing.",
        reply_markup=get_finish_test_keyboard()
    )

@router.message(AdminAddTest.waiting_for_questions, F.document)
async def admin_handle_document(message: Message, state: FSMContext, bot: Bot):
    if message.from_user.id not in SUPER_ADMIN_IDS:
        return
    document = message.document
    file_ext = document.file_name.split('.')[-1].lower()
    file_info = await bot.get_file(document.file_id)
    downloaded_file = await bot.download_file(file_info.file_path)
    
    file_path = f"temp_{document.file_name}"
    with open(file_path, "wb") as f:
        f.write(downloaded_file.read())
        
    extracted_text = ""
    try:
        if file_ext == "pdf":
            reader = pypdf.PdfReader(file_path)
            for page in reader.pages:
                extracted_text += page.extract_text() + "\n"
        elif file_ext in ["docx", "doc"]:
            doc = docx.Document(file_path)
            for para in doc.paragraphs:
                extracted_text += para.text + "\n"
    except Exception as e:
        await message.answer(f"❌ Faylni o'qishda xatolik: {e}")
        return
    finally:
        if os.path.exists(file_path):
            os.remove(file_path)
            
    added_count = await parse_and_add_questions(extracted_text, state)
    data = await state.get_data()
    total_q = len(data.get("questions", []))
    await message.answer(f"✅ Fayldan {added_count} ta savol qo'shildi!\nJami: {total_q} ta.")

@router.message(AdminAddTest.waiting_for_questions, F.text == "✅ Javobni yuklash / Testni saqlash")
async def admin_ask_for_answers(message: Message, state: FSMContext):
    if message.from_user.id not in SUPER_ADMIN_IDS:
        return
    data = await state.get_data()
    questions_list = data.get("questions", [])
    
    if not questions_list:
        await message.answer("❌ Hech qanday savol kiritilmadi!")
        return

    await state.set_state(AdminAddTest.waiting_for_answers)
    await message.answer(
        f"✅ Jami **{len(questions_list)} ta** savol qabul qilindi.\n\n"
        f"🔑 <b>To'g'ri javoblarni yuboring:</b>\n"
        f"<code>A B C D A B C D A B...</code>",
        reply_markup=get_admin_menu()
    )

@router.message(AdminAddTest.waiting_for_questions, F.text)
async def admin_add_bulk_questions_text(message: Message, state: FSMContext):
    if message.from_user.id not in SUPER_ADMIN_IDS:
        return
    added_count = await parse_and_add_questions(message.text, state)
    data = await state.get_data()
    total_q = len(data.get("questions", []))
    await message.answer(f"✅ {added_count} ta savol qo'shildi!\nJami: {total_q} ta.")

async def parse_and_add_questions(text: str, state: FSMContext) -> int:
    data = await state.get_data()
    questions_list = data.get("questions", [])
    
    lines = [line.strip() for line in text.split("\n") if line.strip()]
    added = 0
    i = 0
    while i < len(lines):
        a_idx, b_idx, c_idx, d_idx = -1, -1, -1, -1
        for j in range(i + 1, min(i + 6, len(lines))):
            l_lower = lines[j].lower()
            if l_lower.startswith("a)") or l_lower.startswith("a."): a_idx = j
            elif l_lower.startswith("b)") or l_lower.startswith("b."): b_idx = j
            elif l_lower.startswith("c)") or l_lower.startswith("c."): c_idx = j
            elif l_lower.startswith("d)") or l_lower.startswith("d."): d_idx = j

        if a_idx != -1 and b_idx != -1:
            q_text = " ".join(lines[i:a_idx])
            q_text = re.sub(r'^\d+[\.\)]\s*', '', q_text)
            opt_a = re.sub(r'^[aA][\.\)]\s*', '', lines[a_idx])
            opt_b = re.sub(r'^[bB][\.\)]\s*', '', lines[b_idx])
            opt_c = re.sub(r'^[cC][\.\)]\s*', '', lines[c_idx]) if c_idx != -1 else "Variant C"
            opt_d = re.sub(r'^[dD][\.\)]\s*', '', lines[d_idx]) if d_idx != -1 else "Variant D"
            
            questions_list.append({"text": q_text, "a": opt_a, "b": opt_b, "c": opt_c, "d": opt_d, "correct": "A"})
            added += 1
            i = max(a_idx, b_idx, c_idx if c_idx != -1 else 0, d_idx if d_idx != -1 else 0) + 1
        else:
            if i + 3 < len(lines):
                q_text = lines[i]
                opt_a = re.sub(r'^[aA][\.\)]\s*', '', lines[i+1])
                opt_b = re.sub(r'^[bB][\.\)]\s*', '', lines[i+2])
                opt_c = re.sub(r'^[cC][\.\)]\s*', '', lines[i+3])
                opt_d = re.sub(r'^[dD][\.\)]\s*', '', lines[i+4]) if i + 4 < len(lines) else "Variant D"
                
                questions_list.append({"text": q_text, "a": opt_a, "b": opt_b, "c": opt_c, "d": opt_d, "correct": "A"})
                added += 1
                i += 5
            else:
                i += 1

    await state.update_data(questions=questions_list)
    return added

@router.message(AdminAddTest.waiting_for_answers)
async def admin_save_answers_and_test(message: Message, state: FSMContext):
    if message.from_user.id not in SUPER_ADMIN_IDS:
        return
    text = message.text.upper()
    data = await state.get_data()
    questions_list = data.get("questions", [])
    
    tokens = re.findall(r'[A-D]', text)
    if not tokens:
        await message.answer("❌ To'g'ri javoblar topilmadi! Faqat A, B, C, D harflarini yuboring.")
        return

    for idx, q in enumerate(questions_list):
        if idx < len(tokens):
            q["correct"] = tokens[idx]
        else:
            q["correct"] = "A"

    async with async_session() as session:
        new_test = Test(
            title=data["title"],
            subject=data["subject"],
            grade_level=data["grade"],
            max_attempts=data["max_attempts"],
            duration_seconds_per_question=data["duration"],
            is_active=True
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
    await message.answer("✅ <b>Test muvaffaqiyatli saqlandi!</b>", reply_markup=get_admin_menu())

@router.message(F.text == "⚙️ Testlarni boshqarish")
async def manage_tests_admin(message: Message):
    if message.from_user.id not in SUPER_ADMIN_IDS:
        return
    async with async_session() as session:
        tests = (await session.execute(select(Test))).scalars().all()
        if not tests:
            await message.answer("⚠️ Testlar mavjud emas.")
            return
        
        keyboard_buttons = []
        for t in tests:
            status_icon = "🟢" if t.is_active else "🔴"
            att_text = f"Limit: {t.max_attempts}" if t.max_attempts > 0 else "Limit: Cheksiz"
            keyboard_buttons.append([
                InlineKeyboardButton(text=f"[{t.grade_level}] {t.subject} {status_icon} ({att_text})", callback_data="none"),
                InlineKeyboardButton(text="Holat 🔄", callback_data=f"toggle_test_{t.id}"),
                InlineKeyboardButton(text="O'chirish 🗑", callback_data=f"delete_test_{t.id}")
            ])
            
        markup = InlineKeyboardMarkup(inline_keyboard=keyboard_buttons)
        await message.answer("⚙️ <b>Testlarni boshqarish paneli:</b>", reply_markup=markup)

@router.callback_query(F.data.startswith("toggle_test_"))
async def toggle_test_status(callback: CallbackQuery):
    test_id = int(callback.data.split("_")[2])
    async with async_session() as session:
        test = await session.get(Test, test_id)
        if test:
            test.is_active = not test.is_active
            await session.commit()
            await callback.answer("Holat o'zgartirildi!")
            
            tests = (await session.execute(select(Test))).scalars().all()
            keyboard_buttons = []
            for t in tests:
                status_icon = "🟢" if t.is_active else "🔴"
                att_text = f"Limit: {t.max_attempts}" if t.max_attempts > 0 else "Limit: Cheksiz"
                keyboard_buttons.append([
                    InlineKeyboardButton(text=f"[{t.grade_level}] {t.subject} {status_icon} ({att_text})", callback_data="none"),
                    InlineKeyboardButton(text="Holat 🔄", callback_data=f"toggle_test_{t.id}"),
                    InlineKeyboardButton(text="O'chirish 🗑", callback_data=f"delete_test_{t.id}")
                ])
            markup = InlineKeyboardMarkup(inline_keyboard=keyboard_buttons)
            try:
                await callback.message.edit_reply_markup(reply_markup=markup)
            except Exception:
                pass

@router.callback_query(F.data.startswith("delete_test_"))
async def delete_test_handler(callback: CallbackQuery):
    test_id = int(callback.data.split("_")[2])
    async with async_session() as session:
        test = await session.get(Test, test_id)
        if test:
            await session.delete(test)
            await session.commit()
            await callback.answer("Test butunlay o'chirib yuborildi!", show_alert=True)
            
            tests = (await session.execute(select(Test))).scalars().all()
            if not tests:
                await callback.message.edit_text("⚙️ Hozircha bazada testlar qolmadi.")
                return
                
            keyboard_buttons = []
            for t in tests:
                status_icon = "🟢" if t.is_active else "🔴"
                att_text = f"Limit: {t.max_attempts}" if t.max_attempts > 0 else "Limit: Cheksiz"
                keyboard_buttons.append([
                    InlineKeyboardButton(text=f"[{t.grade_level}] {t.subject} {status_icon} ({att_text})", callback_data="none"),
                    InlineKeyboardButton(text="Holat 🔄", callback_data=f"toggle_test_{t.id}"),
                    InlineKeyboardButton(text="O'chirish 🗑", callback_data=f"delete_test_{t.id}")
                ])
            markup = InlineKeyboardMarkup(inline_keyboard=keyboard_buttons)
            try:
                await callback.message.edit_reply_markup(reply_markup=markup)
            except Exception:
                pass

@router.callback_query(F.data == "none")
async def none_callback(callback: CallbackQuery):
    await callback.answer()

@router.message(F.text == "📊 Jonli statistika")
async def live_statistics(message: Message):
    if message.from_user.id not in SUPER_ADMIN_IDS:
        return
    async with async_session() as session:
        total = await session.scalar(select(func.count(Student.id)))
        completed = await session.scalar(select(func.count(TestSession.id)).where(TestSession.status == "COMPLETED"))
        await message.answer(f"📊 <b>Statistika:</b>\n\nJami o'quvchilar: {total}\nTestni yakunlaganlar: {completed or 0}")

@router.message(F.text == "📥 Excel natijalar")
async def export_excel_results(message: Message, bot: Bot):
    if message.from_user.id not in SUPER_ADMIN_IDS:
        return
    async with async_session() as session:
        rows = (await session.execute(
            select(Student, TestSession, Test)
            .join(TestSession, Student.id == TestSession.student_id)
            .join(Test, TestSession.test_id == Test.id)
            .order_by(TestSession.score.desc())
        )).all()
        
        if not rows:
            await message.answer("⚠️ Natijalar mavjud emas.")
            return
            
        data_list = []
        for s, ts, t in rows:
            data_list.append({
                "ID": s.student_id,
                "Ism": s.first_name,
                "Familiya": s.last_name,
                "Maktab": s.school,
                "Sinf": s.grade,
                "Test Sinf": t.grade_level,
                "Test Fan": t.subject,
                "Ball": ts.score,
                "Foiz (%)": ts.score_percentage,
                "To'g'ri": ts.correct_answers,
                "Noto'g'ri": ts.wrong_answers,
                "Sana": ts.finished_at
            })
            
        df = pd.DataFrame(data_list)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Natijalar')
        output.seek(0)
        
        file_bytes = BufferedInputFile(output.read(), filename="olimpiada_natijalari.xlsx")
        await message.answer_document(file_bytes, caption="📥 Barcha o'quvchilar natijalari Excel ko'rinishida.")

@router.message(F.text == "🧹 Bazani tozalash")
async def reset_database_prompt(message: Message):
    if message.from_user.id not in SUPER_ADMIN_IDS:
        return
    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="Ha, barcha sessiyalarni o'chirish ⚠️", callback_data="confirm_reset_db")],
        [InlineKeyboardButton(text="Bekor qilish ❌", callback_data="cancel_reset")]
    ])
    await message.answer("⚠️ Diqqat! Barcha o'quvchilarning test sessiyalari va natijalari o'chib ketadi (o'quvchilar va testlar saqlanib qoladi). Davom etasizmi?", reply_markup=keyboard)

@router.callback_query(F.data == "confirm_reset_db")
async def confirm_reset(callback: CallbackQuery):
    async with async_session() as session:
        await session.execute(delete(TestSession))
        await session.commit()
    await callback.message.edit_text("✅ Barcha test natijalari va sessiyalar tozalandi!")

@router.callback_query(F.data == "cancel_reset")
async def cancel_reset(callback: CallbackQuery):
    await callback.message.edit_text("❌ Amaliyot bekor qilindi.")

@router.message(F.text == "📢 Xabar yuborish")
async def broadcast_prompt(message: Message, state: FSMContext):
    if message.from_user.id not in SUPER_ADMIN_IDS:
        return
    await state.set_state(AdminBroadcast.waiting_for_message)
    await message.answer("📢 Barcha o'quvchilarga yubormoqchi bo'lgan xabaringizni kiriting (Matn, rasm yoki boshqa formatda bo'lishi mumkin):")

@router.message(AdminBroadcast.waiting_for_message)
async def send_broadcast_message(message: Message, state: FSMContext, bot: Bot):
    if message.from_user.id not in SUPER_ADMIN_IDS:
        return
    
    async with async_session() as session:
        students = (await session.execute(select(Student).where(Student.telegram_id.is_not(None), Student.is_active == True))).scalars().all()
        
    success_count = 0
    fail_count = 0
    
    status_msg = await message.answer("⏳ Xabar yuborilmoqda...")
    
    for s in students:
        try:
            await message.send_copy(chat_id=s.telegram_id)
            success_count += 1
            await asyncio.sleep(0.05)
        except Exception:
            fail_count += 1
            
    await state.clear()
    await status_msg.edit_text(
        f"✅ <b>Xabar yuborish yakunlandi!</b>\n\n"
        f"• Muvaffaqiyatli: {success_count} ta\n"
        f"• Xatolik (bloklaganlar): {fail_count} ta",
        reply_markup=get_admin_menu()
    )

@router.message(F.text == "⬅️ Bosh menyu")
async def back_to_menu_handler(message: Message, state: FSMContext):
    current_state = await state.get_state()
    if current_state == TestProcessState.in_test.state:
        await message.answer("⚠️ Test jarayonida boshqa menyuga o'tib bo'lmaydi!")
        return
    await state.clear()
    async with async_session() as session:
        student = (await session.execute(select(Student).where(Student.telegram_id == message.from_user.id))).scalar_one_or_none()
        if student:
            await message.answer("Asosiy menyu:", reply_markup=get_main_menu())
        elif message.from_user.id in SUPER_ADMIN_IDS:
            await message.answer("Admin menyu:", reply_markup=get_admin_menu())

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
