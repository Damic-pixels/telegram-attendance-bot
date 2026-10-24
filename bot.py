"""
Department Attendance Bot (Telegram)
- Student registration (once)
- Admin opens/closes sessions per course, part 1 (before class) or part 2 (after class)
- Location + class-code check
- Auto-notifies registered students when a session opens
- Admin lists, absentees, Excel export
"""
import asyncio
import functools
import io
import logging
import math
import os
import random
import sqlite3
from datetime import datetime
from zoneinfo import ZoneInfo

from openpyxl import Workbook
from openpyxl.styles import Font
from telegram import (
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    KeyboardButton,
    ReplyKeyboardMarkup,
    ReplyKeyboardRemove,
    Update,
)
from telegram.error import TelegramError
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    ConversationHandler,
    MessageHandler,
    filters,
)

# ------------------------------------------------------------------ CONFIG
BOT_TOKEN = os.environ["BOT_TOKEN"]
SUPER_ADMIN_ID = int(os.environ["SUPER_ADMIN_ID"])  # your Telegram numeric ID
RADIUS_M = int(os.getenv("RADIUS_M", "50"))          # allowed distance from admin
DB_PATH = os.getenv("DB_PATH", "attendance.db")
TZ = ZoneInfo("Africa/Lagos")

FACULTY = "Computing and Informatics"
DEPARTMENT = "Computer Science"
DEFAULT_LEVEL = "200"
PART_LABEL = {1: "Before class", 2: "After class"}

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("attendance")


# ------------------------------------------------------------------ DATABASE
def run(sql, args=(), fetch=None):
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    try:
        cur = con.execute(sql, args)
        con.commit()
        if fetch == "one":
            return cur.fetchone()
        if fetch == "all":
            return cur.fetchall()
        return cur.lastrowid
    finally:
        con.close()


def init_db():
    con = sqlite3.connect(DB_PATH)
    con.executescript(
        """
        CREATE TABLE IF NOT EXISTS students (
            tg_id INTEGER PRIMARY KEY,
            full_name TEXT NOT NULL,
            matric TEXT NOT NULL UNIQUE,
            school_email TEXT,
            personal_email TEXT,
            phone TEXT,
            extra TEXT,
            faculty TEXT,
            department TEXT,
            level TEXT
        );
        CREATE TABLE IF NOT EXISTS admins (tg_id INTEGER PRIMARY KEY);
        CREATE TABLE IF NOT EXISTS courses (
            code TEXT PRIMARY KEY,
            title TEXT
        );
        CREATE TABLE IF NOT EXISTS sessions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            course_code TEXT NOT NULL,
            part INTEGER NOT NULL,
            date TEXT NOT NULL,
            opened_by INTEGER,
            lat REAL,
            lon REAL,
            code TEXT,
            is_open INTEGER DEFAULT 1,
            opened_at TEXT,
            closed_at TEXT,
            UNIQUE(course_code, date, part)
        );
        CREATE TABLE IF NOT EXISTS attendance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            session_id INTEGER NOT NULL,
            tg_id INTEGER NOT NULL,
            marked_at TEXT,
            distance REAL,
            UNIQUE(session_id, tg_id)
        );
        """
    )
    con.close()


# ------------------------------------------------------------------ HELPERS
def now():
    return datetime.now(TZ)


def today():
    return now().strftime("%Y-%m-%d")


def stamp():
    return now().strftime("%Y-%m-%d %H:%M:%S")


def distance_m(lat1, lon1, lat2, lon2):
    r = 6371000
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dphi = p2 - p1
    dl = math.radians(lon2 - lon1)
    a = math.sin(dphi / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * r * math.asin(math.sqrt(a))


def is_admin(uid):
    if uid == SUPER_ADMIN_ID:
        return True
    return run("SELECT 1 FROM admins WHERE tg_id=?", (uid,), "one") is not None


def get_student(uid):
    return run("SELECT * FROM students WHERE tg_id=?", (uid,), "one")


def admin_only(fn):
    @functools.wraps(fn)
    async def wrapper(update: Update, ctx: ContextTypes.DEFAULT_TYPE, *a, **k):
        if not is_admin(update.effective_user.id):
            if update.callback_query:
                await update.callback_query.answer("Admins only", show_alert=True)
            else:
                await update.message.reply_text("Admins only.")
            return ConversationHandler.END
        return await fn(update, ctx, *a, **k)

    return wrapper


async def send_long(update: Update, text: str):
    for i in range(0, len(text), 3800):
        await update.message.reply_text(text[i : i + 3800])


PRIVATE = filters.ChatType.PRIVATE

# ------------------------------------------------------------------ REGISTRATION
NAME, MATRIC, SMAIL, PMAIL, PHONE, EXTRA, CONFIRM = range(7)


async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    if get_student(uid):
        await update.message.reply_text(
            "Welcome back!\n/mark - mark attendance (when open)\n/mystats - your attendance\n/help - all commands"
        )
        return ConversationHandler.END
    await update.message.reply_text(
        f"Welcome to the {DEPARTMENT} Attendance Bot.\nLet's register you (one time only).\n\nEnter your FULL NAME:"
    )
    return NAME


async def reg_name(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    name = update.message.text.strip()
    if len(name) < 5:
        await update.message.reply_text("Please enter your full name.")
        return NAME
    ctx.user_data["reg"] = {"full_name": name}
    await update.message.reply_text("Enter your MATRIC NUMBER:")
    return MATRIC


async def reg_matric(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    matric = update.message.text.strip().upper().replace(" ", "")
    if len(matric) < 4:
        await update.message.reply_text("That doesn't look right. Enter your matric number:")
        return MATRIC
    row = run("SELECT tg_id FROM students WHERE matric=?", (matric,), "one")
    if row and row["tg_id"] != update.effective_user.id:
        await update.message.reply_text(
            "This matric number is already registered on another account. Contact an admin."
        )
        return MATRIC
    ctx.user_data["reg"]["matric"] = matric
    await update.message.reply_text("Enter your SCHOOL EMAIL:")
    return SMAIL


async def reg_smail(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    mail = update.message.text.strip().lower()
    if "@" not in mail or "." not in mail:
        await update.message.reply_text("Invalid email. Enter your school email:")
        return SMAIL
    ctx.user_data["reg"]["school_email"] = mail
    await update.message.reply_text("Enter your PERSONAL EMAIL:")
    return PMAIL


async def reg_pmail(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    mail = update.message.text.strip().lower()
    if "@" not in mail or "." not in mail:
        await update.message.reply_text("Invalid email. Enter your personal email:")
        return PMAIL
    ctx.user_data["reg"]["personal_email"] = mail
    kb = ReplyKeyboardMarkup(
        [[KeyboardButton("Share my phone number", request_contact=True)]],
        resize_keyboard=True,
        one_time_keyboard=True,
    )
    await update.message.reply_text("Tap the button to share your PHONE NUMBER (or type it):", reply_markup=kb)
    return PHONE


async def reg_phone(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if update.message.contact:
        phone = update.message.contact.phone_number
    else:
        phone = update.message.text.strip()
    digits = "".join(c for c in phone if c.isdigit())
    if len(digits) < 10:
        await update.message.reply_text("Invalid phone number. Try again:")
        return PHONE
    ctx.user_data["reg"]["phone"] = phone
    await update.message.reply_text(
        "Any other important info (gender, state of origin, etc.)? Type it, or send 'skip'.",
        reply_markup=ReplyKeyboardRemove(),
    )
    return EXTRA


async def reg_extra(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    ctx.user_data["reg"]["extra"] = "" if text.lower() == "skip" else text
    r = ctx.user_data["reg"]
    summary = (
        "Please confirm your details:\n\n"
        f"Name: {r['full_name']}\n"
        f"Matric: {r['matric']}\n"
        f"School email: {r['school_email']}\n"
        f"Personal email: {r['personal_email']}\n"
        f"Phone: {r['phone']}\n"
        f"Other info: {r['extra'] or '-'}\n"
        f"Faculty: {FACULTY}\n"
        f"Department: {DEPARTMENT}\n"
        f"Level: {DEFAULT_LEVEL}"
    )
    kb = ReplyKeyboardMarkup([["Confirm", "Restart"]], resize_keyboard=True, one_time_keyboard=True)
    await update.message.reply_text(summary, reply_markup=kb)
    return CONFIRM


async def reg_confirm(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    t = update.message.text.strip().lower()
    if t == "restart":
        await update.message.reply_text("Restarting. Enter your FULL NAME:", reply_markup=ReplyKeyboardRemove())
        return NAME
    if t != "confirm":
        await update.message.reply_text("Tap Confirm or Restart.")
        return CONFIRM
    r = ctx.user_data["reg"]
    try:
        run(
            """INSERT INTO students
               (tg_id, full_name, matric, school_email, personal_email, phone, extra, faculty, department, level)
               VALUES (?,?,?,?,?,?,?,?,?,?)""",
            (
                update.effective_user.id,
                r["full_name"],
                r["matric"],
                r["school_email"],
                r["personal_email"],
                r["phone"],
                r["extra"],
                FACULTY,
                DEPARTMENT,
                DEFAULT_LEVEL,
            ),
        )
    except sqlite3.IntegrityError:
        await update.message.reply_text(
            "Registration failed: matric already registered. Contact an admin.",
            reply_markup=ReplyKeyboardRemove(),
        )
        return ConversationHandler.END
    await update.message.reply_text(
        "Registration complete! You'll get a notification when attendance opens.\nUse /mark to mark attendance.",
        reply_markup=ReplyKeyboardRemove(),
    )
    return ConversationHandler.END


async def cancel(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data.clear()
    await update.message.reply_text("Cancelled.", reply_markup=ReplyKeyboardRemove())
    return ConversationHandler.END


# ------------------------------------------------------------------ ADMIN: OPEN SESSION
O_COURSE, O_PART, O_LOC = range(10, 13)


@admin_only
async def open_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    courses = run("SELECT * FROM courses ORDER BY code", fetch="all")
    if not courses:
        await update.message.reply_text("No courses yet. Add one: /addcourse CSC201 Data Structures")
        return ConversationHandler.END
    rows = [[InlineKeyboardButton(c["code"], callback_data=f"oc:{c['code']}")] for c in courses]
    await update.message.reply_text("Select course:", reply_markup=InlineKeyboardMarkup(rows))
    return O_COURSE


async def open_pick_course(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    ctx.user_data["o_code"] = q.data.split(":", 1)[1]
    kb = InlineKeyboardMarkup(
        [
            [InlineKeyboardButton("1 - Before class", callback_data="op:1")],
            [InlineKeyboardButton("2 - After class", callback_data="op:2")],
        ]
    )
    await q.edit_message_text(f"Course: {ctx.user_data['o_code']}\nSelect session:", reply_markup=kb)
    return O_PART


async def open_pick_part(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    ctx.user_data["o_part"] = int(q.data.split(":", 1)[1])
    await q.edit_message_text(
        f"{ctx.user_data['o_code']} - {PART_LABEL[ctx.user_data['o_part']]}\nNow share your location."
    )
    kb = ReplyKeyboardMarkup(
        [[KeyboardButton("Share class location", request_location=True)]],
        resize_keyboard=True,
        one_time_keyboard=True,
    )
    await q.message.reply_text(
        "Stand in the classroom and tap the button. Students must be near this spot to mark attendance.",
        reply_markup=kb,
    )
    return O_LOC


async def open_got_location(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    loc = update.message.location
    code = ctx.user_data["o_code"]
    part = ctx.user_data["o_part"]
    class_code = f"{random.randint(0, 9999):04d}"
    existing = run(
        "SELECT id FROM sessions WHERE course_code=? AND date=? AND part=?", (code, today(), part), "one"
    )
    if existing:
        sid = existing["id"]
        run(
            """UPDATE sessions SET lat=?, lon=?, code=?, is_open=1, opened_by=?, opened_at=?, closed_at=NULL
               WHERE id=?""",
            (loc.latitude, loc.longitude, class_code, update.effective_user.id, stamp(), sid),
        )
    else:
        sid = run(
            """INSERT INTO sessions (course_code, part, date, opened_by, lat, lon, code, is_open, opened_at)
               VALUES (?,?,?,?,?,?,?,1,?)""",
            (code, part, today(), update.effective_user.id, loc.latitude, loc.longitude, class_code, stamp()),
        )
    course = run("SELECT * FROM courses WHERE code=?", (code,), "one")
    title = course["title"] if course and course["title"] else ""
    await update.message.reply_text(
        f"Attendance OPEN\n{code} {title}\n{PART_LABEL[part]} (part {part} of 2)\n\n"
        f"Students must be within {RADIUS_M} m of you to mark.\n"
        f"Close with /close when done.",
        reply_markup=ReplyKeyboardRemove(),
    )
    text = (
        f"ATTENDANCE OPEN\n{code} {title}\nSession: {PART_LABEL[part]} (part {part} of 2)\n\n"
        "Go to class, then send /mark here and share your location."
    )
    ctx.application.create_task(notify_students(ctx, text, update.effective_user.id))
    ctx.user_data.clear()
    return ConversationHandler.END


async def notify_students(ctx, text, admin_id):
    rows = run("SELECT tg_id FROM students", fetch="all")
    sent = 0
    for r in rows:
        try:
            await ctx.bot.send_message(r["tg_id"], text)
            sent += 1
        except TelegramError:
            pass
        await asyncio.sleep(0.05)  # stay under Telegram rate limits
    try:
        await ctx.bot.send_message(admin_id, f"Notification delivered to {sent}/{len(rows)} students.")
    except TelegramError:
        pass


@admin_only
async def code_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    rows = run("SELECT * FROM sessions WHERE is_open=1 ORDER BY id", fetch="all")
    if not rows:
        await update.message.reply_text("No open sessions. Use /open first.")
        return
    await update.message.reply_text(
        "\n".join(f"{r['course_code']} - {PART_LABEL[r['part']]}: CODE {r['code']}" for r in rows)
    )


# ------------------------------------------------------------------ ADMIN: CLOSE
@admin_only
async def close_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    open_s = run("SELECT * FROM sessions WHERE is_open=1 ORDER BY id", fetch="all")
    if not open_s:
        await update.message.reply_text("No open sessions.")
        return
    rows = [
        [InlineKeyboardButton(f"{s['course_code']} - {PART_LABEL[s['part']]}", callback_data=f"cl:{s['id']}")]
        for s in open_s
    ]
    await update.message.reply_text("Which session do you want to close?", reply_markup=InlineKeyboardMarkup(rows))


@admin_only
async def close_pick(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    sid = int(q.data.split(":", 1)[1])
    s = run("SELECT * FROM sessions WHERE id=?", (sid,), "one")
    run("UPDATE sessions SET is_open=0, closed_at=? WHERE id=?", (stamp(), sid))
    n = run("SELECT COUNT(*) c FROM attendance WHERE session_id=?", (sid,), "one")["c"]
    await q.edit_message_text(
        f"Closed: {s['course_code']} - {PART_LABEL[s['part']]}\nStudents marked: {n}\nSee names: /list {s['course_code']}"
    )


# ------------------------------------------------------------------ STUDENT: MARK
M_PICK, M_CODE, M_LOC = range(20, 23)


async def mark_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    if not get_student(uid):
        await update.message.reply_text("Register first with /start.")
        return ConversationHandler.END
    open_s = run("SELECT * FROM sessions WHERE is_open=1 ORDER BY id", fetch="all")
    if not open_s:
        await update.message.reply_text("No attendance is open right now.")
        return ConversationHandler.END
    if len(open_s) == 1:
        return await begin_mark(update, ctx, open_s[0])
    rows = [
        [InlineKeyboardButton(f"{s['course_code']} - {PART_LABEL[s['part']]}", callback_data=f"ms:{s['id']}")]
        for s in open_s
    ]
    await update.message.reply_text("Which class?", reply_markup=InlineKeyboardMarkup(rows))
    return M_PICK


async def mark_pick(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    s = run("SELECT * FROM sessions WHERE id=?", (int(q.data.split(":", 1)[1]),), "one")
    return await begin_mark(update, ctx, s)


async def begin_mark(update: Update, ctx: ContextTypes.DEFAULT_TYPE, s):
    uid = update.effective_user.id
    msg = update.callback_query.message if update.callback_query else update.message
    if not s or not s["is_open"]:
        await msg.reply_text("That session is closed.")
        return ConversationHandler.END
    if run("SELECT 1 FROM attendance WHERE session_id=? AND tg_id=?", (s["id"], uid), "one"):
        await msg.reply_text("You already marked attendance for this session.")
        return ConversationHandler.END
    ctx.user_data["m_sid"] = s["id"]
    kb = ReplyKeyboardMarkup(
        [[KeyboardButton("Share my location", request_location=True)]],
        resize_keyboard=True,
        one_time_keyboard=True,
    )
    await msg.reply_text(
        f"{s['course_code']} - {PART_LABEL[s['part']]}\nTap the button to share your location:",
        reply_markup=kb,
    )
    return M_LOC


async def mark_code(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    s = run("SELECT * FROM sessions WHERE id=?", (ctx.user_data.get("m_sid"),), "one")
    if not s or not s["is_open"]:
        await update.message.reply_text("This session has been closed.")
        return ConversationHandler.END
    if update.message.text.strip() != s["code"]:
        await update.message.reply_text("Wrong code. Try again (or /cancel):")
        return M_CODE
    kb = ReplyKeyboardMarkup(
        [[KeyboardButton("Share my location", request_location=True)]],
        resize_keyboard=True,
        one_time_keyboard=True,
    )
    await update.message.reply_text("Code accepted. Now tap the button to share your location:", reply_markup=kb)
    return M_LOC


async def mark_location(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    msg = update.message
    if msg.forward_origin:
        await msg.reply_text("Forwarded locations are not allowed. Use the button.")
        return M_LOC
    s = run("SELECT * FROM sessions WHERE id=?", (ctx.user_data.get("m_sid"),), "one")
    if not s or not s["is_open"]:
        await msg.reply_text("This session has been closed.", reply_markup=ReplyKeyboardRemove())
        return ConversationHandler.END
    d = distance_m(msg.location.latitude, msg.location.longitude, s["lat"], s["lon"])
    if d > RADIUS_M:
        await msg.reply_text(
            f"You are about {int(d)} m from the class (limit {RADIUS_M} m). Move closer and share again, or /cancel."
        )
        return M_LOC
    try:
        run(
            "INSERT INTO attendance (session_id, tg_id, marked_at, distance) VALUES (?,?,?,?)",
            (s["id"], update.effective_user.id, stamp(), round(d, 1)),
        )
    except sqlite3.IntegrityError:
        await msg.reply_text("You already marked this session.", reply_markup=ReplyKeyboardRemove())
        return ConversationHandler.END
    await msg.reply_text(
        f"Attendance marked for {s['course_code']} ({PART_LABEL[s['part']]}). "
        f"Remember: part 1 and part 2 are 50% each.",
        reply_markup=ReplyKeyboardRemove(),
    )
    ctx.user_data.clear()
    return ConversationHandler.END


async def mystats(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    if not get_student(uid):
        await update.message.reply_text("Register first with /start.")
        return
    lectures = run(
        "SELECT course_code, COUNT(DISTINCT date) n FROM sessions GROUP BY course_code ORDER BY course_code",
        fetch="all",
    )
    if not lectures:
        await update.message.reply_text("No attendance records yet.")
        return
    marks = {
        r["course_code"]: r["c"]
        for r in run(
            """SELECT s.course_code, COUNT(*) c FROM attendance a JOIN sessions s ON s.id=a.session_id
               WHERE a.tg_id=? GROUP BY s.course_code""",
            (uid,),
            "all",
        )
    }
    lines = ["Your attendance:"]
    for l in lectures:
        pct = marks.get(l["course_code"], 0) * 50 / l["n"]
        lines.append(f"{l['course_code']}: {pct:.0f}% ({l['n']} lectures)")
    await update.message.reply_text("\n".join(lines))


# ------------------------------------------------------------------ ADMIN: MANAGEMENT
@admin_only
async def addcourse(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    body = update.message.text.split(None, 1)
    if len(body) < 2:
        await update.message.reply_text(
            "Usage (one or many, one course per line):\n/addcourse CSC201 Data Structures\nCSC202 Algorithms"
        )
        return
    saved = []
    for line in body[1].splitlines():
        parts = line.strip().split(None, 1)
        if not parts:
            continue
        code = parts[0].upper()
        title = parts[1] if len(parts) > 1 else ""
        run("INSERT OR REPLACE INTO courses (code, title) VALUES (?,?)", (code, title))
        saved.append(f"{code} {title}".strip())
    await update.message.reply_text(f"Saved {len(saved)} course(s):\n" + "\n".join(saved))


@admin_only
async def removecourse(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not ctx.args:
        await update.message.reply_text("Usage: /removecourse CSC201")
        return
    run("DELETE FROM courses WHERE code=?", (ctx.args[0].upper(),))
    await update.message.reply_text("Course removed (past attendance records are kept).")


@admin_only
async def courses_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    rows = run("SELECT * FROM courses ORDER BY code", fetch="all")
    if not rows:
        await update.message.reply_text("No courses yet.")
        return
    await update.message.reply_text("\n".join(f"{r['code']} - {r['title']}" for r in rows))


async def addadmin(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != SUPER_ADMIN_ID:
        await update.message.reply_text("Only the super admin can do this.")
        return
    if not ctx.args or not ctx.args[0].isdigit():
        await update.message.reply_text("Usage: /addadmin TELEGRAM_ID (they get it from /myid)")
        return
    run("INSERT OR IGNORE INTO admins (tg_id) VALUES (?)", (int(ctx.args[0]),))
    await update.message.reply_text("Admin added.")


async def removeadmin(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != SUPER_ADMIN_ID:
        await update.message.reply_text("Only the super admin can do this.")
        return
    if not ctx.args or not ctx.args[0].isdigit():
        await update.message.reply_text("Usage: /removeadmin TELEGRAM_ID")
        return
    run("DELETE FROM admins WHERE tg_id=?", (int(ctx.args[0]),))
    await update.message.reply_text("Admin removed.")


async def myid(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(f"Your Telegram ID: {update.effective_user.id}")


@admin_only
async def setlevel(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not ctx.args:
        await update.message.reply_text("Usage:\n/setlevel 300 (everyone)\n/setlevel 300 MATRIC (one student)")
        return
    level = ctx.args[0]
    if len(ctx.args) > 1:
        run("UPDATE students SET level=? WHERE matric=?", (level, ctx.args[1].upper()))
        await update.message.reply_text(f"Updated {ctx.args[1].upper()} to level {level}.")
    else:
        run("UPDATE students SET level=?", (level,))
        await update.message.reply_text(f"All students set to level {level}.")


# ------------------------------------------------------------------ ADMIN: VIEWS
def sessions_for_view(code):
    if code:
        last = run("SELECT MAX(date) d FROM sessions WHERE course_code=?", (code,), "one")
        if not last or not last["d"]:
            return []
        return run(
            "SELECT * FROM sessions WHERE course_code=? AND date=? ORDER BY part", (code, last["d"]), "all"
        )
    return run("SELECT * FROM sessions WHERE is_open=1 ORDER BY id", fetch="all")


@admin_only
async def list_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    code = ctx.args[0].upper() if ctx.args else None
    sessions = sessions_for_view(code)
    if not sessions:
        await update.message.reply_text("No sessions found. Use /list CSC201 for a course.")
        return
    out = []
    for s in sessions:
        rows = run(
            """SELECT st.matric, st.full_name FROM attendance a JOIN students st ON st.tg_id=a.tg_id
               WHERE a.session_id=? ORDER BY st.full_name""",
            (s["id"],),
            "all",
        )
        status = "OPEN" if s["is_open"] else "closed"
        out.append(f"{s['course_code']} | {s['date']} | {PART_LABEL[s['part']]} | {status} | {len(rows)} marked")
        out += [f"{i}. {r['matric']} - {r['full_name']}" for i, r in enumerate(rows, 1)]
        out.append("")
    await send_long(update, "\n".join(out))


@admin_only
async def absent_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    code = ctx.args[0].upper() if ctx.args else None
    sessions = sessions_for_view(code)
    if not sessions:
        await update.message.reply_text("No sessions found. Use /absent CSC201 for a course.")
        return
    out = []
    for s in sessions:
        rows = run(
            """SELECT matric, full_name FROM students
               WHERE tg_id NOT IN (SELECT tg_id FROM attendance WHERE session_id=?)
               ORDER BY full_name""",
            (s["id"],),
            "all",
        )
        out.append(f"ABSENT - {s['course_code']} | {s['date']} | {PART_LABEL[s['part']]} | {len(rows)}")
        out += [f"{i}. {r['matric']} - {r['full_name']}" for i, r in enumerate(rows, 1)]
        out.append("")
    await send_long(update, "\n".join(out))


@admin_only
async def students_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    rows = run("SELECT * FROM students ORDER BY full_name", fetch="all")
    if not rows:
        await update.message.reply_text("No students registered yet.")
        return
    lines = [f"Registered students: {len(rows)}"]
    lines += [f"{i}. {r['matric']} - {r['full_name']} ({r['level']})" for i, r in enumerate(rows, 1)]
    await send_long(update, "\n".join(lines))


# ------------------------------------------------------------------ ADMIN: EXPORT
def bold_header(ws):
    for c in ws[1]:
        c.font = Font(bold=True)


def add_course_sheet(wb, code, students):
    ws = wb.create_sheet(code[:31])
    dates = [
        r["date"]
        for r in run("SELECT DISTINCT date FROM sessions WHERE course_code=? ORDER BY date", (code,), "all")
    ]
    marks = {}
    for r in run(
        """SELECT a.tg_id, s.date, COUNT(*) c FROM attendance a JOIN sessions s ON s.id=a.session_id
           WHERE s.course_code=? GROUP BY a.tg_id, s.date""",
        (code,),
        "all",
    ):
        marks[(r["tg_id"], r["date"])] = r["c"] * 50
    ws.append(["Matric", "Name"] + dates + ["Total %"])
    for st in students:
        vals = [marks.get((st["tg_id"], d), 0) for d in dates]
        total = round(sum(vals) / len(dates), 1) if dates else 0
        ws.append([st["matric"], st["full_name"]] + vals + [total])
    bold_header(ws)
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 32


@admin_only
async def export_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    code = ctx.args[0].upper() if ctx.args else None
    students = run("SELECT * FROM students ORDER BY full_name", fetch="all")
    if code:
        codes = [code]
    else:
        codes = [r["course_code"] for r in run("SELECT DISTINCT course_code FROM sessions ORDER BY 1", fetch="all")]
    if not codes:
        await update.message.reply_text("No attendance data yet.")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws.append(["Matric", "Name", "Level"] + [f"{c} %" for c in codes])
    for st in students:
        row = [st["matric"], st["full_name"], st["level"]]
        for c in codes:
            n = run("SELECT COUNT(DISTINCT date) n FROM sessions WHERE course_code=?", (c,), "one")["n"]
            m = run(
                """SELECT COUNT(*) c FROM attendance a JOIN sessions s ON s.id=a.session_id
                   WHERE a.tg_id=? AND s.course_code=?""",
                (st["tg_id"], c),
                "one",
            )["c"]
            row.append(round(m * 50 / n, 1) if n else 0)
        ws.append(row)
    bold_header(ws)
    ws.column_dimensions["B"].width = 32

    for c in codes:
        add_course_sheet(wb, c, students)

    raw = wb.create_sheet("All Records")
    raw.append(["Date", "Course", "Part", "Matric", "Name", "Time marked", "Distance (m)"])
    sql = """SELECT s.date, s.course_code, s.part, st.matric, st.full_name, a.marked_at, a.distance
             FROM attendance a JOIN sessions s ON s.id=a.session_id JOIN students st ON st.tg_id=a.tg_id"""
    args = ()
    if code:
        sql += " WHERE s.course_code=?"
        args = (code,)
    sql += " ORDER BY s.date, s.course_code, s.part, st.full_name"
    for r in run(sql, args, "all"):
        raw.append(list(tuple(r)))
    bold_header(raw)

    sheet = wb.create_sheet("Students")
    sheet.append(["Matric", "Name", "School email", "Personal email", "Phone", "Other info", "Department", "Level"])
    for st in students:
        sheet.append(
            [st["matric"], st["full_name"], st["school_email"], st["personal_email"], st["phone"],
             st["extra"], st["department"], st["level"]]
        )
    bold_header(sheet)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    name = f"attendance_{code or 'all'}_{today()}.xlsx"
    await update.message.reply_document(document=buf, filename=name)


# ------------------------------------------------------------------ HELP
async def help_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    text = "Student commands:\n/start - register\n/mark - mark attendance\n/mystats - your attendance %\n/myid - your Telegram ID"
    if is_admin(update.effective_user.id):
        text += (
            "\n\nAdmin commands:\n"
            "/open - open a session\n/close - close a session\n"
            "/list [CODE] - who marked\n/absent [CODE] - who is missing\n"
            "/export [CODE] - Excel file (print-ready)\n/students - registered students\n"
            "/addcourse CODE Title\n/removecourse CODE\n/courses\n"
            "/setlevel LEVEL [MATRIC]\n/addadmin ID  /removeadmin ID (super admin)"
        )
    await update.message.reply_text(text)


# ------------------------------------------------------------------ MAIN
def main():
    init_db()
    app = Application.builder().token(BOT_TOKEN).build()

    reg_conv = ConversationHandler(
        entry_points=[CommandHandler("start", start, filters=PRIVATE)],
        states={
            NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, reg_name)],
            MATRIC: [MessageHandler(filters.TEXT & ~filters.COMMAND, reg_matric)],
            SMAIL: [MessageHandler(filters.TEXT & ~filters.COMMAND, reg_smail)],
            PMAIL: [MessageHandler(filters.TEXT & ~filters.COMMAND, reg_pmail)],
            PHONE: [MessageHandler((filters.TEXT & ~filters.COMMAND) | filters.CONTACT, reg_phone)],
            EXTRA: [MessageHandler(filters.TEXT & ~filters.COMMAND, reg_extra)],
            CONFIRM: [MessageHandler(filters.TEXT & ~filters.COMMAND, reg_confirm)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
        allow_reentry=True,
    )

    open_conv = ConversationHandler(
        entry_points=[CommandHandler("open", open_cmd, filters=PRIVATE)],
        states={
            O_COURSE: [CallbackQueryHandler(open_pick_course, pattern="^oc:")],
            O_PART: [CallbackQueryHandler(open_pick_part, pattern="^op:")],
            O_LOC: [MessageHandler(filters.LOCATION, open_got_location)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
        allow_reentry=True,
    )

    mark_conv = ConversationHandler(
        entry_points=[CommandHandler("mark", mark_cmd, filters=PRIVATE)],
        states={
            M_PICK: [CallbackQueryHandler(mark_pick, pattern="^ms:")],
            M_CODE: [MessageHandler(filters.TEXT & ~filters.COMMAND, mark_code)],
            M_LOC: [MessageHandler(filters.LOCATION, mark_location)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
        allow_reentry=True,
    )

    app.add_handler(reg_conv)
    app.add_handler(open_conv)
    app.add_handler(mark_conv)
    app.add_handler(CallbackQueryHandler(close_pick, pattern="^cl:"))

    simple = {
        "close": close_cmd,
        "code": code_cmd,
        "mystats": mystats,
        "addcourse": addcourse,
        "removecourse": removecourse,
        "courses": courses_cmd,
        "addadmin": addadmin,
        "removeadmin": removeadmin,
        "myid": myid,
        "setlevel": setlevel,
        "list": list_cmd,
        "absent": absent_cmd,
        "students": students_cmd,
        "export": export_cmd,
        "help": help_cmd,
    }
    for name, fn in simple.items():
        app.add_handler(CommandHandler(name, fn, filters=PRIVATE))

    log.info("Bot running...")
    asyncio.set_event_loop(asyncio.new_event_loop())  # needed on Python 3.14+
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()