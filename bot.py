import asyncio
import csv
import glob
import json
import logging
import os
import random
import re
import sqlite3
import time
import traceback
from datetime import datetime, timedelta, timezone, tzinfo
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from maxapi import Bot, Dispatcher, F
from maxapi.context import MemoryContext, StatesGroup, State
from maxapi.enums.message_link_type import MessageLinkType
from maxapi.types import (
    BotStarted,
    CallbackButton,
    Command,
    InputMedia,
    MessageCallback,
    MessageCreated,
    NewMessageLink,
)
from maxapi.utils.inline_keyboard import InlineKeyboardBuilder

try:
    from maxapi.types import LinkButton
except ImportError:  # старая версия maxapi — ссылки-кнопки не покажем, останутся обычные
    LinkButton = None

load_dotenv()
logging.basicConfig(level=logging.INFO)
log = logging.getLogger("sprashivatel")

TOKEN = os.getenv("MAX_BOT_TOKEN", "")  # токен — только в .env, не в коде!
# ADMIN_IDS — НЕ роли в боте. Это люди, которым бот присылает технические сообщения об ошибках (можно оставить пустым).
ADMIN_IDS = {int(x) for x in os.getenv("ADMIN_IDS", "").split(",") if x.strip()}
TZ_NAME = os.getenv("TIMEZONE", "Europe/Moscow")
DB_PATH = os.getenv("DB_PATH", "sprashivatel.db")
BOT_USERNAME = os.getenv("BOT_USERNAME", "ideya_fiks_109_bot")  # без @; если пусто — попробуем получить у API
RETENTION_DAYS = int(os.getenv("ANSWERS_RETENTION_DAYS", "0") or 0)  # автоудаление ответов старше N дней (0 — выкл)
EXPORT_DIR = "exports"
BACKUP_DIR = os.getenv("BACKUP_DIR", "backups")
BACKUP_KEEP = 14  # сколько суточных копий БД хранить
PAGE_SIZE = 8  # элементов на странице в списках
REMIND_MIN_LIMIT = 5  # не чаще, чем раз в 5 минут (защита от лимитов API и спама)
REMIND_MAX_LIMIT = 7 * 24 * 60  # не реже, чем раз в неделю
os.makedirs(EXPORT_DIR, exist_ok=True)


def load_tz(name: str) -> tzinfo:
    """Загружает часовой пояс. На Windows нужен пакет tzdata (pip install tzdata);
    если его нет — для Москвы используем фиксированный UTC+3."""
    try:
        return ZoneInfo(name)
    except (ZoneInfoNotFoundError, ValueError):
        log.warning("Часовой пояс %s не найден. Установите: pip install tzdata", name)
        if name == "Europe/Moscow":
            return timezone(timedelta(hours=3), "MSK")
        return timezone.utc


TZ = load_tz(TZ_NAME)

bot = Bot(TOKEN)
dp = Dispatcher()


# ---------- БД ----------
db = sqlite3.connect(DB_PATH, check_same_thread=False)
db.row_factory = sqlite3.Row
db.executescript(
    """
CREATE TABLE IF NOT EXISTS users(
    user_id INTEGER PRIMARY KEY, name TEXT, is_admin INTEGER DEFAULT 0);
CREATE TABLE IF NOT EXISTS polls(
    id INTEGER PRIMARY KEY AUTOINCREMENT, owner_id INTEGER, questions TEXT,
    status TEXT, send_at TEXT, created TEXT);
CREATE TABLE IF NOT EXISTS recipients(
    poll_id INTEGER, user_id INTEGER, sent INTEGER DEFAULT 0, done INTEGER DEFAULT 0,
    PRIMARY KEY(poll_id, user_id));
CREATE TABLE IF NOT EXISTS answers(
    poll_id INTEGER, user_id INTEGER, q_idx INTEGER, question TEXT, answer TEXT, ts TEXT);
CREATE TABLE IF NOT EXISTS chat_log(
    chat_id INTEGER, title TEXT, user_id INTEGER, name TEXT, text TEXT, ts TEXT);
"""
)

# Версии схемы: номер версии = индекс в списке + 1, хранится в PRAGMA user_version.
MIGRATIONS = [
    [  # v1
        "ALTER TABLE users ADD COLUMN chat_id INTEGER",  # адрес личного чата пользователя
        "ALTER TABLE polls ADD COLUMN chat_ids TEXT",  # JSON-список групповых чатов для публикации
        "ALTER TABLE polls ADD COLUMN chats_sent INTEGER DEFAULT 0",
        "ALTER TABLE polls ADD COLUMN access TEXT",  # устарело (оставлено для совместимости)
        "ALTER TABLE users ADD COLUMN grp TEXT",  # устарело
        "UPDATE polls SET access='chat' WHERE access IS NULL AND chat_ids IS NOT NULL",
    ],
    [  # v2
        "ALTER TABLE polls ADD COLUMN deadline TEXT",  # ISO UTC, после него приём закрывается
        "ALTER TABLE polls ADD COLUMN remind_h INTEGER DEFAULT 0",  # устарело: заменено на remind_min (v4)
        "ALTER TABLE polls ADD COLUMN reminded INTEGER DEFAULT 0",  # устарело
        "ALTER TABLE polls ADD COLUMN last_remind TEXT",
        "ALTER TABLE polls ADD COLUMN anon INTEGER DEFAULT 0",
        "CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY, value TEXT)",
        "CREATE TABLE IF NOT EXISTS chat_settings("
        "chat_id INTEGER PRIMARY KEY, title TEXT, logging INTEGER DEFAULT 1, "
        "notified INTEGER DEFAULT 0, last_ts TEXT)",
        "INSERT OR IGNORE INTO chat_settings(chat_id, title, last_ts) "
        "SELECT chat_id, MAX(title), MAX(ts) FROM chat_log GROUP BY chat_id",
    ],
    [  # v3: каждый видит только свои чаты; текст сообщений чатов бот больше не сохраняет
        "CREATE TABLE IF NOT EXISTS chat_users(chat_id INTEGER, user_id INTEGER, PRIMARY KEY(chat_id, user_id))",
        "INSERT OR IGNORE INTO chat_users(chat_id, user_id) "
        "SELECT DISTINCT chat_id, user_id FROM chat_log WHERE user_id IS NOT NULL",
        "ALTER TABLE polls ADD COLUMN audience INTEGER",  # сколько участников в чатах, где опубликован опрос
    ],
    [  # v4: автоматические напоминания с повтором
        "ALTER TABLE polls ADD COLUMN remind_min INTEGER DEFAULT 0",  # интервал повтора в минутах (0 — не напоминать)
        "ALTER TABLE polls ADD COLUMN notified_done INTEGER DEFAULT 0",  # автору уже сообщили, что опрос пройден
    ],
    [  # v5: черновики опросов переживают перезапуск бота
        "CREATE TABLE IF NOT EXISTS drafts(user_id INTEGER PRIMARY KEY, data TEXT, updated TEXT)",
    ],
    [  # v6: повторяющиеся опросы (серии)
        "ALTER TABLE polls ADD COLUMN rep_kind TEXT",  # daily | weekdays | weekly (только у исходного опроса серии)
        "ALTER TABLE polls ADD COLUMN rep_time TEXT",  # HH:MM по местному времени
        "ALTER TABLE polls ADD COLUMN rep_dow INTEGER",  # день недели для weekly
        "ALTER TABLE polls ADD COLUMN rep_next TEXT",  # ISO UTC следующего запуска
        "ALTER TABLE polls ADD COLUMN rep_dl_h INTEGER",  # срок каждого запуска в часах (None — до следующего запуска)
        "ALTER TABLE polls ADD COLUMN rep_series INTEGER",  # у копий: id исходного опроса серии
    ],
    [  # v7: сообщения бота с опросами в чатах — чтобы засчитывать ответы-реплаи
        "CREATE TABLE IF NOT EXISTS chat_posts(mid TEXT PRIMARY KEY, chat_id INTEGER, poll_id INTEGER)",
    ],
]


def migrate() -> None:
    ver = db.execute("PRAGMA user_version").fetchone()[0]
    for i in range(ver, len(MIGRATIONS)):
        for ddl in MIGRATIONS[i]:
            try:
                db.execute(ddl)
            except sqlite3.OperationalError as e:
                if "duplicate column" not in str(e).lower():
                    raise
        db.execute(f"PRAGMA user_version={i + 1}")
        db.commit()


migrate()


def get_setting(key: str, default: str = "") -> str:
    row = db.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    return row["value"] if row else default


def set_setting(key: str, value: str) -> None:
    db.execute("INSERT OR REPLACE INTO settings(key, value) VALUES(?,?)", (key, value))
    db.commit()


def now_utc() -> datetime:
    return datetime.now(timezone.utc)


def fmt_local(iso: str) -> str:
    return datetime.fromisoformat(iso).astimezone(TZ).strftime("%d.%m.%Y %H:%M")


def parse_local_dt(raw: str) -> datetime | None:
    """Разбирает «ДД.ММ.ГГГГ ЧЧ:ММ» в часовом поясе бота; None — формат неверный."""
    try:
        return datetime.strptime(raw.strip(), "%d.%m.%Y %H:%M").replace(tzinfo=TZ)
    except (ValueError, AttributeError):
        return None


def example_dt(plus_hours: int = 1) -> str:
    """Текущие дата и время (+ N часов) для подсказки-примера: пользователю остаётся лишь поправить."""
    return (datetime.now(TZ) + timedelta(hours=plus_hours)).strftime("%d.%m.%Y %H:%M")


def fmt_minutes(m: int) -> str:
    """90 -> «1 ч 30 мин», 10 -> «10 мин», 1440 -> «1 дн.»."""
    if m < 60:
        return f"{m} мин"
    if m % 1440 == 0:
        return f"{m // 1440} дн."
    if m % 60 == 0:
        return f"{m // 60} ч"
    return f"{m // 60} ч {m % 60} мин"


REP_LABEL = {"daily": "каждый день", "weekdays": "по будням (пн–пт)", "weekly": "раз в неделю"}
DOW_NAME = ["понедельникам", "вторникам", "средам", "четвергам", "пятницам", "субботам", "воскресеньям"]


def rep_label(kind: str | None, hhmm: str | None, dow: int | None = None) -> str:
    if not kind:
        return "не повторять"
    if kind == "weekly" and dow is not None:
        return f"по {DOW_NAME[dow]} в {hhmm}"
    return f"{REP_LABEL[kind]} в {hhmm}"


def next_run(kind: str, hhmm: str, after: datetime, dow: int | None = None) -> datetime:
    """Ближайший запуск серии строго после момента after (UTC). Время — по местному часовому поясу."""
    h, m = (int(x) for x in hhmm.split(":"))
    local = after.astimezone(TZ)
    cand = local.replace(hour=h, minute=m, second=0, microsecond=0)
    for _ in range(15):
        if cand > local and (
            kind == "daily"
            or (kind == "weekdays" and cand.weekday() < 5)
            or (kind == "weekly" and cand.weekday() == (dow if dow is not None else local.weekday()))
        ):
            return cand.astimezone(timezone.utc)
        cand += timedelta(days=1)
    return (after + timedelta(days=1)).astimezone(timezone.utc)


def parse_hhmm(raw: str) -> str | None:
    m = re.fullmatch(r"\s*(\d{1,2})[:.\-\s]?(\d{2})\s*", raw or "")
    if not m:
        return None
    h, mi = int(m.group(1)), int(m.group(2))
    return f"{h:02d}:{mi:02d}" if h < 24 and mi < 60 else None


def parse_interval(raw: str) -> int | None:
    """«10», «10 мин», «2 ч», «1.5 часа», «1 д» -> минуты; None — не разобрали."""
    m = re.fullmatch(r"\s*(\d+(?:[.,]\d+)?)\s*([а-яёa-z.]*)\s*", (raw or "").lower())
    if not m:
        return None
    num = float(m.group(1).replace(",", "."))
    unit = m.group(2)
    if unit.startswith(("д", "d")):
        mult = 1440
    elif unit.startswith(("ч", "h")):
        mult = 60
    elif unit == "" or unit.startswith(("м", "m")):
        mult = 1
    else:
        return None
    return int(round(num * mult))


def count(sql: str, params: tuple = ()) -> int:
    row = db.execute(sql, params).fetchone()
    return int(row[0]) if row else 0


def register_user(user_id: int, name: str) -> None:
    """Запоминает имя человека (нужно для выгрузки неанонимных опросов). Никаких ролей: все пользователи равны."""
    db.execute(
        "INSERT INTO users(user_id, name) VALUES(?,?) "
        "ON CONFLICT(user_id) DO UPDATE SET name=COALESCE(NULLIF(excluded.name,''), users.name)",
        (user_id, name or ""),
    )
    db.commit()


def save_chat_id(user_id: int, chat_id) -> None:
    if chat_id:
        db.execute("UPDATE users SET chat_id=? WHERE user_id=?", (chat_id, user_id))
        db.commit()


def get_poll(poll_id: int):
    row = db.execute("SELECT * FROM polls WHERE id=?", (poll_id,)).fetchone()
    if not row:
        return None
    poll = dict(row)
    poll["questions"] = json.loads(poll["questions"])
    return poll


def poll_open(poll: dict) -> bool:
    """Опрос принимает ответы: опубликован, не закрыт и дедлайн не прошёл."""
    if poll["status"] != "sent":
        return False
    dl = poll.get("deadline")
    if dl and datetime.fromisoformat(dl) <= now_utc():
        return False
    return True


def send_at_of(data: dict) -> datetime | None:
    s = data.get("send_at")
    return datetime.fromisoformat(s) if s else None


def compute_deadline(data: dict, send_at_iso: str | None) -> str | None:
    """Дедлайн в ISO UTC: либо задан датой, либо «через N часов после публикации»."""
    if data.get("deadline_abs"):
        return data["deadline_abs"]
    h = data.get("deadline_h")
    if h:
        base = datetime.fromisoformat(send_at_iso) if send_at_iso else now_utc()
        return (base + timedelta(hours=h)).isoformat()
    return None


def safe_cell(v):
    """Защита от формул в Excel: ответ вида =HYPERLINK(...) не выполнится."""
    if isinstance(v, str) and v[:1] in ("=", "+", "-", "@"):
        return "'" + v
    return v


def bot_link(poll_id: int) -> str:
    """Ссылка на опрос: её можно переслать любому контакту или в любой чат."""
    return f"https://max.ru/{BOT_USERNAME}?start=poll_{poll_id}" if BOT_USERNAME else ""


# ---------- ТЕСТЫ С ОЦЕНКОЙ ----------
def quiz_max(poll: dict) -> int:
    """Сколько вопросов с отмеченным правильным вариантом (макс. баллов)."""
    return sum(1 for q in poll["questions"] if q.get("correct") is not None)


def user_score(poll: dict, answers: dict) -> int:
    s = 0
    for i, q in enumerate(poll["questions"]):
        c = q.get("correct")
        if c is not None and q["options"] and answers.get(i) == q["options"][c]:
            s += 1
    return s


def user_answers(poll_id: int, uid: int) -> dict:
    return {
        r["q_idx"]: r["answer"]
        for r in db.execute("SELECT q_idx, answer FROM answers WHERE poll_id=? AND user_id=?", (poll_id, uid))
    }


# ---------- УВЕДОМЛЕНИЯ ОБ ОШИБКАХ, ПОВТОРЫ ЗАПРОСОВ ----------
RETRYABLE = ("429", "too many", "rate", "timeout", "timed out", "temporar", "502", "503", "504")


async def with_retry(fn, *args, tries: int = 4, **kwargs):
    """Повторяет вызов API с нарастающей паузой при временных ошибках и лимитах."""
    for i in range(tries):
        try:
            return await fn(*args, **kwargs)
        except Exception as e:
            if i == tries - 1 or not any(s in str(e).lower() for s in RETRYABLE):
                raise
            await asyncio.sleep(min(10.0, 0.5 * 2 ** i))


_bg_tasks: set = set()


class AdminLogHandler(logging.Handler):
    """Пересылает ошибки (уровень ERROR) людям из ADMIN_IDS; одинаковые — не чаще раза в 10 минут."""
    _last: dict[str, float] = {}

    def emit(self, record: logging.LogRecord) -> None:
        if record.levelno < logging.ERROR or not ADMIN_IDS:
            return
        msg = record.getMessage()
        if record.exc_info:
            msg += "\n" + "".join(traceback.format_exception_only(*record.exc_info[:2])).strip()
        key, now = msg[:120], time.time()
        if now - self._last.get(key, 0) < 600:
            return
        self._last[key] = now
        try:
            loop = asyncio.get_running_loop()
        except RuntimeError:
            return
        task = loop.create_task(notify_admins("🛠 Ошибка бота:\n" + msg[:1500]))
        _bg_tasks.add(task)
        task.add_done_callback(_bg_tasks.discard)


logging.getLogger().addHandler(AdminLogHandler())


# ---------- СОСТОЯНИЯ ----------
class Create(StatesGroup):
    question = State()
    options = State()
    custom_time = State()
    deadline = State()
    remind_custom = State()
    rep_time = State()
    edit_q = State()


class Answer(StatesGroup):
    free_text = State()  # ждём текстовый ответ на вопрос со свободным ответом


class Search(StatesGroup):  # ждём строку поиска по своим опросам
    polls = State()


# ---------- СПРАВКА ПО ЭКРАНАМ ----------
HELP = {
    "menu": (
        "❓ Справка\n\n"
        "Создать опрос — напишите вопросы, и бот даст ссылку. Её можно переслать кому угодно или отправить в чат.\n\n"
        "Мои опросы — результаты, файл с ответами, закрытие и удаление.\n\n"
        "Чтобы бот работал в чате, добавьте его в чат и назначьте администратором. "
        "Затем напишите в чате любое сообщение.\n\n"
        "Напоминания бот отправляет сам. Когда опрос пройден, он напишет Вам.\n"
        "Команда /menu возвращает в главное меню."
    ),
    "q": "❓ Справка: вопрос\n\nНапишите сообщением, о чём хотите спросить.",
    "mode": (
        "❓ Справка: тип ответа\n\n"
        "Выбор из вариантов — человек нажимает на кнопку. Можно отметить правильный ответ, тогда бот посчитает баллы.\n"
        "Свободный ответ — человек пишет своими словами."
    ),
    "opts": "❓ Справка: варианты\n\nНапишите варианты ответа одним сообщением, каждый с новой строки. Нужно минимум два.",
    "correct": (
        "❓ Справка: правильный ответ\n\n"
        "Отметьте правильный вариант, и бот посчитает баллы. Если это обычный опрос, нажмите «Без правильного ответа»."
    ),
    "more": (
        "❓ Справка: вопрос добавлен\n\n"
        "Ещё вопрос — добавить следующий.\n"
        "Изменить/удалить вопросы — исправить или убрать любой вопрос.\n"
        "Дальше — перейти к настройкам."
    ),
    "qlist": "❓ Справка: вопросы\n\nВыберите вопрос, чтобы изменить или удалить его.",
    "qedit": "❓ Справка: вопрос\n\nМожно изменить текст, поднять вопрос выше или удалить его.",
    "dest": (
        "❓ Справка: куда опубликовать\n\n"
        "Отметьте чат, и его участники ответят, не запуская бота.\n\n"
        "Для этого добавьте в чат бота и назначьте его администратором. "
        "Затем напишите в чате любое сообщение — чат появится в списке."
    ),
    "when": "❓ Справка: время\n\nВыберите, когда опубликовать опрос в чате: сразу, через час, завтра или в своё время.",
    "custom": (
        "❓ Справка: своя дата\n\n"
        "Напишите дату и время по образцу. Бот уже подставил сегодняшнюю дату — поправьте время."
    ),
    "extras": (
        "❓ Справка: дополнительно\n\n"
        "Анонимный — имена не показываются.\n"
        "Срок — когда приём ответов закроется сам.\n"
        "Напоминания — бот сам напоминает тем, кто не ответил. Ничего делать не нужно.\n"
        "Повтор — бот сам запускает опрос заново (например, каждое утро) и присылает итоги файлом.\n\n"
        "Всё это необязательно."
    ),
    "dl_custom": "❓ Справка: срок\n\nНапишите дату и время по образцу. Срок должен быть позже времени публикации.",
    "rep_time": "❓ Справка: время запуска\n\nНапишите время, когда бот будет запускать опрос, например 09:00.",
    "rm_custom": (
        "❓ Справка: свой интервал\n\n"
        "Напишите, как часто напоминать: «10 мин», «2 ч» или «1 д». "
        f"Не чаще чем раз в {REMIND_MIN_LIMIT} минут и не реже раза в 7 дней."
    ),
    "rep": (
        "❓ Справка: повтор опроса\n\n"
        "Бот сам запускает этот же опрос заново — например, каждое утро перекличка в чате. "
        "Перед новым запуском он закрывает прошлый и присылает Вам итоги отдельным файлом.\n"
        "Работает для опросов, опубликованных в чате. Остановить повтор можно в карточке опроса."
    ),
    "reopen": (
        "❓ Справка: открыть снова\n\n"
        "Бот предлагает прежние настройки, а срок отсчитывает заново — от сегодняшней даты. "
        "Нажмите «Открыть» или сначала измените срок и напоминания."
    ),
    "confirm": "❓ Справка: проверка\n\nПроверьте опрос. «Подтвердить» — создать и получить ссылку.",
    "my": (
        "❓ Справка: мои опросы\n\n"
        "Нажмите на опрос, чтобы открыть его. В скобках — сколько человек ответило. "
        "Кнопка «Поиск» ищет по тексту вопроса или номеру."
    ),
    "poll": (
        "❓ Справка: опрос\n\n"
        "Результаты — ответы прямо в чате.\n"
        "Файл — ответы в Excel и CSV.\n"
        "Ссылка — сообщение, которое можно переслать.\n"
        "Копировать — создать новый опрос на основе этого.\n"
        "Закрыть приём — ответы больше не принимаются.\n"
        "Удалить — удалить опрос и все ответы."
    ),
    "series": (
        "❓ Справка: серия опросов\n\n"
        "Серия — это опрос, который бот сам запускает заново. Каждый запуск подписан датой и временем. "
        "Нажмите на запуск, чтобы посмотреть результаты и файл за этот день. "
        "«Настройки серии» — карточка исходного опроса."
    ),
    "unfinished": "❓ Справка\n\nЗдесь опросы, которые Вы начали, но не закончили. Бот продолжит с того же места.",
    "take": "❓ Справка: опрос\n\nВыберите вариант кнопкой или напишите ответ. «Назад» — вернуться к прошлому вопросу.",
}

ROOT_SCREENS = {"menu"}  # на главном экране кнопок «Назад»/«Главная» нет


# ---------- ХЕЛПЕРЫ ----------
# Эмодзи съедают место на кнопках, и подписи обрезаются. Поэтому на кнопках картинки убираем автоматически.
EMOJI_RE = re.compile("[\U0001F000-\U0001FAFF\u2300-\u23FF\u25B6\u25C0\u2600-\u27BF\u2B00-\u2BFF\uFE0F\u200D]+")
KEEP_EMOJI = ("ans:", "qv:", "cor:")  # кнопки с текстом вариантов ответа от автора — не трогаем


def plain_label(text: str, payload: str) -> str:
    if payload.startswith(KEEP_EMOJI):
        return text
    t = re.sub(r"\s{2,}", " ", EMOJI_RE.sub("", text)).strip()
    return t or text


def kb_of(rows: list[list[tuple[str, str]]]):
    b = InlineKeyboardBuilder()
    for row in rows:
        btns = []
        for t, p in row:
            t = plain_label(t, p)
            if p.startswith("url:"):  # кнопка-ссылка
                if LinkButton:
                    btns.append(LinkButton(text=t, url=p[4:]))
            else:
                btns.append(CallbackButton(text=t, payload=p))
        if btns:
            b.row(*btns)
    return b.as_markup()


def with_nav(rows, screen: str):
    """Нижний ряд: «Назад» слева, «Главная», «Помощь»."""
    rows = [list(r) for r in (rows or [])]
    if screen in ROOT_SCREENS:
        rows.append([("Помощь", f"help:{screen}")])
    else:
        rows.append([("Назад", f"back:{screen}"), ("Главная", "home"), ("Помощь", f"help:{screen}")])
    return rows


def payload_of(event: MessageCallback) -> str:
    return event.callback.payload or ""


def page_slice(items: list, page: int, size: int = PAGE_SIZE):
    """Возвращает (элементы страницы, номер страницы, всего страниц)."""
    pages = max(1, -(-len(items) // size))
    page = min(max(page, 0), pages - 1)
    return items[page * size:(page + 1) * size], page, pages


def pager_row(prefix: str, page: int, pages: int) -> list[tuple[str, str]]:
    row = []
    if page > 0:
        row.append(("‹ Пред.", f"{prefix}:{page - 1}"))
    row.append((f"{page + 1}/{pages}", "noop"))
    if page < pages - 1:
        row.append(("След. ›", f"{prefix}:{page + 1}"))
    return row


async def send(user_id: int, text: str, rows=None, screen: str | None = None):
    """screen — имя экрана: если задан, снизу добавляются кнопки «Назад», «Главная», «Помощь»."""
    if screen:
        rows = with_nav(rows, screen)
    attachments = [kb_of(rows)] if rows else None
    await with_retry(bot.send_message, user_id=user_id, text=text, attachments=attachments)


async def deliver(user_id: int, text: str, rows=None) -> str | None:
    """Надёжная отправка: сначала по chat_id личного диалога, затем по user_id.
    Возвращает None при успехе или текст ошибки."""
    attachments = [kb_of(rows)] if rows else None
    errors = []
    row = db.execute("SELECT chat_id FROM users WHERE user_id=?", (user_id,)).fetchone()
    if row and row["chat_id"]:
        try:
            await with_retry(bot.send_message, chat_id=row["chat_id"], text=text, attachments=attachments)
            return None
        except Exception as e:
            errors.append(f"chat_id: {e}")
    try:
        await with_retry(bot.send_message, user_id=user_id, text=text, attachments=attachments)
        return None
    except Exception as e:
        errors.append(f"user_id: {e}")
    return "; ".join(errors)


async def notify_admins(text: str) -> None:
    for uid in ADMIN_IDS:
        await deliver(uid, text)


async def refresh(event: MessageCallback, text: str, rows=None, screen: str | None = None):
    """Обновляет сообщение с кнопками (или шлёт новое, если не вышло)."""
    if screen:
        rows = with_nav(rows, screen)
    attachments = [kb_of(rows)] if rows else []
    try:
        await event.answer(new_text=text, attachments=attachments)
    except Exception:
        await send(event.callback.user.user_id, text, rows)


async def render(uid: int, text: str, rows, screen: str, event: MessageCallback | None = None):
    """Показывает экран: правит сообщение (если пришли с кнопки) или шлёт новое."""
    if event is not None:
        await refresh(event, text, rows, screen=screen)
    else:
        await send(uid, text, rows, screen=screen)


async def ack(event: MessageCallback):
    try:
        await event.answer(notification="👌")
    except Exception:
        pass


async def notify(event: MessageCallback, text: str):
    """Всплывающее уведомление; ошибка API не должна ломать обработчик."""
    try:
        await event.answer(notification=text)
    except Exception:
        log.warning("notify failed", exc_info=True)


@dp.message_callback(F.callback.payload == "noop")
async def cb_noop(event: MessageCallback):
    await ack(event)


# ---------- ГЛАВНОЕ МЕНЮ (одинаковое для всех) ----------
def unfinished_polls(uid: int) -> list:
    """Опросы, к которым человек присоединился, но не закончил, и которые ещё открыты."""
    res = []
    for r in db.execute(
        "SELECT p.id FROM polls p JOIN recipients r ON r.poll_id=p.id "
        "WHERE r.user_id=? AND r.done=0 AND r.sent=1 AND p.status='sent' ORDER BY p.id DESC",
        (uid,),
    ).fetchall():
        poll = get_poll(r["id"])
        if poll and poll_open(poll):
            res.append(poll)
    return res


DRAFT_KEYS = ("questions", "chat_sel", "send_at", "anon", "deadline_h", "deadline_abs", "remind_min", "rep_kind", "rep_time")


async def save_draft(uid: int, context: MemoryContext) -> None:
    """Сохраняет недособранный опрос в БД: после перезапуска бота ничего не теряется."""
    try:
        data = await context.get_data()
        if not data.get("questions") or data.get("reopen_id"):
            return
        snap = {k: data.get(k) for k in DRAFT_KEYS}
        db.execute(
            "INSERT OR REPLACE INTO drafts(user_id, data, updated) VALUES(?,?,?)",
            (uid, json.dumps(snap, ensure_ascii=False), now_utc().isoformat()),
        )
        db.commit()
    except Exception:
        log.warning("save_draft failed", exc_info=True)


def load_draft(uid: int) -> dict | None:
    row = db.execute("SELECT data FROM drafts WHERE user_id=?", (uid,)).fetchone()
    if not row:
        return None
    try:
        d = json.loads(row["data"])
    except ValueError:
        return None
    return d if d.get("questions") else None


def delete_draft(uid: int) -> None:
    db.execute("DELETE FROM drafts WHERE user_id=?", (uid,))
    db.commit()


async def restore_draft(uid: int, context: MemoryContext, screen: str = "more") -> bool:
    """Возвращает черновик в мастер и показывает нужный экран (по умолчанию «Вопрос добавлен»)."""
    d = load_draft(uid)
    if not d:
        return False
    await context.clear()
    await context.update_data(cpage=0, **{k: v for k, v in d.items() if v is not None})
    await context.update_data(
        chat_sel=d.get("chat_sel") or [], anon=bool(d.get("anon")), remind_min=d.get("remind_min") or 0,
    )
    if screen == "confirm":  # человек нажал «Подтвердить» — показываем проверку опроса, чтобы нажал ещё раз
        data = await context.get_data()
        await send(
            uid,
            "♻️ Бот перезапускался, но Ваш опрос сохранён. Проверьте его и нажмите «Подтвердить» ещё раз.\n\n"
            + confirm_text(data),
            CONFIRM_ROWS, screen="confirm",
        )
    else:
        await goto(screen, uid, context)
    return True


async def stale_session(event: MessageCallback, context: MemoryContext, screen: str = "more") -> None:
    uid = event.callback.user.user_id
    if await restore_draft(uid, context, screen):
        await notify(event, "Бот перезапускался — черновик опроса восстановлен")
    else:
        await notify(event, "Сессия устарела. Создайте опрос заново.")
        await context.clear()
        await show_menu(uid)


MENU_TEXT = (
    "Я «Спрашиватель»!\n"
    "Помогу создать опросы и сохранить результаты в файл.\n\n"
    "О чём будем спрашивать?"
)

# Настроения Спрашивателя: при каждом открытии главного меню бот «в новом настроении».
# Формат: (лицо, фраза). Новые настроения можно просто дописывать в список.
MOODS = [
    ("(◕‿◕)", "Рад Вас видеть!"),
    ("📋", "Бланки наготове!"),
    ("📝", "Ручка заточена, вопросы ждут!"),
    ("👀", "Любопытствую, что на этот раз?"),
    ("シ", "Настроение — боевое!"),
    ("(•̀ᴗ•́)و", "Готов к опросам!"),
    ("ヽ(´▽`)/", "Какой хороший день для опроса!"),
    ("(｡•̀ᴗ-)✧", "Сегодня всё получится!"),
    ("(ᵔᴥᵔ)", "Опрашивать — моё любимое дело!"),
    ("ʕ•ᴥ•ʔ", "Спрашиваю по-медвежьи основательно!"),
    ("(¬‿¬)", "Кто-то сегодня не ответит… Но я напомню."),
    ("(｀・ω・´)ゞ", "Служу опросам!"),
    ("٩(◕‿◕)۶", "Ну что, поспрашиваем?"),
    ("🧐", "Хм, сейчас всё выясним!"),
    ("🕵️", "Выясню, кто ещё не ответил."),
    ("📊", "Цифры любят точность!"),
    ("☕", "Кофе выпит — можно спрашивать."),
    ("🎯", "Бью точно в вопрос!"),
    ("(☞ﾟヮﾟ)☞", "Давайте что-нибудь спросим!"),
    ("(ง'̀-'́)ง", "Не отстану, пока не ответят!"),
]
_last_mood: dict[int, int] = {}  # чтобы настроение не повторялось два раза подряд у одного человека


def day_part() -> str:
    h = datetime.now(TZ).hour
    if 5 <= h < 12:
        return "Доброе утро!"
    if 12 <= h < 17:
        return "Добрый день!"
    if 17 <= h < 23:
        return "Добрый вечер!"
    return "Не спится? Я тоже на посту!"


def menu_text(uid: int) -> str:
    """Главное меню с настроением: случайное лицо и фраза, иногда — приветствие по времени суток."""
    idx = random.randrange(len(MOODS))
    if len(MOODS) > 1 and idx == _last_mood.get(uid):
        idx = (idx + 1) % len(MOODS)
    _last_mood[uid] = idx
    face, phrase = MOODS[idx]
    head = f"{face}  {phrase}"
    if random.random() < 0.3:  # в каждом третьем меню — приветствие по времени суток
        head = f"{face}  {day_part()}"
    return f"{head}\n\n{MENU_TEXT}"


async def show_menu(uid: int, text: str | None = None):
    if text is None:
        text = menu_text(uid)
    rows = [[("Создать опрос", "new")], [("Мои опросы", "my")]]
    if load_draft(uid):
        rows.insert(1, [("Продолжить черновик опроса", "draft")])
    n = len(unfinished_polls(uid))
    if n:
        rows.append([(f"Продолжить опрос ({n})", "unfinished")])
    await send(uid, text, rows, screen="menu")


@dp.message_callback(F.callback.payload == "draft")
async def cb_draft(event: MessageCallback, context: MemoryContext):
    await ack(event)
    uid = event.callback.user.user_id
    if not await restore_draft(uid, context):
        await send(uid, "Черновик не найден.")
        await show_menu(uid)


@dp.message_callback(F.callback.payload == "home")
async def cb_home(event: MessageCallback, context: MemoryContext):
    await ack(event)
    await context.clear()
    await show_menu(event.callback.user.user_id)


@dp.message_callback(F.callback.payload == "unfinished")
async def cb_unfinished(event: MessageCallback, context: MemoryContext):
    await ack(event)
    await context.set_state(None)
    uid = event.callback.user.user_id
    polls = unfinished_polls(uid)
    if not polls:
        await send(uid, "Незаконченных опросов нет 🙂", screen="unfinished")
        return
    rows = [[(f"{p['questions'][0]['text'][:40]}", f"take:{p['id']}")] for p in polls[:15]]
    await send(uid, "Опросы, которые Вы начали:", rows, screen="unfinished")


# ---------- ВХОД В ОПРОС ПО ССЫЛКЕ ИЛИ ИЗ ЧАТА ----------
JOIN_MSG = {
    "closed": "🔒 Приём ответов по этому опросу закрыт.",
    "done": "Вы уже прошли этот опрос. Спасибо! 🙂",
    "missing": "Опрос не найден (возможно, автор его удалил).",
    "scheduled": "⏳ Этот опрос ещё не открыт — попробуйте позже.",
}


def join_poll(uid: int, poll_id: int) -> str:
    """Допуск человека к опросу: любой, у кого есть ссылка или кто видит его в чате.
    Возвращает 'ok' | 'done' | 'closed' | 'scheduled' | 'missing'."""
    poll = get_poll(poll_id)
    if not poll:
        return "missing"
    if poll["status"] == "scheduled":
        return "scheduled"
    if not poll_open(poll):
        return "closed"
    row = db.execute("SELECT done FROM recipients WHERE poll_id=? AND user_id=?", (poll_id, uid)).fetchone()
    if row is None:
        db.execute("INSERT INTO recipients(poll_id, user_id, sent, done) VALUES(?,?,1,0)", (poll_id, uid))
    elif row["done"]:
        return "done"
    else:
        db.execute("UPDATE recipients SET sent=1 WHERE poll_id=? AND user_id=?", (poll_id, uid))
    db.commit()
    return "ok"


def poll_card_text(poll: dict, prefix: str = "📋 Новый опрос!", with_link: bool = False) -> str:
    n = len(poll["questions"])
    text = f"{prefix}\n\n{poll['questions'][0]['text']}" + (f"\n(вопросов: {n})" if n > 1 else "")
    if poll.get("anon"):
        text += "\n🕶 Опрос анонимный — в отчётах имена не показываются."
    if poll.get("deadline"):
        text += f"\n⏰ Ответить до {fmt_local(poll['deadline'])}"
    if with_link and bot_link(poll["id"]):
        text += f"\n\n🔗 Пройти опрос: {bot_link(poll['id'])}"
    return text


async def present_poll(uid: int, poll_id: int) -> str | None:
    """Показывает человеку карточку опроса с кнопкой «Пройти». Возвращает ошибку доставки или None."""
    res = join_poll(uid, poll_id)
    poll = get_poll(poll_id)
    if res == "ok":
        return await deliver(uid, poll_card_text(poll), [[("Пройти опрос", f"take:{poll_id}")]])
    if res == "done" and poll and poll_open(poll):
        return await deliver(
            uid, "Вы уже прошли этот опрос. Пока он открыт, ответы можно изменить.",
            [[("Изменить ответы", f"redo:{poll_id}")]],
        )
    return await deliver(uid, JOIN_MSG.get(res, JOIN_MSG["missing"]))


def _start_arg(text: str) -> str:
    parts = (text or "").split()
    return parts[1] if len(parts) > 1 else ""


@dp.bot_started()
async def on_bot_started(event: BotStarted):
    uid = event.user.user_id
    register_user(uid, event.user.first_name or "")
    save_chat_id(uid, event.chat_id)
    payload = getattr(event, "payload", None) or ""
    # человек пришёл по ссылке на опрос: ?start=poll_5
    if payload.startswith("poll_") and payload[5:].isdigit():
        await present_poll(uid, int(payload[5:]))
        return
    await show_menu(uid)


@dp.message_created(Command("start"))
async def on_start(event: MessageCreated, context: MemoryContext):
    await context.clear()
    uid = event.message.sender.user_id
    register_user(uid, event.message.sender.first_name or "")
    if "dialog" in str(getattr(event.message.recipient, "chat_type", "")).lower():
        save_chat_id(uid, event.message.recipient.chat_id)  # только личный чат, не групповой
    arg = _start_arg(event.message.body.text if event.message.body else "")
    if arg.startswith("poll_") and arg[5:].isdigit():
        await start_poll_direct(uid, context, int(arg[5:]))
        return
    await show_menu(uid)


async def start_poll_direct(uid: int, context: MemoryContext, poll_id: int):
    """Человек пришёл по ссылке: сразу показываем первый вопрос, без лишней кнопки."""
    res = join_poll(uid, poll_id)
    poll = get_poll(poll_id)
    if res != "ok" or not poll:
        await present_poll(uid, poll_id)  # закрыт / уже прошёл / не найден — объясняем
        return
    answered = count("SELECT COUNT(*) FROM answers WHERE poll_id=? AND user_id=?", (poll_id, uid))
    total = len(poll["questions"])
    start = answered if 0 < answered < total else 0
    if start == 0 and answered:
        db.execute("DELETE FROM answers WHERE poll_id=? AND user_id=?", (poll_id, uid))
        db.commit()
    await context.update_data(redo=False)
    await send(uid, poll_card_text(poll, "📋 Опрос"))
    await ask_question(uid, context, poll_id, start)


@dp.message_created(Command("menu"))
async def on_menu(event: MessageCreated, context: MemoryContext):
    await context.clear()
    await show_menu(event.message.sender.user_id)


@dp.message_created(Command("help"))
async def on_help_cmd(event: MessageCreated):
    await send(event.message.sender.user_id, HELP["menu"])


@dp.message_callback(F.callback.payload.startswith("help:"))
async def cb_help(event: MessageCallback):
    await ack(event)
    screen = payload_of(event).split(":", 1)[1]
    await send(event.callback.user.user_id, HELP.get(screen, "Справка для этого экрана пока не написана."))


# ---------- ТЕКСТЫ И КНОПКИ МАСТЕРА ----------
Q_PROMPT = "✍️ Введите текст вопроса — то, что Вы хотите спросить."
MODE_ROWS = [
    [("Тест: выбор из вариантов", "mode:test")],
    [("Свободный ответ текстом", "mode:free")],
]
OPTS_PROMPT = (
    "🔘 Введите варианты ответа одним сообщением — каждый с новой строки (минимум 2).\n\n"
    "Пример:\nВариант А\nВариант Б\nВариант В"
)
WHEN_ROWS = [
    [("Сразу", "when:now")],
    [("Через 1 час", "when:1h")],
    [("Завтра в 09:00", "when:tom9")],
    [("Выбрать дату и время", "when:custom")],
]
DEST_TEXT = (
    "📤 Куда опубликовать?\n\n"
    "Отметьте чат, и его участники ответят, не запуская бота.\n\n"
    "Для автоматической публикации опроса добавьте в чат бота и назначьте ему права администратора."
)
DEST_EMPTY = (
    "\n\nСписок чатов пуст. Чтобы чат появился: добавьте бота в чат, дайте ему права администратора и напишите там любое сообщение. "
    "Либо просто нажмите «Только ссылка»."
)
EXTRAS_TEXT = (
    "⚙️ Дополнительно (всё необязательно)\n\n"
    "🕶 Анонимный — в результатах и файле не показываются имена.\n"
    "⏰ Срок — после срока приём ответов закрывается сам.\n"
    "🔔 Напоминания — бот сам повторяет их с выбранным интервалом тем, кто начал, но не закончил, "
    "пока человек не ответит или опрос не закроется. Вам придёт только сообщение, что опрос пройден.\n"
    "🔁 Повтор — бот сам запускает опрос заново (например, каждое утро) и присылает итоги отдельным файлом."
)
CONFIRM_ROWS = [[("Подтвердить", "confirm")], [("Отмена", "cancel")]]
REMIND_PRESETS = (10, 30, 60, 180, 1440)


def my_chats(uid: int) -> list:
    """Чаты, где есть бот и где этот человек писал сообщения (чужие чаты не показываются)."""
    return db.execute(
        "SELECT s.chat_id, s.title FROM chat_users u JOIN chat_settings s ON s.chat_id=u.chat_id "
        "WHERE u.user_id=? ORDER BY s.last_ts DESC",
        (uid,),
    ).fetchall()


def dest_screen(uid: int, data: dict):
    chats = my_chats(uid)
    csel = set(data.get("chat_sel", []))
    items, page, pages = page_slice(chats, data.get("cpage") or 0)
    rows = [
        [(("● " if c["chat_id"] in csel else "○ ") + f"{c['title'] or c['chat_id']}", f"ct:{c['chat_id']}")]
        for c in items
    ]
    if pages > 1:
        rows.append(pager_row("cp", page, pages))
    rows.append([(f"Дальше ({len(csel)} чат.)" if csel else "Только ссылка", "dnext")])
    return DEST_TEXT + ("" if chats else DEST_EMPTY), rows


def extras_rows(data: dict):
    dh = data.get("deadline_h") or 0
    dabs = data.get("deadline_abs")
    rm = data.get("remind_min") or 0

    def mk(label: str, on) -> str:
        return ("● " if on else "") + label

    custom_rm = rm and rm not in REMIND_PRESETS
    rows = [[(mk("Анонимный", data.get("anon")), "ex:anon")]]
    rows.append([
        (mk("Без срока", not dh and not dabs), "ex:dl:0"),
        (mk("Срок 24 ч", dh == 24 and not dabs), "ex:dl:24"),
    ])
    rows.append([
        (mk("Срок 3 дня", dh == 72 and not dabs), "ex:dl:72"),
        (mk("Срок 7 дней", dh == 168 and not dabs), "ex:dl:168"),
    ])
    rows.append([(mk("Срок: своя дата", bool(dabs)), "ex:dl:custom")])
    rows.append([(mk("Не напоминать", not rm), "ex:rm:0")])
    rows.append([
        (mk("Раз в 10 мин", rm == 10), "ex:rm:10"),
        (mk("Раз в 30 мин", rm == 30), "ex:rm:30"),
    ])
    rows.append([
        (mk("Раз в 1 час", rm == 60), "ex:rm:60"),
        (mk("Раз в 3 часа", rm == 180), "ex:rm:180"),
    ])
    rows.append([
        (mk("Раз в сутки", rm == 1440), "ex:rm:1440"),
        (mk(f"Свой: {fmt_minutes(rm)}" if custom_rm else "Свой интервал", bool(custom_rm)), "ex:rm:custom"),
    ])
    if not data.get("reopen_id"):
        rk = data.get("rep_kind")
        rows.append([(mk(("Повтор: " + rep_label(rk, data.get("rep_time") or "09:00")) if rk else "Повтор опроса", bool(rk)), "ex:rep")])
    rows.append([("Далее", "ex:next")])
    return rows


def rep_screen(data: dict):
    rk = data.get("rep_kind")
    t = data.get("rep_time") or "09:00"
    if not (data.get("chat_sel")):
        return (
            "🔁 Повтор опроса\n\nПовтор работает для опросов, опубликованных в чате. "
            "Вернитесь на шаг «Куда опубликовать» и отметьте чат.",
            [],
        )
    text = (
        "🔁 Повтор опроса\n\n"
        "Бот сам запустит этот же опрос заново с теми же настройками. Перед новым запуском он закроет "
        "прошлый и пришлёт Вам итоги отдельным файлом. Например, каждое утро — перекличка в чате.\n\n"
        f"Сейчас: {rep_label(rk, t)}"
    )
    if rk and not (data.get("deadline_h") or data.get("deadline_abs")):
        text += "\nСрока нет, поэтому каждый опрос закроется в момент следующего запуска."

    def mk(label, on):
        return ("● " if on else "") + label

    rows = [
        [(mk("Не повторять", not rk), "rep:none")],
        [(mk("Каждый день", rk == "daily"), "rep:daily"), (mk("По будням", rk == "weekdays"), "rep:weekdays")],
        [(mk("Раз в неделю", rk == "weekly"), "rep:weekly")],
        [(mk(f"Время {t}", False), "rep:time")],
        [(mk(x, t == x), "rep:t:" + x.replace(":", "")) for x in ("08:00", "09:00", "12:00", "18:00")],
        [("Готово", "rep:done")],
    ]
    return text, rows


def qlist_screen(data: dict):
    qs = data.get("questions", [])
    if not qs:
        return "Вопросов нет.", [[("Добавить вопрос", "addq")]]
    rows = [[(f"{i + 1}. {q['text'][:40]}", f"qe:{i}")] for i, q in enumerate(qs)]
    return "✏️ Выберите вопрос, чтобы изменить или удалить его:", rows


def qedit_screen(data: dict, i: int):
    q = data["questions"][i]
    lines = [f"Вопрос {i + 1}: {q['text']}"]
    if q["options"]:
        for j, o in enumerate(q["options"]):
            lines.append(("  ✔ " if q.get("correct") == j else "  • ") + o)
    else:
        lines.append("  (свободный ответ)")
    rows = [[("Изменить текст", f"qt:{i}")]]
    if i > 0:
        rows.append([("Выше", f"qu:{i}")])
    rows.append([("Удалить вопрос", f"qd:{i}")])
    return "\n".join(lines), rows


async def goto(screen: str, uid: int, context: MemoryContext):
    """Показывает экран мастера создания опроса (новым сообщением)."""
    if screen != "menu":
        await save_draft(uid, context)
    data = await context.get_data()
    if screen == "menu":
        await context.clear()
        await show_menu(uid)
    elif screen == "q":
        await context.set_state(Create.question)
        await send(uid, Q_PROMPT, screen="q")
    elif screen == "mode":
        await context.set_state(None)
        await send(uid, f"Вопрос: «{data.get('pending_q', '')}»\n\nКак будут отвечать?", MODE_ROWS, screen="mode")
    elif screen == "opts":
        await context.set_state(Create.options)
        await send(uid, OPTS_PROMPT, screen="opts")
    elif screen == "correct":
        await context.set_state(None)
        opts = data.get("pending_opts") or []
        rows = [[(f"{i + 1}. {o[:40]}", f"cor:{i}")] for i, o in enumerate(opts)]
        rows.append([("Без правильного ответа", "cor:-")])
        await send(
            uid,
            "🎯 Какой вариант правильный? Это нужно для оценки (баллы, процент, рейтинг). "
            "Для обычного опроса — «Без правильного ответа».",
            rows,
            screen="correct",
        )
    elif screen == "more":
        await context.set_state(None)
        n = len(data.get("questions", []))
        await send(
            uid,
            f"✅ Вопрос добавлен (всего в опросе: {n}).",
            [
                [("Ещё вопрос", "addq")],
                [("Изменить/удалить вопросы", "qlist")],
                [("Дальше", "dest")],
            ],
            screen="more",
        )
    elif screen == "qlist":
        await context.set_state(None)
        text, rows = qlist_screen(data)
        await send(uid, text, rows, screen="qlist")
    elif screen == "dest":
        await context.set_state(None)
        text, rows = dest_screen(uid, data)
        await send(uid, text, rows, screen="dest")
    elif screen == "when":
        await context.set_state(None)
        await send(uid, "🕒 Когда опубликовать в чате?", WHEN_ROWS, screen="when")
    elif screen == "extras":
        await context.set_state(None)
        await send(uid, EXTRAS_TEXT, extras_rows(data), screen="extras")


@dp.message_callback(F.callback.payload.startswith("back:"))
async def cb_back(event: MessageCallback, context: MemoryContext):
    await ack(event)
    uid = event.callback.user.user_id
    screen = payload_of(event).split(":", 1)[1]
    data = await context.get_data()

    # ---- прохождение опроса: на предыдущий вопрос / в меню
    if screen == "take":
        poll_id, idx = data.get("poll_id"), data.get("idx", 0)
        if poll_id is None or idx == 0:
            await context.clear()
            await show_menu(uid)
        else:
            if not data.get("redo"):  # при изменении ответов старые ответы не трогаем
                db.execute(
                    "DELETE FROM answers WHERE poll_id=? AND user_id=? AND q_idx>=?",
                    (poll_id, uid, idx - 1),
                )
                db.commit()
            await ask_question(uid, context, poll_id, idx - 1)
        return

    # ---- мастер создания
    if screen == "q":
        await goto("more" if data.get("questions") else "menu", uid, context)
    elif screen == "mode":
        await goto("q", uid, context)
    elif screen == "opts":
        await goto("mode", uid, context)
    elif screen == "correct":
        await goto("opts", uid, context)
    elif screen == "more":
        # ничего не удаляем: показываем список вопросов, откуда их можно изменить или удалить
        if data.get("questions"):
            await context.update_data(ql_back=True)
            await goto("qlist", uid, context)
        else:
            await goto("q", uid, context)
    elif screen == "qlist":
        if data.get("ql_back"):  # пришли сюда кнопкой «Назад» с экрана «Вопрос добавлен»
            await context.update_data(ql_back=False)
            await goto("dest", uid, context)
        else:
            await goto("more", uid, context)
    elif screen == "qedit":
        await goto("qlist", uid, context)
    elif screen == "dest":
        await goto("more", uid, context)
    elif screen == "when":
        await goto("dest", uid, context)
    elif screen == "custom":
        await goto("when", uid, context)
    elif screen == "reopen":
        await context.clear()
        await show_polls(uid, context)
    elif screen == "extras" and data.get("reopen_id"):
        await send(uid, reopen_text(data), reopen_rows(data), screen="reopen")
    elif screen == "rep":
        await send(uid, EXTRAS_TEXT, extras_rows(data), screen="extras")
    elif screen == "rep_time":
        text, rows = rep_screen(data)
        await send(uid, text, rows, screen="rep")
    elif screen in ("dl_custom", "rm_custom") and data.get("reopen_id"):
        await send(uid, EXTRAS_TEXT, extras_rows(data), screen="extras")
    elif screen == "extras":
        await goto("when" if data.get("chat_sel") else "dest", uid, context)
    elif screen in ("dl_custom", "rm_custom", "confirm"):
        await goto("extras", uid, context)
    # ---- мои опросы
    elif screen == "series":
        await context.set_state(None)
        await show_polls(uid, context)
    elif screen == "poll":  # карточка/результаты опроса → список опросов
        await context.set_state(None)
        await show_polls(uid, context)
    else:  # my, unfinished и всё остальное
        await goto("menu", uid, context)


# ---------- МОИ ОПРОСЫ: СПИСОК, ПОИСК ----------
def poll_done_count(poll_id: int) -> int:
    return count("SELECT COUNT(*) FROM recipients WHERE poll_id=? AND done=1", (poll_id,))


def series_runs(head_id: int) -> list:
    """Все запуски серии (исходный опрос и его копии), новые сверху."""
    return db.execute(
        "SELECT * FROM polls WHERE id=? OR rep_series=? ORDER BY id DESC", (head_id, head_id)
    ).fetchall()


def is_series_head(p) -> bool:
    return bool(p["rep_kind"]) or count("SELECT COUNT(*) FROM polls WHERE rep_series=?", (p["id"],)) > 0


def owner_polls(owner_id: int, q: str = "") -> list:
    res = []
    # отдельные запуски серий в общем списке не показываются — они внутри серии
    for p in db.execute("SELECT * FROM polls WHERE owner_id=? AND rep_series IS NULL ORDER BY id DESC", (owner_id,)):
        if q:
            text = " ".join(x["text"] for x in json.loads(p["questions"])).lower()
            if q not in text and q.lstrip("#") != str(p["id"]):
                continue
        res.append(p)
    return res


def polls_screen(uid: int, q: str, page: int):
    polls = owner_polls(uid, q)
    if not polls:
        if q:
            return f"Ничего не найдено по запросу «{q}».", [[("Сбросить поиск", "qc:polls")]], "my"
        return "У Вас пока нет опросов. Создайте первый!", [[("Создать опрос", "new")]], "my"
    items, page, pages = page_slice(polls, page)
    rows = []
    for p in items:
        first = json.loads(p["questions"])[0]["text"][:26]
        mark = {"scheduled": "ждёт · ", "closed": "закрыт · "}.get(p["status"], "")
        if is_series_head(p):
            n = len(series_runs(p["id"]))
            rows.append([(f"▸ Серия: {first} ({n} зап.)", f"sr:{p['id']}:0")])
        else:
            rows.append([(f"{mark}#{p['id']} {first} ({poll_done_count(p['id'])})", f"pm:{p['id']}")])
    if pages > 1:
        rows.append(pager_row("mp", page, pages))
    rows.append([("Поиск", "qs:polls")] + ([("Сбросить", "qc:polls")] if q else []))
    text = (
        "📋 Ваши опросы (пометки: «ждёт» — запланирован, «закрыт» — приём закрыт, без пометки — принимает ответы; "
        "в скобках — сколько ответило). Повторяющиеся опросы собраны в серии «▸» — нажмите, чтобы увидеть все запуски.\nВидите их только Вы."
        + (f"\n🔎 Поиск: «{q}»" if q else "")
        + f"\nВсего: {len(polls)}"
    )
    return text, rows, "my"


async def show_polls(uid: int, context: MemoryContext, page: int = 0, event: MessageCallback | None = None):
    data = await context.get_data()
    q = (data.get("pq") or "").strip().lower()
    text, rows, screen = polls_screen(uid, q, page)
    await render(uid, text, rows, screen, event)


@dp.message_callback(F.callback.payload == "my")
async def cb_my(event: MessageCallback, context: MemoryContext):
    await ack(event)
    await context.set_state(None)
    await context.update_data(pq="")
    await show_polls(event.callback.user.user_id, context)


@dp.message_callback(F.callback.payload.startswith("mp:"))
async def cb_polls_page(event: MessageCallback, context: MemoryContext):
    await context.set_state(None)
    page = int(payload_of(event).split(":")[1])
    await show_polls(event.callback.user.user_id, context, page, event)


@dp.message_callback(F.callback.payload == "qs:polls")
async def cb_search_start(event: MessageCallback, context: MemoryContext):
    await ack(event)
    await context.set_state(Search.polls)
    await send(
        event.callback.user.user_id,
        "🔎 Введите часть текста вопроса или номер опроса (например 12):",
        [[("Отмена", "mp:0")]],
    )


@dp.message_callback(F.callback.payload == "qc:polls")
async def cb_search_clear(event: MessageCallback, context: MemoryContext):
    await context.set_state(None)
    await context.update_data(pq="")
    await show_polls(event.callback.user.user_id, context, 0, event)


@dp.message_created(Search.polls)
async def on_search_text(event: MessageCreated, context: MemoryContext):
    text = ((event.message.body.text or "") if event.message.body else "").strip()
    await context.set_state(None)
    await context.update_data(pq=text)
    await show_polls(event.message.sender.user_id, context)


# ---------- СОЗДАНИЕ ОПРОСА ----------
# Шаги: вопросы → (правильный ответ) → куда опубликовать → (время) → дополнительно → подтверждение
@dp.message_callback(F.callback.payload == "new")
async def cb_new(event: MessageCallback, context: MemoryContext):
    await ack(event)
    await context.clear()
    delete_draft(event.callback.user.user_id)
    await context.update_data(
        questions=[], chat_sel=[], cpage=0, send_at=None,
        anon=False, deadline_h=None, deadline_abs=None, remind_min=0, rep_kind=None, rep_time="09:00",
    )
    await goto("q", event.callback.user.user_id, context)


@dp.message_created(Create.question)
async def create_question(event: MessageCreated, context: MemoryContext):
    uid = event.message.sender.user_id
    text = (event.message.body.text or "").strip() if event.message.body else ""
    if not text:
        await send(uid, "Нужен текст вопроса. Напишите его сообщением.", screen="q")
        return
    await context.update_data(pending_q=text)
    await goto("mode", uid, context)


@dp.message_callback(F.callback.payload.startswith("mode:"))
async def cb_mode(event: MessageCallback, context: MemoryContext):
    uid = event.callback.user.user_id
    await ack(event)
    data = await context.get_data()
    if not data.get("pending_q"):  # кнопка из старого сообщения
        return
    if payload_of(event).split(":")[1] == "test":
        await goto("opts", uid, context)
    else:  # свободный ответ
        questions = data.get("questions", [])
        questions.append({"text": data["pending_q"], "options": [], "correct": None})
        await context.update_data(questions=questions, pending_q=None)
        await goto("more", uid, context)


@dp.message_created(Create.options)
async def create_options(event: MessageCreated, context: MemoryContext):
    uid = event.message.sender.user_id
    options: list[str] = []
    text = (event.message.body.text or "") if event.message.body else ""
    for line in text.split("\n"):
        line = line.strip()
        if line and line not in options:
            options.append(line)
    if len(options) < 2:
        await send(uid, "Нужно минимум 2 разных варианта — каждый с новой строки. Попробуйте ещё раз.", screen="opts")
        return
    if len(options) > 20:
        await send(uid, "Слишком много вариантов — максимум 20. Сократите список и отправьте снова.", screen="opts")
        return
    data = await context.get_data()
    if not data.get("pending_q"):
        return
    await context.update_data(pending_opts=options)
    await goto("correct", uid, context)


@dp.message_callback(F.callback.payload.startswith("cor:"))
async def cb_correct(event: MessageCallback, context: MemoryContext):
    uid = event.callback.user.user_id
    await ack(event)
    data = await context.get_data()
    opts = data.get("pending_opts")
    if not data.get("pending_q") or not opts:  # кнопка из старого сообщения
        return
    v = payload_of(event).split(":")[1]
    correct = None if v == "-" else int(v)
    if correct is not None and not (0 <= correct < len(opts)):
        correct = None
    questions = data.get("questions", [])
    questions.append({"text": data["pending_q"], "options": opts, "correct": correct})
    await context.update_data(questions=questions, pending_q=None, pending_opts=None)
    await goto("more", uid, context)


@dp.message_callback(F.callback.payload == "addq")
async def cb_addq(event: MessageCallback, context: MemoryContext):
    await ack(event)
    await context.update_data(ql_back=False)
    await goto("q", event.callback.user.user_id, context)


# ---- изменение и удаление вопросов в мастере ----
@dp.message_callback(F.callback.payload == "qlist")
async def cb_qlist(event: MessageCallback, context: MemoryContext):
    await ack(event)
    await context.update_data(ql_back=False)
    await goto("qlist", event.callback.user.user_id, context)


@dp.message_callback(F.callback.payload.startswith("qe:"))
async def cb_qedit(event: MessageCallback, context: MemoryContext):
    await context.set_state(None)
    i = int(payload_of(event).split(":")[1])
    data = await context.get_data()
    if not (0 <= i < len(data.get("questions", []))):
        await notify(event, "Вопрос не найден")
        return
    text, rows = qedit_screen(data, i)
    await refresh(event, text, rows, screen="qedit")


@dp.message_callback(F.callback.payload.startswith("qt:"))
async def cb_qtext(event: MessageCallback, context: MemoryContext):
    await ack(event)
    i = int(payload_of(event).split(":")[1])
    await context.set_state(Create.edit_q)
    await context.update_data(edit_idx=i)
    await send(event.callback.user.user_id, "✍️ Введите новый текст вопроса:", [[("Отмена", f"qe:{i}")]])


@dp.message_created(Create.edit_q)
async def edit_q_text(event: MessageCreated, context: MemoryContext):
    uid = event.message.sender.user_id
    text = ((event.message.body.text or "") if event.message.body else "").strip()
    if not text:
        await send(uid, "Нужен текст вопроса. Напишите его сообщением.")
        return
    data = await context.get_data()
    i, qs = data.get("edit_idx"), data.get("questions", [])
    await context.set_state(None)
    if i is not None and 0 <= i < len(qs):
        qs[i]["text"] = text
        await context.update_data(questions=qs)
        data = await context.get_data()
        t, rows = qedit_screen(data, i)
        await send(uid, "✅ Текст изменён.\n\n" + t, rows, screen="qedit")
    else:
        await goto("qlist", uid, context)


@dp.message_callback(F.callback.payload.startswith("qu:"))
async def cb_qup(event: MessageCallback, context: MemoryContext):
    i = int(payload_of(event).split(":")[1])
    data = await context.get_data()
    qs = data.get("questions", [])
    if 0 < i < len(qs):
        qs[i - 1], qs[i] = qs[i], qs[i - 1]
        await context.update_data(questions=qs)
    data = await context.get_data()
    text, rows = qlist_screen(data)
    await refresh(event, text, rows, screen="qlist")


@dp.message_callback(F.callback.payload.startswith("qd:"))
async def cb_qdel(event: MessageCallback, context: MemoryContext):
    uid = event.callback.user.user_id
    i = int(payload_of(event).split(":")[1])
    data = await context.get_data()
    qs = data.get("questions", [])
    if 0 <= i < len(qs):
        qs.pop(i)
        await context.update_data(questions=qs)
    if not qs:
        await ack(event)
        await goto("q", uid, context)
        return
    data = await context.get_data()
    text, rows = qlist_screen(data)
    await refresh(event, text, rows, screen="qlist")


# ---- куда опубликовать ----
@dp.message_callback(F.callback.payload == "dest")
async def cb_dest(event: MessageCallback, context: MemoryContext):
    await ack(event)
    await goto("dest", event.callback.user.user_id, context)


async def dest_refresh(event: MessageCallback, context: MemoryContext):
    await save_draft(event.callback.user.user_id, context)
    data = await context.get_data()
    text, rows = dest_screen(event.callback.user.user_id, data)
    await refresh(event, text, rows, screen="dest")


@dp.message_callback(F.callback.payload.startswith("cp:"))
async def cb_dest_page(event: MessageCallback, context: MemoryContext):
    await context.update_data(cpage=int(payload_of(event).split(":")[1]))
    await dest_refresh(event, context)


@dp.message_callback(F.callback.payload.startswith("ct:"))
async def cb_chat_toggle(event: MessageCallback, context: MemoryContext):
    target = int(payload_of(event).split(":")[1])
    if target not in [c["chat_id"] for c in my_chats(event.callback.user.user_id)]:
        await notify(event, "Этот чат недоступен")
        return
    data = await context.get_data()
    csel = data.get("chat_sel", [])
    csel = [x for x in csel if x != target] if target in csel else csel + [target]
    await context.update_data(chat_sel=csel)
    await dest_refresh(event, context)


@dp.message_callback(F.callback.payload == "dnext")
async def cb_dest_next(event: MessageCallback, context: MemoryContext):
    data = await context.get_data()
    if data.get("chat_sel"):
        await refresh(event, "🕒 Когда опубликовать в чате?", WHEN_ROWS, screen="when")
    else:  # только ссылка — публиковать по расписанию незачем
        await context.update_data(send_at=None)
        await save_draft(event.callback.user.user_id, context)
        data = await context.get_data()
        await refresh(event, EXTRAS_TEXT, extras_rows(data), screen="extras")


# ---- время, дополнительные настройки, подтверждение ----
def confirm_text(data: dict) -> str:
    send_at = send_at_of(data)
    lines = []
    for i, q in enumerate(data["questions"], 1):
        if q["options"]:
            opts = ", ".join(("✔" + o) if q.get("correct") == j else o for j, o in enumerate(q["options"]))
            kind = "тест: " + opts
        else:
            kind = "свободный ответ"
        lines.append(f"{i}. {q['text']}\n    ({kind})")
    extra = []
    chats = data.get("chat_sel") or []
    if chats:
        when = "сразу" if send_at is None else send_at.astimezone(TZ).strftime("%d.%m.%Y в %H:%M")
        extra.append(f"💬 Чатов: {len(chats)}, публикация: {when}")
    else:
        extra.append("🔗 Только по ссылке (в чат не публикуем)")
    if data.get("anon"):
        extra.append("🕶 Анонимный опрос")
    dl = compute_deadline(data, data.get("send_at"))
    if dl:
        extra.append(f"⏰ Срок: {fmt_local(dl)}")
    rm = data.get("remind_min") or 0
    if rm:
        extra.append(f"🔔 Напоминания сами: раз в {fmt_minutes(rm)}, пока человек не ответит")
    else:
        extra.append("🔔 Напоминаний нет")
    if data.get("rep_kind") and chats:
        extra.append(f"🔁 Повтор: {rep_label(data['rep_kind'], data.get('rep_time') or '09:00')}; итоги каждого запуска — отдельным файлом")
    quiz = sum(1 for q in data["questions"] if q.get("correct") is not None)
    if quiz:
        extra.append(f"🎯 Оценка: правильные ответы заданы в вопросах: {quiz}")
    return "📋 Проверьте опрос:\n" + "\n".join(lines) + "\n\n" + "\n".join(extra)


@dp.message_callback(F.callback.payload.startswith("when:"))
async def cb_when(event: MessageCallback, context: MemoryContext):
    kind = payload_of(event).split(":")[1]
    uid = event.callback.user.user_id
    if kind == "custom":
        await ack(event)
        await context.set_state(Create.custom_time)
        await send(
            uid,
            f"Введите дату и время в формате ДД.ММ.ГГГГ ЧЧ:ММ\nЧасовой пояс: {TZ_NAME}\nПример: {example_dt()}",
            screen="custom",
        )
        return
    send_at: datetime | None
    if kind == "now":
        send_at = None
    elif kind == "1h":
        send_at = now_utc() + timedelta(hours=1)
    else:
        t = datetime.now(TZ) + timedelta(days=1)
        send_at = t.replace(hour=9, minute=0, second=0, microsecond=0).astimezone(timezone.utc)
    await context.update_data(send_at=send_at.isoformat() if send_at else None)
    await save_draft(uid, context)
    data = await context.get_data()
    await refresh(event, EXTRAS_TEXT, extras_rows(data), screen="extras")


@dp.message_created(Create.custom_time)
async def create_custom_time(event: MessageCreated, context: MemoryContext):
    uid = event.message.sender.user_id
    raw = ((event.message.body.text or "") if event.message.body else "").strip()
    local = parse_local_dt(raw)
    if local is None:
        await send(uid, f"Не получилось разобрать дату 😕 Формат: ДД.ММ.ГГГГ ЧЧ:ММ, например {example_dt()}", screen="custom")
        return
    if local <= datetime.now(TZ):
        await send(uid, "Это время уже прошло. Введите время в будущем.", screen="custom")
        return
    await context.update_data(send_at=local.astimezone(timezone.utc).isoformat())
    await context.set_state(None)
    await save_draft(uid, context)
    data = await context.get_data()
    await send(uid, EXTRAS_TEXT, extras_rows(data), screen="extras")


@dp.message_callback(F.callback.payload.startswith("ex:"))
async def cb_extras(event: MessageCallback, context: MemoryContext):
    uid = event.callback.user.user_id
    parts = payload_of(event).split(":")
    data = await context.get_data()
    if not data.get("questions") and not data.get("reopen_id"):
        await stale_session(event, context, "extras")
        return
    act = parts[1]
    if act == "anon":
        await context.update_data(anon=not data.get("anon"))
    elif act == "dl":
        v = parts[2]
        if v == "custom":
            await ack(event)
            await context.set_state(Create.deadline)
            await send(
                uid,
                f"Введите срок в формате ДД.ММ.ГГГГ ЧЧ:ММ\nЧасовой пояс: {TZ_NAME}\nПример: {example_dt(24)}",
                screen="dl_custom",
            )
            return
        if v == "0":
            await context.update_data(deadline_h=None, deadline_abs=None)
        else:
            await context.update_data(deadline_h=int(v), deadline_abs=None)
    elif act == "rm":
        v = parts[2]
        if v == "custom":
            await ack(event)
            await context.set_state(Create.remind_custom)
            await send(
                uid,
                "🔔 Как часто напоминать? Напишите, например:\n"
                "10 мин — каждые 10 минут\n2 ч — каждые 2 часа\n1 д — раз в сутки\n\n"
                f"Бот будет повторять напоминание, пока человек не ответит или опрос не закроется. "
                f"Минимум {REMIND_MIN_LIMIT} минут, максимум 7 дней.",
                screen="rm_custom",
            )
            return
        await context.update_data(remind_min=int(v))
    elif act == "rep":
        text, rows = rep_screen(data)
        await refresh(event, text, rows, screen="rep")
        return
    elif act == "next":
        if data.get("reopen_id"):
            await refresh(event, reopen_text(data), reopen_rows(data), screen="reopen")
            return
        await refresh(event, confirm_text(data), CONFIRM_ROWS, screen="confirm")
        return
    await save_draft(uid, context)
    data = await context.get_data()
    await refresh(event, EXTRAS_TEXT, extras_rows(data), screen="extras")


@dp.message_callback(F.callback.payload.startswith("rep:"))
async def cb_rep(event: MessageCallback, context: MemoryContext):
    uid = event.callback.user.user_id
    data = await context.get_data()
    if not data.get("questions"):
        await stale_session(event, context, "extras")
        return
    parts = payload_of(event).split(":")
    act = parts[1]
    if act == "done":
        await save_draft(uid, context)
        await refresh(event, EXTRAS_TEXT, extras_rows(data), screen="extras")
        return
    if act == "time":
        await ack(event)
        await context.set_state(Create.rep_time)
        await send(uid, "🕘 Во сколько запускать опрос? Напишите время, например 09:00 (часовой пояс: " + TZ_NAME + ").",
                   screen="rep_time")
        return
    if act == "t":
        v = parts[2]
        await context.update_data(rep_time=f"{v[:2]}:{v[2:]}")
    elif act == "none":
        await context.update_data(rep_kind=None)
    elif act in ("daily", "weekdays", "weekly"):
        await context.update_data(rep_kind=act)
    await save_draft(uid, context)
    data = await context.get_data()
    text, rows = rep_screen(data)
    await refresh(event, text, rows, screen="rep")


@dp.message_created(Create.rep_time)
async def create_rep_time(event: MessageCreated, context: MemoryContext):
    uid = event.message.sender.user_id
    raw = ((event.message.body.text or "") if event.message.body else "").strip()
    t = parse_hhmm(raw)
    if t is None:
        await send(uid, "Не получилось понять 😕 Напишите время, например 09:00 или 18:30.", screen="rep_time")
        return
    await context.update_data(rep_time=t)
    await context.set_state(None)
    await save_draft(uid, context)
    data = await context.get_data()
    text, rows = rep_screen(data)
    await send(uid, text, rows, screen="rep")


@dp.message_created(Create.remind_custom)
async def create_remind_custom(event: MessageCreated, context: MemoryContext):
    uid = event.message.sender.user_id
    raw = ((event.message.body.text or "") if event.message.body else "").strip()
    m = parse_interval(raw)
    if m is None:
        await send(uid, "Не получилось понять 😕 Напишите, например: 10 мин, 2 ч или 1 д.", screen="rm_custom")
        return
    if m < REMIND_MIN_LIMIT:
        await send(uid, f"Слишком часто. Минимальный интервал — {REMIND_MIN_LIMIT} минут.", screen="rm_custom")
        return
    if m > REMIND_MAX_LIMIT:
        await send(uid, "Слишком редко. Максимальный интервал — 7 дней.", screen="rm_custom")
        return
    await context.update_data(remind_min=m)
    await context.set_state(None)
    await save_draft(uid, context)
    data = await context.get_data()
    await send(uid, EXTRAS_TEXT, extras_rows(data), screen="extras")


@dp.message_created(Create.deadline)
async def create_deadline(event: MessageCreated, context: MemoryContext):
    uid = event.message.sender.user_id
    raw = ((event.message.body.text or "") if event.message.body else "").strip()
    local = parse_local_dt(raw)
    if local is None:
        await send(uid, f"Не получилось разобрать дату 😕 Формат: ДД.ММ.ГГГГ ЧЧ:ММ, например {example_dt(24)}", screen="dl_custom")
        return
    data = await context.get_data()
    base = send_at_of(data) or now_utc()
    dl = local.astimezone(timezone.utc)
    if dl <= base:
        await send(uid, "Срок должен быть позже времени публикации. Введите другую дату.", screen="dl_custom")
        return
    await context.update_data(deadline_abs=dl.isoformat(), deadline_h=None)
    await context.set_state(None)
    await save_draft(uid, context)
    data = await context.get_data()
    await send(uid, EXTRAS_TEXT, extras_rows(data), screen="extras")


@dp.message_callback(F.callback.payload == "confirm")
async def cb_confirm(event: MessageCallback, context: MemoryContext):
    uid = event.callback.user.user_id
    data = await context.get_data()
    questions = data.get("questions")
    if not questions:
        await stale_session(event, context, "confirm")
        return
    chat_sel = data.get("chat_sel") or []
    send_at = data.get("send_at") if chat_sel else None
    deadline = compute_deadline(data, send_at)
    remind_min = int(data.get("remind_min") or 0)
    anon = 1 if data.get("anon") else 0
    rep_kind = data.get("rep_kind") if chat_sel else None  # повтор только для опросов в чатах
    rep_time = data.get("rep_time") or "09:00"
    rep_dow = rep_next = rep_dl_h = None
    if rep_kind:
        first = datetime.fromisoformat(send_at) if send_at else now_utc()
        rep_dow = first.astimezone(TZ).weekday()
        rep_next = next_run(rep_kind, rep_time, first, rep_dow).isoformat()
        if deadline:
            rep_dl_h = max(1, round((datetime.fromisoformat(deadline) - first).total_seconds() / 3600))
    # очищаем состояние ДО записи в БД — защита от двойного нажатия «Подтвердить»
    await context.clear()
    delete_draft(uid)
    cur = db.execute(
        "INSERT INTO polls(owner_id, questions, status, send_at, created, chat_ids, "
        "deadline, remind_min, anon, rep_kind, rep_time, rep_dow, rep_next, rep_dl_h) "
        "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
        (uid, json.dumps(questions, ensure_ascii=False), "scheduled",
         send_at or now_utc().isoformat(), now_utc().isoformat(),
         json.dumps(chat_sel) if chat_sel else None, deadline, remind_min, anon,
         rep_kind, rep_time if rep_kind else None, rep_dow, rep_next, rep_dl_h),
    )
    poll_id = cur.lastrowid
    if poll_id is None:
        await refresh(event, "Не удалось сохранить опрос, попробуйте ещё раз.")
        return
    db.commit()
    if send_at is None:
        await refresh(event, "🚀 Создаю опрос…")
        await dispatch_poll(poll_id)  # откроет опрос, опубликует в чатах и пришлёт вам ссылку
    else:
        await refresh(
            event,
            f"📅 Опрос #{poll_id} будет опубликован в чате {fmt_local(send_at)}. "
            "Ссылка заработает в это же время; я пришлю её сразу после публикации.",
        )
        await show_menu(uid, "Что дальше?")


@dp.message_callback(F.callback.payload == "cancel")
async def cb_cancel(event: MessageCallback, context: MemoryContext):
    await context.clear()
    delete_draft(event.callback.user.user_id)
    await refresh(event, "Создание опроса отменено.")
    await show_menu(event.callback.user.user_id)


# ---------- ПУБЛИКАЦИЯ, АВТОМАТИЧЕСКИЕ НАПОМИНАНИЯ ----------
_dispatching: set[int] = set()  # защита от параллельной публикации одного опроса


def chat_buttons(poll: dict):
    """Кнопки под опросом в чате. Один вопрос с вариантами — отвечают сразу кнопками, без запуска бота."""
    qs = poll["questions"]
    rows = []
    if len(qs) == 1 and qs[0]["options"]:
        rows = [[(o, f"qv:{poll['id']}:{i}")] for i, o in enumerate(qs[0]["options"])]
    elif bot_link(poll["id"]) and LinkButton:
        # одна кнопка-ссылка: открывает бота сразу на этом опросе, ничего искать не нужно
        rows.append([("Пройти опрос", f"url:{bot_link(poll['id'])}")])
    else:
        rows.append([("Пройти опрос", f"join:{poll['id']}")])
    if len(qs) == 1 and qs[0]["options"] and bot_link(poll["id"]):
        rows.append([("Открыть в боте", f"url:{bot_link(poll['id'])}")])
    return rows


def is_free_single(poll: dict) -> bool:
    """Опрос из одного вопроса со свободным ответом, не анонимный: на него можно ответить прямо в чате."""
    qs = poll["questions"]
    return len(qs) == 1 and not qs[0]["options"] and not poll.get("anon")


def chat_card_text(poll: dict, prefix: str = "📋 Новый опрос!") -> str:
    text = poll_card_text(poll, prefix)
    qs = poll["questions"]
    if len(qs) == 1 and qs[0]["options"]:
        text += "\n\nНажмите на вариант ниже — запускать бота не нужно."
    elif is_free_single(poll):
        text += "\n\n✍️ Ответьте на это сообщение (кнопка «Ответить»): бот засчитает Ваш ответ и подтвердит это."
    return text


async def chat_audience(chat_id: int) -> int | None:
    try:
        chat = await bot.get_chat_by_id(chat_id)
        n = getattr(chat, "participants_count", None)
        return int(n) if n else None
    except Exception:
        return None


async def chat_members(chat_id: int) -> list[dict] | None:
    """Участники чата (без ботов): [{id, name, username}]. None — список получить не удалось."""
    res, marker = [], None
    try:
        for _ in range(20):  # страницы по 100 человек
            kw = {"count": 100}
            if marker:
                kw["marker"] = marker
            data = await with_retry(bot.get_chat_members, chat_id, **kw)
            for m in getattr(data, "members", None) or []:
                if getattr(m, "is_bot", False):
                    continue
                name = " ".join(x for x in (getattr(m, "first_name", None), getattr(m, "last_name", None)) if x)
                res.append({"id": m.user_id, "name": name or str(m.user_id), "username": getattr(m, "username", None)})
            marker = getattr(data, "marker", None)
            if not marker:
                break
        return res
    except Exception as e:
        log.warning("Не удалось получить участников чата %s: %s", chat_id, e)
        return None


async def pending_text(poll: dict, chat_id: int) -> str:
    """Кто в чате ещё не прошёл опрос: участники чата минус те, кто ответил (и автор опроса)."""
    if poll.get("anon"):  # в анонимном опросе поимённый список выдал бы, кто ответил
        return ""
    members = await chat_members(chat_id)
    if members is None:
        return ""
    done = {r["user_id"] for r in db.execute(
        "SELECT user_id FROM recipients WHERE poll_id=? AND done=1", (poll["id"],))}
    left = [m for m in members if m["id"] not in done and m["id"] != poll["owner_id"]]
    if not left:
        return "\n\n✅ В этом чате все уже ответили."
    shown = ["@" + m["username"] if m["username"] else m["name"] for m in left[:30]]
    more = f" и ещё {len(left) - 30}" if len(left) > 30 else ""
    return f"\n\n⏳ Ещё не прошли ({len(left)}): " + ", ".join(shown) + more


async def post_to_chats(poll: dict, prefix: str = "📋 Новый опрос!", measure: bool = True, pending: bool = False):
    """Публикует опрос в чатах. Возвращает (успешно, ошибки, всего).
    pending=True — в напоминании добавляется список тех, кто ещё не прошёл."""
    ids = json.loads(poll["chat_ids"]) if poll.get("chat_ids") else []
    ok, fails, audience, known = 0, [], 0, True
    for cid in ids:
        try:
            res = await with_retry(
                bot.send_message, chat_id=cid,
                text=chat_card_text(poll, prefix) + (await pending_text(poll, cid) if pending else ""),
                attachments=[kb_of(chat_buttons(poll))],
            )
            ok += 1
            if is_free_single(poll):  # запоминаем сообщение, чтобы засчитывать ответы-реплаи на него
                try:
                    mid = res.message.body.mid
                    db.execute("INSERT OR REPLACE INTO chat_posts(mid, chat_id, poll_id) VALUES(?,?,?)", (mid, cid, poll["id"]))
                    db.commit()
                except Exception:
                    log.warning("Не удалось запомнить сообщение опроса", exc_info=True)
            if measure:
                n = await chat_audience(cid)
                if n:
                    audience += n
                else:
                    known = False
        except Exception as e:
            log.warning("Не удалось опубликовать в чате %s: %s", cid, e)
            fails.append(f"{cid}: {str(e)[:100]}")
        await asyncio.sleep(0.1)
    if measure and ids and known and audience:
        db.execute("UPDATE polls SET audience=? WHERE id=?", (audience, poll["id"]))
        db.commit()
    return ok, fails, len(ids)


async def dispatch_poll(poll_id: int):
    """Открывает опрос: публикует в чатах и присылает автору ссылку."""
    if poll_id in _dispatching:
        return
    _dispatching.add(poll_id)
    try:
        await _dispatch_poll(poll_id)
    finally:
        _dispatching.discard(poll_id)


async def _dispatch_poll(poll_id: int):
    poll = get_poll(poll_id)
    if not poll or poll["status"] != "scheduled":
        return
    # от момента публикации отсчитываем первое напоминание
    db.execute("UPDATE polls SET status='sent', last_remind=? WHERE id=?", (now_utc().isoformat(), poll_id))
    db.commit()
    # старые опросы с заранее выбранными получателями (прежняя версия бота): дошлём им карточку
    for r in db.execute("SELECT user_id FROM recipients WHERE poll_id=? AND sent=0", (poll_id,)).fetchall():
        err = await deliver(r["user_id"], poll_card_text(poll), [[("Пройти опрос", f"take:{poll_id}")]])
        if err is None:
            db.execute("UPDATE recipients SET sent=1 WHERE poll_id=? AND user_id=?", (poll_id, r["user_id"]))
            db.commit()
        await asyncio.sleep(0.1)
    if poll.get("rep_series"):  # очередной запуск серии: коротко, без ссылки
        lines = [f"🔁 По расписанию запущен опрос #{poll_id}."]
        if poll.get("chat_ids"):
            ok, fails, total = await post_to_chats(poll)
            lines.append(f"💬 Опубликовано в чатах: {ok} из {total}")
            lines += [f"• {x}" for x in fails]
        await deliver(poll["owner_id"], "\n".join(lines), [[("Открыть опрос", f"pm:{poll_id}")]])
        return
    lines = [f"✅ Опрос #{poll_id} открыт."]
    if poll.get("chat_ids"):
        ok, fails, total = await post_to_chats(poll)
        lines.append(f"💬 Опубликовано в чатах: {ok} из {total}")
        lines += [f"• {x}" for x in fails]
    if poll.get("remind_min"):
        lines.append(f"🔔 Бот сам напомнит тем, кто не ответил: раз в {fmt_minutes(poll['remind_min'])}.")
    lines.append("Когда опрос будет пройден, я пришлю сообщение — следить не нужно.")
    link = bot_link(poll_id)
    if link:
        lines.append("\nПерешлите следующее сообщение любому контакту или в любой чат — человеку достаточно нажать «Начать» у бота.")
    await deliver(poll["owner_id"], "\n".join(lines), [[("Открыть опрос", f"pm:{poll_id}")]])
    if link:
        await deliver(poll["owner_id"], poll_card_text(get_poll(poll_id), "📋 Опрос", with_link=True))


async def send_reminders(poll_id: int) -> tuple[int, int, int]:
    """Автонапоминание: личные сообщения тем, кто начал, но не закончил, и повтор опроса в чатах.
    Возвращает (личных, ошибок, чатов). Автору ничего не отправляется."""
    poll = get_poll(poll_id)
    if not poll or not poll_open(poll):
        return 0, 0, 0
    db.execute("UPDATE polls SET last_remind=? WHERE id=?", (now_utc().isoformat(), poll_id))
    db.commit()
    rows = db.execute("SELECT user_id FROM recipients WHERE poll_id=? AND done=0 AND sent=1", (poll_id,)).fetchall()
    ok = fail = 0
    for r in rows:
        err = await deliver(
            r["user_id"], poll_card_text(poll, "⏰ Напоминание: Вы не закончили опрос!"),
            [[("Продолжить", f"take:{poll_id}")]],
        )
        if err is None:
            ok += 1
        else:
            fail += 1
        await asyncio.sleep(0.1)
    chats = 0
    if poll.get("chat_ids"):
        audience = poll.get("audience")
        if not (audience and poll_done_count(poll_id) >= audience):  # все ответили — чат не беспокоим
            chats, _, _ = await post_to_chats(poll, "⏰ Напоминание: опрос ещё открыт!", measure=False, pending=True)
    return ok, fail, chats


@dp.message_callback(F.callback.payload.startswith("join:"))
async def cb_join(event: MessageCallback):
    """Кнопка «Пройти опрос» в групповом чате (для опросов из нескольких вопросов или со свободным ответом)."""
    uid = event.callback.user.user_id
    try:
        poll_id = int(payload_of(event).split(":")[1])
        register_user(uid, event.callback.user.first_name or "")
        res = join_poll(uid, poll_id)
        log.info("join: poll=%s user=%s -> %s", poll_id, uid, res)
        if res not in ("ok", "done"):
            await notify(event, JOIN_MSG.get(res, JOIN_MSG["missing"]))
            return
        poll = get_poll(poll_id)
        if res == "ok":
            err = await deliver(uid, poll_card_text(poll), [[("Пройти опрос", f"take:{poll_id}")]])
        else:
            err = await deliver(uid, "Вы уже прошли этот опрос. Пока он открыт, ответы можно изменить.",
                                [[("Изменить ответы", f"redo:{poll_id}")]])
        if err is None:
            await notify(event, "Отправил опрос Вам в личные сообщения 📩")
        else:
            log.warning("join: не удалось отправить карточку %s: %s", uid, err)
            await notify(event, "Не могу написать Вам в личку. Нажмите «Открыть в боте» и «Начать».")
    except Exception:
        log.exception("join failed")
        await notify(event, "Не получилось. Нажмите «Открыть в боте» и «Начать».")


@dp.message_callback(F.callback.payload.startswith("qv:"))
async def cb_quickvote(event: MessageCallback):
    """Быстрый ответ кнопкой прямо в чате (опрос из одного вопроса с вариантами) — бот запускать не нужно."""
    uid = event.callback.user.user_id
    try:
        _, pid, opt = payload_of(event).split(":")
        poll = get_poll(int(pid))
        if not poll:
            await notify(event, JOIN_MSG["missing"])
            return
        qs = poll["questions"]
        if len(qs) != 1 or not qs[0]["options"] or not (0 <= int(opt) < len(qs[0]["options"])):
            await notify(event, "Опрос недоступен")
            return
        if poll["status"] == "scheduled":
            await notify(event, JOIN_MSG["scheduled"])
            return
        if not poll_open(poll):
            await notify(event, JOIN_MSG["closed"])
            return
        register_user(uid, event.callback.user.first_name or "")
        answer = qs[0]["options"][int(opt)]
        changed = count("SELECT COUNT(*) FROM answers WHERE poll_id=? AND user_id=?", (poll["id"], uid)) > 0
        db.execute("INSERT OR IGNORE INTO recipients(poll_id, user_id, sent, done) VALUES(?,?,1,1)", (poll["id"], uid))
        db.execute("UPDATE recipients SET sent=1, done=1 WHERE poll_id=? AND user_id=?", (poll["id"], uid))
        db.execute("DELETE FROM answers WHERE poll_id=? AND user_id=? AND q_idx=0", (poll["id"], uid))
        db.execute(
            "INSERT INTO answers VALUES(?,?,?,?,?,?)",
            (poll["id"], uid, 0, qs[0]["text"], answer, now_utc().isoformat()),
        )
        db.commit()
        await notify(event, ("✏️ Ответ изменён: " if changed else "✅ Ответ принят: ") + answer[:100])
    except Exception:
        log.exception("quickvote failed")
        await notify(event, "Не получилось, попробуйте ещё раз.")


# ---------- ФОНОВЫЕ ЗАДАЧИ: ПУБЛИКАЦИЯ, СРОКИ, НАПОМИНАНИЯ, БЭКАП ----------
def in_series(poll: dict) -> bool:
    return bool(poll.get("rep_kind") or poll.get("rep_series"))


async def close_poll_auto(poll_id: int, reason: str):
    """Закрывает опрос сам (срок вышел или начался следующий запуск серии), пишет автору итог.
    Для опросов из серии итоги присылаются отдельным файлом."""
    poll = get_poll(poll_id)
    if not poll or poll["status"] != "sent":
        return
    db.execute("UPDATE polls SET status='closed', notified_done=1 WHERE id=?", (poll_id,))
    db.commit()
    await deliver(
        poll["owner_id"],
        f"✅ Опрос #{poll_id} пройден: {reason}.\nОтветили: {poll_done_count(poll_id)}. Посмотрите результаты.",
        [[("Результаты", f"res:{poll_id}")], [("Файл Excel/CSV", f"exp:{poll_id}")]],
    )
    if in_series(poll):
        try:
            await export_poll(poll["owner_id"], poll, quiet=True)
        except Exception:
            log.exception("Не удалось отправить итоги серии")


async def close_by_deadline(poll_id: int):
    await close_poll_auto(poll_id, "срок вышел, приём ответов закрыт")


async def spawn_repeat(head_id: int):
    """Очередной запуск серии: закрыть прошлый опрос, создать и опубликовать новый с теми же настройками."""
    head = get_poll(head_id)
    if not head or not head.get("rep_kind"):
        return
    now = now_utc()
    # сначала переносим следующий запуск — чтобы при сбое не запускать серию каждые 15 секунд
    nxt = next_run(head["rep_kind"], head["rep_time"] or "09:00", now, head.get("rep_dow"))
    db.execute("UPDATE polls SET rep_next=? WHERE id=?", (nxt.isoformat(), head_id))
    db.commit()
    for r in db.execute("SELECT id FROM polls WHERE (id=? OR rep_series=?) AND status='sent'", (head_id, head_id)).fetchall():
        await close_poll_auto(r["id"], "начинается следующий запуск")
    deadline = (now + timedelta(hours=head["rep_dl_h"])).isoformat() if head.get("rep_dl_h") else None
    cur = db.execute(
        "INSERT INTO polls(owner_id, questions, status, send_at, created, chat_ids, deadline, remind_min, anon, rep_series) "
        "VALUES(?,?,?,?,?,?,?,?,?,?)",
        (head["owner_id"], json.dumps(head["questions"], ensure_ascii=False), "scheduled", now.isoformat(),
         now.isoformat(), head.get("chat_ids"), deadline, head.get("remind_min") or 0, head.get("anon") or 0, head_id),
    )
    db.commit()
    await dispatch_poll(cur.lastrowid)


async def notify_all_done(poll_id: int):
    """Все участники чатов, где опубликован опрос, ответили — сообщаем автору один раз."""
    poll = get_poll(poll_id)
    if not poll:
        return
    db.execute("UPDATE polls SET notified_done=1 WHERE id=?", (poll_id,))
    db.commit()
    await deliver(
        poll["owner_id"],
        f"✅ Опрос #{poll_id} пройден: ответили все ({poll_done_count(poll_id)}). Посмотрите результаты.",
        [[("Результаты", f"res:{poll_id}")], [("Файл Excel/CSV", f"exp:{poll_id}")]],
    )


def run_backup() -> None:
    """Копия файла БД раз в сутки (в копиях тоже персональные данные — храните их так же бережно)."""
    if DB_PATH == ":memory:":
        return
    os.makedirs(BACKUP_DIR, exist_ok=True)
    dst = os.path.join(BACKUP_DIR, f"sprashivatel_{datetime.now(TZ):%Y%m%d}.db")
    out = sqlite3.connect(dst)
    try:
        db.backup(out)
    finally:
        out.close()
    for old in sorted(glob.glob(os.path.join(BACKUP_DIR, "sprashivatel_*.db")))[:-BACKUP_KEEP]:
        try:
            os.remove(old)
        except OSError:
            pass


def cleanup_old_data() -> None:
    """Автоудаление старых ответов (ANSWERS_RETENTION_DAYS в .env) и устаревшего журнала чатов прежней версии."""
    if RETENTION_DAYS > 0:
        cutoff = (now_utc() - timedelta(days=RETENTION_DAYS)).isoformat()
        n = db.execute("DELETE FROM answers WHERE ts<?", (cutoff,)).rowcount
        log.info("Автоудаление: ответов удалено %s", n)
        db.execute("DELETE FROM chat_log WHERE ts<?", (cutoff,))
    db.execute("DELETE FROM drafts WHERE updated<?", ((now_utc() - timedelta(days=14)).isoformat(),))
    db.commit()


def maintenance() -> None:
    today = datetime.now(TZ).strftime("%Y-%m-%d")
    if get_setting("last_maint") == today:
        return
    run_backup()
    cleanup_old_data()
    set_setting("last_maint", today)


async def scheduler_tick():
    # 1) публикация запланированных опросов
    due = db.execute(
        "SELECT id FROM polls WHERE status='scheduled' AND send_at<=?", (now_utc().isoformat(),)
    ).fetchall()
    for row in due:
        await dispatch_poll(row["id"])
    # 2) автозакрытие по сроку
    for row in db.execute("SELECT id, deadline FROM polls WHERE status='sent' AND deadline IS NOT NULL").fetchall():
        if datetime.fromisoformat(row["deadline"]) <= now_utc():
            await close_by_deadline(row["id"])
    # 3) опрос пройден всеми участниками чатов — одно сообщение автору
    for row in db.execute(
        "SELECT id, audience FROM polls WHERE status='sent' AND notified_done=0 AND audience>0"
    ).fetchall():
        if poll_done_count(row["id"]) >= row["audience"]:
            await notify_all_done(row["id"])
    # 4) автонапоминания с повтором: каждые remind_min минут, пока опрос открыт
    for row in db.execute(
        "SELECT id, send_at, last_remind, remind_min FROM polls WHERE status='sent' AND remind_min>0"
    ).fetchall():
        base = datetime.fromisoformat(row["last_remind"] or row["send_at"])
        if now_utc() >= base + timedelta(minutes=row["remind_min"]):
            await send_reminders(row["id"])
    # 4б) повторяющиеся опросы: пора запускать следующий
    for row in db.execute(
        "SELECT id FROM polls WHERE rep_kind IS NOT NULL AND rep_next IS NOT NULL AND rep_next<=? AND status!='scheduled'",
        (now_utc().isoformat(),),
    ).fetchall():
        await spawn_repeat(row["id"])
    # 5) ежедневное обслуживание
    maintenance()


async def scheduler_loop():
    """Каждые 15 секунд проверяет публикации, сроки, напоминания."""
    while True:
        try:
            await scheduler_tick()
        except Exception:
            log.exception("scheduler error")
        await asyncio.sleep(15)


# ---------- ПРОХОЖДЕНИЕ ОПРОСА ----------
async def ask_question(uid: int, context: MemoryContext, poll_id: int, idx: int):
    poll = get_poll(poll_id)
    if not poll or idx >= len(poll["questions"]):
        await context.clear()
        await send(uid, "Опрос не найден.")
        return
    q = poll["questions"][idx]
    await context.update_data(poll_id=poll_id, idx=idx)
    head = f"Вопрос {idx + 1}/{len(poll['questions'])}:\n{q['text']}"
    if q["options"]:  # режим теста — кнопки
        rows = [[(o, f"ans:{idx}:{i}")] for i, o in enumerate(q["options"])]
        await context.set_state(None)
        await send(uid, head, rows, screen="take")
    else:  # свободный ответ
        await context.set_state(Answer.free_text)
        await send(uid, head + "\n\nНапишите ответ сообщением.", screen="take")


async def finish_poll(uid: int, context: MemoryContext, poll: dict):
    text = "🎉 Спасибо! Ваши ответы сохранены."
    mx = quiz_max(poll)
    if mx:
        s = user_score(poll, user_answers(poll["id"], uid))
        text += f"\n🎯 Ваш результат: {s} из {mx} ({round(s * 100 / mx)}%)"
    rows = [[("Изменить ответы", f"redo:{poll['id']}")]] if poll_open(poll) else None
    await context.clear()
    await send(uid, text, rows)
    await send(uid, "Хотите создать свой опрос?", [[("Создать опрос", "new")], [("Главная", "home")]])


async def save_answer(uid: int, context: MemoryContext, text: str):
    data = await context.get_data()
    poll_id, idx = data.get("poll_id"), data.get("idx")
    if poll_id is None or idx is None:
        await context.clear()
        return
    poll = get_poll(poll_id)
    if not poll:
        await context.clear()
        await send(uid, "Опрос не найден.")
        return
    if not poll_open(poll):  # закрыт автором или прошёл срок, пока человек отвечал
        await context.clear()
        await send(uid, "🔒 Приём ответов по этому опросу закрыт. Последний ответ не сохранён.")
        return
    db.execute("DELETE FROM answers WHERE poll_id=? AND user_id=? AND q_idx=?", (poll_id, uid, idx))
    db.execute(
        "INSERT INTO answers VALUES(?,?,?,?,?,?)",
        (poll_id, uid, idx, poll["questions"][idx]["text"], text, now_utc().isoformat()),
    )
    db.commit()
    if idx + 1 < len(poll["questions"]):
        await send(uid, "✅ Ответ принят.")
        await ask_question(uid, context, poll_id, idx + 1)
    else:
        db.execute("UPDATE recipients SET done=1 WHERE poll_id=? AND user_id=?", (poll_id, uid))
        db.commit()
        await finish_poll(uid, context, poll)


@dp.message_callback(F.callback.payload.startswith("take:"))
async def cb_take(event: MessageCallback, context: MemoryContext):
    uid = event.callback.user.user_id
    poll_id = int(payload_of(event).split(":")[1])
    await ack(event)
    register_user(uid, event.callback.user.first_name or "")
    res = join_poll(uid, poll_id)
    poll = get_poll(poll_id)
    if res == "done":
        if poll and poll_open(poll):
            await send(
                uid, "Вы уже прошли этот опрос. Пока он открыт, ответы можно изменить.",
                [[("Изменить ответы", f"redo:{poll_id}")]],
            )
        else:
            await send(uid, JOIN_MSG["done"])
        return
    if res != "ok":
        await send(uid, JOIN_MSG.get(res, JOIN_MSG["missing"]))
        return
    # продолжаем с места остановки: уже данные ответы не стираем
    answered = count("SELECT COUNT(*) FROM answers WHERE poll_id=? AND user_id=?", (poll_id, uid))
    total = len(poll["questions"])
    start = answered if 0 < answered < total else 0
    if start == 0 and answered:  # на случай мусора в БД: начинаем с чистого листа
        db.execute("DELETE FROM answers WHERE poll_id=? AND user_id=?", (poll_id, uid))
        db.commit()
    await context.update_data(redo=False)
    if start:
        await send(uid, f"▶️ Продолжаем с вопроса {start + 1} из {total}.")
    await ask_question(uid, context, poll_id, start)


@dp.message_callback(F.callback.payload.startswith("redo:"))
async def cb_redo(event: MessageCallback, context: MemoryContext):
    """Изменить свои ответы, пока опрос открыт."""
    uid = event.callback.user.user_id
    poll_id = int(payload_of(event).split(":")[1])
    await ack(event)
    poll = get_poll(poll_id)
    rec = db.execute("SELECT done FROM recipients WHERE poll_id=? AND user_id=?", (poll_id, uid)).fetchone()
    if not poll or not rec or not rec["done"]:
        await send(uid, "Изменить ответы сейчас нельзя.")
    elif not poll_open(poll):
        await send(uid, "🔒 Приём ответов по этому опросу закрыт — изменить ответы уже нельзя.")
    else:
        await context.clear()
        await context.update_data(redo=True)
        await send(uid, "✏️ Отвечаем заново — прежние ответы заменятся новыми.")
        await ask_question(uid, context, poll_id, 0)


@dp.message_callback(F.callback.payload.startswith("ans:"))
async def cb_answer(event: MessageCallback, context: MemoryContext):
    uid = event.callback.user.user_id
    _, q_idx, opt = payload_of(event).split(":")
    await ack(event)
    data = await context.get_data()
    # игнорируем кнопки от старых вопросов
    if data.get("poll_id") is None or data.get("idx") != int(q_idx):
        return
    poll = get_poll(data["poll_id"])
    if not poll:
        return
    await save_answer(uid, context, poll["questions"][int(q_idx)]["options"][int(opt)])


@dp.message_created(Answer.free_text)
async def on_free_text(event: MessageCreated, context: MemoryContext):
    uid = event.message.sender.user_id
    text = ((event.message.body.text or "") if event.message.body else "").strip()
    if not text:
        await send(uid, "Пожалуйста, напишите ответ текстом.", screen="take")
        return
    await save_answer(uid, context, text)


# ---------- КАРТОЧКА ОПРОСА (только для автора) ----------
STATUS_LABEL = {"scheduled": "🕒 запланирован", "sent": "✅ принимает ответы", "closed": "🔒 приём закрыт"}


def poll_card(poll_id: int):
    poll = get_poll(poll_id)
    if not poll:
        return None
    done = poll_done_count(poll_id)
    qs = poll["questions"]
    lines = [f"{i}. {q['text'][:80]}" for i, q in enumerate(qs[:5], 1)]
    if len(qs) > 5:
        lines.append(f"…и ещё вопросов: {len(qs) - 5}")
    text = f"📋 Опрос #{poll_id} — {STATUS_LABEL.get(poll['status'], poll['status'])}\n\n" + "\n".join(lines)
    if poll.get("deadline"):
        text += f"\n\n⏰ Срок: {fmt_local(poll['deadline'])}"
    if poll.get("remind_min"):
        text += f"\n🔔 Автонапоминания: раз в {fmt_minutes(poll['remind_min'])}"
    if poll.get("anon"):
        text += "\n🕶 Анонимный"
    if quiz_max(poll):
        text += f"\n🎯 Тест с оценкой (макс. баллов: {quiz_max(poll)})"
    if poll.get("chat_ids"):
        text += f"\n💬 Опубликован в чатах: {len(json.loads(poll['chat_ids']))}"
    if poll.get("rep_kind"):
        text += f"\n🔁 Повтор: {rep_label(poll['rep_kind'], poll['rep_time'], poll.get('rep_dow'))}"
        if poll.get("rep_next"):
            text += f"; следующий запуск {fmt_local(poll['rep_next'])}"
    elif poll.get("rep_series"):
        text += f"\n🔁 Запуск серии #{poll['rep_series']}"
    text += f"\n📊 Ответили: {done}"
    rows = []
    if poll["status"] == "scheduled":
        rows.append([("Копировать", f"copy:{poll_id}"), ("Удалить", f"del:{poll_id}")])
        if poll.get("rep_kind"):
            rows.append([("Остановить повтор", f"repstop:{poll_id}")])
        return text, rows
    toggle = ("Закрыть приём", f"close:{poll_id}") if poll["status"] == "sent" else ("Открыть снова", f"reopen:{poll_id}")
    rows.append([("Результаты", f"res:{poll_id}"), ("Файл Excel/CSV", f"exp:{poll_id}")])
    rows.append([("Ссылка на опрос", f"lnk:{poll_id}"), toggle])
    rows.append([("Копировать", f"copy:{poll_id}"), ("Удалить", f"del:{poll_id}")])
    if poll.get("rep_kind"):
        rows.append([("Все запуски серии", f"sr:{poll_id}:0"), ("Остановить повтор", f"repstop:{poll_id}")])
    elif poll.get("rep_series"):
        rows.append([("Все запуски серии", f"sr:{poll['rep_series']}:0")])
    return text, rows


def bar(n: int, total: int, width: int = 10) -> str:
    k = round(width * n / total) if total else 0
    return "▓" * k + "░" * (width - k)


def results_text(poll_id: int) -> str:
    poll = get_poll(poll_id)
    done = poll_done_count(poll_id)
    anon = bool(poll.get("anon"))
    audience = poll.get("audience")
    head = f"📊 Результаты опроса #{poll_id} — {STATUS_LABEL.get(poll['status'], poll['status'])}\n"
    if audience and audience >= done:
        head += f"Ответили: {done} из ~{audience} участников чатов ({round(done * 100 / audience)}%)"
    else:
        head += f"Ответили: {done}"
    if poll.get("deadline"):
        head += f"\n⏰ Срок: {fmt_local(poll['deadline'])}"
    if anon:
        head += "\n🕶 Анонимный опрос — имена не показываются"
    parts = [head]
    for i, q in enumerate(poll["questions"]):
        rows = db.execute(
            "SELECT answer, COUNT(*) c FROM answers WHERE poll_id=? AND q_idx=? GROUP BY answer", (poll_id, i)
        ).fetchall()
        n = sum(r["c"] for r in rows)
        lines = [f"{i + 1}. {q['text']}"]
        if q["options"]:
            cnt = {o: 0 for o in q["options"]}
            for r in rows:
                cnt[r["answer"]] = cnt.get(r["answer"], 0) + r["c"]
            correct = q["options"][q["correct"]] if q.get("correct") is not None else None
            for o, c in sorted(cnt.items(), key=lambda x: -x[1]):
                p = round(c * 100 / n) if n else 0
                mark = "✔ " if correct is not None and o == correct else ""
                lines.append(f"   {mark}{o}: {c} ({p}%) {bar(c, n)}")
        else:
            lines.append(f"   Свободных ответов: {n}")
            last = db.execute(
                "SELECT u.name, a.answer FROM answers a LEFT JOIN users u ON u.user_id=a.user_id "
                "WHERE a.poll_id=? AND a.q_idx=? ORDER BY a.ts DESC LIMIT 5",
                (poll_id, i),
            ).fetchall()
            for r in last:
                who = "" if anon else f"{r['name'] or '—'}: "
                lines.append(f"   • {who}{r['answer'][:80]}")
            if n > 5:
                lines.append("   …остальные — в файле Excel")
        parts.append("\n".join(lines))
    mx = quiz_max(poll)
    if mx:
        scores = []
        for r in db.execute(
            "SELECT r.user_id, u.name FROM recipients r LEFT JOIN users u ON u.user_id=r.user_id "
            "WHERE r.poll_id=? AND r.done=1", (poll_id,)
        ).fetchall():
            scores.append((user_score(poll, user_answers(poll_id, r["user_id"])), r["name"] or str(r["user_id"])))
        if scores:
            avg = sum(s for s, _ in scores) / len(scores)
            block = [f"🎯 Тест: средний результат {avg:.1f} из {mx} ({round(avg * 100 / mx)}%)"]
            if not anon:
                block.append("Рейтинг:")
                ranked = sorted(scores, key=lambda x: (-x[0], x[1]))[:10]
                for place, (s, name) in enumerate(ranked, 1):
                    block.append(f"   {place}. {name} — {s}/{mx} ({round(s * 100 / mx)}%)")
            parts.append("\n".join(block))
    unfinished = db.execute(
        "SELECT r.user_id, u.name FROM recipients r LEFT JOIN users u ON u.user_id=r.user_id "
        "WHERE r.poll_id=? AND r.done=0 AND EXISTS("
        "SELECT 1 FROM answers a WHERE a.poll_id=r.poll_id AND a.user_id=r.user_id) ORDER BY u.name",
        (poll_id,),
    ).fetchall()
    if unfinished:
        if anon:
            parts.append(f"⏳ Начали, но не закончили: {len(unfinished)}")
        else:
            names = ", ".join(w["name"] or str(w["user_id"]) for w in unfinished[:15])
            if len(unfinished) > 15:
                names += f" и ещё {len(unfinished) - 15}"
            parts.append(f"⏳ Начали, но не закончили ({len(unfinished)}): {names}")
    return "\n\n".join(parts)[:3800]


async def own_poll(event: MessageCallback, poll_id: int):
    """Опрос виден и доступен только автору. Чужой опрос — как будто его нет."""
    poll = get_poll(poll_id)
    if not poll or poll["owner_id"] != event.callback.user.user_id:
        await send(event.callback.user.user_id, "Опрос не найден.")
        return None
    return poll


WEEKDAYS = ["пн", "вт", "ср", "чт", "пт", "сб", "вс"]


@dp.message_callback(F.callback.payload.startswith("sr:"))
async def cb_series(event: MessageCallback, context: MemoryContext):
    """Серия повторяющихся опросов: список запусков по датам."""
    await context.set_state(None)
    uid = event.callback.user.user_id
    parts = payload_of(event).split(":")
    head_id, page = int(parts[1]), int(parts[2]) if len(parts) > 2 else 0
    head = get_poll(head_id)
    if not head or head["owner_id"] != uid:
        await send(uid, "Опрос не найден.")
        return
    runs = series_runs(head_id)
    total_answers = sum(poll_done_count(r["id"]) for r in runs)
    text = f"🔁 Серия: {head['questions'][0]['text'][:80]}\n"
    if head.get("rep_kind"):
        text += f"Повтор: {rep_label(head['rep_kind'], head['rep_time'], head.get('rep_dow'))}\n"
        if head.get("rep_next"):
            text += f"Следующий запуск: {fmt_local(head['rep_next'])}\n"
    else:
        text += "Повтор остановлен.\n"
    text += f"Запусков: {len(runs)}, ответов всего: {total_answers}\n\nВыберите запуск по дате:"
    items, page, pages = page_slice(runs, page)
    rows = []
    for r in items:
        d = datetime.fromisoformat(r["send_at"] or r["created"]).astimezone(TZ)
        mark = {"scheduled": "ждёт · ", "closed": ""}.get(r["status"], "идёт · ")
        rows.append([(f"{mark}{d:%d.%m} {WEEKDAYS[d.weekday()]} {d:%H:%M} — ответили {poll_done_count(r['id'])}", f"pm:{r['id']}")])
    if pages > 1:
        rows.append(pager_row(f"sr:{head_id}", page, pages))
    rows.append([("Настройки серии", f"pm:{head_id}")])
    if head.get("rep_kind"):
        rows.append([("Остановить повтор", f"repstop:{head_id}")])
    await refresh(event, text, rows, screen="series")


@dp.message_callback(F.callback.payload.startswith("pm:"))
async def cb_poll_menu(event: MessageCallback, context: MemoryContext):
    await context.set_state(None)
    poll_id = int(payload_of(event).split(":")[1])
    if not await own_poll(event, poll_id):
        return
    text, rows = poll_card(poll_id)
    await refresh(event, text, rows, screen="poll")


@dp.message_callback(F.callback.payload.startswith("res:"))
async def cb_results(event: MessageCallback):
    poll_id = int(payload_of(event).split(":")[1])
    if not await own_poll(event, poll_id):
        return
    rows = [[("Обновить", f"res:{poll_id}"), ("Файл", f"exp:{poll_id}")], [("К опросу", f"pm:{poll_id}")]]
    await refresh(event, results_text(poll_id), rows, screen="poll")


@dp.message_callback(F.callback.payload.startswith("repstop:"))
async def cb_repstop(event: MessageCallback):
    poll_id = int(payload_of(event).split(":")[1])
    if not await own_poll(event, poll_id):
        return
    db.execute("UPDATE polls SET rep_kind=NULL, rep_next=NULL WHERE id=?", (poll_id,))
    db.commit()
    text, rows = poll_card(poll_id)
    await refresh(event, "⏹ Повтор остановлен. Текущий опрос продолжает работать.\n\n" + text, rows, screen="poll")


@dp.message_callback(F.callback.payload.startswith("lnk:"))
async def cb_link(event: MessageCallback):
    await ack(event)
    uid = event.callback.user.user_id
    poll_id = int(payload_of(event).split(":")[1])
    poll = await own_poll(event, poll_id)
    if not poll:
        return
    if not bot_link(poll_id):
        await send(uid, "Не знаю имя бота для ссылки. Укажите BOT_USERNAME в файле .env.")
        return
    await send(uid, "Перешлите это сообщение любому контакту или в любой чат 👇")
    await send(uid, poll_card_text(poll, "📋 Опрос", with_link=True))


@dp.message_callback(F.callback.payload.startswith("close:"))
async def cb_close(event: MessageCallback):
    poll_id = int(payload_of(event).split(":")[1])
    poll = await own_poll(event, poll_id)
    if not poll:
        return
    if poll["status"] == "sent":
        db.execute("UPDATE polls SET status='closed' WHERE id=?", (poll_id,))
        db.commit()
    text, rows = poll_card(poll_id)
    await refresh(event, "🔒 Приём ответов закрыт.\n\n" + text, rows, screen="poll")


def reopen_settings(poll: dict) -> dict:
    """Прежние настройки опроса для повторного открытия: срок — тот же период, но от текущего момента."""
    dh = None
    if poll.get("deadline"):
        start = datetime.fromisoformat(poll.get("send_at") or poll["created"])
        dh = max(1, round((datetime.fromisoformat(poll["deadline"]) - start).total_seconds() / 3600))
    chats = json.loads(poll["chat_ids"]) if poll.get("chat_ids") else []
    return dict(
        reopen_id=poll["id"], anon=bool(poll.get("anon")), deadline_h=dh, deadline_abs=None,
        remind_min=int(poll.get("remind_min") or 0), chat_sel=chats, repost=bool(chats), send_at=None,
    )


def reopen_text(data: dict) -> str:
    lines = [f"🔓 Открыть опрос #{data['reopen_id']} снова с такими настройками:"]
    dh = data.get("deadline_h")
    dabs = data.get("deadline_abs")
    if dabs:
        lines.append(f"⏰ Срок: {fmt_local(dabs)}")
    elif dh:
        new_dl = (now_utc() + timedelta(hours=dh)).isoformat()
        lines.append(f"⏰ Срок: {fmt_minutes(dh * 60)} от сейчас (до {fmt_local(new_dl)})")
    else:
        lines.append("⏰ Без срока")
    rm = data.get("remind_min") or 0
    lines.append(f"🔔 Напоминания: раз в {fmt_minutes(rm)}" if rm else "🔔 Напоминаний нет")
    if data.get("anon"):
        lines.append("🕶 Анонимный")
    if data.get("chat_sel"):
        lines.append(
            f"💬 Чатов: {len(data['chat_sel'])}; опубликовать повторно: "
            + ("да" if data.get("repost") else "нет")
        )
    lines.append("\nПрежние настройки сохранены, срок отсчитывается заново — от сегодняшней даты.")
    return "\n".join(lines)


def reopen_rows(data: dict):
    rows = [[("Открыть с этими настройками", "rp:ok")]]
    if data.get("chat_sel"):
        rows.append([("Повторно в чатах: " + ("да" if data.get("repost") else "нет"), "rp:chats")])
    rows.append([("Изменить срок и напоминания", "rp:edit")])
    return rows


@dp.message_callback(F.callback.payload.startswith("reopen:"))
async def cb_reopen(event: MessageCallback, context: MemoryContext):
    poll_id = int(payload_of(event).split(":")[1])
    poll = await own_poll(event, poll_id)
    if not poll:
        return
    if poll["status"] != "closed":
        text, rows = poll_card(poll_id)
        await refresh(event, text, rows, screen="poll")
        return
    await context.clear()
    await context.update_data(**reopen_settings(poll))
    data = await context.get_data()
    await refresh(event, reopen_text(data), reopen_rows(data), screen="reopen")


@dp.message_callback(F.callback.payload.startswith("rp:"))
async def cb_reopen_step(event: MessageCallback, context: MemoryContext):
    uid = event.callback.user.user_id
    data = await context.get_data()
    poll_id = data.get("reopen_id")
    poll = get_poll(poll_id) if poll_id else None
    if not poll or poll["owner_id"] != uid:
        await notify(event, "Сессия устарела, откройте опрос из списка заново")
        return
    act = payload_of(event).split(":")[1]
    if act == "chats":
        await context.update_data(repost=not data.get("repost"))
        data = await context.get_data()
        await refresh(event, reopen_text(data), reopen_rows(data), screen="reopen")
    elif act == "edit":
        await refresh(event, EXTRAS_TEXT, extras_rows(data), screen="extras")
    elif act == "ok":
        await apply_reopen(event, context, poll, data)


async def apply_reopen(event: MessageCallback, context: MemoryContext, poll: dict, data: dict):
    poll_id = poll["id"]
    now = now_utc()
    deadline = compute_deadline(data, now.isoformat())
    db.execute(
        "UPDATE polls SET status='sent', deadline=?, remind_min=?, anon=?, last_remind=?, notified_done=0 WHERE id=?",
        (deadline, int(data.get("remind_min") or 0), 1 if data.get("anon") else 0, now.isoformat(), poll_id),
    )
    db.commit()
    repost = bool(data.get("repost") and data.get("chat_sel"))
    await context.clear()
    note = ""
    if repost:
        fresh = get_poll(poll_id)
        ok, fails, total = await post_to_chats(fresh, "📋 Опрос открыт снова!", measure=False)
        note = f"\n💬 Опубликовано в чатах: {ok} из {total}"
    text, rows = poll_card(poll_id)
    await refresh(event, "🔓 Приём ответов снова открыт." + note + "\n\n" + text, rows, screen="poll")


@dp.message_callback(F.callback.payload.startswith("copy:"))
async def cb_copy(event: MessageCallback, context: MemoryContext):
    """Копирует вопросы опроса в мастер — так можно исправить ошибку и отправить заново."""
    await ack(event)
    uid = event.callback.user.user_id
    poll_id = int(payload_of(event).split(":")[1])
    poll = await own_poll(event, poll_id)
    if not poll:
        return
    await context.clear()
    await context.update_data(
        questions=poll["questions"], chat_sel=[], cpage=0, send_at=None,
        anon=bool(poll.get("anon")), deadline_h=None, deadline_abs=None,
        remind_min=int(poll.get("remind_min") or 0),
        rep_kind=poll.get("rep_kind"), rep_time=poll.get("rep_time") or "09:00",
    )
    await send(
        uid,
        f"📄 Вопросы опроса #{poll_id} загружены в мастер.\n"
        "Любой вопрос можно изменить или удалить кнопкой «Изменить/удалить вопросы», "
        "добавить новый — «Ещё вопрос». Старый опрос можно удалить из его карточки.",
    )
    await goto("more", uid, context)


@dp.message_callback(F.callback.payload.startswith("del:"))
async def cb_delete(event: MessageCallback):
    poll_id = int(payload_of(event).split(":")[1])
    if not await own_poll(event, poll_id):
        return
    n = count("SELECT COUNT(*) FROM answers WHERE poll_id=?", (poll_id,))
    await refresh(
        event,
        f"🗑 Удалить опрос #{poll_id}?\nВместе с ним удалятся все ответы (записей: {n}). Это необратимо.",
        [[("Да, удалить", f"delok:{poll_id}")], [("Нет, оставить", f"pm:{poll_id}")]],
        screen="poll",
    )


@dp.message_callback(F.callback.payload.startswith("delok:"))
async def cb_delete_ok(event: MessageCallback, context: MemoryContext):
    uid = event.callback.user.user_id
    poll_id = int(payload_of(event).split(":")[1])
    if not await own_poll(event, poll_id):
        return
    db.execute("DELETE FROM answers WHERE poll_id=?", (poll_id,))
    db.execute("DELETE FROM recipients WHERE poll_id=?", (poll_id,))
    db.execute("DELETE FROM polls WHERE id=?", (poll_id,))
    db.commit()
    await refresh(event, f"🗑 Опрос #{poll_id} удалён.")
    await show_polls(uid, context)


@dp.message_callback(F.callback.payload.startswith("exp:"))
async def cb_export(event: MessageCallback):
    await ack(event)
    uid = event.callback.user.user_id
    poll_id = int(payload_of(event).split(":")[1])
    poll = get_poll(poll_id)
    if not poll or poll["owner_id"] != uid:  # файл получает только автор
        await send(uid, "Опрос не найден.")
        return
    await export_poll(uid, poll)


async def export_poll(uid: int, poll: dict, quiet: bool = False):
    """Отправляет автору файлы Excel и CSV с ответами. quiet — не писать «ответов нет» (автоматические итоги)."""
    poll_id = poll["id"]
    answers = db.execute(
        "SELECT a.*, u.name FROM answers a LEFT JOIN users u ON u.user_id=a.user_id "
        "WHERE poll_id=? ORDER BY a.user_id, q_idx", (poll_id,)
    ).fetchall()
    if not answers:
        if not quiet:
            await send(uid, "Ответов пока нет.")
        else:
            await send(uid, f"📭 В опросе #{poll_id} никто не ответил — файл не создан.")
        return
    questions = [q["text"] for q in poll["questions"]]
    anon = bool(poll.get("anon"))
    mx = quiz_max(poll)
    table: dict[int, dict] = {}
    for a in answers:
        row = table.setdefault(a["user_id"], {"name": a["name"], "ts": a["ts"], "ans": {}})
        row["ans"][a["q_idx"]] = a["answer"]
        row["ts"] = max(row["ts"], a["ts"])
    items = list(table.items())
    if anon:
        random.shuffle(items)  # порядок строк не должен выдавать, кто есть кто
    header = ["№"] if anon else ["Имя", "ID", "Время ответа"]
    header += questions
    if mx:
        header += ["Баллы", "Процент"]
    lines = []
    for n, (uid_, r) in enumerate(items, 1):
        base = [n] if anon else [r["name"], uid_, fmt_local(r["ts"])]
        line = base + [r["ans"].get(i, "") for i in range(len(questions))]
        if mx:
            s = user_score(poll, r["ans"])
            line += [s, round(s * 100 / mx)]
        lines.append(line)
    # сводка по вопросам с вариантами
    summary = []
    for i, q in enumerate(poll["questions"]):
        if q["options"]:
            counts: dict[str, int] = {o: 0 for o in q["options"]}
            for r in table.values():
                v = r["ans"].get(i)
                if v:
                    counts[v] = counts.get(v, 0) + 1
            summary.append((q["text"], sorted(counts.items(), key=lambda x: -x[1])))
    base_path = os.path.join(EXPORT_DIR, f"poll_{poll_id}_{datetime.now(TZ):%Y%m%d_%H%M}")
    await _write_files(base_path, "Ответы", header, lines, summary)
    await _send_files(uid, base_path, f"📤 Ответы по опросу #{poll_id}" + (" (анонимный)" if anon else ""))


async def _write_files(base, sheet_name, header, lines, summary=None):
    wb = Workbook()
    ws = wb.active
    if ws is None:
        ws = wb.create_sheet()
    ws.title = sheet_name
    ws.append([safe_cell(h) for h in header])
    for line in lines:
        ws.append([safe_cell(v) for v in line])
    for idx, col in enumerate(ws.columns, 1):
        width = max(len(str(c.value or "")) for c in col)
        ws.column_dimensions[get_column_letter(idx)].width = min(60, max(12, width + 2))
    if summary:
        ws2 = wb.create_sheet("Сводка")
        for q, counts in summary:
            ws2.append([safe_cell(q)])
            for val, n in counts:
                ws2.append(["", safe_cell(val), n])
            ws2.append([])
        ws2.column_dimensions["A"].width = 50
        ws2.column_dimensions["B"].width = 30
    wb.save(base + ".xlsx")
    with open(base + ".csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow([safe_cell(h) for h in header])
        w.writerows([[safe_cell(v) for v in line] for line in lines])


async def _send_files(uid: int, base: str, caption: str):
    """Отправляет файлы и удаляет их с диска (в них персональные данные)."""
    try:
        await with_retry(bot.send_message, user_id=uid, text=caption, attachments=[InputMedia(path=base + ".xlsx")])
        await with_retry(bot.send_message, user_id=uid, text="CSV-версия:", attachments=[InputMedia(path=base + ".csv")])
    finally:
        for ext in (".xlsx", ".csv"):
            try:
                os.remove(base + ext)
            except OSError:
                pass


# ---------- ЧАТЫ: БОТ ЗАПОМИНАЕТ ТОЛЬКО «ЧАТ + КТО В НЁМ ПИСАЛ» (текст сообщений не сохраняется) ----------
_chat_titles: dict[int, str | None] = {}  # кеш названий чатов


def touch_chat(chat_id: int, title: str | None, user_id: int | None) -> None:
    ts = now_utc().isoformat()
    db.execute(
        "INSERT INTO chat_settings(chat_id, title, last_ts) VALUES(?,?,?) "
        "ON CONFLICT(chat_id) DO UPDATE SET title=COALESCE(excluded.title, chat_settings.title), last_ts=excluded.last_ts",
        (chat_id, title, ts),
    )
    if user_id:
        db.execute("INSERT OR IGNORE INTO chat_users(chat_id, user_id) VALUES(?,?)", (chat_id, user_id))
    db.commit()


async def chat_title(chat_id: int) -> str | None:
    if chat_id not in _chat_titles:
        try:
            chat = await bot.get_chat_by_id(chat_id)
            _chat_titles[chat_id] = chat.title
        except Exception:
            _chat_titles[chat_id] = None
    return _chat_titles[chat_id]


try:
    from maxapi.types import BotAdded
except ImportError:  # в старых версиях maxapi события нет — чат запомнится по первому сообщению
    BotAdded = None

if BotAdded is not None and hasattr(dp, "bot_added"):
    @dp.bot_added()
    async def on_bot_added(event):
        try:
            cid, uid = event.chat_id, event.user.user_id
            title = await chat_title(cid)
            touch_chat(cid, title, uid)
            await deliver(uid, f"✅ Бот добавлен в чат «{title or cid}». Теперь этот чат можно выбрать при создании опроса.")
        except Exception:
            log.warning("bot_added handler failed", exc_info=True)


@dp.message_created()
async def on_any_message(event: MessageCreated):
    msg = event.message
    chat_type = str(getattr(msg.recipient, "chat_type", "")).lower()
    if "dialog" in chat_type:  # личка без активного диалога — показываем меню
        uid = msg.sender.user_id
        register_user(uid, msg.sender.first_name or "")
        save_chat_id(uid, msg.recipient.chat_id)
        await show_menu(uid)
        return
    # групповой чат: запоминаем, что бот здесь есть и что этот человек здесь пишет (чтобы он мог выбрать чат для опроса)
    chat_id = msg.recipient.chat_id
    if chat_id is None or msg.sender is None:
        return
    register_user(msg.sender.user_id, msg.sender.first_name or "")
    touch_chat(chat_id, await chat_title(chat_id), msg.sender.user_id)
    await try_chat_answer(msg, chat_id)


async def try_chat_answer(msg, chat_id: int) -> None:
    """Ответ на опрос прямо в чате: человек отвечает (reply) на сообщение бота с опросом."""
    link = getattr(msg, "link", None)
    if not link or link.type != MessageLinkType.REPLY or not link.message:
        return
    row = db.execute("SELECT poll_id FROM chat_posts WHERE mid=? AND chat_id=?", (link.message.mid, chat_id)).fetchone()
    if not row:
        return
    text = ((msg.body.text or "") if msg.body else "").strip()
    if not text:
        return
    poll = get_poll(row["poll_id"])
    uid = msg.sender.user_id
    name = msg.sender.first_name or ""

    async def reply(t: str):
        try:
            await with_retry(
                bot.send_message, chat_id=chat_id, text=t,
                link=NewMessageLink(type=MessageLinkType.REPLY, mid=msg.body.mid),
            )
        except Exception:
            log.warning("Не удалось ответить в чате", exc_info=True)

    if not poll or not is_free_single(poll):
        return
    if poll["status"] != "sent" or not poll_open(poll):
        await reply("🔒 Приём ответов по этому опросу закрыт.")
        return
    changed = count("SELECT COUNT(*) FROM answers WHERE poll_id=? AND user_id=?", (poll["id"], uid)) > 0
    db.execute("INSERT OR IGNORE INTO recipients(poll_id, user_id, sent, done) VALUES(?,?,1,1)", (poll["id"], uid))
    db.execute("UPDATE recipients SET sent=1, done=1 WHERE poll_id=? AND user_id=?", (poll["id"], uid))
    db.execute("DELETE FROM answers WHERE poll_id=? AND user_id=? AND q_idx=0", (poll["id"], uid))
    db.execute(
        "INSERT INTO answers VALUES(?,?,?,?,?,?)",
        (poll["id"], uid, 0, poll["questions"][0]["text"], text, now_utc().isoformat()),
    )
    db.commit()
    who = f"{name}, " if name else ""
    await reply(f"✅ {who}ответ " + ("изменён" if changed else "принят") + ". Спасибо!")


# ---------- ЗАПУСК ----------
async def main():
    global BOT_USERNAME
    if not TOKEN:
        raise SystemExit("Укажите MAX_BOT_TOKEN в файле .env")
    if not BOT_USERNAME:
        try:
            me = await bot.get_me()
            BOT_USERNAME = getattr(me, "username", None) or ""
        except Exception:
            log.warning("Не удалось получить имя бота. Укажите BOT_USERNAME в .env для ссылок на опросы.")
    asyncio.create_task(scheduler_loop())
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())