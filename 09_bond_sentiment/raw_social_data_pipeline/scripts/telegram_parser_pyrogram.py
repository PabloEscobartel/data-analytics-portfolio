import asyncio
import os
import sqlite3
import logging
import random
import json
from datetime import datetime, timezone, timedelta

from telethon import TelegramClient
from telethon.tl.functions.channels import GetFullChannelRequest
from telethon.tl.types import Channel
from telethon.errors import FloodWaitError, MsgIdInvalidError

# ─── CONFIG ─────────────────────────────────────────

API_ID = int(os.environ["TELEGRAM_API_ID"])
API_HASH = os.environ["TELEGRAM_API_HASH"]
SESSION  = "bond_parser"

CHANNELS = [
    "bunkerobligatcy",
    "Bonds_lab",
    "rusbondsnews",
    "russianjunkbonds",
    "if_bonds",
    "cbonds",
    "angrybonds",
    "birzhevikbondsofficial1",
    "marythebond",
    "bondovik",
    "Bondholders",
    "Erabond",
    "bondsmartlab",
    "corpbonds",
    "rus_bonds_news",
    "philippovich_bonds",
    "oblig_news",
    "silenceandmoney",
    "pro_bonds",
    "mozginvest",
    "DolgosrokInvest",
    "bonds_wizzard",
    "tonsofbonds",
]

DATE_FROM = datetime(2018, 1, 1, tzinfo=timezone.utc)
DATE_TO   = datetime(2026, 1, 1, tzinfo=timezone.utc)

DUPLICATE_THRESHOLD = 500
REQUEST_INTERVAL    = 1.3

# ─── LOGGING ────────────────────────────────────────

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(message)s")
log = logging.getLogger(__name__)
logging.getLogger("telethon").setLevel(logging.WARNING)

# ─── DB ─────────────────────────────────────────────

def init_db():
    conn = sqlite3.connect("data.db")
    conn.execute("""
        CREATE TABLE IF NOT EXISTS messages (
            id               INTEGER,
            channel          TEXT,
            text             TEXT,
            date             TEXT,
            views            INTEGER,
            reactions        INTEGER,
            reactions_detail TEXT,
            parent_id        INTEGER,
            reply_count      INTEGER DEFAULT 0,
            PRIMARY KEY (id, channel)
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS processed_posts (
            post_id INTEGER,
            channel TEXT,
            PRIMARY KEY (post_id, channel)
        )
    """)
    try:
        conn.execute("ALTER TABLE messages ADD COLUMN reply_count INTEGER DEFAULT 0")
        conn.commit()
    except sqlite3.OperationalError:
        pass
    conn.commit()
    return conn


def save(conn, rows):
    if not rows:
        return
    conn.executemany("""
        INSERT OR IGNORE INTO messages
        VALUES (:id, :channel, :text, :date, :views, :reactions,
                :reactions_detail, :parent_id, :reply_count)
    """, rows)
    conn.commit()


# ─── THROTTLE ───────────────────────────────────────

LAST = 0

async def throttle():
    global LAST
    now = asyncio.get_event_loop().time()
    wait = REQUEST_INTERVAL - (now - LAST)
    if wait > 0:
        await asyncio.sleep(wait)
    LAST = asyncio.get_event_loop().time()


# ─── HELPERS ────────────────────────────────────────

def parse_reactions(msg):
    if not msg.reactions:
        return 0, ""
    results = msg.reactions.results
    total = sum(r.count for r in results)
    detail = {}
    for r in results:
        try:
            emoticon = r.reaction.emoticon
        except AttributeError:
            emoticon = str(r.reaction)
        detail[emoticon] = r.count
    return total, json.dumps(detail, ensure_ascii=False)


def row(msg, channel, parent=None):
    if not msg.message:
        return None
    dt = msg.date.astimezone(timezone.utc)
    if not (DATE_FROM <= dt <= DATE_TO):
        return None
    if parent is None and msg.reply_to:
        parent = msg.reply_to.reply_to_msg_id
    return {
        "id":               msg.id,
        "channel":          channel,
        "text":             msg.message,
        "date":             msg.date.isoformat(),
        "views":            msg.views or 0,
        "reactions":        parse_reactions(msg)[0],
        "reactions_detail": parse_reactions(msg)[1],
        "parent_id":        parent,
        "reply_count":      msg.replies.replies if msg.replies else 0,
    }


def msg_exists(conn, msg_id, channel):
    return conn.execute(
        "SELECT 1 FROM messages WHERE id=? AND channel=?", (msg_id, channel)
    ).fetchone() is not None


# ─── СБОР СООБЩЕНИЙ ────────────────────────────────

async def collect_messages(client, conn, entity, username, offset_date, stop_date):
    posts = []
    total_new = 0
    dupes_in_a_row = 0

    async for msg in client.iter_messages(entity, offset_date=offset_date):
        if msg.date < stop_date:
            break

        r = row(msg, username)
        if not r:
            continue

        if msg_exists(conn, msg.id, username):
            dupes_in_a_row += 1
            if dupes_in_a_row >= DUPLICATE_THRESHOLD:
                log.info(f"  [{username}] {DUPLICATE_THRESHOLD} дубликатов подряд — остальное собрано")
                break
            continue

        dupes_in_a_row = 0
        posts.append(r)
        total_new += 1

        if total_new % 100 == 0:
            log.info(f"  [{username}] собрано: {total_new}")

        if len(posts) >= 200:
            save(conn, posts)
            posts = []

        if random.random() < 0.05:
            await asyncio.sleep(3)

    save(conn, posts)
    return total_new


# ─── CORE ───────────────────────────────────────────

async def scrape_channel(client, conn, username):
    log.info(f"{'─'*20} START {username} {'─'*20}")

    await throttle()
    entity = await client.get_entity(username)
    await throttle()
    await client(GetFullChannelRequest(entity))

    is_chat = not (isinstance(entity, Channel) and entity.broadcast)
    log.info(f"  [{username}] тип: {'чат' if is_chat else 'канал'}")

    # ─── STEP 1: ПОСТЫ / СООБЩЕНИЯ ──────────────────

    log.info(f"  [{username}] Сбор {DATE_FROM.date()} — {DATE_TO.date()}")

    # Основной проход: от DATE_TO вниз
    new_top = await collect_messages(
        client, conn, entity, username,
        offset_date=DATE_TO, stop_date=DATE_FROM
    )
    log.info(f"  [{username}] новых: {new_top}")

    # Дособираем старые (если DATE_FROM сдвинули назад)
    min_date_str = conn.execute(
        "SELECT MIN(date) FROM messages WHERE channel=?", (username,)
    ).fetchone()[0]

    if min_date_str:
        min_date = datetime.fromisoformat(min_date_str)
        if min_date > DATE_FROM + timedelta(days=1):
            log.info(f"  [{username}] Дособираем старые (до {min_date_str[:10]})")
            new_old = await collect_messages(
                client, conn, entity, username,
                offset_date=min_date, stop_date=DATE_FROM
            )
            log.info(f"  [{username}] старых дособрано: {new_old}")

    # ─── STEP 2: КОММЕНТАРИИ (только для каналов) ───

    if is_chat:
        log.info(f"  [{username}] чат — комментарии уже в шаге 1")
        return

    # Посты с комментариями, которые ещё не обработаны
    all_with_comments = {r[0] for r in conn.execute(
        "SELECT id FROM messages WHERE channel=? AND parent_id IS NULL AND reply_count > 0",
        (username,)
    ).fetchall()}

    processed = {r[0] for r in conn.execute(
        "SELECT post_id FROM processed_posts WHERE channel=?", (username,)
    ).fetchall()}

    remaining = list(all_with_comments - processed)

    log.info(f"  [{username}] комментарии: осталось {len(remaining)} "
             f"(обработано: {len(processed)})")

    if not remaining:
        log.info(f"  [{username}] все комментарии собраны!")
        return

    random.shuffle(remaining)
    total_comments = 0

    for i, pid in enumerate(remaining):
        try:
            await throttle()
            log.info(f"  [{username}] пост {pid} | {i+1}/{len(remaining)}")
            comments = []

            async for c in client.iter_messages(entity, reply_to=pid):
                r = row(c, username, parent=pid)
                if r:
                    comments.append(r)

            save(conn, comments)
            conn.execute(
                "INSERT OR IGNORE INTO processed_posts VALUES (?, ?)",
                (pid, username)
            )
            conn.commit()
            total_comments += len(comments)

            log.info(f"  [{username}] пост {pid}: {len(comments)} комм. "
                     f"| всего: {total_comments}")

            await asyncio.sleep(2 + random.uniform(1, 2))

            if (i + 1) % 20 == 0:
                log.info(f"  [{username}] пауза...")
                await asyncio.sleep(15 + random.uniform(5, 10))

        except MsgIdInvalidError:
            conn.execute(
                "INSERT OR IGNORE INTO processed_posts VALUES (?, ?)",
                (pid, username)
            )
            conn.commit()
        except FloodWaitError as e:
            log.warning(f"  FloodWait {e.seconds}с — ждём...")
            await asyncio.sleep(e.seconds + 5)

    log.info(f"  [{username}] DONE | комментариев: {total_comments}")


# ─── MAIN ───────────────────────────────────────────

async def main():
    conn = init_db()

    client = TelegramClient(
        SESSION, API_ID, API_HASH,
        device_model="MacBook Air",
        system_version="4.16.30-vxCUSTOM",
        app_version="1.0",
        lang_code="ru",
        system_lang_code="ru-RU",
    )
    await client.start()
    await client.get_dialogs(limit=5)
    await asyncio.sleep(2)

    for ch in CHANNELS:
        await scrape_channel(client, conn, ch)
        await asyncio.sleep(10 + random.uniform(3, 5))

    # Статистика
    print("\n" + "═"*50)
    for ch in CHANNELS:
        total = conn.execute(
            "SELECT COUNT(*) FROM messages WHERE channel=?", (ch,)
        ).fetchone()[0]
        no_parent = conn.execute(
            "SELECT COUNT(*) FROM messages WHERE channel=? AND parent_id IS NULL", (ch,)
        ).fetchone()[0]
        has_parent = conn.execute(
            "SELECT COUNT(*) FROM messages WHERE channel=? AND parent_id IS NOT NULL", (ch,)
        ).fetchone()[0]
        with_comments = conn.execute(
            "SELECT COUNT(*) FROM messages WHERE channel=? AND reply_count > 0", (ch,)
        ).fetchone()[0]
        print(f"{ch}: всего={total}, без parent={no_parent}, "
              f"с parent={has_parent}, с комментариями={with_comments}")
    print("═"*50)

if __name__ == "__main__":
    asyncio.run(main())
