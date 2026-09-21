"""
Smart-lab parser v2 — написан на основе реальной структуры HTML
Источник: smart-lab.ru/bonds/extra/

Структура сайта:
- Список постов: /bonds/extra/ и /bonds/extra/page{N}/  (до ~924 страниц)
- Каждый пост:   /blog/{post_id}.php
- Комментарии:   /comments/{tid}/  (загружаются отдельным запросом)

Запуск:
    python smartlab_parser_v2.py --db smartlab_bonds.db
    python smartlab_parser_v2.py --db smartlab_bonds.db --resume
    python smartlab_parser_v2.py --db smartlab_bonds.db --pages 10  # тест на 10 стр.
"""

import requests
from bs4 import BeautifulSoup
import sqlite3
import time
import random
import logging
import argparse
import re
from datetime import datetime
from typing import Optional
from tqdm import tqdm
import cloudscraper

# ─── Logging ─────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("smartlab_parser.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)
log = logging.getLogger(__name__)

# ─── Config ───────────────────────────────────────────────────────────────────

BASE_URL    = "https://smart-lab.ru"
LIST_URL    = "https://smart-lab.ru/bonds/extra/"          # стр. 1
LIST_URL_P  = "https://smart-lab.ru/bonds/extra/page{}/"  # стр. N

DATE_FROM = datetime(2018, 1, 1)
DATE_TO   = datetime(2018, 12, 31)

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) "
                  "Chrome/122.0.0.0 Safari/537.36",
    "Accept-Language": "ru-RU,ru;q=0.9",
    "Referer": "https://smart-lab.ru/bonds/extra/",
}

# ─── DB ───────────────────────────────────────────────────────────────────────

def init_db(path: str) -> sqlite3.Connection:
    conn = sqlite3.connect(path)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    c = conn.cursor()

    c.execute("""
        CREATE TABLE IF NOT EXISTS posts (
            post_id      TEXT PRIMARY KEY,   -- числовой ID из URL /blog/1271183.php
            tid          TEXT,               -- tid из div.topic (для запроса комментов)
            url          TEXT,
            title        TEXT,
            author       TEXT,
            author_url   TEXT,
            published_at TEXT,               -- ISO datetime
            text         TEXT,               -- полный текст поста
            tags         TEXT,               -- через запятую
            special_section TEXT,            -- спецраздел (облигации и др.)
            views        INTEGER DEFAULT 0,
            votes        INTEGER DEFAULT 0,
            comments_count INTEGER DEFAULT 0,
            collected_at TEXT
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS comments (
            comment_id   TEXT PRIMARY KEY,
            post_id      TEXT,
            author       TEXT,
            author_url   TEXT,
            published_at TEXT,
            text         TEXT,
            votes        INTEGER DEFAULT 0,
            is_reply     INTEGER DEFAULT 0,  -- 1 если ответ на другой комментарий
            collected_at TEXT,
            FOREIGN KEY (post_id) REFERENCES posts(post_id)
        )
    """)

    # Прогресс по страницам списка
    c.execute("""
        CREATE TABLE IF NOT EXISTS pages_progress (
            page_num     INTEGER PRIMARY KEY,
            status       TEXT,   -- 'done', 'error', 'out_of_range'
            posts_found  INTEGER DEFAULT 0,
            processed_at TEXT
        )
    """)

    # Прогресс по отдельным постам
    c.execute("""
        CREATE TABLE IF NOT EXISTS posts_progress (
            post_id      TEXT PRIMARY KEY,
            status       TEXT,   -- 'done', 'error', 'skipped'
            processed_at TEXT
        )
    """)

    conn.commit()
    return conn


def get_processed_pages(conn) -> set:
    return {r[0] for r in conn.execute(
        "SELECT page_num FROM pages_progress WHERE status='done'"
    )}


def get_processed_posts(conn) -> set:
    return {r[0] for r in conn.execute(
        "SELECT post_id FROM posts_progress"
    )}


# ─── HTTP ─────────────────────────────────────────────────────────────────────

def fetch(url: str, session: requests.Session, retries=3) -> Optional[BeautifulSoup]:
    for attempt in range(retries):
        try:
            r = session.get(url, headers=HEADERS, timeout=20)
            if r.status_code == 404:
                return None
            if r.status_code == 429:
                log.warning("Rate limit — ждём 60 сек")
                time.sleep(60)
                continue
            if r.status_code != 200:
                log.warning(f"HTTP {r.status_code}: {url}")
                time.sleep(5)
                continue
            r.encoding = "utf-8"
            return BeautifulSoup(r.text, "html.parser")
        except Exception as e:
            log.warning(f"Попытка {attempt+1}/{retries}: {e}")
            time.sleep(5 * (attempt + 1))
    return None


# ─── Date parsing ─────────────────────────────────────────────────────────────

RU_MONTHS = {
    "января": "01", "февраля": "02", "марта": "03",
    "апреля": "04", "мая": "05", "июня": "06",
    "июля": "07", "августа": "08", "сентября": "09",
    "октября": "10", "ноября": "11", "декабря": "12",
}

def parse_date(raw: str, ref_date: Optional[str] = None) -> Optional[str]:
    """
    Парсит дату из Smart-lab форматов.
    ref_date — дата поста (ISO), нужна для resolve относительных дат.

    Форматы:
      "27 февраля 2026, 21:43"  → абсолютная
      "Вчера в 20:48"           → ref_date - 1 день
      "Сегодня в 01:15"         → ref_date (тот же день)
      "2026-02-27T21:43"        → ISO
    """
    if not raw:
        return None
    raw = raw.strip()

    # ISO формат
    m = re.search(r"(\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2})", raw)
    if m:
        return m.group(1).replace(" ", "T")

    # Относительные: "Вчера в HH:MM" / "Сегодня в HH:MM"
    m_rel = re.search(r"(Вчера|Сегодня)\s+в\s+(\d{2}:\d{2})", raw, re.IGNORECASE)
    if m_rel:
        keyword, time_ = m_rel.groups()
        # Берём дату из ref_date, либо текущую дату как fallback
        if ref_date:
            try:
                base = datetime.fromisoformat(ref_date[:10])
            except Exception:
                base = datetime.now()
        else:
            base = datetime.now()
        if "Вчера" in keyword:
            from datetime import timedelta
            base = base - timedelta(days=1)
        return f"{base.strftime('%Y-%m-%d')}T{time_}"

    # "15 марта 2022, 14:30" или "15 марта 2022"
    m = re.search(r"(\d{1,2})\s+(\w+)\s+(\d{4})(?:,?\s*(\d{2}:\d{2}))?", raw)
    if m:
        day, mon_ru, year, time_ = m.groups()
        mon = RU_MONTHS.get(mon_ru.lower())
        if mon:
            dt = f"{year}-{mon}-{day.zfill(2)}"
            return f"{dt}T{time_}" if time_ else dt

    return None


def in_range(dt_str: Optional[str]) -> bool:
    if not dt_str:
        return True
    try:
        dt = datetime.fromisoformat(dt_str[:16])
        return DATE_FROM <= dt <= DATE_TO
    except Exception:
        return True


# ─── Parse list page ──────────────────────────────────────────────────────────

def parse_list_page(soup: BeautifulSoup) -> list:
    """
    Парсит страницу /bonds/extra/ или /bonds/extra/page{N}/
    Возвращает список метаданных постов (без полного текста).

    Структура li:
    <li class="bluid_XXXXX">
      <b>N_comments <u>+votes</u></b>
      HH:MM или DD/MM
      <i><a href="/profile/USER/">AUTHOR</a></i>
      <a href="/blog/POST_ID.php" title="TITLE">TITLE</a>
    </li>
    """
    posts = []
    items = soup.find_all("li", class_=lambda c: c and any(
        "bluid_" in cls for cls in (c if isinstance(c, list) else [c])
    ))

    for item in items:
        try:
            # Ссылка на пост
            post_link = item.find("a", href=lambda h: h and "/blog/" in h)
            if not post_link:
                continue
            href = post_link["href"]
            m = re.search(r"/blog/(\d+)", href)
            if not m:
                continue
            post_id = m.group(1)
            url = BASE_URL + href if not href.startswith("http") else href
            title = post_link.get("title") or post_link.get_text(strip=True)

            # Автор
            author_tag = item.find("a", href=lambda h: h and "/profile/" in h)
            author = author_tag.get_text(strip=True) if author_tag else None
            author_url = BASE_URL + author_tag["href"] if author_tag else None

            # Комментарии и голоса из <b>N <u>+M</u></b>
            b_tag = item.find("b")
            comments_count = 0
            votes = 0
            if b_tag:
                u_tag = b_tag.find("u")
                vote_text = u_tag.get_text(strip=True) if u_tag else ""
                m_v = re.search(r"[+-]?\d+", vote_text)
                votes = int(m_v.group()) if m_v else 0
                # Текст b без u — это количество комментариев
                b_text = b_tag.get_text(separator=" ").split()[0] if b_tag else "0"
                m_c = re.search(r"\d+", b_text)
                comments_count = int(m_c.group()) if m_c else 0

            # Время (просто HH:MM или DD/MM — полную дату берём со страницы поста)
            time_text = item.get_text(separator=" ", strip=True)

            posts.append({
                "post_id": post_id,
                "url": url,
                "title": title,
                "author": author,
                "author_url": author_url,
                "comments_count": comments_count,
                "votes_list": votes,  # предварительное значение
            })

        except Exception as e:
            log.debug(f"Ошибка парсинга li: {e}")

    return posts


# ─── Parse post page ──────────────────────────────────────────────────────────

def parse_post_page(post_id: str, soup: BeautifulSoup) -> Optional[dict]:
    """
    Парсит полную страницу поста /blog/{post_id}.php

    Структура:
    <div class="topic bluid_XXXXX" tid="TID" bid="BID">
      <h1 class="title">...</h1>
      <div class="content">...</div>
      <ul class="tags"><li><a>тег</a></li>...</ul>
      <ul class="ext_tags">...</ul>
    </div>
    <div class="views-total-topic" id="tviews_TID">
      <span ...>4K</span>
    </div>
    <ul class="voting guest">
      <li class="total" title="всего проголосовало: N"><a>N</a></li>
    </ul>
    """
    topic = soup.find("div", class_="topic") or soup.find("div", attrs={"tid": True})
    if not topic:
        return None

    tid = topic.get("tid")

    # Заголовок
    h1 = topic.find("h1")
    title = h1.get_text(strip=True) if h1 else None

    # Полный текст
    content_div = topic.find("div", class_="content")
    text = content_div.get_text(separator="\n", strip=True) if content_div else None

    # Теги
    tags_ul = topic.find("ul", class_="tags")
    tags = ", ".join(a.get_text(strip=True) for a in tags_ul.find_all("a")) if tags_ul else ""

    # Спецраздел (облигации, акции и др.)
    ext_tags_ul = topic.find("ul", class_="ext_tags")
    special = ", ".join(a.get_text(strip=True) for a in ext_tags_ul.find_all("a")) if ext_tags_ul else ""

    # Просмотры
    views_div = soup.find("div", class_="views-total-topic")
    views_text = views_div.get_text(strip=True) if views_div else "0"
    # Может быть "4K" или "499"
    views = 0
    if views_text:
        m = re.search(r"([\d.]+)\s*([KkКк]?)", views_text)
        if m:
            num = float(m.group(1))
            views = int(num * 1000) if m.group(2).upper() in ("K", "К") else int(num)

    # Голоса
    voting_ul = topic.find("ul", class_="voting")
    votes = 0
    if voting_ul:
        total_li = voting_ul.find("li", class_="total")
        if total_li:
            a = total_li.find("a")
            m = re.search(r"-?\d+", a.get_text() if a else "")
            votes = int(m.group()) if m else 0

    # Автор и дата — в ul.action.blog_more
    # <li class="date">27 февраля 2026, 21:43</li>
    # <li class="author"><a href="/profile/...">Имя</a></li>
    author = None
    author_url = None
    published_at = None

    blog_more = topic.find("ul", class_="blog_more")
    if blog_more:
        date_li = blog_more.find("li", class_="date")
        if date_li:
            published_at = parse_date(date_li.get_text(strip=True))

        author_li = blog_more.find("li", class_="author")
        if author_li:
            a_tag = author_li.find("a", href=lambda h: h and "/profile/" in h)
            if a_tag:
                author = a_tag.get_text(strip=True)
                href = a_tag["href"]
                author_url = BASE_URL + href if not href.startswith("http") else href

    # Количество комментариев из div.comments_total
    comments_count = 0
    comments_total = soup.find("div", class_="comments_total")
    if comments_total:
        m_c = re.search(r"\d+", comments_total.get_text())
        comments_count = int(m_c.group()) if m_c else 0

    return {
        "post_id": post_id,
        "tid": tid,
        "title": title,
        "text": text,
        "tags": tags,
        "special_section": special,
        "views": views,
        "votes": votes,
        "author": author,
        "author_url": author_url,
        "published_at": published_at,
        "comments_count": comments_count,
    }


# ─── Parse comments ───────────────────────────────────────────────────────────

def parse_comments_from_soup(post_id: str, soup: BeautifulSoup,
                             post_date: Optional[str] = None) -> list:
    """
    Парсит комментарии прямо из HTML страницы поста.
    Комментарии НЕ загружаются отдельно — они уже есть на странице поста.

    Структура:
    <div class="comments">
      <div class="comments_total">11 комментариев</div>
      <div class="comment bluid_XXXXX" id="comment_id_CID" cid="CID">
        <div class="voting guest">
          <div class="total"><a>+2</a></div>   ← голоса (пусто если 0)
        </div>
        <div class="content" id="comment_content_id_CID">
          <div class="text">текст комментария</div>
        </div>
        <div class="info">
          <div class="author"><a href="/profile/USER/">Имя</a></div>
          <ul class="chat_wrapper">
            <li class="date">Вчера в 20:48</li>  ← относительная дата!
          </ul>
        </div>
      </div>
      <!-- comment_child = ответ на комментарий -->
      <div class="comment comment_child bluid_XXXXX" cid="CID2">...</div>
    </div>

    Даты относительные ("Вчера в HH:MM", "Сегодня в HH:MM") — resolve через дату поста.
    """
    comments_div = soup.find("div", class_="comments")
    if not comments_div:
        return []

    # Все комментарии определяем по наличию атрибута cid
    comment_blocks = comments_div.find_all("div", cid=True)
    comments = []

    for div in comment_blocks:
        try:
            comment_id = div.get("cid")
            if not comment_id:
                continue

            classes = div.get("class", [])
            is_reply = 1 if "comment_child" in classes else 0

            # Автор
            author = None
            author_url = None
            author_div = div.find("div", class_="author")
            if author_div:
                a_tag = author_div.find("a", href=lambda h: h and "/profile/" in h)
                if a_tag:
                    author = a_tag.get_text(strip=True)
                    href = a_tag["href"]
                    author_url = BASE_URL + href if not href.startswith("http") else href

            # Дата из li.date внутри div.info
            # Формат: "Вчера в 20:48" / "Сегодня в 01:15" / "27 февраля 2026, 21:43"
            published_at = None
            date_li = div.find("li", class_="date")
            if date_li:
                published_at = parse_date(date_li.get_text(strip=True), post_date)

            # Текст — div.text внутри div#comment_content_id_CID
            text = None
            content_div = div.find("div", id=f"comment_content_id_{comment_id}")
            if content_div:
                text_div = content_div.find("div", class_="text")
                text = text_div.get_text(separator="\n", strip=True) if text_div else content_div.get_text(strip=True)

            # Голоса — div.total > a (пустой если 0)
            votes = 0
            total_div = div.find("div", class_="total")
            if total_div:
                a_tag = total_div.find("a")
                if a_tag:
                    m = re.search(r"[+-]?\d+", a_tag.get_text(strip=True))
                    votes = int(m.group()) if m else 0

            comments.append({
                "comment_id": comment_id,
                "post_id": post_id,
                "author": author,
                "author_url": author_url,
                "published_at": published_at,
                "text": text,
                "votes": votes,
                "is_reply": is_reply,
                "collected_at": datetime.now().isoformat(),
            })

        except Exception as e:
            log.debug(f"Ошибка парсинга комментария cid={div.get('cid')}: {e}")

    return comments


# ─── Save helpers ─────────────────────────────────────────────────────────────

def save_post(conn, data: dict):
    conn.execute("""
        INSERT OR IGNORE INTO posts
        (post_id, tid, url, title, author, author_url, published_at,
         text, tags, special_section, views, votes, comments_count, collected_at)
        VALUES (:post_id,:tid,:url,:title,:author,:author_url,:published_at,
                :text,:tags,:special_section,:views,:votes,:comments_count,:collected_at)
    """, {**data, "collected_at": datetime.now().isoformat()})


def save_comments(conn, comments: list):
    conn.executemany("""
        INSERT OR IGNORE INTO comments
        (comment_id, post_id, author, author_url, published_at, text, votes, is_reply, collected_at)
        VALUES (:comment_id,:post_id,:author,:author_url,:published_at,:text,:votes,:is_reply,:collected_at)
    """, comments)


# ─── Main pipeline ────────────────────────────────────────────────────────────

def process_post(post_meta: dict, conn, session, delay: float,
                 done_posts: set) -> bool:
    """Загружает и сохраняет один пост с комментариями."""
    post_id = post_meta["post_id"]
    if post_id in done_posts:
        return False

    url = post_meta["url"]
    soup = fetch(url, session)
    if not soup:
        conn.execute("INSERT OR REPLACE INTO posts_progress VALUES (?,?,?)",
                     (post_id, "error", datetime.now().isoformat()))
        conn.commit()
        return False

    post_data = parse_post_page(post_id, soup)
    if not post_data:
        conn.execute("INSERT OR REPLACE INTO posts_progress VALUES (?,?,?)",
                     (post_id, "error", datetime.now().isoformat()))
        conn.commit()
        return False

    # Фильтр по дате
    if not in_range(post_data.get("published_at")):
        conn.execute("INSERT OR REPLACE INTO posts_progress VALUES (?,?,?)",
                     (post_id, "skipped", datetime.now().isoformat()))
        conn.commit()
        return False

    # Дополняем данными из списка
    post_data["url"] = url
    post_data["comments_count"] = post_meta.get("comments_count", 0)
    if not post_data.get("author"):
        post_data["author"] = post_meta.get("author")
    if not post_data.get("author_url"):
        post_data["author_url"] = post_meta.get("author_url")

    save_post(conn, post_data)

    # Комментарии — парсим прямо из уже загруженной страницы поста
    if post_data.get("comments_count", 0) > 0:
        comments = parse_comments_from_soup(post_id, soup, post_data.get("published_at"))
        if comments:
            save_comments(conn, comments)
            log.debug(f"  Пост {post_id}: {len(comments)} комментариев")

    conn.execute("INSERT OR REPLACE INTO posts_progress VALUES (?,?,?)",
                 (post_id, "done", datetime.now().isoformat()))
    conn.commit()
    done_posts.add(post_id)
    return True


def main():
    parser = argparse.ArgumentParser(description="Smart-lab /bonds/extra/ parser v2")
    parser.add_argument("--db", default="smartlab_bonds.db")
    parser.add_argument("--delay", type=float, default=1.5,
                        help="Задержка между запросами (сек)")
    parser.add_argument("--resume", action="store_true",
                        help="Пропустить уже обработанные страницы и посты")
    parser.add_argument("--pages", type=int, default=None,
                        help="Ограничить число страниц (для теста)")
    args = parser.parse_args()

    conn = init_db(args.db)
    session = cloudscraper.create_scraper(
        browser={"browser": "chrome", "platform": "windows", "mobile": False}
    )
    session.headers.update(HEADERS)

    done_pages = get_processed_pages(conn) if args.resume else set()
    done_posts = get_processed_posts(conn) if args.resume else set()

    log.info(f"Уже обработано страниц: {len(done_pages)}, постов: {len(done_posts)}")

    # Сначала определяем сколько страниц всего
    log.info("Загружаем первую страницу для определения пагинации...")
    soup_p1 = fetch(LIST_URL, session)
    if not soup_p1:
        log.error("Не удалось загрузить первую страницу!")
        return

    # Максимальный номер страницы из пагинации
    # <a class="page gradient last" href="/bonds/extra/page924/">→</a>
    max_page = 1
    pagination = soup_p1.find("div", id="pagination")
    if pagination:
        page_links = pagination.find_all("a", href=lambda h: h and "/bonds/extra/page" in h)
        for link in page_links:
            m = re.search(r"/bonds/extra/page(\d+)/", link["href"])
            if m:
                max_page = max(max_page, int(m.group(1)))

    log.info(f"Всего страниц: {max_page}")

    if args.pages:
        max_page = min(max_page, args.pages)
        log.info(f"Ограничение: {max_page} страниц")

    # Статистика
    total_posts_saved = 0
    total_out_of_range = 0

    pages = range(1, max_page + 1)

    for page_num in tqdm(pages, desc="Страницы /bonds/extra/", unit="стр"):
        if args.resume and page_num in done_pages:
            continue

        url = LIST_URL if page_num == 1 else LIST_URL_P.format(page_num)
        soup = fetch(url, session)

        if not soup:
            conn.execute("INSERT OR REPLACE INTO pages_progress VALUES (?,?,?,?)",
                         (page_num, "error", 0, datetime.now().isoformat()))
            conn.commit()
            time.sleep(args.delay)
            continue

        post_metas = parse_list_page(soup)
        page_saved = 0
        page_out = 0

        for meta in post_metas:
            saved = process_post(meta, conn, session, args.delay, done_posts)
            if saved:
                page_saved += 1
                total_posts_saved += 1
            else:
                page_out += 1

            time.sleep(args.delay + random.uniform(0, 0.7))

        conn.execute("INSERT OR REPLACE INTO pages_progress VALUES (?,?,?,?)",
                     (page_num, "done", page_saved, datetime.now().isoformat()))
        conn.commit()

        log.info(f"Стр {page_num}/{max_page}: сохранено {page_saved}, вне диапазона {page_out}")
        time.sleep(args.delay + random.uniform(0, 1.0))

    conn.close()
    log.info(f"""
    ═══════════════════════════════════════
    Парсинг завершён
    Постов сохранено:  {total_posts_saved}
    ═══════════════════════════════════════
    """)


if __name__ == "__main__":
    main()
