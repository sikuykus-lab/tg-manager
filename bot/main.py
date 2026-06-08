# -*- coding: utf-8 -*-
"""
Telegram-бот: вехи из листа «Экспорт», цепочки зависимостей (M / N), прогноз сроков, импорт из Ctrl+V.
Зависимость: у строки A поле N (от какой вехи) совпадает с полем M (у кого зависят…) предшествующей вехи.
"""
from __future__ import annotations

import asyncio
import html
import io
import os
import re
import sqlite3
import statistics
import uuid
from collections import defaultdict, deque
from datetime import date, datetime, timedelta
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple

from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command, CommandObject
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    CallbackQuery,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    KeyboardButton,
    Message,
    ReplyKeyboardMarkup,
)

# --- конфиг ---
API_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "").strip()
if not API_TOKEN:
    raise SystemExit("Задайте TELEGRAM_BOT_TOKEN (например в /root/tg_manager/.env через systemd)")

DB_PATH = os.environ.get("MILESTONE_DB_PATH", "/root/tg_manager/data.db")

# CRM → Google: только ваши скрипты листов (cron тоже на них).
# Рабочий каталог: ILFLAT_SCRIPTS_CWD или (совместимость) COMETA_CRM_SYNC_DIR.
SCRIPTS_CWD = (
    os.environ.get("ILFLAT_SCRIPTS_CWD")
    or os.environ.get("COMETA_CRM_SYNC_DIR")
    or "/opt/scripts"
).strip()
bot = Bot(token=API_TOKEN)
dp = Dispatcher(storage=MemoryStorage())


class Form(StatesGroup):
    """Поиск: 📂 только корпус → затем inline; 📌 корпус → веха → inline."""
    waiting_search_corpus = State()
    waiting_pair_corpus = State()
    waiting_pair_milestone = State()

# --- даты Excel ---
_EXCEL_BASE = datetime(1899, 12, 30)


def parse_excel_serial(raw: str) -> Optional[date]:
    if raw is None:
        return None
    s = str(raw).strip().strip('"').replace(",", ".")
    if not s or s.upper() in ("#N/A", "N/A", "-"):
        return None
    try:
        n = float(s)
    except ValueError:
        # dd.mm.yyyy
        for sep in (".", "/"):
            parts = s.split(sep)
            if len(parts) == 3:
                try:
                    d, m, y = int(parts[0]), int(parts[1]), int(parts[2])
                    if y < 100:
                        y += 2000
                    return date(y, m, d)
                except (ValueError, IndexError):
                    pass
        return None
    if n < 1:
        return None
    dt = _EXCEL_BASE + timedelta(days=int(n))
    return dt.date()


def fmt_date(d: Optional[date]) -> str:
    if not d:
        return "—"
    return d.strftime("%d.%m.%Y")


def date_from_iso(s: Optional[str]) -> Optional[date]:
    if not s:
        return None
    try:
        return date.fromisoformat(s)
    except ValueError:
        return None


def milestone_detail_block(r: sqlite3.Row) -> str:
    """Колонки как в Excel: B, C, E→G, F, H | L."""
    b = (r["short_code"] or "").strip() or "—"
    c = (r["name"] or "").strip() or "—"
    e = date_from_iso(r["target_date"])
    g = date_from_iso(r["no_fact_date"])
    e_g = f"{fmt_date(e)} → {fmt_date(g)}"
    f_d = fmt_date(date_from_iso(r["forecast_fact_date"]))
    h = (r["status"] or "").strip() or "—"
    l_d = fmt_date(date_from_iso(r["modified_date"]))
    h_l = f"{h} | {l_d}"
    return (
        f"B: {b}\n"
        f"C: {c}\n"
        f"E→G: {e_g}\n"
        f"F: {f_d}\n"
        f"H | L: {h_l}"
    )


def filter_rows_corpus(
    rows: Sequence[sqlite3.Row], corpus_substr: str
) -> List[sqlite3.Row]:
    s = (corpus_substr or "").strip().lower()
    if not s:
        return list(rows)
    return [r for r in rows if s in (r["corpus"] or "").lower()]


def filter_rows_milestone(
    rows: Sequence[sqlite3.Row], needle: str
) -> List[sqlite3.Row]:
    s = (needle or "").strip().lower()
    if not s:
        return list(rows)
    out: List[sqlite3.Row] = []
    for r in rows:
        if s in (r["short_code"] or "").lower():
            out.append(r)
            continue
        if s in (r["name"] or "").lower():
            out.append(r)
            continue
        if s in (str(r["num"] or "").lower()):
            out.append(r)
    return out


def compact_links_line(
    r: sqlite3.Row,
    by_m: Dict[str, sqlite3.Row],
    dependents: Dict[str, List[str]],
) -> str:
    pred = r["n_predecessor"]
    if pred and pred in by_m:
        pr = by_m[pred]
        mark = "✓" if is_done(pr["status"]) else "…"
        lab = (pr["short_code"] or "?")[:14]
        up = f"↑{mark}{lab}"
    elif pred:
        up = "↑вне базы"
    else:
        up = "↑корень"
    k = len(dependents.get(r["m_uuid"], []))
    return f"{up} · ↓{k} след."


def milestone_compact_card(
    r: sqlite3.Row,
    by_m: Dict[str, sqlite3.Row],
    dependents: Dict[str, List[str]],
) -> str:
    """Читаемая короткая карточка (номер, код, название, даты одной строкой, связи)."""
    st = "✅" if is_done(r["status"]) else "⏳"
    b = (r["short_code"] or "—").strip()
    nm = (r["name"] or "").strip()
    if len(nm) > 54:
        nm = nm[:51] + "…"
    e = date_from_iso(r["target_date"])
    g = date_from_iso(r["no_fact_date"])
    f_d = date_from_iso(r["forecast_fact_date"])
    h = (r["status"] or "—").strip()
    l_d = date_from_iso(r["modified_date"])
    corp = ((r["corpus"] or "—").strip())[:40]
    line1 = f"{st} №{r['num'] or '—'} · {b} · 🏗 {corp}"
    line2 = f"   {nm}"
    line3 = (
        f"   E→G: {fmt_date(e)} → {fmt_date(g)}   ·   F: {fmt_date(f_d)}   ·   "
        f"H|L: {h} | {fmt_date(l_d)}"
    )
    line4 = f"   {compact_links_line(r, by_m, dependents)}"
    return "\n".join([line1, line2, line3, line4])


def build_brief_list(
    rows: Sequence[sqlite3.Row],
    corpus_q: str,
    mile_q: str,
    limit: int = 200,
) -> str:
    w = filter_rows_milestone(filter_rows_corpus(rows, corpus_q), mile_q)
    w = sorted(w, key=lambda x: ((x["corpus"] or ""), str(x["num"] or "")))
    lines = [f"📋 Краткий список: {len(w)} вех\n"]
    for i, r in enumerate(w[:limit], 1):
        mark = "✓" if is_done(r["status"]) else "○"
        sn = (r["short_code"] or "—")[:16]
        title = (r["name"] or "")[:36]
        if len(r["name"] or "") > 36:
            title += "…"
        lines.append(
            f"{i}. {mark} {sn} — {title} · E {fmt_date(date_from_iso(r['target_date']))}"
        )
    if len(w) > limit:
        lines.append(f"\n… и ещё {len(w) - limit} (сузьте запрос)")
    if not w:
        lines.append("Ничего не найдено — измените корпус или веху.")
    return "\n".join(lines)


UUID_RE = re.compile(
    r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}"
)


def norm_header(h: str) -> str:
    return " ".join(h.replace("\n", " ").split()).lower()


def detect_columns(headers: Sequence[str]) -> Dict[str, int]:
    """Сопоставление колонок листа «Экспорт» по заголовкам."""
    idx: Dict[str, int] = {}
    for i, h in enumerate(headers):
        n = norm_header(h)
        if "№" in h or n.startswith("номер") or n == "n":
            idx.setdefault("num", i)
        if "короткое" in n and "название" in n:
            idx.setdefault("short", i)
        if n.startswith("название") and "короткое" not in n:
            idx.setdefault("name", i)
        if "корпус" in n:
            idx.setdefault("corpus", i)
        if "целевая" in n or ("дата" in n and "исполнен" in n and "без" not in n and "факт" not in n):
            idx.setdefault("target", i)
        if "прогноз" in n or ("факт" in n and "прогноз" in n.replace(" ", "")):
            idx.setdefault("forecast", i)
        if "без учета факта" in n or "без учёта факта" in n:
            idx.setdefault("no_fact", i)
        if "статус" in n and "вех" in n:
            idx.setdefault("status", i)
        if "ответств" in n and "email" not in n:
            idx.setdefault("owner", i)
        if "email" in n:
            idx.setdefault("email", i)
        if "роль" in n:
            idx.setdefault("role", i)
        if "изменен" in n:
            idx.setdefault("modified", i)
        if "у кого зависят" in n:
            idx.setdefault("m_uuid", i)
        if "от какой вехи" in n:
            idx.setdefault("n_pred", i)
    return idx


def default_column_map() -> Dict[str, int]:
    return {
        "num": 0,
        "short": 1,
        "name": 2,
        "corpus": 3,
        "target": 4,
        "forecast": 5,
        "no_fact": 6,
        "status": 7,
        "owner": 8,
        "email": 9,
        "role": 10,
        "modified": 11,
        "m_uuid": 12,
        "n_pred": 13,
    }


def row_get(row: List[str], key: str, cmap: Dict[str, int]) -> str:
    j = cmap.get(key)
    if j is None or j >= len(row):
        return ""
    return row[j].strip() if row[j] else ""


def is_technical_row(row: List[str], cmap: Dict[str, int]) -> bool:
    num = row_get(row, "num", cmap)
    short = row_get(row, "short", cmap)
    if num == "№" or "короткое" in short.lower():
        return True
    if num.isdigit() and short.isdigit():
        return True
    return False


def init_db() -> None:
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name='milestones'"
    )
    if c.fetchone():
        cols = {r[1] for r in c.execute("PRAGMA table_info(milestones)")}
        if "m_uuid" not in cols or "n_predecessor" not in cols:
            c.execute("DROP TABLE IF EXISTS milestone_dependencies")
            c.execute("DROP TABLE IF EXISTS milestones")
    c.execute(
        """CREATE TABLE IF NOT EXISTS milestones (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        num TEXT,
        short_code TEXT,
        name TEXT,
        corpus TEXT,
        target_date TEXT,
        forecast_fact_date TEXT,
        no_fact_date TEXT,
        status TEXT,
        owner TEXT,
        email TEXT,
        role TEXT,
        modified_date TEXT,
        m_uuid TEXT UNIQUE,
        n_predecessor TEXT,
        updated_at TEXT
    )"""
    )
    c.execute(
        "CREATE INDEX IF NOT EXISTS idx_milestones_n ON milestones (n_predecessor)"
    )
    c.execute(
        "CREATE INDEX IF NOT EXISTS idx_milestones_corpus ON milestones (corpus)"
    )
    conn.commit()
    conn.close()


def get_conn() -> sqlite3.Connection:
    return sqlite3.connect(DB_PATH)


def row_to_record(
    row: List[str], cmap: Dict[str, int]
) -> Optional[Dict[str, Any]]:
    if len(row) < 8 or is_technical_row(row, cmap):
        return None
    short = row_get(row, "short", cmap)
    name = row_get(row, "name", cmap)
    if not short and not name:
        return None

    m_raw = row_get(row, "m_uuid", cmap)
    n_raw = row_get(row, "n_pred", cmap)
    m_uuid = UUID_RE.search(m_raw)
    m_val = m_raw if m_uuid else ""
    n_uuid = UUID_RE.search(n_raw)
    n_val = n_raw if n_uuid else ""

    if not m_val:
        base = f"{short}|{name}|{row_get(row, 'corpus', cmap)}"
        m_val = str(uuid.uuid5(uuid.NAMESPACE_URL, base))

    t = parse_excel_serial(row_get(row, "target", cmap))
    f = parse_excel_serial(row_get(row, "forecast", cmap))
    nf = parse_excel_serial(row_get(row, "no_fact", cmap))
    mod = parse_excel_serial(row_get(row, "modified", cmap))

    return {
        "num": row_get(row, "num", cmap),
        "short_code": short,
        "name": name,
        "corpus": row_get(row, "corpus", cmap),
        "target_date": t.isoformat() if t else None,
        "forecast_fact_date": f.isoformat() if f else None,
        "no_fact_date": nf.isoformat() if nf else None,
        "status": row_get(row, "status", cmap),
        "owner": row_get(row, "owner", cmap),
        "email": row_get(row, "email", cmap),
        "role": row_get(row, "role", cmap),
        "modified_date": mod.isoformat() if mod else None,
        "m_uuid": m_val.strip(),
        "n_predecessor": n_val.strip() or None,
    }


def import_rows(rows: List[List[str]], replace_all: bool) -> Tuple[int, str]:
    if not rows:
        return 0, "Пустой фрагмент."
    header_row = rows[0]
    cmap = detect_columns(header_row)
    if len(cmap) < 5:
        cmap = default_column_map()
        data_rows = rows
    else:
        data_rows = rows[1:]

    recs: List[Dict[str, Any]] = []
    for row in data_rows:
        if not any(x.strip() for x in row):
            continue
        while len(row) < 15:
            row.append("")
        r = row_to_record(row, cmap)
        if r:
            recs.append(r)

    if not recs:
        return 0, "Не удалось разобрать ни одной строки (нужны колонки как на листе «Экспорт»)."

    now = datetime.utcnow().isoformat(timespec="seconds") + "Z"
    conn = get_conn()
    c = conn.cursor()
    if replace_all:
        c.execute("DELETE FROM milestones")
    merged = 0
    for r in recs:
        c.execute("SELECT id FROM milestones WHERE m_uuid = ?", (r["m_uuid"],))
        ex = c.fetchone()
        if ex:
            c.execute(
                """UPDATE milestones SET
                num=?, short_code=?, name=?, corpus=?,
                target_date=?, forecast_fact_date=?, no_fact_date=?,
                status=?, owner=?, email=?, role=?, modified_date=?,
                n_predecessor=?, updated_at=?
                WHERE m_uuid=?""",
                (
                    r["num"],
                    r["short_code"],
                    r["name"],
                    r["corpus"],
                    r["target_date"],
                    r["forecast_fact_date"],
                    r["no_fact_date"],
                    r["status"],
                    r["owner"],
                    r["email"],
                    r["role"],
                    r["modified_date"],
                    r["n_predecessor"],
                    now,
                    r["m_uuid"],
                ),
            )
        else:
            c.execute(
                """INSERT INTO milestones (
                num, short_code, name, corpus, target_date, forecast_fact_date, no_fact_date,
                status, owner, email, role, modified_date, m_uuid, n_predecessor, updated_at
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (
                    r["num"],
                    r["short_code"],
                    r["name"],
                    r["corpus"],
                    r["target_date"],
                    r["forecast_fact_date"],
                    r["no_fact_date"],
                    r["status"],
                    r["owner"],
                    r["email"],
                    r["role"],
                    r["modified_date"],
                    r["m_uuid"],
                    r["n_predecessor"],
                    now,
                ),
            )
        merged += 1
    conn.commit()
    conn.close()
    return merged, "OK"


def load_all() -> List[sqlite3.Row]:
    conn = get_conn()
    conn.row_factory = sqlite3.Row
    rows = conn.execute("SELECT * FROM milestones").fetchall()
    conn.close()
    return rows


def by_m_uuid(rows: Sequence[sqlite3.Row]) -> Dict[str, sqlite3.Row]:
    return {r["m_uuid"]: r for r in rows if r["m_uuid"]}


def build_dependents_map(
    rows: Sequence[sqlite3.Row],
) -> Dict[str, List[str]]:
    """Кто зависит от данной вехи: у них N = наш M."""
    dep: Dict[str, List[str]] = defaultdict(list)
    for r in rows:
        p = r["n_predecessor"]
        if p:
            dep[p].append(r["m_uuid"])
    return dep


def chain_depth_back(
    m_uuid: str,
    by_m: Dict[str, sqlite3.Row],
    memo: Optional[Dict[str, int]] = None,
    visiting: Optional[set] = None,
) -> int:
    """Глубина цепочки назад по предшественникам (0 = корень). Циклы в N→M обрываем."""
    if memo is None:
        memo = {}
    if visiting is None:
        visiting = set()
    if m_uuid in memo:
        return memo[m_uuid]
    if m_uuid in visiting:
        return 0
    visiting.add(m_uuid)
    try:
        r = by_m.get(m_uuid)
        if not r:
            memo[m_uuid] = 0
            return 0
        p = r["n_predecessor"]
        if not p or p not in by_m:
            memo[m_uuid] = 0
            return 0
        d = 1 + chain_depth_back(p, by_m, memo, visiting)
        memo[m_uuid] = d
        return d
    finally:
        visiting.discard(m_uuid)


def compute_slip_profiles(
    rows: Sequence[sqlite3.Row],
) -> Tuple[float, Dict[str, float], Dict[Tuple[str, str], float], Dict[str, int], Dict[Tuple[str, str], int], int]:
    """Глобальная, по корпусу (≥4 завершённых), по (корпус+код) (≥3)."""
    deltas_all: List[int] = []
    by_corpus: Dict[str, List[int]] = defaultdict(list)
    by_pair: Dict[Tuple[str, str], List[int]] = defaultdict(list)
    for r in rows:
        if not is_done(r["status"]):
            continue
        t, f = r["target_date"], r["forecast_fact_date"]
        if not t or not f:
            continue
        try:
            d = (date.fromisoformat(f) - date.fromisoformat(t)).days
        except ValueError:
            continue
        deltas_all.append(d)
        c = r["corpus"] or "—"
        by_corpus[c].append(d)
        sc = (r["short_code"] or "—").strip() or "—"
        by_pair[(c, sc)].append(d)
    glob = float(statistics.median(deltas_all)) if deltas_all else 0.0
    corp_med = {k: float(statistics.median(v)) for k, v in by_corpus.items() if len(v) >= 4}
    pair_med = {k: float(statistics.median(v)) for k, v in by_pair.items() if len(v) >= 3}
    corp_cnt = {k: len(v) for k, v in by_corpus.items()}
    pair_cnt = {k: len(v) for k, v in by_pair.items()}
    return glob, corp_med, pair_med, corp_cnt, pair_cnt, len(deltas_all)


def dep_bucket(n_dep: int) -> int:
    """Грубая категория «сколько вех ждёт эту»."""
    if n_dep <= 0:
        return 0
    if n_dep <= 2:
        return 1
    return 2


def depth_bucket(depth: int) -> int:
    if depth <= 0:
        return 0
    if depth <= 2:
        return 1
    return 2


class DoneAnalog:
    __slots__ = (
        "corpus",
        "short_code",
        "pred_short",
        "lag_fe",
        "span_fp",
        "dep_b",
        "depth_b",
    )

    def __init__(
        self,
        corpus: str,
        short_code: str,
        pred_short: Optional[str],
        lag_fe: int,
        span_fp: Optional[int],
        dep_b: int,
        depth_b: int,
    ) -> None:
        self.corpus = corpus
        self.short_code = short_code
        self.pred_short = pred_short
        self.lag_fe = lag_fe
        self.span_fp = span_fp
        self.dep_b = dep_b
        self.depth_b = depth_b


def build_done_analogs(
    rows: Sequence[sqlite3.Row], by_m: Dict[str, sqlite3.Row]
) -> List[DoneAnalog]:
    """Снимки завершённых вех с метриками для поиска аналогов."""
    dependents = build_dependents_map(rows)
    depth_memo: Dict[str, int] = {}
    out: List[DoneAnalog] = []
    for r in rows:
        if not is_done(r["status"]):
            continue
        t_s, f_s = r["target_date"], r["forecast_fact_date"]
        if not t_s or not f_s:
            continue
        try:
            d_e = date.fromisoformat(t_s)
            d_f = date.fromisoformat(f_s)
        except ValueError:
            continue
        lag = (d_f - d_e).days
        c = (r["corpus"] or "—").strip() or "—"
        sc = ((r["short_code"] or "—").strip() or "—")
        p = r["n_predecessor"]
        p_short: Optional[str] = None
        span_fp: Optional[int] = None
        if p and p in by_m:
            pr = by_m[p]
            p_short = ((pr["short_code"] or "—").strip() or "—")
            if is_done(pr["status"]):
                fp = effective_actual_date(pr)
                if fp is not None:
                    sp = (d_f - fp).days
                    if sp >= 0:
                        span_fp = sp
        uid = r["m_uuid"]
        dep_b = dep_bucket(len(dependents.get(uid, [])))
        depth_b = depth_bucket(chain_depth_back(uid, by_m, depth_memo))
        out.append(DoneAnalog(c, sc, p_short, lag, span_fp, dep_b, depth_b))
    return out


def open_row_pred_short(r: sqlite3.Row, by_m: Dict[str, sqlite3.Row]) -> Optional[str]:
    p = r["n_predecessor"]
    if not p or p not in by_m:
        return None
    return ((by_m[p]["short_code"] or "—").strip() or "—")


def pick_delay_from_history(
    r: sqlite3.Row,
    analogs: Sequence[DoneAnalog],
    by_m: Dict[str, sqlite3.Row],
    dependents: Dict[str, List[str]],
    depth_memo: Dict[str, int],
    pred_finish: Optional[date],
    base_dt: date,
) -> Tuple[float, str, str]:
    """
    Задержка в днях и пояснение. mode: 'span' | 'lag'.
    Прогноз строится из медиан по **похожим завершённым** вехам.
    """
    c = (r["corpus"] or "—").strip() or "—"
    sc = ((r["short_code"] or "—").strip() or "—")
    psc = open_row_pred_short(r, by_m)
    nd = dep_bucket(len(dependents.get(r["m_uuid"], [])))
    dpt = depth_bucket(chain_depth_back(r["m_uuid"], by_m, depth_memo))
    has_pred = bool(r["n_predecessor"] and r["n_predecessor"] in by_m)

    def filt(pred_fn: Callable[[DoneAnalog], bool]) -> List[DoneAnalog]:
        return [a for a in analogs if pred_fn(a)]

    tiers: List[Tuple[str, Callable[[DoneAnalog], bool], int]] = [
        (
            "тот же объект + тип вехи + тип предшественника + похожая нагрузка (число следующих)",
            lambda a: a.corpus == c
            and a.short_code == sc
            and a.pred_short == psc
            and a.dep_b == nd,
            2,
        ),
        (
            "тот же объект + тип вехи + тип предшественника",
            lambda a: a.corpus == c and a.short_code == sc and a.pred_short == psc,
            2,
        ),
        (
            "тот же объект + тип вехи + глубина цепочки + нагрузка",
            lambda a: a.corpus == c
            and a.short_code == sc
            and a.depth_b == dpt
            and a.dep_b == nd,
            2,
        ),
        (
            "тот же объект + тип вехи + глубина цепочки",
            lambda a: a.corpus == c and a.short_code == sc and a.depth_b == dpt,
            2,
        ),
        (
            "тот же объект + тип вехи",
            lambda a: a.corpus == c and a.short_code == sc,
            2,
        ),
        (
            "тот же тип вехи и предшественник (все объекты)",
            lambda a: a.short_code == sc and a.pred_short == psc,
            3,
        ),
        (
            "тот же тип вехи (все объекты)",
            lambda a: a.short_code == sc,
            3,
        ),
        (
            "тот же объект (все типы вех)",
            lambda a: a.corpus == c,
            4,
        ),
    ]

    all_lags = [a.lag_fe for a in analogs]

    for label, pred_fn, min_n in tiers:
        sub = filt(pred_fn)
        if len(sub) < min_n:
            continue
        spans = [a.span_fp for a in sub if a.span_fp is not None]
        if has_pred and pred_finish is not None and len(spans) >= min_n:
            d = float(statistics.median(spans))
            expl = (
                f"Похожие завершённые вехи ({label}): n={len(spans)}.\n"
                f"Медиана дней от факта предшественника до факта вехи: {d:.0f} дн."
            )
            return d, expl, "span"
        lags = [a.lag_fe for a in sub]
        if len(lags) >= min_n:
            d = float(statistics.median(lags))
            expl = (
                f"Похожие завершённые вехи ({label}): n={len(lags)}.\n"
                f"Медиана сдвига план(E)→факт(F) у аналогов: {d:.0f} дн."
            )
            return d, expl, "lag"

    if len(all_lags) >= 3:
        d = float(statistics.median(all_lags))
        expl = (
            f"Мало узких аналогов; взята медиана по всей истории завершённых (n={len(all_lags)}): {d:.0f} дн. (F−E)."
        )
        return d, expl, "lag"
    if all_lags:
        d = float(statistics.median(all_lags))
        expl = f"Мало данных; медиана по доступным завершённым (n={len(all_lags)}): {d:.0f} дн."
        return d, expl, "lag"
    return 0.0, "Нет завершённых вех с датами E и F — задержка 0.", "lag"


async def run_python_script_subprocess(script_path: str) -> tuple[int, str]:
    """Одиночный python-скрипт в SCRIPTS_CWD (наследует env процесса бота / systemd)."""
    proc = await asyncio.create_subprocess_exec(
        "python3",
        script_path,
        cwd=SCRIPTS_CWD,
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.STDOUT,
        env=os.environ.copy(),
    )
    out_b, _ = await proc.communicate()
    out = out_b.decode("utf-8", errors="replace") if out_b else ""
    return int(proc.returncode or 0), out


def main_reply_kb() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[
            [
                KeyboardButton(text="📜 История"),
                KeyboardButton(text="📊 Прогноз"),
            ],
            [
                KeyboardButton(text="📂 По корпусу"),
                KeyboardButton(text="📌 Корпус + веха"),
            ],
            [
                KeyboardButton(text="❓ Справка"),
                KeyboardButton(text="🗑 Очистить базу"),
            ],
        ],
        resize_keyboard=True,
        input_field_placeholder="Таблица из Excel или кнопка",
    )


def cancel_reply_kb() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text="Отмена")]],
        resize_keyboard=True,
    )


def clear_confirm_ikb() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(text="✅ Да, удалить всё", callback_data="clear:yes"),
                InlineKeyboardButton(text="Отмена", callback_data="clear:no"),
            ]
        ]
    )


def corpus_next_ikb() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(text="📜 История", callback_data="corpact:story"),
                InlineKeyboardButton(text="📊 Прогноз", callback_data="corpact:forecast"),
            ],
            [InlineKeyboardButton(text="📋 Краткий список", callback_data="corpact:brief")],
        ]
    )


def is_done(status: Optional[str]) -> bool:
    if not status:
        return False
    s = status.strip().lower()
    return "выполн" in s or s in ("done", "готово", "закрыт")


def effective_actual_date(r: sqlite3.Row) -> Optional[date]:
    if is_done(r["status"]):
        for key in ("forecast_fact_date", "no_fact_date", "target_date"):
            d = r[key]
            if d:
                try:
                    return date.fromisoformat(d)
                except ValueError:
                    pass
    return None


def topo_groups(rows: Sequence[sqlite3.Row]) -> List[List[sqlite3.Row]]:
    """Группы по корпусу, внутри — топологический порядок (предшественники раньше)."""
    by_corpus: Dict[str, List[sqlite3.Row]] = defaultdict(list)
    for r in rows:
        by_corpus[r["corpus"] or "—"].append(r)

    ordered: List[List[sqlite3.Row]] = []
    for corp, lst in sorted(by_corpus.items(), key=lambda x: x[0]):
        mmap = by_m_uuid(lst)
        pred_count: Dict[str, int] = {r["m_uuid"]: 0 for r in lst}
        children: Dict[str, List[str]] = defaultdict(list)
        ids = {r["m_uuid"] for r in lst}
        for r in lst:
            p = r["n_predecessor"]
            if p and p in ids:
                pred_count[r["m_uuid"]] += 1
                children[p].append(r["m_uuid"])

        q = deque([uid for uid in ids if pred_count[uid] == 0])
        chain: List[sqlite3.Row] = []
        while q:
            u = q.popleft()
            chain.append(mmap[u])
            for v in children[u]:
                pred_count[v] -= 1
                if pred_count[v] == 0:
                    q.append(v)
        rest = [r for r in lst if r not in chain]
        chain.extend(rest)
        ordered.append(chain)
    return ordered


MAX_STORY_ITEMS = 100


def build_story(
    rows: Sequence[sqlite3.Row],
    corpus_substr: str = "",
    milestone_substr: str = "",
    max_items: int = MAX_STORY_ITEMS,
) -> str:
    if not rows:
        return "База пуста. Вставьте фрагмент листа «Экспорт» (Ctrl+C / Ctrl+V)."
    total_all = len(rows)
    work = filter_rows_corpus(rows, corpus_substr)
    ms = (milestone_substr or "").strip()
    if ms:
        work = filter_rows_milestone(work, ms)
    if not work:
        if ms:
            return (
                f"Ничего не найдено.\nКорпус: «{corpus_substr.strip() or '(все)'}»\n"
                f"Веха (код/название/№): «{ms}»"
            )
        if corpus_substr.strip():
            return f"Нет вех, где корпус содержит «{corpus_substr.strip()}»."
        return "Пустая выборка."

    shown = 0
    truncated = False
    lines: List[str] = ["📜 История вех\n"]
    if corpus_substr.strip():
        lines.append(f"Корпус: «{corpus_substr.strip()}»")
    if ms:
        lines.append(f"Веха содержит: «{ms}»")
    lines.append(f"В выборке: {len(work)} вех (порядок по графу внутри корпуса).\n")
    if not corpus_substr.strip() and not ms and total_all > max_items:
        lines.append(
            f"⚠️ Во всей базе {total_all} вех. Показано до {max_items}. "
            f"Используйте «📂 По корпусу» или /корпус …\n"
        )

    mmap = by_m_uuid(load_all())
    dependents = build_dependents_map(load_all())

    for grp in topo_groups(work):
        corp = (grp[0]["corpus"] or "—").strip()
        lines.append(f"── {corp[:55]}{'…' if len(corp) > 55 else ''} ──")
        for r in grp:
            if shown >= max_items:
                truncated = True
                break
            lines.append(milestone_compact_card(r, mmap, dependents))
            lines.append("")
            shown += 1
        if shown >= max_items:
            truncated = True
            break
    if truncated:
        lines.append(f"… показано {max_items} из {len(work)}. Сузьте корпус или веху.")
    return "\n".join(lines).strip()


MAX_FORECAST_ITEMS = 80


def build_forecast(
    corpus_substr: str = "",
    milestone_substr: str = "",
    max_open: int = MAX_FORECAST_ITEMS,
) -> str:
    all_r = load_all()
    if not all_r:
        return "Нет данных для прогноза."
    mmap = by_m_uuid(all_r)
    dependents = build_dependents_map(all_r)
    glob, _, _, _, _, n_done = compute_slip_profiles(all_r)
    analogs = build_done_analogs(all_r, mmap)
    depth_memo_fc: Dict[str, int] = {}

    memo_fin: Dict[str, Optional[date]] = {}
    finish_visiting: set = set()

    def finish_of(uid: str) -> Optional[date]:
        if uid in memo_fin:
            return memo_fin[uid]
        row = mmap.get(uid)
        if not row:
            memo_fin[uid] = None
            return None
        if is_done(row["status"]):
            d = effective_actual_date(row)
            memo_fin[uid] = d
            return d
        if uid in finish_visiting:
            # цикл в цепочке открытых вех — временно только план E, без записи в memo
            return date_from_iso(row["target_date"])
        finish_visiting.add(uid)
        try:
            preds: List[date] = []
            p = row["n_predecessor"]
            if p and p in mmap:
                pd = finish_of(p)
                if pd:
                    preds.append(pd)
            tgt = date_from_iso(row["target_date"])
            if preds:
                base_dt = max(preds)
                if tgt:
                    base_dt = max(base_dt, tgt)
            else:
                base_dt = tgt
            if base_dt is None:
                memo_fin[uid] = None
                return None
            pred_finish_date = max(preds) if preds else None
            d_days, _expl, mode = pick_delay_from_history(
                row,
                analogs,
                mmap,
                dependents,
                depth_memo_fc,
                pred_finish_date,
                base_dt,
            )
            if mode == "span" and pred_finish_date is not None:
                cand = pred_finish_date + timedelta(
                    days=int(round(max(0.0, d_days)))
                )
                memo_fin[uid] = max(cand, base_dt)
            else:
                memo_fin[uid] = base_dt + timedelta(
                    days=int(round(max(0.0, d_days)))
                )
            return memo_fin[uid]
        finally:
            finish_visiting.discard(uid)

    def explain_row(r: sqlite3.Row) -> str:
        uid = r["m_uuid"]
        preds: List[date] = []
        p = r["n_predecessor"]
        pred_lbl = "нет"
        if p and p in mmap:
            pd = finish_of(p)
            pr = mmap[p]
            pred_lbl = pr["short_code"] or pr["name"][:30]
            if pd:
                preds.append(pd)
        tgt = date_from_iso(r["target_date"])
        if preds:
            base_dt = max(preds)
            base_src = (
                f"Опорная дата (не раньше неё): max(прогноз/факт предшественников, план E) = {fmt_date(base_dt)}"
            )
            if tgt:
                base_dt = max(base_dt, tgt)
                base_src = (
                    f"Опорная дата: max(предш.: {fmt_date(max(preds))}, план E: {fmt_date(tgt)}) = {fmt_date(base_dt)}"
                )
        elif tgt:
            base_dt = tgt
            base_src = f"Опорная дата: план E = {fmt_date(tgt)} (предшественник вне графа или без даты)"
        else:
            return "Нет ни плана E, ни даты предшественника — прогноз невозможен."

        pred_finish_date = max(preds) if preds else None
        d_days, hist_expl, mode = pick_delay_from_history(
            r,
            analogs,
            mmap,
            dependents,
            depth_memo_fc,
            pred_finish_date,
            base_dt,
        )
        parts = [base_src, "", hist_expl]
        if mode == "span" and pred_finish_date is not None:
            parts.append(
                f"\nФормула: дата готовности предшественника ({fmt_date(pred_finish_date)}) "
                f"+ типичный интервал по аналогам ({d_days:.0f} дн.), не раньше {fmt_date(base_dt)}."
            )
        else:
            parts.append(
                f"\nФормула: опорная дата ({fmt_date(base_dt)}) + типичный сдвиг E→F по аналогам ({d_days:.0f} дн.)."
            )
        return "\n".join(parts)

    pool = filter_rows_milestone(filter_rows_corpus(all_r, corpus_substr), milestone_substr)
    if corpus_substr.strip() and not filter_rows_corpus(all_r, corpus_substr):
        return f"Нет вех с корпусом, содержащим «{corpus_substr.strip()}»."
    if (milestone_substr or "").strip() and not pool:
        return (
            f"Нет вех по корпусу «{corpus_substr.strip() or '(все)'}» "
            f"и запросу вехи «{milestone_substr.strip()}»."
        )

    open_rows = [r for r in pool if not is_done(r["status"])]
    open_rows.sort(key=lambda x: (x["corpus"] or "", x["num"] or ""))

    lines = [
        "📊 Прогноз (по истории похожих завершённых вех)\n"
        "Отбор аналогов: корпус → тип вехи → предшественник → нагрузка/глубина; "
        "интервал предш→факт или сдвиг E→F.\n"
        f"В базе {n_done} завершённых с E и F; глобальная медиана (F−E) ≈ {glob:.0f} дн.\n",
    ]
    if corpus_substr.strip():
        lines.append(f"Корпус: «{corpus_substr.strip()}»")
    if (milestone_substr or "").strip():
        lines.append(f"Веха: «{milestone_substr.strip()}»")
    lines.append(f"Открытых в выборке: {len(open_rows)}\n")

    shown = 0
    truncated = False
    for r in open_rows:
        if shown >= max_open:
            truncated = True
            break
        pr = finish_of(r["m_uuid"])
        expl = explain_row(r)
        lines.append("──────────────")
        lines.append(milestone_compact_card(r, mmap, dependents))
        lines.append(f"⏱ Прогноз окончания: {fmt_date(pr)}")
        lines.append("📐 Как посчитано:\n" + expl)
        lines.append("")
        shown += 1
    if len(lines) <= 6 and not open_rows:
        lines.append("Все вехи в статусе «выполнено» или нет открытых в этой выборке.")
    if truncated:
        lines.append(
            f"… показано {max_open} из {len(open_rows)}. Сузьте: 📂 / 📌 или /прогноз …"
        )
    return "\n".join(lines).strip()


def looks_like_table(text: str) -> bool:
    lines = text.strip().splitlines()
    if len(lines) < 2:
        return False
    tabs = text.count("\t")
    return tabs >= 5 or (tabs >= 3 and len(lines) >= 3)


def parse_pasted_table(text: str) -> List[List[str]]:
    # UTF-8; Excel on Mac often uses \t
    reader = io.StringIO(text.strip())
    rows: List[List[str]] = []
    for line in reader:
        if not line.strip():
            continue
        rows.append(line.rstrip("\n\r").split("\t"))
    return rows


def try_parse_freeform_add(text: str) -> Optional[Dict[str, Any]]:
    """Добавление одной вехи из текста: ищем UUID, короткое имя, название, даты, предшественника."""
    uuids = UUID_RE.findall(text)
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    data: Dict[str, str] = {}
    for ln in lines:
        if ":" in ln or "=" in ln:
            sep = ":" if ":" in ln else "="
            k, v = ln.split(sep, 1)
            data[norm_header(k)] = v.strip()
    short = data.get("короткое", data.get("short", ""))
    name = data.get("название", data.get("name", ""))
    corpus = data.get("корпус", data.get("corpus", ""))
    status = data.get("статус", data.get("status", "Новая"))
    target = data.get("целевая", data.get("дата", data.get("target", "")))
    if not short and not name:
        # одна строка — возможно «РНС<TAB>...»
        if "\t" in text:
            parts = text.split("\t")
            if len(parts) >= 2:
                short, name = parts[0].strip(), parts[1].strip()
    if not short and not name:
        return None

    m_uuid = ""
    for u in uuids:
        if data.get("m") == u or data.get("m_uuid") == u:
            m_uuid = u
    if not m_uuid:
        if uuids and "предшественник" not in text.lower() and "от какой" not in text.lower():
            m_uuid = uuids[0]
        elif len(uuids) >= 2:
            m_uuid = uuids[0]
        else:
            m_uuid = str(uuid.uuid4())

    pred = data.get("от какой вехи", data.get("предшественник", data.get("n", "")))
    if not pred:
        for u in uuids:
            if u != m_uuid:
                pred = u
                break

    t = parse_excel_serial(target) if target else None
    now = datetime.utcnow().isoformat(timespec="seconds") + "Z"
    return {
        "num": data.get("№", data.get("num", "")),
        "short_code": short,
        "name": name,
        "corpus": corpus,
        "target_date": t.isoformat() if t else None,
        "forecast_fact_date": None,
        "no_fact_date": None,
        "status": status,
        "owner": data.get("ответственный", data.get("owner", "")),
        "email": data.get("email", ""),
        "role": data.get("роль", ""),
        "modified_date": None,
        "m_uuid": m_uuid.strip(),
        "n_predecessor": pred.strip() if pred else None,
        "_updated_at": now,
    }


def insert_single(rec: Dict[str, Any]) -> str:
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT id FROM milestones WHERE m_uuid = ?", (rec["m_uuid"],))
    ex = c.fetchone()
    now = rec.get("_updated_at") or datetime.utcnow().isoformat(timespec="seconds") + "Z"
    if ex:
        c.execute(
            """UPDATE milestones SET
            num=?, short_code=?, name=?, corpus=?,
            target_date=?, forecast_fact_date=?, no_fact_date=?,
            status=?, owner=?, email=?, role=?, modified_date=?,
            n_predecessor=?, updated_at=?
            WHERE m_uuid=?""",
            (
                rec["num"],
                rec["short_code"],
                rec["name"],
                rec["corpus"],
                rec["target_date"],
                rec["forecast_fact_date"],
                rec["no_fact_date"],
                rec["status"],
                rec["owner"],
                rec["email"],
                rec["role"],
                rec["modified_date"],
                rec["n_predecessor"],
                now,
                rec["m_uuid"],
            ),
        )
        msg = "Обновлена существующая веха."
    else:
        c.execute(
            """INSERT INTO milestones (
            num, short_code, name, corpus, target_date, forecast_fact_date, no_fact_date,
            status, owner, email, role, modified_date, m_uuid, n_predecessor, updated_at
        ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (
                rec["num"],
                rec["short_code"],
                rec["name"],
                rec["corpus"],
                rec["target_date"],
                rec["forecast_fact_date"],
                rec["no_fact_date"],
                rec["status"],
                rec["owner"],
                rec["email"],
                rec["role"],
                rec["modified_date"],
                rec["m_uuid"],
                rec["n_predecessor"],
                now,
            ),
        )
        msg = "Добавлена новая веха."
    conn.commit()
    conn.close()
    return msg


async def send_chunks(
    message: Message,
    text: str,
    reply_markup: Optional[InlineKeyboardMarkup | ReplyKeyboardMarkup] = None,
) -> None:
    kb = reply_markup if reply_markup is not None else main_reply_kb()
    max_len = 4000
    if len(text) <= max_len:
        await message.answer(text, reply_markup=kb)
        return
    for i in range(0, len(text), max_len):
        await message.answer(text[i : i + max_len], reply_markup=kb)


MENU_BUTTONS = {
    "📜 История",
    "📊 Прогноз",
    "❓ Справка",
    "📂 По корпусу",
    "📌 Корпус + веха",
    "🗑 Очистить базу",
}


async def show_search_result_menu(
    message: Message,
    state: FSMContext,
    corpus: str,
    milestone: str,
) -> None:
    corpus = (corpus or "").strip()
    milestone = (milestone or "").strip()
    await state.update_data(pending_corpus=corpus, pending_milestone=milestone)
    n = len(filter_rows_milestone(filter_rows_corpus(load_all(), corpus), milestone))
    if n == 0:
        await message.answer(
            "По такому запросу ничего не найдено.",
            reply_markup=main_reply_kb(),
        )
        await state.clear()
        return
    sub = f"Корпус: «{corpus}»" if corpus else "Корпус: (все)"
    if milestone:
        sub += f"\nВеха: «{milestone}» (код, фрагмент названия или №)"
    await message.answer(
        f"Найдено вех: {n}\n{sub}\n\nВыберите действие:",
        reply_markup=corpus_next_ikb(),
    )


@dp.callback_query(F.data == "clear:yes")
async def cb_clear_yes(query: CallbackQuery) -> None:
    conn = get_conn()
    conn.execute("DELETE FROM milestones")
    conn.commit()
    conn.close()
    await query.message.answer("Все вехи удалены из базы бота.", reply_markup=main_reply_kb())
    await query.answer("Готово")


@dp.callback_query(F.data == "clear:no")
async def cb_clear_no(query: CallbackQuery) -> None:
    await query.message.answer("Очистка отменена.", reply_markup=main_reply_kb())
    await query.answer()


@dp.callback_query(F.data.startswith("corpact:"))
async def cb_corpus_action(query: CallbackQuery, state: FSMContext) -> None:
    action = query.data.split(":", 1)[1]
    data = await state.get_data()
    corpus = (data.get("pending_corpus") or "").strip()
    mile = (data.get("pending_milestone") or "").strip()
    if not corpus and not mile:
        await query.answer(
            "Сначала задайте поиск: кнопки «По корпусу» / «Корпус+веха» или /корпус …",
            show_alert=True,
        )
        return
    rows = load_all()
    if action == "story":
        txt = build_story(rows, corpus_substr=corpus, milestone_substr=mile)
    elif action == "forecast":
        txt = build_forecast(corpus_substr=corpus, milestone_substr=mile)
    else:
        txt = build_brief_list(rows, corpus, mile)
    await state.clear()
    await send_chunks(query.message, txt)
    await query.answer()


@dp.message(Command("start"))
@dp.message(Command("help"))
async def cmd_help(message: Message, state: FSMContext) -> None:
    await state.clear()
    await message.answer(
        "Вехи НДК — лист «Экспорт».\n\n"
        "• 📂 По корпусу — вводите фрагмент названия корпуса, затем кнопки: история / прогноз / краткий список.\n"
        "• 📌 Корпус + веха — сначала корпус, затем код вехи или слово из названия.\n"
        "• История / Прогноз — по всей базе (много вех, лучше сузить через 📂).\n"
        "• Команды: /корпус или /corpus …  |  /веха или /milestone <корпус> <веха>  |  /история …  |  /прогноз …\n"
        "• Импорт: вставьте таблицу из Excel.\n"
        "Связи: N = M другой строки. В карточках: E→G, F, H|L.",
        reply_markup=main_reply_kb(),
    )


@dp.message(Command("история"))
async def cmd_story(message: Message, state: FSMContext, command: CommandObject) -> None:
    await state.clear()
    arg = (command.args or "").strip()
    rows = load_all()
    await send_chunks(message, build_story(rows, corpus_substr=arg))


@dp.message(Command("прогноз"))
async def cmd_forecast(
    message: Message, state: FSMContext, command: CommandObject
) -> None:
    await state.clear()
    arg = (command.args or "").strip()
    await send_chunks(message, build_forecast(corpus_substr=arg))


@dp.message(Command("корпус"))
@dp.message(Command("corpus"))
async def cmd_corpus(message: Message, state: FSMContext, command: CommandObject) -> None:
    arg = (command.args or "").strip()
    if not arg:
        await message.answer(
            "Укажите фрагмент корпуса, например:\n/корпус ПБ_ф10_к5",
            reply_markup=main_reply_kb(),
        )
        return
    await show_search_result_menu(message, state, arg, "")


@dp.message(Command("веха"))
@dp.message(Command("milestone"))
async def cmd_veha(message: Message, state: FSMContext, command: CommandObject) -> None:
    raw = (command.args or "").strip()
    if not raw:
        await message.answer(
            "Формат: /веха <корпус> <код или название>\n"
            "Пример: /веха ПБ_ф10_к5 РНС",
            reply_markup=main_reply_kb(),
        )
        return
    parts = raw.split(None, 1)
    if len(parts) < 2:
        await message.answer(
            "Нужны два поля через пробел: корпус и веха.\n"
            "Пример: /веха ПБ_ф10_к5 экспертиза",
            reply_markup=main_reply_kb(),
        )
        return
    co, ve = parts[0], parts[1].strip()
    await show_search_result_menu(message, state, co, ve)


@dp.message(Command("очистить"))
async def cmd_clear(message: Message, state: FSMContext) -> None:
    await state.clear()
    await message.answer(
        "Удалить все вехи из базы бота?",
        reply_markup=clear_confirm_ikb(),
    )


@dp.message(Form.waiting_search_corpus, F.text)
async def fsm_search_corpus(message: Message, state: FSMContext) -> None:
    text = (message.text or "").strip()
    if text.lower() in ("отмена", "cancel"):
        await state.clear()
        await message.answer("Отменено.", reply_markup=main_reply_kb())
        return
    if looks_like_table(text):
        await state.clear()
        rows = parse_pasted_table(text)
        n, _ = import_rows(rows, replace_all=False)
        await message.answer(f"Импортировано/обновлено строк: {n}", reply_markup=main_reply_kb())
        return
    if not text:
        await message.answer("Введите непустой фрагмент корпуса или «Отмена».")
        return
    await show_search_result_menu(message, state, text, "")


@dp.message(Form.waiting_pair_corpus, F.text)
async def fsm_pair_corpus(message: Message, state: FSMContext) -> None:
    text = (message.text or "").strip()
    if text.lower() in ("отмена", "cancel"):
        await state.clear()
        await message.answer("Отменено.", reply_markup=main_reply_kb())
        return
    if looks_like_table(text):
        await state.clear()
        rows = parse_pasted_table(text)
        n, _ = import_rows(rows, replace_all=False)
        await message.answer(f"Импортировано/обновлено строк: {n}", reply_markup=main_reply_kb())
        return
    if not text:
        await message.answer("Введите фрагмент корпуса.")
        return
    await state.update_data(pair_corpus=text)
    await state.set_state(Form.waiting_pair_milestone)
    await message.answer(
        f"Корпус: «{text}»\nТеперь введите короткий код вехи (ЭКС, РНС…) "
        f"или слово из полного названия.\nОтмена — «Отмена».",
        reply_markup=cancel_reply_kb(),
    )


@dp.message(Form.waiting_pair_milestone, F.text)
async def fsm_pair_milestone(message: Message, state: FSMContext) -> None:
    text = (message.text or "").strip()
    if text.lower() in ("отмена", "cancel"):
        await state.clear()
        await message.answer("Отменено.", reply_markup=main_reply_kb())
        return
    if looks_like_table(text):
        await state.clear()
        rows = parse_pasted_table(text)
        n, _ = import_rows(rows, replace_all=False)
        await message.answer(f"Импортировано/обновлено строк: {n}", reply_markup=main_reply_kb())
        return
    data = await state.get_data()
    co = (data.get("pair_corpus") or "").strip()
    if not co:
        await state.clear()
        await message.answer("Сбой шага. Начните с «📌 Корпус + веха».", reply_markup=main_reply_kb())
        return
    if not text:
        await message.answer("Введите запрос по вехе.")
        return
    await show_search_result_menu(message, state, co, text)


@dp.message(F.text.in_(MENU_BUTTONS))
async def handle_menu_buttons(message: Message, state: FSMContext) -> None:
    text = message.text or ""
    if text == "📂 По корпусу":
        await state.set_state(Form.waiting_search_corpus)
        await message.answer(
            "Введите фрагмент названия корпуса (без учёта регистра).\n"
            "Потом выберите: история, прогноз или краткий список.\n"
            "Отмена — «Отмена».",
            reply_markup=cancel_reply_kb(),
        )
        return
    if text == "📌 Корпус + веха":
        await state.set_state(Form.waiting_pair_corpus)
        await message.answer(
            "Шаг 1/2: введите фрагмент корпуса.\nОтмена — «Отмена».",
            reply_markup=cancel_reply_kb(),
        )
        return
    await state.clear()
    if text == "📜 История":
        await send_chunks(message, build_story(load_all(), ""))
    elif text == "📊 Прогноз":
        await send_chunks(message, build_forecast(""))
    elif text == "❓ Справка":
        await cmd_help(message, state)
    elif text == "🗑 Очистить базу":
        await message.answer(
            "Удалить все вехи?",
            reply_markup=clear_confirm_ikb(),
        )


@dp.message(F.text)
async def on_text(message: Message, state: FSMContext) -> None:
    text = message.text or ""
    if looks_like_table(text):
        await state.clear()
        rows = parse_pasted_table(text)
        n, _ = import_rows(rows, replace_all=False)
        await message.answer(
            f"Импортировано/обновлено строк: {n}",
            reply_markup=main_reply_kb(),
        )
        return
    add = try_parse_freeform_add(text)
    if add:
        await state.clear()
        msg = insert_single(add)
        await message.answer(
            f"{msg}\nM={add['m_uuid']}\nN→{add['n_predecessor'] or '—'}",
            reply_markup=main_reply_kb(),
        )
        return
    await message.answer(
        "Не понял сообщение. Используйте кнопки внизу или вставьте таблицу «Экспорт».",
        reply_markup=main_reply_kb(),
    )


async def main() -> None:
    init_db()
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
