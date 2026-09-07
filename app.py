# -*- coding: utf-8 -*-
"""
每日單字測驗 — 單檔 Streamlit Web App
=====================================

資料來源（三擇一，於側邊欄切換）
  1. Google Sheet 公開連結（自動轉成 CSV export 網址）
  2. 上傳 CSV / Excel
  3. 內建範例單字庫 sample_vocab.csv

單字庫欄位（中英文標題皆可，缺的欄位會自動補空白）
  word / 單字         *必填
  meaning / 中文意思  *必填
  pos / 詞性
  example / 例句
  example_zh / 例句翻譯
  note / 備註          ← 熟詞偏義說明寫這裡
  tags / 標籤          ← 含「熟詞偏義」字樣者會被特別標註

執行： streamlit run app.py
"""
from __future__ import annotations

import datetime as dt
import io
import json
import random
import re
import string
from pathlib import Path

import pandas as pd
import streamlit as st

# --------------------------------------------------------------------------
# 基本設定
# --------------------------------------------------------------------------
APP_TITLE = "每日單字測驗"
HERE = Path(__file__).resolve().parent
SAMPLE_FILE = HERE / "sample_vocab.csv"
PROGRESS_FILE = HERE / "progress.json"

CANON_COLUMNS = ["word", "meaning", "pos", "example", "example_zh", "note", "tags"]
REQUIRED_COLUMNS = ("word", "meaning")

COLUMN_ALIASES = {
    "word": ["word", "單字", "英文", "英文單字", "vocabulary", "vocab", "term", "en"],
    "meaning": ["meaning", "中文", "中文意思", "中文解釋", "中文翻譯", "中文釋義",
                "意思", "解釋", "釋義", "字義", "翻譯", "中譯", "定義",
                "definition", "zh", "chinese"],
    "pos": ["pos", "詞性", "partofspeech", "part_of_speech"],
    "example": ["example", "例句", "英文例句", "sentence", "examplesentence", "用法"],
    "example_zh": ["example_zh", "例句翻譯", "例句中文", "中文例句", "sentence_zh"],
    "note": ["note", "notes", "備註", "註記", "說明", "筆記", "remark", "熟詞偏義", "陷阱"],
    "tags": ["tags", "tag", "標籤", "分類", "類別", "category", "主題", "level", "程度"],
}
TRICKY_KEYWORDS = ("熟詞偏義", "偏義", "陷阱", "tricky", "trap")

KIND_LABELS = {
    "en2zh": "英 → 中（四選一）",
    "zh2en": "中 → 英（四選一）",
    "cloze": "例句填空（四選一）",
    "spell": "例句填空（拼字輸入）",
}
DEFAULT_KINDS = ["en2zh", "zh2en", "cloze"]

MODES = {
    "daily": "每日精選（同一天題目固定）",
    "free": "自由練習（每次重抽）",
    "wrong": "錯題複習",
    "tricky": "熟詞偏義特訓",
}

BLANK = "＿＿＿＿＿"

st.set_page_config(page_title=APP_TITLE, page_icon="📘", layout="centered",
                   initial_sidebar_state="collapsed")

st.markdown(
    """
    <style>
      .block-container {padding-top: 2.2rem; padding-bottom: 4rem; max-width: 820px;}
      h1 {font-size: 2rem !important;}
      @media (max-width: 480px) {
          h1 {font-size: 1.55rem !important;}
          .block-container {padding-left: .9rem; padding-right: .9rem;}
      }
      div.stButton > button {
          min-height: 3.1rem; font-size: 1.02rem; line-height: 1.45;
          white-space: normal; padding: 0.6rem 0.95rem; justify-content: flex-start;
      }
      /* 選項文字靠左，長中文釋義換行後才對齊得好看 */
      div.stButton > button, div.stButton > button div, div.stButton > button p {text-align: left;}
      div.stButton > button > div {width: 100%; justify-content: flex-start;}
      .q-card {
          border: 1px solid rgba(128,128,128,.28); border-radius: 14px;
          padding: 1.1rem 1.2rem; margin-bottom: .9rem;
      }
      .q-prompt {font-size: 1.6rem; font-weight: 700; line-height: 1.5; margin: .25rem 0 .35rem;}
      .q-sent {font-size: 1.18rem; line-height: 1.8; margin: .35rem 0;}
      .q-meta {font-size: .9rem; opacity: .7;}
      .opt-row {font-size: 1.02rem; padding: .6rem .85rem; border-radius: 10px;
                margin-bottom: .4rem; border: 1px solid rgba(128,128,128,.25);}
      .opt-ok {background: rgba(33,195,84,.16); border-color: rgba(33,195,84,.5);}
      .opt-ng {background: rgba(255,75,75,.14); border-color: rgba(255,75,75,.45);}
      .pill {display:inline-block; padding:.14rem .6rem; border-radius:999px;
             font-size:.78rem; margin-right:.35rem; background:rgba(128,128,128,.18);}
      .pill-tricky {background:rgba(255,170,0,.28);}
    </style>
    """,
    unsafe_allow_html=True,
)


# --------------------------------------------------------------------------
# 資料讀取與正規化
# --------------------------------------------------------------------------
def _canon(name) -> str:
    return re.sub(r"[\s_\-()（）]", "", str(name)).strip().lower()


ALIAS_LOOKUP = {_canon(a): canon for canon, alist in COLUMN_ALIASES.items() for a in alist}

CJK_RE = re.compile(r"[㐀-鿿]")
TAIL_PAREN_RE = re.compile(r"[（(]([^（()）]*)[)）]\s*$")


def split_bilingual(sentence: str):
    """把「English sentence. (中文翻譯。)」拆成 (英文句, 中文句)。

    很多人的單字表會把例句和翻譯塞在同一格。若不拆開，出填空題時
    中文翻譯會緊跟在空格旁邊 → 等於直接把答案講出來。
    """
    s = str(sentence).strip()
    if not s or not CJK_RE.search(s):
        return s, ""

    m = TAIL_PAREN_RE.search(s)                     # 情況一：翻譯放在句尾括號裡
    if m and CJK_RE.search(m.group(1)):
        return s[: m.start()].strip(), m.group(1).strip()

    m = CJK_RE.search(s)                            # 情況二：中文直接接在英文後面
    head, tail = s[: m.start()].strip(), s[m.start():].strip()
    if len(head.split()) >= 3:                      # 前面要真的是一個英文句子才切
        return head, tail.strip("（()）").strip()
    return s, ""


def normalize_frame(raw: pd.DataFrame) -> pd.DataFrame:
    df = raw.copy()
    rename = {}
    for col in df.columns:
        target = ALIAS_LOOKUP.get(_canon(col))
        if target and target not in rename.values():
            rename[col] = target
    df = df.rename(columns=rename)

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            "缺少必要欄位：" + "、".join(missing)
            + "。表頭請至少包含 word/單字 與 meaning/中文意思。目前讀到的欄位："
            + "、".join(map(str, raw.columns))
        )

    for col in CANON_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[CANON_COLUMNS]

    for col in CANON_COLUMNS:
        df[col] = df[col].fillna("").astype(str).str.strip()

    df = df[(df["word"] != "") & (df["meaning"] != "")]
    df = df.drop_duplicates(subset=["word"], keep="first").reset_index(drop=True)

    # 例句與翻譯若擠在同一格，拆開；已另有翻譯欄位者不動
    split = df["example"].apply(split_bilingual)
    df["example"] = [en for en, _ in split]
    df["example_zh"] = [zh if zh and not old else old
                        for (_, zh), old in zip(split, df["example_zh"])]
    df["tricky"] = (df["tags"] + " " + df["note"]).str.lower().apply(
        lambda s: any(k.lower() in s for k in TRICKY_KEYWORDS)
    )
    return df


def sheet_url_to_csv(url: str) -> str:
    """把 Google Sheet 一般網址轉成 CSV 匯出網址；已是 CSV 連結則原樣返回。"""
    url = url.strip()
    if not url or "/spreadsheets/" not in url:
        return url
    if "format=csv" in url or url.endswith(".csv"):
        return url
    m = re.search(r"/spreadsheets/d/([A-Za-z0-9_\-]+)", url)
    if not m:
        return url
    gid_m = re.search(r"[#&?]gid=([0-9]+)", url)
    gid = gid_m.group(1) if gid_m else "0"
    return f"https://docs.google.com/spreadsheets/d/{m.group(1)}/export?format=csv&gid={gid}"


def explain_load_error(e: Exception, url: str) -> str:
    """把 pandas / urllib 的錯誤翻成「該怎麼修」。"""
    msg = str(e)
    if "401" in msg or "403" in msg or "Unauthorized" in msg or "Forbidden" in msg:
        return ("這份 Google Sheet 還沒開放公開讀取（HTTP 401/403）。\n\n"
                "請到 Sheet 右上角 **共用** → 一般存取權改成 "
                "**「知道連結的任何人」＋「檢視者」** → 完成，再按側邊欄的「🔄 重新載入單字庫」。")
    if "404" in msg or "Not Found" in msg:
        return "找不到這份 Sheet（HTTP 404）。請確認網址正確、分頁的 gid 沒填錯。"
    if "缺少必要欄位" in msg:
        return msg + "\n\n第一列請放表頭，至少要有 `word`（或「單字」）與 `meaning`（或「中文意思」）兩欄。"
    if "No columns to parse" in msg or "EmptyData" in msg:
        return "這個分頁是空的，或 gid 指到了空白工作表。"
    return f"讀取失敗：{msg}"


@st.cache_data(ttl=600, show_spinner="讀取單字庫中…")
def load_from_url(url: str) -> pd.DataFrame:
    return normalize_frame(pd.read_csv(sheet_url_to_csv(url)))


@st.cache_data(show_spinner=False)
def load_from_bytes(data: bytes, name: str) -> pd.DataFrame:
    if name.lower().endswith((".xlsx", ".xlsm", ".xls")):
        raw = pd.read_excel(io.BytesIO(data))
    else:
        try:
            raw = pd.read_csv(io.BytesIO(data))
        except UnicodeDecodeError:
            raw = pd.read_csv(io.BytesIO(data), encoding="big5")
    return normalize_frame(raw)


@st.cache_data(show_spinner=False)
def load_sample() -> pd.DataFrame:
    return normalize_frame(pd.read_csv(SAMPLE_FILE, encoding="utf-8"))


# --------------------------------------------------------------------------
# 學習紀錄（寫本機檔案，best-effort；雲端容器重啟會清空，可下載備份）
# --------------------------------------------------------------------------
def load_progress() -> dict:
    if "progress" in st.session_state:
        return st.session_state.progress
    data = {"history": {}, "wrong": {}}
    try:
        if PROGRESS_FILE.exists():
            data = json.loads(PROGRESS_FILE.read_text(encoding="utf-8"))
            data.setdefault("history", {})
            data.setdefault("wrong", {})
    except Exception:
        pass
    st.session_state.progress = data
    return data


def save_progress() -> None:
    try:
        PROGRESS_FILE.write_text(
            json.dumps(st.session_state.progress, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except Exception:
        pass  # 雲端檔案系統唯讀時忽略


def record_result(word: str, ok: bool) -> None:
    prog = load_progress()
    today = dt.date.today().isoformat()
    day = prog["history"].setdefault(today, {"total": 0, "correct": 0})
    day["total"] += 1
    day["correct"] += int(ok)

    entry = prog["wrong"].get(word)
    if ok:
        if entry:
            entry["streak"] = entry.get("streak", 0) + 1
            if entry["streak"] >= 2:            # 連續答對兩次就畢業，移出錯題本
                prog["wrong"].pop(word, None)
    else:
        entry = entry or {"count": 0}
        entry["count"] = entry.get("count", 0) + 1
        entry["streak"] = 0
        entry["last"] = today
        prog["wrong"][word] = entry
    save_progress()


def study_streak() -> int:
    prog = load_progress()
    days = {d for d, v in prog["history"].items() if v.get("total", 0) > 0}
    if not days:
        return 0
    today = dt.date.today()
    cur = today
    if today.isoformat() not in days:
        cur = today - dt.timedelta(days=1)
        if cur.isoformat() not in days:
            return 0
    n = 0
    while cur.isoformat() in days:
        n += 1
        cur -= dt.timedelta(days=1)
    return n


# --------------------------------------------------------------------------
# 出題邏輯
# --------------------------------------------------------------------------
def pick_distractors(pool: pd.DataFrame, row: pd.Series, field: str,
                     rng: random.Random, k: int = 3) -> list:
    """干擾選項：優先同詞性 → 同標籤 → 全庫隨機，且不與正解重複。"""
    correct = str(row[field]).strip()
    pos = str(row.get("pos", "")).strip()
    tag = str(row.get("tags", "")).strip()

    tiers = []
    if pos:
        tiers.append(pool[pool["pos"].str.strip() == pos])
    if tag:
        tiers.append(pool[pool["tags"].str.strip() == tag])
    tiers.append(pool)

    seen = {correct.lower()}
    out = []
    for tier in tiers:
        vals = [str(v).strip() for v in tier[field].tolist() if str(v).strip()]
        rng.shuffle(vals)
        for v in vals:
            if v.lower() in seen:
                continue
            seen.add(v.lower())
            out.append(v)
            if len(out) >= k:
                return out
    return out


def make_blank(sentence: str, word: str):
    """把例句中的目標字（含常見詞形變化）挖空。回傳 (挖空句, 原字面) 或 None。"""
    sentence, word = str(sentence).strip(), str(word).strip()
    if not sentence or not word:
        return None
    base = re.escape(word)
    variants = [base + r"(?:e?s|ed|d|ing|ly|'s)?"]
    if word.endswith("e") and len(word) > 2:
        variants.append(re.escape(word[:-1]) + r"(?:ing|ed|ion)")
    if word.endswith("y") and len(word) > 2:
        variants.append(re.escape(word[:-1]) + r"(?:ies|ied|ily)")
    if len(word) > 3 and word[-1].lower() not in "aeiouwxy":
        variants.append(base + re.escape(word[-1]) + r"(?:ing|ed|er)")
    pattern = re.compile(r"\b(" + "|".join(variants) + r")\b", re.IGNORECASE)
    m = pattern.search(sentence)
    if not m:
        return None
    return sentence[: m.start()] + BLANK + sentence[m.end():], m.group(0)


def make_question(pool: pd.DataFrame, row: pd.Series, kind: str, rng: random.Random):
    """pool = 抽干擾選項用的完整單字庫（不含 row 自己）。"""
    q = {
        "word": row["word"], "meaning": row["meaning"], "pos": row["pos"],
        "example": row["example"], "example_zh": row["example_zh"],
        "note": row["note"], "tags": row["tags"], "tricky": bool(row["tricky"]),
        "kind": kind, "sub": "",
    }

    if kind in ("cloze", "spell"):
        blanked = make_blank(row["example"], row["word"])
        if not blanked:
            return None
        q["sentence"], q["surface"] = blanked
        if kind == "spell":
            q.update(prompt="把空格填回原本的單字（詞形不限）", answer=q["surface"])
            return q
        opts = pick_distractors(pool, row, "word", rng, 3)
        if len(opts) < 3:
            return None
        options = opts + [row["word"]]
        rng.shuffle(options)
        q.update(prompt="選出最適合填入空格的字", answer=row["word"], options=options)
        return q

    if kind == "en2zh":
        opts = pick_distractors(pool, row, "meaning", rng, 3)
        if len(opts) < 3:
            return None
        options = opts + [row["meaning"]]
        rng.shuffle(options)
        q.update(prompt=row["word"], sub="這個字的意思是？",
                 answer=row["meaning"], options=options)
        return q

    if kind == "zh2en":
        opts = pick_distractors(pool, row, "word", rng, 3)
        if len(opts) < 3:
            return None
        options = opts + [row["word"]]
        rng.shuffle(options)
        q.update(prompt=row["meaning"], sub="對應的英文單字是？",
                 answer=row["word"], options=options)
        return q
    return None


def filter_by_mode(df: pd.DataFrame, mode: str):
    if mode == "tricky":
        sub = df[df["tricky"]]
        if sub.empty:
            return df, "單字庫中沒有標記為「熟詞偏義」的字（請在 tags 或 note 欄寫上），已改用全部單字。"
        return sub.reset_index(drop=True), f"熟詞偏義特訓：共 {len(sub)} 個標記字。"
    if mode == "wrong":
        bank = set(load_progress()["wrong"].keys())
        sub = df[df["word"].isin(bank)]
        if sub.empty:
            return df, "錯題本目前是空的，先來一輪測驗吧！已改用全部單字。"
        return sub.reset_index(drop=True), f"錯題複習：錯題本共 {len(sub)} 個字。"
    return df, ""


# --------------------------------------------------------------------------
# 側邊欄
# --------------------------------------------------------------------------
def sidebar() -> dict:
    st.sidebar.header("⚙️ 設定")

    default_url = ""
    try:
        default_url = st.secrets.get("sheet_url", "")
    except Exception:
        pass

    source = st.sidebar.radio("單字庫來源",
                              ["Google Sheet 連結", "上傳檔案", "內建範例"],
                              index=0 if default_url else 2)

    df, err = None, None
    if source == "Google Sheet 連結":
        url = st.sidebar.text_input(
            "貼上公開的 Google Sheet 連結", value=default_url,
            placeholder="https://docs.google.com/spreadsheets/d/.../edit#gid=0",
            help="Sheet 需設為「知道連結的任何人皆可檢視」。一般編輯網址會自動轉成 CSV 匯出網址。",
        )
        if url.strip():
            try:
                df = load_from_url(url.strip())
            except Exception as e:  # noqa: BLE001
                err = explain_load_error(e, url)
        else:
            st.sidebar.caption("尚未填連結，先用內建範例。")
    elif source == "上傳檔案":
        up = st.sidebar.file_uploader("CSV 或 Excel", type=["csv", "xlsx", "xls", "xlsm"])
        if up is not None:
            try:
                df = load_from_bytes(up.getvalue(), up.name)
            except Exception as e:  # noqa: BLE001
                err = f"讀取失敗：{e}"

    if err:
        st.sidebar.error("單字庫讀取失敗，詳見主畫面。")
    if df is None:
        df, source_label = load_sample(), "內建範例"
    else:
        source_label = source

    if st.sidebar.button("🔄 重新載入單字庫", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    st.sidebar.divider()
    mode = st.sidebar.selectbox("練習模式", list(MODES.keys()), format_func=lambda k: MODES[k])
    n = st.sidebar.slider("每回題數", 5, 40, 12)
    kinds = st.sidebar.multiselect("題型", list(KIND_LABELS.keys()), default=DEFAULT_KINDS,
                                   format_func=lambda k: KIND_LABELS[k]) or DEFAULT_KINDS
    hint = st.sidebar.toggle("顯示提示（詞性／首字母）", value=False)

    st.sidebar.divider()
    prog = load_progress()
    today = prog["history"].get(dt.date.today().isoformat(), {"total": 0, "correct": 0})
    c1, c2, c3 = st.sidebar.columns(3)
    c1.metric("今日作答", today["total"])
    c2.metric("今日答對", today["correct"])
    c3.metric("連續天數", study_streak())
    st.sidebar.caption(f"單字庫 {len(df)} 字（{source_label}）｜錯題本 {len(prog['wrong'])} 字")
    st.sidebar.download_button(
        "⬇️ 下載學習紀錄 JSON",
        json.dumps(prog, ensure_ascii=False, indent=2).encode("utf-8"),
        file_name="progress.json", mime="application/json", use_container_width=True,
    )

    return {"df": df, "mode": mode, "n": n, "kinds": sorted(kinds), "hint": hint,
            "error": err, "source": source_label}


# --------------------------------------------------------------------------
# 測驗狀態
# --------------------------------------------------------------------------
def quiz_signature(cfg: dict) -> str:
    df = cfg["df"]
    fingerprint = f"{len(df)}:{'|'.join(df['word'].head(20).tolist())}"
    parts = [cfg["mode"], str(cfg["n"]), ",".join(cfg["kinds"]), fingerprint,
             str(st.session_state.get("salt", 0))]
    if cfg["mode"] == "daily":
        parts.append(dt.date.today().isoformat())
    return "|".join(parts)


def build_questions(cfg: dict, pool: pd.DataFrame, rng: random.Random) -> list:
    """從 pool 抽題；干擾選項一律取自完整單字庫，順便複習其他字。"""
    full = cfg["df"]
    order = list(pool.index)
    rng.shuffle(order)
    questions = []
    for idx in order:
        if len(questions) >= cfg["n"]:
            break
        row = pool.loc[idx]
        others = full[full["word"] != row["word"]]
        trial = list(cfg["kinds"])
        rng.shuffle(trial)
        for kind in trial:
            q = make_question(others, row, kind, rng)
            if q:
                questions.append(q)
                break
    return questions


def ensure_quiz(cfg: dict):
    sig = quiz_signature(cfg)
    if st.session_state.get("sig") == sig and st.session_state.get("questions") is not None:
        return st.session_state.questions, st.session_state.get("mode_note", "")

    pool, note = filter_by_mode(cfg["df"], cfg["mode"])
    # 每日模式：以日期為亂數種子 → 同一天不論重整幾次，題目都一樣
    rng = random.Random(f"{dt.date.today().isoformat()}|{sig}") if cfg["mode"] == "daily" \
        else random.Random()

    st.session_state.sig = sig
    st.session_state.questions = build_questions(cfg, pool, rng)
    st.session_state.answers = {}
    st.session_state.idx = 0
    st.session_state.mode_note = note
    return st.session_state.questions, note


def new_round() -> None:
    st.session_state.salt = st.session_state.get("salt", 0) + 1
    st.session_state.pop("sig", None)


def norm_answer(s) -> str:
    return str(s).strip().lower().strip(string.punctuation + " ")


def submit(idx: int, q: dict, user: str) -> None:
    ok = norm_answer(user) == norm_answer(q["answer"])
    if q["kind"] == "spell" and not ok:      # 拼字題也接受原形
        ok = norm_answer(user) == norm_answer(q["word"])
    st.session_state.answers[idx] = {"user": user, "ok": ok}
    record_result(q["word"], ok)


# --------------------------------------------------------------------------
# 畫面
# --------------------------------------------------------------------------
def render_details(q: dict) -> None:
    if q["example"]:
        sent = q["example"]
        if q.get("surface"):
            sent = sent.replace(q["surface"], f"**:orange[{q['surface']}]**")
        st.markdown(f"📖 {sent}")
    if q["example_zh"]:
        st.caption(q["example_zh"])
    if q["note"]:
        (st.warning if q["tricky"] else st.info)(
            ("⚠️ 熟詞偏義：" if q["tricky"] else "💡 ") + q["note"]
        )


def render_question(cfg: dict, q: dict, idx: int, total: int) -> None:
    st.progress(idx / total, text=f"第 {idx + 1} / {total} 題")

    pills = f"<span class='pill'>{KIND_LABELS[q['kind']]}</span>"
    if q["tricky"]:
        pills += "<span class='pill pill-tricky'>熟詞偏義</span>"
    if cfg["hint"] and q["pos"]:
        pills += f"<span class='pill'>{q['pos']}</span>"

    if q["kind"] in ("cloze", "spell"):
        body = f"{pills}<div class='q-sent'>{q['sentence']}</div><div class='q-meta'>{q['prompt']}</div>"
    else:
        body = f"{pills}<div class='q-prompt'>{q['prompt']}</div><div class='q-meta'>{q['sub']}</div>"
    st.markdown(f"<div class='q-card'>{body}</div>", unsafe_allow_html=True)

    answered = idx in st.session_state.answers

    if q["kind"] == "spell":
        if not answered:
            if cfg["hint"]:
                st.caption(f"提示：{q['answer'][0]} + {len(q['answer']) - 1} 個字母")
            with st.form(f"spell_{idx}"):
                text = st.text_input("你的答案", placeholder="輸入單字後按 Enter")
                if st.form_submit_button("送出", use_container_width=True) and text.strip():
                    submit(idx, q, text)
                    st.rerun()
    elif not answered:
        cols = st.columns(2)
        for i, opt in enumerate(q["options"]):
            if cols[i % 2].button(opt, key=f"opt_{idx}_{i}", use_container_width=True):
                submit(idx, q, opt)
                st.rerun()
    else:
        user = st.session_state.answers[idx]["user"]
        for opt in q["options"]:
            cls, mark = "opt-row", ""
            if opt == q["answer"]:
                cls, mark = cls + " opt-ok", "　✅"
            elif opt == user:
                cls, mark = cls + " opt-ng", "　❌ 你的選擇"
            st.markdown(f"<div class='{cls}'>{opt}{mark}</div>", unsafe_allow_html=True)

    if answered:
        res = st.session_state.answers[idx]
        if res["ok"]:
            st.success(f"答對了！**{q['word']}**（{q['pos'] or '—'}）{q['meaning']}")
        else:
            st.error(f"正解：**{q['answer']}**　—　**{q['word']}** {q['meaning']}")
        render_details(q)
        label = "下一題 ▶" if idx + 1 < total else "看結果 🎉"
        if st.button(label, type="primary", use_container_width=True, key=f"next_{idx}"):
            st.session_state.idx = idx + 1
            st.rerun()


def render_result(questions: list) -> None:
    answers = st.session_state.answers
    total = len(questions)
    correct = sum(1 for a in answers.values() if a["ok"])
    rate = correct / total if total else 0.0

    st.subheader("🎉 本回合完成")
    c1, c2, c3 = st.columns(3)
    c1.metric("題數", total)
    c2.metric("答對", correct)
    c3.metric("正確率", f"{rate:.0%}")
    st.progress(rate)

    wrong = [(i, q) for i, q in enumerate(questions) if not answers.get(i, {}).get("ok")]
    if wrong:
        st.markdown(f"#### 需要再看一次的 {len(wrong)} 個字")
        for i, q in wrong:
            with st.expander(f"❌ {q['word']} — {q['meaning']}"):
                st.caption(f"{KIND_LABELS[q['kind']]}　你的答案：{answers.get(i, {}).get('user', '—')}"
                           f"　／　正解：{q['answer']}")
                render_details(q)
        rows = [{"word": q["word"], "meaning": q["meaning"], "pos": q["pos"],
                 "your_answer": answers.get(i, {}).get("user", ""), "answer": q["answer"],
                 "example": q["example"], "note": q["note"]} for i, q in wrong]
        st.download_button(
            "⬇️ 下載錯題 CSV",
            pd.DataFrame(rows).to_csv(index=False).encode("utf-8-sig"),
            file_name=f"wrong_{dt.date.today().isoformat()}.csv",
            mime="text/csv", use_container_width=True,
        )
    else:
        st.balloons()
        st.success("全部答對，今天狀態很好！")

    c1, c2 = st.columns(2)
    if c1.button("🔁 再來一回（重抽題目）", type="primary", use_container_width=True):
        new_round()
        st.rerun()
    if c2.button("📕 立刻重練這回錯字", use_container_width=True, disabled=not wrong):
        st.session_state.questions = [q for _, q in wrong]
        st.session_state.answers = {}
        st.session_state.idx = 0
        st.rerun()


def main() -> None:
    st.title("📘 " + APP_TITLE)
    cfg = sidebar()
    if cfg["error"]:
        st.error(cfg["error"])
        st.caption("目前先用內建範例單字庫作答，修好後按側邊欄的「🔄 重新載入單字庫」即可切換。")
    questions, note = ensure_quiz(cfg)

    if note:
        st.info(note)
    if not questions:
        st.warning("產不出題目：單字庫太小（四選一至少需要 4 個字），"
                   "或所選題型（例句填空）需要 example 欄位有內容。")
        st.dataframe(cfg["df"].head(20), use_container_width=True)
        return

    st.caption(f"{MODES[cfg['mode']]}　·　{dt.date.today():%Y/%m/%d}　·　"
               "點左上角箭頭展開設定，可換單字庫與題型")

    idx = st.session_state.idx
    if idx >= len(questions):
        render_result(questions)
    else:
        render_question(cfg, questions[idx], idx, len(questions))
        st.divider()
        if st.button("🔀 換一批題目", use_container_width=True):
            new_round()
            st.rerun()


if __name__ == "__main__":
    main()
