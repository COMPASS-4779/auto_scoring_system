# -*- coding: utf-8 -*-
# =====================================================================
# auto_scoring_system  app.py  ―― 答案自動採点システム
#   バージョンと更新内容は CHANGELOG.md（先頭の版を画面ヘッダーに表示）
# =====================================================================
# 【画面構成】 📥 取込 ／ 📊 集計結果 ／ ⚙️ マスタ管理 ／ 📜 更新履歴
#
# 【取込の流れ】 ①種類 → ②生徒・科目 → ③種類ごとの設定 → ④写真 → ⑤AIで読み取る
#               → ⑥表で確認・修正 → ⑦記録（ここで初めて Drive / シート / メール）
#  ・📄 テキスト  : ページ番号 → 目次マスタを逆引きして 章/節/節タイトル を埋める
#  ・📝 確認テスト: 理解度確認テスト・復習テスト。テストのタイトル(M列)と出題元の
#                   テキスト名(D)/章(F)/節(G) を読み取る（無ければ単元見出しで代替）
#  ・🎓 過去問    : 実施日・学校名・学部・科目・方式・年度・得点・満点を読み取り、
#                   「過去問」タブへ記録。スケジュール管理のテスト結果に表示される
#
# 【スプレッドシート】
#  ・1枚目（結果）: 日時, 生徒名, 科目, テキスト名, ページ, 章, 節, 問題番号, 写真リンク,
#                   総問題数(小問数), 節タイトル, （空）, テストのタイトル        … A:M
#  ・過去問       : 登録日時, 生徒名, 実施日, 学校名, 学部, 科目, 方式, 年度, 得点, 満点,
#                   写真リンク, 記録ID                                          … A:L
#  ・目次マスタ / 生徒名簿 / 科目マスタ
#  ※ B列＝生徒名はスケジュール管理側の照合キー（ログインID または登録氏名）
#
# 【Secrets】 必須: GEMINI_API_KEY / SENDER_EMAIL / APP_PASSWORD / GOOGLE_TOKEN_JSON
#            任意: SPREADSHEET_ID / PARENT_FOLDER_ID / NOTIFICATION_EMAIL / STUDENT_SEED
# =====================================================================
import streamlit as st
import os
import io
import csv
import zipfile
import datetime
import json
import re
import statistics
import tempfile
import time
import uuid
import fitz  # PyMuPDF
import pandas as pd
from PIL import Image
try:
    import pillow_heif
    pillow_heif.register_heif_opener()   # iPhoneのHEIC/HEIFを読めるように登録
    _HEIC_OK = True
except Exception:
    _HEIC_OK = False
import smtplib
from email.mime.text import MIMEText
from email.header import Header
from email.utils import formatdate

from google import genai
from google.genai import types
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# ==========================================
# 設定情報
# ==========================================
APP_DIR = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH = os.path.join(APP_DIR, "assets", "compass_logo.png")      # COMPASS online ロゴ
MARK_PATH = os.path.join(APP_DIR, "assets", "compass_mark.png")      # ロゴの紋章部分（ファビコン）
CHANGELOG_PATH = os.path.join(APP_DIR, "CHANGELOG.md")               # 更新履歴（バージョンの出どころ）

st.set_page_config(page_title="答案自動採点システム",
                   page_icon=(MARK_PATH if os.path.exists(MARK_PATH) else "📝"), layout="wide")

missing_keys = []
for key in ["GEMINI_API_KEY", "SENDER_EMAIL", "APP_PASSWORD", "GOOGLE_TOKEN_JSON"]:
    if key not in st.secrets:
        missing_keys.append(key)
if missing_keys:
    st.error(f"🚨 Secrets に以下の鍵が見つかりません: {', '.join(missing_keys)}")
    st.stop()

try:
    GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
    SENDER_EMAIL = st.secrets["SENDER_EMAIL"]
    APP_PASSWORD = st.secrets["APP_PASSWORD"]
    GOOGLE_TOKEN_DICT = json.loads(st.secrets["GOOGLE_TOKEN_JSON"])
except json.JSONDecodeError as e:
    st.error(f"🚨 token.json の読み込みに失敗: {e}")
    st.stop()
except Exception as e:
    st.error(f"🚨 予期せぬエラー: {e}")
    st.stop()

# --- 以下の ID / メールアドレス / 生徒名は Secrets に置くことを推奨 ---
#     Streamlit Cloud の Settings > Secrets に下記を追加すると、そちらが優先されます。
#       SPREADSHEET_ID   = "..."      結果を書き込むスプレッドシートID
#       PARENT_FOLDER_ID = "..."      生徒フォルダを作る Drive の親フォルダID
#       NOTIFICATION_EMAIL = "..."    進捗通知の宛先
#       STUDENT_SEED     = "生徒A,生徒B"  「生徒名簿」タブが空のときだけ使う初期値
#     Secrets への移行が済んだら、下の第2引数（フォールバック値）を "" にしてください。
def _cfg(key, default=""):
    try:
        v = st.secrets.get(key, "")
    except Exception:
        v = ""
    return str(v).strip() or default

SPREADSHEET_ID = _cfg("SPREADSHEET_ID", "1B8BKKY8SfR-V3ysirsNG6fqlrVzXqPBF_AdjFDc5fCc")
PARENT_FOLDER_ID = _cfg("PARENT_FOLDER_ID", "1DS7anMs-ruhTtVxZNqsVhZSbeQFCww_2")
MASTER_DIR = "master_texts"
NOTIFICATION_EMAIL = _cfg("NOTIFICATION_EMAIL", SENDER_EMAIL)
MASTER_TAB = "目次マスタ"  # [統合] テキスト目次マスタを保存するタブ名
MASTER_HEADER = ["テキスト名", "章", "節", "節タイトル", "開始ページ", "終了ページ"]
STUDENT_TAB = "生徒名簿"   # [統合] 生徒名を保存するタブ
SUBJECT_TAB = "科目マスタ"  # [統合] 科目を保存するタブ
DEFAULT_SUBJECTS = ["国語", "数学", "英語", "英文法", "古文", "理科", "社会"]

# 生徒名はソースに書かない。通常はスプレッドシートの「生徒名簿」タブから読み込む。
# STUDENT_SEED（Secrets・カンマ区切り）は、名簿タブがまだ空のときの初期値としてのみ使う。
STUDENT_LIST = [x.strip() for x in _cfg("STUDENT_SEED").split(",") if x.strip()]
os.makedirs(MASTER_DIR, exist_ok=True)


# ==========================================
# 共通ユーティリティ
# ==========================================
def _extract_json_array(text):
    """Gemini の応答から最初の JSON 配列を取り出して list を返す（失敗時は []）。
       ```json フェンス・前後の説明文・文中の別の [ ] に強い（括弧の対応を数える）。"""
    if not text:
        return []
    s = str(text)
    m = re.search(r"```(?:json)?\s*(.+?)```", s, re.DOTALL)   # コードフェンスがあれば中身を優先
    if m:
        s = m.group(1)
    for start, ch in enumerate(s):
        if ch != "[":
            continue
        depth, in_str, esc = 0, False, False
        for i in range(start, len(s)):
            c = s[i]
            if in_str:
                if esc:      esc = False
                elif c == "\\": esc = True
                elif c == '"':   in_str = False
                continue
            if c == '"':
                in_str = True
            elif c == "[":
                depth += 1
            elif c == "]":
                depth -= 1
                if depth == 0:                      # 対応する ] まで来た
                    try:
                        v = json.loads(s[start:i + 1])
                    except Exception:
                        break                       # 壊れていたら次の [ を試す
                    return v if isinstance(v, list) else []
    return []


# ==========================================
# [統合] テキスト目次マスタ（Google Sheet 永続化）
# ==========================================
def _to_int(v):
    if v is None: return None
    m = re.search(r"\d+", str(v).translate(str.maketrans("０１２３４５６７８９", "0123456789")))
    return int(m.group()) if m else None

def _text_name_from_filename(fn):
    """ファイル名からテキスト名を推定（末尾の _text_章節リスト / _章節リスト / 拡張子を除去）。"""
    base = re.sub(r"\.[^.]+$", "", str(fn))
    base = re.sub(r"[_\s]*(text)?[_\s]*章節リスト$", "", base, flags=re.IGNORECASE)
    base = re.sub(r"[_\s]*目次$", "", base)
    return base.strip()

def _page_from_filename(name):
    """ファイル名から『ページ番号』を推定（p45 / page45 / 45ページ / 頁 等の明示マーカーがある時のみ）。"""
    s = str(name).translate(str.maketrans("０１２３４５６７８９", "0123456789"))
    m = re.search(r"(?:p\.?|page|ｐ|ページ|頁)\s*([0-9]{1,3})", s, re.IGNORECASE)
    if m:
        return m.group(1)
    m = re.search(r"([0-9]{1,3})\s*(?:ページ|頁)", s)
    if m:
        return m.group(1)
    return ""

def _sheets(creds):
    return build('sheets', 'v4', credentials=creds)

def ensure_master_tab(creds):
    """目次マスタ タブが無ければ作成し、ヘッダー行を入れる。"""
    svc = _sheets(creds)
    meta = svc.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    titles = [s['properties']['title'] for s in meta.get('sheets', [])]
    if MASTER_TAB not in titles:
        svc.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"requests": [{"addSheet": {"properties": {"title": MASTER_TAB}}}]}
        ).execute()
        svc.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID, range=f"{MASTER_TAB}!A1",
            valueInputOption="RAW", body={"values": [MASTER_HEADER]}
        ).execute()

def load_master_index(creds):
    """目次マスタ タブを読み込み {テキスト名: [ {chapter,section,title,start,end} ]} を返す。"""
    index = {}
    try:
        ensure_master_tab(creds)
        res = _sheets(creds).spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID, range=f"{MASTER_TAB}!A:F").execute()
        rows = res.get('values', [])
        if len(rows) < 2:
            return index
        header = rows[0]
        def idx_of(*names):
            for n in names:
                if n in header:
                    return header.index(n)
            return None
        ci = {k: idx_of(*v) for k, v in {
            "text": ["テキスト名", "教材名"], "chapter": ["章"], "section": ["節"],
            "title": ["節タイトル", "タイトル"], "start": ["開始ページ", "ページ"], "end": ["終了ページ"],
        }.items()}
        for r in rows[1:]:
            def cell(key):
                i = ci.get(key)
                return (r[i].strip() if (i is not None and i < len(r) and r[i] is not None) else "")
            t = cell("text")
            if not t:
                continue
            index.setdefault(t, []).append({
                "chapter": cell("chapter"), "section": cell("section"), "title": cell("title"),
                "start": _to_int(cell("start")), "end": _to_int(cell("end")),
            })
        # 終了ページ補完
        for t, lst in index.items():
            lst.sort(key=lambda r: (r["start"] is None, r["start"] or 0))
            for i, r in enumerate(lst):
                if r["start"] is not None and r["end"] is None:
                    nxt = lst[i + 1]["start"] if i + 1 < len(lst) else None
                    r["end"] = (nxt - 1) if nxt else r["start"]
    except Exception as e:
        print(f"[統合] load_master_index error: {e}")
    return index

def register_master_csv(creds, file_bytes, default_text_name=""):
    """目次CSVを取り込み、目次マスタ タブへ反映（同名テキストは置換）。
       テキスト名列が無いCSV（章節リスト等）は default_text_name（ファイル名由来）を使う。
       列は 章/節/節タイトル/開始ページ/終了ページ を自動判別（開始問題番号などは無視）。"""
    text = file_bytes.decode("utf-8-sig", errors="replace")
    reader = csv.DictReader(io.StringIO(text))
    new_rows, new_names = [], set()
    for row in reader:
        def col(*names):
            for n in names:
                if n in row and row[n] is not None:
                    return str(row[n]).strip()
            return ""
        tname = col("テキスト名", "教材名") or (default_text_name or "").strip()
        if not tname:
            continue
        chapter = col("章", "編")
        section = col("節", "項")
        title = col("節タイトル", "タイトル", "見出し")
        start = col("開始ページ", "ページ", "頁")
        end = col("終了ページ", "終了")
        # 章/節 の2階層のみ（節タイトル列なし）の場合、節を節タイトルにも入れて見やすく
        if not title and section:
            title = section
        new_names.add(tname)
        new_rows.append([tname, chapter, section, title, start, end])
    if not new_rows:
        return 0, 0
    # 既存を読み、同名テキストを除外して結合
    ensure_master_tab(creds)
    res = _sheets(creds).spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=f"{MASTER_TAB}!A:F").execute()
    existing = res.get('values', [])
    body_rows = []
    if existing and existing[0] == MASTER_HEADER:
        for r in existing[1:]:
            if r and (r[0].strip() not in new_names):
                body_rows.append(r)
    body_rows.extend(new_rows)
    # 全書き換え
    _sheets(creds).spreadsheets().values().clear(
        spreadsheetId=SPREADSHEET_ID, range=f"{MASTER_TAB}!A:F").execute()
    _sheets(creds).spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID, range=f"{MASTER_TAB}!A1",
        valueInputOption="RAW", body={"values": [MASTER_HEADER] + body_rows}).execute()
    return len(new_names), len(new_rows)

def lookup_section(index, text_name, page):
    """テキスト名＋ページ → (章, 節, 節タイトル)。無ければ ('','','')。"""
    p = _to_int(page)
    if p is None:
        return ("", "", "")
    cands = index.get(text_name)
    if not cands:
        for k, v in index.items():
            if text_name and (text_name in k or k in text_name):
                cands = v
                break
    for r in (cands or []):
        if r["start"] is not None and r["end"] is not None and r["start"] <= p <= r["end"]:
            return (r["chapter"], r["section"], r["title"])
    return ("", "", "")


# ==========================================
# [統合] 生徒名・科目（単一列タブ）の管理
# ==========================================
def ensure_list_tab(creds, tab, header_label, seed=None):
    svc = _sheets(creds)
    meta = svc.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    titles = [s['properties']['title'] for s in meta.get('sheets', [])]
    if tab not in titles:
        svc.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"requests": [{"addSheet": {"properties": {"title": tab}}}]}).execute()
        vals = [[header_label]] + [[s] for s in (seed or [])]
        svc.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID, range=f"{tab}!A1",
            valueInputOption="RAW", body={"values": vals}).execute()

def load_list(creds, tab, header_label, seed=None):
    try:
        ensure_list_tab(creds, tab, header_label, seed)
        res = _sheets(creds).spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID, range=f"{tab}!A:A").execute()
        rows = res.get('values', [])
        out = []
        for r in rows[1:]:
            v = (r[0].strip() if r and r[0] is not None else "")
            if v and v not in out:
                out.append(v)
        return out
    except Exception as e:
        print(f"[統合] load_list error ({tab}): {e}")
        return list(seed or [])

def add_list_item(creds, tab, header_label, value):
    value = (value or "").strip()
    if not value:
        return False
    cur = load_list(creds, tab, header_label)
    if value in cur:
        return False
    _sheets(creds).spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID, range=f"{tab}!A:A",
        valueInputOption="RAW", body={"values": [[value]]}).execute()
    return True

def remove_list_item(creds, tab, header_label, value):
    cur = load_list(creds, tab, header_label)
    new = [x for x in cur if x != value]
    _sheets(creds).spreadsheets().values().clear(
        spreadsheetId=SPREADSHEET_ID, range=f"{tab}!A:A").execute()
    _sheets(creds).spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID, range=f"{tab}!A1",
        valueInputOption="RAW", body={"values": [[header_label]] + [[x] for x in new]}).execute()
    return True




# ==========================================
# [統合] PDF目次解析（逆引きアプリのPython移植 / PyMuPDF）
# ==========================================
ZEN = "０１２３４５６７８９"
KAN = "〇一二三四五六七八九"

def _half(s): return str(s).translate(str.maketrans(ZEN, "0123456789"))
def _has_jp(s): return bool(re.search(r"[ぁ-んァ-ヶ一-龥]", s or ""))
def _parse_num(s):
    t = re.sub(r"[^\d]", "", _half(str(s)))
    if not t: return None
    n = int(t)
    return n if 1 <= n <= 1999 else None
def _kan2num(s):
    s = _half(s)
    if s.isdigit(): return int(s)
    if s == "十": return 10
    m = re.match(r"^(.?)十(.?)$", s)
    if m:
        t = KAN.find(m.group(1)) if m.group(1) else 1
        o = KAN.find(m.group(2)) if m.group(2) else 0
        if t < 0: t = 1
        if o < 0: o = 0
        return t*10+o
    i = KAN.find(s)
    return i if i > 0 else None

def _page_spans(page):
    out = []
    for b in page.get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            for sp in l.get("spans", []):
                t = sp["text"]
                if t and t.strip():
                    x0, y0, x1, y1 = sp["bbox"]
                    out.append({"s": t, "x": x0, "y": y0, "h": sp["size"]})
    return out

def _group_lines(items, tol=6):
    lines = []
    for it in sorted(items, key=lambda a: (a["y"], a["x"])):
        g = next((L for L in lines if abs(L["y"]-it["y"]) <= tol), None)
        if not g:
            g = {"y": it["y"], "parts": []}; lines.append(g)
        g["parts"].append(it)
    for L in lines:
        L["parts"].sort(key=lambda a: a["x"])
    return lines

def _clean_name(s):
    s = re.sub(r"[.．。・…‥､、，]+", "", str(s))
    s = re.sub(r"[0-9０-９]+\s*$", "", s)
    return re.sub(r"\s|　", "", s).strip()

# ---------- 埋め込み目次 ----------
def _is_junk_outline(toc):
    titles = [t[1].strip() for t in toc]
    if not titles: return True
    junk = sum(1 for t in titles if re.match(r"^p(age|\.)?\s*\d+$", t, re.I) or t.isdigit() or t == "")
    return junk/len(titles) >= 0.6

def _rows_from_outline(doc, toc):
    out = []
    has_child = any(t[0] >= 2 for t in toc)
    cur = ""
    for level, title, page in toc:
        title = (title or "").strip()
        if has_child and level == 1:
            cur = title
            out.append({"chapter": title, "section": "（章扉）", "title": title, "start": page})
        elif has_child:
            out.append({"chapter": cur, "section": title, "title": title, "start": page})
        else:
            out.append({"chapter": "", "section": title, "title": title, "start": page})
    return [r for r in out if r["start"]]

# ---------- 目次ページ解析 ----------
PART_RE = re.compile(r"第\s*([0-9０-９一二三四五六七八九十]+)\s*[部章編節]")
CHAP_RE = re.compile(r"第\s*([0-9０-９一二三四五六七八九十]+)\s*章")

def _right_col(items, w):
    nums = []
    for L in _group_lines([it for it in items if it["x"] >= w*0.8 and re.search(r"[\d０-９]", it["s"])], 6):
        v = _parse_num("".join(p["s"] for p in L["parts"]))
        if v is not None:
            nums.append({"y": L["y"], "x": min(p["x"] for p in L["parts"]), "val": v})
    return nums

def _toc_column(nums):
    if len(nums) < 6: return None
    best = []
    for a in nums:
        g = [b for b in nums if abs(b["x"]-a["x"]) <= 12]
        if len(g) > len(best): best = g
    if len(best) < 6: return None
    vals = [n["val"] for n in sorted(best, key=lambda a: a["y"])]
    if len(set(vals)) < 5: return None
    asc = sum(1 for i in range(1, len(vals)) if vals[i] >= vals[i-1])
    if asc/(len(vals)-1) < 0.6: return None
    return best

def _overview_topics(pages):
    for items, w in pages[:15]:
        nums = [{"y": L["y"], "val": _parse_num("".join(p["s"] for p in L["parts"]))}
                for L in _group_lines([it for it in items if it["x"] >= w*0.66 and re.search(r"[\d０-９]", it["s"])], 6)]
        nums = [n for n in nums if n["val"] is not None]
        if len(nums) < 3: continue
        lefts = [{"y": L["y"], "raw": "".join(p["s"] for p in L["parts"])}
                 for L in _group_lines([it for it in items if it["x"] < w*0.62], 6)]
        lefts = [l for l in lefts if _has_jp(l["raw"])]
        entries = []
        for l in lefts:
            best, bd = None, 99
            for n in nums:
                d = abs(n["y"]-l["y"])
                if d <= 20 and d < bd: bd, best = d, n
            entries.append({"raw": l["raw"], "page": best["val"] if best else None})
        if len(entries) < 3: continue
        pgs = [e["page"] for e in entries if e["page"] is not None]
        gaps = sorted(abs(pgs[i]-pgs[i-1]) for i in range(1, len(pgs)))
        if not gaps or gaps[len(gaps)//2] < 4: continue
        cur, topics = "第1部", []
        for k, e in enumerate(entries):
            m = PART_RE.search(e["raw"])
            if m:
                n = _kan2num(m.group(1))
                if n is not None: cur = f"第{n}部"
            nm = _clean_name(PART_RE.sub("", e["raw"]))
            if len(nm) < 2: continue
            if re.search(r"[：:]", e["raw"]) or re.search(r"解説|著者|編集|まえがき", nm): continue
            if not re.search(r"[ぁ-んァ-ヶ]", nm) and not re.search(r"[0-9０-９]", nm) and len(nm) <= 3: continue
            if e["page"] is not None and k+1 < len(entries) and entries[k+1]["page"] == e["page"]: continue
            topics.append({"part": cur, "name": nm})
        if len(topics) >= 2:
            return topics
    return None

def _interpolate(rows, key="page"):
    known = [i for i, r in enumerate(rows) if r[key] is not None]
    if not known: return
    for k in range(known[0]): rows[k][key] = max(1, rows[known[0]][key]-(known[0]-k))
    last = known[-1]
    for k in range(last+1, len(rows)): rows[k][key] = rows[last][key]+(k-last)
    for a in range(len(known)-1):
        i, j = known[a], known[a+1]
        pi, pj = rows[i][key], rows[j][key]
        for k in range(i+1, j):
            rows[k][key] = round(pi+(pj-pi)*(k-i)/(j-i))

TOP_RE = re.compile(r"第\s*[0-9０-９一二三四五六七八九十]+\s*[編部]")
SUB_RE = re.compile(r"^第\s*([0-9０-９一二三四五六七八九十]+)\s*[章節]")
SPECIAL_TOP = re.compile(r"(特集|巻末特集|付録|総合問題|序章|終章|巻頭)")
BULLET_RE2 = re.compile(r"^\s*([❶-❿①-⑳]|\([0-9０-９]+\)|[0-9０-９]+[\.．])\s*")
_GARBAGE = set("国醒ヨ駈田四■誼遍囲團圖□〇・·-—|ー")

def _strip_box(s):
    s = s.strip()
    i = 0
    while i < len(s) and (s[i] in _GARBAGE or s[i] in "0123456789０１２３４５６７８９ \t　"):
        i += 1
    return s[i:].strip()

def _collect_toc_lines(pages):
    chap_re3 = re.compile(r"第\s*[0-9０-９一二三四五六七八九十]+\s*[部章編節]")
    out = []
    for items, w in pages:
        nums = _toc_column(_right_col(items, w))
        if not nums: continue
        rec = []
        for L in _group_lines([it for it in items if it["x"] < w*0.66], 6):
            title = re.sub(r"\s+", " ", "".join(p["s"] for p in L["parts"])).strip()
            if not _has_jp(title) or len(title) < 2: continue
            rec.append({"y": L["y"], "title": title,
                        "size": max(p["h"] for p in L["parts"]),
                        "x": min(p["x"] for p in L["parts"])})
        if not rec: continue
        lens = sorted(len(r["title"]) for r in rec)
        if lens[len(lens)//2] > 30 or sum(1 for r in rec if len(r["title"]) <= 30) < max(4, len(rec)*0.5):
            continue  # 散文ばかりの本文/解答ページを除外
        med_x = sorted(r["x"] for r in rec)[len(rec)//2]
        paired = 0
        for r in rec:
            best, bd = None, 99
            for n in nums:
                d = abs(n["y"]-r["y"])
                if d <= 16 and d < bd: bd, best = d, n
            r["page"] = best["val"] if best else None
            if best: paired += 1
            r["is_top"] = r["x"] <= med_x - 8           # 左端の見出し＝上位区分（編/部/特集 等）
            r["isChap"] = bool(chap_re3.search(r["title"]))
        # 「見出し＋ページ番号」が大半でなければ目次ページでない（解答/本文の誤検出を除外）
        if paired < max(5, len(rec)*0.45):
            continue
        out.extend(rec)
    return out

def _hier_map(items):
    rows = []
    hen, cur_name, special = 1, "", None
    chap_no = 0          # 書籍の章番号（編配下、連番で欠番OCRを補完）
    item_no = 0          # 特集等の項目番号
    cur_chap = None
    CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"
    chapline = re.compile(r"^第\s*([0-9０-９一二三四五六七八九十]*)\s*[章節]")
    for it in items:
        t = it["title"].strip()
        if it.get("is_top"):
            if "巻末" in t:
                special = "巻末特集"; cur_chap = None; item_no = 0; continue
            if ("トレーニング" in t) or ("特集" in t):
                special = "特集"; cur_chap = None; item_no = 0; continue
            name = _strip_box(t)
            if name: cur_name = name
            continue
        clean = _strip_box(t)
        if len(clean) < 1:
            continue
        chap = special if special else ("第%d編" % hen + ((" " + cur_name) if cur_name else ""))
        if chap != cur_chap:
            cur_chap = chap; item_no = 0
        m = chapline.match(t)
        if (special is None) and m:
            n = _kan2num(m.group(1)) if m.group(1) else None
            chap_no = n if (n is not None) else chap_no + 1
            rows.append({"chapter": chap, "section": "第%d章" % chap_no,
                         "title": t[m.end():].strip(), "start": it["page"]})
            continue
        if re.match(r"^編[末未]問題", t):
            rows.append({"chapter": chap, "section": "編末問題", "title": "", "start": it["page"]})
            if special is None:
                hen += 1; cur_name = ""
            continue
        # 特集/巻末などの項目 → 節は連番ラベル、節タイトルは名称（先頭の丸記号は除去）
        item_no += 1
        sec = CIRCLED[item_no-1] if item_no <= len(CIRCLED) else str(item_no)
        sectitle = re.sub(r"^[〇◎●○◯❶-❿①-⑳・\s]+", "", clean).strip() or clean
        rows.append({"chapter": chap, "section": sec, "title": sectitle, "start": it["page"]})
    _interpolate(rows, "start")
    return rows

def _rows_from_toc(pages):
    all_items = _collect_toc_lines(pages)
    if len(all_items) < 3: return []
    has_top = any(it.get("is_top") for it in all_items)
    has_sub = any(SUB_RE.match(it["title"]) or BULLET_RE2.match(it["title"]) for it in all_items)
    if has_top and has_sub:
        return _hier_map(all_items)
    # ---- 従来方式（概観トピック名 + ● + 区分） ----
    last = 0
    for r in all_items:
        if r["page"] is None: continue
        if last <= r["page"] <= last+50: last = r["page"]
        else: r["page"] = None
    _interpolate(all_items)
    ov = _overview_topics(pages)
    cur_part, topic, expl, bullet = "第1部", 0, None, False
    rows = []
    bullet_re = re.compile(r"^\s*[●○◯◆■▼▶・]\s*(.+)$")
    for r in all_items:
        t = r["title"]
        m = PART_RE.search(t)
        if m:
            n = _kan2num(m.group(1))
            if n is not None: cur_part = f"第{n}部"
        hm = bullet_re.match(t)
        if hm:
            if not bullet and cur_part == "第1部": cur_part, bullet = "第2部", True
            nm = _clean_name(hm.group(1))
            if len(nm) >= 2: expl = nm
            continue
        sec = expl or ((ov[topic]["name"] if ov and topic < len(ov) else f"区分{topic+1}"))
        rows.append({"chapter": cur_part, "section": sec, "title": t, "start": r["page"]})
        if re.search(r"演習題|解答", t): topic += 1; expl = None
    return rows

# ---------- 本文見出し走査 ----------
def _rows_from_heading_scan(pages):
    allh = []
    pinfo = []
    for items, w, h in pages:
        sizes = [it["h"] for it in items]
        allh += sizes
        pinfo.append({"items": items, "w": w, "h": h, "n": len(items),
                      "full": "".join(it["s"] for it in items),
                      "maxH": max(sizes) if sizes else 0})
    gmed = statistics.median(allh) if allh else 12
    sec_h = gmed*1.6
    def big_line(p):
        if p["maxH"] < sec_h: return None
        cand = [it for it in p["items"] if it["h"] >= p["maxH"]*0.9 and it["y"] <= p["h"]*0.58]
        if not cand: return None
        cand.sort(key=lambda a: (a["y"], a["x"]))
        return re.sub(r"\s+", " ", "".join(c["s"] for c in cand)).strip() or None
    def is_toc(p): return bool(re.search(r"CONTENTS|目次", p["full"])) or len(re.findall(r"第\s*[0-9０-９一二三四五六七八九十]+\s*章", p["full"])) >= 3
    def bad(s):
        if not s or len(s) < 4: return True
        if len(re.findall(r"[ぁ-んァ-ヶ一-龥]", s)) < 3: return True
        if re.match(r"^(図|表|囲|團|圖|E\s)", s): return True
        if re.search(r"CONTENTS|目次", s): return True
        return False
    def simkey(s): return re.sub(r"[「」『』（）()【】\[\]:：・.,。、!！?？~〜\-Ff]", "", re.sub(r"\s|　", "", s).lower())
    cur_chap_n, cur_chap, found, prev = 0, "", False, None
    rows = []
    for i, p in enumerate(pinfo):
        if is_toc(p): continue
        cm = CHAP_RE.search(p["full"])
        if cm and p["n"] <= 18:
            n = _kan2num(cm.group(1))
            if n is not None and n > cur_chap_n:
                cur_chap_n = n; found = True
                big = big_line(p); nm = big if (big and not CHAP_RE.search(big)) else ""
                cur_chap = f"第{n}章" + (" "+nm if nm else ""); prev = cur_chap
                rows.append({"chapter": cur_chap, "section": "（章扉）", "title": nm or cur_chap, "start": i+1, "isChap": True})
                continue
        big = big_line(p)
        if not big or bad(big): continue
        if prev and simkey(big) == simkey(prev): continue
        prev = big
        rows.append({"chapter": cur_chap, "section": big, "title": big, "start": i+1})
    out = rows
    if found:
        fc = next((k for k, r in enumerate(rows) if r.get("isChap")), 0)
        out = [r for k, r in enumerate(rows) if r.get("isChap") or k > fc]
    return [{"chapter": r["chapter"], "section": r["section"], "title": r["title"], "start": r["start"]} for r in out]

def _fill_ranges(rows, max_page):
    rows.sort(key=lambda r: (r["start"] is None, r["start"] or 0))
    for i, r in enumerate(rows):
        if r["start"] is None: continue
        if r.get("end") in (None, ""):
            r["end"] = (rows[i+1]["start"]-1) if (i+1 < len(rows) and rows[i+1]["start"] is not None) else (max_page or r["start"])
    return rows

def _garbled_ratio(rows):
    if not rows: return 0.0
    bad = 0
    for r in rows:
        s = (r.get("section") or "") + (r.get("title") or "")
        jp = len(re.findall(r"[ぁ-んァ-ヶ一-龥]", s))
        sym = len(re.findall(r"[\u25a0\u25a1\u3010\u3011\uff5c|\uff1d=:：；;、。\[\]「」『』]", s))
        if jp < 2 or sym > jp:
            bad += 1
    return bad / len(rows)

def analyze_pdf(doc, text_name):
    pages_simple = []
    pages_full = []
    for i in range(doc.page_count):
        pg = doc.load_page(i)
        items = _page_spans(pg)
        w, h = pg.rect.width, pg.rect.height
        pages_simple.append((items, w))
        pages_full.append((items, w, h))
    rows, method, pdf_mode = [], "", True
    toc = doc.get_toc()
    if toc and not _is_junk_outline(toc):
        rows = _rows_from_outline(doc, toc); method = "埋め込み目次"
    if not rows:
        rows = _rows_from_toc(pages_simple)
        if rows: method, pdf_mode = "目次ページ解析（印刷ページ）", False
    if not rows:
        rows = _rows_from_heading_scan(pages_full); method = "本文走査（見出し推定）"
        empty_chap = sum(1 for r in rows if not (r.get("chapter") or "").strip()) / max(1, len(rows))
        if _garbled_ratio(rows) > 0.45 or empty_chap > 0.8:   # フォント破損で文字化け/章不明→自動解析不可
            rows = []
            method = "解析不可（目次・見出しが文字化け）：CSV登録または手入力をご利用ください"
    _fill_ranges(rows, doc.page_count if pdf_mode else None)
    for r in rows:
        r["text"] = text_name
    return rows, method

# ==========================================
# [統合] Gemini画像解析（文字化けPDFの目次を画像から読む・高精度）
# ==========================================
def analyze_pdf_gemini(doc, text_name, api_key):
    client = genai.Client(api_key=api_key)
    model = get_best_model(client)
    toc_pages = []
    for i in range(min(25, doc.page_count)):
        pg = doc.load_page(i)
        if _toc_column(_right_col(_page_spans(pg), pg.rect.width)):
            toc_pages.append(i)
    if not toc_pages:
        toc_pages = [min(3, doc.page_count - 1)]
    prompt = (
        "これは学習参考書の目次ページの画像です。階層は『編または部 ＞ 章 ＞ 項目（題名）』です。\n"
        "見出し行を JSON 配列で返してください。各要素は "
        "{\"chapter\": \"\", \"section\": \"\", \"title\": \"\", \"page\": 0}。\n"
        "・chapter = 最上位区分（例: 第1編 力と運動 / 特集 / 巻末特集）。同じ編の各行に同じ chapter を入れる。\n"
        "・section = 章レベルのラベル（例: 第1章。編末問題は \"編末問題\"。特集の項目は ①②③ 等）。\n"
        "・title = 章/項目の題名（編末問題は空文字 \"\"）。\n"
        "・page = その行の右にある開始ページ番号（半角整数）。\n"
        "2段組のときは左列→右列の順。JSON配列だけを出力。"
    )
    rows = []
    for i in toc_pages[:4]:
        pix = doc.load_page(i).get_pixmap(dpi=200)
        tmp = os.path.join(tempfile.gettempdir(), f"toc_{uuid.uuid4().hex}.png")
        pix.save(tmp)
        try:
            af = client.files.upload(file=tmp)
            while af.state.name == 'PROCESSING':
                time.sleep(1); af = client.files.get(name=af.name)
            resp = client.models.generate_content(model=model, contents=[af, prompt])
            _arr = _extract_json_array(resp.text)
            if _arr:
                for o in _arr:
                    rows.append({"chapter": str(o.get("chapter", "")).strip(),
                                 "section": str(o.get("section", "")).strip(),
                                 "title": str(o.get("title", "")).strip(),
                                 "start": _to_int(o.get("page"))})
        finally:
            try: os.remove(tmp)
            except: pass
    seen, uniq = set(), []
    for r in rows:
        k = (r["chapter"], r["section"], r["title"], r["start"])
        if k in seen: continue
        seen.add(k); uniq.append(r)
    _fill_ranges(uniq, None)
    for r in uniq:
        r["text"] = text_name
    return uniq, "AI画像解析（Gemini）"


# ==========================================
# メール・Drive・結果Sheets
# ==========================================
def send_notification_email_plan_b(subject, body):
    try:
        msg = MIMEText(body, "plain", "utf-8")
        msg['Subject'] = Header(subject, "utf-8")
        msg['From'] = SENDER_EMAIL
        msg['To'] = NOTIFICATION_EMAIL
        msg['Date'] = formatdate(localtime=True)
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(SENDER_EMAIL, APP_PASSWORD)
        server.send_message(msg)
        server.quit()
    except Exception as e:
        print(f"メール送信失敗: {e}")

def ensure_drive_folder(student_name, creds):
    """親フォルダ(PARENT_FOLDER_ID)配下に同名フォルダが無ければ新規作成。(folder_id, 新規作成したか) を返す。"""
    service = build('drive', 'v3', credentials=creds)
    safe = str(student_name).replace("'", "\'")
    query = f"'{PARENT_FOLDER_ID}' in parents and name = '{safe}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    folders = service.files().list(q=query, fields="files(id)").execute().get('files', [])
    if folders:
        return folders[0]['id'], False
    created = service.files().create(
        body={'name': str(student_name), 'mimeType': 'application/vnd.google-apps.folder', 'parents': [PARENT_FOLDER_ID]},
        fields='id').execute()
    return created.get('id'), True

DONE_FOLDER = "実施済答案"      # 生徒フォルダの下に作る、答案写真の保存先


def _find_or_create_folder(service, name, parent_id):
    safe = str(name).replace("\\", "\\\\").replace("'", "\\'")
    q = (f"'{parent_id}' in parents and name = '{safe}' "
         "and mimeType = 'application/vnd.google-apps.folder' and trashed = false")
    hit = service.files().list(q=q, fields="files(id)").execute().get('files', [])
    if hit:
        return hit[0]['id']
    return service.files().create(
        body={'name': name, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [parent_id]},
        fields='id').execute().get('id')


def ensure_done_folder(student_name, creds):
    """生徒フォルダの下の「実施済答案」フォルダ（無ければ作成）の ID を返す。答案写真はすべてここへ保存する。"""
    student_folder, _ = ensure_drive_folder(student_name, creds)
    return _find_or_create_folder(build('drive', 'v3', credentials=creds), DONE_FOLDER, student_folder)


def _list_children(service, parent_id, extra_q=""):
    """フォルダ直下の子（ページングを辿って全件）を返す。"""
    out, token = [], None
    while True:
        res = service.files().list(
            q=f"'{parent_id}' in parents and trashed = false" + extra_q,
            fields="nextPageToken, files(id, name, mimeType)", pageSize=1000, pageToken=token).execute()
        out.extend(res.get('files', []))
        token = res.get('nextPageToken')
        if not token:
            return out


def plan_move_to_done(creds):
    """各生徒フォルダの直下に置かれた答案写真（画像ファイル）を洗い出す。
       戻り値: [{"student", "folder_id", "files": [{"id", "name"}]}]（対象がある生徒のみ）"""
    svc = build('drive', 'v3', credentials=creds)
    plan = []
    folders = _list_children(svc, PARENT_FOLDER_ID, " and mimeType = 'application/vnd.google-apps.folder'")
    for f in sorted(folders, key=lambda x: x.get('name', '')):
        files = [c for c in _list_children(svc, f['id'], " and mimeType contains 'image/'")]
        if files:
            plan.append({"student": f['name'], "folder_id": f['id'],
                         "files": [{"id": c['id'], "name": c['name']} for c in files]})
    return plan


def move_to_done(plan, creds, on_progress=None):
    """plan_move_to_done の結果に沿って、写真を各生徒の「実施済答案」へ移動する。戻り値: (移動件数, エラー一覧)"""
    svc = build('drive', 'v3', credentials=creds)
    total = sum(len(p["files"]) for p in plan)
    done, errors = 0, []
    for p in plan:
        dest = _find_or_create_folder(svc, DONE_FOLDER, p["folder_id"])
        for f in p["files"]:
            if on_progress:
                on_progress(done + len(errors), total, f"{p['student']} / {f['name']}")
            try:
                svc.files().update(fileId=f["id"], addParents=dest,
                                   removeParents=p["folder_id"], fields="id").execute()
                done += 1
            except Exception as e:
                errors.append(f"{p['student']} / {f['name']}: {e}")
    return done, errors


def upload_to_drive(filepath, filename, folder_id, creds):
    service = build('drive', 'v3', credentials=creds)
    media = MediaFileUpload(filepath, mimetype='image/jpeg', resumable=True)
    file = service.files().create(body={'name': filename, 'parents': [folder_id]}, media_body=media, fields='webViewLink').execute()
    return file.get('webViewLink')

def get_spreadsheet_data(creds):
    try:
        res = _sheets(creds).spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range='A:M').execute()
        rows = res.get('values', [])
        if not rows:
            return pd.DataFrame()
        # 見出しより長い行（M列追加前の古い見出しなど）でも落ちないように幅を揃える
        header = [str(h).strip() or f"列{i + 1}" for i, h in enumerate(rows[0])]
        width = max([len(header)] + [len(r) for r in rows[1:]])
        header += [f"列{i + 1}" for i in range(len(header), width)]
        data = [list(r) + [''] * (width - len(r)) for r in rows[1:]]
        return pd.DataFrame(data, columns=header)
    except Exception:
        return pd.DataFrame()

def process_master_file_from_path(filepath, client):
    ai_files = []
    if filepath.lower().endswith('.pdf'):
        doc = fitz.open(filepath)
        for i in range(len(doc)):
            page = doc.load_page(i); pix = page.get_pixmap(dpi=150)
            tmp = os.path.join(tempfile.gettempdir(), f"master_{uuid.uuid4().hex}.png")
            pix.save(tmp)
            try:
                af = client.files.upload(file=tmp)
                while af.state.name == 'PROCESSING': time.sleep(2); af = client.files.get(name=af.name)
                if af.state.name == 'ACTIVE': ai_files.append(af)
            finally:
                try: os.remove(tmp)
                except: pass
    else:
        af = client.files.upload(file=filepath)
        while af.state.name == 'PROCESSING': time.sleep(2); af = client.files.get(name=af.name)
        if af.state.name == 'ACTIVE': ai_files.append(af)
    return ai_files

def get_best_model(client):
    # 安定版を先頭に。プレビュー版は提供終了で消えることがあるため後ろに置く。
    preferred = ['gemini-2.5-flash', 'gemini-2.5-pro',
                 'gemini-2.5-flash-preview-05-20', 'gemini-2.5-pro-exp-03-25',
                 'gemini-1.5-pro-latest', 'gemini-1.5-pro']
    try:
        available = [m.name.replace('models/', '') for m in client.models.list()]
        for model in preferred:
            if model in available:
                return model
    except Exception:
        pass
    return 'gemini-2.5-flash'


# ==========================================
# 採点写真の読み取り（前面処理）→ 確認フォーム → 記録
# ==========================================
REVIEW_COLUMNS = ["ファイル", "テキスト名", "ページ", "章", "節", "問題番号", "小問数",
                  "節タイトル", "テストのタイトル"]
# 結果シートの見出し（A〜M。L列は空けてある）
RESULT_HEADER = ["日時", "生徒名", "科目", "テキスト名", "ページ", "章", "節", "問題番号",
                 "写真リンク", "総問題数", "節タイトル", "提出F", "テストのタイトル"]
SUBMIT_FLAG_OLD = "弱点補強テスト実施F"   # L列の旧見出し（見つけたら「提出F」に改名する）

_MARK_RULES = (
    "【採点記号の意味 ＝ 最重要ルール】\n"
    "・問題番号が赤い〇（丸・楕円）で囲まれている → その問題は【正解】。wrong に入れない。\n"
    "・問題番号のそばに赤い『レ点』『チェック(✓)』『斜線(／)』『×』のいずれかが付いている → 【間違い】。wrong に入れる。\n"
    "・□（チェックボックス）が黒く塗りつぶされている → 【間違い】。\n"
    "・重要：赤い〇（丸）は必ず【正解】です。丸を間違いと誤認しないこと。\n"
    "・〇でも×系でもなく、□も塗られていない無印は、正解として扱う。\n"
    "・計算の途中式や答えの数値（例: -8, 3/4）は問題番号ではない。\n"
    "・問題番号は「1」「(2)」「問3」「(ア)」「①」のような番号表記。"
    "手書きで番号（ア・イ・ウ や (1)・① 等）が振られている場合は、その手書き番号を優先して読み取る。\n"
)


def _photo_prompt(mode, text_name):
    """採点済み写真から間違いを読み取らせるプロンプト（元の問題PDFは使わない）。"""
    if mode == "confirm":
        return (
            "これは採点済みのテスト答案の写真です（理解度確認テスト・復習テストなど）。"
            "赤ペンの採点記号を1問ずつ判定してください。\n\n"
            + _MARK_RULES +
            "\n【答案の構成】用紙の一番上にテスト名、太字・色付きの見出しが『単元』、"
            "その下の 1. 2. 3. が『大問』、(1)(2)… が『小問』です。"
            "大問番号は単元ごとに 1 から振り直されることがあるので、"
            "必ず『どの単元の大問か』も答えてください。\n\n"
            "【出題元の情報も読み取る】このテストは市販テキスト・問題集から出題されています。"
            "用紙の見出し・欄外・大問のそばに、元になったテキスト名や『第2章』『2-1』"
            "『P.32〜』のような章・節の表記があれば、必ず読み取ってください。\n\n"
            "【出力形式】JSON配列のみ（説明文は不要）。大問ごとに1要素。\n"
            '[{"test_title":"第3回 理解度確認テスト","text":"新中学問題集 数学1年",'
            '"chapter":"第2章 文字と式","section":"第1節 文字を使った式","unit":"正負の数",'
            '"daimon":"1","total":4,"wrong":[{"number":"(1)"}]}]\n'
            "・test_title = 用紙上部のテスト名（例:「第3回 理解度確認テスト」「復習テスト②」）。読めなければ \"\"\n"
            "・text    = 出題元のテキスト・問題集の名前。読めなければ \"\"\n"
            "・chapter = 出題元の章（例:「第2章 文字と式」）。書かれていなければ \"\"\n"
            "・section = 出題元の節（例:「第1節 文字を使った式」「2-1」）。書かれていなければ \"\"\n"
            "・unit    = 大問の上にある単元見出し。無ければ \"\"\n"
            "・daimon  = 大問番号（半角数字。読めなければ \"\"）\n"
            "・total   = その大問に含まれる小問 (1)(2)(3)… の個数。数えられなければ 0\n"
            "・wrong   = 間違いだった小問の番号だけを並べる（1問も無ければ空配列）\n"
            "推測で埋めないこと。読めない項目は必ず空文字にする。"
        )
    return (
        "これは採点済みの答案（テキスト・問題集）の写真です。赤ペンの採点記号を1問ずつ判定してください。\n\n"
        + _MARK_RULES +
        "・写真内に印刷されている『ページ番号』を読み取り、page に半角数字で入れる。読めなければ \"\"。\n\n"
        "【出力形式】JSON配列のみ（説明文は不要）。\n"
        '[{"chapter":"' + str(text_name or "") + '","section":"項目名","page":"8","total":4,'
        '"wrong":[{"page":"8","number":"(1)"}]}]\n'
        "・chapter = \"" + str(text_name or "") + "\"（固定）\n"
        "・section = その問題群の項目名（読めなければ \"\"）\n"
        "・total   = その項目に含まれる問題の総数。数えられなければ 0\n"
        "・wrong   = 間違いだった問題だけを並べる（page と number／1問も無ければ空配列）"
    )


def _upload_photo_to_gemini(client, photo_path):
    ai_photo = client.files.upload(file=photo_path)
    while ai_photo.state.name == 'PROCESSING':
        time.sleep(1)
        ai_photo = client.files.get(name=ai_photo.name)
    return ai_photo


def analyze_photos_for_review(images, mode, text_name, master_index, api_key,
                              selected_master_path=None, test_title="", on_progress=None):
    """採点済み写真を1枚ずつ読み取り、確認フォーム用の行リストを返す。
       この段階では Drive にもスプレッドシートにも一切書き込まない。
       戻り値: (rows, errors, 使用モデル名)"""
    client = genai.Client(api_key=api_key)
    model = get_best_model(client)
    ai_master_files = (process_master_file_from_path(selected_master_path, client)
                       if (mode != "confirm" and selected_master_path) else [])
    rows, errors = [], []
    for i, (path, name) in enumerate(images):
        if on_progress:
            on_progress(i, len(images), name)
        # ファイル名にページ番号があれば、AIが読めなかったときの候補にする（例: 数学_p45.jpg）
        fallback_page = _page_from_filename(name) if mode != "confirm" else ""
        n_before = len(rows)
        try:
            ai_photo = _upload_photo_to_gemini(client, path)
            resp = client.models.generate_content(
                model=model, contents=ai_master_files + [ai_photo, _photo_prompt(mode, text_name)])
            result = _extract_json_array(resp.text)
        except Exception as e:
            errors.append(f"{name}: {e}")
            result = []
        for s in (result or []):
            if not isinstance(s, dict):
                continue
            total = _to_int(s.get("total")) or 0
            wrongs = [w for w in (s.get("wrong") or []) if isinstance(w, dict)]
            if mode == "confirm":
                # 出題元（テキスト名・章・節）を優先。読めなければ画面で入力した既定値／単元見出しで補う
                ttl = str(s.get("test_title", "") or "").strip() or test_title
                src = str(s.get("text", "") or "").strip() or text_name
                unit = str(s.get("unit", "") or "").strip()
                ch = str(s.get("chapter", "") or "").strip()
                se = str(s.get("section", "") or "").strip()
                if not se:
                    se = unit          # 節が印刷されていなければ単元見出しを節に
                elif not ch:
                    ch = unit          # 節はあるが章が無ければ単元見出しを章に
                daimon = _to_int(s.get("daimon"))
                dm = str(daimon) if daimon else ""
                for w in wrongs:
                    rows.append({"ファイル": name, "テキスト名": src,
                                 "ページ": dm, "章": ch, "節": se,
                                 "問題番号": str(w.get("number", "") or "").strip(),
                                 "小問数": total, "節タイトル": "", "テストのタイトル": ttl})
            else:
                ai_ch = str(s.get("chapter", "") or "").strip()
                ai_se = str(s.get("section", "") or "").strip()
                sec_page = _to_int(s.get("page"))
                for w in wrongs:
                    pg = _to_int(w.get("page")) or sec_page
                    page = str(pg) if pg else fallback_page
                    m_ch, m_se, m_ti = lookup_section(master_index, text_name, page)
                    rows.append({"ファイル": name, "テキスト名": text_name,
                                 "ページ": page, "章": m_ch or ai_ch, "節": m_se or ai_se,
                                 "問題番号": str(w.get("number", "") or "").strip(),
                                 "小問数": total, "節タイトル": m_ti, "テストのタイトル": ""})
        if len(rows) == n_before:
            # 読み取れなかった（または間違いが1問も無かった）→ 手入力用の空行を1行だけ置く
            rows.append({"ファイル": name, "テキスト名": text_name, "ページ": fallback_page,
                         "章": "", "節": "", "問題番号": "", "小問数": 0, "節タイトル": "",
                         "テストのタイトル": (test_title if mode == "confirm" else "")})
    return rows, errors, model


def relookup_rows(rows, master_index):
    """フォームで直したページ番号から、章・節・節タイトルを目次マスタで引き直す。"""
    out = []
    for r in rows:
        r = dict(r)
        m_ch, m_se, m_ti = lookup_section(master_index, str(r.get("テキスト名", "") or ""),
                                          r.get("ページ", ""))
        if m_ch or m_se or m_ti:
            r["章"], r["節"], r["節タイトル"] = m_ch, m_se, m_ti
        out.append(r)
    return out


def ensure_result_header(creds):
    """結果シート1行目の見出しを確認する。L列の旧見出し「弱点補強テスト実施F」は「提出F」に改名し、
       M列（テストのタイトル）が無ければ書き足す。"""
    try:
        svc = _sheets(creds)
        row = (svc.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID, range="A1:M1").execute().get("values") or [[]])[0]
        if len(row) >= 12 and str(row[11]).strip() == SUBMIT_FLAG_OLD:
            svc.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID, range="L1",
                valueInputOption="RAW", body={"values": [["提出F"]]}).execute()
        if len(row) >= 13 and str(row[12]).strip():
            return
        if not [c for c in row if str(c).strip()]:      # 見出しが空 → まとめて作る
            svc.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID, range="A1:M1",
                valueInputOption="RAW", body={"values": [RESULT_HEADER]}).execute()
        else:                                          # 既存の見出しは触らず M1 だけ足す
            svc.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID, range="M1",
                valueInputOption="RAW", body={"values": [["テストのタイトル"]]}).execute()
    except Exception as e:
        print(f"ensure_result_header: {e}")


def record_reviewed_rows(rows, student_name, subject_name, images, creds, on_progress=None,
                         submit_flag=False):
    """確認フォームで確定した行をスプレッドシートへ記録し、写真は正解・不正解にかかわらず
       すべて生徒の「実施済答案」フォルダへ保存する。問題番号が空の行はシートには書かない。
       submit_flag=True なら L列（提出F）に 1 を立てる（送付した理解度確認テストの答案）。
       戻り値: (記録件数, {ファイル名: 写真リンク})"""
    def _s(r, k):
        v = r.get(k, "")
        if v is None:
            return ""
        v = str(v).strip()
        return "" if v.lower() in ("nan", "none") else v

    valid = [r for r in rows if _s(r, "問題番号") not in ("", "-")]
    folder_id = ensure_done_folder(student_name, creds)
    first_row = {}                     # 写真ごとの代表行（ファイル名の見出しに使う）
    for r in valid:
        first_row.setdefault(_s(r, "ファイル"), r)
    links = {}
    for i, (path, fn) in enumerate(images):   # 間違いが無い写真も含めて全部保存する
        if on_progress:
            on_progress(i, len(images), fn)
        r = first_row.get(fn, {})
        head = "_".join([x for x in (_s(r, "章"), _s(r, "節"), _s(r, "節タイトル")) if x]) if r else ""
        pg = _s(r, "ページ") if r else ""
        prefix = ""
        if head:
            prefix = ("[" + head + "]").replace("/", "／") + (f"_p{pg}" if pg else "") + "_"
        links[fn] = upload_to_drive(path, prefix + fn, folder_id, creds)
    if not valid:
        return 0, links
    now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    values = []
    for r in valid:
        values.append([
            now, student_name, subject_name, _s(r, "テキスト名"),
            _s(r, "ページ"), _s(r, "章"), _s(r, "節"), _s(r, "問題番号"),
            links.get(_s(r, "ファイル"), ""), (_to_int(r.get("小問数")) or 0), _s(r, "節タイトル"),
            "1" if submit_flag else "",      # L列: 提出F
            _s(r, "テストのタイトル"),         # M列
        ])
    ensure_result_header(creds)
    _sheets(creds).spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID, range='A1',
        valueInputOption='USER_ENTERED', body={'values': values}
    ).execute()
    return len(values), links



# ==========================================
# 🎓 過去問の取込（テキスト・確認テストとは別の取込）
# ==========================================
KAKOMON_TAB = "過去問"
KAKOMON_HEADER = ["登録日時", "生徒名", "実施日", "学校名", "学部", "科目", "方式", "年度",
                  "得点", "満点", "写真リンク", "記録ID"]
KAKOMON_COLUMNS = ["写真", "実施日", "学校名", "学部", "科目", "方式", "年度", "得点", "満点"]


def ensure_kakomon_tab(creds):
    """「過去問」タブが無ければ作成し、見出し行を入れる。"""
    svc = _sheets(creds)
    meta = svc.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    titles = [s['properties']['title'] for s in meta.get('sheets', [])]
    if KAKOMON_TAB not in titles:
        svc.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"requests": [{"addSheet": {"properties": {"title": KAKOMON_TAB}}}]}).execute()
        svc.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID, range=f"{KAKOMON_TAB}!A1",
            valueInputOption="RAW", body={"values": [KAKOMON_HEADER]}).execute()


def _kakomon_prompt(n_photos):
    return (
        f"これは生徒が解いた入試の過去問（採点済みの答案・問題冊子）の写真です（全{n_photos}枚。"
        "送った順に 1, 2, 3 … と番号を付けます）。\n"
        "写真から次の情報を読み取り、JSON配列で返してください。同じ試験の複数ページは1件にまとめ、"
        "学校・学部・科目・年度のどれかが違えば別の要素にしてください。\n\n"
        "【出力形式】JSON配列のみ（説明文は不要）。\n"
        '[{"photos":[1,2],"date":"2026-09-15","school":"早稲田大学","faculty":"商学部",'
        '"subject":"英語","method":"一般選抜","year":"2024","score":72,"max_score":100}]\n'
        "・photos    = その試験が写っている写真の番号\n"
        "・date      = 実施日（答案に書かれた日付。YYYY-MM-DD。書かれていなければ \"\"）\n"
        "・school    = 学校名（大学名・高校名）\n"
        "・faculty   = 学部・学科（無ければ \"\"）\n"
        "・subject   = 科目（例: 英語、数学IA、日本史）\n"
        "・method    = 入試方式（例: 一般選抜、共通テスト利用、学校推薦型、前期日程、A方式。無ければ \"\"）\n"
        "・year      = 過去問の年度（例: 2024。表紙や欄外の『2024年度』『令和6年度』を読む）\n"
        "・score     = 採点後の合計得点（赤字で書かれた合計点。無ければ null）\n"
        "・max_score = 満点（配点の合計。無ければ null）\n"
        '・wrong     = 間違えた問題（赤の×・レ点・斜線などが付いた問題）の大問と問題番号。'
        '例 [{"daimon":"Ⅰ","number":"問3"}]。無ければ []\n'
        "推測で埋めないこと。読めない項目は \"\" または null にする。"
    )


def _parse_exam_date(v, today=None):
    """『2026-09-15』『2026/9/15』『2026年9月15日』『9/15』→ datetime.date。読めなければ None。"""
    today = today or datetime.date.today()
    s = _half(str(v or "")).strip()
    m = re.search(r"(20\d{2})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})", s)
    if m:
        y, mo, d = (int(x) for x in m.groups())
    else:
        m = re.search(r"(\d{1,2})\s*[/月]\s*(\d{1,2})", s)
        if not m:
            return None
        y, mo, d = today.year, int(m.group(1)), int(m.group(2))
    try:
        return datetime.date(y, mo, d)
    except ValueError:
        return None


def _norm_year(v):
    """年度を西暦4桁に（『2024年度』→2024、『令和6年度』『R6』→2024、『平成30』→2018）。"""
    s = _half(str(v or "")).strip()
    if s.lower() in ("nan", "none"):
        return ""
    m = re.search(r"(?:19|20)\d{2}", s)
    if m:
        return m.group(0)
    m = re.search(r"(?:令和|R)\s*(\d{1,2}|元)", s, re.I)
    if m:
        return str(2018 + (1 if m.group(1) == "元" else int(m.group(1))))
    m = re.search(r"(?:平成|H)\s*(\d{1,2}|元)", s, re.I)
    if m:
        return str(1988 + (1 if m.group(1) == "元" else int(m.group(1))))
    return s


def _num(v):
    """『72』『72点』『７２』→ 72.0。読めなければ None。"""
    if v is None:
        return None
    s = _half(str(v)).replace(",", "").strip()
    m = re.search(r"-?\d+(?:\.\d+)?", s)
    return float(m.group(0)) if m else None


def _fmt_num(v):
    n = _num(v)
    if n is None:
        return ""
    return str(int(n)) if float(n).is_integer() else str(n)


def _as_date_str(v):
    """表（data_editor）から戻った実施日を 'YYYY-MM-DD' に。空なら ""。"""
    if v is None or str(v).strip().lower() in ("", "nan", "nat", "none"):
        return ""
    if isinstance(v, datetime.datetime):
        return v.date().isoformat()
    if isinstance(v, datetime.date):
        return v.isoformat()
    d = _parse_exam_date(v)
    return d.isoformat() if d else str(v).strip()


def analyze_kakomon(images, api_key, default_subject="", on_progress=None, wrongs_out=None):
    """過去問の写真をまとめて読み取り、確認フォーム用の行（1行＝1回分の過去問）を返す。
       wrongs_out にリストを渡すと、間違えた問題（KK_WRONG_COLUMNS の形）をそこへ追加する。
       この段階では Drive にもスプレッドシートにも書き込まない。戻り値: (rows, errors, 使用モデル名)"""
    client = genai.Client(api_key=api_key)
    model = get_best_model(client)
    uploaded, idx_map, errors = [], [], []
    for i, (path, name) in enumerate(images):
        if on_progress:
            on_progress(i, len(images), name)
        try:
            uploaded.append(_upload_photo_to_gemini(client, path))
            idx_map.append(i)
        except Exception as e:
            errors.append(f"{name}: {e}")
    result = []
    if uploaded:
        try:
            resp = client.models.generate_content(
                model=model, contents=uploaded + [_kakomon_prompt(len(uploaded))])
            result = _extract_json_array(resp.text)
        except Exception as e:
            errors.append(f"読み取り: {e}")
    today = datetime.date.today()
    all_nums = ",".join(str(i + 1) for i in range(len(images)))
    rows = []
    for o in (result or []):
        if not isinstance(o, dict):
            continue
        nums = []
        for k in (o.get("photos") or []):
            k = _to_int(k)
            if k and 1 <= k <= len(idx_map) and (idx_map[k - 1] + 1) not in nums:
                nums.append(idx_map[k - 1] + 1)
        rows.append({
            "写真": ",".join(str(n) for n in sorted(nums)) or all_nums,
            "実施日": _parse_exam_date(o.get("date"), today) or today,
            "学校名": str(o.get("school", "") or "").strip(),
            "学部": str(o.get("faculty", "") or "").strip(),
            "科目": str(o.get("subject", "") or "").strip() or default_subject,
            "方式": str(o.get("method", "") or "").strip(),
            "年度": _norm_year(o.get("year")),
            "得点": _num(o.get("score")),
            "満点": _num(o.get("max_score")),
        })
        if wrongs_out is not None:
            for w in (o.get("wrong") or []):
                if isinstance(w, dict) and (str(w.get("daimon", "") or "").strip() or str(w.get("number", "") or "").strip()):
                    wrongs_out.append({"過去問": str(len(rows)), "大問": str(w.get("daimon", "") or "").strip(),
                                       "問題番号": str(w.get("number", "") or "").strip(),
                                       "分野（見出し）": "", "タイトル": ""})
    if not rows:   # 読み取れなかった → 手入力用の行を1行だけ置く
        rows.append({"写真": all_nums, "実施日": today, "学校名": "", "学部": "",
                     "科目": default_subject, "方式": "", "年度": "", "得点": None, "満点": None})
    return rows, errors, model


def record_kakomon(rows, student_name, images, creds, on_progress=None, wrongs=None):
    """確認フォームで確定した過去問を、写真は Drive、記録は「過去問」タブへ書き込む。
       wrongs（間違えた問題の表）があれば、分野ごとに結果シート（1枚目）にも記録する
       （章＝問題の先頭の見出し、節＝問題番号のタイトル、M列＝どの過去問か）。
       学校名も科目も空の行は記録しない。戻り値: (記録件数, メール用の要約行)"""
    def _s(v):
        if v is None:
            return ""
        v = str(v).strip()
        return "" if v.lower() in ("nan", "none", "nat") else v

    def _photo_nums(r):
        out = []
        for x in re.split(r"[,\s、，・]+", _half(_s(r.get("写真")))):
            k = _to_int(x)
            if k and 1 <= k <= len(images) and k not in out:
                out.append(k)
        return out

    valid = [r for r in rows if _s(r.get("学校名")) or _s(r.get("科目"))]
    if not valid:
        return 0, []
    folder_id = ensure_done_folder(student_name, creds)
    todo, seen = [], set()
    for r in valid:
        for k in _photo_nums(r):
            if k not in seen:
                seen.add(k)
                todo.append((k, r))
    for k in range(1, len(images) + 1):       # どの行にも割り当てられなかった写真も保存する
        if k not in seen:
            seen.add(k)
            todo.append((k, {}))
    link_of = {}
    for i, (k, r) in enumerate(todo):
        path, name = images[k - 1]
        if on_progress:
            on_progress(i, len(todo), name)
        yr = _norm_year(r.get("年度")) if r else ""
        head = "_".join(x for x in ("過去問", _s(r.get("学校名")), _s(r.get("学部")), _s(r.get("科目")),
                                    (yr + "年度") if yr else "") if x)
        link_of[k] = upload_to_drive(path, ("[" + head + "]").replace("/", "／") + "_" + name,
                                     folder_id, creds)
    now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    today = datetime.date.today().isoformat()
    values, summary = [], []
    for r in valid:
        links = [link_of[k] for k in _photo_nums(r) if k in link_of]
        row = [now, student_name, _as_date_str(r.get("実施日")) or today,
               _s(r.get("学校名")), _s(r.get("学部")), _s(r.get("科目")), _s(r.get("方式")),
               _norm_year(r.get("年度")), _fmt_num(r.get("得点")), _fmt_num(r.get("満点")),
               "\n".join(links), uuid.uuid4().hex[:12]]
        values.append(row)
        score = (row[8] + ("/" + row[9] if row[9] else "") + "点") if row[8] else "得点なし"
        summary.append(f"{row[2]} {row[3]} {row[4]} {row[5]} {row[6]} "
                       f"{(row[7] + '年度') if row[7] else ''} … {score}\n  " + " ".join(links))
    ensure_kakomon_tab(creds)
    # 文字列のまま保存（RAW）し、列ごとの型を揃える（スケジュール管理側の CSV 取込で値が欠けないように）
    _sheets(creds).spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID, range=f"{KAKOMON_TAB}!A1",
        valueInputOption='RAW', body={'values': values}).execute()

    # ---- 間違えた問題 → 結果シート（分野ごとの弱点としてスケジュール管理・復習に使える）
    wrong_values = []
    for w in (wrongs or []):
        idx = _to_int(w.get("過去問"))
        exam = rows[idx - 1] if idx and 1 <= idx <= len(rows) else (valid[0] if len(valid) == 1 else None)
        num = " ".join(x for x in (_s(w.get("大問")), _s(w.get("問題番号"))) if x)
        if exam is None or exam not in valid or not num:
            continue
        yr = _norm_year(exam.get("年度"))
        label = " ".join(x for x in (_s(exam.get("学校名")), _s(exam.get("学部")), (yr + "年度") if yr else "",
                                     _s(exam.get("方式")), _s(exam.get("科目"))) if x)
        links = [link_of[k] for k in _photo_nums(exam) if k in link_of]
        wrong_values.append([now, student_name, _s(exam.get("科目")),
                             " ".join(x for x in ("過去問", _s(exam.get("学校名")), _s(exam.get("学部"))) if x),
                             "", _s(w.get("分野（見出し）")), _s(w.get("タイトル")), num,
                             "\n".join(links), "", "", "", "過去問 " + label])
    if wrong_values:
        ensure_result_header(creds)
        _sheets(creds).spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_ID, range='A1',
            valueInputOption='USER_ENTERED', body={'values': wrong_values}).execute()
        summary.append(f"間違えた問題 {len(wrong_values)} 件を分野ごとに結果シートへ記録しました")
    return len(values), summary


# ==========================================
# 問題（問題用紙・問題のページ）から、間違えた問題の分野を読み取る
#   章 ＝ その問題が含まれるまとまりの先頭の見出し、節 ＝ 問題番号のところのタイトル
# ==========================================
KK_WRONG_COLUMNS = ["過去問", "大問", "問題番号", "分野（見出し）", "タイトル"]


def _classify_prompt(n_pages, targets):
    lines = "\n".join(
        f"{t['key']}: " + " ".join(x for x in (
            f"大問{t['daimon']}" if t.get("daimon") else "", str(t.get("number") or ""),
            f"（p.{t['page']}）" if t.get("page") else "",
            f"［手がかり: {t['hint']}］" if t.get("hint") else "") if x)
        for t in targets)
    return (
        f"これは問題用紙（問題のページ）の画像です（全{n_pages}枚）。\n"
        "下の一覧の各問題が問題用紙のどこにあるかを探し、その問題の分野を次の2つで答えてください。\n"
        "・heading = その問題が含まれるまとまりの先頭にある見出し（例:「第2章 二次関数」「Ⅰ 長文読解」「文法・語法」）\n"
        "・title   = 問題番号のところに書かれたタイトル（例:「最大・最小」「仮定法過去」）。番号だけでタイトルが無ければ \"\"\n"
        "見出し・タイトルは用紙に書かれている文字をそのまま使い、推測で作らないこと。見つからない問題は両方 \"\" にする。\n\n"
        f"【問題の一覧】\n{lines}\n\n"
        "【出力形式】JSON配列のみ（説明文は不要）。\n"
        '[{"key":"1","heading":"第2章 二次関数","title":"最大・最小"}]'
    )


def classify_by_questions(targets, qimages, api_key, on_progress=None):
    """間違えた問題 targets（key・大問・問題番号・ページ・手がかり）の分野を、問題 qimages の
       見出しとタイトルから読み取る。戻り値: ({key: (見出し, タイトル)}, errors)"""
    if not targets or not qimages:
        return {}, []
    client = genai.Client(api_key=api_key)
    model = get_best_model(client)
    uploaded, errors = [], []
    for i, (path, name) in enumerate(qimages):
        if on_progress:
            on_progress(i, len(qimages), name)
        try:
            uploaded.append(_upload_photo_to_gemini(client, path))
        except Exception as e:
            errors.append(f"{name}: {e}")
    if not uploaded:
        return {}, errors
    try:
        resp = client.models.generate_content(
            model=model, contents=uploaded + [_classify_prompt(len(uploaded), targets)])
        arr = _extract_json_array(resp.text)
    except Exception as e:
        return {}, errors + [f"分野の読み取り: {e}"]
    out = {}
    for o in arr:
        if isinstance(o, dict) and str(o.get("key", "")).strip():
            out[str(o["key"]).strip()] = (str(o.get("heading", "") or "").strip(),
                                          str(o.get("title", "") or "").strip())
    return out, errors


def apply_fields_to_review_rows(rows, qimages, api_key, mode, on_progress=None):
    """テキスト・確認テスト：間違えた問題の行に、問題から読んだ分野を入れる（章＝見出し、節＝タイトル）。
       テキストは目次で章・節が引けた行（節タイトルあり）をそのままにし、引けなかった行だけ補う。
       戻り値: (分野を入れた行数, errors)"""
    targets = []
    for i, r in enumerate(rows):
        if not str(r.get("問題番号", "") or "").strip():
            continue
        if mode == "text" and str(r.get("節タイトル", "") or "").strip():
            continue                                    # 目次で引けている
        targets.append({"key": str(i + 1), "number": str(r.get("問題番号")),
                        "daimon": str(r.get("ページ") or "") if mode == "confirm" else "",
                        "page": str(r.get("ページ") or "") if mode == "text" else "",
                        "hint": str(r.get("節") or r.get("章") or "")})
    got, errors = classify_by_questions(targets, qimages, api_key, on_progress)
    n = 0
    for t in targets:
        h, ti = got.get(t["key"], ("", ""))
        if not (h or ti):
            continue
        r = rows[int(t["key"]) - 1]
        if h:
            r["章"] = h
        if ti:
            r["節"] = ti
        n += 1
    return n, errors


def apply_fields_to_kakomon_wrongs(wrongs, exams, qimages, api_key, on_progress=None):
    """過去問：間違えた問題の表に、問題冊子から読んだ分野を入れる。戻り値: (分野を入れた行数, errors)"""
    targets = []
    for i, w in enumerate(wrongs):
        idx = _to_int(w.get("過去問"))
        exam = exams[idx - 1] if idx and 1 <= idx <= len(exams) else {}
        targets.append({"key": str(i + 1), "daimon": str(w.get("大問") or ""), "number": str(w.get("問題番号") or ""),
                        "hint": " ".join(x for x in (str(exam.get("学校名") or ""), str(exam.get("科目") or "")) if x)})
    got, errors = classify_by_questions(targets, qimages, api_key, on_progress)
    n = 0
    for t in targets:
        h, ti = got.get(t["key"], ("", ""))
        if h or ti:
            wrongs[int(t["key"]) - 1]["分野（見出し）"] = h
            wrongs[int(t["key"]) - 1]["タイトル"] = ti
            n += 1
    return n, errors


LEDGER_TAB = "送付テスト"   # 宿題自動送信（LINE）で送ったテストの台帳。宿題自動送信側が書き込む


def get_ledger_data(creds):
    """「送付テスト」タブ（送付日時・生徒名・テスト名・テストID・処理F・提出F・正解率・状態など）を返す。"""
    try:
        rows = _sheets(creds).spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID, range=f"{LEDGER_TAB}!A:R").execute().get('values', [])
        if len(rows) < 2:
            return pd.DataFrame()
        width = len(rows[0])
        return pd.DataFrame([(list(r) + [''] * width)[:width] for r in rows[1:]], columns=rows[0])
    except Exception:
        return pd.DataFrame()


def get_kakomon_data(creds):
    """「過去問」タブを DataFrame で返す（無ければ空）。"""
    try:
        rows = _sheets(creds).spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID, range=f"{KAKOMON_TAB}!A:L").execute().get('values', [])
        if len(rows) < 2:
            return pd.DataFrame()
        width = len(rows[0])
        return pd.DataFrame([(list(r) + [''] * width)[:width] for r in rows[1:]], columns=rows[0])
    except Exception:
        return pd.DataFrame()


# ==========================================
# 更新履歴・バージョン（CHANGELOG.md をサーバー上の正とする）
# ==========================================
def load_changelog():
    """CHANGELOG.md を読み、(全文, 最新バージョン, 更新日) を返す。無ければ空文字。"""
    try:
        text = io.open(CHANGELOG_PATH, encoding="utf-8").read()
    except Exception:
        return "", "", ""
    m = re.search(r"^##\s*v?(\d+(?:\.\d+)*)\s*[—–-]+\s*(\d{4}-\d{2}-\d{2})", text, re.M)
    return text, (m.group(1) if m else ""), (m.group(2) if m else "")


# ==========================================
# [統合] アップロード(画像/PDF/Word) → 画像群に展開
# ==========================================
def expand_uploaded_to_images(uploaded_file):
    """画像/PDF/Word を Gemini に渡す画像の一時ファイル群に展開して [(path, name), ...] を返す。"""
    name = uploaded_file.name
    ext = name.rsplit('.', 1)[-1].lower() if '.' in name else ''
    data = uploaded_file.getvalue()
    out = []
    try:
        if ext in ('heic', 'heif'):
            # iPhoneのHEIC/HEIF → JPEGへ変換（Geminiが読める形式にする）
            img = Image.open(io.BytesIO(data)).convert('RGB')
            tmp = os.path.join(tempfile.gettempdir(), f"photo_{uuid.uuid4().hex}.jpg")
            img.save(tmp, 'JPEG', quality=95)
            out.append((tmp, re.sub(r'\.(heic|heif)$', '.jpg', name, flags=re.IGNORECASE)))
        elif ext in ('jpg', 'jpeg', 'png', 'bmp', 'webp', 'gif'):
            tmp = os.path.join(tempfile.gettempdir(), f"photo_{uuid.uuid4().hex}.{ext if ext!='jpeg' else 'jpg'}")
            with open(tmp, 'wb') as f: f.write(data)
            out.append((tmp, name))
        elif ext == 'pdf':
            doc = fitz.open(stream=data, filetype="pdf")
            for i in range(len(doc)):
                pix = doc.load_page(i).get_pixmap(dpi=150)
                tmp = os.path.join(tempfile.gettempdir(), f"photo_{uuid.uuid4().hex}.png")
                pix.save(tmp)
                out.append((tmp, f"{name}_p{i+1}"))
        elif ext in ('docx', 'doc'):
            z = zipfile.ZipFile(io.BytesIO(data))
            media = [n for n in z.namelist()
                     if n.startswith('word/media/') and n.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif'))]
            for k, n in enumerate(media):
                e = n.rsplit('.', 1)[-1].lower()
                tmp = os.path.join(tempfile.gettempdir(), f"photo_{uuid.uuid4().hex}.{e}")
                with open(tmp, 'wb') as f: f.write(z.read(n))
                out.append((tmp, f"{name}_img{k+1}"))
    except Exception as e:
        print(f"展開エラー({name}): {e}")
    return out


# ==========================================
# Streamlit Web UI
# ==========================================
creds_ui = Credentials.from_authorized_user_info(GOOGLE_TOKEN_DICT)
master_index = load_master_index(creds_ui)  # [統合] Sheetから常時参照
students = load_list(creds_ui, STUDENT_TAB, "生徒名", STUDENT_LIST)   # [統合] 生徒名簿
subjects = load_list(creds_ui, SUBJECT_TAB, "科目", DEFAULT_SUBJECTS)  # [統合] 科目マスタ
APP_CHANGELOG, APP_VERSION, APP_UPDATED = load_changelog()

KIND_LABELS = {"text": "📄 テキスト", "confirm": "📝 確認テスト", "kakomon": "🎓 過去問"}
KIND_HELP = {
    "text": "テキスト・問題集の答案。ページ番号から目次マスタを逆引きして、章・節を記録します。",
    "confirm": "理解度確認テスト・復習テストの答案。テストのタイトルと出題元のテキスト名・章・節を写真から読み取ります。",
    "kakomon": "入試過去問の答案。実施日・学校名・学部・科目・方式・年度・得点を読み取り、"
               "スケジュール管理（schedule.compassonline.site）のテスト結果に表示します。",
}
PHOTO_TYPES = ['jpg', 'jpeg', 'png', 'heic', 'heif', 'pdf', 'docx', 'doc']

# ---- ヘッダー：ロゴ・システム名・バージョン ----
_hl, _hr = st.columns([1, 4])
with _hl:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=170)
with _hr:
    st.markdown(
        "<div style='padding-top:14px'>"
        "<span style='font-size:1.8rem;font-weight:800'>答案自動採点システム</span>"
        "<span style='margin-left:12px;padding:3px 12px;border-radius:999px;background:#0A1628;"
        f"color:#C9A24A;font-size:.85rem;font-weight:700;vertical-align:middle'>Ver. {APP_VERSION or '―'}</span>"
        + (f"<div style='opacity:.65;font-size:.8rem;margin-top:4px'>最終更新 {APP_UPDATED}"
           "　／　更新内容は「📜 更新履歴」タブ</div>" if APP_UPDATED else "")
        + "</div>", unsafe_allow_html=True)


@st.cache_data(ttl=60, show_spinner=False)
def _result_table(_creds, which):
    """集計結果の表（1分キャッシュ。記録・更新ボタンで破棄）。"""
    if which == "ledger":
        return get_ledger_data(_creds)
    return get_kakomon_data(_creds) if which == "kakomon" else get_spreadsheet_data(_creds)


def _flash_show():
    f = st.session_state.pop("flash", None)
    if f:
        getattr(st, f[0])(f[1])
        if f[0] == "success":
            st.balloons()


def _clear_review(mode):
    if mode == "kakomon":
        st.session_state.pop("kk_rows", None)
        st.session_state.pop("kk_wrongs", None)
    elif st.session_state.get("review_mode") == mode:
        st.session_state.pop("review_rows", None)


def _photo_uploader(mode):
    """種類ごとに独立したアップロード欄。内容が変わった時だけ画像へ展開（PDF=各ページ、Word=埋め込み画像）。"""
    nonce = st.session_state.get(f"up_nonce_{mode}", 0)
    files = st.file_uploader("採点済みの写真／PDF／Word（複数可・iPhoneのHEICも可）",
                             type=PHOTO_TYPES, accept_multiple_files=True, key=f"up_{mode}_{nonce}")
    k_imgs, k_sig = f"imgs_{mode}", f"sig_{mode}"
    if files:
        sig = tuple((f.name, f.size) for f in files)
        if st.session_state.get(k_sig) != sig:
            imgs = []
            for f in files:
                imgs.extend(expand_uploaded_to_images(f))
            st.session_state[k_imgs] = imgs          # [(path, name), ...]
            st.session_state[k_sig] = sig
            _clear_review(mode)                     # 写真が変わったら読み取り結果は破棄
    else:
        st.session_state.pop(k_imgs, None)
        st.session_state.pop(k_sig, None)
        _clear_review(mode)
    return st.session_state.get(k_imgs, [])


def _question_uploader(mode):
    """答案に対応する問題（問題用紙・問題のページ）の取込欄（任意）。分野の読み取りに使う。"""
    nonce = st.session_state.get(f"up_nonce_{mode}", 0)
    files = st.file_uploader("答案に対応する問題（問題用紙・問題のページ）※任意",
                             type=PHOTO_TYPES, accept_multiple_files=True, key=f"qup_{mode}_{nonce}",
                             help="取り込むと、間違えた問題ごとに、問題の先頭の見出し（→章）と"
                                  "問題番号のタイトル（→節）から分野を読み取ります。")
    k_imgs, k_sig = f"qimgs_{mode}", f"qsig_{mode}"
    if files:
        sig = tuple((f.name, f.size) for f in files)
        if st.session_state.get(k_sig) != sig:
            q = []
            for f in files:
                q.extend(expand_uploaded_to_images(f))
            st.session_state[k_imgs] = q
            st.session_state[k_sig] = sig
    else:
        st.session_state.pop(k_imgs, None)
        st.session_state.pop(k_sig, None)
    return st.session_state.get(k_imgs, [])


def _finish(mode, imgs, msg):
    """記録後の後片付け：一時ファイル削除・読み取り結果の破棄・アップロード欄を空にする。"""
    for path, _nm in list(imgs) + list(st.session_state.get(f"qimgs_{mode}", [])):
        try:
            os.remove(path)
        except Exception:
            pass
    for k in (f"imgs_{mode}", f"sig_{mode}", f"qimgs_{mode}", f"qsig_{mode}"):
        st.session_state.pop(k, None)
    _clear_review(mode)
    st.session_state[f"up_nonce_{mode}"] = st.session_state.get(f"up_nonce_{mode}", 0) + 1
    _result_table.clear()
    st.session_state["flash"] = ("success", msg)
    st.rerun()


def render_review_form(mode, imgs, student_name, subject_name, text_name, test_title):
    """⑥ テキスト・確認テストの読み取り結果の確認・修正 → ⑦ 記録。"""
    st.markdown("#### ⑥ 読み取り結果の確認・修正")
    st.caption(f"日時: 記録時に自動で入ります　／　生徒名: **{student_name or '未選択'}**　／　"
               f"科目: **{subject_name or '未選択'}**")
    st.caption("AIが読めなかった項目は空欄です。ここで入力・修正してから記録してください。行の追加・削除もできます。"
               + ("確認テストの『テキスト名・章・節』には、写真から読み取った**出題元**が入ります。"
                  if mode == "confirm" else "")
               + "**問題番号が空の行はシートに記録されません**（全問正解の写真は空行のままで構いません）。"
               "写真は正解・不正解にかかわらず、すべて生徒の「実施済答案」フォルダに保存します。")
    names = [nm for _, nm in imgs]
    df = pd.DataFrame(st.session_state["review_rows"], columns=REVIEW_COLUMNS)
    edited = st.data_editor(
        df, num_rows="dynamic", width="stretch",
        key=f"review_editor_{st.session_state.get('review_ver', 0)}",
        column_config={
            "ファイル": st.column_config.SelectboxColumn("ファイル（写真）", options=names, width="medium"),
            "テキスト名": st.column_config.TextColumn("テキスト名"),
            "ページ": st.column_config.TextColumn("ページ", help="確認テストでは大問番号が入ります"),
            "章": st.column_config.TextColumn("章"),
            "節": st.column_config.TextColumn("節"),
            "問題番号": st.column_config.TextColumn("問題番号", help="間違えた問題の番号。空欄の行は記録されません"),
            "小問数": st.column_config.NumberColumn("小問数", min_value=0, step=1,
                                                 help="その大問（項目）に含まれる問題数。総問題数の列に入ります"),
            "節タイトル": st.column_config.TextColumn("節タイトル"),
            "テストのタイトル": st.column_config.TextColumn(
                "テストのタイトル", help="M列に記録されます（理解度確認テスト・復習テストなどの名称）"),
        })
    rows_now = edited.to_dict("records")

    c1, c2, c3 = st.columns([1.4, 1, 1])
    with c1:
        do_record = st.button("✅ ⑦ この内容で記録", type="primary", key=f"record_{mode}")
    with c2:
        if mode == "text" and st.button("📖 ページから章・節を再取得", key="relookup"):
            st.session_state["review_rows"] = relookup_rows(rows_now, master_index)
            st.session_state["review_ver"] = st.session_state.get("review_ver", 0) + 1
            st.rerun()
    with c3:
        if st.button("✖ 破棄", key=f"discard_{mode}"):
            _clear_review(mode)
            st.rerun()
    if not do_record:
        return
    if not student_name:
        st.error("生徒名を選んでください")
        return
    bar = st.progress(0.0, text="記録を開始します…")

    def _prog(i, n, nm):
        bar.progress(i / max(1, n), text=f"写真をDriveへ保存中 {i + 1}/{n}： {nm}")

    try:
        n, links = record_reviewed_rows(rows_now, student_name, subject_name, imgs, creds_ui,
                                        on_progress=_prog,
                                        submit_flag=(mode == "confirm" and bool(st.session_state.get("submit_flag"))))
    except Exception as e:
        bar.empty()
        st.error(f"記録エラー: {e}")
        return
    send_notification_email_plan_b(
        f"【完了】{student_name} さんの記録（{test_title or text_name}）",
        f"種類: {KIND_LABELS[mode]}\n科目: {subject_name}\nテスト: {test_title}\nテキスト名: {text_name}\n"
        f"記録件数: {n}件\n\n" + "\n".join(f"{k}: {v}" for k, v in links.items()))
    _finish(mode, imgs, (f"✅ {n} 件を記録しました" if n else "✅ 間違いの記録はありません（全問正解）")
            + f"（写真 {len(links)} 枚を「{DONE_FOLDER}」に保存）。")


def render_kakomon_form(imgs, student_name):
    """⑥ 過去問の読み取り結果の確認・修正 → ⑦ 記録。"""
    st.markdown("#### ⑥ 読み取り結果の確認・修正（過去問）")
    st.caption(f"生徒名: **{student_name or '未選択'}**　／　1行＝1回分の過去問です（学校・学部・科目・年度ごと）。"
               "読めなかった項目は空欄なので入力してください。実施日は、答案に書かれていなければ今日の日付が入っています。"
               "**学校名も科目も空の行は記録されません。**")
    with st.expander(f"📷 写真の番号（{len(imgs)} 枚）"):
        st.markdown("\n".join(f"{i + 1}. {nm}" for i, (_, nm) in enumerate(imgs)))
    df = pd.DataFrame(st.session_state["kk_rows"], columns=KAKOMON_COLUMNS)
    for c in ("得点", "満点"):
        df[c] = pd.to_numeric(df[c], errors="coerce")
    edited = st.data_editor(
        df, num_rows="dynamic", width="stretch",
        key=f"kk_editor_{st.session_state.get('kk_ver', 0)}",
        column_config={
            "写真": st.column_config.TextColumn("写真番号", help="この過去問が写っている写真の番号（例: 1,2）"),
            "実施日": st.column_config.DateColumn("実施日", format="YYYY-MM-DD"),
            "学校名": st.column_config.TextColumn("学校名"),
            "学部": st.column_config.TextColumn("学部・学科"),
            "科目": st.column_config.TextColumn("科目"),
            "方式": st.column_config.TextColumn("方式", help="一般選抜・共通テスト利用・学校推薦型・前期日程・A方式 など"),
            "年度": st.column_config.TextColumn("年度", help="過去問の年度（例: 2024）。令和・平成は西暦に直して記録します"),
            "得点": st.column_config.NumberColumn("得点", min_value=0),
            "満点": st.column_config.NumberColumn("満点", min_value=0),
        })
    rows_now = edited.to_dict("records")

    st.markdown("##### 間違えた問題（分野ごとに結果シートへ記録）")
    st.caption("「過去問」は上の表の何行目の過去問かを表します。分野は、問題冊子を取り込むと"
               "問題の先頭の見出し（→章）と問題番号のタイトル（→節）から読み取ります。"
               "大問も問題番号も空の行は記録されません。")
    wdf = pd.DataFrame(st.session_state.get("kk_wrongs") or [], columns=KK_WRONG_COLUMNS)
    wedited = st.data_editor(
        wdf, num_rows="dynamic", width="stretch",
        key=f"kk_wrong_editor_{st.session_state.get('kk_ver', 0)}",
        column_config={
            "過去問": st.column_config.TextColumn("過去問（上の表の行番号）", help="例: 1"),
            "大問": st.column_config.TextColumn("大問"),
            "問題番号": st.column_config.TextColumn("問題番号"),
            "分野（見出し）": st.column_config.TextColumn("分野（見出し）→章"),
            "タイトル": st.column_config.TextColumn("タイトル → 節"),
        })
    wrongs_now = wedited.to_dict("records")

    c1, c2 = st.columns([1.4, 2])
    with c1:
        do_record = st.button("✅ ⑦ この内容で記録", type="primary", key="record_kakomon")
    with c2:
        if st.button("✖ 破棄", key="discard_kakomon"):
            _clear_review("kakomon")
            st.rerun()
    if not do_record:
        return
    if not student_name:
        st.error("生徒名を選んでください")
        return
    bar = st.progress(0.0, text="記録を開始します…")

    def _prog(i, n, nm):
        bar.progress(i / max(1, n), text=f"写真をDriveへ保存中 {i + 1}/{n}： {nm}")

    try:
        n, summary = record_kakomon(rows_now, student_name, imgs, creds_ui, on_progress=_prog, wrongs=wrongs_now)
    except Exception as e:
        bar.empty()
        st.error(f"記録エラー: {e}")
        return
    if n == 0:
        bar.empty()
        st.warning("学校名・科目が入力された行がないため、記録しませんでした。")
        return
    send_notification_email_plan_b(f"【完了】{student_name} さんの過去問記録（{n}件）", "\n\n".join(summary))
    _finish("kakomon", imgs, f"✅ 過去問 {n} 件を記録しました。スケジュール管理のテスト結果には、"
                             "生徒がダッシュボードを開いたとき（またはスプレッドシート取込ボタン）に反映されます。")


_flash_show()   # 記録・登録などの完了メッセージ（どのタブで操作しても見えるようタブの外に出す）
tab_in, tab_res, tab_master, tab_log = st.tabs(["📥 取込", "📊 集計結果", "⚙️ マスタ管理", "📜 更新履歴"])

# ================================================================== 📥 取込
with tab_in:
    mode = st.radio("① 取込の種類", options=list(KIND_LABELS), format_func=KIND_LABELS.get,
                    horizontal=True, key="kind")
    st.caption(KIND_HELP[mode])

    st.markdown("**② 生徒・科目**")
    _c1, _c2 = st.columns(2)
    with _c1:
        student_name = st.selectbox("生徒名", options=students, index=None,
                                    placeholder="選択してください", key="student_pick")
    with _c2:
        subj_pick = st.selectbox("科目" + ("（写真から読めなかったときに使います）" if mode == "kakomon" else ""),
                                 options=subjects + ["（手入力）"], index=None,
                                 placeholder="選択してください", key="subject_pick")
        subject_name = (st.text_input("科目（手入力）", key="subject_free")
                        if subj_pick == "（手入力）" else (subj_pick or ""))
    st.caption("生徒・科目の追加や削除は「⚙️ マスタ管理」タブで行えます。")

    selected_master_path, text_name, test_title = None, "", ""
    if mode == "text":
        st.markdown("**③ テキスト**")
        text_options = list(master_index.keys())
        if text_options:
            picked = st.selectbox("テキスト名（目次マスタから選択）", options=text_options + ["（手入力）"],
                                  index=None, placeholder="選択してください", key="text_pick")
            text_name = (st.text_input("テキスト名（手入力）", key="text_free")
                         if picked == "（手入力）" else (picked or ""))
        else:
            text_name = st.text_input("テキスト名", key="text_free")
            st.caption("目次マスタが未登録のため、章・節の逆引きはできません（「⚙️ マスタ管理」で登録できます）。")
        with st.expander("詳細設定：採点比較用のマスター画像/PDF（任意）"):
            master_option = st.radio("マスターテキスト", ["💾 保存済みを使う", "🆕 新規アップロード", "❌ 指定しない"],
                                     horizontal=True, key="master_opt")
            if master_option == "💾 保存済みを使う":
                master_files = [f for f in os.listdir(MASTER_DIR) if f.endswith(('.pdf', '.png', '.jpg'))]
                if master_files:
                    selected_master_path = os.path.join(MASTER_DIR, st.selectbox("テキストを選択", master_files,
                                                                                 key="master_saved"))
                else:
                    st.caption("保存済みのマスターはありません。")
            elif master_option == "🆕 新規アップロード":
                um = st.file_uploader("マスターPDF/画像", type=['pdf', 'jpg', 'png'], key="master_up")
                if um:
                    selected_master_path = os.path.join(MASTER_DIR, um.name)
                    with open(selected_master_path, "wb") as f:
                        f.write(um.getvalue())
        if selected_master_path:
            st.caption(f"採点比較用マスター: {os.path.basename(selected_master_path)}")
    elif mode == "confirm":
        st.markdown("**③ テストの情報（任意）**")
        _c3, _c4 = st.columns(2)
        with _c3:
            test_title = st.text_input("テストのタイトル", placeholder="例: 第3回 理解度確認テスト", key="test_title",
                                       help="M列に記録されます。写真から読み取れなかった行にこの値が入ります。")
        with _c4:
            text_name = st.text_input("出題元のテキスト名", key="conf_text",
                                      help="D列に記録されます。写真から読み取れなかった行にこの値が入ります。")
        st.checkbox("📮 送付した理解度確認テストの答案として記録する（L列「提出F」に 1）", key="submit_flag",
                    help="提出F が立った行は、テスト作成システムの手動の復習テスト作成の対象から外れます"
                         "（LINE で回収した答案の復習は自動のループで作成されるため）。")
    else:
        st.markdown("**③ 過去問の情報**")
        st.caption("実施日・学校名・学部・科目・方式・年度・得点は写真から読み取るので、ここでの入力は不要です。"
                   "読み取り後の表で確認・修正できます。")

    st.markdown("**④ 採点済みの写真**")
    imgs = _photo_uploader(mode)
    if imgs:
        st.caption(f"📷 {len(imgs)} 枚（PDFは1ページ＝1枚として数えます）")
    qimgs = _question_uploader(mode)
    if qimgs:
        st.caption(f"📄 問題 {len(qimgs)} 枚：間違えた問題の分野を、問題の先頭の見出し（→章）と問題番号の"
                   "タイトル（→節）から読み取ります" + ("（目次で引けなかった問題だけ）" if mode == "text" else "")
                   + "。取り込んだ後に変えた場合は、もう一度「AIで読み取る」を押してください。")

    st.markdown("**⑤ 読み取り**")
    if st.button("🔍 AIで読み取る", type="primary", key=f"read_{mode}"):
        missing = [label for label, ok in (("生徒名", student_name), ("写真", imgs),
                                           ("テキスト名", text_name or mode != "text")) if not ok]
        if missing:
            st.error("・".join(missing) + " を入力してください")
        else:
            bar = st.progress(0.0, text="読み取りを開始します…")

            def _prog(i, n, nm):
                bar.progress(i / max(1, n), text=(f"写真を送信中 {i + 1}/{n}： {nm}" if mode == "kakomon"
                                                  else f"読み取り中 {i + 1}/{n}： {nm}"))

            try:
                def _qprog(i, n, nm):
                    bar.progress(i / max(1, n), text=f"問題から分野を読み取り中 {i + 1}/{n}： {nm}")

                nf = None
                if mode == "kakomon":
                    wrongs = []
                    rows, errs, used_model = analyze_kakomon(imgs, GEMINI_API_KEY, subject_name,
                                                             on_progress=_prog, wrongs_out=wrongs)
                    if qimgs and wrongs:
                        nf, qerrs = apply_fields_to_kakomon_wrongs(wrongs, rows, qimgs, GEMINI_API_KEY, on_progress=_qprog)
                        errs += qerrs
                    st.session_state["kk_rows"] = rows
                    st.session_state["kk_wrongs"] = wrongs
                    st.session_state["kk_ver"] = st.session_state.get("kk_ver", 0) + 1
                else:
                    rows, errs, used_model = analyze_photos_for_review(
                        imgs, mode, text_name, master_index, GEMINI_API_KEY, selected_master_path,
                        test_title=test_title, on_progress=_prog)
                    if qimgs:
                        nf, qerrs = apply_fields_to_review_rows(rows, qimgs, GEMINI_API_KEY, mode, on_progress=_qprog)
                        errs += qerrs
                    st.session_state["review_rows"] = rows
                    st.session_state["review_mode"] = mode
                    st.session_state["review_ver"] = st.session_state.get("review_ver", 0) + 1
                bar.progress(1.0, text="読み取り完了")
                if errs:
                    st.warning("一部の写真で読み取りに失敗しました： " + " / ".join(errs[:3]))
                st.info(f"使用モデル: {used_model}／{len(rows)} 行を読み取りました。"
                        + (f"問題から {nf} 問の分野を読み取りました。" if nf is not None else "")
                        + "下の表で確認し、必要なら直してから記録してください。")
            except Exception as e:
                bar.empty()
                st.error(f"読み取りエラー: {e}")

    if mode == "kakomon" and st.session_state.get("kk_rows") is not None:
        st.divider()
        render_kakomon_form(imgs, student_name)
    elif (mode != "kakomon" and st.session_state.get("review_rows") is not None
          and st.session_state.get("review_mode") == mode):
        st.divider()
        render_review_form(mode, imgs, student_name, subject_name, text_name, test_title)

# ================================================================== 📊 集計結果
with tab_res:
    if st.button("🔄 最新のデータを読み込む", key="refresh_results"):
        _result_table.clear()
    _r1, _r2, _r3 = st.tabs(["📝 採点記録（テキスト・確認テスト）", "🎓 過去問", "📨 送付テスト（提出状況）"])
    with _r1:
        _df = _result_table(creds_ui, "result")
        if _df.empty:
            st.info("まだ記録がありません。")
        else:
            st.dataframe(_df.iloc[::-1], height=600, width="stretch")
    with _r2:
        _kdf = _result_table(creds_ui, "kakomon")
        if _kdf.empty:
            st.info("まだ過去問の記録がありません。")
        else:
            st.dataframe(_kdf.iloc[::-1], height=600, width="stretch")
    with _r3:
        st.caption("テスト作成システムから LINE で送った理解度確認テスト・復習テスト・過去問の一覧です。"
                   "送付した時点で処理F=1、答案が届くと提出F=1 になり、正解率と状態が更新されます（宿題自動送信が自動で記録）。")
        _ldf = _result_table(creds_ui, "ledger")
        if _ldf.empty:
            st.info("まだ LINE で送付したテストはありません。")
        else:
            _only = st.checkbox("未提出だけ表示", key="ledger_unsubmitted")
            if _only and "提出F" in _ldf.columns:
                _ldf = _ldf[_ldf["提出F"].astype(str).str.strip() != "1"]
            st.dataframe(_ldf.iloc[::-1], height=600, width="stretch", hide_index=True)

# ================================================================== ⚙️ マスタ管理
with tab_master:
    _m1, _m2, _m3, _m4, _m5 = st.tabs(["📚 目次マスタ（PDFから）", "📚 目次マスタ（CSVから）", "👤 生徒", "📕 科目",
                                       "🗂 Drive整理"])

    with _m1:   # ---- PDFから目次マスタを登録（逆引きアプリのPDF解析を統合） ----
        st.caption("テキストのPDFを解析して目次（章・節・ページ）を取り出し、確認・修正してから登録します。")
        pdf_up = st.file_uploader("テキストのPDF", type=["pdf"], key="pdf_master_up")
        pdf_name_in = st.text_input("登録テキスト名（空欄ならファイル名）", key="pdf_name_in")
        use_gemini = st.checkbox("🤖 AI画像解析（Gemini）を使う（文字化けPDF・複雑な目次向け／高精度）",
                                 key="pdf_use_gemini")
        if pdf_up is not None and st.button("🔎 PDFを解析", key="pdf_analyze"):
            try:
                _doc = fitz.open(stream=pdf_up.getvalue(), filetype="pdf")
                _name = (pdf_name_in or "").strip() or pdf_up.name.rsplit(".", 1)[0]
                if use_gemini:
                    _rows, _method = analyze_pdf_gemini(_doc, _name, GEMINI_API_KEY)
                else:
                    _rows, _method = analyze_pdf(_doc, _name)
                    if (not _rows) or _method.startswith("解析不可"):
                        try:
                            gr, gm = analyze_pdf_gemini(_doc, _name, GEMINI_API_KEY)
                            if gr:
                                _rows, _method = gr, gm + "（自動切替）"
                        except Exception as ge:
                            st.warning(f"AI画像解析に失敗: {ge}")
                st.session_state["pdf_rows"] = _rows
                st.session_state["pdf_name"] = _name
                st.session_state["pdf_method"] = _method
                st.success(f"{len(_rows)} 件抽出しました（方式: {_method}）。下の表で確認・修正して登録してください。")
            except Exception as e:
                st.error(f"解析エラー: {e}")
        if st.session_state.get("pdf_rows"):
            st.caption(f"テキスト名: {st.session_state['pdf_name']} ／ 方式: {st.session_state.get('pdf_method', '')}")
            _pdf_df = pd.DataFrame([{"章": r.get("chapter", ""), "節": r.get("section", ""),
                                     "節タイトル": r.get("title", ""), "開始ページ": r.get("start"),
                                     "終了ページ": r.get("end")} for r in st.session_state["pdf_rows"]])
            _pdf_edited = st.data_editor(_pdf_df, num_rows="dynamic", width="stretch", key="pdf_editor")
            if st.button("✅ このテキストを目次マスタに登録", type="primary", key="pdf_register"):
                try:
                    _out = _pdf_edited.copy()
                    _out.insert(0, "テキスト名", st.session_state["pdf_name"])
                    n_t, n_r = register_master_csv(creds_ui, _out.to_csv(index=False).encode("utf-8-sig"))
                    for k in ("pdf_rows", "pdf_name", "pdf_method"):
                        st.session_state.pop(k, None)
                    st.session_state["flash"] = ("success", f"目次マスタに登録しました（{n_t} テキスト / {n_r} 行）。")
                    st.rerun()
                except Exception as e:
                    st.error(f"登録エラー: {e}")

    with _m2:   # ---- 目次CSVから登録 ----
        if master_index:
            st.success("登録済みテキスト：\n- " + "\n- ".join(master_index.keys()))
        else:
            st.warning("未登録です。逆引きアプリの『マスタを書き出す』で出力したCSVを登録してください。")
        ups = st.file_uploader("目次CSV を登録（複数可・章節リストCSVも可）", type=["csv"],
                               accept_multiple_files=True, key="master_csv")
        st.caption("テキスト名の列が無いCSV（例: 『〇〇_章節リスト.csv』）は、ファイル名をテキスト名として登録します。")
        if ups and st.button("⬆️ マスタに登録（目次マスタタブへ保存）", key="csv_register"):
            tot_t, tot_r, errs = 0, 0, []
            for up in ups:
                try:
                    nt, nr = register_master_csv(creds_ui, up.getvalue(),
                                                 default_text_name=_text_name_from_filename(up.name))
                    tot_t += nt
                    tot_r += nr
                except Exception as e:
                    errs.append(f"{up.name}: {e}")
            if errs:
                st.error("一部失敗： " + " / ".join(errs))
            elif tot_r:
                st.session_state["flash"] = ("success", f"{tot_t} テキスト / {tot_r} 行を登録しました。")
                st.rerun()

    with _m3:   # ---- 生徒 ----
        _ns = st.text_input("生徒名を追加", key="new_student",
                            help="スケジュール管理のログインID（または登録氏名）と同じ表記にしてください。")
        if st.button("➕ 追加して Drive フォルダを作成", key="add_student"):
            _nm = (_ns or "").strip()
            if not _nm:
                st.warning("生徒名を入力してください")
            else:
                add_list_item(creds_ui, STUDENT_TAB, "生徒名", _nm)   # 名簿へ（重複は無視）
                try:
                    _fid, _created = ensure_drive_folder(_nm, creds_ui)
                    st.session_state["flash"] = ("success", f"「{_nm}」を追加しました"
                                                 f"（Driveフォルダ：{'新規作成' if _created else '既存を使用'}）")
                except Exception as _e:
                    st.session_state["flash"] = ("warning", f"名簿には追加しましたが、Driveフォルダ作成でエラー: {_e}")
                st.rerun()
        _ds = st.selectbox("削除する生徒", options=["（選択）"] + students, key="del_student")
        if st.button("🗑 生徒を削除", key="del_student_btn") and _ds and _ds != "（選択）":
            remove_list_item(creds_ui, STUDENT_TAB, "生徒名", _ds)
            st.session_state["flash"] = ("success", f"「{_ds}」を名簿から削除しました（Driveのフォルダは残ります）")
            st.rerun()

    with _m4:   # ---- 科目 ----
        _nsub = st.text_input("科目を追加", key="new_subject")
        if st.button("➕ 科目を追加", key="add_subject"):
            if add_list_item(creds_ui, SUBJECT_TAB, "科目", _nsub):
                st.session_state["flash"] = ("success", "科目を追加しました")
                st.rerun()
            else:
                st.warning("空欄、または既に登録済みです")
        _dsub = st.selectbox("削除する科目", options=["（選択）"] + subjects, key="del_subject")
        if st.button("🗑 科目を削除", key="del_subject_btn") and _dsub and _dsub != "（選択）":
            remove_list_item(creds_ui, SUBJECT_TAB, "科目", _dsub)
            st.session_state["flash"] = ("success", f"「{_dsub}」を削除しました")
            st.rerun()

    with _m5:   # ---- 生徒フォルダ直下の写真を「実施済答案」へ移す（一度だけ使う整理ツール） ----
        st.caption(f"生徒フォルダの直下に置かれている答案写真（画像ファイル）を、各生徒の「{DONE_FOLDER}」フォルダへ移動します。"
                   "画像以外のファイルやサブフォルダは動かしません。今後の記録は最初から「実施済答案」に保存されます。")
        if st.button("🔍 移動する写真を調べる", key="plan_move"):
            with st.spinner("Drive を確認しています…"):
                try:
                    st.session_state["move_plan"] = plan_move_to_done(creds_ui)
                except Exception as e:
                    st.error(f"確認エラー: {e}")
        _plan = st.session_state.get("move_plan")
        if _plan is not None:
            _total = sum(len(x["files"]) for x in _plan)
            if not _total:
                st.success("生徒フォルダの直下に答案写真はありません。整理済みです。")
            else:
                st.dataframe(pd.DataFrame([{"生徒": x["student"], "移動する写真": len(x["files"]),
                                            "例": x["files"][0]["name"]} for x in _plan]),
                             width="stretch", hide_index=True)
                if st.button(f"📦 {_total} 枚を「{DONE_FOLDER}」へ移動する", type="primary", key="do_move"):
                    _bar = st.progress(0.0, text="移動を開始します…")

                    def _mprog(i, n, nm):
                        _bar.progress(i / max(1, n), text=f"移動中 {i + 1}/{n}： {nm}")

                    _n, _errs = move_to_done(_plan, creds_ui, on_progress=_mprog)
                    st.session_state.pop("move_plan", None)
                    st.session_state["flash"] = (("warning" if _errs else "success"),
                                                 f"{_n} 枚を「{DONE_FOLDER}」へ移動しました。"
                                                 + (f"失敗 {len(_errs)} 件: " + " / ".join(_errs[:3]) if _errs else ""))
                    st.rerun()

# ================================================================== 📜 更新履歴
with tab_log:
    if APP_CHANGELOG:
        st.caption(f"現在のバージョン: Ver. {APP_VERSION}（{APP_UPDATED} 更新）　／　"
                   "サーバー上の CHANGELOG.md を表示しています。")
        st.markdown(APP_CHANGELOG)
    else:
        st.info("CHANGELOG.md が見つかりません。")
