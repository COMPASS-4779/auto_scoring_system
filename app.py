# -*- coding: utf-8 -*-
# =====================================================================
# auto_scoring_system  app.py  ―― 逆引きアプリ統合版（テキスト目次マスタ内蔵）
# =====================================================================
# 逆引きアプリ（ページ番号→章/節/節タイトル）の機能を採点集計システムにマージ。
#
# 【統合した機能】
#  1. テキスト目次マスタを Google スプレッドシートの「目次マスタ」タブに永続保存
#     （Streamlit Cloud は再起動でファイルが消えるため、Sheet に保存して常時参照）
#  2. 逆引きアプリで書き出した「テキスト目次マスタ.csv」をアップロードして登録
#     （列: テキスト名, 章, 節, 節タイトル, 開始ページ, 終了ページ／同名テキストは置換）
#  3. 採点写真の解析時、Gemini が読み取った印刷ページ番号からマスタを逆引きし、
#     正確な 章 / 節 / 節タイトル を結果スプレッドシートに書き込む
#
# 【結果スプレッドシートの列】 A:K
#  日時, 生徒名, 科目, テキスト名, ページ, 章(マスタ), 節(マスタ), 問題番号, 写真リンク, 総問題数, 節タイトル(マスタ)
#
# 【Secrets】 元のまま: GEMINI_API_KEY / SENDER_EMAIL / APP_PASSWORD / GOOGLE_TOKEN_JSON
#
# ※ 実環境（Gemini/Google認証/Streamlit）が無いため未実行です。構文・CSV/逆引きロジックは検証済み。
#    デプロイ前に必ずテスト実行してください。
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
import threading
import fitz  # PyMuPDF
import pandas as pd
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
st.set_page_config(page_title="AI集計システム（逆引き統合版）", layout="wide")

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

SPREADSHEET_ID = "1B8BKKY8SfR-V3ysirsNG6fqlrVzXqPBF_AdjFDc5fCc"
PARENT_FOLDER_ID = "1DS7anMs-ruhTtVxZNqsVhZSbeQFCww_2"
MASTER_DIR = "master_texts"
NOTIFICATION_EMAIL = "info@compassesonline.com"
MASTER_TAB = "目次マスタ"  # [統合] テキスト目次マスタを保存するタブ名
MASTER_HEADER = ["テキスト名", "章", "節", "節タイトル", "開始ページ", "終了ページ"]
STUDENT_TAB = "生徒名簿"   # [統合] 生徒名を保存するタブ
SUBJECT_TAB = "科目マスタ"  # [統合] 科目を保存するタブ
DEFAULT_SUBJECTS = ["国語", "数学", "英語", "英文法", "古文", "理科", "社会"]

STUDENT_LIST = ["上原百華", "上原遥人", "浅井渉", "荒木陽向", "谷川瑠依", "momokauehara"]
os.makedirs(MASTER_DIR, exist_ok=True)


# ==========================================
# [統合] テキスト目次マスタ（Google Sheet 永続化）
# ==========================================
def _to_int(v):
    if v is None: return None
    m = re.search(r"\d+", str(v).translate(str.maketrans("０１２３４５６７８９", "0123456789")))
    return int(m.group()) if m else None

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

def register_master_csv(creds, file_bytes):
    """逆引きアプリの「テキスト目次マスタ.csv」を取り込み、目次マスタ タブへ反映（同名テキストは置換）。"""
    text = file_bytes.decode("utf-8-sig", errors="replace")
    reader = csv.DictReader(io.StringIO(text))
    new_rows, new_names = [], set()
    for row in reader:
        def col(*names):
            for n in names:
                if n in row and row[n] is not None:
                    return str(row[n]).strip()
            return ""
        tname = col("テキスト名", "教材名")
        if not tname:
            continue
        new_names.add(tname)
        new_rows.append([
            tname, col("章"), col("節"), col("節タイトル", "タイトル"),
            col("開始ページ", "ページ"), col("終了ページ"),
        ])
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
            m = re.search(r'\[.*\]', resp.text, re.DOTALL)
            if m:
                for o in json.loads(m.group(0)):
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

def get_drive_folder_id(student_name, creds):
    service = build('drive', 'v3', credentials=creds)
    query = f"'{PARENT_FOLDER_ID}' in parents and name = '{student_name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    results = service.files().list(q=query, fields="files(id)").execute()
    folders = results.get('files', [])
    return folders[0]['id'] if folders else None

def upload_to_drive(filepath, filename, folder_id, creds):
    service = build('drive', 'v3', credentials=creds)
    media = MediaFileUpload(filepath, mimetype='image/jpeg', resumable=True)
    file = service.files().create(body={'name': filename, 'parents': [folder_id]}, media_body=media, fields='webViewLink').execute()
    return file.get('webViewLink')

def save_to_spreadsheet(student_name, subject, text_name, section_results, drive_link, creds, master_index, user_page=None):
    """[統合] ページ番号（ユーザー入力優先）からマスタ逆引きし 章/節/節タイトル を付けて結果へ書き込む。"""
    service = _sheets(creds)
    now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    up = (str(user_page).strip() if user_page not in (None, "") else "")
    values = []
    for s in section_results:
        ai_chapter = s.get('chapter', '')
        ai_section = s.get('section', '')
        total = s.get('total', 0)
        sec_page = s.get('page', '')
        for p in s.get('wrong', []):
            page = up or p.get('page', '') or sec_page or '-'   # ユーザー入力ページを最優先
            m_ch, m_sec, m_title = lookup_section(master_index, text_name, page)
            chapter = m_ch or ai_chapter   # マスタ優先、無ければAI推定
            section = m_sec or ai_section
            values.append([
                now, student_name, subject, text_name,
                page, chapter, section,
                p.get('number', '-'), drive_link, total, m_title,
            ])
    if values:
        service.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_ID, range='A1',
            valueInputOption='USER_ENTERED', body={'values': values}
        ).execute()

def get_spreadsheet_data(creds):
    try:
        res = _sheets(creds).spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range='A:K').execute()
        rows = res.get('values', [])
        if not rows:
            return pd.DataFrame()
        header = rows[0]
        data = [r + [''] * (len(header) - len(r)) for r in rows[1:]]
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
    preferred = ['gemini-2.5-flash-preview-05-20', 'gemini-2.5-pro-exp-03-25',
                 'gemini-2.5-flash', 'gemini-2.5-pro', 'gemini-1.5-pro-latest', 'gemini-1.5-pro']
    try:
        available = [m.name.replace('models/', '') for m in client.models.list()]
        for model in preferred:
            if model in available:
                return model
    except Exception:
        pass
    return 'gemini-1.5-pro'

def background_processing_task(student_name, subject_name, text_name, selected_master_path, photos_data, api_key, token_dict, master_index):
    try:
        creds = Credentials.from_authorized_user_info(token_dict)
        client = genai.Client(api_key=api_key)
        folder_id = get_drive_folder_id(student_name, creds)
        best_model = get_best_model(client)
        send_notification_email_plan_b("【進捗】処理開始", f"生徒: {student_name}／使用モデル: {best_model}")

        ai_master_files = process_master_file_from_path(selected_master_path, client) if selected_master_path else []

        for _item in photos_data:
            photo_filepath, photo_name = _item[0], _item[1]
            user_page = _item[2] if len(_item) > 2 else None
            try:
                common_rules = (
                    "【重要な判断基準】\n"
                    "・×や✗、赤ペンで訂正されている問題 = 間違い\n"
                    "・○や無印の問題 = 正解（wrongに含めない）\n"
                    "・計算の途中式や答えの数値（例: -8, 3/4）は問題番号ではない\n"
                    "・問題番号は「1」「(2)」「問3」「(ア)」「①」のような番号表記。"
                    "手書きで番号（ア・イ・ウ や (1)・① 等）が振られている場合は、その手書き番号を正として優先的に読み取る。\n"
                    "・写真内に印刷されている『ページ番号』を読み取り、各間違い問題の page に半角数字で入れる。"
                    "読めない場合のみ \"-\"（※最終的なページ番号は利用者入力を優先します）。\n\n"
                    f"【出力形式】\n"
                    f"chapterは常に \"{text_name}\"。sectionは項目内容。totalは総問題数。"
                    "wrongは間違いのみで page と number を入れる。\n\n"
                    "[{\"chapter\": \"" + text_name + "\", \"section\": \"項目名\", \"total\": 4, "
                    "\"wrong\": [{\"page\": \"8\", \"number\": \"(1)\"}]}]"
                )
                ai_photo = client.files.upload(file=photo_filepath)
                while ai_photo.state.name == 'PROCESSING': time.sleep(1); ai_photo = client.files.get(name=ai_photo.name)

                if ai_master_files:
                    prompt = "採点済み答案とマスター（正解）を比較し、sectionごとに総問題数と間違いをJSONで返す。\n\n" + common_rules
                    contents = ai_master_files + [ai_photo, prompt]
                else:
                    prompt = "採点済み答案を見て、sectionごとに総問題数と間違いをJSONで返す。\n\n" + common_rules
                    contents = [ai_photo, prompt]

                response = client.models.generate_content(model=best_model, contents=contents)
                match = re.search(r'\[.*\]', response.text, re.DOTALL)
                section_results = json.loads(match.group(0)) if match else []

                # [統合] ヘッダー（章_節_節タイトル）はユーザー入力ページから判定し、答案用紙を区別
                ch, se, ti = lookup_section(master_index, text_name, user_page)
                if not (se or ti):  # 入力が無ければGeminiの読み取りページで代替
                    for s in section_results:
                        for w in s.get('wrong', []):
                            ch, se, ti = lookup_section(master_index, text_name, w.get('page', ''))
                            if se or ti: break
                        if se or ti: break
                header_label = f"[{ch}_{se}_{ti}]".replace("/", "／") if (se or ti) else ""
                pg_label = (str(user_page).strip() if user_page not in (None, "") else "")
                save_name = ((header_label + ("_p" + pg_label if pg_label else "") + "_") if header_label else "") + photo_name
                drive_link = upload_to_drive(photo_filepath, save_name, folder_id, creds)

                save_to_spreadsheet(student_name, subject_name, text_name, section_results, drive_link, creds, master_index, user_page=user_page)
                send_notification_email_plan_b(f"【進捗】記録完了 ({photo_name})",
                                               json.dumps(section_results, ensure_ascii=False, indent=2) + f"\n\nリンク: {drive_link}")
            finally:
                try: os.remove(photo_filepath)
                except: pass
        send_notification_email_plan_b("【完了】全処理終了", f"{student_name} さんの全画像処理が完了しました。")
    except Exception as e:
        send_notification_email_plan_b("【警告】システムエラー", f"エラー内容: {e}")


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
        if ext in ('jpg', 'jpeg', 'png', 'bmp', 'webp', 'gif'):
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
st.title("📝 採点済みプリント 自動集計システム（逆引き統合版）")

# ---- PDFから目次マスタを登録（逆引きアプリのPDF解析を統合） ----
with st.expander("📄 PDFから目次マスタを登録（自動解析→確認・修正→登録）", expanded=False):
    _creds_pdf = Credentials.from_authorized_user_info(GOOGLE_TOKEN_DICT)
    pdf_up = st.file_uploader("テキストのPDF", type=["pdf"], key="pdf_master_up")
    pdf_name_in = st.text_input("登録テキスト名（空欄ならファイル名）", key="pdf_name_in")
    use_gemini = st.checkbox("🤖 AI画像解析（Gemini）を使う（文字化けPDF・複雑な目次向け／高精度）", key="pdf_use_gemini")
    if pdf_up is not None and st.button("🔎 PDFを解析"):
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
        st.caption(f"テキスト名: {st.session_state['pdf_name']} ／ 方式: {st.session_state.get('pdf_method','')}")
        _df = pd.DataFrame([{ "章": r.get("chapter", ""), "節": r.get("section", ""),
                              "節タイトル": r.get("title", ""), "開始ページ": r.get("start"),
                              "終了ページ": r.get("end") } for r in st.session_state["pdf_rows"]])
        _edited = st.data_editor(_df, num_rows="dynamic", width="stretch", key="pdf_editor")
        if st.button("✅ このテキストを目次マスタに登録", type="primary"):
            try:
                _out = _edited.copy()
                _out.insert(0, "テキスト名", st.session_state["pdf_name"])
                _csv = _out.to_csv(index=False).encode("utf-8-sig")
                n_t, n_r = register_master_csv(_creds_pdf, _csv)
                st.success(f"目次マスタに登録しました（{n_t} テキスト / {n_r} 行）。")
                for k in ("pdf_rows", "pdf_name", "pdf_method"):
                    st.session_state.pop(k, None)
                st.rerun()
            except Exception as e:
                st.error(f"登録エラー: {e}")


creds_ui = Credentials.from_authorized_user_info(GOOGLE_TOKEN_DICT)
master_index = load_master_index(creds_ui)  # [統合] Sheetから常時参照
students = load_list(creds_ui, STUDENT_TAB, "生徒名", STUDENT_LIST)   # [統合] 生徒名簿
subjects = load_list(creds_ui, SUBJECT_TAB, "科目", DEFAULT_SUBJECTS)  # [統合] 科目マスタ

# ---- サイドバー：テキスト目次マスタ管理 ----
with st.sidebar:
    st.subheader("📚 テキスト目次マスタ")
    if master_index:
        st.success("登録済みテキスト：\n- " + "\n- ".join(master_index.keys()))
    else:
        st.warning("未登録です。逆引きアプリの『マスタを書き出す』で出力したCSVを登録してください。")
    up = st.file_uploader("テキスト目次マスタ.csv を登録", type=["csv"], key="master_csv")
    if up is not None and st.button("⬆️ マスタに登録（目次マスタタブへ保存）"):
        try:
            n_text, n_row = register_master_csv(creds_ui, up.getvalue())
            st.success(f"{n_text} テキスト / {n_row} 行を登録しました。")
            st.rerun()
        except Exception as e:
            st.error(f"登録エラー: {e}")

    st.divider()
    st.subheader("👤 生徒名の管理")
    ns = st.text_input("生徒名を追加", key="new_student")
    if st.button("➕ 生徒を追加"):
        if add_list_item(creds_ui, STUDENT_TAB, "生徒名", ns):
            st.success("追加しました"); st.rerun()
        else:
            st.warning("空欄、または既に登録済みです")
    ds = st.selectbox("削除する生徒", options=["（選択）"] + students, key="del_student")
    if st.button("🗑 生徒を削除") and ds and ds != "（選択）":
        remove_list_item(creds_ui, STUDENT_TAB, "生徒名", ds); st.success("削除しました"); st.rerun()

    st.divider()
    st.subheader("📕 科目の管理")
    nsub = st.text_input("科目を追加", key="new_subject")
    if st.button("➕ 科目を追加"):
        if add_list_item(creds_ui, SUBJECT_TAB, "科目", nsub):
            st.success("追加しました"); st.rerun()
        else:
            st.warning("空欄、または既に登録済みです")
    dsub = st.selectbox("削除する科目", options=["（選択）"] + subjects, key="del_subject")
    if st.button("🗑 科目を削除") and dsub and dsub != "（選択）":
        remove_list_item(creds_ui, SUBJECT_TAB, "科目", dsub); st.success("削除しました"); st.rerun()

col_left, col_right = st.columns([1, 1])

with col_left:
    st.subheader("👤 講師用アップロード画面")
    student_name = st.selectbox("生徒名", options=students, index=None)
    subj_pick = st.selectbox("科目", options=subjects + ["（手入力）"], index=None)
    subject_name = st.text_input("科目（手入力）") if subj_pick == "（手入力）" else (subj_pick or "")
    # [統合] テキスト名は登録済みマスタから選択可（手入力も可）
    text_options = list(master_index.keys())
    if text_options:
        picked = st.selectbox("テキスト名（マスタから選択／逆引き対象）", options=text_options + ["（手入力）"], index=None)
        text_name = st.text_input("テキスト名（手入力）") if picked == "（手入力）" else (picked or "")
    else:
        text_name = st.text_input("テキスト名")

    master_option = st.radio("マスターテキスト（採点比較用の画像/PDF）", ["💾 保存済みを使う", "🆕 新規アップロード", "❌ 指定しない"])
    selected_master_path = None
    if master_option == "💾 保存済みを使う":
        master_files = [f for f in os.listdir(MASTER_DIR) if f.endswith(('.pdf', '.png', '.jpg'))]
        if master_files:
            selected_master_path = os.path.join(MASTER_DIR, st.selectbox("テキストを選択", master_files))
    elif master_option == "🆕 新規アップロード":
        um = st.file_uploader("マスターPDF/画像", type=['pdf', 'jpg', 'png'])
        if um:
            selected_master_path = os.path.join(MASTER_DIR, um.name)
            with open(selected_master_path, "wb") as f: f.write(um.getvalue())

    uploaded_photos = st.file_uploader("採点済み写真／PDF／Word（複数可）", type=['jpg', 'jpeg', 'png', 'pdf', 'docx', 'doc'], accept_multiple_files=True)

    # [統合] アップロード内容が変わった時だけ画像へ展開（PDF=各ページ、Word=埋め込み画像）
    if uploaded_photos:
        sig = tuple((f.name, f.size) for f in uploaded_photos)
        if st.session_state.get("img_sig") != sig:
            imgs = []
            for f in uploaded_photos:
                imgs.extend(expand_uploaded_to_images(f))
            st.session_state["pending_images"] = imgs   # [(path, name), ...]
            st.session_state["img_sig"] = sig
            st.session_state["page_df"] = pd.DataFrame([{"画像": n, "ページ番号": ""} for (_pp, n) in imgs])
    else:
        for k in ("pending_images", "img_sig", "page_df"):
            st.session_state.pop(k, None)

    if st.session_state.get("pending_images"):
        st.markdown("**📄 各画像（答案用紙）のページ番号を入力してください**　"
                    "— このページ番号からテキスト目次マスタを逆引きして、章・節・節タイトルを記入します。")
        _imgs0 = st.session_state["pending_images"]
        _cols = st.columns(2) if len(_imgs0) > 1 else [st]
        for idx, (path, name) in enumerate(_imgs0):
            with (_cols[idx % len(_cols)]):
                st.text_input(f"{idx+1}. {name}", key=f"pgin_{idx}", placeholder="ページ番号（例: 8）")

    if st.button("🚀 送信して完了", type="primary"):
        if not student_name or not st.session_state.get("pending_images") or not text_name:
            st.error("生徒名・テキスト名・画像は必須です")
        else:
            imgs = st.session_state["pending_images"]
            missing = [name for i, (path, name) in enumerate(imgs) if not str(st.session_state.get(f"pgin_{i}", "") or "").strip()]
            photos_data = []
            for i, (path, name) in enumerate(imgs):
                pg = str(st.session_state.get(f"pgin_{i}", "") or "").strip()
                photos_data.append((path, name, pg))
            if missing:
                st.warning("ページ番号が未入力の画像があります（そのまま送信すると章・節は空になります）：" + " / ".join(missing[:5]))
            threading.Thread(
                target=background_processing_task,
                args=(student_name, subject_name, text_name, selected_master_path, photos_data, GEMINI_API_KEY, GOOGLE_TOKEN_DICT, master_index)
            ).start()
            st.success(f"✅ 受付完了！（{len(photos_data)} 枚の画像を処理します）進捗はメールで通知されます。")
            st.balloons()
            for k in ("pending_images", "img_sig", "page_df"):
                st.session_state.pop(k, None)

with col_right:
    st.subheader("📊 現在の集計結果")
    if st.button("🔄 データを更新"): st.rerun()
    df = get_spreadsheet_data(creds_ui)
    if not df.empty:
        st.dataframe(df.iloc[::-1], height=600, width='stretch')
