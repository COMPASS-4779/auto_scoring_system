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
import datetime
import json
import re
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

STUDENT_LIST = ["上原百華", "上原遥人", "浅井渉", "荒木陽向", "谷川瑠依", "momokauehara"]
os.makedirs(MASTER_DIR, exist_ok=True)


# ==========================================
# [統合] テキスト目次マスタ（Google Sheet 永続化）
# ==========================================
def _to_int(v):
    m = re.search(r"\d+", str(v)) if v is not None else None
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

def save_to_spreadsheet(student_name, subject, text_name, section_results, drive_link, creds, master_index):
    """[統合] ページ番号からマスタ逆引きし 章/節/節タイトル を付けて結果へ書き込む。"""
    service = _sheets(creds)
    now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    values = []
    for s in section_results:
        ai_chapter = s.get('chapter', '')
        ai_section = s.get('section', '')
        total = s.get('total', 0)
        sec_page = s.get('page', '')
        for p in s.get('wrong', []):
            page = p.get('page', '') or sec_page or '-'
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

        for (photo_filepath, photo_name) in photos_data:
            try:
                common_rules = (
                    "【重要な判断基準】\n"
                    "・×や✗、赤ペンで訂正されている問題 = 間違い\n"
                    "・○や無印の問題 = 正解（wrongに含めない）\n"
                    "・計算の途中式や答えの数値（例: -8, 3/4）は問題番号ではない\n"
                    "・問題番号は「1」「(2)」「問3」のような番号表記のみ\n"
                    "・写真内に印刷されている『ページ番号』を読み取り、各間違い問題の page に半角数字で入れる。"
                    "読めない場合のみ \"-\"。\n\n"
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

                # [統合] 保存名の先頭に 章_節_節タイトル ヘッダーを付与
                header_label = ""
                for s in section_results:
                    for w in s.get('wrong', []):
                        ch, se, ti = lookup_section(master_index, text_name, w.get('page', ''))
                        if se or ti:
                            header_label = f"[{ch}_{se}_{ti}]".replace("/", "／")
                            break
                    if header_label:
                        break
                save_name = (header_label + "_" + photo_name) if header_label else photo_name
                drive_link = upload_to_drive(photo_filepath, save_name, folder_id, creds)

                save_to_spreadsheet(student_name, subject_name, text_name, section_results, drive_link, creds, master_index)
                send_notification_email_plan_b(f"【進捗】記録完了 ({photo_name})",
                                               json.dumps(section_results, ensure_ascii=False, indent=2) + f"\n\nリンク: {drive_link}")
            finally:
                try: os.remove(photo_filepath)
                except: pass
        send_notification_email_plan_b("【完了】全処理終了", f"{student_name} さんの全画像処理が完了しました。")
    except Exception as e:
        send_notification_email_plan_b("【警告】システムエラー", f"エラー内容: {e}")


# ==========================================
# Streamlit Web UI
# ==========================================
st.title("📝 採点済みプリント 自動集計システム（逆引き統合版）")

creds_ui = Credentials.from_authorized_user_info(GOOGLE_TOKEN_DICT)
master_index = load_master_index(creds_ui)  # [統合] Sheetから常時参照

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

col_left, col_right = st.columns([1, 1])

with col_left:
    st.subheader("👤 講師用アップロード画面")
    student_name = st.selectbox("生徒名", options=STUDENT_LIST, index=None)
    subject_name = st.text_input("科目")
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

    uploaded_photos = st.file_uploader("採点済み写真", type=['jpg', 'jpeg', 'png'], accept_multiple_files=True)

    if st.button("🚀 送信して完了", type="primary"):
        if not student_name or not uploaded_photos or not text_name:
            st.error("生徒名・テキスト名・写真は必須です")
        else:
            photos_data = []
            for photo in uploaded_photos:
                tmp = os.path.join(tempfile.gettempdir(), f"photo_{uuid.uuid4().hex}.jpg")
                with open(tmp, "wb") as f: f.write(photo.getvalue())
                photos_data.append((tmp, photo.name))
            threading.Thread(
                target=background_processing_task,
                args=(student_name, subject_name, text_name, selected_master_path, photos_data, GEMINI_API_KEY, GOOGLE_TOKEN_DICT, master_index)
            ).start()
            st.success("✅ 受付完了！進捗はメールで通知されます。")
            st.balloons()

with col_right:
    st.subheader("📊 現在の集計結果")
    if st.button("🔄 データを更新"): st.rerun()
    df = get_spreadsheet_data(creds_ui)
    if not df.empty:
        st.dataframe(df.iloc[::-1], height=600, width='stretch')
