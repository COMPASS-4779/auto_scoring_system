import streamlit as st
import os
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
# 🌟 設定情報（原因特定・強化版）
# ==========================================
st.set_page_config(page_title="AI集計システム", layout="wide")

# 足りない鍵を具体的にチェックする機能
missing_keys = []
for key in ["GEMINI_API_KEY", "SENDER_EMAIL", "APP_PASSWORD", "GOOGLE_TOKEN_JSON"]:
    if key not in st.secrets:
        missing_keys.append(key)

if missing_keys:
    st.error(f"🚨 金庫（Secrets）の中に以下の鍵が見つかりません: {', '.join(missing_keys)}")
    st.info("Streamlit右下の「Manage app」 > 「Settings」 > 「Secrets」を開き、名前や入力形式を確認してください。")
    st.stop()

try:
    GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
    SENDER_EMAIL = st.secrets["SENDER_EMAIL"]
    APP_PASSWORD = st.secrets["APP_PASSWORD"]
    GOOGLE_TOKEN_JSON_STR = st.secrets["GOOGLE_TOKEN_JSON"]
    # 🌟 token.jsonの文字が崩れていないかチェック
    GOOGLE_TOKEN_DICT = json.loads(GOOGLE_TOKEN_JSON_STR)
except json.JSONDecodeError as e:
    st.error(f"🚨 token.json の読み込みに失敗しました。貼り付けた中身が途中で途切れているか、余計な文字が入っている可能性があります。\n詳細エラー: {e}")
    st.stop()
except Exception as e:
    st.error(f"🚨 予期せぬエラー: {e}")
    st.stop()

SPREADSHEET_ID = "1B8BKKY8SfR-V3ysirsNG6fqlrVzXqPBF_AdjFDc5fCc"
PARENT_FOLDER_ID = "1DS7anMs-ruhTtVxZNqsVhZSbeQFCww_2"
MASTER_DIR = "master_texts"
NOTIFICATION_EMAIL = "info@compassesonline.com"

STUDENT_LIST = ["上原百華", "上原遥人", "浅井渉", "荒木陽向", "谷川瑠依", "momokauehara"]
os.makedirs(MASTER_DIR, exist_ok=True)

# ==========================================
# 🌟 各種関数
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

def save_to_spreadsheet(student_name, subject, text_name, wrong_problems, drive_link, creds):
    service = build('sheets', 'v4', credentials=creds)
    now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    values = []
    if not isinstance(wrong_problems, list): wrong_problems = [wrong_problems]
    for p in wrong_problems:
        if isinstance(p, dict):
            values.append([now, student_name, subject, text_name, p.get('page','-'), p.get('chapter','-'), p.get('section','-'), p.get('number','-'), drive_link])
        else:
            values.append([now, student_name, subject, text_name, '-', '-', '-', str(p), drive_link])
    if values:
        service.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_ID, range='A1',
            valueInputOption='USER_ENTERED', body={'values': values}
        ).execute()

def get_spreadsheet_data(creds):
    try:
        service = build('sheets', 'v4', credentials=creds)
        result = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range='A:I').execute()
        rows = result.get('values', [])
        if not rows: return pd.DataFrame()
        return pd.DataFrame(rows[1:], columns=rows[0])
    except: return pd.DataFrame()

def process_master_file_from_path(filepath, client):
    ai_files = []
    if filepath.lower().endswith('.pdf'):
        doc = fitz.open(filepath)
        for i in range(len(doc)):
            page = doc.load_page(i); pix = page.get_pixmap(dpi=150)
            tmp_img_path = os.path.join(tempfile.gettempdir(), f"master_{uuid.uuid4().hex}.png")
            pix.save(tmp_img_path)
            try:
                ai_file = client.files.upload(file=tmp_img_path)
                while ai_file.state.name == 'PROCESSING': time.sleep(2); ai_file = client.files.get(name=ai_file.name)
                if ai_file.state.name == 'ACTIVE': ai_files.append(ai_file)
            finally:
                try: os.remove(tmp_img_path)
                except: pass
    else:
        ai_file = client.files.upload(file=filepath)
        while ai_file.state.name == 'PROCESSING': time.sleep(2); ai_file = client.files.get(name=ai_file.name)
        if ai_file.state.name == 'ACTIVE': ai_files.append(ai_file)
    return ai_files

def background_processing_task(student_name, subject_name, text_name, selected_master_path, photos_data, api_key, token_dict):
    try:
        creds = Credentials.from_authorized_user_info(token_dict)
        client = genai.Client(api_key=api_key)
        folder_id = get_drive_folder_id(student_name, creds)
        
        send_notification_email_plan_b("【進捗】AI集計システム 処理開始", f"生徒: {student_name} さんの処理を開始しました。")

        ai_master_files = []
        if selected_master_path:
            ai_master_files = process_master_file_from_path(selected_master_path, client)

        for i, (photo_filepath, photo_name) in enumerate(photos_data):
            try:
                drive_link = upload_to_drive(photo_filepath, photo_name, folder_id, creds)
                ai_photo = client.files.upload(file=photo_filepath)
                while ai_photo.state.name == 'PROCESSING': time.sleep(1); ai_photo = client.files.get(name=ai_photo.name)
                
                if ai_master_files:
                    prompt = "マスターと比較して間違っている問題を抽出してJSONで返してください。\n[{\"page\": \"12\", \"chapter\": \"第1章\", \"section\": \"五感（3）聴覚\", \"number\": \"問2\"}]"
                    contents = ai_master_files + [ai_photo, prompt]
                else:
                    prompt = "写真だけを見て間違っている問題をJSONで返してください。章は単元名、節は詳細項目、番号は1⃣（1）形式にしてください。\n[{\"page\": \"-\", \"chapter\": \"植物\", \"section\": \"光合成\", \"number\": \"1⃣（1）\"}]"
                    contents = [ai_photo, prompt]

                response = client.models.generate_content(model='gemini-3.1-pro-preview', contents=contents)
                match = re.search(r'\[.*\]', response.text, re.DOTALL)
                wrong_problems = json.loads(match.group(0)) if match else []
                save_to_spreadsheet(student_name, subject_name, text_name, wrong_problems, drive_link, creds)
                
                result_text = json.dumps(wrong_problems, ensure_ascii=False, indent=2)
                send_notification_email_plan_b(f"【進捗】写真記録完了 ({photo_name})", f"解析完了:\n{result_text}\n\nリンク: {drive_link}")
            finally:
                try: os.remove(photo_filepath)
                except: pass
        
        send_notification_email_plan_b("【完了】すべての処理が終了しました", f"{student_name} さんの全画像処理が完了しました。")
    except Exception as e:
        send_notification_email_plan_b("【警告】システムエラー", f"エラー内容: {e}")

# ==========================================
# 🌟 Streamlit Web UI
# ==========================================
st.title("📝 採点済みプリント 自動集計システム")

col_left, col_right = st.columns([1, 1])

with col_left:
    st.subheader("👤 講師用アップロード画面")
    student_name = st.selectbox("生徒名", options=STUDENT_LIST, index=None)
    subject_name = st.text_input("科目")
    text_name = st.text_input("テキスト名")
    master_option = st.radio("マスターテキスト", ["💾 保存済みを使う", "🆕 新規アップロード", "❌ 指定しない"])
    
    selected_master_path = None
    if master_option == "💾 保存済みを使う":
        master_files = [f for f in os.listdir(MASTER_DIR) if f.endswith(('.pdf', '.png', '.jpg'))]
        if master_files:
            selected_file = st.selectbox("テキストを選択", master_files)
            selected_master_path = os.path.join(MASTER_DIR, selected_file)
    elif master_option == "🆕 新規アップロード":
        uploaded_master = st.file_uploader("マスターPDF/画像", type=['pdf', 'jpg', 'png'])
        if uploaded_master:
            selected_master_path = os.path.join(MASTER_DIR, uploaded_master.name)
            with open(selected_master_path, "wb") as f: f.write(uploaded_master.getvalue())

    uploaded_photos = st.file_uploader("採点済み写真", type=['jpg', 'jpeg', 'png'], accept_multiple_files=True)

    if st.button("🚀 送信して完了", type="primary"):
        if not student_name or not uploaded_photos: st.error("入力不足です")
        else:
            st.info("📤 アップロードを開始しました。")
            photos_data = []
            for photo in uploaded_photos:
                tmp_filepath = os.path.join(tempfile.gettempdir(), f"photo_{uuid.uuid4().hex}.jpg")
                with open(tmp_filepath, "wb") as f: f.write(photo.getvalue())
                photos_data.append((tmp_filepath, photo.name))
            
            threading.Thread(target=background_processing_task, args=(student_name, subject_name, text_name, selected_master_path, photos_data, GEMINI_API_KEY, GOOGLE_TOKEN_DICT)).start()
            st.success("✅ 受付完了！画面を閉じてOKです。進捗はメールで通知されます。")
            st.balloons()

with col_right:
    st.subheader("📊 現在の集計結果")
    if st.button("🔄 データを更新"): st.rerun()
    df = get_spreadsheet_data(Credentials.from_authorized_user_info(GOOGLE_TOKEN_DICT))
    if not df.empty:
        st.dataframe(df.iloc[::-1], height=600, width='stretch', column_config={"写真リンク": st.column_config.LinkColumn("写真リンク", display_text="リンクを開く")})
