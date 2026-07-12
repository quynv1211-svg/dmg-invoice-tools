"""
DMG — Branch Receipt App (Phase 1)
Chi nhánh nhập: NCC, ngày, mặt hàng, số lượng nhận, ảnh giấy giao nhận.
Dữ liệu ghi vào Google Sheet (tab 'branch_receipts'), ảnh lưu Google Drive.

Cách chạy local:
    pip install -r requirements.txt
    streamlit run branch_receipt_app.py

Cách cấu hình (BẮT BUỘC trước khi chạy):
    1. Tạo 1 Google Sheet mới, đặt tên tùy ý, tạo sẵn 1 tab tên "branch_receipts"
       với dòng tiêu đề: record_id | branch_code | ncc_name | date | item | qty_received | photo_url | submitted_by | submitted_at
    2. Tạo 1 Google Cloud Service Account, bật Google Sheets API + Google Drive API,
       tải file credentials JSON, share quyền Editor Sheet trên cho email service account đó.
    3. Deploy lên Streamlit Cloud: dán nội dung JSON vào mục Settings > Secrets theo định dạng TOML bên dưới:

        [gcp_service_account]
        type = "service_account"
        project_id = "..."
        private_key_id = "..."
        private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
        client_email = "...@...iam.gserviceaccount.com"
        client_id = "..."
        token_uri = "https://oauth2.googleapis.com/token"

        [app]
        sheet_id = "DÁN_ID_GOOGLE_SHEET_CỦA_BẠN_VÀO_ĐÂY"
        drive_folder_id = "DÁN_ID_FOLDER_GOOGLE_DRIVE_ĐỂ_LƯU_ẢNH"

    Khi chạy local (không dùng Streamlit Cloud), có thể thay st.secrets bằng đọc file
    service_account.json trực tiếp — xem phần "LOCAL DEV MODE" bên dưới.
"""

import streamlit as st
import pandas as pd
from datetime import datetime, date
import uuid
import io

import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# ============ CẤU HÌNH ============
BRANCHES = ["CN1", "CN2 - Gò Vấp", "CN3 - Q12", "CN4", "CN5", "CN6", "CN7"]
SHEET_TAB_NAME = "branch_receipts"
SHEET_HEADERS = ["record_id", "branch_code", "ncc_name", "date", "item",
                 "qty_received", "photo_url", "submitted_by", "submitted_at"]

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
]


@st.cache_resource
def get_google_clients():
    """Khởi tạo kết nối Google Sheets + Drive từ Streamlit Secrets."""
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    except Exception:
        # ---- LOCAL DEV MODE: thay bằng đường dẫn file JSON tải từ Google Cloud ----
        creds = Credentials.from_service_account_file("service_account.json", scopes=SCOPES)

    gc = gspread.authorize(creds)
    drive_service = build("drive", "v3", credentials=creds)
    return gc, drive_service


def get_sheet_id_and_folder():
    try:
        return st.secrets["app"]["sheet_id"], st.secrets["app"]["drive_folder_id"]
    except Exception:
        # ---- LOCAL DEV MODE: điền trực tiếp 2 ID này khi test trên máy ----
        return "DÁN_SHEET_ID_KHI_TEST_LOCAL", "DÁN_DRIVE_FOLDER_ID_KHI_TEST_LOCAL"


def get_worksheet(gc, sheet_id):
    sh = gc.open_by_key(sheet_id)
    try:
        ws = sh.worksheet(SHEET_TAB_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=SHEET_TAB_NAME, rows=1000, cols=len(SHEET_HEADERS))
        ws.append_row(SHEET_HEADERS)
    return ws


def upload_photo_to_drive(drive_service, folder_id, file_bytes, filename, mime_type):
    file_metadata = {"name": filename, "parents": [folder_id]}
    media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mime_type, resumable=False)
    uploaded = drive_service.files().create(
        body=file_metadata, media_body=media, fields="id"
    ).execute()
    file_id = uploaded["id"]
    # Cho phép xem qua link (không public toàn bộ Drive, chỉ file này)
    drive_service.permissions().create(
        fileId=file_id, body={"role": "reader", "type": "anyone"}
    ).execute()
    return f"https://drive.google.com/file/d/{file_id}/view"


def append_receipt_row(ws, row: dict):
    ws.append_row([row.get(h, "") for h in SHEET_HEADERS])


def load_recent_receipts(ws, branch_code, limit=15):
    records = ws.get_all_records()
    df = pd.DataFrame(records)
    if df.empty:
        return df
    df = df[df["branch_code"] == branch_code]
    return df.sort_values("submitted_at", ascending=False).head(limit)


# ============ GIAO DIỆN ============
st.set_page_config(page_title="ĐMG — Nhập Chứng Từ Nhận Hàng", page_icon="📦", layout="centered")

st.title("📦 Nhập Chứng Từ Nhận Hàng")
st.caption("Dành cho nhân viên chi nhánh — nhập số lượng thực nhận và tải ảnh giấy giao nhận từ NCC")

with st.form("receipt_form", clear_on_submit=True):
    col1, col2 = st.columns(2)
    with col1:
        branch_code = st.selectbox("Chi nhánh *", BRANCHES)
        receipt_date = st.date_input("Ngày nhận hàng *", value=date.today())
    with col2:
        ncc_name = st.text_input("Tên nhà cung cấp (NCC) *", placeholder="VD: Fresh Đại Phát")
        submitted_by = st.text_input("Tên người nhập *", placeholder="Tên của bạn")

    st.markdown("---")
    st.markdown("**Danh sách mặt hàng nhận được** — thêm từng dòng")

    if "receipt_items" not in st.session_state:
        st.session_state.receipt_items = [{"item": "", "qty": 0.0}]

    for i, it in enumerate(st.session_state.receipt_items):
        c1, c2 = st.columns([3, 1])
        with c1:
            st.session_state.receipt_items[i]["item"] = st.text_input(
                f"Mặt hàng #{i+1}", value=it["item"], key=f"item_{i}"
            )
        with c2:
            st.session_state.receipt_items[i]["qty"] = st.number_input(
                f"Số lượng #{i+1}", value=it["qty"], min_value=0.0, step=0.1, key=f"qty_{i}"
            )

    add_col, _ = st.columns([1, 3])
    add_item = st.form_submit_button("➕ Thêm mặt hàng")

    photo = st.file_uploader("Ảnh giấy giao nhận (chụp rõ, có chữ ký nếu có) *", type=["jpg", "jpeg", "png"])

    submit = st.form_submit_button("✅ Gửi chứng từ", type="primary")

    if add_item:
        st.session_state.receipt_items.append({"item": "", "qty": 0.0})
        st.rerun()

    if submit:
        errors = []
        if not ncc_name.strip():
            errors.append("Chưa nhập tên NCC.")
        if not submitted_by.strip():
            errors.append("Chưa nhập tên người nhập.")
        if photo is None:
            errors.append("Chưa tải ảnh giấy giao nhận.")
        valid_items = [it for it in st.session_state.receipt_items if it["item"].strip() and it["qty"] > 0]
        if not valid_items:
            errors.append("Cần ít nhất 1 mặt hàng có số lượng > 0.")

        if errors:
            for e in errors:
                st.error(e)
        else:
            with st.spinner("Đang lưu chứng từ..."):
                gc, drive_service = get_google_clients()
                sheet_id, folder_id = get_sheet_id_and_folder()
                ws = get_worksheet(gc, sheet_id)

                photo_bytes = photo.read()
                filename = f"{branch_code}_{ncc_name}_{receipt_date}_{uuid.uuid4().hex[:6]}.jpg"
                photo_url = upload_photo_to_drive(drive_service, folder_id, photo_bytes, filename, photo.type)

                now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                for it in valid_items:
                    row = {
                        "record_id": uuid.uuid4().hex[:10],
                        "branch_code": branch_code,
                        "ncc_name": ncc_name.strip(),
                        "date": str(receipt_date),
                        "item": it["item"].strip(),
                        "qty_received": it["qty"],
                        "photo_url": photo_url,
                        "submitted_by": submitted_by.strip(),
                        "submitted_at": now_str,
                    }
                    append_receipt_row(ws, row)

            st.success(f"Đã gửi {len(valid_items)} dòng chứng từ cho {ncc_name} — chi nhánh {branch_code}.")
            st.session_state.receipt_items = [{"item": "", "qty": 0.0}]

st.markdown("---")
st.subheader("📋 Chứng từ đã gửi gần đây")

view_branch = st.selectbox("Xem theo chi nhánh", BRANCHES, key="view_branch")
if st.button("🔄 Tải lại danh sách"):
    st.rerun()

try:
    gc, drive_service = get_google_clients()
    sheet_id, folder_id = get_sheet_id_and_folder()
    ws = get_worksheet(gc, sheet_id)
    recent = load_recent_receipts(ws, view_branch)
    if recent.empty:
        st.info("Chưa có chứng từ nào cho chi nhánh này.")
    else:
        st.dataframe(
            recent[["date", "ncc_name", "item", "qty_received", "submitted_by", "submitted_at"]],
            use_container_width=True, hide_index=True,
        )
except Exception as e:
    st.warning("Chưa kết nối được Google Sheets — kiểm tra lại cấu hình Secrets. Chi tiết lỗi (cho việc debug): " + str(e))
