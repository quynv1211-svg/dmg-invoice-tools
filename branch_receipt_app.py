"""
DMG — Branch Receipt App (Phase 1 — v2)
Chi nhánh nhập: Nhóm hàng, NCC (lọc theo nhóm hàng), mặt hàng (lọc theo NCC, tự điền tên + ĐVT),
số lượng nhận, ảnh giấy giao nhận (không bắt buộc).
Dữ liệu ghi vào Google Sheet (tab 'branch_receipts'), ảnh lưu qua ImgBB (dịch vụ ảnh miễn phí).

Cách chạy local:
    pip install -r requirements.txt
    streamlit run branch_receipt_app.py

Cách cấu hình (BẮT BUỘC trước khi chạy):
    1. Tạo 1 Google Sheet mới, tạo sẵn 1 tab tên "branch_receipts" (app sẽ tự tạo cột nếu chưa có).
    2. Tạo Google Cloud Service Account, bật Google Sheets API, tải file JSON, share Sheet cho email đó (Editor).
    3. Đăng ký API Key miễn phí tại https://api.imgbb.com/
    4. File "master_ncc.json" (danh mục Nhóm hàng/NCC/Mã hàng) PHẢI nằm CÙNG THƯ MỤC với file này khi deploy.
    5. Deploy lên Streamlit Cloud, khai báo Secrets:

        [gcp_service_account]
        type = "service_account"
        project_id = "..."
        private_key_id = "..."
        private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
        client_email = "...@...iam.gserviceaccount.com"
        client_id = "..."
        token_uri = "https://oauth2.googleapis.com/token"

        [app]
        sheet_id = "..."
        imgbb_api_key = "..."

        [branch_accounts]
        "Tên chi nhánh 1" = "mat_khau_1"
        ...
"""

import streamlit as st
import pandas as pd
from datetime import datetime, date
from zoneinfo import ZoneInfo

VN_TZ = ZoneInfo("Asia/Ho_Chi_Minh")
import uuid
import base64
import json
import os

import requests
import gspread
from google.oauth2.service_account import Credentials

# ============ CẤU HÌNH ============
BRANCHES = [
    "Quận 1: Nguyễn Thị Minh Khai",
    "Tân Phú: Hà Thị Đát",
    "Gò Vấp: Quang Trung",
    "Quận 12: Dương Thị Mười",
    "Thủ Đức (CN1): Lương Khải Siêu",
    "Thủ Đức (CN2): Hiệp Bình Chánh",
    "Quận 9: Nguyễn Văn Tăng",
]

SHEET_TAB_NAME = "branch_receipts"
SHEET_HEADERS = ["record_id", "branch_code", "category", "ncc_name", "date",
                 "item_code", "item_name", "unit", "qty_received",
                 "photo_url", "submitted_by", "submitted_at"]

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

MASTER_DATA_PATH = os.path.join(os.path.dirname(__file__), "master_ncc.json")


@st.cache_data
def load_master_data():
    with open(MASTER_DATA_PATH, encoding="utf-8") as f:
        return json.load(f)


@st.cache_resource
def get_google_client():
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    except Exception:
        creds = Credentials.from_service_account_file("service_account.json", scopes=SCOPES)
    gc = gspread.authorize(creds)
    return gc


def get_sheet_id_and_imgbb_key():
    try:
        return st.secrets["app"]["sheet_id"], st.secrets["app"]["imgbb_api_key"]
    except Exception:
        return "DÁN_SHEET_ID_KHI_TEST_LOCAL", "DÁN_IMGBB_API_KEY_KHI_TEST_LOCAL"


def get_branch_passwords():
    try:
        return dict(st.secrets["branch_accounts"])
    except Exception:
        return {b: "1234" for b in BRANCHES}


def get_worksheet(gc, sheet_id):
    sh = gc.open_by_key(sheet_id)
    try:
        ws = sh.worksheet(SHEET_TAB_NAME)
        if not st.session_state.get("_header_checked", False):
            existing_header = ws.row_values(1)
            if existing_header != SHEET_HEADERS:
                ws.update(values=[SHEET_HEADERS], range_name="A1")
            st.session_state["_header_checked"] = True
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=SHEET_TAB_NAME, rows=1000, cols=len(SHEET_HEADERS))
        ws.append_row(SHEET_HEADERS)
        st.session_state["_header_checked"] = True
    return ws


def upload_photo_to_imgbb(api_key, file_bytes, filename):
    b64_image = base64.b64encode(file_bytes).decode("utf-8")
    resp = requests.post(
        "https://api.imgbb.com/1/upload",
        data={"key": api_key, "image": b64_image, "name": filename},
        timeout=30,
    )
    resp.raise_for_status()
    result = resp.json()
    if not result.get("success"):
        raise RuntimeError(f"ImgBB upload thất bại: {result}")
    return result["data"]["url"]


def append_receipt_row(ws, row: dict):
    ws.append_row([row.get(h, "") for h in SHEET_HEADERS], table_range="A1")


def load_recent_receipts(ws, branch_code, limit=15):
    all_values = ws.get_all_values()
    if len(all_values) <= 1:
        return pd.DataFrame(columns=SHEET_HEADERS)
    data_rows = all_values[1:]  # bỏ dòng tiêu đề thật trên Sheet, tự dùng SHEET_HEADERS của code
    # Cắt/đệm mỗi dòng cho khớp đúng số cột mong đợi, tránh lỗi lệch cột dữ liệu cũ
    normalized_rows = []
    for r in data_rows:
        if len(r) < len(SHEET_HEADERS):
            r = r + [""] * (len(SHEET_HEADERS) - len(r))
        normalized_rows.append(r[:len(SHEET_HEADERS)])
    df = pd.DataFrame(normalized_rows, columns=SHEET_HEADERS)
    df = df[df["branch_code"] == branch_code]
    if df.empty:
        return df
    return df.sort_values("submitted_at", ascending=False).head(limit)


# ============ ĐĂNG NHẬP THEO CHI NHÁNH ============
st.set_page_config(page_title="ĐMG — Nhập Chứng Từ Nhận Hàng", page_icon="📦", layout="centered")

if "authenticated_branch" not in st.session_state:
    st.session_state.authenticated_branch = None

if st.session_state.authenticated_branch is None:
    st.title("🔒 Đăng Nhập Chi Nhánh")
    st.caption("Chọn chi nhánh của bạn và nhập mật khẩu được cấp để tiếp tục")

    login_branch = st.selectbox("Chi nhánh", BRANCHES)
    login_password = st.text_input("Mật khẩu chi nhánh", type="password")
    login_btn = st.button("Đăng nhập", type="primary")

    if login_btn:
        passwords = get_branch_passwords()
        correct_password = passwords.get(login_branch)
        if correct_password is not None and login_password == correct_password:
            st.session_state.authenticated_branch = login_branch
            st.rerun()
        else:
            st.error("Sai mật khẩu cho chi nhánh này. Vui lòng thử lại hoặc liên hệ P.MH.")

    st.stop()


# ============ GIAO DIỆN CHÍNH ============
authenticated_branch = st.session_state.authenticated_branch
master = load_master_data()

top_col1, top_col2 = st.columns([4, 1])
with top_col1:
    st.title("📦 Nhập Chứng Từ Nhận Hàng")
with top_col2:
    if st.button("Đăng xuất"):
        st.session_state.authenticated_branch = None
        st.rerun()

st.caption(f"Đang đăng nhập: **{authenticated_branch}** — nhập số lượng thực nhận, ảnh giấy giao nhận không bắt buộc")

if st.session_state.get("last_submit_message"):
    st.success(st.session_state["last_submit_message"])
    del st.session_state["last_submit_message"]

# --- Không dùng st.form ở đây: cần rerun ngay khi đổi Nhóm hàng / NCC để lọc dropdown tiếp theo ---
col1, col2 = st.columns(2)
with col1:
    st.text_input("Chi nhánh *", value=authenticated_branch, disabled=True)
    receipt_date = st.date_input("Ngày nhận hàng *", value=datetime.now(VN_TZ).date())
with col2:
    submitted_by = st.text_input("Tên người nhập *", placeholder="Tên của bạn", key="submitted_by_input")

st.markdown("---")
cat_col, ncc_col = st.columns(2)
with cat_col:
    category = st.selectbox("Nhóm hàng *", master["categories"], key="category_select")
with ncc_col:
    ncc_options = master["cat_to_ncc"].get(category, [])
    ncc_name = st.selectbox("Nhà cung cấp (NCC) *", ncc_options, key="ncc_select")

items_key = f"{category}||{ncc_name}"
available_items = master["catncc_to_items"].get(items_key, [])
item_code_options = [it["code"] for it in available_items]
item_lookup = {it["code"]: it for it in available_items}

st.markdown("---")
st.markdown("**Danh sách mặt hàng nhận được** — thêm từng dòng")

if "receipt_items" not in st.session_state:
    st.session_state.receipt_items = [{"code": None, "qty": 0.0}]

if not item_code_options:
    st.warning("Không tìm thấy mã hàng nào cho NCC này trong danh mục — vui lòng báo P.MH bổ sung vào Master_NCC.")

for i, it in enumerate(st.session_state.receipt_items):
    c1, c2, c3 = st.columns([2, 2, 1])
    selected_code = None
    with c1:
        current_code = it["code"] if it["code"] in item_code_options else (item_code_options[0] if item_code_options else None)
        if item_code_options:
            selected_code = st.selectbox(
                f"Mã hàng #{i+1}", item_code_options,
                index=item_code_options.index(current_code) if current_code in item_code_options else 0,
                key=f"item_code_{i}",
            )
        st.session_state.receipt_items[i]["code"] = selected_code
    with c2:
        if selected_code:
            info = item_lookup.get(selected_code, {})
            st.markdown(f"**Tên SP #{i+1}**")
            st.info(f"{info.get('name','(chưa rõ tên)')} — ĐVT: {info.get('unit','?')}")
    with c3:
        st.session_state.receipt_items[i]["qty"] = st.number_input(
            f"Số lượng #{i+1}", value=it["qty"], min_value=0.0, step=0.1, key=f"item_qty_{i}"
        )

btn_col1, btn_col2 = st.columns([1, 1])
with btn_col1:
    if st.button("➕ Thêm mặt hàng"):
        st.session_state.receipt_items.append({"code": None, "qty": 0.0})
        st.rerun()
with btn_col2:
    if len(st.session_state.receipt_items) > 1 and st.button("➖ Xóa dòng cuối"):
        st.session_state.receipt_items.pop()
        st.rerun()

st.markdown("---")
photo = st.file_uploader(
    "Ảnh giấy giao nhận (không bắt buộc — có thể bổ sung sau)",
    type=["jpg", "jpeg", "png"],
)

if st.button("✅ Gửi chứng từ", type="primary"):
    print("[DEBUG] Đã bấm nút Gửi chứng từ", flush=True)
    errors = []
    if not submitted_by.strip():
        errors.append("Chưa nhập tên người nhập.")
    valid_items = [it for it in st.session_state.receipt_items if it["code"] and it["qty"] > 0]
    if not valid_items:
        errors.append("Cần ít nhất 1 mặt hàng có chọn mã hàng và số lượng > 0.")

    print(f"[DEBUG] errors={errors}, valid_items={valid_items}", flush=True)

    if errors:
        for e in errors:
            st.error(e)
    else:
        try:
            print("[DEBUG] Bắt đầu try block", flush=True)
            with st.spinner("Đang lưu chứng từ..."):
                print("[DEBUG] Gọi get_google_client()", flush=True)
                gc = get_google_client()
                print("[DEBUG] get_google_client() xong", flush=True)
                sheet_id, imgbb_api_key = get_sheet_id_and_imgbb_key()
                print(f"[DEBUG] sheet_id={sheet_id[:8]}..., có imgbb_key={bool(imgbb_api_key)}", flush=True)
                ws = get_worksheet(gc, sheet_id)
                print("[DEBUG] get_worksheet() xong", flush=True)

                photo_url = ""
                if photo is not None:
                    print("[DEBUG] Bắt đầu upload ảnh lên ImgBB", flush=True)
                    photo_bytes = photo.read()
                    filename = f"{authenticated_branch}_{ncc_name}_{receipt_date}_{uuid.uuid4().hex[:6]}.jpg"
                    photo_url = upload_photo_to_imgbb(imgbb_api_key, photo_bytes, filename)
                    print("[DEBUG] Upload ảnh xong:", photo_url, flush=True)
                else:
                    print("[DEBUG] Không có ảnh, bỏ qua upload", flush=True)

                now_str = datetime.now(VN_TZ).strftime("%Y-%m-%d %H:%M:%S")
                for it in valid_items:
                    info = item_lookup.get(it["code"], {})
                    row = {
                        "record_id": uuid.uuid4().hex[:10],
                        "branch_code": authenticated_branch,
                        "category": category,
                        "ncc_name": ncc_name,
                        "date": str(receipt_date),
                        "item_code": it["code"],
                        "item_name": info.get("name", ""),
                        "unit": info.get("unit", ""),
                        "qty_received": it["qty"],
                        "photo_url": photo_url,
                        "submitted_by": submitted_by.strip(),
                        "submitted_at": now_str,
                    }
                    print("[DEBUG] Chuẩn bị append_receipt_row:", row, flush=True)
                    append_receipt_row(ws, row)
                    print("[DEBUG] append_receipt_row xong", flush=True)

            print("[DEBUG] Toàn bộ khối try thành công, chuẩn bị hiện success", flush=True)
            st.session_state["last_submit_message"] = f"Đã gửi {len(valid_items)} dòng chứng từ cho {ncc_name} — chi nhánh {authenticated_branch}."
            st.session_state.receipt_items = [{"code": None, "qty": 0.0}]
            st.rerun()
        except Exception as e:
            print(f"[DEBUG] LỖI trong try block: {repr(e)}", flush=True)
            st.error(f"⚠️ Có lỗi xảy ra khi lưu chứng từ, vui lòng thử lại hoặc báo P.MH. Chi tiết lỗi (cho việc debug): {e}")

st.markdown("---")
st.subheader(f"📋 Chứng từ đã gửi gần đây — {authenticated_branch}")

if st.button("🔄 Tải lại danh sách"):
    st.rerun()

try:
    gc = get_google_client()
    sheet_id, imgbb_api_key = get_sheet_id_and_imgbb_key()
    ws = get_worksheet(gc, sheet_id)
    recent = load_recent_receipts(ws, authenticated_branch)
    if recent.empty:
        st.info("Chưa có chứng từ nào cho chi nhánh này.")
    else:
        st.dataframe(
            recent[["date", "category", "ncc_name", "item_code", "item_name", "unit", "qty_received", "submitted_by", "submitted_at"]],
            use_container_width=True, hide_index=True,
        )
except Exception as e:
    st.warning("Chưa kết nối được Google Sheets — kiểm tra lại cấu hình Secrets. Chi tiết lỗi (cho việc debug): " + str(e))
