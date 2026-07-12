"""
DMG — Supplier Draft Invoice App (Phase 2)
NCC đăng nhập, nhập draft hóa đơn: chi nhánh giao hàng, ngày, mặt hàng, số lượng, đơn giá,
tải ảnh/scan draft hóa đơn. Dữ liệu ghi vào Google Sheet (tab 'supplier_drafts'), ảnh lưu qua ImgBB.

Dùng chung 1 Google Sheet với branch_receipt_app.py (khác tab), dùng chung Secrets [gcp_service_account] + [app].

Cách chạy local:
    pip install -r requirements.txt
    streamlit run supplier_draft_app.py

Cách cấu hình BỔ SUNG (thêm vào Secrets đã có sẵn từ Phase 1):
    1. Sửa danh sách SUPPLIERS bên dưới thành đúng tên các NCC thật của công ty bạn.
    2. Thêm mục [supplier_accounts] vào Secrets (Streamlit Cloud) — mỗi NCC 1 mật khẩu riêng:

        [supplier_accounts]
        "Tên NCC 1" = "mat_khau_ncc1"
        "Tên NCC 2" = "mat_khau_ncc2"

    3. Sheet ID và ImgBB API key dùng lại đúng như Phase 1, không cần cấu hình thêm.
"""

import streamlit as st
import pandas as pd
from datetime import datetime, date
import uuid
import base64
import requests

import gspread
from google.oauth2.service_account import Credentials

# ============ CẤU HÌNH ============
# ⚠️ SỬA danh sách này thành tên các chi nhánh thật (nên khớp với branch_receipt_app.py)
BRANCHES = [
    "Quận 1: Nguyễn Thị Minh Khai",
    "Tân Phú: Hà Thị Đát",
    "Gò Vấp: Quang Trung",
    "Quận 12: Dương Thị Mười",
    "Thủ Đức (CN1): Lương Khải Siêu",
    "Thủ Đức (CN2): Hiệp Bình Chánh",
    "Quận 9: Nguyễn Văn Tăng",
]

# ⚠️ SỬA danh sách này thành tên các NCC thật đang hợp tác với công ty
SUPPLIERS = [
    "An Phát Thành",
    "CRAFTSMANSHIP",
    "Coca",
    "Cát Tiên",
    "GAS NĂM SAO",
    "HBCO",
    "Hoàng Kim Tín",
    "Hoàng Đào",
    "Huy Hòa",
    "Huỳnh Huy",
    "Hà Thanh",
    "Hưng Thịnh",
    "Hướng Dương",
    "Hạnh Phúc",
    "Khánh Hân",
    "Kim Anh",
    "Kim Ngân",
    "Liên An",
    "Lộc Thảo Phát",
    "Mega Eco",
    "Minh Trí",
    "Minh Tuệ",
    "Minh Tâm",
    "NS Kiên Long",
    "Nam Sơn",
    "Nam Việt Sin",
    "Namilux",
    "Nano",
    "Ngân Thy",
    "Ngọc Oanh",
    "Nhựa Tốt",
    "Paper Store",
    "Phú Thịnh",
    "Phương Dung",
    "Quang Tính",
    "Sao Việt",
    "Song Phương",
    "Sơn Bích",
    "TB Vina",
    "Thành Phát",
    "Thái Thịnh",
    "Thảo Tiên",
    "Tiến Phát",
    "Trung Tiến",
    "Trương Huy",
    "Tuấn Phong",
    "Tuấn Thanh",
    "Tín Nghĩa",
    "Vĩnh Tân",
    "Wood one",
    "Yến Phát",
    "Ánh Dương Xanh",
    "Đệ Nhất",
    "Đồng Xanh",
]

SHEET_TAB_NAME = "supplier_drafts"
SHEET_HEADERS = ["record_id", "ncc_name", "branch_code", "date", "item",
                 "qty_invoiced", "unit_price", "photo_url", "submitted_at"]

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


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


def get_supplier_passwords():
    try:
        return dict(st.secrets["supplier_accounts"])
    except Exception:
        return {s: "1234" for s in SUPPLIERS}


def get_worksheet(gc, sheet_id):
    sh = gc.open_by_key(sheet_id)
    try:
        ws = sh.worksheet(SHEET_TAB_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=SHEET_TAB_NAME, rows=1000, cols=len(SHEET_HEADERS))
        ws.append_row(SHEET_HEADERS)
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


def append_draft_row(ws, row: dict):
    ws.append_row([row.get(h, "") for h in SHEET_HEADERS])


def load_recent_drafts(ws, ncc_name, limit=15):
    records = ws.get_all_records()
    df = pd.DataFrame(records)
    if df.empty:
        return df
    df = df[df["ncc_name"] == ncc_name]
    return df.sort_values("submitted_at", ascending=False).head(limit)


# ============ ĐĂNG NHẬP THEO NCC ============
st.set_page_config(page_title="ĐMG — Nhập Draft Hóa Đơn (NCC)", page_icon="🧾", layout="centered")

if "authenticated_supplier" not in st.session_state:
    st.session_state.authenticated_supplier = None

if st.session_state.authenticated_supplier is None:
    st.title("🔒 Đăng Nhập Nhà Cung Cấp")
    st.caption("Chọn tên công ty của bạn và nhập mật khẩu được cấp để tiếp tục")

    login_supplier = st.selectbox("Nhà cung cấp", SUPPLIERS)
    login_password = st.text_input("Mật khẩu", type="password")
    login_btn = st.button("Đăng nhập", type="primary")

    if login_btn:
        passwords = get_supplier_passwords()
        correct_password = passwords.get(login_supplier)
        if correct_password is not None and login_password == correct_password:
            st.session_state.authenticated_supplier = login_supplier
            st.rerun()
        else:
            st.error("Sai mật khẩu. Vui lòng thử lại hoặc liên hệ P.MH Đắng Mà Ghiền.")

    st.stop()


# ============ GIAO DIỆN ============
authenticated_supplier = st.session_state.authenticated_supplier

top_col1, top_col2 = st.columns([4, 1])
with top_col1:
    st.title("🧾 Nhập Draft Hóa Đơn")
with top_col2:
    if st.button("Đăng xuất"):
        st.session_state.authenticated_supplier = None
        st.rerun()

st.caption(f"Đang đăng nhập: **{authenticated_supplier}** — nhập thông tin draft hóa đơn để P.MH & Kế toán đối chiếu")

with st.form("draft_form", clear_on_submit=True):
    col1, col2 = st.columns(2)
    with col1:
        st.text_input("Nhà cung cấp *", value=authenticated_supplier, disabled=True)
        ncc_name = authenticated_supplier
        branch_code = st.selectbox("Chi nhánh giao hàng *", BRANCHES)
    with col2:
        invoice_date = st.date_input("Ngày giao hàng *", value=date.today())

    st.markdown("---")
    st.markdown("**Danh sách mặt hàng trong draft** — thêm từng dòng")

    if "draft_items" not in st.session_state:
        st.session_state.draft_items = [{"item": "", "qty": 0.0, "price": 0.0}]

    for i, it in enumerate(st.session_state.draft_items):
        c1, c2, c3 = st.columns([3, 1, 1])
        with c1:
            st.session_state.draft_items[i]["item"] = st.text_input(
                f"Mặt hàng #{i+1}", value=it["item"], key=f"ditem_{i}"
            )
        with c2:
            st.session_state.draft_items[i]["qty"] = st.number_input(
                f"Số lượng #{i+1}", value=it["qty"], min_value=0.0, step=0.1, key=f"dqty_{i}"
            )
        with c3:
            st.session_state.draft_items[i]["price"] = st.number_input(
                f"Đơn giá #{i+1}", value=it["price"], min_value=0.0, step=1000.0, key=f"dprice_{i}"
            )

    add_item = st.form_submit_button("➕ Thêm mặt hàng")

    photo = st.file_uploader(
        "Ảnh/scan draft hóa đơn (chụp rõ hoặc chụp màn hình nếu là file PDF) *",
        type=["jpg", "jpeg", "png"],
    )

    submit = st.form_submit_button("✅ Gửi draft hóa đơn", type="primary")

    if add_item:
        st.session_state.draft_items.append({"item": "", "qty": 0.0, "price": 0.0})
        st.rerun()

    if submit:
        errors = []
        if photo is None:
            errors.append("Chưa tải ảnh draft hóa đơn.")
        valid_items = [it for it in st.session_state.draft_items if it["item"].strip() and it["qty"] > 0]
        if not valid_items:
            errors.append("Cần ít nhất 1 mặt hàng có số lượng > 0.")

        if errors:
            for e in errors:
                st.error(e)
        else:
            with st.spinner("Đang lưu draft hóa đơn..."):
                gc = get_google_client()
                sheet_id, imgbb_api_key = get_sheet_id_and_imgbb_key()
                ws = get_worksheet(gc, sheet_id)

                photo_bytes = photo.read()
                filename = f"{ncc_name}_{branch_code}_{invoice_date}_{uuid.uuid4().hex[:6]}.jpg"
                photo_url = upload_photo_to_imgbb(imgbb_api_key, photo_bytes, filename)

                now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                for it in valid_items:
                    row = {
                        "record_id": uuid.uuid4().hex[:10],
                        "ncc_name": ncc_name,
                        "branch_code": branch_code,
                        "date": str(invoice_date),
                        "item": it["item"].strip(),
                        "qty_invoiced": it["qty"],
                        "unit_price": it["price"],
                        "photo_url": photo_url,
                        "submitted_at": now_str,
                    }
                    append_draft_row(ws, row)

            st.success(f"Đã gửi {len(valid_items)} dòng draft hóa đơn cho chi nhánh {branch_code}.")
            st.session_state.draft_items = [{"item": "", "qty": 0.0, "price": 0.0}]

st.markdown("---")
st.subheader(f"📋 Draft đã gửi gần đây — {authenticated_supplier}")

if st.button("🔄 Tải lại danh sách"):
    st.rerun()

try:
    gc = get_google_client()
    sheet_id, imgbb_api_key = get_sheet_id_and_imgbb_key()
    ws = get_worksheet(gc, sheet_id)
    recent = load_recent_drafts(ws, authenticated_supplier)
    if recent.empty:
        st.info("Chưa có draft nào được gửi.")
    else:
        st.dataframe(
            recent[["date", "branch_code", "item", "qty_invoiced", "unit_price", "submitted_at"]],
            use_container_width=True, hide_index=True,
        )
except Exception as e:
    st.warning("Chưa kết nối được Google Sheets — kiểm tra lại cấu hình Secrets. Chi tiết lỗi (cho việc debug): " + str(e))
