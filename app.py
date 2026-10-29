################################################################################
# [UNKNOWN BIKE OS - v6 Enterprise  /  무결성 검증 완료 + UI 전면 개편 + 공급업체 CRM]
################################################################################
import streamlit as st
import pandas as pd
import gspread
from gspread.utils import rowcol_to_a1
from datetime import datetime, timedelta, timezone
import re
import os
import json
import html
import hmac
import base64
import time
import traceback
from io import BytesIO
from PIL import Image, ImageDraw, ImageFont
import streamlit.components.v1 as components
from google import genai
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

try:
    import printers
except ImportError:
    pass

st.set_page_config(page_title="UNKNOWN BIKE OS", layout="wide")

# 💡 [UI/UX 개편] Material 아이콘(_arrow_right) 깨짐 원천 차단 + 밝고 선명한 엔터프라이즈 테마
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;600;700;800&display=swap');
    
    /* 아이콘 폰트(Material Symbols)를 침범하지 않도록 텍스트 태그에만 정확히 폰트 지정 */
    html, body, p, label, input, textarea, button, th, td, li, h1, h2, h3, h4, h5, h6 {
        font-family: 'Noto Sans KR', -apple-system, BlinkMacSystemFont, sans-serif;
    }

    /* 사이드바: 가독성이 뛰어난 밝은 오피스 그레이 배경 */
    [data-testid="stSidebar"] {
        background-color: #F4F6F9 !important;
        border-right: 1px solid #E2E8F0 !important;
    }
    .menu-title {
        font-size: 13px;
        font-weight: 800;
        color: #64748B;
        margin: 15px 0px 12px 6px;
        letter-spacing: 1.2px;
    }

    /* 사이드바 메뉴 라디오 버튼을 큼직하고 선명한 탭 카드 형태로 디자인 */
    div[data-testid="stSidebar"] div[data-testid="stRadio"] > div[role="radiogroup"] > label {
        background-color: #FFFFFF !important;
        border: 1px solid #CBD5E1 !important;
        padding: 15px 18px !important;
        border-radius: 10px !important;
        margin-bottom: 10px !important;
        cursor: pointer !important;
        transition: all 0.15s ease-in-out !important;
        box-shadow: 0 1px 2px rgba(0,0,0,0.03) !important;
    }
    div[data-testid="stSidebar"] div[data-testid="stRadio"] > div[role="radiogroup"] > label:hover {
        border-color: #94A3B8 !important;
        background-color: #F8FAFC !important;
    }
    /* 사이드바 라디오 기본 동그라미(점) 숨기기 */
    div[data-testid="stSidebar"] div[data-testid="stRadio"] > div[role="radiogroup"] > label > div:first-child {
        display: none !important;
    }
    /* 사이드바 메뉴 글씨: 진한 네이비 블랙, 큼직한 사이즈 */
    div[data-testid="stSidebar"] div[data-testid="stRadio"] > div[role="radiogroup"] > label p {
        font-size: 17px !important;
        font-weight: 700 !important;
        color: #1E293B !important;
        margin: 0 !important;
    }
    /* 선택된 메뉴: 선명한 레드 포인트 배경 + 흰색 글씨 */
    div[data-testid="stSidebar"] div[data-testid="stRadio"] > div[role="radiogroup"] > label[data-checked="true"],
    div[data-testid="stSidebar"] div[data-testid="stRadio"] > div[role="radiogroup"] > label:has(input:checked) {
        background-color: #CC0000 !important;
        border-color: #B30000 !important;
        box-shadow: 0 4px 10px rgba(204, 0, 0, 0.22) !important;
    }
    div[data-testid="stSidebar"] div[data-testid="stRadio"] > div[role="radiogroup"] > label[data-checked="true"] p,
    div[data-testid="stSidebar"] div[data-testid="stRadio"] > div[role="radiogroup"] > label:has(input:checked) p {
        color: #FFFFFF !important;
        font-weight: 800 !important;
    }

    /* 상단 요약 카드(metric-card): 눈의 피로도를 낮추는 화이트 카드 + 고정폭 숫자 정렬 */
    .metric-card {
        background-color: #FFFFFF;
        padding: 20px 22px;
        border-radius: 12px;
        border: 1px solid #E2E8F0;
        border-left: 5px solid #CC0000;
        margin-bottom: 18px;
        box-shadow: 0 2px 6px rgba(15, 23, 42, 0.04);
    }
    .metric-title {
        font-size: 13px;
        font-weight: 600;
        color: #64748B;
        margin-bottom: 6px;
    }
    .metric-value {
        font-size: 24px;
        font-weight: 800;
        color: #0F172A;
        font-variant-numeric: tabular-nums;
    }

    @media print {
        header, [data-testid="stSidebar"], .stTabs [data-baseweb="tab-list"],
        button, input, select, [data-testid="stToolbar"], .metric-card { display: none !important; }
        body { background-color: white !important; padding: 0 !important; margin: 0 !important; }
        img { width: 100% !important; height: 95vh !important; object-fit: contain !important; page-break-after: always !important; }
        @page { size: landscape; margin: 0; }
    }
    </style>
    """, unsafe_allow_html=True)

################################################################################
# 1. 설정 / 상수
################################################################################
KST = timezone(timedelta(hours=9))

def now_kst():
    return datetime.now(KST)

def sec(name, default="", legacy=False):
    try: v = st.secrets.get(name, None)
    except Exception: v = None
    if v is None or str(v) == "":
        return default
    return str(v)

TRACKING_SHEET_ID = sec("SHIPMENT_SHEET_ID", "1S8IuINkkN6lCNEB0ruCesUPqX97KuIfKwPMpIj6Ofdo", True)
CUST_SHEET_ID = sec("PARTNER_SHEET_ID", "1lGEQyNgFI2x-GuSadmPK9DAqX5zRbmah3HSFa9PaHn0", True)
GY_SHEET_ID = sec("KEUNYOUNG_SHEET_ID", "1fsopleTjYdceyK4TpoWR-mOCds1bqjJiv4Qc_u4eshE", True)
WECHAT_SHEET_ID = sec("WECHAT_SHEET_ID", "1iFMujcOcEWAN7LaW2B2eo7-pZ9fY4cOfzyyQM1TC6SA", True)
DRIVE_FOLDER_ID = sec("DRIVE_FOLDER_ID", "1KWabdTHCjGseHquHJVV89S53xI8bqbgN", True)
WECHAT_PASSWORD = sec("WECHAT_PASSWORD", "unk", True)
APP_PASSWORD = sec("APP_PASSWORD", "")
GEMINI_MODEL = sec("GEMINI_MODEL", "gemini-3.8-flash")
MODEL_CANDIDATES = [GEMINI_MODEL, "gemini-3.8-flash", "gemini-3.5-flash-lite", "gemini-3.5-flash"]

# 💡 [요금 정보 명칭 반영: 기본->내용물, 휠/프레임(선택), 매장완차->완차, 개인완차->완성차]
RATE_ITEMS = [
    {"key": "q_base", "label": "기본", "memo": "내용물", "rate": 5500},
    {"key": "q_wheel", "label": "휠/프레임", "memo": "휠", "rate": 6000},
    {"key": "q_shop", "label": "매장완차", "memo": "완차", "rate": 11000},
    {"key": "q_personal", "label": "개인완차", "memo": "완성차", "rate": 35000},
]

TRACK_HEADERS = ['날짜', '트레킹번호', '받는사람', '택배사', '박스수량', '내용물', '메모']
GY_HEADERS = ["날짜", "받는 곳", "연락처", "주소", "운송장번호", "수량", "운임", "비고"]
CUST_HEADERS = ['종류', '이름', '전화번호', '이메일', '주소', '메모']
WECHAT_HEADERS = ["날짜", "공장", "태그", "요약제목", "상세내용"]
TASK_HEADERS = ["생성일", "공장", "구분", "담당", "내용", "기한", "상태", "캘린더ID"]
PRICE_HEADERS = ["날짜", "공장", "품목", "내용"]
SUPPLIER_HEADERS = ["업체명", "담당자명", "메모"]

current_dir = os.path.dirname(os.path.abspath(__file__))
LOGO_FILE_PNG = os.path.join(current_dir, 'unknown2026.png')
FONT_BOLD = os.path.join(current_dir, "NotoSansKR-Bold.ttf")
FONT_REG = os.path.join(current_dir, "NotoSansKR-Regular.ttf")

TAG_MAP = {"긴급": "🔴[긴급이슈]", "생산납기": "🔵[생산/납기]", "단가스펙": "🟡[단가/스펙]", "일반": "🟢[일반진행]"}
TAG_OPTIONS = list(TAG_MAP.values())
RAW_CELL_LIMIT = 30000     
CELL_TOTAL_LIMIT = 48000   

################################################################################
# 2. 앱 전체 로그인 
################################################################################
if APP_PASSWORD and not st.session_state.get('app_auth', False):
    st.markdown("<h2 style='text-align:center; color:#cc0000; margin-top:80px;'>🔒 UNKNOWN BIKE OS</h2>", unsafe_allow_html=True)
    _l, _m, _r = st.columns([1, 1, 1])
    with _m:
        _pw = st.text_input("비밀번호", type="password", key="app_pw_input")
        if st.button("로그인", use_container_width=True, type="primary"):
            if hmac.compare_digest(str(_pw), APP_PASSWORD):
                st.session_state['app_auth'] = True
                st.rerun()
            else:
                st.error("❌ 비밀번호가 일치하지 않습니다.")
    st.stop()

################################################################################
# 3. 구글 연결 (캘린더 API 제거로 속도 향상 및 에러 방지)
################################################################################
@st.cache_resource(ttl=3000)
def init_master_connection_v3():
    creds_dict = dict(st.secrets["gcp_service_account"])
    if "\\n" in creds_dict["private_key"]:
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    credentials = service_account.Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(credentials), build('drive', 'v3', credentials=credentials)

try:
    client, drive_service = init_master_connection_v3()
except Exception as _e:
    client, drive_service = None, None
    _conn_err = str(_e)

if not client:
    st.error("🚨 서버 인증 실패: Secrets 환경설정의 암호를 확인해주세요.")
    st.stop()

try:
    gemini_key = st.secrets.get("GEMINI_API_KEY", "")
except Exception: pass

# --- 세션 상태 초기화 ---
if 'deli_list' not in st.session_state: st.session_state['deli_list'] = []
if 'bulk_input_data' not in st.session_state: st.session_state['bulk_input_data'] = [["", "", "한진", "1", "", ""] for _ in range(20)]
if 'wechat_auth' not in st.session_state: st.session_state['wechat_auth'] = False
if 'sd_date' not in st.session_state: st.session_state['sd_date'] = now_kst().date()
if 'wf_type' not in st.session_state: st.session_state['wf_type'] = "휠"

for _k, _v in [('bulk_ver', 0), ('trk_ver', 0), ('print_imgs', None), ('wc_draft', None), ('wc_ver', 0),
               ('draft_checked', False), ('draft_found', None), ('wc_limit', 20), ('wc_sig', ''), ('sup_ver', 0)]:
    if _k not in st.session_state: st.session_state[_k] = _v

form_keys = ['in_name', 'in_phone', 'in_addr', 'in_inv', 'q_base', 'q_wheel', 'q_shop', 'q_personal', 'in_qty', 'in_fee', 'in_memo', 'cust_select', 'in_repeat']
for k in form_keys:
    if k not in st.session_state:
        if k in ['q_base', 'q_wheel', 'q_shop', 'q_personal', 'in_fee']: st.session_state[k] = 0
        elif k in ('in_qty', 'in_repeat'): st.session_state[k] = 1
        elif k == 'cust_select': st.session_state[k] = "직접 입력"
        else: st.session_state[k] = ""

################################################################################
# 4. 공통 유틸
################################################################################
def s(v):
    if v is None: return ""
    try:
        if pd.isna(v): return ""
    except (TypeError, ValueError): pass
    return str(v).strip()

def safe_int(v, default=0):
    try:
        if v is None: return default
        if isinstance(v, float) and v != v: return default
        t = str(v).replace(',', '').strip()
        if t == "" or t.lower() in ("nan", "none"): return default
        return int(float(t))
    except (ValueError, TypeError): return default

def valid_date(t):
    t = s(t)
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", t): return False
    try:
        datetime.strptime(t, "%Y-%m-%d")
        return True
    except ValueError: return False

def norm_date(v):
    t = s(v)
    if not t: return ""
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", t): return t
    try:
        d = pd.to_datetime(t.replace("마감", "").strip(), errors="coerce")
        if pd.isna(d): return t
        return d.strftime("%Y-%m-%d")
    except Exception: return t

def to_native(v):
    if isinstance(v, float) and v != v: return ""
    if hasattr(v, "item") and not isinstance(v, (str, bytes)):
        try: return v.item()
        except Exception: pass
    return v

def clean_rows(rows): return [[to_native(c) for c in r] for r in rows]

def rows_to_df(data, headers):
    if not data: return pd.DataFrame(columns=headers)
    hdr = data[0]
    df = pd.DataFrame(data[1:], columns=hdr).fillna('') if len(data) > 1 else pd.DataFrame(columns=hdr)
    for h in headers:
        if h not in df.columns: df[h] = ''
    return df[headers].copy()

def fix_trk(v):
    t = s(v)
    if t.endswith('.0'): t = t[:-2]
    if len(t) == 9 and t.startswith('19'): t = '0' + t
    return t

def inc_inv(inv, n=1):
    m = re.search(r'(\d+)$', inv)
    if not m: return inv
    num = m.group(1)
    return inv[:m.start()] + str(int(num) + n).zfill(len(num))

def flash(kind, msg): st.session_state.setdefault('_flash', []).append((kind, msg))
def show_flash():
    for kind, msg in st.session_state.pop('_flash', []):
        getattr(st, kind, st.info)(msg)

def get_hex_color(courier): return {"한진": "#FFFFFF", "건영": "#FCE4D6", "천일": "#FFF2CC", "대신": "#DDEBF7", "우리": "#E2EFDA"}.get(courier, "#F2F2F2")
def get_rgb_dict(courier): return {"한진": {"red": 1.0, "green": 1.0, "blue": 1.0}, "건영": {"red": 0.988, "green": 0.894, "blue": 0.839}, "천일": {"red": 1.0, "green": 0.949, "blue": 0.800}, "대신": {"red": 0.866, "green": 0.921, "blue": 0.968}, "우리": {"red": 0.886, "green": 0.937, "blue": 0.854}}.get(courier, {"red": 0.95, "green": 0.95, "blue": 0.95})
def rgb_tuple(courier):
    d = get_rgb_dict(courier)
    return (d["red"], d["green"], d["blue"])

def get_factory_color(name):
    c_list = ["#2563EB", "#D97706", "#7C3AED", "#0D9488", "#DC2626", "#EA580C", "#0284C7", "#4F46E5", "#16A34A", "#9333EA"]
    h = sum(ord(c) for c in str(name))
    return c_list[h % len(c_list)]

def set_sd_date(d):
    st.session_state['sd_date'] = d

################################################################################
# 5. 구글 시트 안전 입출력 헬퍼
################################################################################
@st.cache_resource(ttl=3000)
def _open_sheet_cached(key): return client.open_by_key(key)

def get_sh(key):
    try: return _open_sheet_cached(key)
    except Exception: return client.open_by_key(key)

def ws_update(ws, values, range_name="A1"):
    try: return ws.update(values=values, range_name=range_name, value_input_option='RAW')
    except TypeError: return ws.update(range_name, values)

def ensure_grid(ws, rows, cols):
    try:
        if ws.row_count < rows or ws.col_count < cols:
            ws.resize(rows=max(ws.row_count, rows), cols=max(ws.col_count, cols))
    except Exception: pass

def get_or_create_ws(sh, title, headers=None, rows=1000, cols=10):
    try: return sh.worksheet(title)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=title, rows=str(rows), cols=str(max(cols, len(headers) if headers else cols)))
        if headers: ws_update(ws, [headers], "A1")
        return ws

def safe_write(sh, ws, values, old_values=None):
    values = clean_rows(values)
    if old_values is None: old_values = ws.get_all_values()
    ncols = max((len(r) for r in values), default=1)
    ensure_grid(ws, len(values) + 12, ncols)
    ws_update(ws, values, "A1")
    old_rows = len(old_values)
    old_cols = max((len(r) for r in old_values), default=0)
    ranges = []
    end = rowcol_to_a1(max(old_rows, len(values), 1), max(old_cols, ncols))
    if old_rows > len(values): ranges.append(f"A{len(values) + 1}:{end}")
    if old_cols > ncols: ranges.append(f"{rowcol_to_a1(1, ncols + 1)}:{end}")
    if ranges: ws.batch_clear(ranges)

BLACK = {"red": 0, "green": 0, "blue": 0}
GRAY = {"red": 0.8, "green": 0.8, "blue": 0.8}

def build_format_requests(sheet_id, specs, ncols, total_rows):
    reqs = [{"updateCells": {"range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": total_rows + 10, "startColumnIndex": 0, "endColumnIndex": ncols}, "fields": "userEnteredFormat(backgroundColor,borders,textFormat)"}}]
    i = 0
    while i < len(specs):
        spec = specs[i]
        j = i + 1
        if not spec[2]:
            while j < len(specs) and specs[j] == spec: j += 1
        bg = {"red": spec[0][0], "green": spec[0][1], "blue": spec[0][2]}
        reqs.append({"repeatCell": {"range": {"sheetId": sheet_id, "startRowIndex": i + 1, "endRowIndex": j + 1, "startColumnIndex": 0, "endColumnIndex": ncols},
                                    "cell": {"userEnteredFormat": {"backgroundColor": bg, "textFormat": {"foregroundColor": BLACK, "bold": spec[1]},
                                                                   "borders": {"top": {"style": "SOLID_MEDIUM" if spec[2] else "SOLID", "color": BLACK if spec[2] else GRAY},
                                                                               "bottom": {"style": "SOLID", "color": GRAY},
                                                                               "left": {"style": "SOLID_MEDIUM", "color": BLACK},
                                                                               "right": {"style": "SOLID_MEDIUM", "color": BLACK}}}},
                                    "fields": "userEnteredFormat(backgroundColor,textFormat,borders)"}})
        i = j
    return reqs

def upload_to_drive(name, data_bytes, mimetype):
    if not drive_service: return False, "Drive 연결 없음"
    try:
        meta = {'name': name, 'parents': [DRIVE_FOLDER_ID]}
        media = MediaIoBaseUpload(BytesIO(data_bytes), mimetype=mimetype or 'application/octet-stream', resumable=True)
        up = drive_service.files().create(body=meta, media_body=media, fields='id, webViewLink').execute()
        return True, up.get('webViewLink', '')
    except Exception as e: return False, str(e)

################################################################################
# 6. 데이터 페칭 (캐시) & 공급업체 상세 정보 연동
################################################################################
@st.cache_data(ttl=600)
def fetch_suppliers_df_v2():
    """채팅방목록 시트에서 [업체명, 담당자명, 메모]를 안전하게 불러옵니다 (기존 1열 시트 완벽 호환)"""
    try:
        ws = get_sh(WECHAT_SHEET_ID).worksheet("채팅방목록")
        vals = ws.get_all_values()
        if len(vals) <= 1:
            return pd.DataFrame(columns=SUPPLIER_HEADERS)
        rows = []
        for idx, r in enumerate(vals[1:], start=2):
            name = s(r[0]) if len(r) > 0 else ""
            if not name: continue
            mgr = s(r[1]) if len(r) > 1 else ""
            memo = s(r[2]) if len(r) > 2 else ""
            rows.append({"_id": idx, "업체명": name, "담당자명": mgr, "메모": memo})
        return pd.DataFrame(rows) if rows else pd.DataFrame(columns=["_id"] + SUPPLIER_HEADERS)
    except Exception:
        return pd.DataFrame(columns=["_id"] + SUPPLIER_HEADERS)

@st.cache_data(ttl=600)
def fetch_wechat_rooms_v2():
    df = fetch_suppliers_df_v2()
    if df.empty or "업체명" not in df.columns: return []
    return [str(v) for v in df["업체명"].tolist() if str(v).strip()]

@st.cache_data(ttl=600)
def fetch_wechat_history_v2():
    try:
        vals = get_sh(WECHAT_SHEET_ID).worksheet("위챗기록").get_all_values()
        if len(vals) > 1: return pd.DataFrame(vals[1:], columns=vals[0]).fillna('')
    except Exception: pass
    return pd.DataFrame(columns=["날짜", "공장", "태그", "요약제목", "상세내용"])

@st.cache_data(ttl=600)
def fetch_tasks_v2():
    try:
        sh = get_sh(WECHAT_SHEET_ID)
        vals = sh.worksheet("할일").get_all_values()
        if len(vals) > 1:
            df = rows_to_df(vals, TASK_HEADERS)
            df['_row'] = range(2, len(df) + 2)
            return df
    except Exception: pass
    df = pd.DataFrame(columns=TASK_HEADERS)
    df['_row'] = []
    return df

@st.cache_data(ttl=600)
def fetch_price_v2():
    try:
        vals = get_sh(WECHAT_SHEET_ID).worksheet("단가이력").get_all_values()
        if len(vals) > 1: return rows_to_df(vals, PRICE_HEADERS)
    except Exception: pass
    return pd.DataFrame(columns=PRICE_HEADERS)

@st.cache_data(ttl=600)
def fetch_customer_data_v2():
    try:
        vals = get_sh(CUST_SHEET_ID).sheet1.get_all_values()
        if len(vals) > 1: return pd.DataFrame(vals[1:], columns=vals[0]).fillna('')
    except Exception: pass
    return pd.DataFrame(columns=CUST_HEADERS)

@st.cache_data(ttl=600)
def fetch_gy_history_v2(month_str):
    try:
        sh = get_sh(GY_SHEET_ID)
        try: ws = sh.worksheet(month_str)
        except gspread.exceptions.WorksheetNotFound: return pd.DataFrame(columns=GY_HEADERS)
        vals = ws.get_all_values()
        if len(vals) > 1: return pd.DataFrame(vals[1:], columns=vals[0]).fillna('')
    except Exception: pass
    return pd.DataFrame(columns=GY_HEADERS)

@st.cache_data(ttl=600)
def fetch_tracking_data_v2(sheet_name):
    try:
        sh = get_sh(TRACKING_SHEET_ID)
        try: ws = sh.worksheet(sheet_name)
        except gspread.exceptions.WorksheetNotFound: return pd.DataFrame(columns=TRACK_HEADERS)
        vals = ws.get_all_values()
        if len(vals) > 1: return pd.DataFrame(vals[1:], columns=vals[0]).fillna('')
    except Exception: pass
    return pd.DataFrame(columns=TRACK_HEADERS)

def clear_all_caches():
    for fn in (fetch_customer_data_v2, fetch_gy_history_v2, fetch_tracking_data_v2,
               fetch_suppliers_df_v2, fetch_wechat_rooms_v2, fetch_wechat_history_v2, fetch_tasks_v2, fetch_price_v2):
        try: fn.clear()
        except Exception: pass

@st.cache_data(ttl=600)
def fetch_drive_archive_v2(search_query=""):
    if not drive_service: return pd.DataFrame()
    try:
        q = "trashed = false"
        if search_query:
            safe_q = search_query.replace("\\", "\\\\").replace("'", "\\'")
            q += f" and (name contains '{safe_q}' or fullText contains '{safe_q}')"
        else:
            q += " and fullText contains '수원자료'"
            
        results = drive_service.files().list(
            q=q,
            fields="files(id, name, mimeType, webViewLink, modifiedTime)",
            pageSize=50,
            orderBy="modifiedTime desc"
        ).execute()
        
        items = results.get('files', [])
        if not items: return pd.DataFrame()
        
        df = pd.DataFrame(items)
        df['modifiedTime'] = pd.to_datetime(df['modifiedTime']).dt.tz_convert('Asia/Seoul').dt.strftime('%Y-%m-%d %H:%M')
        def get_icon(mime):
            if 'folder' in mime: return '📁 폴더'
            elif 'sheet' in mime or 'excel' in mime: return '📊 엑셀/시트'
            elif 'pdf' in mime: return '📕 PDF'
            elif 'image' in mime: return '🖼️ 이미지'
            elif 'zip' in mime: return '📦 압축파일'
            return '📄 문서'
        df['종류'] = df['mimeType'].apply(get_icon)
        return df[['modifiedTime', '종류', 'name', 'webViewLink']]
    except Exception:
        return pd.DataFrame()

def rename_supplier_across_sheets(sh, rename_map):
    """공급업체 이름이 변경되었을 때 위챗기록, 할일, 단가이력 시트의 옛날 업체명도 자동 일괄 변경합니다."""
    if not rename_map: return
    for ws_name, col_name in [("위챗기록", "공장"), ("할일", "공장"), ("단가이력", "공장")]:
        try:
            ws = sh.worksheet(ws_name)
            vals = ws.get_all_values()
            if len(vals) <= 1: continue
            hdr = vals[0]
            if col_name not in hdr: continue
            c_idx = hdr.index(col_name)
            updates = []
            for r_i, row in enumerate(vals[1:], start=2):
                if len(row) > c_idx:
                    cur_val = s(row[c_idx])
                    if cur_val in rename_map:
                        updates.append({'range': rowcol_to_a1(r_i, c_idx + 1), 'values': [[rename_map[cur_val]]]})
            if updates:
                ws.batch_update(updates, value_input_option='RAW')
        except Exception:
            pass

################################################################################
# 7. 화물택배(건영) / 출고 액션 함수
################################################################################
def update_totals():
    qs = [safe_int(st.session_state.get(i["key"], 0)) for i in RATE_ITEMS]
    st.session_state.in_qty = max(1, sum(qs))
    st.session_state.in_fee = sum(q * i["rate"] for q, i in zip(qs, RATE_ITEMS))
    parts = []
    for q, i in zip(qs, RATE_ITEMS):
        if q:
            memo_label = st.session_state.get("wf_type", "휠") if i["key"] == "q_wheel" else i["memo"]
            parts.append(f"{memo_label}x{q}")
    st.session_state.in_memo = f"[{', '.join(parts)}] " if parts else ""

def on_customer_select():
    df = fetch_customer_data_v2()
    sel = st.session_state.cust_select
    if sel != "직접 입력":
        rows = df[df['이름'] == sel] if '이름' in df.columns else df.iloc[0:0]
        if rows.empty: return
        row = rows.iloc[0]
        st.session_state.in_name = sel
        st.session_state.in_phone = str(row.get('전화번호', ''))
        st.session_state.in_addr = str(row.get('주소', ''))
    else:
        st.session_state.in_name, st.session_state.in_phone, st.session_state.in_addr = "", "", ""

def save_draft(lst):
    try:
        ws = get_or_create_ws(get_sh(GY_SHEET_ID), "임시저장", None, rows=10, cols=2)
        ws_update(ws, [[json.dumps(lst, ensure_ascii=False)]], "A1")
    except Exception: pass

def load_draft():
    try:
        ws = get_sh(GY_SHEET_ID).worksheet("임시저장")
        txt = ws.acell("A1").value
        if txt:
            data = json.loads(txt)
            if isinstance(data, list): return data
    except Exception: pass
    return []

def add_to_list():
    name, inv = s(st.session_state.in_name), s(st.session_state.in_inv)
    if not name or not inv:
        st.error("업체명과 운송장번호를 입력하세요."); return
    repeat = max(1, min(30, safe_int(st.session_state.in_repeat, 1)))
    existing = {i['inv_no'] for i in st.session_state['deli_list']}
    added, dups = 0, []
    for n in range(repeat):
        cur = inc_inv(inv, n) if n else inv
        if cur in existing:
            dups.append(cur); continue
        st.session_state['deli_list'].append({'name': name, 'phone': s(st.session_state.in_phone), 'addr': s(st.session_state.in_addr), 'inv_no': cur,
                                              'qty': st.session_state.in_qty, 'fee': st.session_state.in_fee, 'memo': s(st.session_state.in_memo)})
        existing.add(cur); added += 1
    if added == 0:
        st.warning("⚠️ 이미 작업 리스트에 존재하는 송장번호입니다!"); return
    if dups: st.warning(f"⚠️ 중복 송장 제외: {', '.join(dups)}")
    st.session_state.in_inv = inc_inv(inv, repeat)
    for k in ['in_name', 'in_phone', 'in_addr', 'in_memo']: st.session_state[k] = ""
    for k in ['q_base', 'q_wheel', 'q_shop', 'q_personal', 'in_fee']: st.session_state[k] = 0
    st.session_state.in_qty = 1; st.session_state.in_repeat = 1; st.session_state.cust_select = "직접 입력"
    save_draft(st.session_state['deli_list'])

def sync_customer_database_batch(items):
    if not client or not items: return True, ""
    try:
        ws = get_sh(CUST_SHEET_ID).sheet1
        vals = ws.get_all_values()
        if not vals:
            ws_update(ws, [CUST_HEADERS], "A1"); vals = [CUST_HEADERS]
        headers = vals[0]
        for h in ('이름', '전화번호', '주소'):
            if h not in headers: return False, f"고객DB 첫 행에 '{h}' 열이 없습니다."
        ci = {}
        for i, h in enumerate(headers):
            if h and h not in ci: ci[h] = i
        name_row = {}
        for r_i, row in enumerate(vals[1:], start=2):
            nm = row[ci['이름']].strip() if len(row) > ci['이름'] else ''
            if nm and nm not in name_row: name_row[nm] = (r_i, row)
        updates, new_rows, seen = [], [], set()
        for item in items:
            name, phone, addr = s(item.get('name')), s(item.get('phone')), s(item.get('addr'))
            if not name: continue
            if name in name_row:
                r_i, row = name_row[name]
                for col, val in (('전화번호', phone), ('주소', addr)):
                    if not val: continue
                    cur = row[ci[col]] if len(row) > ci[col] else ''
                    if cur != val: updates.append({'range': rowcol_to_a1(r_i, ci[col] + 1), 'values': [[val]]})
            elif name not in seen:
                seen.add(name)
                nr = [''] * len(headers)
                for col, val in (('종류', '개인(자동저장)'), ('이름', name), ('전화번호', phone), ('주소', addr)):
                    if col in ci: nr[ci[col]] = val
                new_rows.append(nr)
        if updates: ws.batch_update(updates, value_input_option='RAW')
        if new_rows: ws.append_rows(new_rows, value_input_option='RAW')
        return True, ""
    except Exception as e: return False, str(e)

def organize_and_format_gsheet(sheet_name):
    if not client: return False, "구글 서버 연결 실패"
    try:
        sh = get_sh(TRACKING_SHEET_ID)
        try: ws = sh.worksheet(sheet_name)
        except gspread.exceptions.WorksheetNotFound: return True, "데이터 없음"
        data = ws.get_all_values()
        if len(data) <= 1: return True, "데이터 없음"
        df = rows_to_df(data, TRACK_HEADERS)
        valid = (df['받는사람'].map(s) != '') | (df['트레킹번호'].map(s) != '')
        df = df[valid].copy()
        df['날짜'] = df['날짜'].map(norm_date).replace('', pd.NA).ffill().fillna('')
        df = df.sort_values(by=['날짜', '택배사'], ascending=[True, True], kind='stable').reset_index(drop=True)
        show_date = df['날짜'].where(~df['날짜'].duplicated(), '')
        out_data = [TRACK_HEADERS]
        specs = []
        for i, row in df.iterrows():
            r = [to_native(x) for x in row.tolist()]
            r[0] = show_date[i]
            out_data.append(r)
            is_new = (show_date[i] != '')
            specs.append((rgb_tuple(s(row['택배사'])), is_new, is_new))
        safe_write(sh, ws, out_data, old_values=data)
        requests = build_format_requests(ws.id, specs, len(TRACK_HEADERS), len(out_data))
        requests.append({"updateBorders": {"range": {"sheetId": ws.id, "startRowIndex": len(out_data) - 1, "endRowIndex": len(out_data), "startColumnIndex": 0, "endColumnIndex": len(TRACK_HEADERS)}, "bottom": {"style": "SOLID_MEDIUM", "color": BLACK}}})
        sh.batch_update({'requests': requests})
        return True, "성공"
    except Exception as e: return False, str(e)

def append_to_tracking_gsheet(new_data_list, sheet_name):
    if not client: return
    sh = get_sh(TRACKING_SHEET_ID)
    ws = get_or_create_ws(sh, sheet_name, TRACK_HEADERS, rows=1000, cols=10)
    data = ws.get_all_values()
    df_existing = rows_to_df(data, TRACK_HEADERS) if len(data) > 1 else pd.DataFrame(columns=TRACK_HEADERS)
    df_new = pd.DataFrame(new_data_list, columns=TRACK_HEADERS).fillna('')
    df_new['트레킹번호'] = df_new['트레킹번호'].map(fix_trk)
    if not df_existing.empty:
        new_nos = [t for t in df_new['트레킹번호'].tolist() if s(t) != '']
        if new_nos: df_existing = df_existing[~df_existing['트레킹번호'].map(s).isin(new_nos)]
        df_combined = pd.concat([df_existing, df_new], ignore_index=True)
    else: df_combined = df_new
    safe_write(sh, ws, [TRACK_HEADERS] + df_combined.values.tolist(), old_values=data)
    organize_and_format_gsheet(sheet_name)

def sync_gy_and_format(new_items, save_date):
    month_str = f"건영-{save_date.year}-{save_date.month}"
    if not client: return
    sh = get_sh(GY_SHEET_ID)
    ws = get_or_create_ws(sh, month_str, GY_HEADERS, rows=1000, cols=10)
    data = ws.get_all_values()
    df = rows_to_df(data, GY_HEADERS) if len(data) > 1 else pd.DataFrame(columns=GY_HEADERS)
    df = df[df['운송장번호'].map(s) != 'TOTAL']
    if new_items:
        sd_str = save_date.strftime('%Y-%m-%d')
        new_df = pd.DataFrame([[sd_str, i['name'], i['phone'], i['addr'], i['inv_no'], safe_int(i['qty'], 1), safe_int(i['fee']), i['memo']] for i in new_items], columns=GY_HEADERS)
        df = pd.concat([df, new_df], ignore_index=True)
        inv_ok = df['운송장번호'].map(s) != ''
        df = df[~(inv_ok & df.duplicated(subset=['운송장번호'], keep='last'))] 
    df['날짜'] = df['날짜'].map(norm_date).replace('', pd.NA).ffill().fillna('')
    df = df.sort_values(by=['날짜'], kind='stable').reset_index(drop=True)
    df['수량'] = pd.to_numeric(df['수량'], errors='coerce').fillna(0).astype(int)
    df['운임'] = pd.to_numeric(df['운임'], errors='coerce').fillna(0).astype(int)
    final_data = [GY_HEADERS]
    curr_date, d_qty, d_fee = None, 0, 0
    for _, row in df.iterrows():
        r_date = str(row['날짜'])
        if curr_date != r_date:
            if curr_date is not None: final_data.append([f"{curr_date} 마감", "당일 발송 총계 ➔", "", "", "TOTAL", d_qty, d_fee, ""])
            curr_date = r_date; d_qty = 0; d_fee = 0
        final_data.append(row.tolist())
        d_qty += int(row['수량']); d_fee += int(row['운임'])
    if curr_date is not None: final_data.append([f"{curr_date} 마감", "당일 발송 총계 ➔", "", "", "TOTAL", d_qty, d_fee, ""])
    safe_write(sh, ws, final_data, old_values=data)
    specs, prev_date = [], None
    for row in final_data[1:]:
        is_total = (str(row[4]).strip() == 'TOTAL')
        row_date = str(row[0]).replace(" 마감", "")
        is_new_date = (row_date != prev_date) and not is_total
        specs.append(((0.85, 0.85, 0.85) if is_total else (1.0, 1.0, 1.0), is_total, is_new_date))
        if not is_total: prev_date = row_date
    sh.batch_update({'requests': build_format_requests(ws.id, specs, len(GY_HEADERS), len(final_data))})

def month_stats(df):
    empty_top = pd.DataFrame(columns=['받는 곳', '박스수'])
    if df is None or df.empty or '운송장번호' not in df.columns: return 0, 0, empty_top
    real = df[df['운송장번호'].astype(str).str.strip() != 'TOTAL'].copy()
    real['수량'] = pd.to_numeric(real['수량'], errors='coerce').fillna(0)
    real['운임'] = pd.to_numeric(real['운임'], errors='coerce').fillna(0)
    named = real[real['받는 곳'].map(s) != '']
    top = named.groupby('받는 곳')['수량'].sum().sort_values(ascending=False).head(5).reset_index()
    top.columns = ['받는 곳', '박스수']
    return int(real['수량'].sum()), int(real['운임'].sum()), top

################################################################################
# 8. 인쇄 이미지 생성 (운송장번호 제거 -> 상품정보 추가 & 자동 줄바꿈/행높이 확장)
################################################################################
def wrap_text_by_pixel(draw, text, font, max_width):
    """글자 폭(픽셀)을 계산하여 칸 너비를 넘어가면 잘리지 않게 자동으로 줄바꿈해주는 함수"""
    text = s(text)
    if not text:
        return [""]
    lines = []
    for paragraph in text.split('\n'):
        if not paragraph:
            lines.append("")
            continue
        current_line = ""
        for char in paragraph:
            test_line = current_line + char
            bbox = draw.textbbox((0, 0), test_line, font=font)
            if (bbox[2] - bbox[0]) <= max_width:
                current_line = test_line
            else:
                if current_line:
                    lines.append(current_line)
                current_line = char
        if current_line:
            lines.append(current_line)
    return lines if lines else [""]

def generate_courier_images(deli_list):
    if not deli_list: return [], "출력할 목록이 없습니다."
    if not os.path.exists(FONT_BOLD) or not os.path.exists(FONT_REG):
        return [], "폰트 파일(NotoSansKR-Bold.ttf / NotoSansKR-Regular.ttf)을 app.py 와 같은 폴더에 두세요."
    try:
        width, height = 3508, 2480
        f_title, f_head, f_th, f_txt, f_foot = [
            ImageFont.truetype(f, sz) for f, sz in [
                (FONT_BOLD, 100), (FONT_BOLD, 45), (FONT_BOLD, 40), (FONT_REG, 36), (FONT_BOLD, 50)
            ]
        ]

        cols = [
            ("NO", 150, "mm", 80),
            ("받는 곳 (업체명)", 230, "lm", 430),
            ("연락처", 690, "lm", 360),
            ("주소", 1080, "lm", 1060),
            ("상품정보 (내역)", 2180, "lm", 780),
            ("수량", 3060, "mm", 130),
            ("운임", 3400, "rs", 200)
        ]

        dummy_img = Image.new('RGB', (10, 10))
        dummy_draw = ImageDraw.Draw(dummy_img)

        processed_rows = []
        sum_qty, sum_fee = 0, 0
        line_height = 48

        for idx, row in enumerate(deli_list):
            q = safe_int(row.get('qty', 0))
            fee = safe_int(row.get('fee', 0))
            sum_qty += q
            sum_fee += fee

            name_lines = wrap_text_by_pixel(dummy_draw, row.get('name', ''), f_txt, cols[1][3])
            phone_lines = wrap_text_by_pixel(dummy_draw, row.get('phone', ''), f_txt, cols[2][3])
            addr_lines = wrap_text_by_pixel(dummy_draw, row.get('addr', ''), f_txt, cols[3][3])
            memo_lines = wrap_text_by_pixel(dummy_draw, row.get('memo', ''), f_txt, cols[4][3])

            max_lines = max(len(name_lines), len(phone_lines), len(addr_lines), len(memo_lines), 1)
            row_h = max(115, max_lines * line_height + 36)

            processed_rows.append({
                "no": str(idx + 1),
                "name": name_lines,
                "phone": phone_lines,
                "addr": addr_lines,
                "memo": memo_lines,
                "qty": str(q),
                "fee": f"{fee:,}",
                "row_h": row_h
            })

        y_tbl = 580
        header_h = 115
        max_y = height - 340

        pages_data = []
        cur_page = []
        cur_y = y_tbl + header_h

        for prow in processed_rows:
            if cur_page and (cur_y + prow["row_h"] > max_y):
                pages_data.append(cur_page)
                cur_page = [prow]
                cur_y = y_tbl + header_h + prow["row_h"]
            else:
                cur_page.append(prow)
                cur_y += prow["row_h"]
        if cur_page:
            pages_data.append(cur_page)

        total_pages = max(1, len(pages_data))
        img_base64_list = []

        for p_idx, page_rows in enumerate(pages_data):
            img = Image.new('RGB', (width, height), (255, 255, 255))
            draw = ImageDraw.Draw(img)

            if os.path.exists(LOGO_FILE_PNG):
                try:
                    logo = Image.open(LOGO_FILE_PNG).convert("RGBA")
                    logo.thumbnail((2000, 420))
                    img.paste(logo, (100, 80), logo)
                except Exception:
                    pass
            else:
                draw.text((100, 80), "UNKNOWN BIKE", font=f_title, fill=(0, 0, 0))

            draw.text((3400, 80), "건영택배 일일배송요청서", font=f_title, fill=(0, 0, 0), anchor="rs")
            draw.text((3400, 220), f"발송일자: {now_kst().strftime('%Y-%m-%d')}", font=f_head, fill=(100, 100, 100), anchor="rs")
            draw.text((3400, 300), "보내는 사람 : 언노운바이크\n주소 : 경기도 양주시 삼숭로 58번길 84-10\n연락처 : 010-4214-3944", font=f_head, fill=(0, 0, 0), anchor="rs", align="right", spacing=20)
            draw.text((3400, 500), f"[{p_idx + 1} / {total_pages} 페이지]", font=f_head, fill=(50, 50, 50), anchor="rs")
            draw.line([(100, 530), (width - 100, 530)], fill=(0, 0, 0), width=5)

            draw.rectangle([(100, y_tbl), (width - 100, y_tbl + header_h)], fill=(230, 230, 230))
            for t, x, a, _ in cols:
                draw.text((x, y_tbl + header_h // 2), t, font=f_th, fill=(0, 0, 0), anchor=a)

            y_r = y_tbl + header_h
            for prow in page_rows:
                rh = prow["row_h"]
                cy = y_r + rh // 2

                draw.text((cols[0][1], cy), prow["no"], font=f_txt, fill=(0, 0, 0), anchor="mm")
                draw.text((cols[5][1], cy), prow["qty"], font=f_txt, fill=(0, 0, 0), anchor="mm")
                draw.text((cols[6][1], cy), prow["fee"], font=f_txt, fill=(0, 0, 0), anchor="rm")

                for col_idx, key in [(1, "name"), (2, "phone"), (3, "addr"), (4, "memo")]:
                    lines = prow[key]
                    total_text_h = len(lines) * line_height
                    start_y = y_r + (rh - total_text_h) // 2 + (line_height // 2)
                    for l_idx, line_str in enumerate(lines):
                        draw.text((cols[col_idx][1], start_y + l_idx * line_height), line_str, font=f_txt, fill=(0, 0, 0), anchor="lm")

                draw.line([(100, y_r + rh), (width - 100, y_r + rh)], fill=(200, 200, 200), width=1)
                y_r += rh

            if p_idx == total_pages - 1:
                draw.rectangle([(100, y_r), (width - 100, y_r + header_h)], fill=(240, 240, 240))
                draw.text((cols[4][1] + 500, y_r + header_h // 2), "TOTAL :", font=f_th, fill=(0, 0, 0), anchor="mm")
                draw.text((cols[5][1], y_r + header_h // 2), str(sum_qty), font=f_th, fill=(0, 0, 0), anchor="mm")
                draw.text((cols[6][1], y_r + header_h // 2), f"{sum_fee:,}", font=f_th, fill=(0, 0, 0), anchor="rm")

            draw.text((width // 2, height - 130), "위와 같이 배송을 요청합니다.", font=f_foot, fill=(0, 0, 0), anchor="mm")

            buffer = BytesIO()
            img.save(buffer, format="JPEG", quality=95)
            img_base64_list.append(base64.b64encode(buffer.getvalue()).decode())

        return img_base64_list, ""
    except Exception as e:
        return [], f"이미지 생성 오류: {e}"

def build_pdf_from_b64(imgs):
    pil = [Image.open(BytesIO(base64.b64decode(b))).convert("RGB") for b in imgs]
    buf = BytesIO()
    pil[0].save(buf, format="PDF", save_all=True, append_images=pil[1:], resolution=300.0)
    return buf.getvalue()


################################################################################
# 9. 새 위챗·생산관리 연결 (원본 기록 읽기 + 별도 V2 저장소)
################################################################################
def call_gemini(prompt, images):
    key = sec("GEMINI_API_KEY", "")
    if not key:
        raise RuntimeError("GEMINI_API_KEY가 없습니다. 원문 초안을 먼저 저장하고 관리자에게 문의하세요.")
    api = genai.Client(api_key=key, http_options={"timeout": 90000})
    names = list(dict.fromkeys([GEMINI_MODEL] + [x.strip() for x in sec("GEMINI_FALLBACK_MODELS", "").split(",") if x.strip()]))
    failures = []
    try:
        for name in names:
            try:
                response = api.models.generate_content(model=name, contents=[prompt] + images,
                    config={"response_mime_type": "application/json", "temperature": 0.1})
                if not response.text:
                    raise RuntimeError("AI 응답이 비어 있습니다.")
                return response.text, name
            except Exception as err:
                failures.append(f"{name}: {type(err).__name__}")
        raise RuntimeError("AI 분석을 완료하지 못했습니다. 모델 설정·권한·사용량을 확인하세요. " + "; ".join(failures))
    finally:
        api.close()


def workflow_context():
    return {
        "key": sec("PRODUCTION_SHEET_ID", WECHAT_SHEET_ID),
        "wechat_key": WECHAT_SHEET_ID,
        "get_sh": get_sh,
        "wechat_password": WECHAT_PASSWORD,
        "read_only": sec("NEW_WORKFLOW_READ_ONLY", "true").lower() not in ("false", "0", "no"),
        "call_ai": call_gemini,
        "upload_to_drive": upload_to_drive,
        "drive_service": drive_service,
        "drive_folder": DRIVE_FOLDER_ID,
        "gspread_client": client,
        "restore_share_email": sec("RESTORE_SHARE_EMAIL", ""),
        "backup_sources": {
            "택배": TRACKING_SHEET_ID, "거래처": CUST_SHEET_ID, "건영": GY_SHEET_ID,
            "위챗": WECHAT_SHEET_ID, "생산": sec("PRODUCTION_SHEET_ID", WECHAT_SHEET_ID),
        },
    }

################################################################################
# 10. 사이드바 (이모티콘 제거, 가독성 개편, 보안안내 제거)
################################################################################
with st.sidebar:
    st.markdown('<p class="menu-title">WORK MENU</p>', unsafe_allow_html=True)
    menu = st.radio("M", ["화물택배발송", "일일출고관리", "생산관리", "위챗요약"], label_visibility="collapsed")
    if APP_PASSWORD and st.button("로그아웃", use_container_width=True):
        st.session_state['app_auth'] = False
        st.rerun()

show_flash()

################################################################################
# 11. 메인 메뉴 1: 화물택배발송
################################################################################
if menu == "화물택배발송":
    col_t, col_b = st.columns([3, 1])
    col_t.title("화물택배 발송 시스템")
    _now = now_kst()
    current_month_str = f"건영-{_now.year}-{_now.month}"
    if col_b.button("✨ 건영 시트 꼬임 복구 (리프레쉬)", use_container_width=True):
        try:
            with st.spinner("엑셀 데이터를 재정렬 중입니다..."):
                sync_gy_and_format([], _now)
                clear_all_caches()
            flash("success", "건영택배 엑셀 정비 완료!")
        except Exception:
            st.error("🚨 정비 중 오류가 발생했습니다.")
            st.code(traceback.format_exc())
        else:
            st.rerun()

    tab1, tab2 = st.tabs(["📝 금일 발송 작업", "📅 과거 기록 조회 및 복원"])

    with tab1:
        if not st.session_state['draft_checked']:
            st.session_state['draft_checked'] = True
            if not st.session_state['deli_list']:
                _d = load_draft()
                if _d: st.session_state['draft_found'] = _d
        if st.session_state['draft_found'] and not st.session_state['deli_list']:
            dc1, dc2, dc3 = st.columns([3, 1, 1])
            dc1.info(f"💾 이전에 작업하던 임시저장 목록이 있습니다 ({len(st.session_state['draft_found'])}건)")
            if dc2.button("불러오기", use_container_width=True):
                st.session_state['deli_list'] = st.session_state['draft_found']
                st.session_state['draft_found'] = None
                st.rerun()
            if dc3.button("버리기", use_container_width=True):
                save_draft([])
                st.session_state['draft_found'] = None
                st.rerun()

        dash_df = fetch_gy_history_v2(current_month_str)
        month_box_total, month_fee_total, top5 = month_stats(dash_df)
        _prev = _now.replace(day=1) - timedelta(days=1)
        prev_box, prev_fee, _ = month_stats(fetch_gy_history_v2(f"건영-{_prev.year}-{_prev.month}"))

        d_col1, d_col2, d_col3 = st.columns(3)
        d_col1.markdown(f'<div class="metric-card"><div class="metric-title">📅 조회 기준</div><div class="metric-value">{current_month_str} 탭 기록</div></div>', unsafe_allow_html=True)
        d_col2.markdown(f'<div class="metric-card"><div class="metric-title">📦 이번 달 총 발송 박스</div><div class="metric-value">{month_box_total} 개</div></div>', unsafe_allow_html=True)
        d_col3.markdown(f'<div class="metric-card"><div class="metric-title">💳 이번 달 예상 총 운임비</div><div class="metric-value">{month_fee_total:,} 원</div></div>', unsafe_allow_html=True)
        
        with st.expander("📊 월별 통계 (지난달 대비 · 업체별 TOP 5)"):
            m1, m2 = st.columns(2)
            m1.metric("발송 박스", f"{month_box_total} 개", f"{month_box_total - prev_box:+d} (지난달 {prev_box}개)")
            m2.metric("운임비", f"{month_fee_total:,} 원", f"{month_fee_total - prev_fee:+,} (지난달 {prev_fee:,}원)")
            if not top5.empty:
                st.markdown("**🏆 이번 달 발송량 TOP 5**")
                st.dataframe(top5, hide_index=True, use_container_width=True)
                st.bar_chart(top5.set_index('받는 곳'))
            else:
                st.caption("이번 달 기록이 없습니다.")

        # 💡 [입력 폼 카드 그룹핑으로 시선 집중도 향상]
        with st.container(border=True):
            cust_df = fetch_customer_data_v2()
            cust_names = cust_df['이름'].tolist() if (not cust_df.empty and '이름' in cust_df.columns) else []
            st.selectbox("🔍 수령인 검색 (업체정보 DB 자동 연동)", ["직접 입력"] + cust_names, key="cust_select", on_change=on_customer_select)

            c1, c2 = st.columns(2)
            c1.text_input("업체/개인명 (지정 시 DB에 자동 저장됩니다)", key="in_name")
            c2.text_input("운송장번호", key="in_inv")
            col_p, col_a = st.columns(2)
            col_p.text_input("연락처", key="in_phone")
            col_a.text_input("주소", key="in_addr")

            st.divider()
            qcols = st.columns(len(RATE_ITEMS))
            for qc, item in zip(qcols, RATE_ITEMS):
                if item['key'] == 'q_wheel':
                    qc.radio("종류 선택", ["휠", "프레임"], horizontal=True, key="wf_type", on_change=update_totals, label_visibility="collapsed")
                qc.number_input(f"{item['label']}({item['rate']:,}원)", min_value=0, key=item['key'], on_change=update_totals)

            cc1, cc2, cc3, cc4 = st.columns([1, 1, 2, 1])
            cc1.number_input("최종 수량", min_value=1, key="in_qty")
            cc2.number_input("최종 운임", min_value=0, step=100, key="in_fee")
            cc3.text_input("비고 (인쇄 시 상품정보 칸에 출력됩니다)", key="in_memo")
            cc4.number_input("연속 N건 추가", min_value=1, max_value=30, key="in_repeat", help="같은 업체에 송장번호가 연속인 박스를 여러 개 보낼 때 (각 건에 동일 수량·운임 적용)")

            st.button("➕ 목록에 추가", type="primary", on_click=add_to_list)

        if st.session_state['deli_list']:
            st.write("")
            df_display = pd.DataFrame(st.session_state['deli_list'])
            df_display = df_display[['name', 'phone', 'addr', 'inv_no', 'qty', 'fee', 'memo']]
            df_display.columns = ['업체명', '연락처', '주소', '운송장번호', '수량', '운임', '비고']

            edited_df = st.data_editor(
                df_display, num_rows="dynamic", use_container_width=True,
                column_config={
                    "업체명": st.column_config.TextColumn("업체명", width=140),
                    "연락처": st.column_config.TextColumn("연락처", width=150),
                    "주소": st.column_config.TextColumn("주소", width=320),
                    "운송장번호": st.column_config.TextColumn("운송장번호", width=180),
                    "수량": st.column_config.NumberColumn("수량", width=70),
                    "운임": st.column_config.NumberColumn("운임", width=90),
                    "비고": st.column_config.TextColumn("비고", width=250)
                }
            )
            _new_list = []
            for _, r in edited_df.iterrows():
                nm, iv = s(r['업체명']), s(r['운송장번호'])
                if not nm and not iv: continue
                _new_list.append({'name': nm, 'phone': s(r['연락처']), 'addr': s(r['주소']), 'inv_no': iv,
                                  'qty': safe_int(r['수량'], 1), 'fee': safe_int(r['운임'], 0), 'memo': s(r['비고'])})
            st.session_state['deli_list'] = _new_list

            ca, cb = st.columns(2)
            with ca:
                if st.button("🖨️ 인쇄 화면 띄우기 (로컬 PDF 자동저장)", use_container_width=True):
                    with st.spinner("출력 화면을 준비하고 있습니다..."):
                        imgs, err = generate_courier_images(st.session_state['deli_list'])
                        if imgs: 
                            st.session_state['print_imgs'] = imgs
                            try:
                                pdf_dir = r"C:\Users\AD\Documents\건영택배 일일출고리스트"
                                os.makedirs(pdf_dir, exist_ok=True)
                                pdf_path = os.path.join(pdf_dir, f"건영택배_일일배송요청서_{now_kst().strftime('%Y%m%d_%H%M%S')}.pdf")
                                pdf_bytes = build_pdf_from_b64(imgs)
                                with open(pdf_path, "wb") as f:
                                    f.write(pdf_bytes)
                                flash("success", f"✅ PDF가 로컬 PC에 자동 저장되었습니다:\n{pdf_path}")
                            except Exception as e:
                                flash("warning", f"⚠️ 로컬 자동 저장에 실패했습니다. 수동으로 다운로드해주세요. ({e})")
                        else: st.error(f"🚨 {err}")
                if st.button("💾 임시저장 (새로고침 대비)", use_container_width=True):
                    save_draft(st.session_state['deli_list'])
                    st.success("임시저장 완료")
            with cb:
                save_date = st.date_input("마감 날짜", now_kst().date(), label_visibility="collapsed")
                if st.button("💾 구글 시트 일괄 저장 (API 과부하 방어)", use_container_width=True, type="primary"):
                    target_month = f"건영-{save_date.year}-{save_date.month}"
                    df_hist = fetch_gy_history_v2(target_month)
                    new_items_to_save, dupe_invs = [], []
                    if not df_hist.empty:
                        saved_invs = set(df_hist['운송장번호'].tolist())
                        for item in st.session_state['deli_list']:
                            if item['inv_no'] in saved_invs: dupe_invs.append(item['inv_no'])
                            else: new_items_to_save.append(item)
                    else: new_items_to_save = list(st.session_state['deli_list'])

                    if not new_items_to_save:
                        st.error("⚠️ 리스트의 모든 송장이 이미 이번 달 구글 시트에 저장되어 있습니다.")
                    else:
                        try:
                            with st.spinner("데이터 저장 중입니다. 잠시만 기다려주세요..."):
                                ok_c, msg_c = sync_customer_database_batch(new_items_to_save)
                                sync_gy_and_format(new_items_to_save, save_date)
                                new_track_items = [{'날짜': save_date.strftime('%Y-%m-%d'), '트레킹번호': i['inv_no'], '받는사람': i['name'], '택배사': '건영', '박스수량': i['qty'], '내용물': i['memo'], '메모': ''} for i in new_items_to_save]
                                append_to_tracking_gsheet(new_track_items, f"{save_date.year}-{save_date.month}")
                                clear_all_caches()
                            st.session_state['deli_list'] = []
                            st.session_state['print_imgs'] = None
                            save_draft([])
                            if dupe_invs: flash("success", f"✅ 중복 송장 제외, {len(new_items_to_save)}건 동기화 및 엑셀 디자인 정렬 완료!")
                            else: flash("success", "✅ 발송 마감 완료! 3곳의 시트에 깔끔하게 분산 저장되었습니다.")
                            if not ok_c: flash("warning", f"⚠️ 고객 DB 자동 저장은 실패했습니다 (발송 기록은 저장됨): {msg_c}")
                        except Exception:
                            st.error("🚨 저장 중 오류가 발생했습니다. 목록은 그대로 유지됩니다. 다시 시도하면 이미 저장된 송장은 자동으로 제외됩니다.")
                            st.code(traceback.format_exc())
                        else:
                            st.rerun()

        if st.session_state['print_imgs']:
            imgs = st.session_state['print_imgs']
            html_content = f"""
            <!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8"><style>
            body {{ margin: 0; padding: 10px; background-color: #f4f4f4; text-align: center; font-family: sans-serif; }}
            .print-btn {{ display: block; width: 100%; max-width: 800px; margin: 0 auto 5px auto; background-color: #cc0000; color: white; border: none; padding: 15px; font-size: 18px; font-weight: bold; border-radius: 8px; cursor: pointer; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }}
            .print-btn:hover {{ background-color: #aa0000; }}
            .desc {{ color: #666; font-size: 14px; margin-bottom: 20px; }}
            img {{ width: 100%; max-width: 1200px; margin-bottom: 20px; box-shadow: 0 2px 5px rgba(0,0,0,0.2); }}
            @media print {{ .print-btn, .desc {{ display: none !important; }} body {{ margin: 0 !important; padding: 0 !important; background-color: white !important; }} img {{ width: 100% !important; max-width: none !important; height: 98vh !important; object-fit: contain !important; page-break-after: always !important; box-shadow: none !important; margin: 0 !important; }} @page {{ size: landscape; margin: 0; }} }}
            </style></head><body>
            <button class="print-btn" onclick="window.print()">🖨️ 이곳을 눌러 바로 가로 인쇄하기</button>
            <p class="desc">(버튼을 누르거나 Ctrl+P를 누르시면 됩니다)</p>
            {''.join([f'<img src="data:image/jpeg;base64,{img}">' for img in imgs])}
            </body></html>
            """
            components.html(html_content, height=800, scrolling=True)
            pd1, pd2, pd3 = st.columns(3)
            try:
                pd1.download_button("📄 PDF 수동 다운로드", data=build_pdf_from_b64(imgs), file_name=f"건영택배_{now_kst().strftime('%Y%m%d')}.pdf", mime="application/pdf", use_container_width=True)
            except Exception as e:
                pd1.caption(f"PDF 생성 실패: {e}")
            pd2.download_button("🖼️ 1페이지 JPG 다운로드", data=base64.b64decode(imgs[0]), file_name=f"건영택배_{now_kst().strftime('%Y%m%d')}.jpg", mime="image/jpeg", use_container_width=True)
            if pd3.button("✖ 미리보기 닫기", use_container_width=True):
                st.session_state['print_imgs'] = None
                st.rerun()

    with tab2:
        st.subheader("📅 과거 배송 기록 조회 및 리콜")
        c_cal, c_search = st.columns([1, 2])
        search_date = c_cal.date_input("조회 날짜", now_kst().date())
        s_date_str = search_date.strftime("%Y-%m-%d")
        s_month_str = f"건영-{search_date.year}-{search_date.month}"

        df_hist = fetch_gy_history_v2(s_month_str)
        if not df_hist.empty:
            s_query = c_search.text_input("🔍 이름/송장번호 검색")
            df_view = df_hist[df_hist['날짜'].astype(str).str.contains(s_date_str)]
            if s_query: df_view = df_view[df_view.apply(lambda r: r.astype(str).str.contains(s_query, case=False, regex=False).any(), axis=1)]
            if not df_view.empty:
                st.dataframe(df_view, use_container_width=True)
                if st.button("↩️ 리스트 작업창으로 복원", type="primary"):
                    st.session_state['deli_list'] = [
                        {'name': s(r.get('받는 곳', '')), 'phone': s(r.get('연락처', '')), 'addr': s(r.get('주소', '')), 'inv_no': s(r.get('운송장번호', '')),
                         'qty': safe_int(r.get('수량', 1), 1), 'fee': safe_int(r.get('운임', 0)), 'memo': s(r.get('비고', ''))}
                        for _, r in df_view.iterrows() if s(r.get('운송장번호', '')) not in ("TOTAL", "")]
                    st.success("복원 완료! '📝 금일 발송 작업' 탭으로 이동하세요.")
            else: st.info(f"해당 날짜({s_date_str})에 기록된 데이터가 없습니다.")
        else:
            st.info(f"선택하신 달({s_month_str})에 생성된 건영택배 탭이 아직 없습니다.")

################################################################################
# 12. 메인 메뉴 2: 일일출고관리
################################################################################
elif menu == "일일출고관리":
    col_t, col_b = st.columns([3, 1])
    col_t.title("일일출고 통합 관리")
    
    c_btn1, c_btn2, c_cal, c_o = st.columns([0.5, 0.5, 1, 2])
    c_btn1.button("어제", on_click=set_sd_date, args=(now_kst().date() - timedelta(days=1),), use_container_width=True)
    c_btn2.button("오늘", on_click=set_sd_date, args=(now_kst().date(),), use_container_width=True)
    
    sd = c_cal.date_input("날짜 선택", value=st.session_state['sd_date'], key="sd_date_input", on_change=lambda: set_sd_date(st.session_state.sd_date_input))
    tn = c_o.text_input("구글 시트 탭 이름", value=f"{sd.year}-{sd.month}").strip()
    
    if tn and not re.fullmatch(r"\d{4}-\d{1,2}", tn):
        c_o.caption("⚠️ 탭 이름이 'YYYY-M' 형식이 아닙니다. 없는 이름이면 새 탭이 만들어집니다.")

    if col_b.button("✨ 엑셀 꼬임/서식 강제 리프레쉬", use_container_width=True):
        with st.spinner("구글 시트의 유령 데이터를 지우고 디자인을 재정비 중입니다..."):
            success, msg = organize_and_format_gsheet(tn)
            clear_all_caches()
        if success: flash("success", "✅ 구글 시트 정렬 및 디자인 복구가 완료되었습니다!")
        else: flash("error", f"⛔ 에러가 발생했습니다: {msg}")
        st.rerun()

    st.divider()
    st.subheader("➕ 대량 일괄 등록 (한진택배 기본)")
    quick_df = pd.DataFrame(st.session_state['bulk_input_data'], columns=['트레킹번호', '받는사람', '택배사', '박스수량', '내용물', '메모'])
    bulk_editor = st.data_editor(
        quick_df, use_container_width=True, height=500, key=f"bulk_editor_{st.session_state['bulk_ver']}",
        column_config={
            "트레킹번호": st.column_config.TextColumn("트레킹번호", width=180),
            "받는사람": st.column_config.TextColumn("받는사람", width=140),
            "택배사": st.column_config.SelectboxColumn("택배사", options=["한진", "건영", "천일", "대신", "우리"], width=90),
            "박스수량": st.column_config.SelectboxColumn("박스수량", options=[str(i) for i in range(1, 11)], width=80),
            "내용물": st.column_config.TextColumn("내용물", width=250),
            "메모": st.column_config.TextColumn("메모", width=200)
        }
    )

    valid_mask = (bulk_editor['트레킹번호'].map(s) != '') | (bulk_editor['받는사람'].map(s) != '')
    valid_data = bulk_editor[valid_mask]
    if not valid_data.empty:
        st.markdown("### 📦 금일 작성 내용 요약표")
        summary = valid_data.copy()
        summary['택배사'] = summary['택배사'].map(lambda v: s(v) or '한진')
        summary = summary.groupby('택배사').size().reset_index(name='출고 건수 (박스수)')
        def style_summary(row): return [f"background-color: {get_hex_color(row['택배사'])}; color: black;"] * len(row)
        st.dataframe(summary.style.apply(style_summary, axis=1), hide_index=True)

    if st.button("🚀 위 리스트 구글 시트로 일괄 등록", type="primary"):
        final_bulk = []
        for _, r in valid_data.iterrows():
            final_bulk.append({'날짜': sd.strftime('%Y-%m-%d'), '트레킹번호': s(r.get('트레킹번호', '')), '받는사람': s(r.get('받는사람', '')),
                               '택배사': s(r.get('택배사', '')) or '한진', '박스수량': s(r.get('박스수량', '')) or '1',
                               '내용물': s(r.get('내용물', '')), '메모': s(r.get('메모', ''))})
        if not tn:
            st.warning("⚠️ 구글 시트 탭 이름을 입력하세요.")
        elif final_bulk:
            try:
                with st.spinner("구글 시트에 등록 중입니다..."):
                    append_to_tracking_gsheet(final_bulk, tn)
                    clear_all_caches()
                st.session_state['bulk_input_data'] = [["", "", "한진", "1", "", ""] for _ in range(20)]
                st.session_state['bulk_ver'] += 1
                flash("success", f"✅ 총 {len(final_bulk)}건 구글 시트 등록 완료!")
            except Exception:
                st.error("🚨 등록 중 오류가 발생했습니다. 입력 내용은 그대로 유지됩니다.")
                st.code(traceback.format_exc())
            else:
                st.rerun()
        else: st.warning("⚠️ 입력된 데이터가 없습니다.")

    st.divider()
    df_t = fetch_tracking_data_v2(tn) if tn else pd.DataFrame(columns=TRACK_HEADERS)
    if not df_t.empty:
        df_t = rows_to_df([df_t.columns.tolist()] + df_t.values.tolist(), TRACK_HEADERS)
        valid_mask_t = (df_t['받는사람'].map(s) != '') | (df_t['트레킹번호'].map(s) != '')
        df_t = df_t[valid_mask_t].copy()
        df_t['날짜'] = df_t['날짜'].map(norm_date).replace('', pd.NA).ffill().fillna('')
        df_t = df_t.sort_values(by=['날짜', '택배사'], ascending=[True, True], kind='stable')
        df_t.reset_index(drop=True, inplace=True)

        def style_dataframe(df):
            styles = pd.DataFrame('', index=df.index, columns=df.columns)
            prev_date = None
            for i, row in df.iterrows():
                cur_date = str(row['날짜']).strip()
                bg_color = get_hex_color(row['택배사'])
                border = "border-top: 2px solid #555;" if prev_date and cur_date != prev_date else "border-top: 1px solid #ddd;"
                styles.iloc[i, :] = f"background-color: {bg_color}; color: black; {border}"
                prev_date = cur_date
            return styles

        target_date_str = sd.strftime('%Y-%m-%d')
        st.subheader(f"📍 '{target_date_str}' 금일 발송 내역 수정")
        df_today = df_t[df_t['날짜'] == target_date_str].copy()
        df_today.reset_index(drop=True, inplace=True)

        if not df_today.empty:
            edited_today = st.data_editor(
                df_today.style.apply(lambda _: style_dataframe(df_today), axis=None),
                use_container_width=True, num_rows="dynamic", key=f"editor_today_{st.session_state['trk_ver']}",
                column_config={
                    "날짜": st.column_config.TextColumn("날짜", width=120),
                    "트레킹번호": st.column_config.TextColumn("트레킹번호", width=180),
                    "받는사람": st.column_config.TextColumn("받는사람", width=140),
                    "내용물": st.column_config.TextColumn("내용물", width=250),
                    "메모": st.column_config.TextColumn("메모", width=200)
                }
            )
            if st.button("💾 금일 수정한 내용 저장", type="primary"):
                try:
                    df_others = df_t[df_t['날짜'] != target_date_str]
                    edited_today_clean = edited_today.fillna('').copy()
                    edited_today_clean['날짜'] = edited_today_clean['날짜'].map(lambda v: norm_date(v) or target_date_str)
                    df_final = pd.concat([df_others, edited_today_clean], ignore_index=True)
                    df_final['날짜'] = df_final['날짜'].replace('', pd.NA).ffill().fillna('')
                    df_final = df_final.sort_values(by=['날짜', '택배사'], ascending=[True, True], kind='stable')
                    with st.spinner("수정된 내용을 구글 시트에 저장 중입니다..."):
                        sh = get_sh(TRACKING_SHEET_ID)
                        ws = sh.worksheet(tn)
                        safe_write(sh, ws, [TRACK_HEADERS] + df_final[TRACK_HEADERS].values.tolist())
                        organize_and_format_gsheet(tn)
                        clear_all_caches()
                    st.session_state['trk_ver'] += 1
                    flash("success", "✅ 금일 내역 수정이 구글 시트에 완료되었습니다!")
                except Exception:
                    st.error("🚨 저장 중 오류가 발생했습니다.")
                    st.code(traceback.format_exc())
                else:
                    st.rerun()
        else: st.info(f"'{target_date_str}' 날짜에 등록된 내역이 없습니다.")

        st.divider()
        st.subheader(f"📅 '{tn}' 이번 달 전체 내역 한눈에 보기 & 수정")
        edited_month = st.data_editor(
            df_t.style.apply(lambda _: style_dataframe(df_t), axis=None),
            use_container_width=True, num_rows="dynamic", height=600, key=f"editor_month_{st.session_state['trk_ver']}",
            column_config={
                "날짜": st.column_config.TextColumn("날짜", width=120),
                "트레킹번호": st.column_config.TextColumn("트레킹번호", width=180),
                "받는사람": st.column_config.TextColumn("받는사람", width=140),
                "내용물": st.column_config.TextColumn("내용물", width=250),
                "메모": st.column_config.TextColumn("메모", width=200)
            }
        )
        if st.button("💾 전체 표에서 수정한 내용 한 번에 저장"):
            try:
                edited_month_clean = edited_month.fillna('').copy()
                edited_month_clean['날짜'] = edited_month_clean['날짜'].map(norm_date).replace('', pd.NA).ffill().fillna('')
                edited_month_clean = edited_month_clean.sort_values(by=['날짜', '택배사'], ascending=[True, True], kind='stable')
                with st.spinner("전체 내용을 구글 시트에 덮어쓰고 있습니다..."):
                    sh = get_sh(TRACKING_SHEET_ID)
                    ws = sh.worksheet(tn)
                    safe_write(sh, ws, [TRACK_HEADERS] + edited_month_clean[TRACK_HEADERS].values.tolist())
                    organize_and_format_gsheet(tn)
                    clear_all_caches()
                st.session_state['trk_ver'] += 1
                flash("success", "✅ 전체 내역 덮어쓰기 및 구글 시트 디자인 복구가 완료되었습니다!")
            except Exception:
                st.error("🚨 저장 중 오류가 발생했습니다.")
                st.code(traceback.format_exc())
            else:
                st.rerun()
    else: st.info(f"'{tn}' 구글 시트에 데이터가 없습니다.")


################################################################################
# 12-2. 기존 데이터 보존형 생산관리 / 위챗 업무
################################################################################
elif menu in ("생산관리", "위챗요약"):
    try:
        import production_mgmt
        ctx = workflow_context()
        if menu == "생산관리":
            production_mgmt.render(ctx)
        else:
            production_mgmt.render_wechat(ctx)
    except ImportError as error:
        st.error("모듈을 불러오지 못했습니다. production_mgmt.py와 requirements.txt를 함께 적용하세요.")
        st.caption(str(error))
    except Exception as error:
        st.error("데이터 연결 또는 화면 처리에 문제가 있습니다. 기존 시트를 초기화하지 말고 백업·연결 설정을 확인하세요.")
        with st.expander("오류 상세"):
            st.code(traceback.format_exc())
