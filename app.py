import io
import re
from decimal import Decimal, InvalidOperation
from typing import Optional, Tuple, Dict

import pandas as pd
import streamlit as st

# Excel IO
# - openpyxl is required by pandas to read .xlsx files
# - xlrd reads .xls templates
# - xlwt writes .xls output
import xlwt
# Encrypted xlsx support
import msoffcrypto


FIXED_PASSWORD = "0000"
DELIVERY_METHODS = ["택배", "등기", "소포"]
DEFAULT_DELIVERY_METHOD = "택배"


# -----------------------------
# Helpers
# -----------------------------
def _norm(s: str) -> str:
    return re.sub(r"\s+", "", str(s or "")).strip().lower()


def find_col(df: pd.DataFrame, keywords) -> Optional[str]:
    cols = list(df.columns)
    norm_map = {_norm(c): c for c in cols}
    for kw in keywords:
        kw_n = _norm(kw)
        for n, orig in norm_map.items():
            if kw_n and kw_n in n:
                return orig
    return None


def to_plain_str(v) -> str:
    """
    Convert tracking/order numbers that may arrive as:
      - float (e.g., 3.13936e+11)
      - scientific string (e.g., "3.13936E+11")
      - int
      - string with hyphens
    into a plain digit string (or original string if non-numeric).
    """
    if v is None:
        return ""
    if isinstance(v, str):
        s = v.strip()
        if s == "" or s.lower() in {"nan", "none"}:
            return ""
        # keep hyphenated tracking numbers as-is
        if "-" in s:
            return s
        # scientific notation in string?
        if re.fullmatch(r"[+-]?\d+(\.\d+)?[eE][+-]?\d+", s):
            try:
                d = Decimal(s)
                # quantize to whole number
                return format(d.quantize(Decimal(1)), "f").split(".")[0]
            except (InvalidOperation, ValueError):
                return s
        # digits-only
        if re.fullmatch(r"\d+", s):
            return s
        # digits with .0
        if re.fullmatch(r"\d+\.0+", s):
            return s.split(".")[0]
        return s

    # numeric types
    try:
        # pandas may give numpy types
        if pd.isna(v):
            return ""
    except Exception:
        pass

    if isinstance(v, (int, )):
        return str(v)

    if isinstance(v, (float, )):
        if math.isnan(v) or math.isinf(v):
            return ""
        # Convert via Decimal using string representation to avoid binary float artifacts
        try:
            d = Decimal(str(v))
            # if looks like integer
            return format(d.quantize(Decimal(1)), "f").split(".")[0]
        except Exception:
            # fallback
            return str(int(v))

    # fallback
    return str(v).strip()


def read_encrypted_xlsx(uploaded_file, password: str) -> pd.DataFrame:
    """
    Decrypt an encrypted Excel file (xlsx) using msoffcrypto and return DataFrame.
    """
    raw = uploaded_file.read()
    office_file = msoffcrypto.OfficeFile(io.BytesIO(raw))
    office_file.load_key(password=password)
    decrypted = io.BytesIO()
    office_file.decrypt(decrypted)
    decrypted.seek(0)
    return pd.read_excel(decrypted, dtype=str)


def read_excel_any(uploaded_file) -> pd.DataFrame:
    """
    Read xlsx or xls. Prefer dtype=str to keep identifiers stable.
    """
    name = (uploaded_file.name or "").lower()
    data = uploaded_file.read()
    bio = io.BytesIO(data)
    if name.endswith(".xls"):
        # xlrd required for .xls
        return pd.read_excel(bio, dtype=str, engine="xlrd")
    return pd.read_excel(bio, dtype=str)


def make_address(df: pd.DataFrame) -> Tuple[pd.Series, Dict[str, str]]:
    """
    Build a best-effort address string from common SmartStore columns.
    Returns (address_series, debug mapping)
    """
    mapping = {}
    base = find_col(df, ["주소", "배송지", "수령자주소", "수령주소", "배송주소"])
    detail = find_col(df, ["상세주소", "주소2", "배송지상세", "상세"])
    zipc = find_col(df, ["우편번호", "우편"])

    mapping["addr_base"] = base or ""
    mapping["addr_detail"] = detail or ""
    mapping["zip"] = zipc or ""

    addr = df[base].fillna("").astype(str) if base else pd.Series([""] * len(df))
    if detail:
        addr = (addr.str.strip() + " " + df[detail].fillna("").astype(str).str.strip()).str.strip()
    if zipc:
        z = df[zipc].fillna("").astype(str).str.strip()
        # if zip exists, prefix in brackets
        addr = (z.where(z != "", "")).map(lambda x: f"[{x}] " if x else "") + addr
    return addr, mapping


def build_output(orders_df: pd.DataFrame, ship_df: pd.DataFrame, template_df: pd.DataFrame) -> pd.DataFrame:
    # ---- Identify essential columns in 1 (orders) ----
    o_order = find_col(orders_df, ["상품주문번호", "주문번호"])
    o_buyer = find_col(orders_df, ["주문자", "구매자"])
    o_recv = find_col(orders_df, ["수령자", "받는사람", "수취인"])
    addr_series, _addr_map = make_address(orders_df)

    if not o_order:
        raise ValueError("1번(스마트스토어) 파일에서 '상품주문번호' 컬럼을 찾지 못했어요.")
    if not o_buyer:
        raise ValueError("1번(스마트스토어) 파일에서 '주문자' 컬럼을 찾지 못했어요.")
    if not o_recv:
        raise ValueError("1번(스마트스토어) 파일에서 '수령자' 컬럼을 찾지 못했어요.")

    od = orders_df.copy()
    od["_order_no"] = od[o_order].map(to_plain_str)
    od["_buyer"] = od[o_buyer].fillna("").astype(str).str.strip()
    od["_recv"] = od[o_recv].fillna("").astype(str).str.strip()
    od["_addr"] = addr_series.fillna("").astype(str).str.strip()

    # ---- Identify essential columns in 2 (shipping) ----
    s_order = find_col(ship_df, ["상품주문번호", "주문번호"])
    s_track = find_col(ship_df, ["운송장번호", "송장번호", "운송장", "송장"])
    if not s_order or not s_track:
        raise ValueError("2번(운송장/출고) 파일에서 '상품주문번호' 또는 '운송장번호/송장번호' 컬럼을 찾지 못했어요.")

    sd = ship_df.copy()
    sd["_order_no"] = sd[s_order].map(to_plain_str)
    sd["_track"] = sd[s_track].map(to_plain_str)

    order_to_track: Dict[str, str] = {}
    for _, r in sd.iterrows():
        ono = (r.get("_order_no") or "").strip()
        trk = (r.get("_track") or "").strip()
        if ono and trk and ono not in order_to_track:
            order_to_track[ono] = trk

    od["_track_by_order"] = od["_order_no"].map(lambda x: order_to_track.get(x, ""))

    # ---- Group rule: same buyer/receiver/address -> same tracking number ----
    od["_group_key"] = (od["_buyer"] + "||" + od["_recv"] + "||" + od["_addr"])
    # choose first non-empty tracking within group
    group_track = (
        od.sort_values(by=["_order_no"])
          .groupby("_group_key")["_track_by_order"]
          .apply(lambda s: next((x for x in s.tolist() if x), ""))
          .to_dict()
    )
    od["_group_track"] = od["_group_key"].map(lambda k: group_track.get(k, ""))

    # ---- Prepare a lookup from order_no -> (buyer, recv, addr, group_track) ----
    lookup = (
        od.drop_duplicates(subset=["_order_no"])
          .set_index("_order_no")[["_buyer", "_recv", "_addr", "_group_track"]]
    )

    # ---- Apply to template (3) ----
    out = template_df.copy()

    t_order = find_col(out, ["상품주문번호", "주문번호"])
    if not t_order:
        raise ValueError("3번(템플릿) 파일에서 '상품주문번호' 컬럼을 찾지 못했어요.")

    out["_order_no"] = out[t_order].map(to_plain_str)
    merged = out.merge(lookup, how="left", left_on="_order_no", right_index=True)

    # Fill common fields if present
    t_buyer = find_col(out, ["주문자", "구매자"])
    t_recv = find_col(out, ["수령자", "수취인", "받는사람"])
    t_addr = find_col(out, ["주소", "배송지", "수령자주소", "배송주소"])
    t_detail = find_col(out, ["상세주소", "주소2", "배송지상세", "상세"])
    t_track = find_col(out, ["송장번호", "운송장번호", "운송장", "송장"])
    t_method = find_col(out, ["배송방법"])
    t_courier = find_col(out, ["택배사", "택배사명", "배송사", "운송사"])

    # Buyer/receiver/address
    if t_buyer:
        merged[t_buyer] = merged["_buyer"].fillna(merged.get(t_buyer))
    if t_recv:
        merged[t_recv] = merged["_recv"].fillna(merged.get(t_recv))
    if t_addr:
        # if template has separate detail, keep it; otherwise write full address into addr col
        if t_detail and t_addr:
            # best effort: split into base + detail by last space if detail empty
            full = merged["_addr"].fillna("")
            base = full
            det = ""
            # only fill base if present
            merged[t_addr] = base.where(base != "", merged.get(t_addr))
            # if detail col exists, leave it unless empty
            if t_detail:
                merged[t_detail] = merged.get(t_detail).fillna(det)
        else:
            merged[t_addr] = merged["_addr"].where(merged["_addr"].fillna("") != "", merged.get(t_addr))

    # Tracking number
    tracking = merged["_group_track"].fillna("")
    if t_track:
        merged[t_track] = tracking.where(tracking != "", merged.get(t_track))
    else:
        # if no tracking column, create one
        merged["송장번호"] = tracking

    # Delivery method (B col request)
    if t_method:
        merged[t_method] = DEFAULT_DELIVERY_METHOD
    else:
        merged["배송방법"] = DEFAULT_DELIVERY_METHOD

    # Courier (C col request)
    courier_val = tracking.map(lambda x: "컬리넥스트마일" if ("-" in str(x)) else ("롯데택배" if str(x).strip() else ""))
    if t_courier:
        merged[t_courier] = courier_val.where(courier_val != "", merged.get(t_courier))
    else:
        merged["택배사"] = courier_val

    # Clean helper columns
    for c in ["_order_no", "_buyer", "_recv", "_addr", "_group_track"]:
        if c in merged.columns:
            pass
    merged = merged.drop(columns=[c for c in merged.columns if c.startswith("_")], errors="ignore")

    return merged


def df_to_xls_bytes(df: pd.DataFrame, sheet_name: str = "발송처리") -> bytes:
    """Write DataFrame to legacy .xls and return bytes.

    NOTE: .xls has limitations (max rows 65,536). If exceeded, raise a clear error.
    """
    if len(df) > 65535:
        raise ValueError(f".xls 형식은 최대 65,536행까지 지원해요. 현재 행 수: {len(df)}")

    wb = xlwt.Workbook()
    ws = wb.add_sheet(sheet_name[:31])

    header_style = xlwt.easyxf("font: bold on; align: vert centre;")
    # default style = 'General' in Excel (no explicit number format)
    default_style = xlwt.easyxf("align: vert centre;")

    # Write header
    for j, col in enumerate(df.columns):
        ws.write(0, j, str(col), header_style)

    # Write rows
    for i, row in enumerate(df.itertuples(index=False, name=None), start=1):
        for j, v in enumerate(row):
            if v is None or (isinstance(v, float) and pd.isna(v)) or pd.isna(v):
                ws.write(i, j, "", default_style)
            else:
                # Keep as string to prevent scientific notation / precision loss
                ws.write(i, j, str(v), default_style)

    # Auto width (rough)
    for j, col in enumerate(df.columns):
        sample = df.iloc[:200, j].astype(str).fillna("").tolist()
        max_len = max([len(str(col))] + [len(x) for x in sample])
        ws.col(j).width = int(min(max(10, max_len + 2), 40) * 256)

    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


# -----------------------------
# UI
# -----------------------------
st.set_page_config(page_title="발송처리 자동 채움", layout="wide")

st.title("📦 발송처리(3번) 자동 채움")

# Section 1
st.markdown(
    "<div style='font-size:20px; font-weight:700;'>스마트스토어 엑셀(비번0000)</div>",
    unsafe_allow_html=True
)
st.write("")  # one line spacing
smartstore_file = st.file_uploader(
    label="",
    type=["xlsx"],
    accept_multiple_files=False,
    key="smartstore",
    label_visibility="collapsed",
)

st.write("")  # one line spacing

# Section 2
st.markdown(
    "<div style='font-size:20px; font-weight:700;'>운송장/출고 엑셀</div>",
    unsafe_allow_html=True
)
st.write("")  # one line spacing
shipping_file = st.file_uploader(
    label="",
    type=["xlsx", "xls"],
    accept_multiple_files=False,
    key="shipping",
    label_visibility="collapsed",
)

st.write("")  # spacing

# Template uploader (3)
st.markdown(
    "<div style='font-size:18px; font-weight:700;'>발송처리 템플릿(3번 엑셀)</div>",
    unsafe_allow_html=True
)
template_file = st.file_uploader(
    label="",
    type=["xlsx", "xls"],
    accept_multiple_files=False,
    key="template",
    label_visibility="collapsed",
)

st.write("")

run = st.button("✅ 자동 채움 실행", type="primary", use_container_width=True)

if run:
    if not smartstore_file or not shipping_file or not template_file:
        st.error("1번(스마트스토어), 2번(운송장/출고), 3번(템플릿) 파일을 모두 올려줘.")
        st.stop()

    try:
        with st.spinner("1번(암호화 엑셀) 해독 중..."):
            orders_df = read_encrypted_xlsx(smartstore_file, FIXED_PASSWORD)

        with st.spinner("2번/3번 엑셀 읽는 중..."):
            ship_df = read_excel_any(shipping_file)
            template_df = read_excel_any(template_file)

        with st.spinner("데이터 매칭 & 채우는 중..."):
            out_df = build_output(orders_df, ship_df, template_df)

        # Identify key columns in output for formatting/validation
        delivery_col = find_col(out_df, ["배송방법"]) or "배송방법"
        order_col = find_col(out_df, ["상품주문번호", "주문번호"]) or "상품주문번호"
        track_col = find_col(out_df, ["송장번호", "운송장번호", "운송장", "송장"]) or "송장번호"

        xls_bytes = df_to_xls_bytes(out_df)
# Result header: a bit larger than the ones above
        st.markdown(
            "<div style='font-size:24px; font-weight:800; margin-top:8px;'>3번 결과</div>",
            unsafe_allow_html=True
        )

        st.dataframe(out_df, use_container_width=True, hide_index=True)

        st.download_button(
            "⬇️ 3번 결과 엑셀 다운로드",
            data=xls_bytes,
            file_name="엑셀일괄발송.xls",
            mime="application/vnd.ms-excel",
            use_container_width=True,
        )

    except Exception as e:
        st.exception(e)
