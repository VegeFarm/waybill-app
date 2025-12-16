import io
import re
from decimal import Decimal, InvalidOperation
from typing import Optional, Tuple, Dict

import pandas as pd
import streamlit as st

# -------------------------
# 고정 비밀번호 (요청사항)
# -------------------------
FIXED_PASSWORD = "0000"

ROMAN_MAP = str.maketrans({
    "Ⅰ": "1", "Ⅱ": "2", "Ⅲ": "3", "Ⅳ": "4", "Ⅴ": "5",
    "Ⅵ": "6", "Ⅶ": "7", "Ⅷ": "8", "Ⅸ": "9", "Ⅹ": "10",
    "ⅰ": "1", "ⅱ": "2", "ⅲ": "3", "ⅳ": "4", "ⅴ": "5",
    "ⅵ": "6", "ⅶ": "7", "ⅷ": "8", "ⅸ": "9", "ⅹ": "10",
})


def norm_text(s) -> str:
    """공백/특수문자 제거 + 로마숫자(Ⅱ 등) 숫자로 변환."""
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    s = str(s).strip().translate(ROMAN_MAP)
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"[^0-9A-Za-z가-힣]", "", s)
    return s


def to_plain_number_str(x) -> str:
    """3.13936E+11 같은 표기를 '313936000000'처럼 보이게 변환."""
    if x is None:
        return ""
    try:
        if isinstance(x, float) and pd.isna(x):
            return ""
    except Exception:
        pass

    s = str(x).strip()
    if s == "" or s.lower() == "nan":
        return ""

    s = s.replace(",", "")
    if re.fullmatch(r"-?\d+\.0+", s):  # '123.0' 형태
        return s.split(".")[0]

    try:
        d = Decimal(s)
        if d == d.to_integral():
            return format(d.to_integral(), "f")
        plain = format(d, "f").rstrip("0").rstrip(".")
        return plain
    except (InvalidOperation, ValueError):
        return s


def to_plain_tracking_str(x) -> str:
    """운송장번호: '-' 있으면 그대로, 숫자면 과학표기 방지 변환."""
    if x is None:
        return ""
    try:
        if isinstance(x, float) and pd.isna(x):
            return ""
    except Exception:
        pass

    s = str(x).strip()
    if s == "" or s.lower() == "nan":
        return ""

    if "-" in s:
        return s
    return to_plain_number_str(s)


def decrypt_office_excel(file_bytes: bytes, password: str) -> io.BytesIO:
    """암호화된 스마트스토어 엑셀(xlsx)을 해제해서 BytesIO로 반환"""
    import msoffcrypto  # requirements.txt에 포함 필요

    decrypted = io.BytesIO()
    office_file = msoffcrypto.OfficeFile(io.BytesIO(file_bytes))
    office_file.load_key(password=password)
    office_file.decrypt(decrypted)
    decrypted.seek(0)
    return decrypted


def find_header_row(df: pd.DataFrame, must_have: Tuple[str, ...], max_scan: int = 30) -> int:
    """header=None로 읽은 df에서 컬럼명 행을 찾는다."""
    scan = min(max_scan, len(df))
    for i in range(scan):
        row = df.iloc[i].astype(str).tolist()
        if all(any(m in cell for cell in row) for m in must_have):
            return i
    return -1


def choose_tracking(series: pd.Series) -> Optional[str]:
    """같은 key에서 운송장번호가 여러 개면 최빈값(동률이면 먼저 나온 값) 선택"""
    s = series.dropna().astype(str)
    if s.empty:
        return None
    vc = s.value_counts()
    top = vc.max()
    candidates = vc[vc == top].index.tolist()
    if len(candidates) == 1:
        return candidates[0]
    for v in s:  # tie-break: 먼저 나온 값
        if v in candidates:
            return v
    return candidates[0]


def build_output(df1: pd.DataFrame, df2: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    # 1번에서 필요한 컬럼
    col_buyer = "구매자명"
    col_recv = "수취인명"
    col_addr = "통합배송지"
    col_po = "상품주문번호"

    # 2번에서 필요한 컬럼
    col2_buyer = "주문자"
    col2_recv = "수령자"
    col2_addr = "수령자 주소(상세포함)"
    col2_track = "운송장번호"

    df1 = df1.copy()
    df2 = df2.copy()

    # 주문자/수령자/주소가 같으면 같은 송장번호로 묶기 위한 key
    df1["__key"] = df1[col_buyer].map(norm_text) + "|" + df1[col_recv].map(norm_text) + "|" + df1[col_addr].map(norm_text)
    df2["__key"] = df2[col2_buyer].map(norm_text) + "|" + df2[col2_recv].map(norm_text) + "|" + df2[col2_addr].map(norm_text)

    # key → 운송장번호 매핑
    map_track: Dict[str, Optional[str]] = df2.groupby("__key")[col2_track].apply(choose_tracking).to_dict()
    df1["송장번호"] = df1["__key"].map(map_track)

    # 참고용: 같은 key에서 운송장번호가 여러 개인 경우
    dup_info = (
        df2.groupby("__key")[col2_track]
        .nunique(dropna=True)
        .reset_index(name="운송장번호_종류수")
        .query("운송장번호_종류수 > 1")
        .sort_values("운송장번호_종류수", ascending=False)
    )

    df1["_상품주문번호_plain"] = df1[col_po].apply(to_plain_number_str)
    df1["_송장번호_plain"] = df1["송장번호"].apply(to_plain_tracking_str)

    out = pd.DataFrame({
        "상품주문번호": df1["_상품주문번호_plain"],
        "배송방법": ["택배"] * len(df1),  # 기본값
        "택배사": df1["_송장번호_plain"].apply(
            lambda x: "컬리넥스트마일" if "-" in str(x) else ("롯데택배" if str(x).strip() else "")
        ),
        "송장번호": df1["_송장번호_plain"],
    })
    return out, dup_info


def export_xls(out_df: pd.DataFrame) -> bytes:
    """
    .xls 생성 (xlwt)
    - .xls는 드롭다운(DataValidation) 강제 적용이 제한적이라 B열은 값만 '택배'로 채움
    - A/D는 문자열로 써서 과학표기 방지
    """
    import xlwt

    wb = xlwt.Workbook()
    ws = wb.add_sheet("발송처리")

    header_style = xlwt.easyxf("font: bold on; align: horiz center, vert center;")
    center_style = xlwt.easyxf("align: horiz center, vert center;")
    left_style = xlwt.easyxf("align: horiz left, vert center;")

    # 컬럼 폭(대략)
    col_widths = [24, 10, 16, 32]
    for c, w in enumerate(col_widths):
        ws.col(c).width = int(w * 256)

    # 헤더
    for c, name in enumerate(out_df.columns):
        ws.write(0, c, name, header_style)

    # 데이터
    for r, row in enumerate(out_df.itertuples(index=False), start=1):
        vals = list(row)
        for c, v in enumerate(vals):
            v_str = "" if v is None else str(v)

            # A(상품주문번호), D(송장번호) → 문자열로 써서 E+11 방지
            if c in (0, 3):
                ws.write(r, c, v_str, left_style)
            # B,C는 가운데 정렬
            else:
                ws.write(r, c, v_str, center_style)

    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


# ---------------- UI ----------------
st.set_page_config(page_title="송장일괄발송", layout="wide")
st.title("📦 송장일괄발송")

st.markdown("- 1번 파일은 **비밀번호 0000 고정**으로 열어서 처리합니다.")
st.markdown("- 3번 결과는 **xls**로 다운로드됩니다.")

st.markdown("""
<style>
.upload-title { font-size: 20px; font-weight: 700; margin-bottom: 2px; }
.result-title { font-size: 22px; font-weight: 800; margin-top: 8px; }
</style>
""", unsafe_allow_html=True)

# ✅ 1) 업로드 (제목 바로 밑에 Drag&Drop)
st.markdown('<div class="upload-title">1) 스마트스토어 엑셀(비번0000)</div>', unsafe_allow_html=True)
f1 = st.file_uploader(
    label="스마트스토어 엑셀 업로드",
    type=["xlsx"],
    key="smartstore_file",
    label_visibility="collapsed",
)

# ✅ 한 칸 띄우고 2) 업로드
st.markdown("<br>", unsafe_allow_html=True)

st.markdown('<div class="upload-title">2) 운송장/출고 엑셀</div>', unsafe_allow_html=True)
f2 = st.file_uploader(
    label="운송장/출고 엑셀 업로드",
    type=["xlsx", "xls"],
    key="tracking_file",
    label_visibility="collapsed",
)

st.markdown("<br>", unsafe_allow_html=True)

run = st.button("자동 채우기", type="primary", disabled=(f1 is None or f2 is None))

if run:
    # 1번 decrypt + read
    try:
        decrypted = decrypt_office_excel(f1.read(), FIXED_PASSWORD)
        raw1 = pd.read_excel(decrypted, header=None)
    except Exception as e:
        st.error("1번 파일을 열지 못했습니다. 비밀번호(0000) 또는 파일 형식을 확인해 주세요.")
        st.exception(e)
        st.stop()

    header_idx = find_header_row(raw1, must_have=("구매자명", "수취인명", "통합배송지", "상품주문번호"))
    if header_idx < 0:
        st.error("1번 파일에서 컬럼명 행(구매자명/수취인명/통합배송지/상품주문번호)을 찾지 못했습니다.")
        st.stop()

    header = raw1.iloc[header_idx].tolist()
    df1 = raw1.iloc[header_idx + 1:].copy()
    df1.columns = header
    df1 = df1.reset_index(drop=True)

    # 2번 read
    try:
        df2 = pd.read_excel(f2)
    except Exception as e:
        st.error("2번 파일을 읽지 못했습니다.")
        st.exception(e)
        st.stop()

    need1 = {"구매자명", "수취인명", "통합배송지", "상품주문번호"}
    need2 = {"주문자", "수령자", "수령자 주소(상세포함)", "운송장번호"}
    if not need1.issubset(set(df1.columns)):
        st.error(f"1번 파일에 필요한 컬럼이 없습니다: {sorted(list(need1 - set(df1.columns)))}")
        st.stop()
    if not need2.issubset(set(df2.columns)):
        st.error(f"2번 파일에 필요한 컬럼이 없습니다: {sorted(list(need2 - set(df2.columns)))}")
        st.stop()

    out_df, dup_info = build_output(df1, df2)

    st.subheader("미리보기")
    st.dataframe(out_df.head(30), use_container_width=True)

    miss = (out_df["송장번호"].isna() | (out_df["송장번호"].astype(str).str.strip() == "")).sum()
    st.write(f"총 {len(out_df)}건 / 송장번호 누락 {miss}건")

    if not dup_info.empty:
        with st.expander("⚠️ (참고) 같은 주문자/수령자/주소인데 운송장번호가 여러 개인 경우"):
            st.dataframe(dup_info.head(50), use_container_width=True)

    st.markdown('<div class="result-title">3) 결과 다운로드</div>', unsafe_allow_html=True)

    xls_bytes = export_xls(out_df)
    st.download_button(
        "✅ 3번(발송처리) 엑셀 다운로드",
        data=xls_bytes,
        file_name="3_발송처리_자동채움.xls",
        mime="application/vnd.ms-excel",
    )
