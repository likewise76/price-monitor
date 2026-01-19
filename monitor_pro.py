import re
import html
from io import BytesIO
from urllib.parse import quote, urlsplit, urlunsplit
from datetime import datetime

import requests
import pandas as pd
import streamlit as st

# =========================================================
# [1] 기본 설정 및 스타일
# =========================================================
st.set_page_config(page_title="상품가 유통 모니터링", layout="wide")

st.markdown(
    """
<style>
/* 전체 배경 및 폰트 설정 */
.stApp { background-color: #F8F9FA; }

/* 버튼 스타일 */
div.stButton > button {
    width: 100%;
    border-radius: 6px;
    height: 45px;
    font-weight: bold;
    font-size: 16px;
    background-color: #0D6EFD;
    color: white;
    border: none;
    transition: background-color 0.3s;
}
div.stButton > button:hover {
    background-color: #0b5ed7;
    color: #ffffff;
}

/* 상단 불필요한 여백 및 UI 숨김 */
#MainMenu { visibility: hidden; }
footer { visibility: hidden; }
header { visibility: hidden; }
.block-container { padding-top: 2rem; padding-bottom: 2rem; }
</style>
""",
    unsafe_allow_html=True,
)

st.markdown(
    """
<div style="text-align:center; margin-bottom: 30px;">
  <div style="font-size: 32px; font-weight: 800; color:#333;">상품가 유통 모니터링</div>
  <div style="font-size: 14px; color:#666; margin-top: 6px;">네이버 쇼핑 검색 API 기준</div>
</div>
""",
    unsafe_allow_html=True,
)

# =========================================================
# [2] API 키 로드
# =========================================================
try:
    client_id = st.secrets["NAVER_CLIENT_ID"]
    client_secret = st.secrets["NAVER_CLIENT_SECRET"]
except Exception:
    st.error("API 키 설정이 확인되지 않습니다.")
    st.info("로컬 실행 시: .streamlit/secrets.toml 파일을 확인해주세요.")
    st.info("웹 배포 시: Streamlit Cloud의 Secrets 설정에 키를 등록해주세요.")
    st.stop()

NAVER_SHOP_URL = "https://openapi.naver.com/v1/search/shop.json"

# =========================================================
# [3] 유틸리티 함수
# =========================================================
def strip_bold_tags(s: str) -> str:
    """네이버 API 결과의 <b> 태그 제거"""
    if not isinstance(s, str):
        return ""
    s = re.sub(r"</?b>", "", s)
    return html.unescape(s).strip()

def to_int(v, default=None):
    """안전한 정수 변환"""
    try:
        return int(v)
    except Exception:
        return default

def apply_exclude_words(query: str, exclude_words: str) -> str:
    """제외 검색어 적용 로직"""
    q = (query or "").strip()
    if not exclude_words:
        return q

    clean_words = re.sub(r"[,;\t\n]+", " ", exclude_words)
    words = [w.strip() for w in clean_words.split(" ") if w.strip()]
    words = list(set(words))

    if not words:
        return q

    minus_str = " ".join([f"-{w}" for w in words])
    return f"{q} {minus_str}".strip()

def sanitize_url_for_excel(url: str) -> str:
    """엑셀 하이퍼링크 오류 방지"""
    u = (url or "").strip()
    if not u:
        return ""
    if not (u.startswith("http://") or u.startswith("https://")):
        return u

    try:
        parts = urlsplit(u)
        path = quote(parts.path, safe="/%._-~")
        query = quote(parts.query, safe="=&%._-~")
        fragment = quote(parts.fragment, safe="=%&%._-~")
        return urlunsplit((parts.scheme, parts.netloc, path, query, fragment))
    except Exception:
        return u

@st.cache_data(ttl=600)
def call_naver_shop_api(query: str, display: int, start: int, sort: str = "asc") -> dict:
    """네이버 쇼핑 API 호출 (캐싱 적용)"""
    headers = {
        "X-Naver-Client-Id": client_id,
        "X-Naver-Client-Secret": client_secret,
    }
    params = {"query": query, "display": display, "start": start, "sort": sort}

    r = requests.get(NAVER_SHOP_URL, headers=headers, params=params, timeout=10)
    if r.status_code != 200:
        raise RuntimeError(f"API 호출 실패 (HTTP {r.status_code}): {r.text[:200]}")
    return r.json()

def safe_filename(name: str) -> str:
    """파일명 특수문자 제거"""
    name = (name or "").strip()
    name = re.sub(r"[\\/:*?\"<>|]", "_", name)
    return name or "상품"

def build_excel(df_for_excel: pd.DataFrame) -> BytesIO:
    """엑셀 파일 생성 (서식 적용)"""
    output = BytesIO()
    engine = "xlsxwriter"

    with pd.ExcelWriter(output, engine=engine) as writer:
        df_for_excel.to_excel(writer, index=False, sheet_name="Sheet1")
        workbook = writer.book
        worksheet = writer.sheets["Sheet1"]

        header_fmt = workbook.add_format({"bold": True, "align": "center", "bg_color": "#D7E4BC", "border": 1})
        link_fmt = workbook.add_format({"font_color": "blue", "underline": 1})
        num_fmt = workbook.add_format({"num_format": "#,##0"})
        red_fmt = workbook.add_format({"font_color": "red", "bold": True, "num_format": "#,##0"})
        blue_fmt = workbook.add_format({"font_color": "blue", "bold": True, "num_format": "#,##0"})

        for col_num, value in enumerate(df_for_excel.columns.values):
            worksheet.write(0, col_num, value, header_fmt)

        link_col_idx = df_for_excel.columns.get_loc("판매 링크(클릭)")
        rawurl_col_idx = df_for_excel.columns.get_loc("원본 URL(복사용)")
        diff_col_idx = df_for_excel.columns.get_loc("차액")
        price_col_idx = df_for_excel.columns.get_loc("판매가")
        guide_col_idx = df_for_excel.columns.get_loc("가이드가")

        for row_num, (_, row) in enumerate(df_for_excel.iterrows(), start=1):
            raw_url = str(row.get("원본 URL(복사용)", "")).strip()
            safe_url = sanitize_url_for_excel(raw_url)

            worksheet.write_string(row_num, rawurl_col_idx, raw_url)

            if safe_url.startswith("http"):
                worksheet.write_url(row_num, link_col_idx, safe_url, link_fmt, string="바로가기")
            else:
                worksheet.write_string(row_num, link_col_idx, "링크없음", link_fmt)

            diff_val = to_int(row.get("차액"), 0) or 0
            price_fmt = red_fmt if diff_val < 0 else blue_fmt

            worksheet.write(row_num, diff_col_idx, diff_val, price_fmt)
            worksheet.write(row_num, price_col_idx, to_int(row.get("판매가"), 0) or 0, num_fmt)
            worksheet.write(row_num, guide_col_idx, to_int(row.get("가이드가"), 0) or 0, num_fmt)

        worksheet.set_column(0, 0, 15)
        worksheet.set_column(1, 1, 20)
        worksheet.set_column(2, 4, 12)
        worksheet.set_column(5, 5, 55)
        worksheet.set_column(6, 6, 12)
        worksheet.set_column(7, 7, 40)

    output.seek(0)
    return output

def show_check_points():
    """결과 0건 시 가이드"""
    st.markdown("### 검색 결과가 0건인가요?")
    st.info(
        """
        1. 정렬을 "최고가순(dsc)" 또는 "정확도순(sim)"으로 바꿔보세요.
        2. 검색할 페이지 수를 늘려보세요.
        3. 가격 필터를 잠시 끄고(체크 해제) 결과가 나오는지 먼저 확인해보세요.
        4. 제품명을 단순하게 바꿔보세요. (예: '페로리 SG15' -> '대성 페로리')
        5. 제외할 단어가 너무 많지 않은지 확인해주세요.
        """
    )

def pick_sort_for_price_band(price_filter_enabled: bool, min_price: int, max_price: int, user_sort: str) -> str:
    """
    사용자가 '정렬'을 직접 선택한 경우 그대로 사용합니다.
    다만 가격필터가 켜져 있고, 사용자가 선택한 정렬이 asc인데 최소가가 높은 편이면
    dsc로 바꿔서 뒷페이지에 몰린 고가 구간을 앞에서 더 빨리 찾게 할 수 있습니다.
    """
    # 사용자가 명시적으로 선택한 정렬을 우선 존중
    # 아래 보정은 "최저가순(asc)" 선택 시에만 가볍게 적용
    if not price_filter_enabled:
        return user_sort

    if user_sort != "asc":
        return user_sort

    # 경험적으로 min_price가 큰 경우(예: 30만원 이상) asc는 앞쪽 페이지에 저가만 잔뜩 나와 0건이 자주 발생
    if to_int(min_price, 0) >= 300000:
        return "dsc"

    return user_sort

# =========================================================
# [4] 화면 UI 구성
# =========================================================
st.subheader("검색 조건 설정")

c1, c2, c3, c4 = st.columns([2, 1, 1, 1])
with c1:
    target_name = st.text_input("제품명", value="페로리 SG15", placeholder="브랜드 + 모델명")
with c2:
    guide_price = st.number_input("가이드가(원)", value=124900, step=1000, min_value=0)

price_filter_enabled = st.checkbox("가격 필터 사용", value=True)

with c3:
    min_price = st.number_input(
        "최소 가격(원)",
        value=60000,
        step=1000,
        min_value=0,
        disabled=not price_filter_enabled,
    )
with c4:
    max_price = st.number_input(
        "최대 가격(원)",
        value=200000,
        step=1000,
        min_value=0,
        disabled=not price_filter_enabled,
    )

c5, c6, c7, c8 = st.columns([1, 1, 2, 1])
with c5:
    display = st.number_input("페이지당 개수", value=50, step=10, min_value=10, max_value=100)
with c6:
    pages = st.number_input("검색할 페이지 수", value=5, step=1, min_value=1, max_value=20)
with c7:
    exclude_words = st.text_input("제외할 단어 (선택)", value="", placeholder="공백으로 구분 (예: 중고 렌탈)")
with c8:
    sort_label = st.selectbox("정렬", ["최저가순(asc)", "최고가순(dsc)", "정확도순(sim)", "최신순(date)"], index=0)
    sort_map = {
        "최저가순(asc)": "asc",
        "최고가순(dsc)": "dsc",
        "정확도순(sim)": "sim",
        "최신순(date)": "date",
    }
    user_sort = sort_map[sort_label]

st.write("")
start_btn = st.button("모니터링 시작", type="primary")
log_placeholder = st.empty()

# =========================================================
# [5] 분석 실행 로직
# =========================================================
if start_btn:
    if price_filter_enabled and (min_price > max_price):
        log_placeholder.error("최소 가격이 최대 가격보다 클 수 없습니다.")
        st.stop()

    base_query = (target_name or "").strip()
    if not base_query:
        log_placeholder.error("제품명을 입력해주세요.")
        st.stop()

    query = apply_exclude_words(base_query, exclude_words)

    all_rows = []
    scanned_items = 0
    scanned_prices = []
    dropped_by_price = 0
    dropped_by_invalid_price = 0

    # 정렬 보정(필요 시)
    sort = pick_sort_for_price_band(price_filter_enabled, int(min_price), int(max_price), user_sort)

    try:
        log_placeholder.info(f"'{base_query}' 검색 중... (API 연결)")

        for i in range(int(pages)):
            start = 1 + i * int(display)
            if start > 1000:
                break

            data = call_naver_shop_api(query=query, display=int(display), start=int(start), sort=sort)
            items = data.get("items", [])
            scanned_items += len(items)

            for it in items:
                lprice = to_int(it.get("lprice"), 0)
                if lprice is None or lprice <= 0:
                    dropped_by_invalid_price += 1
                    continue

                scanned_prices.append(int(lprice))

                if price_filter_enabled:
                    if not (int(min_price) <= int(lprice) <= int(max_price)):
                        dropped_by_price += 1
                        continue

                title = strip_bold_tags(it.get("title", ""))
                link = (it.get("link", "") or "").strip()
                mall = (it.get("mallName", "") or "").strip() or "판매처미상"

                diff = int(lprice) - int(guide_price)

                status = "정상"
                if diff < 0:
                    status = "가이드가 미준수"
                elif diff > 0:
                    status = "고가"

                all_rows.append(
                    {
                        "상태": status,
                        "판매처": mall,
                        "판매가": int(lprice),
                        "가이드가": int(guide_price),
                        "차액": int(diff),
                        "제품명": title,
                        "링크": link,
                    }
                )

        # 진단 패널(0건 원인 확인용)
        with st.expander("진단 정보(0건 원인 확인)", expanded=False):
            st.write(f"검색어(적용 후): {query}")
            st.write(f"정렬: {sort} (선택: {user_sort})")
            st.write(f"스캔 결과 수: {scanned_items}개")
            if scanned_prices:
                st.write(f"스캔된 가격대: {min(scanned_prices):,}원 ~ {max(scanned_prices):,}원")
            else:
                st.write("스캔된 가격 데이터가 없습니다(lprice 누락/0원 처리 가능).")
            st.write(f"가격정보 누락/0원 제외: {dropped_by_invalid_price}개")
            if price_filter_enabled:
                st.write(f"가격필터: ON ({int(min_price):,}원 ~ {int(max_price):,}원)")
                st.write(f"가격필터로 제외된 항목: {dropped_by_price}개")
            else:
                st.write("가격필터: OFF")

        if not all_rows:
            log_placeholder.warning("조건에 맞는 상품이 없습니다.")
            with st.expander("점검 가이드 보기", expanded=True):
                show_check_points()

            # 사용자 체감상 '가격범위를 삭제해야 될 정도' 원인을 명확히 보이기 위한 추가 메시지
            if price_filter_enabled and scanned_prices:
                if int(max_price) < min(scanned_prices) or int(min_price) > max(scanned_prices):
                    st.info("입력하신 가격범위가 이번에 스캔된 가격대와 겹치지 않습니다. 정렬/페이지를 바꾸거나 가격필터를 잠시 끄고 범위를 재설정해보시기 바랍니다.")
            st.stop()

        df = pd.DataFrame(all_rows)
        df = df.drop_duplicates(subset=["링크"])
        df = df.sort_values(by="판매가", ascending=True).reset_index(drop=True)

        log_placeholder.success("분석이 완료되었습니다.")

        # 결과 리포트
        st.markdown("### 분석 결과 리포트")
        st.info(f"총 {scanned_items}개 검색 결과 중 유효 상품 {len(df)}개 발견")

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("검색결과", f"{len(df)}개")
        m2.metric("현재 최저가", f"{df['판매가'].min():,}원")
        m3.metric(
            "최저가 차액",
            f"{df['판매가'].min() - int(guide_price):,}원",
            delta_color="off" if df["판매가"].min() >= int(guide_price) else "inverse",
        )
        m4.metric("미준수 건수", f"{len(df[df['차액'] < 0])}개", delta_color="inverse")

        # Top 5 미준수
        viol = df[df["차액"] < 0].copy()
        if len(viol) > 0:
            st.markdown("#### 미준수 판매채널 TOP 5")
            agg = (
                viol.groupby("판매처", dropna=False)
                .agg(적발건수=("링크", "count"), 최저가=("판매가", "min"), 최저차액=("차액", "min"))
                .reset_index()
                .sort_values(by=["적발건수", "최저차액"], ascending=[False, True])
                .head(5)
            )
            st.dataframe(
                agg,
                column_config={
                    "최저가": st.column_config.NumberColumn(format="%d원"),
                    "최저차액": st.column_config.NumberColumn(format="%d원"),
                },
                use_container_width=True,
                hide_index=True,
            )

        # 상세 리스트
        st.markdown("### 상세 모니터링 리스트")
        df_display = df.copy()
        df_display["판매 링크"] = df_display["링크"]

        st.dataframe(
            df_display[["상태", "판매처", "판매가", "가이드가", "차액", "제품명", "판매 링크"]],
            column_config={
                "판매 링크": st.column_config.LinkColumn("바로가기", display_text="링크이동"),
                "판매가": st.column_config.NumberColumn(format="%d원"),
                "가이드가": st.column_config.NumberColumn(format="%d원"),
                "차액": st.column_config.NumberColumn(format="%d원"),
                "상태": st.column_config.TextColumn("상태"),
            },
            use_container_width=True,
            height=600,
            hide_index=True,
        )

        # 엑셀 다운로드
        df_for_excel = df.copy()
        df_for_excel.insert(df_for_excel.columns.get_loc("링크") + 1, "원본 URL(복사용)", df_for_excel["링크"])
        df_for_excel.insert(df_for_excel.columns.get_loc("링크") + 1, "판매 링크(클릭)", "바로가기")
        df_for_excel = df_for_excel.drop(columns=["링크"])
        df_for_excel = df_for_excel[
            ["상태", "판매처", "판매가", "가이드가", "차액", "제품명", "판매 링크(클릭)", "원본 URL(복사용)"]
        ]

        output = build_excel(df_for_excel)

        today_str = datetime.now().strftime("%Y%m%d")
        safe_query = safe_filename(base_query)
        file_name = f"모니터링_{safe_query}_{today_str}.xlsx"

        st.download_button(
            label="엑셀 리포트 다운로드",
            data=output,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

        st.caption("엑셀에서 '바로가기' 클릭이 안 될 경우 '원본 URL(복사용)'을 이용해주세요.")

    except Exception as e:
        log_placeholder.error(f"오류 발생: {e}")
        with st.expander("상세 오류 보기"):
            st.write(e)
