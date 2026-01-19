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
.stApp { background-color: #F8F9FA; }
div.stButton > button {
    width: 100%;
    border-radius: 6px;
    height: 45px;
    font-weight: bold;
    font-size: 16px;
    background-color: #0D6EFD;
    color: white;
    border: none;
}
div.stButton > button:hover { background-color: #0b5ed7; color: #ffffff; }
#MainMenu { visibility: hidden; }
footer { visibility: hidden; }
header { visibility: hidden; }
.block-container { padding-top: 1.6rem; padding-bottom: 2rem; }
</style>
""",
    unsafe_allow_html=True,
)

st.markdown(
    """
<div style="text-align:center; margin-bottom: 18px;">
  <div style="font-size: 30px; font-weight: 800; color:#333;">상품가 유통 모니터링</div>
  <div style="font-size: 13px; color:#666; margin-top: 6px;">네이버 쇼핑 검색 API 기준</div>
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
    if not isinstance(s, str):
        return ""
    s = re.sub(r"</?b>", "", s)
    return html.unescape(s).strip()

def to_int(v, default=None):
    try:
        return int(v)
    except Exception:
        return default

def normalize_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())

def parse_space_words(s: str) -> list:
    s = normalize_spaces(s)
    if not s:
        return []
    return [w.strip() for w in s.split(" ") if w.strip()]

def title_contains_all(title: str, must_words: str) -> bool:
    words = parse_space_words(must_words)
    if not words:
        return True
    t = (strip_bold_tags(title or "")).upper()
    return all(w.upper() in t for w in words)

def title_contains_any(title: str, any_words: str) -> bool:
    words = parse_space_words(any_words)
    if not words:
        return False
    t = (strip_bold_tags(title or "")).upper()
    return any(w.upper() in t for w in words)

def sanitize_url_for_excel(url: str) -> str:
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

def safe_filename(name: str) -> str:
    name = (name or "").strip()
    name = re.sub(r"[\\/:*?\"<>|]", "_", name)
    return name or "상품"

def extract_matched_terms_from_raw_title(raw_title: str) -> str:
    """
    네이버 쇼핑 API title의 <b>...</b> 구간을 추출해 매칭키워드로 사용
    """
    if not raw_title:
        return ""
    hits = re.findall(r"<b>(.*?)</b>", raw_title, flags=re.IGNORECASE)
    hits = [strip_bold_tags(h) for h in hits if h]
    uniq = []
    for h in hits:
        if h and h not in uniq:
            uniq.append(h)
    return ", ".join(uniq[:8])

@st.cache_data(ttl=600)
def call_naver_shop_api(query: str, display: int, start: int, sort: str, exclude: str = "") -> dict:
    headers = {
        "X-Naver-Client-Id": client_id,
        "X-Naver-Client-Secret": client_secret,
    }
    params = {"query": query, "display": display, "start": start, "sort": sort}
    if exclude:
        params["exclude"] = exclude
    r = requests.get(NAVER_SHOP_URL, headers=headers, params=params, timeout=10)
    if r.status_code != 200:
        raise RuntimeError(f"API 호출 실패 (HTTP {r.status_code}): {r.text[:200]}")
    return r.json()

def build_excel(df_for_excel: pd.DataFrame) -> BytesIO:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_for_excel.to_excel(writer, index=False, sheet_name="Sheet1")
        workbook = writer.book
        ws = writer.sheets["Sheet1"]

        header_fmt = workbook.add_format({"bold": True, "align": "center", "bg_color": "#D7E4BC", "border": 1})
        link_fmt = workbook.add_format({"font_color": "blue", "underline": 1})
        num_fmt = workbook.add_format({"num_format": "#,##0"})
        red_fmt = workbook.add_format({"font_color": "red", "bold": True, "num_format": "#,##0"})
        blue_fmt = workbook.add_format({"font_color": "blue", "bold": True, "num_format": "#,##0"})

        for col_num, value in enumerate(df_for_excel.columns.values):
            ws.write(0, col_num, value, header_fmt)

        # 링크 컬럼이 있는 경우에만 서식 적용
        if "판매 링크(클릭)" in df_for_excel.columns and "원본 URL(복사용)" in df_for_excel.columns:
            link_col_idx = df_for_excel.columns.get_loc("판매 링크(클릭)")
            rawurl_col_idx = df_for_excel.columns.get_loc("원본 URL(복사용)")
            diff_col_idx = df_for_excel.columns.get_loc("차액") if "차액" in df_for_excel.columns else None
            price_col_idx = df_for_excel.columns.get_loc("판매가") if "판매가" in df_for_excel.columns else None
            guide_col_idx = df_for_excel.columns.get_loc("가이드가") if "가이드가" in df_for_excel.columns else None

            for row_num, (_, row) in enumerate(df_for_excel.iterrows(), start=1):
                raw_url = str(row.get("원본 URL(복사용)", "")).strip()
                safe_url = sanitize_url_for_excel(raw_url)

                ws.write_string(row_num, rawurl_col_idx, raw_url)

                if safe_url.startswith("http"):
                    ws.write_url(row_num, link_col_idx, safe_url, link_fmt, string="바로가기")
                else:
                    ws.write_string(row_num, link_col_idx, "링크없음", link_fmt)

                if diff_col_idx is not None:
                    diff_val = to_int(row.get("차액"), 0) or 0
                    ws.write(row_num, diff_col_idx, diff_val, red_fmt if diff_val < 0 else blue_fmt)
                if price_col_idx is not None:
                    ws.write(row_num, price_col_idx, to_int(row.get("판매가"), 0) or 0, num_fmt)
                if guide_col_idx is not None:
                    ws.write(row_num, guide_col_idx, to_int(row.get("가이드가"), 0) or 0, num_fmt)

        ws.set_column(0, min(10, len(df_for_excel.columns)-1), 18)
    output.seek(0)
    return output

# =========================================================
# [4] UI (단순화)
# =========================================================
st.subheader("검색")

c1, c2, c3 = st.columns([2, 1, 1])
with c1:
    query = st.text_input("검색어", value="페로리 SG15", placeholder="예: 보일러 DOC / 페로리 SG15")
with c2:
    guide_price = st.number_input("가이드가(원)", value=124900, step=1000, min_value=0)
with c3:
    pages = st.number_input("페이지 수", value=5, step=1, min_value=1, max_value=20)

c4, c5, c6 = st.columns([1, 1, 1])
with c4:
    price_filter_enabled = st.checkbox("가격 필터", value=True)
with c5:
    min_price = st.number_input("최소(원)", value=0, step=1000, min_value=0, disabled=not price_filter_enabled)
with c6:
    max_price = st.number_input("최대(원)", value=2000000, step=1000, min_value=0, disabled=not price_filter_enabled)

st.write("")
st.subheader("상품군 섞임 방지(선택)")

d1, d2 = st.columns([1, 1])
with d1:
    must_in_title = st.text_input("제목 필수 포함 단어", value="", placeholder="예: 페로리 온수기 대성쎌틱")
with d2:
    ban_in_title = st.text_input("제목 제외 단어", value="", placeholder="예: 기타 앰프")

with st.expander("고급 설정", expanded=False):
    a1, a2, a3 = st.columns([1, 1, 2])
    with a1:
        display = st.number_input("페이지당 개수", value=50, step=10, min_value=10, max_value=100)
    with a2:
        sort_label = st.selectbox("정렬", ["sim(정확도)", "asc(저가)", "dsc(고가)", "date(최신)"], index=0)
        sort_map = {"sim(정확도)": "sim", "asc(저가)": "asc", "dsc(고가)": "dsc", "date(최신)": "date"}
        sort = sort_map[sort_label]
    with a3:
        exclude_opts = st.multiselect(
            "API exclude",
            options=["used(중고)", "rental(렌탈)", "cbshop(해외직구/구매대행)"],
            default=["used(중고)", "rental(렌탈)"],
        )

start_btn = st.button("모니터링 시작", type="primary")
log_placeholder = st.empty()

# =========================================================
# [5] 실행
# =========================================================
if start_btn:
    q = normalize_spaces(query)
    if not q:
        log_placeholder.error("검색어를 입력해주세요.")
        st.stop()
    if price_filter_enabled and min_price > max_price:
        log_placeholder.error("최소 가격이 최대 가격보다 클 수 없습니다.")
        st.stop()

    exclude_map = {"used(중고)": "used", "rental(렌탈)": "rental", "cbshop(해외직구/구매대행)": "cbshop"}
    exclude_val = ":".join([exclude_map[x] for x in exclude_opts if x in exclude_map])

    all_rows = []
    scanned_items = 0
    debug_raw_titles = []

    try:
        log_placeholder.info("검색 중입니다...")

        for i in range(int(pages)):
            start = 1 + i * int(display)
            if start > 1000:
                break  # start 최대 1000 :contentReference[oaicite:2]{index=2}

            data = call_naver_shop_api(query=q, display=int(display), start=int(start), sort=sort, exclude=exclude_val)
            items = data.get("items", [])
            scanned_items += len(items)

            for it in items:
                lprice = to_int(it.get("lprice"), 0)
                if not lprice or lprice <= 0:
                    continue

                if price_filter_enabled:
                    if not (int(min_price) <= int(lprice) <= int(max_price)):
                        continue

                raw_title = it.get("title", "") or ""
                title = strip_bold_tags(raw_title)
                if len(debug_raw_titles) < 5 and raw_title:
                    debug_raw_titles.append(raw_title)

                # 제목 기반 1차 정리(상품군 섞임 방지)
                if must_in_title and (not title_contains_all(title, must_in_title)):
                    continue
                if ban_in_title and title_contains_any(title, ban_in_title):
                    continue

                matched_terms = extract_matched_terms_from_raw_title(raw_title)
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
                        "매칭키워드": matched_terms,
                        "링크": link,
                    }
                )

        if not all_rows:
            log_placeholder.warning("조건에 맞는 상품이 없습니다.")
            with st.expander("진단(왜 0건인지 확인)", expanded=True):
                st.write(f"스캔 수: {scanned_items}개")
                st.write(f"정렬: {sort}, exclude: {exclude_val if exclude_val else '(미사용)'}")
                st.write("원본 title 예시(<b> 매칭 확인용):")
                for t in debug_raw_titles:
                    st.write(t)
            st.stop()

        df = pd.DataFrame(all_rows).drop_duplicates(subset=["링크"]).sort_values("판매가").reset_index(drop=True)

        # =========================================================
        # [6] 매칭키워드 선택(체크) 필터
        # =========================================================
        # 매칭키워드 후보 목록 만들기
        terms = []
        for s in df["매칭키워드"].fillna(""):
            parts = [p.strip() for p in str(s).split(",") if p.strip()]
            for p in parts:
                if p not in terms:
                    terms.append(p)

        st.markdown("### 결과")
        st.info(f"스캔 {scanned_items}개 중 유효 {len(df)}개")

        # 매칭키워드가 있을 때만 노출
        df_view = df.copy()
        if terms:
            st.markdown("### 매칭키워드 선택(필터)")
            selected_terms = st.multiselect(
                "표시할 매칭키워드(복수 선택 가능)",
                options=terms,
                default=terms,
            )
            apply_term_filter = st.checkbox("선택한 매칭키워드로 필터 적용", value=True)

            if apply_term_filter and selected_terms:
                def hit_selected(s: str) -> bool:
                    s = str(s or "")
                    return any(t in s for t in selected_terms)

                df_view = df_view[df_view["매칭키워드"].apply(hit_selected)].copy()

        # 요약 지표는 df_view 기준으로 보여드리는 편이 직관적입니다.
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("표시 상품", f"{len(df_view)}개")
        m2.metric("현재 최저가", f"{df_view['판매가'].min():,}원" if len(df_view) else "-")
        m3.metric("미준수 건수", f"{len(df_view[df_view['차액'] < 0])}개" if len(df_view) else "0개", delta_color="inverse")
        m4.metric("정렬", sort)

        st.markdown("### 상세 모니터링 리스트")
        df_display = df_view.copy()
        df_display["판매 링크"] = df_display["링크"]

        st.dataframe(
            df_display[["상태", "판매처", "판매가", "가이드가", "차액", "제품명", "매칭키워드", "판매 링크"]],
            column_config={
                "판매 링크": st.column_config.LinkColumn("바로가기", display_text="링크이동"),
                "판매가": st.column_config.NumberColumn(format="%d원"),
                "가이드가": st.column_config.NumberColumn(format="%d원"),
                "차액": st.column_config.NumberColumn(format="%d원"),
            },
            use_container_width=True,
            height=600,
            hide_index=True,
        )

        # =========================================================
        # [7] 엑셀 다운로드(df_view 기준)
        # =========================================================
        df_for_excel = df_view.copy()
        df_for_excel.insert(df_for_excel.columns.get_loc("링크") + 1, "원본 URL(복사용)", df_for_excel["링크"])
        df_for_excel.insert(df_for_excel.columns.get_loc("링크") + 1, "판매 링크(클릭)", "바로가기")
        df_for_excel = df_for_excel.drop(columns=["링크"])
        df_for_excel = df_for_excel[
            ["상태", "판매처", "판매가", "가이드가", "차액", "제품명", "매칭키워드", "판매 링크(클릭)", "원본 URL(복사용)"]
        ]

        output = build_excel(df_for_excel)
        today_str = datetime.now().strftime("%Y%m%d")
        file_name = f"모니터링_{safe_filename(q)}_{today_str}.xlsx"

        st.download_button(
            label="엑셀 리포트 다운로드",
            data=output,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

        with st.expander("진단(매칭 확인)", expanded=False):
            st.write(f"정렬: {sort}, exclude: {exclude_val if exclude_val else '(미사용)'}")
            st.write("원본 title 예시(<b> 매칭 확인용):")
            for t in debug_raw_titles:
                st.write(t)

        log_placeholder.success("완료되었습니다.")

    except Exception as e:
        log_placeholder.error(f"오류 발생: {e}")
        with st.expander("상세 오류 보기"):
            st.write(e)
