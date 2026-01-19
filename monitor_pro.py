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
    transition: background-color 0.3s;
}
div.stButton > button:hover {
    background-color: #0b5ed7;
    color: #ffffff;
}

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
<div style="text-align:center; margin-bottom: 26px;">
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
    if not isinstance(s, str):
        return ""
    s = re.sub(r"</?b>", "", s)
    return html.unescape(s).strip()

def to_int(v, default=None):
    try:
        return int(v)
    except Exception:
        return default

def apply_exclude_words(query: str, exclude_words: str) -> str:
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

def build_excel(df_for_excel: pd.DataFrame) -> BytesIO:
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

        worksheet.set_column(0, 0, 20)   # 상태
        worksheet.set_column(1, 1, 18)   # 제목 플래그
        worksheet.set_column(2, 2, 35)   # 제목 진단
        worksheet.set_column(3, 3, 18)   # 판매처
        worksheet.set_column(4, 6, 12)   # 가격/가이드/차액
        worksheet.set_column(7, 7, 55)   # 제품명
        worksheet.set_column(8, 8, 16)   # 매칭키워드
        worksheet.set_column(9, 9, 12)   # 바로가기
        worksheet.set_column(10, 10, 40) # 원본 URL

    output.seek(0)
    return output

def show_check_points():
    st.markdown("### 검색 결과가 0건인가요?")
    st.info(
        """
        정렬을 sim(정확도순)로 포함한 2회전 수집을 켜보세요.
        검색할 페이지 수를 늘려보세요.
        가격 필터를 잠시 끄고 결과가 나오는지 먼저 확인해보세요.
        검색어 변형(자동/수동)을 사용해보세요.
        """
    )

def extract_matched_terms_from_raw_title(raw_title: str) -> str:
    if not raw_title:
        return ""
    hits = re.findall(r"<b>(.*?)</b>", raw_title, flags=re.IGNORECASE)
    hits = [strip_bold_tags(h) for h in hits if h]
    uniq = []
    for h in hits:
        if h and h not in uniq:
            uniq.append(h)
    return ", ".join(uniq[:6])

def normalize_for_match(s: str) -> str:
    s = (s or "")
    s = strip_bold_tags(s)
    s = s.upper()
    s = re.sub(r"[\s\-_]+", "", s)
    return s

def extract_model_code_from_query(q: str):
    if not q:
        return None
    nq = normalize_for_match(q)
    m = re.search(r"([A-Z]{1,6})(\d{1,4})", nq)
    if not m:
        return None
    prefix = m.group(1)
    num = m.group(2)
    full = f"{prefix}{num}"
    return prefix, num, full

def diagnose_title(title: str, base_query: str, bad_words: str) -> str:
    t_raw = strip_bold_tags(title or "")
    t_up = (t_raw or "").upper()
    reasons = []

    if bad_words:
        clean = re.sub(r"[,;\t\n]+", " ", bad_words)
        words = [w.strip() for w in clean.split(" ") if w.strip()]
        hit = [w for w in words if w and (w.upper() in t_up)]
        if hit:
            reasons.append(f"위험단어:{'/'.join(hit[:3])}" + ("…" if len(hit) > 3 else ""))

    model = extract_model_code_from_query(base_query)
    if model:
        prefix, num, full = model
        nt = normalize_for_match(t_raw)

        if full not in nt:
            reasons.append("모델불일치(목표코드없음)")

        others = set(re.findall(rf"{re.escape(prefix)}(\d{{1,4}})", nt))
        if num in others:
            others.discard(num)
        if others:
            others_list = sorted(list(others))[:3]
            reasons.append(f"다른모델동시표기:{prefix}{'/'.join(others_list)}" + ("…" if len(others) > 3 else ""))

    return " | ".join(reasons)

def generate_query_variations(base_query: str, enable: bool, max_vars: int) -> list:
    b = (base_query or "").strip()
    if not b:
        return []
    if not enable:
        return [b]

    vars_ = []
    def add(q):
        q = (q or "").strip()
        if q and q not in vars_:
            vars_.append(q)

    add(b)

    # 공백 정리
    add(re.sub(r"\s+", " ", b))

    # 하이픈/언더스코어 제거 버전
    add(re.sub(r"[-_]+", " ", b))

    # 모델코드가 있으면, 코드 단독/브랜드+코드 형태도 후보로 추가(너무 노이즈면 UI에서 max_vars로 제한)
    model = extract_model_code_from_query(b)
    if model:
        _, _, full = model
        add(full)

    return vars_[: max(1, int(max_vars))]

def pick_second_sort(price_filter_enabled: bool, min_price: int, max_price: int) -> str:
    if not price_filter_enabled:
        return "asc"
    # 고가 구간을 보려면 dsc가 빠르게 잡힘, 저가 구간은 asc가 빠름
    if to_int(min_price, 0) >= 300000:
        return "dsc"
    return "asc"

@st.cache_data(ttl=600)
def call_naver_shop_api(query: str, display: int, start: int, sort: str, exclude: str = "") -> dict:
    headers = {
        "X-Naver-Client-Id": client_id,
        "X-Naver-Client-Secret": client_secret,
    }
    params = {
        "query": query,
        "display": display,
        "start": start,
        "sort": sort,
    }
    if exclude:
        params["exclude"] = exclude

    r = requests.get(NAVER_SHOP_URL, headers=headers, params=params, timeout=10)
    if r.status_code != 200:
        raise RuntimeError(f"API 호출 실패 (HTTP {r.status_code}): {r.text[:200]}")
    return r.json()

# =========================================================
# [4] 화면 UI 구성
# =========================================================
st.subheader("검색 조건 설정")

c1, c2, c3, c4 = st.columns([2, 1, 1, 1])
with c1:
    target_name = st.text_input("제품명(대표 검색어)", value="페로리 SG15", placeholder="브랜드 + 모델명")
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
    exclude_words = st.text_input("제외 검색어(선택, query에 -단어로 반영)", value="", placeholder="공백으로 구분 (예: 중고 렌탈)")
with c8:
    auto_vars = st.checkbox("검색어 변형 자동 적용", value=True)
    max_vars = st.number_input("변형 개수", value=2, step=1, min_value=1, max_value=6)

st.write("")
st.subheader("수집 전략")

s1, s2, s3, s4 = st.columns([1, 1, 1, 1])
with s1:
    use_two_pass = st.checkbox("2회전 수집(sim + 가격순)", value=True)
with s2:
    user_primary_sort = st.selectbox("1차 정렬", ["sim(정확도)", "date(최신)", "asc(저가)", "dsc(고가)"], index=0)
    sort_map = {"sim(정확도)": "sim", "date(최신)": "date", "asc(저가)": "asc", "dsc(고가)": "dsc"}
    primary_sort = sort_map[user_primary_sort]
with s3:
    exclude_opts = st.multiselect(
        "API exclude(1차 제외)",
        options=["used(중고)", "rental(렌탈)", "cbshop(해외직구/구매대행)"],
        default=["used(중고)", "rental(렌탈)"],
    )
with s4:
    manual_variations = st.text_area(
        "추가 검색어(선택, 줄바꿈으로 여러 개)",
        value="",
        height=90,
        placeholder="예: 대성쎌틱 DOC-20K\nDOC-20K",
    )

st.write("")
st.subheader("제목(상품명) 진단/표시(제외하지 않음)")

t1, t2, t3 = st.columns([2, 1, 1])
with t1:
    title_bad_words = st.text_input(
        "제목 위험 단어(표시용)",
        value="중고 리퍼 렌탈 반품 전시 병행수입 해외직구 구매대행 정품아님 호환",
        placeholder="공백으로 구분",
    )
with t2:
    hide_title_warning = st.checkbox("상세 리스트에서 제목주의 숨기기", value=False)
with t3:
    agg_exclude_title_warning = st.checkbox("TOP5 집계에서 제목주의 제외", value=False)

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

    # exclude 조합
    exclude_map = {"used(중고)": "used", "rental(렌탈)": "rental", "cbshop(해외직구/구매대행)": "cbshop"}
    exclude_val = ":".join([exclude_map[x] for x in exclude_opts if x in exclude_map])

    # 검색어 구성(대표 + 자동 변형 + 수동 추가)
    queries = []
    for q in generate_query_variations(base_query, enable=auto_vars, max_vars=int(max_vars)):
        queries.append(q)

    if manual_variations.strip():
        for line in manual_variations.splitlines():
            line = line.strip()
            if line:
                queries.append(line)

    # query에 -단어 반영(사용자 제외검색어)
    queries = [apply_exclude_words(q, exclude_words) for q in queries]
    # 중복 제거
    uniq_queries = []
    for q in queries:
        if q and q not in uniq_queries:
            uniq_queries.append(q)
    queries = uniq_queries

    # 정렬 구성
    sorts = [primary_sort]
    if use_two_pass:
        second_sort = pick_second_sort(price_filter_enabled, int(min_price), int(max_price))
        if second_sort not in sorts:
            sorts.append(second_sort)

    all_rows = []
    scanned_items = 0
    scanned_prices = []
    dropped_by_price = 0
    dropped_by_invalid_price = 0
    per_query_stats = []
    debug_titles = []

    try:
        log_placeholder.info(f"'{base_query}' 검색 중... (API 연결)")

        for q in queries:
            for sort in sorts:
                local_scanned = 0
                local_kept = 0

                for i in range(int(pages)):
                    start = 1 + i * int(display)
                    if start > 1000:
                        break

                    data = call_naver_shop_api(query=q, display=int(display), start=int(start), sort=sort, exclude=exclude_val)
                    items = data.get("items", [])
                    scanned_items += len(items)
                    local_scanned += len(items)

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

                        raw_title = it.get("title", "") or ""
                        title = strip_bold_tags(raw_title)
                        if len(debug_titles) < 10 and raw_title:
                            debug_titles.append(raw_title)

                        matched_terms = extract_matched_terms_from_raw_title(raw_title)
                        title_issue = diagnose_title(title=title, base_query=base_query, bad_words=title_bad_words)
                        title_flag = "정상" if not title_issue else "제목주의"

                        link = (it.get("link", "") or "").strip()
                        mall = (it.get("mallName", "") or "").strip() or "판매처미상"

                        diff = int(lprice) - int(guide_price)
                        status = "정상"
                        if diff < 0:
                            status = "가이드가 미준수"
                        elif diff > 0:
                            status = "고가"

                        if title_flag == "제목주의":
                            status = f"{status} / 제목주의"

                        all_rows.append(
                            {
                                "상태": status,
                                "제목 플래그": title_flag,
                                "제목 진단": title_issue,
                                "판매처": mall,
                                "판매가": int(lprice),
                                "가이드가": int(guide_price),
                                "차액": int(diff),
                                "제품명": title,
                                "매칭키워드": matched_terms,
                                "링크": link,
                                "검색어": q,
                                "정렬": sort,
                            }
                        )
                        local_kept += 1

                per_query_stats.append({"검색어": q, "정렬": sort, "스캔": local_scanned, "추출": local_kept})

        with st.expander("진단 정보(0건/오탐 원인 확인)", expanded=False):
            st.write(f"API exclude: {exclude_val if exclude_val else '(미사용)'}")
            st.write(f"검색어 개수: {len(queries)}개, 정렬: {', '.join(sorts)}")
            st.write(f"총 스캔 결과 수: {scanned_items}개")
            if scanned_prices:
                st.write(f"스캔된 가격대: {min(scanned_prices):,}원 ~ {max(scanned_prices):,}원")
            else:
                st.write("스캔된 가격 데이터가 없습니다(lprice 누락/0원 처리 가능).")
            st.write(f"가격정보 누락/0원 제외: {dropped_by_invalid_price}개")
            if price_filter_enabled:
                st.write(f"가격필터: ON ({int(min_price):,}원 ~ {int(max_price):,}원), 필터로 제외: {dropped_by_price}개")
            else:
                st.write("가격필터: OFF")
            if debug_titles:
                st.write("원본 title 예시(<b> 매칭 확인용, 최대 10개):")
                for t in debug_titles[:10]:
                    st.write(t)

            if per_query_stats:
                st.write("검색어/정렬별 스캔/추출 요약:")
                st.dataframe(pd.DataFrame(per_query_stats), use_container_width=True, hide_index=True)

        if not all_rows:
            log_placeholder.warning("조건에 맞는 상품이 없습니다.")
            with st.expander("점검 가이드 보기", expanded=True):
                show_check_points()
            st.stop()

        df = pd.DataFrame(all_rows)

        # 링크 기준 중복 제거(다중 검색어/정렬로 수집되므로)
        if "링크" in df.columns:
            df = df.drop_duplicates(subset=["링크"])

        df = df.sort_values(by="판매가", ascending=True).reset_index(drop=True)

        log_placeholder.success("분석이 완료되었습니다.")

        st.markdown("### 분석 결과 리포트")
        st.info(f"총 {scanned_items}개 스캔 결과 중 유효 상품 {len(df)}개 발견")

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("유효 상품", f"{len(df)}개")
        m2.metric("현재 최저가", f"{df['판매가'].min():,}원")
        m3.metric("미준수 건수", f"{len(df[df['차액'] < 0])}개", delta_color="inverse")
        m4.metric("제목주의 건수", f"{len(df[df['제목 플래그'] == '제목주의'])}개")
        m5.metric("스캔/정렬/변형", f"{len(queries)}x{len(sorts)}")

        viol = df[df["차액"] < 0].copy()
        if agg_exclude_title_warning and "제목 플래그" in viol.columns:
            viol = viol[viol["제목 플래그"] != "제목주의"].copy()

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

        st.markdown("### 상세 모니터링 리스트")
        df_display = df.copy()
        df_display["판매 링크"] = df_display["링크"]

        if hide_title_warning and "제목 플래그" in df_display.columns:
            df_display = df_display[df_display["제목 플래그"] != "제목주의"].copy()

        st.dataframe(
            df_display[
                [
                    "상태",
                    "제목 플래그",
                    "제목 진단",
                    "판매처",
                    "판매가",
                    "가이드가",
                    "차액",
                    "제품명",
                    "매칭키워드",
                    "판매 링크",
                ]
            ],
            column_config={
                "판매 링크": st.column_config.LinkColumn("바로가기", display_text="링크이동"),
                "판매가": st.column_config.NumberColumn(format="%d원"),
                "가이드가": st.column_config.NumberColumn(format="%d원"),
                "차액": st.column_config.NumberColumn(format="%d원"),
                "제목 진단": st.column_config.TextColumn("제목 진단"),
                "매칭키워드": st.column_config.TextColumn("매칭키워드"),
            },
            use_container_width=True,
            height=600,
            hide_index=True,
        )

        # 엑셀 다운로드 데이터 구성
        df_for_excel = df.copy()

        # 링크 컬럼 분리(엑셀 하이퍼링크 안정화)
        df_for_excel.insert(df_for_excel.columns.get_loc("링크") + 1, "원본 URL(복사용)", df_for_excel["링크"])
        df_for_excel.insert(df_for_excel.columns.get_loc("링크") + 1, "판매 링크(클릭)", "바로가기")
        df_for_excel = df_for_excel.drop(columns=["링크"])

        # 엑셀 컬럼 순서
        wanted_cols = [
            "상태",
            "제목 플래그",
            "제목 진단",
            "판매처",
            "판매가",
            "가이드가",
            "차액",
            "제품명",
            "매칭키워드",
            "판매 링크(클릭)",
            "원본 URL(복사용)",
        ]
        # 일부 컬럼이 없을 수 있으니 안전하게
        wanted_cols = [c for c in wanted_cols if c in df_for_excel.columns]
        df_for_excel = df_for_excel[wanted_cols]

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
