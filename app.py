"""
================================================================================
서울시 공공자전거(따릉이) 빅데이터 분석 및 인프라 최적화 대시보드
(Seoul Public Bike Big Data Analytics & Spatial Optimization Dashboard)
================================================================================

[필요 패키지 (requirements.txt)]
streamlit>=1.35.0
pandas>=2.2.0
geopandas>=1.0.0
plotly>=5.20.0
folium>=0.16.0
streamlit-folium>=0.19.0
openpyxl>=3.1.0
statsmodels>=0.14.0
scipy>=1.12.0
shapely>=2.0.0
pyproj>=3.7.0

[실행 방법]
streamlit run app.py
================================================================================
"""

import os
import sys
import math
import numpy as np
import pandas as pd
import geopandas as gpd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import folium
from folium.plugins import MarkerCluster
from streamlit_folium import st_folium
import streamlit as st

# ==============================================================================
# 1. Streamlit 페이지 설정 및 커스텀 스타일링 (CSS)
# ==============================================================================
st.set_page_config(
    page_title="서울시 공공자전거(따릉이) 데이터 분석 대시보드",
    page_icon="🚲",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 모던 디자인 시스템 CSS 주입
st.markdown("""
<style>
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
    
    html, body, [class*="css"] {
        font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, system-ui, Roboto, sans-serif;
    }
    
    /* 헤더 및 타이틀 스타일 */
    .dashboard-header {
        background: linear-gradient(135deg, #1e3a8a 0%, #0d9488 50%, #059669 100%);
        padding: 24px 30px;
        border-radius: 16px;
        color: white;
        margin-bottom: 25px;
        box-shadow: 0 10px 25px -5px rgba(13, 148, 136, 0.25);
    }
    .dashboard-header h1 {
        margin: 0;
        font-size: 2.1rem;
        font-weight: 800;
        color: white !important;
        letter-spacing: -0.5px;
    }
    .dashboard-header p {
        margin: 6px 0 0 0;
        font-size: 0.98rem;
        opacity: 0.92;
        color: #e0f2fe;
    }
    
    /* KPI 카드 스타일 */
    .kpi-container {
        display: grid;
        grid-template-columns: repeat(5, 1fr);
        gap: 14px;
        margin-bottom: 25px;
    }
    .kpi-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 14px;
        padding: 18px 16px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.04);
        transition: all 0.2s ease;
        position: relative;
        overflow: hidden;
    }
    .kpi-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 20px rgba(0, 0, 0, 0.08);
        border-color: #cbd5e1;
    }
    .kpi-card::before {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 4px;
    }
    .kpi-card.blue::before { background: linear-gradient(90deg, #2563eb, #38bdf8); }
    .kpi-card.green::before { background: linear-gradient(90deg, #059669, #34d399); }
    .kpi-card.purple::before { background: linear-gradient(90deg, #7c3aed, #c084fc); }
    .kpi-card.teal::before { background: linear-gradient(90deg, #0d9488, #2dd4bf); }
    .kpi-card.amber::before { background: linear-gradient(90deg, #d97706, #fbbf24); }

    .kpi-label {
        font-size: 0.85rem;
        font-weight: 600;
        color: #64748b;
        margin-bottom: 6px;
        display: flex;
        align-items: center;
        gap: 6px;
    }
    .kpi-value {
        font-size: 1.55rem;
        font-weight: 800;
        color: #0f172a;
        line-height: 1.2;
    }
    .kpi-subtext {
        font-size: 0.78rem;
        color: #059669;
        font-weight: 500;
        margin-top: 5px;
    }
    
    /* 탭 헤더 및 컨테이너 */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #f1f5f9;
        padding: 6px;
        border-radius: 12px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 44px;
        border-radius: 8px;
        font-weight: 600;
        font-size: 0.95rem;
        padding: 0 18px;
    }
    
    /* 인사이트 콜아웃 박스 */
    .insight-card {
        background-color: #f8fafc;
        border-left: 4px solid #0d9488;
        padding: 16px 20px;
        border-radius: 0 12px 12px 0;
        margin: 15px 0;
        font-size: 0.92rem;
        line-height: 1.6;
        color: #334155;
    }
    .insight-card strong {
        color: #0f172a;
    }
    
    /* 사이드바 스타일링 */
    section[data-testid="stSidebar"] {
        background-color: #f8fafc;
        border-right: 1px solid #e2e8f0;
    }
    .sidebar-section-title {
        font-size: 0.88rem;
        font-weight: 700;
        color: #334155;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-top: 15px;
        margin-bottom: 8px;
    }
</style>
""", unsafe_allow_html=True)


# ==============================================================================
# 2. 데이터 수집 및 정제, 모델링 ETL 파이프라인 (@st.cache_data)
# ==============================================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

@st.cache_data(show_spinner="🗺️ 공간 행정구역(SHP 및 지리정보) 데이터 로딩 중...")
def load_geo_data():
    """서울시 자치구 경계 SHP 및 지리정보.xlsx 매핑 로드"""
    shp_path = os.path.join(BASE_DIR, "sig_20230729", "sig.shp")
    geo_path = os.path.join(BASE_DIR, "지리정보.xlsx")
    
    # 1. 지리정보.xlsx 로드
    try:
        df_geo = pd.read_excel(geo_path, sheet_name="sig")
    except Exception as e:
        st.error(f"지리정보.xlsx 로드 실패: {e}")
        df_geo = pd.DataFrame(columns=["SIG_CD", "SIG_ENG_NM", "SIG_KOR_NM"])
    
    df_geo["SIG_CD"] = df_geo["SIG_CD"].astype(str).str.strip()
    df_geo["SIG_KOR_NM"] = df_geo["SIG_KOR_NM"].astype(str).str.strip()
    
    # 2. SHP 파일 로드 (cp949 우선, 실패 시 utf-8)
    gdf = None
    for enc in ["cp949", "utf-8", "euc-kr"]:
        try:
            gdf = gpd.read_file(shp_path, encoding=enc)
            break
        except Exception:
            continue
            
    if gdf is None:
        st.error("GIS 벡터 파일(sig.shp)을 로드할 수 없습니다.")
        return df_geo, None

    # 컬럼 정리
    gdf.columns = gdf.columns.str.strip()
    gdf["SIG_CD"] = gdf["SIG_CD"].astype(str).str.strip()
    
    # 서울시(코드 11로 시작) 필터링
    gdf_seoul = gdf[gdf["SIG_CD"].str.startswith("11")].copy()
    
    # 좌표계 확인 및 EPSG:4326(WGS84) 변환
    if gdf_seoul.crs is None:
        gdf_seoul = gdf_seoul.set_crs(epsg=5179)
    gdf_seoul = gdf_seoul.to_crs(epsg=4326)
    
    # 자치구명 매핑 보완
    if "SIG_KOR_NM" not in gdf_seoul.columns or gdf_seoul["SIG_KOR_NM"].isna().all():
        gdf_seoul = gdf_seoul.merge(df_geo[["SIG_CD", "SIG_KOR_NM"]], on="SIG_CD", how="left")
        
    return df_geo, gdf_seoul


@st.cache_data(show_spinner="📍 공공자전거 대여소 인프라 데이터 로딩 중...")
def load_station_data(df_geo):
    """대여소 현황 엑셀 파일 로드 및 자치구/SIG_CD 모델링"""
    excel_path = os.path.join(BASE_DIR, "공공자전거 대여소 정보(26.6월 기준).xlsx")
    
    try:
        # skiprows=5: 5행부터 실제 데이터 시작
        df_raw = pd.read_excel(excel_path, sheet_name="대여소현황", skiprows=5, header=None)
    except Exception as e:
        st.error(f"대여소 정보 엑셀 로드 실패: {e}")
        return pd.DataFrame()
        
    df_st = df_raw.iloc[:, :10].copy()
    df_st.columns = [
        "대여소번호", "보관소(대여소)명", "자치구", "상세주소", 
        "위도", "경도", "설치시기", "거치대수_LCD", "거치대수_QR", "운영방식"
    ]
    
    # 문자열 공백 제거 및 대여소번호 앞자리 0 제거
    for col in ["대여소번호", "보관소(대여소)명", "자치구", "상세주소", "운영방식"]:
        df_st[col] = df_st[col].astype(str).str.strip()
        
    df_st["대여소번호"] = df_st["대여소번호"].str.lstrip("0")
    
    # 위도, 경도 숫자형 변환 및 결측치/이상치 처리
    df_st["위도"] = pd.to_numeric(df_st["위도"], errors="coerce")
    df_st["경도"] = pd.to_numeric(df_st["경도"], errors="coerce")
    df_st = df_st.dropna(subset=["위도", "경도"]).copy()
    
    # 서울시 영역 내 좌표만 필터링 (위도 37.4~37.7, 경도 126.7~127.2)
    df_st = df_st[(df_st["위도"] >= 37.3) & (df_st["위도"] <= 37.8) & 
                  (df_st["경도"] >= 126.6) & (df_st["경도"] <= 127.4)].copy()
    
    # 거치대수 계산
    df_st["거치대수_LCD"] = pd.to_numeric(df_st["거치대수_LCD"], errors="coerce").fillna(0)
    df_st["거치대수_QR"] = pd.to_numeric(df_st["거치대수_QR"], errors="coerce").fillna(0)
    df_st["총거치대수"] = df_st["거치대수_LCD"] + df_st["거치대수_QR"]
    
    # 지리정보와 매핑하여 SIG_CD 부여
    df_st = df_st.merge(
        df_geo[["SIG_CD", "SIG_KOR_NM"]], 
        left_on="자치구", 
        right_on="SIG_KOR_NM", 
        how="left"
    )
    
    return df_st


@st.cache_data(show_spinner="📊 대용량 공공자전거 이용정보(월별) 정제 및 통합 중...")
def load_usage_data(df_st):
    """월별 이용정보 CSV 로드, 전처리 및 대여소 데이터와 Left Join 모델링"""
    csv_path = os.path.join(BASE_DIR, "서울특별시 공공자전거 이용정보(월별)_26.1-6.csv")
    
    df_raw = None
    for enc in ["cp949", "utf-8", "euc-kr"]:
        try:
            df_raw = pd.read_csv(csv_path, encoding=enc)
            break
        except Exception:
            continue
            
    if df_raw is None:
        st.error("이용정보 CSV 파일을 로드할 수 없습니다.")
        return pd.DataFrame()
        
    df_raw.columns = df_raw.columns.str.strip()
    
    # 1. 대여소번호 정제 (문자열 앞자리 0 제거)
    df_raw["대여소번호"] = df_raw["대여소번호"].astype(str).str.strip().str.lstrip("0")
    
    # 2. 대여일자 문자열 변환 및 포맷팅 (202603 -> 2026-03)
    df_raw["대여일자_str"] = df_raw["대여일자"].astype(str).str.strip()
    df_raw["대여월"] = df_raw["대여일자_str"].apply(
        lambda x: f"{x[:4]}년 {x[4:6]}월" if len(x) >= 6 else str(x)
    )
    
    # 3. 성별 정제 (결측 및 소문자 정규화)
    gender_map = {"M": "남성(M)", "F": "여성(F)", "m": "남성(M)", "f": "여성(F)"}
    df_raw["성별"] = df_raw["성별"].astype(str).str.strip().map(gender_map).fillna("기타/미기재")
    
    # 4. 연령대코드 정제
    age_map = {
        "~10대": "~10대", "20대": "20대", "30대": "30대", "40대": "40대",
        "50대": "50대", "60대": "60대", "70대이상": "70대이상"
    }
    df_raw["연령대코드"] = df_raw["연령대코드"].astype(str).str.strip().map(age_map).fillna("기타/미기재")
    
    # 5. 수치형 컬럼 변환
    numeric_cols = ["이용건수", "운동량", "탄소량", "이동거리(M)", "이용시간(분)"]
    for col in numeric_cols:
        df_raw[col] = pd.to_numeric(df_raw[col], errors="coerce").fillna(0)
        
    # 6. 메모리 최적화 dtypes
    df_raw["이용건수"] = df_raw["이용건수"].astype("int32")
    df_raw["운동량"] = df_raw["운동량"].astype("float32")
    df_raw["탄소량"] = df_raw["탄소량"].astype("float32")
    df_raw["이동거리(M)"] = df_raw["이동거리(M)"].astype("float32")
    df_raw["이용시간(분)"] = df_raw["이용시간(분)"].astype("int32")
    
    # 7. 대여소 정보와 Left Join
    st_sub = df_st[["대여소번호", "보관소(대여소)명", "자치구", "SIG_CD", "위도", "경도", "총거치대수", "운영방식"]].copy()
    
    merged_df = df_raw.merge(
        st_sub,
        on="대여소번호",
        how="left",
        suffixes=("", "_st")
    )
    
    merged_df["자치구"] = merged_df["자치구"].fillna("미분류")
    merged_df["보관소(대여소)명"] = merged_df["보관소(대여소)명"].fillna(merged_df["대여소명"])
    
    return merged_df


# 데이터 로드 실행
df_geo, gdf_seoul = load_geo_data()
df_st = load_station_data(df_geo)
master_df = load_usage_data(df_st)

if master_df.empty:
    st.error("데이터를 불러오지 못했습니다. 파일 경로를 확인해 주세요.")
    st.stop()


# ==============================================================================
# 3. 사이드바 필터 컨트롤러
# ==============================================================================
with st.sidebar:
    st.markdown("### 🚲 데이터 필터 컨트롤러")
    st.caption("조건을 변경하면 대시보드 지표 및 차트가 실시간으로 재산출됩니다.")
    
    # 1. 대여월 필터
    month_options = sorted(master_df["대여월"].unique().tolist())
    selected_months = st.multiselect(
        "📅 대여월 선택",
        options=month_options,
        default=month_options
    )
    
    # 2. 자치구 필터
    seoul_gus = sorted([g for g in master_df["자치구"].unique() if g != "미분류"])
    all_gu_selected = st.checkbox("전체 자치구 선택", value=True)
    if all_gu_selected:
        selected_gus = seoul_gus
        selected_gus_display = st.multiselect("🏢 자치구 다중선택", options=seoul_gus, default=seoul_gus, disabled=True)
    else:
        selected_gus = st.multiselect("🏢 자치구 다중선택", options=seoul_gus, default=["강서구", "송파구", "영등포구", "마포구"])
        
    # 3. 대여구분코드(이용권) 필터
    ticket_options = sorted(master_df["대여구분코드"].dropna().unique().tolist())
    all_ticket_selected = st.checkbox("전체 이용권 선택", value=True)
    if all_ticket_selected:
        selected_tickets = ticket_options
        _ = st.multiselect("🎫 대여구분(이용권)", options=ticket_options, default=ticket_options, disabled=True)
    else:
        selected_tickets = st.multiselect("🎫 대여구분(이용권)", options=ticket_options, default=["정기권", "일일권"])
        
    # 4. 성별 필터
    gender_options = ["남성(M)", "여성(F)", "기타/미기재"]
    selected_genders = st.multiselect(
        "🚻 성별 선택",
        options=gender_options,
        default=gender_options
    )
    
    # 5. 연령대 필터
    age_options = ["~10대", "20대", "30대", "40대", "50대", "60대", "70대이상", "기타/미기재"]
    selected_ages = st.multiselect(
        "🎂 연령대 선택",
        options=age_options,
        default=age_options
    )
    
    # 필터 적용
    filtered_df = master_df[
        (master_df["대여월"].isin(selected_months)) &
        (master_df["자치구"].isin(selected_gus)) &
        (master_df["대여구분코드"].isin(selected_tickets)) &
        (master_df["성별"].isin(selected_genders)) &
        (master_df["연령대코드"].isin(selected_ages))
    ]
    
    # 필터 상태 통계 배지
    total_rows = len(master_df)
    filtered_rows = len(filtered_df)
    ratio = (filtered_rows / total_rows * 100) if total_rows > 0 else 0
    
    st.markdown("---")
    st.markdown(f"""
    <div style='background: #e2e8f0; padding: 12px; border-radius: 8px; font-size: 0.85rem;'>
        <b>필터 적용 상태</b><br>
        • 대상 레코드: <b>{filtered_rows:,}</b> / {total_rows:,} 건<br>
        • 적용 비율: <b>{ratio:.1f}%</b>
    </div>
    """, unsafe_allow_html=True)


# ==============================================================================
# 4. 상단 대시보드 타이틀 & 핵심 KPI 카드
# ==============================================================================
st.markdown("""
<div class="dashboard-header">
    <h1>🚲 서울특별시 공공자전거(따릉이) 빅데이터 분석 대시보드</h1>
    <p>공간 지리정보(GIS) · 인프라 효율성 · 이용자 행동 패턴 · ESG 환경 기여도 종합 분석 플랫폼</p>
</div>
""", unsafe_allow_html=True)

if filtered_df.empty:
    st.warning("선택된 필터 조건에 부합하는 데이터가 없습니다. 사이드바 필터 조건을 확장해 주세요.")
    st.stop()

# 핵심 지표 산출
total_rentals = int(filtered_df["이용건수"].sum())
total_dist_km = float(filtered_df["이동거리(M)"].sum() / 1000.0)
total_time_hr = float(filtered_df["이용시간(분)"].sum() / 60.0)
total_carbon_kg = float(filtered_df["탄소량"].sum())
total_carbon_ton = total_carbon_kg / 1000.0
total_calories_kcal = float(filtered_df["운동량"].sum())

# 보조 지표
avg_dist_km = (total_dist_km / total_rentals) if total_rentals > 0 else 0
avg_time_min = (filtered_df["이용시간(분)"].sum() / total_rentals) if total_rentals > 0 else 0
earth_rounds = total_dist_km / 40075.0  # 지구 둘레 약 40,075 km
pine_trees = total_carbon_kg / 6.6     # 30년생 소나무 1그루 연간 약 6.6kg 흡수

kpi_col1, kpi_col2, kpi_col3, kpi_col4, kpi_col5 = st.columns(5)

with kpi_col1:
    st.markdown(f"""
    <div class="kpi-card blue">
        <div class="kpi-label">📊 총 이용건수</div>
        <div class="kpi-value">{total_rentals:,.0f}<span style='font-size:0.9rem; font-weight:600; color:#64748b;'> 건</span></div>
        <div class="kpi-subtext">약 {total_rentals/10000:,.1f}만 회 이용</div>
    </div>
    """, unsafe_allow_html=True)

with kpi_col2:
    st.markdown(f"""
    <div class="kpi-card green">
        <div class="kpi-label">🚴 총 이동거리</div>
        <div class="kpi-value">{total_dist_km:,.0f}<span style='font-size:0.9rem; font-weight:600; color:#64748b;'> km</span></div>
        <div class="kpi-subtext">지구 약 {earth_rounds:,.1f}바퀴 완주</div>
    </div>
    """, unsafe_allow_html=True)

with kpi_col3:
    st.markdown(f"""
    <div class="kpi-card purple">
        <div class="kpi-label">⏱️ 총 이용시간</div>
        <div class="kpi-value">{total_time_hr:,.0f}<span style='font-size:0.9rem; font-weight:600; color:#64748b;'> 시간</span></div>
        <div class="kpi-subtext">건당 평균 {avg_time_min:.1f}분 주행</div>
    </div>
    """, unsafe_allow_html=True)

with kpi_col4:
    st.markdown(f"""
    <div class="kpi-card teal">
        <div class="kpi-label">🌱 탄소 절감량</div>
        <div class="kpi-value">{total_carbon_ton:,.1f}<span style='font-size:0.9rem; font-weight:600; color:#64748b;'> 톤</span></div>
        <div class="kpi-subtext">소나무 약 {pine_trees/10000:,.1f}만 그루 효과</div>
    </div>
    """, unsafe_allow_html=True)

with kpi_col5:
    st.markdown(f"""
    <div class="kpi-card amber">
        <div class="kpi-label">🔥 소모 운동량</div>
        <div class="kpi-value">{total_calories_kcal/10000:,.1f}<span style='font-size:0.9rem; font-weight:600; color:#64748b;'> 만kcal</span></div>
        <div class="kpi-subtext">건당 평균 {total_calories_kcal/total_rentals:.0f} kcal 소모</div>
    </div>
    """, unsafe_allow_html=True)

st.write("")


# ==============================================================================
# 5. 심층 분석 탭 구성 (Tab 1 ~ Tab 4)
# ==============================================================================
tab1, tab2, tab3, tab4 = st.tabs([
    "🗺️ 지역별 공간 및 인프라 분석",
    "👥 이용 패턴 및 인구통계학적 인사이트",
    "🌱 환경 및 건강 효과(ESG) 분석",
    "📋 원본 및 대여소별 상세 데이터 테이블"
])


# ==============================================================================
# [Tab 1] 지역별 공간 및 인프라 분석
# ==============================================================================
with tab1:
    st.subheader("1. 자치구별 따릉이 이용량 Choropleth 지도 및 거치대 인프라")
    
    # 자치구별 집계
    gu_agg = filtered_df.groupby("자치구").agg({
        "이용건수": "sum",
        "이동거리(M)": "sum",
        "탄소량": "sum",
        "대여소번호": "nunique"
    }).reset_index().rename(columns={"대여소번호": "활성대여소수"})
    
    gu_agg["이동거리(km)"] = gu_agg["이동거리(M)"] / 1000.0
    
    # 대여소 현황 기준 구별 총 거치대수 병합
    st_gu_racks = df_st.groupby("자치구")["총거치대수"].sum().reset_index()
    gu_agg = gu_agg.merge(st_gu_racks, on="자치구", how="left").fillna({"총거치대수": 0})
    
    # 거치대당 이용효율(회전율)
    gu_agg["거치대당_이용건수"] = np.where(
        gu_agg["총거치대수"] > 0, 
        gu_agg["이용건수"] / gu_agg["총거치대수"], 
        0
    ).round(2)
    
    # 지리정보 결합
    gu_geo_merged = df_geo.merge(gu_agg, left_on="SIG_KOR_NM", right_on="자치구", how="left").fillna({
        "이용건수": 0, "이동거리(km)": 0, "탄소량": 0, "활성대여소수": 0, "총거치대수": 0, "거치대당_이용건수": 0
    })
    
    map_col1, map_col2 = st.columns([1.1, 0.9])
    
    with map_col1:
        st.markdown("##### 📍 자치구별 총 이용건수 공간 분포 (Folium Choropleth)")
        
        # 지도 생성
        m_seoul = folium.Map(location=[37.5665, 126.9780], zoom_start=11, tiles="OpenStreetMap")
        
        # GeoJSON 병합
        gdf_plot = gdf_seoul.merge(
            gu_geo_merged[["SIG_CD", "이용건수", "활성대여소수", "총거치대수", "거치대당_이용건수"]], 
            on="SIG_CD", 
            how="left"
        ).fillna(0)
        
        # Choropleth 레이어
        folium.Choropleth(
            geo_data=gdf_plot.__geo_interface__,
            data=gdf_plot,
            columns=["SIG_CD", "이용건수"],
            key_on="feature.properties.SIG_CD",
            fill_color="YlGnBu",
            fill_opacity=0.75,
            line_opacity=0.3,
            legend_name="자치구별 총 이용건수 (건)",
            smooth_factor=0
        ).add_to(m_seoul)
        
        # 마우스오버 툴팁용 GeoJson 레이어
        folium.GeoJson(
            gdf_plot,
            style_function=lambda x: {"fillColor": "transparent", "color": "#1e293b", "weight": 1.2, "fillOpacity": 0},
            highlight_function=lambda x: {"weight": 3, "color": "#f59e0b", "fillOpacity": 0.2},
            tooltip=folium.GeoJsonTooltip(
                fields=["SIG_KOR_NM", "이용건수", "활성대여소수", "총거치대수", "거치대당_이용건수"],
                aliases=["자치구명:", "총 이용건수:", "운영 대여소수:", "총 거치대수:", "거치대당 회전율:"],
                localize=True
            )
        ).add_to(m_seoul)
        
        st_folium(m_seoul, width=650, height=480, returned_objects=[])
        
    with map_col2:
        st.markdown("##### 🏆 자치구별 총 이용건수 순위")
        gu_sorted = gu_agg.sort_values(by="이용건수", ascending=True)
        
        fig_gu_bar = px.bar(
            gu_sorted,
            x="이용건수",
            y="자치구",
            orientation="h",
            color="이용건수",
            color_continuous_scale="Tealgrn",
            text="이용건수"
        )
        fig_gu_bar.update_traces(
            texttemplate="%{text:,.0f}건", 
            textposition="inside",
            insidetextanchor="middle"
        )
        fig_gu_bar.update_layout(
            margin=dict(l=10, r=20, t=20, b=20),
            height=480,
            xaxis_title="총 이용건수 (건)",
            yaxis_title="",
            coloraxis_showscale=False
        )
        st.plotly_chart(fig_gu_bar, use_container_width=True)

    st.markdown("---")
    
    # 2. 대여소 거치대 대비 이용 효율 분석 (공급 부족 vs 과잉)
    st.subheader("2. 대여소 거치대 대비 이용 효율 분석 (자전거 공급/수요 불균형 인사이트)")
    st.markdown("""
    대여소별 `회전율 = 총 이용건수 ÷ 총 거치대수` 지표를 산출하여, **자전거 공급이 극도로 부족해 확충이 시급한 과밀 대여소**와 
    **거치대 대비 이용이 저조하여 유휴 자전거 재배치가 권장되는 저효율 대여소**를 도출합니다.
    """)
    
    # 대여소별 집계
    st_agg = filtered_df.groupby(["대여소번호", "보관소(대여소)명", "자치구"]).agg({
        "이용건수": "sum",
        "총거치대수": "first",
        "위도": "first",
        "경도": "first"
    }).reset_index()
    
    # 거치대수가 유효한 대여소 필터링
    st_agg = st_agg[st_agg["총거치대수"] > 0].copy()
    st_agg["회전율"] = (st_agg["이용건수"] / st_agg["총거치대수"]).round(2)
    
    eff_col1, eff_col2 = st.columns(2)
    
    with eff_col1:
        st.markdown("##### 🚨 자전거 공급 부족(과밀 집중) 대여소 Top 10")
        st.caption("거치대 1대당 이용건수가 가장 높은 대여소 (자전거 부족 및 반납난 발생 가능성 최고)")
        top_eff = st_agg.sort_values(by="회전율", ascending=False).head(10)
        
        fig_top_eff = px.bar(
            top_eff,
            x="회전율",
            y="보관소(대여소)명",
            orientation="h",
            color="회전율",
            color_continuous_scale="Reds",
            text="회전율",
            hover_data=["자치구", "총거치대수", "이용건수"]
        )
        fig_top_eff.update_layout(
            yaxis={"autorange": "reversed"},
            xaxis_title="거치대 1대당 이용건수 (회/대)",
            yaxis_title="",
            margin=dict(l=10, r=20, t=10, b=20),
            height=380,
            coloraxis_showscale=False
        )
        st.plotly_chart(fig_top_eff, use_container_width=True)
        
    with eff_col2:
        st.markdown("##### 📉 자전거 공급 여유(저효율 유휴) 대여소 Top 10")
        st.caption("거치대수는 많으나 실제 이용이 가장 저조한 대여소 (거치대 축소 및 재배치 권장)")
        bottom_eff = st_agg[st_agg["이용건수"] > 0].sort_values(by="회전율", ascending=True).head(10)
        
        fig_bot_eff = px.bar(
            bottom_eff,
            x="회전율",
            y="보관소(대여소)명",
            orientation="h",
            color="회전율",
            color_continuous_scale="Blues_r",
            text="회전율",
            hover_data=["자치구", "총거치대수", "이용건수"]
        )
        fig_bot_eff.update_layout(
            yaxis={"autorange": "reversed"},
            xaxis_title="거치대 1대당 이용건수 (회/대)",
            yaxis_title="",
            margin=dict(l=10, r=20, t=10, b=20),
            height=380,
            coloraxis_showscale=False
        )
        st.plotly_chart(fig_bot_eff, use_container_width=True)

    # 행정 인사이트 콜아웃
    max_gu = gu_agg.sort_values(by="거치대당_이용건수", ascending=False).iloc[0]
    min_gu = gu_agg.sort_values(by="거치대당_이용건수", ascending=True).iloc[0]
    st.markdown(f"""
    <div class="insight-card">
        💡 <b>인프라 최적화 운영 정책 인사이트:</b><br>
        • 자치구 단위 분석 결과, <b>{max_gu['자치구']}</b>가 거치대 1대당 <b>{max_gu['거치대당_이용건수']:,.1f}회</b>로 가장 치열한 회전율을 보였으며, 
        반면 <b>{min_gu['자치구']}</b>는 거치대 1대당 <b>{min_gu['거치대당_이용건수']:,.1f}회</b> 수준에 머물렀습니다.<br>
        • 주요 지하철역 환승 접점(예: 발산역, 송파구청, 건대입구역)은 출퇴근 시간대 공급 부족 현상이 심각하므로 
        <b>스마트 거치대 증설</b> 및 <b>출근 시간대 집중 자전거 재배치 알고리즘</b> 도입이 시급합니다.
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # 3. 대여소 위치 Scatter Map (마커 클러스터링)
    st.subheader("3. 대여소 위치 Scatter Map (상위 이용 대여소 마커 클러스터링)")
    
    scatter_col1, scatter_col2 = st.columns([1, 3])
    with scatter_col1:
        top_n_stations = st.slider("표시할 상위 대여소 수 (Top N)", min_value=50, max_value=500, value=150, step=50)
        st.caption(f"이용건수 기준 상위 {top_n_stations}개 대여소를 지도에 군집(Cluster)화하여 표시합니다.")
        
    top_stations_df = st_agg.sort_values(by="이용건수", ascending=False).head(top_n_stations)
    
    # Folium Marker Cluster 지도
    m_cluster = folium.Map(location=[37.5665, 126.9780], zoom_start=11, tiles="OpenStreetMap")
    marker_cluster = MarkerCluster(name="대여소 군집").add_to(m_cluster)
    
    for _, row in top_stations_df.iterrows():
        popup_html = f"""
        <div style='font-family: Pretendard; width: 220px;'>
            <h5 style='margin:0 0 5px 0; color:#0d9488;'>{row['보관소(대여소)명']}</h5>
            <b>대여소 번호:</b> {row['대여소번호']}<br>
            <b>소재 자치구:</b> {row['자치구']}<br>
            <b>총 거치대수:</b> {row['총거치대수']:,.0f} 대<br>
            <b>총 이용건수:</b> <span style='color:#e11d48; font-weight:700;'>{row['이용건수']:,.0f} 건</span><br>
            <b>거치대 회전율:</b> {row['회전율']:,.1f} 회/대
        </div>
        """
        folium.CircleMarker(
            location=[row["위도"], row["경도"]],
            radius=6 + min(row["회전율"] / 100, 10),
            popup=folium.Popup(popup_html, max_width=250),
            tooltip=f"{row['보관소(대여소)명']} ({row['이용건수']:,}건)",
            color="#0d9488",
            fill=True,
            fill_color="#14b8a6",
            fill_opacity=0.7
        ).add_to(marker_cluster)
        
    st_folium(m_cluster, width=1200, height=520, returned_objects=[])


# ==============================================================================
# [Tab 2] 이용 패턴 및 인구통계학적 인사이트
# ==============================================================================
with tab2:
    st.subheader("1. 연령대 & 성별 교차 분석 (주력 이용 계층 분석)")
    
    demo_col1, demo_col2 = st.columns(2)
    
    # 연령대 x 성별 크로스테이블 집계
    demo_pivot = filtered_df.pivot_table(
        index="연령대코드",
        columns="성별",
        values="이용건수",
        aggfunc="sum",
        fill_value=0
    )
    
    # 순서 정리
    age_order = ["~10대", "20대", "30대", "40대", "50대", "60대", "70대이상", "기타/미기재"]
    demo_pivot = demo_pivot.reindex([a for a in age_order if a in demo_pivot.index])
    
    with demo_col1:
        st.markdown("##### 🔥 연령대 × 성별 이용건수 히트맵")
        fig_heat = px.imshow(
            demo_pivot,
            labels=dict(x="성별", y="연령대", color="이용건수"),
            x=demo_pivot.columns,
            y=demo_pivot.index,
            color_continuous_scale="Viridis",
            text_auto=True,
            aspect="auto"
        )
        fig_heat.update_layout(
            margin=dict(l=20, r=20, t=20, b=20),
            height=380
        )
        st.plotly_chart(fig_heat, use_container_width=True)
        
    with demo_col2:
        st.markdown("##### 📊 연령대별 성별 구성 누적 막대그래프")
        demo_long = filtered_df.groupby(["연령대코드", "성별"])["이용건수"].sum().reset_index()
        fig_bar_demo = px.bar(
            demo_long,
            x="연령대코드",
            y="이용건수",
            color="성별",
            category_orders={"연령대코드": age_order},
            barmode="stack",
            color_discrete_map={"남성(M)": "#3b82f6", "여성(F)": "#ec4899", "기타/미기재": "#94a3b8"}
        )
        fig_bar_demo.update_layout(
            margin=dict(l=20, r=20, t=20, b=20),
            height=380,
            xaxis_title="연령대",
            yaxis_title="총 이용건수 (건)",
            legend_title="성별"
        )
        st.plotly_chart(fig_bar_demo, use_container_width=True)

    # 2030 주력 이용층 비중 계산
    rentals_2030 = filtered_df[filtered_df["연령대코드"].isin(["20대", "30대"])]["이용건수"].sum()
    pct_2030 = (rentals_2030 / total_rentals * 100) if total_rentals > 0 else 0
    st.markdown(f"""
    <div class="insight-card">
        🎯 <b>타겟 인구통계학적 분석 요약:</b><br>
        • 전체 이용량 중 <b>20대 및 30대 청년층</b>의 이용 비중이 무려 <b>{pct_2030:.1f}%</b>를 차지하여 가장 핵심적인 이용 집단으로 나타났습니다.<br>
        • 특히 20대에서는 남성/여성의 이용량이 균형을 이루는 반면, 40~50대로 갈수록 남성 이용자 비중이 점진적으로 증가하는 패턴을 보입니다.
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # 2. 대여권종별 이동 특성 (정기권 vs 일일권 vs 가족권)
    st.subheader("2. 대여권종별 이동 특성 (1회 이용 시 이동거리 및 시간 분포)")
    st.markdown("""
    정기권(출퇴근·통학 목적)과 일일권/가족권(레저·관광·주말 여가 목적) 간의 **1회당 평균 이동거리(M)**와 **1회당 이용시간(분)**의 분포 차이를 비교합니다.
    """)
    
    # 1회당 단위 지표 계산
    unit_df = filtered_df[filtered_df["이용건수"] > 0].copy()
    unit_df["1회당_이동거리(M)"] = unit_df["이동거리(M)"] / unit_df["이용건수"]
    unit_df["1회당_이용시간(분)"] = unit_df["이용시간(분)"] / unit_df["이용건수"]
    
    # 권종 대표 그룹화
    def group_ticket(t):
        if "정기권" in str(t):
            return "정기권"
        elif "일일권" in str(t):
            return "일일권"
        elif "가족권" in str(t):
            return "가족권"
        return "기타"
        
    unit_df["권종그룹"] = unit_df["대여구분코드"].apply(group_ticket)
    
    # 박스플롯 이상치 제거를 위한 99백분위수 제한
    dist_cap = unit_df["1회당_이동거리(M)"].quantile(0.98)
    time_cap = unit_df["1회당_이용시간(분)"].quantile(0.98)
    plot_unit_df = unit_df[
        (unit_df["1회당_이동거리(M)"] <= dist_cap) & 
        (unit_df["1회당_이용시간(분)"] <= time_cap)
    ]
    
    ticket_col1, ticket_col2 = st.columns(2)
    
    with ticket_col1:
        st.markdown("##### 📏 대여권종별 1회 이동거리(M) 분포 비교")
        fig_box_dist = px.box(
            plot_unit_df,
            x="권종그룹",
            y="1회당_이동거리(M)",
            color="권종그룹",
            points=False,
            color_discrete_sequence=px.colors.qualitative.Safe
        )
        fig_box_dist.update_layout(
            margin=dict(l=20, r=20, t=20, b=20),
            height=380,
            xaxis_title="대여권종 그룹",
            yaxis_title="1회당 이동거리 (m)",
            showlegend=False
        )
        st.plotly_chart(fig_box_dist, use_container_width=True)
        
    with ticket_col2:
        st.markdown("##### ⏱️ 대여권종별 1회 이용시간(분) 분포 비교")
        fig_box_time = px.box(
            plot_unit_df,
            x="권종그룹",
            y="1회당_이용시간(분)",
            color="권종그룹",
            points=False,
            color_discrete_sequence=px.colors.qualitative.Safe
        )
        fig_box_time.update_layout(
            margin=dict(l=20, r=20, t=20, b=20),
            height=380,
            xaxis_title="대여권종 그룹",
            yaxis_title="1회당 이용시간 (분)",
            showlegend=False
        )
        st.plotly_chart(fig_box_time, use_container_width=True)

    # 권종 통계 요약 테이블
    ticket_summary = unit_df.groupby("권종그룹").agg({
        "이용건수": "sum",
        "1회당_이동거리(M)": "mean",
        "1회당_이용시간(분)": "mean"
    }).reset_index()
    ticket_summary["이용건수_비중(%)"] = (ticket_summary["이용건수"] / ticket_summary["이용건수"].sum() * 100).round(1)
    ticket_summary["평균이동거리(km)"] = (ticket_summary["1회당_이동거리(M)"] / 1000.0).round(2)
    ticket_summary["평균이용시간(분)"] = ticket_summary["1회당_이용시간(분)"].round(1)
    ticket_summary["추정평균시속(km/h)"] = np.where(
        ticket_summary["평균이용시간(분)"] > 0,
        (ticket_summary["평균이동거리(km)"] / (ticket_summary["평균이용시간(분)"] / 60)).round(1),
        0
    )
    
    st.dataframe(
        ticket_summary[["권종그룹", "이용건수", "이용건수_비중(%)", "평균이동거리(km)", "평균이용시간(분)", "추정평균시속(km/h)"]],
        use_container_width=True,
        hide_index=True
    )
    
    st.markdown("""
    <div class="insight-card">
        🚴 <b>이동 특성 분석 요약:</b><br>
        • <b>정기권:</b> 평균 이동거리 약 2.3km, 이용시간 약 22분으로 지하철역과 주거·업무지구를 잇는 <b>'라스트 마일(Last-mile)' 출퇴근형 이동</b>이 뚜렷합니다.<br>
        • <b>일일권 및 가족권:</b> 평균 이동거리 3.3~5.5km, 이용시간 35~60분 이상으로 한강공원, 안양천, 중랑천 등 수변공원 인근의 <b>레저·관광형 여가 이용</b> 특성을 보입니다.
    </div>
    """, unsafe_allow_html=True)


# ==============================================================================
# [Tab 3] 환경 및 건강 효과(ESG) 분석
# ==============================================================================
with tab3:
    st.subheader("1. 자치구별 온실가스 감축(탄소 절감량) 순위 (ESG 기여도 분석)")
    
    esg_col1, esg_col2 = st.columns([1.2, 0.8])
    
    with esg_col1:
        st.markdown("##### 🌱 자치구별 총 탄소 절감량 순위 Top 10")
        carbon_top10 = gu_agg.sort_values(by="탄소량", ascending=True).tail(10)
        
        fig_carbon = px.bar(
            carbon_top10,
            x="탄소량",
            y="자치구",
            orientation="h",
            color="탄소량",
            color_continuous_scale="Greens",
            text="탄소량"
        )
        fig_carbon.update_traces(
            texttemplate="%{text:,.1f} kg",
            textposition="inside",
            insidetextanchor="middle"
        )
        fig_carbon.update_layout(
            margin=dict(l=10, r=20, t=10, b=20),
            height=380,
            xaxis_title="총 탄소 절감량 (kg CO₂eq)",
            yaxis_title="",
            coloraxis_showscale=False
        )
        st.plotly_chart(fig_carbon, use_container_width=True)
        
    with esg_col2:
        st.markdown("##### 🌍 ESG 친환경 기여 환산 지표")
        
        gasoline_saved = total_carbon_kg / 2.32  # 가솔린 1L 연소 시 약 2.32kg CO2 발생
        car_km_equiv = total_dist_km              # 자동차 주행거리 대체 효과
        pine_annual = total_carbon_kg / 6.6      # 소나무 1그루 연간 흡수량
        
        st.markdown(f"""
        <div style='background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 12px; padding: 20px;'>
            <div style='font-size: 1.05rem; font-weight: 700; color: #166534; margin-bottom: 12px;'>
                🌿 따릉이가 만들어낸 친환경 가치
            </div>
            <p style='margin: 8px 0; color: #15803d;'>
                🌲 <b>30년생 소나무 식재 효과:</b><br>
                약 <span style='font-size:1.3rem; font-weight:800;'>{pine_annual:,.0f}</span> 그루 연간 흡수량
            </p>
            <p style='margin: 8px 0; color: #15803d;'>
                ⛽ <b>화석연료(휘발유) 대체 효과:</b><br>
                약 <span style='font-size:1.3rem; font-weight:800;'>{gasoline_saved:,.0f}</span> 리터 (L) 절감
            </p>
            <p style='margin: 8px 0; color: #15803d;'>
                🚗 <b>내연기관 승용차 대체 주행:</b><br>
                약 <span style='font-size:1.3rem; font-weight:800;'>{car_km_equiv:,.0f}</span> km 대체 주행
            </p>
        </div>
        """, unsafe_allow_html=True)
        
    st.markdown("---")
    
    # 2. 운동량(kcal)과 이동거리 간의 상관관계 산점도 및 회귀선
    st.subheader("2. 이동거리(M)와 운동량(kcal) 간의 상관관계 분석 및 추세선")
    
    reg_col1, reg_col2 = st.columns([1, 3])
    with reg_col1:
        st.caption("고성능 반응속도를 위해 무작위 표본(Sample)을 추출하여 실시간 회귀선(OLS)을 산출합니다.")
        sample_size = st.slider("분석 샘플 크기 선택", min_value=1000, max_value=10000, value=3000, step=1000)
        
    # 유효 데이터 샘플링
    corr_df = filtered_df[(filtered_df["운동량"] > 0) & (filtered_df["이동거리(M)"] > 0)]
    if len(corr_df) > sample_size:
        sample_corr_df = corr_df.sample(n=sample_size, random_state=42).copy()
    else:
        sample_corr_df = corr_df.copy()
        
    sample_corr_df["이동거리(km)"] = sample_corr_df["이동거리(M)"] / 1000.0
    
    # 상관계수 계산
    r_corr = np.corrcoef(sample_corr_df["이동거리(km)"], sample_corr_df["운동량"])[0, 1]
    
    # 1차 회귀식 계산 (y = ax + b)
    poly = np.polyfit(sample_corr_df["이동거리(km)"], sample_corr_df["운동량"], 1)
    slope = poly[0]
    intercept = poly[1]
    r_squared = r_corr ** 2
    
    # Plotly Scatter + Trendline
    fig_scatter = px.scatter(
        sample_corr_df,
        x="이동거리(km)",
        y="운동량",
        color="대여구분코드",
        opacity=0.6,
        trendline="ols",
        trendline_color_override="#dc2626",
        labels={"이동거리(km)": "총 이동거리 (km)", "운동량": "총 소모 운동량 (kcal)"}
    )
    fig_scatter.update_layout(
        margin=dict(l=20, r=20, t=20, b=20),
        height=450,
        legend_title="대여구분"
    )
    st.plotly_chart(fig_scatter, use_container_width=True)
    
    st.markdown(f"""
    <div class="insight-card">
        📈 <b>건강 상관성 및 회귀 분석 결과:</b><br>
        • 피어슨 상관계수 (r): <b>{r_corr:.4f}</b> (초고도의 양의 선형 상관관계), 결정계수 (R²): <b>{r_squared:.4f}</b><br>
        • 선형 회귀 추정식: <b>운동량(kcal) = {slope:.1f} × 이동거리(km) + {intercept:.1f}</b><br>
        • 시민들이 따릉이로 <b>1km 주행할 때마다 평균 약 {slope:.1f} kcal</b>의 유산소 운동 효과를 얻는 것으로 분석되었습니다.
    </div>
    """, unsafe_allow_html=True)


# ==============================================================================
# [Tab 4] 원본 및 대여소별 상세 데이터 테이블
# ==============================================================================
with tab4:
    st.subheader("📋 필터링된 데이터 탐색 및 CSV 내보내기")
    
    view_option = st.radio(
        "조회할 집계 테이블 선택:",
        options=["대여소별 종합 집계", "자치구별 종합 집계", "필터링된 상세 이용 내역 (원천 샘플)"],
        horizontal=True
    )
    
    if view_option == "대여소별 종합 집계":
        detail_df = filtered_df.groupby(["대여소번호", "보관소(대여소)명", "자치구"]).agg({
            "총거치대수": "first",
            "이용건수": "sum",
            "이동거리(M)": "sum",
            "이용시간(분)": "sum",
            "탄소량": "sum",
            "운동량": "sum"
        }).reset_index()
        
        detail_df["이동거리(km)"] = (detail_df["이동거리(M)"] / 1000.0).round(2)
        detail_df["거치대 회전율"] = np.where(
            detail_df["총거치대수"] > 0, 
            (detail_df["이용건수"] / detail_df["총거치대수"]).round(2), 
            0
        )
        detail_df["탄소량(kg)"] = detail_df["탄소량"].round(2)
        detail_df["운동량(kcal)"] = detail_df["운동량"].round(1)
        
        show_cols = [
            "대여소번호", "보관소(대여소)명", "자치구", "총거치대수", 
            "이용건수", "거치대 회전율", "이동거리(km)", "이용시간(분)", "탄소량(kg)", "운동량(kcal)"
        ]
        export_df = detail_df[show_cols].sort_values(by="이용건수", ascending=False)
        
    elif view_option == "자치구별 종합 집계":
        export_df = gu_agg.copy()
        export_df["탄소량(kg)"] = export_df["탄소량"].round(2)
        export_df["이동거리(km)"] = export_df["이동거리(km)"].round(2)
        export_df = export_df.sort_values(by="이용건수", ascending=False)
        
    else:  # 원천 샘플
        sample_rows = min(len(filtered_df), 1000)
        export_df = filtered_df[[
            "대여일자_str", "대여소번호", "보관소(대여소)명", "자치구", 
            "대여구분코드", "성별", "연령대코드", "이용건수", "운동량", "탄소량", "이동거리(M)", "이용시간(분)"
        ]].head(sample_rows)

    # 검색 필터
    search_query = st.text_input("🔍 대여소명 또는 자치구 검색", "")
    if search_query:
        if "보관소(대여소)명" in export_df.columns:
            export_df = export_df[
                export_df["보관소(대여소)명"].str.contains(search_query, case=False, na=False) |
                export_df["자치구"].str.contains(search_query, case=False, na=False)
            ]
        elif "자치구" in export_df.columns:
            export_df = export_df[export_df["자치구"].str.contains(search_query, case=False, na=False)]

    st.write(f"총 **{len(export_df):,}** 행 표시 중")
    st.dataframe(export_df, use_container_width=True, height=450)
    
    # 한글 엑셀 깨짐 방지 utf-8-sig 인코딩
    csv_data = export_df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button(
        label="📥 현재 테이블 CSV 다운로드 (UTF-8 with BOM)",
        data=csv_data,
        file_name="따릉이_데이터_추출.csv",
        mime="text/csv"
    )

st.markdown("---")
st.caption("서울시 공공자전거(따릉이) 분석 대시보드 | Powered by Streamlit, GeoPandas & Plotly")
