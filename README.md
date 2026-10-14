# 🚲 서울시 공공자전거(따릉이) 빅데이터 분석 및 인프라 최적화 대시보드
> **Seoul Public Bike Big Data Analytics & Spatial Optimization Dashboard**

서울시 공공자전거(따릉이) 1,057만 건의 이용 빅데이터와 2,789개 대여소 인프라, GIS 공간정보(SHP)를 융합하여 인프라 운영 효율성, 이용자 행동 패턴, ESG 환경 기여도를 다각도로 분석하는 통합 데이터 플랫폼입니다.

---

## 🌟 주요 기능 (Key Features)

### 1. 🗺️ 공간 지리정보(GIS) 및 대여소 회전율 분석
- **Choropleth 지도**: 자치구별 따릉이 총 이용량 공간 분포 시각화 (Folium)
- **거치대 회전율 지표 ($Turnover = \frac{이용건수}{거치대수}$)** 산출
- **수급 불균형 진단**: 출퇴근 시간대 극심한 자전거 고갈/만차가 발생하는 **공급부족 대여소 Top 10** vs **유휴 대여소 Top 10** 자동 판별
- **마커 클러스터링**: 상위 이용 대여소 위치 및 운영 현황 공간 시각화

### 2. 👥 이용자 행동 패턴 및 인구통계학적 인사이트
- **연령대 × 성별 교차 분석**: 2030 세대(53.8%) 중심의 주력 이용 패턴 및 연령대별 성별 구성비 히트맵
- **대여권종별 이동 특성**:
  - **정기권 (78.0%)**: 평균 2.29km, 21.8분 (지하철 연계형 출퇴근 라스트마일)
  - **일일권 (21.7%)**: 평균 3.39km, 37.9분 (한강공원·지천 중심 레저 여가)
  - **가족권 (0.3%)**: 평균 5.16km, 54.3분 (주말 장거리 라이딩)

### 3. 🌱 ESG 환경 기여도 및 건강 효과 분석
- **온실가스 감축량**: 총 6,106.0톤 CO₂eq 감축 (30년생 소나무 약 92.5만 그루 연간 흡수량 상당)
- **화석연료 대체**: 내연기관 승용차 대체 주행 2,644만 km, 휘발유 약 263만 리터 절감
- **유산소 운동량 OLS 선형 회귀 분석**: 피어슨 상관계수 $r = 0.978$ (1km 주행 시 약 25.4 kcal 소모)

### 4. 📊 16:9 파워포인트(PPTX) 자동 보고서 생성기
- `create_presentation.py` 실행 시 분석 결과, 고해상도 차트 6종, 정책 제언이 포함된 12장의 발표용 PPTX 자동 생성

---

## 🛠️ 기술 스택 (Tech Stack)

- **Language**: Python 3.12+
- **Interactive Web App**: Streamlit
- **Data Engineering**: Pandas, NumPy, OpenPyXL
- **Spatial GIS**: GeoPandas, Shapely, PyProj, Folium, Streamlit-Folium
- **Visualization**: Plotly, Matplotlib
- **Statistical Modeling**: Statsmodels, SciPy
- **Automated Reporting**: python-pptx

---

## 🚀 빠른 시작 (Getting Started)

### 1. 레포지토리 클론
```bash
git clone https://github.com/jakekim4410/publicbikedashboard.git
cd publicbikedashboard
```

### 2. 필수 라이브러리 설치
```bash
pip install -r requirements.txt
```

### 3. 로컬 대시보드 실행
```bash
streamlit run app.py
```
브라우저에서 `http://localhost:8501`로 접속합니다.

### 4. PPT 보고서 생성
```bash
python create_presentation.py
```
`서울시_공공자전거_빅데이터_분석_결과보고서.pptx` 파일이 자동으로 생성됩니다.

---

## 📂 프로젝트 구조 (Project Structure)

```
.
├── app.py                                   # Streamlit 메인 대시보드 애플리케이션
├── create_presentation.py                   # PPTX 결과보고서 자동 생성 파이프라인
├── generate_charts.py                       # 고화질 분석 차트 생성 스크립트
├── requirements.txt                         # 패키지 의존성 정의
├── .gitignore                               # Git 제외 대상 설정
├── README.md                                # 프로젝트 안내 문서
├── 서울시_공공자전거_빅데이터_분석_결과보고서.pptx  # 16:9 와이드스크린 결과 보고서
├── report_charts/                           # 시각화 차트 이미지 에셋 (PNG)
├── sig_20230729/                            # 서울시 행정구역 GIS 공간 데이터 (SHP)
├── 지리정보.xlsx                            # 서울시 자치구 코드 매핑 테이블
├── 공공자전거 대여소 정보(26.6월 기준).xlsx   # 따릉이 대여소 마스터 인프라 데이터
└── 서울특별시 공공자전거 이용정보(월별)_26.1-6.csv # 공공자전거 월별 이용 빅데이터
```

---

## 📄 라이선스 (License)
본 프로젝트는 공공데이터포털 및 서울 열린데이터광장의 공공누리(제1유형)에 따라 개방된 공공데이터를 기반으로 제작되었습니다.
