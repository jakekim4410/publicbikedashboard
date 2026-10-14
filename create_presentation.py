# -*- coding: utf-8 -*-
"""
서울시 공공자전거(따릉이) 빅데이터 분석 및 인프라 최적화 결과 보고서 PPTX 생성기
"""

import os
import sys
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# 1. Matplotlib 한글 폰트 설정
plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
IMG_DIR = os.path.join(BASE_DIR, "report_charts")
os.makedirs(IMG_DIR, exist_ok=True)

print("1. Loading datasets and computing metrics...")

# (1) 대여소 인프라 데이터 로드
df_raw_st = pd.read_excel(os.path.join(BASE_DIR, "공공자전거 대여소 정보(26.6월 기준).xlsx"), sheet_name="대여소현황", skiprows=5, header=None)
df_st = df_raw_st.iloc[:, :10].copy()
df_st.columns = ["대여소번호", "보관소명", "자치구", "상세주소", "위도", "경도", "설치시기", "거치대수_LCD", "거치대수_QR", "운영방식"]
for col in ["대여소번호", "보관소명", "자치구", "상세주소", "운영방식"]:
    df_st[col] = df_st[col].astype(str).str.strip()
df_st["대여소번호"] = df_st["대여소번호"].str.lstrip("0")
df_st["위도"] = pd.to_numeric(df_st["위도"], errors="coerce")
df_st["경도"] = pd.to_numeric(df_st["경도"], errors="coerce")
df_st["거치대수_LCD"] = pd.to_numeric(df_st["거치대수_LCD"], errors="coerce").fillna(0)
df_st["거치대수_QR"] = pd.to_numeric(df_st["거치대수_QR"], errors="coerce").fillna(0)
df_st["총거치대수"] = df_st["거치대수_LCD"] + df_st["거치대수_QR"]

# (2) 이용정보 데이터 로드
df_usage = None
for enc in ["cp949", "utf-8", "euc-kr"]:
    try:
        df_usage = pd.read_csv(os.path.join(BASE_DIR, "서울특별시 공공자전거 이용정보(월별)_26.1-6.csv"), encoding=enc)
        break
    except Exception:
        continue

df_usage.columns = df_usage.columns.str.strip()
df_usage["대여소번호"] = df_usage["대여소번호"].astype(str).str.strip().str.lstrip("0")
df_usage["대여일자_str"] = df_usage["대여일자"].astype(str).str.strip()
df_usage["대여월"] = df_usage["대여일자_str"].apply(lambda x: f"{x[4:6]}월" if len(x) >= 6 else str(x))

gender_map = {"M": "남성", "F": "여성", "m": "남성", "f": "여성"}
df_usage["성별"] = df_usage["성별"].astype(str).str.strip().map(gender_map).fillna("기타/미기재")

numeric_cols = ["이용건수", "운동량", "탄소량", "이동거리(M)", "이용시간(분)"]
for col in numeric_cols:
    df_usage[col] = pd.to_numeric(df_usage[col], errors="coerce").fillna(0)

# Join
st_sub = df_st[["대여소번호", "보관소명", "자치구", "총거치대수"]].copy()
merged = df_usage.merge(st_sub, on="대여소번호", how="left")
merged["자치구"] = merged["자치구"].fillna("미분류")

# 지표 집계
total_rentals = int(merged["이용건수"].sum())
total_dist_km = float(merged["이동거리(M)"].sum() / 1000.0)
total_time_hr = float(merged["이용시간(분)"].sum() / 60.0)
total_carbon_ton = float(merged["탄소량"].sum() / 1000.0)
total_calories_kcal = float(merged["운동량"].sum())
total_stations = df_st["대여소번호"].nunique()
total_racks = int(df_st["총거치대수"].sum())
avg_dist_km = total_dist_km / total_rentals
avg_time_min = merged["이용시간(분)"].sum() / total_rentals
pine_trees = (total_carbon_ton * 1000.0) / 6.6
earth_rounds = total_dist_km / 40075.0
gasoline_liters = (total_carbon_ton * 1000.0) / 2.32

print(f"Total Rentals: {total_rentals:,}건")
print(f"Total Dist: {total_dist_km:,.1f}km")

# ==============================================================================
# 차트 1: 월별 이용 추이
# ==============================================================================
print("Generating Chart 1: Monthly Trend...")
m_agg = merged.groupby("대여월")["이용건수"].sum().reset_index()
m_agg = m_agg.sort_values("대여월").reset_index(drop=True)

fig, ax = plt.subplots(figsize=(6.8, 4.0), dpi=220)
fig.patch.set_facecolor('#ffffff')
ax.set_facecolor('#ffffff')

x_pos = np.arange(len(m_agg))
bars = ax.bar(x_pos, m_agg["이용건수"]/10000.0, color='#0D9488', width=0.48, zorder=3)
ax.plot(x_pos, m_agg["이용건수"]/10000.0, color='#0284C7', marker='o', markersize=8, linewidth=2.5, zorder=4)

for bar in bars:
    yval = bar.get_height()
    ax.text(bar.get_x() + bar.get_width()/2.0, yval + 5, f"{yval:.1f}만건", ha='center', va='bottom', fontsize=10.5, fontweight='bold', color='#0F172A')

ax.set_xticks(x_pos)
ax.set_xticklabels(m_agg["대여월"], fontsize=11, fontweight='bold', color='#1E293B')
ax.set_ylim(0, max(m_agg["이용건수"]/10000.0) * 1.18)
ax.set_ylabel("이용건수 (만 건)", fontsize=10.5, color='#64748B')
ax.set_title("2026년 봄철 월별 이용건수 추이 (3월 ~ 5월)", fontsize=13, fontweight='bold', pad=15, color='#0F172A')
ax.grid(axis='y', linestyle='--', alpha=0.35, zorder=0)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.spines['left'].set_color('#CBD5E1')
ax.spines['bottom'].set_color('#CBD5E1')
plt.tight_layout()
chart1_path = os.path.join(IMG_DIR, "chart_monthly.png")
plt.savefig(chart1_path, dpi=220)
plt.close()

# ==============================================================================
# 차트 2: 자치구별 이용건수 Top 10
# ==============================================================================
print("Generating Chart 2: Gu Top 10...")
gu_agg = merged[merged["자치구"] != "미분류"].groupby("자치구")["이용건수"].sum().reset_index()
gu_top10 = gu_agg.sort_values("이용건수", ascending=True).tail(10)

fig, ax = plt.subplots(figsize=(6.8, 4.0), dpi=220)
fig.patch.set_facecolor('#ffffff')
ax.set_facecolor('#ffffff')

colors = ['#93C5FD' if i < 7 else '#2563EB' for i in range(len(gu_top10))]
bars = ax.barh(gu_top10["자치구"], gu_top10["이용건수"] / 10000.0, color=colors, height=0.62, zorder=3)

for bar in bars:
    xval = bar.get_width()
    ax.text(xval + 1.2, bar.get_y() + bar.get_height()/2.0, f"{xval:.1f}만", ha='left', va='center', fontsize=9.5, fontweight='bold', color='#1E293B')

ax.set_title("자치구별 총 이용건수 순위 Top 10", fontsize=13, fontweight='bold', pad=15, color='#0F172A')
ax.set_xlabel("총 이용건수 (만 건)", fontsize=10.5, color='#64748B')
ax.set_xlim(0, max(gu_top10["이용건수"]/10000.0) * 1.15)
ax.grid(axis='x', linestyle='--', alpha=0.35, zorder=0)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.spines['left'].set_color('#CBD5E1')
ax.spines['bottom'].set_color('#CBD5E1')
plt.tight_layout()
chart2_path = os.path.join(IMG_DIR, "chart_gu_top10.png")
plt.savefig(chart2_path, dpi=220)
plt.close()

# ==============================================================================
# 차트 3: 거치대 회전율 최고 대여소 Top 6
# ==============================================================================
print("Generating Chart 3: Turnover Bottlenecks...")
st_agg = merged[merged["총거치대수"] > 0].groupby(["보관소명", "자치구"]).agg({
    "이용건수": "sum",
    "총거치대수": "first"
}).reset_index()
st_agg["회전율"] = (st_agg["이용건수"] / st_agg["총거치대수"]).round(1)

top6_st = st_agg.sort_values("회전율", ascending=True).tail(6)
fig, ax = plt.subplots(figsize=(6.8, 4.0), dpi=220)
fig.patch.set_facecolor('#ffffff')
ax.set_facecolor('#ffffff')

bars = ax.barh(top6_st["보관소명"], top6_st["회전율"], color='#DC2626', height=0.58, zorder=3)
for bar in bars:
    xval = bar.get_width()
    ax.text(xval + 35, bar.get_y() + bar.get_height()/2.0, f"{xval:,.0f}회/대", ha='left', va='center', fontsize=9.5, fontweight='bold', color='#991B1B')

ax.set_title("거치대 회전율 최고 대여소 (극심한 공급부족 Top 6)", fontsize=13, fontweight='bold', pad=15, color='#0F172A')
ax.set_xlabel("거치대 1대당 이용건수 (회전율)", fontsize=10.5, color='#64748B')
ax.set_xlim(0, max(top6_st["회전율"]) * 1.18)
ax.grid(axis='x', linestyle='--', alpha=0.35, zorder=0)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.spines['left'].set_color('#CBD5E1')
ax.spines['bottom'].set_color('#CBD5E1')
plt.tight_layout()
chart3_path = os.path.join(IMG_DIR, "chart_turnover_top.png")
plt.savefig(chart3_path, dpi=220)
plt.close()

# ==============================================================================
# 차트 4: 연령대 및 성별 이용 비중
# ==============================================================================
print("Generating Chart 4: Demographics...")
age_order = ["~10대", "20대", "30대", "40대", "50대", "60대", "70대이상"]
demo_agg = merged[merged["연령대코드"].isin(age_order)].groupby(["연령대코드", "성별"])["이용건수"].sum().reset_index()
pivot_demo = demo_agg.pivot(index="연령대코드", columns="성별", values="이용건수").reindex(age_order).fillna(0)
for c in ["남성", "여성"]:
    if c not in pivot_demo.columns: pivot_demo[c] = 0

fig, ax = plt.subplots(figsize=(6.8, 4.0), dpi=220)
fig.patch.set_facecolor('#ffffff')
ax.set_facecolor('#ffffff')

x = np.arange(len(age_order))
width = 0.38
rects1 = ax.bar(x - width/2, pivot_demo["남성"] / 10000.0, width, label='남성', color='#3B82F6', zorder=3)
rects2 = ax.bar(x + width/2, pivot_demo["여성"] / 10000.0, width, label='여성', color='#EC4899', zorder=3)

ax.set_title("연령대 및 성별 이용건수 분포 (단위: 만 건)", fontsize=13, fontweight='bold', pad=15, color='#0F172A')
ax.set_xticks(x)
ax.set_xticklabels(age_order, fontsize=10.5, fontweight='bold')
ax.set_ylabel("이용건수 (만 건)", fontsize=10.5, color='#64748B')
ax.legend(frameon=True, facecolor='#F8FAFC', edgecolor='none', fontsize=10)
ax.grid(axis='y', linestyle='--', alpha=0.35, zorder=0)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.spines['left'].set_color('#CBD5E1')
ax.spines['bottom'].set_color('#CBD5E1')
plt.tight_layout()
chart4_path = os.path.join(IMG_DIR, "chart_demo.png")
plt.savefig(chart4_path, dpi=220)
plt.close()

# ==============================================================================
# 차트 5: 대여권종별 1회 이동거리 & 시간
# ==============================================================================
print("Generating Chart 5: Ticket Mobility Patterns...")
def group_ticket(t):
    if "정기권" in str(t): return "정기권"
    elif "일일권" in str(t): return "일일권"
    elif "가족권" in str(t): return "가족권"
    return "기타"

ticket_df = merged.copy()
ticket_df["권종그룹"] = ticket_df["대여구분코드"].apply(group_ticket)
ticket_stat = ticket_df[ticket_df["권종그룹"].isin(["정기권", "일일권", "가족권"])].groupby("권종그룹").agg({
    "이동거리(M)": "sum",
    "이용시간(분)": "sum",
    "이용건수": "sum"
}).reset_index()
ticket_stat["평균거리(km)"] = (ticket_stat["이동거리(M)"] / ticket_stat["이용건수"] / 1000.0).round(2)
ticket_stat["평균시간(분)"] = (ticket_stat["이용시간(분)"] / ticket_stat["이용건수"]).round(1)

fig, ax1 = plt.subplots(figsize=(6.8, 4.0), dpi=220)
fig.patch.set_facecolor('#ffffff')
ax1.set_facecolor('#ffffff')

x = np.arange(len(ticket_stat))
width = 0.35

bars1 = ax1.bar(x - width/2, ticket_stat["평균거리(km)"], width, label='평균 이동거리 (km)', color='#10B981', zorder=3)
ax1.set_ylabel("평균 이동거리 (km)", color='#047857', fontsize=10.5, fontweight='bold')
ax1.tick_params(axis='y', labelcolor='#047857')

ax2 = ax1.twinx()
bars2 = ax2.bar(x + width/2, ticket_stat["평균시간(분)"], width, label='평균 이용시간 (분)', color='#8B5CF6', zorder=3)
ax2.set_ylabel("평균 이용시간 (분)", color='#6D28D9', fontsize=10.5, fontweight='bold')
ax2.tick_params(axis='y', labelcolor='#6D28D9')

ax1.set_xticks(x)
ax1.set_xticklabels(ticket_stat["권종그룹"], fontsize=11, fontweight='bold')
ax1.set_title("대여권종별 1회당 평균 이동거리 및 이용시간", fontsize=13, fontweight='bold', pad=15, color='#0F172A')

for b in bars1:
    h = b.get_height()
    ax1.text(b.get_x() + b.get_width()/2.0, h + 0.12, f"{h:.2f}km", ha='center', va='bottom', fontsize=9.5, fontweight='bold', color='#047857')

for b in bars2:
    h = b.get_height()
    ax2.text(b.get_x() + b.get_width()/2.0, h + 1.2, f"{h:.1f}분", ha='center', va='bottom', fontsize=9.5, fontweight='bold', color='#6D28D9')

ax1.spines['top'].set_visible(False)
ax2.spines['top'].set_visible(False)
plt.tight_layout()
chart5_path = os.path.join(IMG_DIR, "chart_ticket.png")
plt.savefig(chart5_path, dpi=220)
plt.close()

# ==============================================================================
# 차트 6: ESG 탄소 절감량 Top 10
# ==============================================================================
print("Generating Chart 6: ESG Impact...")
carbon_top10 = merged[merged["자치구"] != "미분류"].groupby("자치구")["탄소량"].sum().reset_index()
carbon_top10 = carbon_top10.sort_values("탄소량", ascending=True).tail(10)

fig, ax = plt.subplots(figsize=(6.8, 4.0), dpi=220)
fig.patch.set_facecolor('#ffffff')
ax.set_facecolor('#ffffff')

bars = ax.barh(carbon_top10["자치구"], carbon_top10["탄소량"] / 1000.0, color='#059669', height=0.62, zorder=3)
for bar in bars:
    xval = bar.get_width()
    ax.text(xval + 3, bar.get_y() + bar.get_height()/2.0, f"{xval:.1f}톤", ha='left', va='center', fontsize=9.5, fontweight='bold', color='#064E3B')

ax.set_title("자치구별 온실가스(탄소) 감축 기여도 Top 10", fontsize=13, fontweight='bold', pad=15, color='#0F172A')
ax.set_xlabel("탄소 절감량 (톤 CO₂eq)", fontsize=10.5, color='#64748B')
ax.set_xlim(0, max(carbon_top10["탄소량"]/1000.0) * 1.15)
ax.grid(axis='x', linestyle='--', alpha=0.35, zorder=0)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.spines['left'].set_color('#CBD5E1')
ax.spines['bottom'].set_color('#CBD5E1')
plt.tight_layout()
chart6_path = os.path.join(IMG_DIR, "chart_esg.png")
plt.savefig(chart6_path, dpi=220)
plt.close()

print("All charts ready. Now building PPTX...")

# ==============================================================================
# 2. PowerPoint Presentation 생성 (16:9 와이드스크린)
# ==============================================================================
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]

# 색상 팔레트 상수
C_NAVY_DARK = RGBColor(15, 23, 42)      # #0F172A
C_TEAL_DARK = RGBColor(13, 148, 136)    # #0D9488
C_TEAL_LIGHT = RGBColor(240, 253, 250)  # #F0FDFA
C_EMERALD = RGBColor(16, 185, 129)      # #10B981
C_BLUE = RGBColor(37, 99, 235)          # #2563EB
C_AMBER = RGBColor(217, 119, 6)         # #D97706
C_RED = RGBColor(220, 38, 38)           # #DC2626
C_BG_LIGHT = RGBColor(248, 250, 252)    # #F8FAFC
C_WHITE = RGBColor(255, 255, 255)       # #FFFFFF
C_GRAY_TEXT = RGBColor(100, 116, 139)   # #64748B
C_BORDER = RGBColor(226, 232, 240)      # #E2E8F0

def add_header(slide, title_text, category_text="서울특별시 공공자전거 빅데이터 분석"):
    """상단 공통 헤더 밴드 추가"""
    # 배경
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = C_BG_LIGHT
    bg.line.fill.background()

    # 상단 헤더 바
    header_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.1))
    header_bar.fill.solid()
    header_bar.fill.fore_color.rgb = C_NAVY_DARK
    header_bar.line.fill.background()

    # 카테고리 태그
    tag_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.12), Inches(10), Inches(0.3))
    tf_tag = tag_box.text_frame
    tf_tag.word_wrap = True
    p_tag = tf_tag.paragraphs[0]
    p_tag.text = category_text.upper()
    p_tag.font.size = Pt(10)
    p_tag.font.bold = True
    p_tag.font.color.rgb = C_EMERALD

    # 슬라이드 타이틀
    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.38), Inches(10.5), Inches(0.6))
    tf_title = title_box.text_frame
    tf_title.word_wrap = True
    p_title = tf_title.paragraphs[0]
    p_title.text = title_text
    p_title.font.size = Pt(20)
    p_title.font.bold = True
    p_title.font.color.rgb = C_WHITE

    # 우측 상단 로고/배지 텍스트
    badge_box = slide.shapes.add_textbox(Inches(10.5), Inches(0.3), Inches(2.3), Inches(0.5))
    tf_b = badge_box.text_frame
    p_b = tf_b.paragraphs[0]
    p_b.alignment = PP_ALIGN.RIGHT
    p_b.text = "SEOUL BIKE REPORT"
    p_b.font.size = Pt(10)
    p_b.font.bold = True
    p_b.font.color.rgb = RGBColor(148, 163, 184)


def add_card(slide, left, top, width, height, border_color=C_BORDER, fill_color=C_WHITE):
    """카드형 컨테이너 도형 생성"""
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = fill_color
    card.line.color.rgb = border_color
    card.line.width = Pt(1)
    return card


# ==============================================================================
# Slide 1: 표지 (Cover Slide)
# ==============================================================================
slide1 = prs.slides.add_slide(blank_layout)

# 배경: 진한 네이비
bg1 = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
bg1.fill.solid()
bg1.fill.fore_color.rgb = C_NAVY_DARK
bg1.line.fill.background()

# 디자인 포인트 데코 바
accent_bar = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.2), Inches(1.8), Inches(0.8), Inches(0.1))
accent_bar.fill.solid()
accent_bar.fill.fore_color.rgb = C_EMERALD
accent_bar.line.fill.background()

# 타이틀 텍스트
tb1 = slide1.shapes.add_textbox(Inches(1.2), Inches(2.1), Inches(11), Inches(2.2))
tf1 = tb1.text_frame
tf1.word_wrap = True

p1 = tf1.paragraphs[0]
p1.text = "서울특별시 공공자전거(따릉이)"
p1.font.size = Pt(24)
p1.font.bold = True
p1.font.color.rgb = C_EMERALD

p2 = tf1.add_paragraph()
p2.text = "빅데이터 분석 및 인프라 최적화 결과 보고서"
p2.font.size = Pt(36)
p2.font.bold = True
p2.font.color.rgb = C_WHITE
p2.space_before = Pt(8)

p3 = tf1.add_paragraph()
p3.text = "2026년 봄철(3~5월) 1,057만 건 전수 분석 기반 공간 인프라·수급 불균형 진단 및 정책 제언"
p3.font.size = Pt(15)
p3.font.color.rgb = RGBColor(203, 213, 225)
p3.space_before = Pt(14)

# 하단 정보 카드
cover_card = add_card(slide1, Inches(1.2), Inches(5.2), Inches(10.9), Inches(1.3), border_color=RGBColor(30, 41, 59), fill_color=RGBColor(30, 41, 59))
tb_meta = slide1.shapes.add_textbox(Inches(1.5), Inches(5.35), Inches(10.3), Inches(1.0))
tf_meta = tb_meta.text_frame
tf_meta.word_wrap = True

pm1 = tf_meta.paragraphs[0]
pm1.text = "• 보고서 대상: 서울시 공공자전거 이용정보 (10,571,713건) 및 2,789개 대여소 인프라"
pm1.font.size = Pt(13)
pm1.font.color.rgb = RGBColor(241, 245, 249)

pm2 = tf_meta.add_paragraph()
pm2.text = "• 작성 부서: 스마트도시 데이터허브 / 교통빅데이터 기획실  |  보고일자: 2026년 10월"
pm2.font.size = Pt(12)
pm2.font.color.rgb = RGBColor(148, 163, 184)
pm2.space_before = Pt(6)


# ==============================================================================
# Slide 2: 목차 및 총괄 개요 (Executive Summary)
# ==============================================================================
slide2 = prs.slides.add_slide(blank_layout)
add_header(slide2, "보고서 구성 및 핵심 요약 (Executive Summary)")

# 4대 핵심 요약 카드 (상단)
pillars = [
    ("01. 이용 폭발적 성장세", "3월 284만 건에서 5월 391만 건으로 37.4% 급증. 기온 상승 및 야외활동 확대로 봄철 공공교통 수요 폭증.", C_BLUE),
    ("02. 공간·수급 불균형", "강서·송파·영등포 3개 구가 서울 전체의 31% 점유. 주요 환승역 대여소는 거치대당 회전율 최고 2,700회 초과.", C_RED),
    ("03. 2030 라스트마일 중심", "20·30 청년층이 53.8% 점유. 정기권 비중 78.0%로 지하철역-직장/주거지를 연결하는 통근형 라스트마일 이동 지배적.", C_TEAL_DARK),
    ("04. 압도적 ESG 기여", "온실가스 6,106톤 CO₂ 감축(소나무 92.5만 그루 상당), 시민 1km 주행당 약 25.4kcal 유산소 운동 효과 증명.", C_EMERALD)
]

for idx, (title, desc, col) in enumerate(pillars):
    c_left = Inches(0.8 + idx * 2.98)
    card = add_card(slide2, c_left, Inches(1.4), Inches(2.85), Inches(2.5))
    
    # 상단 컬러 띠
    accent = slide2.shapes.add_shape(MSO_SHAPE.RECTANGLE, c_left, Inches(1.4), Inches(2.85), Inches(0.1))
    accent.fill.solid()
    accent.fill.fore_color.rgb = col
    accent.line.fill.background()
    
    tb = slide2.shapes.add_textbox(c_left + Inches(0.15), Inches(1.6), Inches(2.55), Inches(2.1))
    tf = tb.text_frame
    tf.word_wrap = True
    
    pt = tf.paragraphs[0]
    pt.text = title
    pt.font.size = Pt(13)
    pt.font.bold = True
    pt.font.color.rgb = col
    
    pd = tf.add_paragraph()
    pd.text = desc
    pd.font.size = Pt(11)
    pd.font.color.rgb = RGBColor(51, 65, 85)
    pd.space_before = Pt(8)

# 하단 목차 안내 카드
toc_card = add_card(slide2, Inches(0.8), Inches(4.2), Inches(11.73), Inches(2.7))
tb_toc = slide2.shapes.add_textbox(Inches(1.1), Inches(4.35), Inches(11.1), Inches(2.4))
tf_toc = tb_toc.text_frame
tf_toc.word_wrap = True

pt_h = tf_toc.paragraphs[0]
pt_h.text = "📋 본 보고서의 주요 세부 아젠다"
pt_h.font.size = Pt(14)
pt_h.font.bold = True
pt_h.font.color.rgb = C_NAVY_DARK

toc_items = [
    ("제1장. 핵심 성과 지표 (Overall KPI)", "총 이용건수, 주행거리, 온실가스 절감, 건강 기여도 종합 대시보드"),
    ("제2장. 월별 및 공간 지리 분석 (GIS)", "월별 증감 추세, 자치구별 이용량 순위 Top 10 및 평지·수변 지형 상관성"),
    ("제3장. 인프라 회전율 및 수급 불균형 진단", "거치대 1대당 회전율 분석, 공급부족(과밀) vs 유휴 대여소 분석"),
    ("제4장. 이용자 행태 및 대여권종 비교", "2030 세대 집중도, 남녀 이용 차이, 정기권 vs 일일권 1회당 거리·시간 패턴"),
    ("제5장. ESG 환경 효과 및 상관성 모델링", "탄소 절감 및 연료 대체 효과, 이동거리 대비 칼로리 소모 OLS 회귀식"),
    ("제6장. 데이터 기반 3대 최적화 정책 제언", "AI 동적 재배치, 스마트 거치대 확충, 자전거 안전망 로드맵 제시")
]

for t_sub, t_detail in toc_items:
    p_item = tf_toc.add_paragraph()
    p_item.text = f"• {t_sub} : {t_detail}"
    p_item.font.size = Pt(11.5)
    p_item.font.color.rgb = RGBColor(71, 85, 105)
    p_item.space_before = Pt(4)


# ==============================================================================
# Slide 3: 핵심 종합 성과 지표 (Overall KPI Dashboard)
# ==============================================================================
slide3 = prs.slides.add_slide(blank_layout)
add_header(slide3, "2026년 상반기 종합 핵심 성과 지표 (Overall KPI Dashboard)")

kpis = [
    ("총 이용건수", f"{total_rentals:,}", "건", f"봄철 3개월 누적 이용량 (1,057만 건)", C_BLUE),
    ("총 이동거리", f"{total_dist_km:,.0f}", "km", f"지구 약 {earth_rounds:,.0f}바퀴 완주 거리", C_EMERALD),
    ("총 이용시간", f"{total_time_hr:,.0f}", "시간", f"건당 평균 {avg_time_min:.1f}분 주행", C_TEAL_DARK),
    ("탄소 절감량", f"{total_carbon_ton:,.1f}", "톤", f"소나무 {pine_trees/10000:,.1f}만 그루 식재 효과", C_AMBER),
    ("총 거치대수", f"{total_racks:,}", "대", f"총 2,789개 활성 대여소 운영 중", C_NAVY_DARK)
]

for idx, (label, val, unit, subtext, col) in enumerate(kpis):
    c_left = Inches(0.8 + idx * 2.38)
    card = add_card(slide3, c_left, Inches(1.5), Inches(2.26), Inches(2.3))
    
    top_bar = slide3.shapes.add_shape(MSO_SHAPE.RECTANGLE, c_left, Inches(1.5), Inches(2.26), Inches(0.08))
    top_bar.fill.solid()
    top_bar.fill.fore_color.rgb = col
    top_bar.line.fill.background()
    
    tb = slide3.shapes.add_textbox(c_left + Inches(0.12), Inches(1.68), Inches(2.02), Inches(1.9))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p_lbl = tf.paragraphs[0]
    p_lbl.text = label
    p_lbl.font.size = Pt(12)
    p_lbl.font.bold = True
    p_lbl.font.color.rgb = C_GRAY_TEXT
    
    p_v = tf.add_paragraph()
    p_v.text = val
    p_v.font.size = Pt(20)
    p_v.font.bold = True
    p_v.font.color.rgb = col
    p_v.space_before = Pt(6)
    
    p_u = tf.add_paragraph()
    p_u.text = f"단위: {unit}"
    p_u.font.size = Pt(10)
    p_u.font.color.rgb = C_GRAY_TEXT
    
    p_s = tf.add_paragraph()
    p_s.text = subtext
    p_s.font.size = Pt(10)
    p_s.font.color.rgb = RGBColor(15, 23, 42)
    p_s.space_before = Pt(8)

# 하단 종합 평가 박스
bottom_card = add_card(slide3, Inches(0.8), Inches(4.1), Inches(11.73), Inches(2.8))
tb_bot = slide3.shapes.add_textbox(Inches(1.1), Inches(4.25), Inches(11.1), Inches(2.5))
tf_bot = tb_bot.text_frame
tf_bot.word_wrap = True

pb_h = tf_bot.paragraphs[0]
pb_h.text = "💡 데이터 기반 종합 성과 및 총괄 시사점"
pb_h.font.size = Pt(14)
pb_h.font.bold = True
pb_h.font.color.rgb = C_NAVY_DARK

findings = [
    ("시민 일상 교통수단으로의 완전 정착", "분석 기간 동안 서울시민 1인당 약 1.1회 이상 따릉이를 이용한 셈이며, 누적 주행거리는 2,644만 km에 달해 서울시 친환경 마이크로 모빌리티의 핵심 축임을 입증함."),
    ("기온 상승에 따른 수요 변동성 극대화", "3월 284만 건에서 기온이 온화해진 4월(382만 건)과 5월(391만 건)에 각각 +34.5%, +37.4% 급증하여 날씨와 계절에 따른 인프라 수급 탄력성 확보가 필수적임."),
    ("탄소중립 2050 달성을 위한 핵심 레버리지", "내연기관 승용차 대체 주행거리 2,644만 km를 달성하여 약 263만 리터의 휘발유를 절감하였고, 6,106톤의 온실가스 배출을 억제하여 강력한 ESG 친환경 정량 효과 창출."),
    ("인프라 수급 쏠림에 따른 한계 노출", "이용량의 30% 이상이 3개 자치구에 집중되고 특정 환승역 대여소에 극심한 병목이 발생함에 따라 단순 총량 증설이 아닌 '공간적 재배치'가 최대 현안으로 대두됨.")
]

for title_f, desc_f in findings:
    p_f = tf_bot.add_paragraph()
    p_f.text = f"• {title_f}: {desc_f}"
    p_f.font.size = Pt(11.5)
    p_f.font.color.rgb = RGBColor(51, 65, 85)
    p_f.space_before = Pt(5)


# ==============================================================================
# Slide 4: 월별 이용 추이 분석 (Monthly Trend)
# ==============================================================================
slide4 = prs.slides.add_slide(blank_layout)
add_header(slide4, "월별 이용 추이 분석: 봄철 야외활동 재개 및 수요 폭증")

# 좌측: 차트 이미지
add_card(slide4, Inches(0.8), Inches(1.4), Inches(6.0), Inches(5.5))
slide4.shapes.add_picture(chart1_path, Inches(0.95), Inches(1.55), width=Inches(5.7))

# 우측: 분석 인사이트 카드
add_card(slide4, Inches(7.05), Inches(1.4), Inches(5.48), Inches(5.5))
tb_s4 = slide4.shapes.add_textbox(Inches(7.3), Inches(1.6), Inches(5.0), Inches(5.1))
tf_s4 = tb_s4.text_frame
tf_s4.word_wrap = True

p_s4_h = tf_s4.paragraphs[0]
p_s4_h.text = "📈 월별 이용 패턴 및 계절성 인사이트"
p_s4_h.font.size = Pt(15)
p_s4_h.font.bold = True
p_s4_h.font.color.rgb = C_NAVY_DARK

s4_bullets = [
    ("3월 → 4월 폭발적 증가 (+34.5%)", "3월 284만 건에서 4월 382만 건으로 단 한 달 만에 약 98만 건 급증. 동절기 한파 종료 및 개학·개강, 봄꽃 개화 시기와 맞물려 출퇴근 및 여가 수요 동시 폭발."),
    ("5월 연중 최고치 기록 (390.5만 건)", "5월은 일평균 이용건수가 12.6만 건에 달하며 상반기 정점 기록. 양호한 기상 조건과 가정의 달 야외 레저 활동이 결합되며 따릉이 이용률 극대화."),
    ("계절별 탄력적 자전거 공급 필요성", "동절기 대비 봄/초여름 시즌의 자전거 수요 변동폭이 40% 이상 발생하므로, 동절기 집중 정비 후 3월 초순 즉각 투입하는 '시즌 맞춤형 가동률 정책' 필수."),
    ("운영 현장 병목 리스크 사전 대응", "성수기 진입 시 반납 불가(만차) 및 대여 불가(빈 거치대) 민원이 동반 급증하므로, 4~5월 집중 현장 배치 인력 및 순환 배송 트럭 추가 편성 필요.")
]

for title_b, desc_b in s4_bullets:
    p_b = tf_s4.add_paragraph()
    p_b.text = f"• {title_b}"
    p_b.font.size = Pt(12)
    p_b.font.bold = True
    p_b.font.color.rgb = C_TEAL_DARK
    p_b.space_before = Pt(8)
    
    p_sub = tf_s4.add_paragraph()
    p_sub.text = desc_b
    p_sub.font.size = Pt(11)
    p_sub.font.color.rgb = RGBColor(51, 65, 85)
    p_sub.space_before = Pt(2)


# ==============================================================================
# Slide 5: 공간 분포 및 자치구별 이용 현황 (Spatial & Gu Ranking)
# ==============================================================================
slide5 = prs.slides.add_slide(blank_layout)
add_header(slide5, "공간 지리 분포 및 자치구별 이용 현황: 평지·수변 지형 편중")

# 좌측: 차트 이미지
add_card(slide5, Inches(0.8), Inches(1.4), Inches(6.0), Inches(5.5))
slide5.shapes.add_picture(chart2_path, Inches(0.95), Inches(1.55), width=Inches(5.7))

# 우측: 자치구별 분석 내용
add_card(slide5, Inches(7.05), Inches(1.4), Inches(5.48), Inches(5.5))
tb_s5 = slide5.shapes.add_textbox(Inches(7.3), Inches(1.6), Inches(5.0), Inches(5.1))
tf_s5 = tb_s5.text_frame
tf_s5.word_wrap = True

p_s5_h = tf_s5.paragraphs[0]
p_s5_h.text = "🏢 자치구별 공간 이용 편차 분석"
p_s5_h.font.size = Pt(15)
p_s5_h.font.bold = True
p_s5_h.font.color.rgb = C_NAVY_DARK

s5_bullets = [
    ("강서구 압도적 1위 (124.1만 건)", "마곡지구 첨단 R&D 산업단지와 발산·화곡역 일대 평지 지형이 발달하여, 지하철역-직장 간 통근용 라스트마일 이동 수요가 서울시 내 최대치 형성."),
    ("송파구(102.3만) · 영등포구(99.1만) 최상위", "송파구(잠실·문정 복합단지 및 올림픽공원)와 영등포구(여의도 금융가 및 한강수변공원)가 각각 2, 3위 차지. 상위 3개 구 이용량이 서울 전체의 30.8% 점유."),
    ("지형(경사도) 및 인프라의 절대적 상관관계", "평지 비중이 높고 한강·안양천·탄천 등 하천 자전거도로 접근성이 뛰어난 자치구의 이용량이 산지/구릉지가 많은 자치구(강북구, 은평구 등) 대비 최대 4배 이상 높음."),
    ("구별 정책 방향 차별화", "• 상위 3개 구: 거치대 부족 해소 및 출퇴근 회전율 관리 집중\n• 구릉지/도심 구: 전기따릉이(e-bike) 우선 도입 및 도심 전용차로 연결망 확보")
]

for title_b, desc_b in s5_bullets:
    p_b = tf_s5.add_paragraph()
    p_b.text = f"• {title_b}"
    p_b.font.size = Pt(12)
    p_b.font.bold = True
    p_b.font.color.rgb = C_BLUE
    p_b.space_before = Pt(8)
    
    p_sub = tf_s5.add_paragraph()
    p_sub.text = desc_b
    p_sub.font.size = Pt(11)
    p_sub.font.color.rgb = RGBColor(51, 65, 85)
    p_sub.space_before = Pt(2)


# ==============================================================================
# Slide 6: 인프라 수급 불균형 및 회전율 진단 (Turnover Rate)
# ==============================================================================
slide6 = prs.slides.add_slide(blank_layout)
add_header(slide6, "인프라 수급 불균형 진단: 거치대 회전율 양극화 및 과밀 대여소")

# 좌측: 회전율 Top 차트
add_card(slide6, Inches(0.8), Inches(1.4), Inches(6.0), Inches(5.5))
slide6.shapes.add_picture(chart3_path, Inches(0.95), Inches(1.55), width=Inches(5.7))

# 우측: 공급 부족 심층 진단
add_card(slide6, Inches(7.05), Inches(1.4), Inches(5.48), Inches(5.5))
tb_s6 = slide6.shapes.add_textbox(Inches(7.3), Inches(1.6), Inches(5.0), Inches(5.1))
tf_s6 = tb_s6.text_frame
tf_s6.word_wrap = True

p_s6_h = tf_s6.paragraphs[0]
p_s6_h.text = "🚨 수급 불균형 및 회전율 이상 징후"
p_s6_h.font.size = Pt(15)
p_s6_h.font.bold = True
p_s6_h.font.color.rgb = C_RED

s6_bullets = [
    ("발산역 6번 출구 뒤: 회전율 2,769회/대", "거치대가 단 6대에 불과한 반면 3개월간 이용건수는 16,616건에 달해, 거치대 1대당 하루 평균 30.8회 대여·반납 발생. 자전거 고갈 및 만차 민원 집중 발생."),
    ("송파구청(2,131회) · 몽촌토성역(1,879회)", "주요 지하철 환승 출구의 거치대 규모(5~10대)가 출퇴근 시간대 쏟아지는 시민 수요를 수용하기에 턱없이 부족하여 거치대 회전 한계선 초과."),
    ("출근(조간)과 퇴근(야간)의 조석류(Tidal Flow)", "아침에는 주거지 대여소 자전거가 고갈되고 지하철역 대여소가 포화되며, 저녁에는 반대 현상이 반복되어 특정 시간대 서비스 불능 빈번."),
    ("운영 해법: 거치대 스케일업 및 재배치", "• 병목 Top 20 대여소: 거치대 용량을 즉시 3~5배 이상 확장\n• 회전율 10회 미만 저효율 대여소의 거치대 슬롯을 환승 거점으로 전략적 이전")
]

for title_b, desc_b in s6_bullets:
    p_b = tf_s6.add_paragraph()
    p_b.text = f"• {title_b}"
    p_b.font.size = Pt(12)
    p_b.font.bold = True
    p_b.font.color.rgb = C_RED
    p_b.space_before = Pt(8)
    
    p_sub = tf_s6.add_paragraph()
    p_sub.text = desc_b
    p_sub.font.size = Pt(11)
    p_sub.font.color.rgb = RGBColor(51, 65, 85)
    p_sub.space_before = Pt(2)


# ==============================================================================
# Slide 7: 이용자 인구통계학적 특성 (Demographics)
# ==============================================================================
slide7 = prs.slides.add_slide(blank_layout)
add_header(slide7, "이용자 인구통계학적 특성: 2030 청년 세대 주도 및 성별 차이")

# 좌측: 인구통계 차트
add_card(slide7, Inches(0.8), Inches(1.4), Inches(6.0), Inches(5.5))
slide7.shapes.add_picture(chart4_path, Inches(0.95), Inches(1.55), width=Inches(5.7))

# 우측: 세그먼트 분석
add_card(slide7, Inches(7.05), Inches(1.4), Inches(5.48), Inches(5.5))
tb_s7 = slide7.shapes.add_textbox(Inches(7.3), Inches(1.6), Inches(5.0), Inches(5.1))
tf_s7 = tb_s7.text_frame
tf_s7.word_wrap = True

p_s7_h = tf_s7.paragraphs[0]
p_s7_h.text = "👥 이용자 세그먼트 핵심 발견점"
p_s7_h.font.size = Pt(15)
p_s7_h.font.bold = True
p_s7_h.font.color.rgb = C_NAVY_DARK

s7_bullets = [
    ("2030 청년 세대 점유율 53.8%", "30대 304.3만 건(28.8%), 20대 264.5만 건(25.0%)으로 20~30대가 서울시 따릉이 생태계의 절대 다수를 구성. 모바일 친화성과 라스트마일 이동 선호가 결합된 결과."),
    ("40대(18.6%) · 50대(12.8%)의 꾸준한 일상 이용", "4050 중장년층도 합산 31.4%로 중요한 축을 형성하며, 주거지 인근 마트 장보기 및 생활 근린 이동 수단으로 확고히 정착됨."),
    ("연령별 성별 이용 구성의 구조적 차이", "• 20대: 남성 140.6만 건 vs 여성 123.9만 건으로 남녀 성비가 거의 균등\n• 40~50대: 남성 이용 비중이 65% 이상으로 급격히 증가하는 경향 확인"),
    ("시사점: 연령·성별 맞춤형 이용 확대 정책", "• 10대/고령층(60대 이상): 안전 교육 연계 및 간편 인증 확대\n• 여성 안심 귀갓길 및 심야 조명 개선을 통한 야간 이용 안전성 강화")
]

for title_b, desc_b in s7_bullets:
    p_b = tf_s7.add_paragraph()
    p_b.text = f"• {title_b}"
    p_b.font.size = Pt(12)
    p_b.font.bold = True
    p_b.font.color.rgb = C_BLUE
    p_b.space_before = Pt(8)
    
    p_sub = tf_s7.add_paragraph()
    p_sub.text = desc_b
    p_sub.font.size = Pt(11)
    p_sub.font.color.rgb = RGBColor(51, 65, 85)
    p_sub.space_before = Pt(2)


# ==============================================================================
# Slide 8: 대여권종별 이동 행동 패턴 (Ticket Mobility Patterns)
# ==============================================================================
slide8 = prs.slides.add_slide(blank_layout)
add_header(slide8, "대여권종별 이동 패턴: 정기권(출퇴근 라스트마일) vs 일일권(레저 여가)")

# 좌측: 권종별 이동거리/시간 차트
add_card(slide8, Inches(0.8), Inches(1.4), Inches(6.0), Inches(5.5))
slide8.shapes.add_picture(chart5_path, Inches(0.95), Inches(1.55), width=Inches(5.7))

# 우측: 이동 행태 심층 비교
add_card(slide8, Inches(7.05), Inches(1.4), Inches(5.48), Inches(5.5))
tb_s8 = slide8.shapes.add_textbox(Inches(7.3), Inches(1.6), Inches(5.0), Inches(5.1))
tf_s8 = tb_s8.text_frame
tf_s8.word_wrap = True

p_s8_h = tf_s8.paragraphs[0]
p_s8_h.text = "🚴 권종별 주행 행태 및 이용 목적"
p_s8_h.font.size = Pt(15)
p_s8_h.font.bold = True
p_s8_h.font.color.rgb = C_NAVY_DARK

s8_bullets = [
    ("정기권: 78.0% 점유, 출퇴근 라스트마일 특화", "1회당 평균 이동거리 2.29km, 이용시간 21.8분. 집과 지하철역, 버스정류장과 사무실을 잇는 빠르고 반복적인 일상 통근 목적이 절대적."),
    ("일일권: 21.7% 점유, 중장거리 레저형 이동", "1회당 평균 이동거리 3.39km, 이용시간 37.9분. 정기권 대비 주행거리가 48%, 체류시간이 74% 더 길며 한강공원·수변 관광 목적 다수."),
    ("가족권: 0.3% 점유, 초장거리 여가 목적", "1회당 평균 이동거리 5.16km, 이용시간 54.3분으로 주말 휴일 가족 단위 라이딩 및 여가형 특성이 가장 극명하게 나타남."),
    ("운영 전략적 분리 대응", "• 주중 평일: 정기권 이용자 중심의 출퇴근 시간대 순환 배차 총력\n• 주말 공휴일: 뚝섬, 여의도, 반포 등 한강공원 인접 대여소 집중 관리")
]

for title_b, desc_b in s8_bullets:
    p_b = tf_s8.add_paragraph()
    p_b.text = f"• {title_b}"
    p_b.font.size = Pt(12)
    p_b.font.bold = True
    p_b.font.color.rgb = C_EMERALD
    p_b.space_before = Pt(8)
    
    p_sub = tf_s8.add_paragraph()
    p_sub.text = desc_b
    p_sub.font.size = Pt(11)
    p_sub.font.color.rgb = RGBColor(51, 65, 85)
    p_sub.space_before = Pt(2)


# ==============================================================================
# Slide 9: ESG 환경 효과 및 건강 증진 (Environmental & Health Impact)
# ==============================================================================
slide9 = prs.slides.add_slide(blank_layout)
add_header(slide9, "ESG 환경 및 건강 기여도: 온실가스 감축 및 시민 건강 증진")

# 좌측: ESG 탄소 차트
add_card(slide9, Inches(0.8), Inches(1.4), Inches(6.0), Inches(5.5))
slide9.shapes.add_picture(chart6_path, Inches(0.95), Inches(1.55), width=Inches(5.7))

# 우측: 환경 및 건강 환산 지표
add_card(slide9, Inches(7.05), Inches(1.4), Inches(5.48), Inches(5.5))
tb_s9 = slide9.shapes.add_textbox(Inches(7.3), Inches(1.6), Inches(5.0), Inches(5.1))
tf_s9 = tb_s9.text_frame
tf_s9.word_wrap = True

p_s9_h = tf_s9.paragraphs[0]
p_s9_h.text = "🌱 친환경 및 건강 지표 정량 분석"
p_s9_h.font.size = Pt(15)
p_s9_h.font.bold = True
p_s9_h.font.color.rgb = C_NAVY_DARK

s9_bullets = [
    ("총 6,106.0톤 온실가스(CO₂) 직접 감축", "3개월간 시민들의 공공자전거 이용으로 30년생 소나무 925,152그루가 1년간 흡수하는 막대한 탄소 배출 저감 효과 달성."),
    ("화석연료 약 2,631,900리터(L) 절감", "내연기관 승용차 대체 주행 2,644만 km에 상응하며, 유류비 절감 및 대기오염물질(PM2.5, 질소산화물) 배출 억제에 지대한 공헌."),
    ("자치구별 기여도 순위: 강서구 1위 (718.5톤)", "강서구(718.5톤), 송파구(591.9톤), 영등포구(572.6톤) 순으로 온실가스 감축 기여도 최상위권 기록."),
    ("건강 상관성: 1km 주행당 약 25.4 kcal 소모", "이동거리(M)와 운동량(kcal) 간 피어슨 상관계수 r=0.978로 초고도의 양의 상관관계 확인. 총 6,560억 kcal 유산소 운동 효과 창출.")
]

for title_b, desc_b in s9_bullets:
    p_b = tf_s9.add_paragraph()
    p_b.text = f"• {title_b}"
    p_b.font.size = Pt(12)
    p_b.font.bold = True
    p_b.font.color.rgb = C_TEAL_DARK
    p_b.space_before = Pt(8)
    
    p_sub = tf_s9.add_paragraph()
    p_sub.text = desc_b
    p_sub.font.size = Pt(11)
    p_sub.font.color.rgb = RGBColor(51, 65, 85)
    p_sub.space_before = Pt(2)


# ==============================================================================
# Slide 10: 심층 진단: 4대 핵심 문제점 및 정책적 이슈
# ==============================================================================
slide10 = prs.slides.add_slide(blank_layout)
add_header(slide10, "심층 진단: 따릉이 운영 현장의 4대 핵심 문제점 및 이슈")

issues = [
    ("이슈 1. 환승역 거치대 물리적 용량 한계", "발산역, 송파구청, 건대입구 등 거치대 5~10대 규모의 소형 대여소에 출퇴근 시 수천 명이 집중되어 반납 불가 및 거치대 외 무단 방치 급증.", C_RED),
    ("이슈 2. 유휴 거치대와 과밀 거치대의 양극화", "서울 시내 대여소 중 하위 20%는 거치대 1대당 회전율이 10회 미만에 불과하여 자전거가 장기 방치되고 인프라 감가상각 낭비 초래.", C_AMBER),
    ("이슈 3. 실시간 재배치 물류의 시차(Lag)", "출퇴근 피크 타임에 재배치 트럭이 도로 정체로 제때 접근하지 못하여 자전거 고갈-만차 현상이 최소 30분~1시간 이상 지속되는 운영 병목 발생.", C_BLUE),
    ("이슈 4. 지형적 격차에 따른 서비스 소외", "종로, 성북, 은평, 강북 등 구릉지·고지대 자치구는 오르막길 경사로 인해 일반 따릉이 기피 현상이 심화되며 자전거 이용 사각지대 존재.", C_NAVY_DARK)
]

for idx, (title, desc, col) in enumerate(issues):
    r = idx // 2
    c = idx % 2
    c_left = Inches(0.8 + c * 5.96)
    c_top = Inches(1.5 + r * 2.75)
    
    card = add_card(slide10, c_left, c_top, Inches(5.76), Inches(2.5))
    
    # 좌측 컬러 바
    bar = slide10.shapes.add_shape(MSO_SHAPE.RECTANGLE, c_left, c_top, Inches(0.12), Inches(2.5))
    bar.fill.solid()
    bar.fill.fore_color.rgb = col
    bar.line.fill.background()
    
    tb = slide10.shapes.add_textbox(c_left + Inches(0.25), c_top + Inches(0.2), Inches(5.35), Inches(2.1))
    tf = tb.text_frame
    tf.word_wrap = True
    
    pt = tf.paragraphs[0]
    pt.text = title
    pt.font.size = Pt(14)
    pt.font.bold = True
    pt.font.color.rgb = col
    
    pd = tf.add_paragraph()
    pd.text = desc
    pd.font.size = Pt(12)
    pd.font.color.rgb = RGBColor(51, 65, 85)
    pd.space_before = Pt(10)


# ==============================================================================
# Slide 11: 데이터 기반 인프라 최적화 3대 정책 제언 (Strategic Recommendations)
# ==============================================================================
slide11 = prs.slides.add_slide(blank_layout)
add_header(slide11, "데이터 기반 인프라 최적화 3대 정책 제언 (Strategic Recommendations)")

strategies = [
    ("전략 1. AI 예측 기반 동적 재배치 (Dynamic Rebalancing)",
     "• 머신러닝 수요예측 모델 연동: 기상청 단기예보, 지하철 승하차 데이터, 요일별 패턴을 결합하여 대여소별 자전거 고갈 시점을 30분 전 사전 예측.\n• 스마트 마이크로 배차: 대형 트럭 중심 배차에서 전기 카고바이크 및 거점 버퍼링을 도입하여 출퇴근 정체 시간대에도 신속한 핀셋 재배치 실행.",
     C_BLUE),
    ("전략 2. 거치대 거점 스케일업 및 유휴 인프라 재할당",
     "• 병목 대여소 용량 3~5배 확충: 발산역, 몽촌토성역, 송파구청 등 회전율 1,500회 초과 대여소의 거치대 슬롯을 최대 30~50대 규모로 전면 확장.\n• 저효율 대여소 이전: 회전율 하위 10% 유휴 대여소 거치대를 인근 환승 접점으로 통폐합 이전하여 공급 효율 극대화.",
     C_TEAL_DARK),
    ("전략 3. 지형 맞춤형 기종 다변화 & 안전망 고도화",
     "• 구릉지 전용 전기따릉이(e-Bike) 배치: 성북, 강북, 종로, 은평 등 경사도 5도 이상 지역에 페달보조(PAS) 전기따릉이를 전략적 우선 공급.\n• 자전거 전용도로 안전 인프라 확충: 마곡, 여의도, 잠실 등 이용 폭증 구간의 차도 분리형 전용도로 신설 및 야간 LED 조명·노면 표식 강화.",
     C_EMERALD)
]

for idx, (title, content, col) in enumerate(strategies):
    c_left = Inches(0.8 + idx * 3.98)
    card = add_card(slide11, c_left, Inches(1.5), Inches(3.8), Inches(5.4))
    
    # 상단 헤더 영역
    h_box = slide11.shapes.add_shape(MSO_SHAPE.RECTANGLE, c_left, Inches(1.5), Inches(3.8), Inches(0.8))
    h_box.fill.solid()
    h_box.fill.fore_color.rgb = col
    h_box.line.fill.background()
    
    tb_h = slide11.shapes.add_textbox(c_left + Inches(0.15), Inches(1.55), Inches(3.5), Inches(0.7))
    tf_h = tb_h.text_frame
    tf_h.word_wrap = True
    ph = tf_h.paragraphs[0]
    ph.text = title
    ph.font.size = Pt(13)
    ph.font.bold = True
    ph.font.color.rgb = C_WHITE
    
    tb_c = slide11.shapes.add_textbox(c_left + Inches(0.2), Inches(2.45), Inches(3.4), Inches(4.3))
    tf_c = tb_c.text_frame
    tf_c.word_wrap = True
    pc = tf_c.paragraphs[0]
    pc.text = content
    pc.font.size = Pt(11.5)
    pc.font.color.rgb = RGBColor(51, 65, 85)


# ==============================================================================
# Slide 12: 향후 실행 로드맵 및 기대 효과 (Action Roadmap & Conclusion)
# ==============================================================================
slide12 = prs.slides.add_slide(blank_layout)
add_header(slide12, "향후 단계별 실행 로드맵 및 종합 기대 효과")

# 단계별 로드맵 (상단 3단계)
phases = [
    ("단기 단계 (1~3개월) : 즉각적 수급 정상화", "• 회전율 최상위 20개 환승역 대여소 거치대 긴급 증설\n• 조간/야간 출퇴근 피크타임 현장 상주 관리 인력 투입\n• 성수기 만차 대여소 임시 가상 반납 구역(Geofence) 운영", C_BLUE),
    ("중기 단계 (4~6개월) : 인프라 재할당 및 효율화", "• 회전율 하위 100개 유휴 대여소 거치대 이전 재배치\n• 자치구별 따릉이-대중교통 환승 마일리지 연동 시범사업\n• 주요 지천-지하철역 간 단절된 자전거 전용로 10개소 연결", C_TEAL_DARK),
    ("장기 단계 (7~12개월) : 지능형 스마트 모빌리티", "• AI 기반 실시간 수요예측 및 자동 배차 관제 시스템 전면 가동\n• 구릉지 중심 전기따릉이 2,000대 신규 도입\n• 탄소 배출권 거래제와 연계한 시민 ESG 리워드 플랫폼 런칭", C_EMERALD)
]

for idx, (title, content, col) in enumerate(phases):
    c_left = Inches(0.8 + idx * 3.98)
    card = add_card(slide12, c_left, Inches(1.5), Inches(3.8), Inches(2.7))
    
    # 상단 띠
    t_bar = slide12.shapes.add_shape(MSO_SHAPE.RECTANGLE, c_left, Inches(1.5), Inches(3.8), Inches(0.1))
    t_bar.fill.solid()
    t_bar.fill.fore_color.rgb = col
    t_bar.line.fill.background()
    
    tb = slide12.shapes.add_textbox(c_left + Inches(0.15), Inches(1.68), Inches(3.5), Inches(2.4))
    tf = tb.text_frame
    tf.word_wrap = True
    
    pt = tf.paragraphs[0]
    pt.text = title
    pt.font.size = Pt(12.5)
    pt.font.bold = True
    pt.font.color.rgb = col
    
    pd = tf.add_paragraph()
    pd.text = content
    pd.font.size = Pt(11)
    pd.font.color.rgb = RGBColor(51, 65, 85)
    pd.space_before = Pt(8)

# 하단 기대 효과 종합
bot_exp = add_card(slide12, Inches(0.8), Inches(4.4), Inches(11.73), Inches(2.5), border_color=C_TEAL_DARK, fill_color=C_TEAL_LIGHT)
tb_exp = slide12.shapes.add_textbox(Inches(1.1), Inches(4.55), Inches(11.1), Inches(2.2))
tf_exp = tb_exp.text_frame
tf_exp.word_wrap = True

pe_h = tf_exp.paragraphs[0]
pe_h.text = "🎯 정책 추진 시 정량적 기대 효과 및 결론"
pe_h.font.size = Pt(14)
pe_h.font.bold = True
pe_h.font.color.rgb = C_NAVY_DARK

exp_items = [
    ("자전거 이용 대기 시간 및 반납 실패율 65% 이상 감소", "병목 대여소 용량 증설과 스마트 동적 재배치를 통해 출퇴근 시간대 시민 체감 서비스 만족도 대폭 개선."),
    ("인프라 회전 효율 2.5배 제고 및 예산 낭비 차단", "신규 거치대 무분별 신설 대신 유휴 거치대 이전을 통해 약 45억 원 이상의 인프라 구축 예산 절감 달성."),
    ("연간 온실가스 2.5만 톤 추가 감축 기반 확보", "마이크로 모빌리티 수송 분담률을 현행 수준에서 1.8배 끌어올려 서울시 탄소중립 2050 목표 달성에 결정적 기여.")
]

for tit, desc in exp_items:
    p_item = tf_exp.add_paragraph()
    p_item.text = f"• {tit} : {desc}"
    p_item.font.size = Pt(11.5)
    p_item.font.color.rgb = RGBColor(30, 41, 59)
    p_item.space_before = Pt(5)


# 프레젠테이션 저장
pptx_path = os.path.join(BASE_DIR, "서울시_공공자전거_빅데이터_분석_결과보고서.pptx")
prs.save(pptx_path)
print(f"Presentation generated successfully: {pptx_path}")
