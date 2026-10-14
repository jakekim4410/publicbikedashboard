# -*- coding: utf-8 -*-
"""
서울시 공공자전거(따릉이) 분석 결과 PPTX 보고서 자동 생성 스크립트
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

# 1. 폰트 및 스타일 설정
plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
IMG_DIR = os.path.join(BASE_DIR, "report_charts")
os.makedirs(IMG_DIR, exist_ok=True)

print("Starting data loading and analysis...")

# 2. 데이터 로드 및 정제
# (1) 대여소 데이터
df_raw_st = pd.read_excel(os.path.join(BASE_DIR, "공공자전거 대여소 정보(26.6월 기준).xlsx"), sheet_name="대여소현황", skiprows=5, header=None)
df_st = df_raw_st.iloc[:, :10].copy()
df_st.columns = ["대여소번호", "보관소(대여소)명", "자치구", "상세주소", "위도", "경도", "설치시기", "거치대수_LCD", "거치대수_QR", "운영방식"]
for col in ["대여소번호", "보관소(대여소)명", "자치구", "상세주소", "운영방식"]:
    df_st[col] = df_st[col].astype(str).str.strip()
df_st["대여소번호"] = df_st["대여소번호"].str.lstrip("0")
df_st["위도"] = pd.to_numeric(df_st["위도"], errors="coerce")
df_st["경도"] = pd.to_numeric(df_st["경도"], errors="coerce")
df_st["거치대수_LCD"] = pd.to_numeric(df_st["거치대수_LCD"], errors="coerce").fillna(0)
df_st["거치대수_QR"] = pd.to_numeric(df_st["거치대수_QR"], errors="coerce").fillna(0)
df_st["총거치대수"] = df_st["거치대수_LCD"] + df_st["거치대수_QR"]

# (2) 이용정보 데이터
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

# Join with station info
st_sub = df_st[["대여소번호", "보관소(대여소)명", "자치구", "총거치대수"]].copy()
merged = df_usage.merge(st_sub, on="대여소번호", how="left")
merged["자치구"] = merged["자치구"].fillna("미분류")

# 지표 요약
total_rentals = int(merged["이용건수"].sum())
total_dist_km = float(merged["이동거리(M)"].sum() / 1000.0)
total_time_hr = float(merged["이용시간(분)"].sum() / 60.0)
total_carbon_ton = float(merged["탄소량"].sum() / 1000.0)
total_calories_kcal = float(merged["운동량"].sum())
total_stations = df_st["대여소번호"].nunique()
total_racks = int(df_st["총거치대수"].sum())
avg_dist_km = total_dist_km / total_rentals if total_rentals > 0 else 0
avg_time_min = (merged["이용시간(분)"].sum() / total_rentals) if total_rentals > 0 else 0
pine_trees = (total_carbon_ton * 1000.0) / 6.6
earth_rounds = total_dist_km / 40075.0

print(f"Total Rentals: {total_rentals:,} 건")
print(f"Total Dist: {total_dist_km:,.1f} km (지구 {earth_rounds:.1f}바퀴)")
print(f"Total Carbon: {total_carbon_ton:,.1f} 톤 (소나무 {pine_trees/10000:.1f}만 그루)")
print(f"Stations: {total_stations:,} 개소, 거치대: {total_racks:,} 대")

# ==============================================================================
# 차트 1: 월별 이용량 추이 (상반기 성장세)
# ==============================================================================
monthly_agg = merged.groupby("대여월")["이용건수"].sum().reset_index()
# 순서 정렬
month_order = ["01월", "02월", "03월", "04월", "05월", "06월"]
monthly_agg["sort_key"] = monthly_agg["대여월"].apply(lambda x: month_order.index(x) if x in month_order else 99)
monthly_agg = monthly_agg.sort_values("sort_key").reset_index(drop=True)

fig, ax = plt.subplots(figsize=(7.5, 4.2), dpi=200)
fig.patch.set_facecolor('#ffffff')
ax.set_facecolor('#ffffff')

bars = ax.bar(monthly_agg["대여월"], monthly_agg["이용건수"] / 10000.0, color='#0D9488', width=0.55, edgecolor='none', zorder=3)
ax.plot(monthly_agg["대여월"], monthly_agg["이용건수"] / 10000.0, color='#0284C7', marker='o', markersize=7, linewidth=2.5, zorder=4)

for bar in bars:
    yval = bar.get_height()
    ax.text(bar.get_x() + bar.get_width()/2.0, yval + 3, f"{yval:.1f}만건", ha='center', va='bottom', fontsize=10, fontweight='bold', color='#0F172A')

ax.set_title("2026년 상반기 월별 따릉이 이용건수 추이 (단위: 만 건)", fontsize=13, fontweight='bold', pad=15, color='#0F172A')
ax.set_ylabel("이용건수 (만 건)", fontsize=10, color='#64748B')
ax.grid(axis='y', linestyle='--', alpha=0.3, zorder=0)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.spines['left'].set_color('#CBD5E1')
ax.spines['bottom'].set_color('#CBD5E1')
plt.tight_layout()
chart1_path = os.path.join(IMG_DIR, "chart_monthly.png")
plt.savefig(chart1_path, dpi=200)
plt.close()

# ==============================================================================
# 차트 2: 자치구별 이용건수 Top 10
# ==============================================================================
gu_agg = merged[merged["자치구"] != "미분류"].groupby("자치구")["이용건수"].sum().reset_index()
gu_top10 = gu_agg.sort_values("이용건수", ascending=True).tail(10)

fig, ax = plt.subplots(figsize=(7.5, 4.2), dpi=200)
fig.patch.set_facecolor('#ffffff')
ax.set_facecolor('#ffffff')

colors = ['#93C5FD' if i < 7 else '#2563EB' for i in range(len(gu_top10))]
bars = ax.barh(gu_top10["자치구"], gu_top10["이용건수"] / 10000.0, color=colors, height=0.6, zorder=3)

for bar in bars:
    xval = bar.get_width()
    ax.text(xval + 1, bar.get_y() + bar.get_height()/2.0, f"{xval:.1f}만", ha='left', va='center', fontsize=9.5, fontweight='bold', color='#1E293B')

ax.set_title("서울시 자치구별 따릉이 총 이용건수 Top 10 (단위: 만 건)", fontsize=13, fontweight='bold', pad=15, color='#0F172A')
ax.set_xlabel("총 이용건수 (만 건)", fontsize=10, color='#64748B')
ax.grid(axis='x', linestyle='--', alpha=0.3, zorder=0)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.spines['left'].set_color('#CBD5E1')
ax.spines['bottom'].set_color('#CBD5E1')
plt.tight_layout()
chart2_path = os.path.join(IMG_DIR, "chart_gu_top10.png")
plt.savefig(chart2_path, dpi=200)
plt.close()

# ==============================================================================
# 차트 3: 거치대 회전율 Top 5 (공급부족) vs Bottom 5 (유휴)
# ==============================================================================
st_agg = merged[merged["총거치대수"] > 0].groupby(["보관소(대여소)명", "자치구"]).agg({
    "이용건수": "sum",
    "총거치대수": "first"
}).reset_index()
st_agg["회전율"] = (st_agg["이용건수"] / st_agg["총거치대수"]).round(1)

top5_st = st_agg.sort_values("회전율", ascending=True).tail(5)
fig, ax = plt.subplots(figsize=(7.5, 4.2), dpi=200)
fig.patch.set_facecolor('#ffffff')
ax.set_facecolor('#ffffff')

bars = ax.barh(top5_st["보관소(대여소)명"], top5_st["회전율"], color='#EF4444', height=0.55, zorder=3)
for bar in bars:
    xval = bar.get_width()
    ax.text(xval + 5, bar.get_y() + bar.get_height()/2.0, f"{xval:.0f}회/대", ha='left', va='center', fontsize=9.5, fontweight='bold', color='#991B1B')

ax.set_title("🚨 거치대 1대당 회전율 최고 대여소 (자전거 공급부족 Top 5)", fontsize=13, fontweight='bold', pad=15, color='#0F172A')
ax.set_xlabel("회전율 (이용건수 / 총 거치대수)", fontsize=10, color='#64748B')
ax.grid(axis='x', linestyle='--', alpha=0.3, zorder=0)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.spines['left'].set_color('#CBD5E1')
ax.spines['bottom'].set_color('#CBD5E1')
plt.tight_layout()
chart3_path = os.path.join(IMG_DIR, "chart_turnover_top.png")
plt.savefig(chart3_path, dpi=200)
plt.close()

# ==============================================================================
# 차트 4: 연령대별 & 성별 이용 비중
# ==============================================================================
age_order = ["~10대", "20대", "30대", "40대", "50대", "60대", "70대이상"]
demo_agg = merged[merged["연령대코드"].isin(age_order)].groupby(["연령대코드", "성별"])["이용건수"].sum().reset_index()

pivot_demo = demo_agg.pivot(index="연령대코드", columns="성별", values="이용건수").reindex(age_order).fillna(0)
for c in ["남성", "여성"]:
    if c not in pivot_demo.columns:
        pivot_demo[c] = 0

fig, ax = plt.subplots(figsize=(7.5, 4.2), dpi=200)
fig.patch.set_facecolor('#ffffff')
ax.set_facecolor('#ffffff')

x = np.arange(len(age_order))
width = 0.38
rects1 = ax.bar(x - width/2, pivot_demo["남성"] / 10000.0, width, label='남성', color='#3B82F6', zorder=3)
rects2 = ax.bar(x + width/2, pivot_demo["여성"] / 10000.0, width, label='여성', color='#EC4899', zorder=3)

ax.set_title("연령대 및 성별 따릉이 이용건수 분포 (단위: 만 건)", fontsize=13, fontweight='bold', pad=15, color='#0F172A')
ax.set_xticks(x)
ax.set_xticklabels(age_order, fontsize=10)
ax.set_ylabel("이용건수 (만 건)", fontsize=10, color='#64748B')
ax.legend(frameon=True, facecolor='#F8FAFC', edgecolor='none', fontsize=10)
ax.grid(axis='y', linestyle='--', alpha=0.3, zorder=0)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.spines['left'].set_color('#CBD5E1')
ax.spines['bottom'].set_color('#CBD5E1')
plt.tight_layout()
chart4_path = os.path.join(IMG_DIR, "chart_demo.png")
plt.savefig(chart4_path, dpi=200)
plt.close()

# ==============================================================================
# 차트 5: 대여권종별 1회 이동거리 & 시간 비교
# ==============================================================================
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

fig, ax1 = plt.subplots(figsize=(7.5, 4.2), dpi=200)
fig.patch.set_facecolor('#ffffff')
ax1.set_facecolor('#ffffff')

x = np.arange(len(ticket_stat))
width = 0.35

bars1 = ax1.bar(x - width/2, ticket_stat["평균거리(km)"], width, label='평균 이동거리 (km)', color='#10B981', zorder=3)
ax1.set_ylabel("평균 이동거리 (km)", color='#047857', fontsize=10, fontweight='bold')
ax1.tick_params(axis='y', labelcolor='#047857')

ax2 = ax1.twinx()
bars2 = ax2.bar(x + width/2, ticket_stat["평균시간(분)"], width, label='평균 이용시간 (분)', color='#8B5CF6', zorder=3)
ax2.set_ylabel("평균 이용시간 (분)", color='#6D28D9', fontsize=10, fontweight='bold')
ax2.tick_params(axis='y', labelcolor='#6D28D9')

ax1.set_xticks(x)
ax1.set_xticklabels(ticket_stat["권종그룹"], fontsize=11, fontweight='bold')
ax1.set_title("대여권종별 1회당 평균 이동거리 vs 이용시간 비교", fontsize=13, fontweight='bold', pad=15, color='#0F172A')

for b in bars1:
    h = b.get_height()
    ax1.text(b.get_x() + b.get_width()/2.0, h + 0.15, f"{h:.2f}km", ha='center', va='bottom', fontsize=9.5, fontweight='bold', color='#047857')

for b in bars2:
    h = b.get_height()
    ax2.text(b.get_x() + b.get_width()/2.0, h + 1.5, f"{h:.1f}분", ha='center', va='bottom', fontsize=9.5, fontweight='bold', color='#6D28D9')

ax1.spines['top'].set_visible(False)
ax2.spines['top'].set_visible(False)
plt.tight_layout()
chart5_path = os.path.join(IMG_DIR, "chart_ticket.png")
plt.savefig(chart5_path, dpi=200)
plt.close()

# ==============================================================================
# 차트 6: 자치구별 탄소 절감량 Top 10 (ESG)
# ==============================================================================
carbon_top10 = merged[merged["자치구"] != "미분류"].groupby("자치구")["탄소량"].sum().reset_index()
carbon_top10 = carbon_top10.sort_values("탄소량", ascending=True).tail(10)

fig, ax = plt.subplots(figsize=(7.5, 4.2), dpi=200)
fig.patch.set_facecolor('#ffffff')
ax.set_facecolor('#ffffff')

bars = ax.barh(carbon_top10["자치구"], carbon_top10["탄소량"] / 1000.0, color='#059669', height=0.6, zorder=3)
for bar in bars:
    xval = bar.get_width()
    ax.text(xval + 2, bar.get_y() + bar.get_height()/2.0, f"{xval:.1f}톤", ha='left', va='center', fontsize=9.5, fontweight='bold', color='#064E3B')

ax.set_title("자치구별 온실가스(탄소) 감축 기여도 Top 10 (단위: 톤 CO₂)", fontsize=13, fontweight='bold', pad=15, color='#0F172A')
ax.set_xlabel("탄소 절감량 (톤)", fontsize=10, color='#64748B')
ax.grid(axis='x', linestyle='--', alpha=0.3, zorder=0)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.spines['left'].set_color('#CBD5E1')
ax.spines['bottom'].set_color('#CBD5E1')
plt.tight_layout()
chart6_path = os.path.join(IMG_DIR, "chart_esg.png")
plt.savefig(chart6_path, dpi=200)
plt.close()

print("All charts generated successfully!")
