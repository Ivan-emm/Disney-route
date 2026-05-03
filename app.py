# ============================================================
# 迪士尼路径规划可视化系统 - 后端服务
# 适配 Render 云端部署版本（使用相对路径 + 0.0.0.0 监听）
# ============================================================

import os
import math
import re
import json
import heapq
import itertools
import pandas as pd
import numpy as np
from pathlib import Path
from difflib import get_close_matches
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

# ============================================================
# 0. 全局配置 - 使用当前脚本所在目录的相对路径
# ============================================================

# 获取当前脚本所在目录
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
TEMPLATE_DIR = os.path.join(BASE_DIR, "templates")

# 数据文件路径（云端用）
REALTIME_PATH = os.path.join(DATA_DIR, "disney_2026_05_01_hourly_wait_clean.xlsx")
ATTR_PATH = os.path.join(DATA_DIR, "修正餐饮购物0.6后的8维向量数据_20260502_231148.xlsx")
HIST_PATH = os.path.join(DATA_DIR, "上海迪士尼_QueueTimes_20240301_20260301.xlsx")
INFO_PATH = os.path.join(DATA_DIR, "最终项目数据表_中英文对应_填充版3.2.xlsx")

TARGET_DATE = "2026-05-01"
QUEUE_DEADLINE = "21:30"
ENTRANCE_NAME = "上海迪士尼入口"

DIM_COLS = ["刺激", "沉浸", "互动", "休闲", "见面会", "演出", "餐饮", "购物"]
CORE_DIM_COLS = ["刺激", "沉浸", "互动", "休闲", "见面会", "演出"]

ETA_TIME = 2.0
NF_FOR_SATISFACTION = 2
NS_FOR_SATISFACTION = 1

PARK_CLOSE_TIME = "22:00"
LUNCH_WINDOW = ("11:00", "13:00")
DINNER_WINDOW = ("17:00", "19:00")
SHOP_TIME_MIN = 30
MAX_SHOP_TIMES = 3

AREA_GRAPH = {
    "米奇大街": {"奇想花园": 408, "探险岛": 529, "明日世界": 717},
    "奇想花园": {"米奇大街": 408, "梦幻世界": 260, "宝藏湾": 482, "探险岛": 236, "疯狂动物城": 463, "玩具总动员": 546},
    "梦幻世界": {"奇想花园": 260, "玩具总动员": 406, "疯狂动物城": 408},
    "玩具总动员": {"梦幻世界": 406, "明日世界": 244, "奇想花园": 546},
    "明日世界": {"玩具总动员": 244, "米奇大街": 717},
    "宝藏湾": {"奇想花园": 482, "探险岛": 312},
    "探险岛": {"米奇大街": 529, "奇想花园": 236, "宝藏湾": 312},
    "疯狂动物城": {"奇想花园": 463, "梦幻世界": 408},
}

mu_dict = {
    "亲子": {"刺激": 0.20, "沉浸": 0.50, "互动": 0.90, "休闲": 0.80,
             "见面会": 0.95, "演出": 0.70, "餐饮": 0.80, "购物": 0.80},
    "情侣": {"刺激": 0.85, "沉浸": 0.75, "互动": 0.80, "休闲": 0.35,
             "见面会": 0.70, "演出": 0.80, "餐饮": 0.80, "购物": 0.90},
    "普通": {"刺激": 0.65, "沉浸": 0.55, "互动": 0.70, "休闲": 0.55,
             "见面会": 0.85, "演出": 0.60, "餐饮": 0.80, "购物": 0.70},
}

K_dict = {
    "亲子": {"刺激": 1, "沉浸": 3, "互动": 5, "休闲": 4,
             "见面会": 5, "演出": 3, "餐饮": 4, "购物": 1},
    "情侣": {"刺激": 5, "沉浸": 4, "互动": 2, "休闲": 1,
             "见面会": 3, "演出": 4, "餐饮": 3, "购物": 5},
    "普通": {"刺激": 3, "沉浸": 3, "互动": 3, "休闲": 3,
             "见面会": 3, "演出": 3, "餐饮": 2, "购物": 3},
}

manual_en_to_cn_raw = {
    "Become Iron Man at the Marvel Universe": "漫威英雄总部",
    "Marvel Universe": "漫威英雄总部",
    "Challenge Trails at Camp Discovery": "古迹探索营",
    "Vista Trail at Camp Discovery": "古迹探索营",
    "Junior Explorers Camp": "小勇者营地",
    "Pixar Adventurous Journey": "皮克斯奇旅",
    "Space Chat With Stitch": "太空幸会史迪奇",
    "TRON Lightcycle Power Run": "创极速光轮",
    "TRON Lightcycle Power Run - Presented by Chevrolet": "创极速光轮",
    "TRON Lightcycle Power Run - Presented by Chevrolet (Virtual)": "创极速光轮",
    "Selfie Spot with Disney Jungle Characters at Happy Circle": "丛林里的迪士尼朋友见面会",
    "Selfie Spot with Spider Man": "蜘蛛侠见面会",
    "attWestZootopiaLandEntry": "疯狂动物城・热力追踪",
}

manual_info_to_attr_raw = {
    "创极速光轮": "创极速光轮－雪佛兰呈献",
    "太空幸会史迪奇": "太空对话史迪奇",
    "疯狂动物城・热力追踪": "疯狂动物城：热力追踪",
    "丛林里的迪士尼朋友见面会": "丛林里的迪士尼朋友见面会",
    "蜘蛛侠见面会": "蜘蛛侠见面会",
}

# ============================================================
# 1. 工具函数
# ============================================================

def normalize_name(x):
    if pd.isna(x):
        return ""
    s = str(x).strip().lower()
    s = s.replace("'", "'").replace("'", "'").replace('"', '"').replace('"', '"')
    s = s.replace("–", "-").replace("—", "-")
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"[·•：:，,。！!（）()\[\]【】_\-]", "", s)
    return s

def classify_date_type(date_str):
    holiday_dates = {"2026-05-01", "2026-05-02", "2026-05-03", "2026-05-04", "2026-05-05"}
    dt = pd.to_datetime(date_str)
    ds = dt.strftime("%Y-%m-%d")
    if ds in holiday_dates:
        return "节假日"
    elif dt.weekday() >= 5:
        return "周末"
    else:
        return "工作日"

def time_str_to_datetime(date_str, hhmm):
    return pd.to_datetime(f"{date_str} {hhmm}:00")

def floor_to_hour(dt):
    dt = pd.to_datetime(dt)
    return dt.replace(minute=0, second=0, microsecond=0)

def find_col_by_keywords(df, keywords, required=True):
    if isinstance(keywords, str):
        keywords = [keywords]
    for kw in keywords:
        for c in df.columns:
            if kw in str(c):
                return c
    if required:
        raise ValueError(f"没有找到包含关键词 {keywords} 的列。现有列名：{list(df.columns)}")
    return None

def z_food(nf):
    if nf <= 0: return 0.0
    return 0.6 + 0.4 * (1 - math.exp(-0.55 * (nf - 1)))

def z_shop(ns):
    if ns <= 0: return 0.0
    return 0.6 + 0.4 * (1 - math.exp(-0.75 * (ns - 1)))

def calc_k9_from_eta(eta):
    return math.log(2) / (1 - math.exp(-(eta - 1))) ** 2

def shortest_area_distance(area_a, area_b):
    if area_a == area_b: return 0.0
    if area_a not in AREA_GRAPH or area_b not in AREA_GRAPH: return np.nan
    pq = [(0.0, area_a)]
    dist = {area_a: 0.0}
    visited = set()
    while pq:
        d, u = heapq.heappop(pq)
        if u in visited: continue
        visited.add(u)
        if u == area_b: return float(d)
        for v, w in AREA_GRAPH[u].items():
            nd = d + float(w)
            if v not in dist or nd < dist[v]:
                dist[v] = nd
                heapq.heappush(pq, (nd, v))
    return np.nan

# ============================================================
# 2. 数据加载（服务启动时执行一次）
# ============================================================

print("=" * 60)
print("  迪士尼路径规划系统 - 正在加载数据...")
print("=" * 60)

# --- 实时排队数据 ---
real_long = pd.read_excel(REALTIME_PATH, sheet_name="按小时长表_建模主表")
real_time_col = find_col_by_keywords(real_long, ["datetime", "时间"])
real_ride_col = find_col_by_keywords(real_long, ["ride_name", "项目"])
real_wait_col = find_col_by_keywords(real_long, ["avg_wait_time_min", "平均等待", "等待"])

# --- 八维向量数据 ---
attr_df = pd.read_excel(ATTR_PATH, sheet_name=0)
attr_name_col = find_col_by_keywords(attr_df, ["project_name", "项目", "名称"])

dim_col_map = {}
for dim in DIM_COLS:
    for c in attr_df.columns:
        if dim in str(c):
            dim_col_map[dim] = c
            break

# --- 历史排队数据 ---
hist_df = pd.read_excel(HIST_PATH, sheet_name="项目_日期类型汇总")
hist_date_type_col = find_col_by_keywords(hist_df, ["date_type", "日期类型"])
hist_ride_col = find_col_by_keywords(hist_df, ["ride_name", "项目"])
hist_wait_col = find_col_by_keywords(hist_df, ["avg_wait_min", "平均等待", "等待"])

# --- 项目基础信息 ---
info_df = pd.read_excel(INFO_PATH, sheet_name=0)
info_cn_col = find_col_by_keywords(info_df, ["中文名称", "中文", "项目中文"])
info_en_col = find_col_by_keywords(info_df, ["英文项目名称", "英文", "项目英文"])
info_area_col = find_col_by_keywords(info_df, ["园区"], required=False)
info_type_col = find_col_by_keywords(info_df, ["大类", "类别", "类型"], required=False)
info_height_col = find_col_by_keywords(info_df, ["身高限制", "身高"], required=False)
info_duration_col = find_col_by_keywords(info_df, ["时长", "游玩时间", "体验时间"], required=False)

DATE_TYPE = classify_date_type(TARGET_DATE)
K9 = calc_k9_from_eta(ETA_TIME)

print(f"  目标日期：{TARGET_DATE}，判定为：{DATE_TYPE}")
print(f"  K9 = {K9:.6f}")

# ============================================================
# 3. 名称映射建立
# ============================================================

manual_en_to_cn = {normalize_name(en): cn for en, cn in manual_en_to_cn_raw.items()}

project_info_full = info_df.copy()
project_info_full["项目中文名"] = project_info_full[info_cn_col].astype(str).str.strip()
project_info_full = project_info_full[
    project_info_full["项目中文名"].notna() &
    (project_info_full["项目中文名"].astype(str).str.lower() != "nan") &
    (project_info_full["项目中文名"].astype(str).str.strip() != "")
].copy()
project_info_full = project_info_full.drop_duplicates(subset=["项目中文名"]).reset_index(drop=True)

info_project_names = project_info_full["项目中文名"].dropna().astype(str).str.strip().unique().tolist()
info_norm_to_cn = {normalize_name(x): x for x in info_project_names if normalize_name(x)}
info_norm_list = list(info_norm_to_cn.keys())

exact_en_to_cn = {}
for _, row in project_info_full.iterrows():
    en = row.get(info_en_col, np.nan)
    cn = row.get("项目中文名", np.nan)
    if pd.notna(en) and str(en).strip():
        exact_en_to_cn[normalize_name(en)] = str(cn).strip()

def align_to_info_name(cn_name, allow_raw=False):
    if pd.isna(cn_name): return None
    cn_name = str(cn_name).strip()
    n = normalize_name(cn_name)
    if n in info_norm_to_cn: return info_norm_to_cn[n]
    matches = get_close_matches(n, info_norm_list, n=1, cutoff=0.72)
    if matches: return info_norm_to_cn[matches[0]]
    return cn_name if allow_raw else None

def map_project_to_info_chinese(name):
    if pd.isna(name): return None
    raw = str(name).strip()
    raw_no_virtual = raw.replace("(Virtual)", "").replace("Virtual", "").strip()
    n = normalize_name(raw)
    n2 = normalize_name(raw_no_virtual)
    if n in manual_en_to_cn: cn = manual_en_to_cn[n]
    elif n2 in manual_en_to_cn: cn = manual_en_to_cn[n2]
    elif n in exact_en_to_cn: cn = exact_en_to_cn[n]
    elif n2 in exact_en_to_cn: cn = exact_en_to_cn[n2]
    else: cn = raw_no_virtual
    return align_to_info_name(cn, allow_raw=False)

# 实时表映射
real_long = real_long.copy()
real_long["项目中文名"] = real_long[real_ride_col].apply(map_project_to_info_chinese)
real_long_valid = real_long.dropna(subset=["项目中文名"]).copy()
real_long_valid[real_time_col] = pd.to_datetime(real_long_valid[real_time_col], errors="coerce")
real_long_valid = real_long_valid.dropna(subset=[real_time_col]).copy()

real_hourly = (
    real_long_valid
    .groupby(["项目中文名", real_time_col], as_index=False)[real_wait_col]
    .mean()
    .sort_values(by=["项目中文名", real_time_col])
    .reset_index(drop=True)
)

open_projects_today = set(real_hourly["项目中文名"].unique())

# 历史表映射
hist_df = hist_df.copy()
hist_df["项目中文名"] = hist_df[hist_ride_col].apply(map_project_to_info_chinese)
hist_valid = hist_df.dropna(subset=["项目中文名"]).copy()
hist_use = hist_valid[hist_valid[hist_date_type_col] == DATE_TYPE].copy()
hist_wait_df = (
    hist_use.groupby("项目中文名", as_index=False)[hist_wait_col]
    .mean()
    .rename(columns={hist_wait_col: "历史平均等待_min"})
)
hist_wait_dict = dict(zip(hist_wait_df["项目中文名"], hist_wait_df["历史平均等待_min"]))

# 拥挤系数列
def find_crowd_col_by_date_type(date_type):
    if date_type == "节假日": date_keywords = ["节假日"]
    elif date_type in ["周末", "双休日"]: date_keywords = ["双休日", "周末"]
    else: date_keywords = ["工作日"]
    for c in project_info_full.columns:
        s = str(c)
        if any(k in s for k in date_keywords) and ("拥挤" in s) and ("系数" in s):
            return c
    for c in project_info_full.columns:
        s = str(c)
        if any(k in s for k in date_keywords) and ("拥挤" in s):
            return c
    for c in project_info_full.columns:
        s = str(c)
        if ("拥挤" in s) and ("系数" in s):
            return c
    return None

crowd_col = find_crowd_col_by_date_type(DATE_TYPE)

# 八维向量查找表
attr_lookup_df = attr_df.copy()
attr_lookup_df[attr_name_col] = attr_lookup_df[attr_name_col].astype(str).str.strip()
attr_lookup_df = attr_lookup_df.drop_duplicates(subset=[attr_name_col])
attr_lookup_df = attr_lookup_df.set_index(attr_name_col)
attr_projects_raw = attr_lookup_df.index.astype(str).tolist()

manual_info_to_attr = {normalize_name(k): v for k, v in manual_info_to_attr_raw.items()}

def map_info_name_to_attr_name(info_name):
    if pd.isna(info_name): return None
    name = str(info_name).strip()
    n = normalize_name(name)
    if n in manual_info_to_attr:
        target = manual_info_to_attr[n]
        for a in attr_projects_raw:
            if normalize_name(a) == normalize_name(target):
                return a
    for a in attr_projects_raw:
        if normalize_name(a) == n: return a
    contain_candidates = []
    for a in attr_projects_raw:
        na = normalize_name(a)
        if n and (n in na or na in n):
            contain_candidates.append(a)
    if contain_candidates:
        return sorted(contain_candidates, key=lambda x: len(str(x)))[0]
    attr_norm_list = [normalize_name(x) for x in attr_projects_raw]
    matches = get_close_matches(n, attr_norm_list, n=1, cutoff=0.72)
    if matches:
        idx = attr_norm_list.index(matches[0])
        return attr_projects_raw[idx]
    return None

project_info_full["八维表对应项目名"] = project_info_full["项目中文名"].apply(map_info_name_to_attr_name)
project_info_lookup = project_info_full.set_index("项目中文名", drop=False)

print(f"  当天开放项目数：{len(open_projects_today)}")
print(f"  历史等待项目数：{len(hist_wait_dict)}")
print(f"  拥挤系数列：{crowd_col}")
print("=" * 60)
print("  数据加载完成！")
print("=" * 60)

# ============================================================
# 4. 核心算法函数
# ============================================================

def get_project_area(project_name):
    if project_name == ENTRANCE_NAME: return ENTRANCE_NAME
    if project_name == "米奇大街": return "米奇大街"
    if project_name not in project_info_lookup.index: return None
    if info_area_col is None or info_area_col not in project_info_lookup.columns: return None
    area = project_info_lookup.loc[project_name, info_area_col]
    if pd.isna(area): return None
    return str(area).strip()

def calc_distance_m(from_loc, to_project):
    to_area = get_project_area(to_project)
    if to_area is None: return np.nan
    if from_loc == ENTRANCE_NAME:
        if to_area == "米奇大街": return 60.0
        d_area = shortest_area_distance("米奇大街", to_area)
        if pd.isna(d_area): return np.nan
        return 60.0 + d_area
    if from_loc in AREA_GRAPH or from_loc == "米奇大街":
        from_area = from_loc
    else:
        from_area = get_project_area(from_loc)
    if from_area is None: return np.nan
    if from_area == to_area: return 60.0
    return shortest_area_distance(from_area, to_area)

def get_walk_speed(user_type):
    return 50.0 if user_type == "亲子" else 60.0

def get_crowd_factor(project_name):
    if project_name not in project_info_lookup.index: return 1.0
    val = project_info_lookup.loc[project_name, crowd_col]
    if pd.isna(val): return 1.0
    try: return float(val)
    except: return 1.0

def calc_walk_time_min(from_loc, to_project, user_type):
    d = calc_distance_m(from_loc, to_project)
    if pd.isna(d): return np.nan, np.nan, np.nan
    crowd = get_crowd_factor(to_project)
    speed = get_walk_speed(user_type)
    walk_time = d / speed * crowd
    return float(d), float(walk_time), float(crowd)

def get_real_wait(project_cn, query_time):
    if project_cn not in open_projects_today:
        return np.nan, None, "当天未开放"
    ht = floor_to_hour(query_time)
    sub = real_hourly[real_hourly["项目中文名"] == project_cn].copy()
    if sub.empty: return np.nan, None, "当天未开放"
    exact = sub[sub[real_time_col] == ht]
    if not exact.empty:
        return float(exact[real_wait_col].iloc[0]), ht, "向下取整整点匹配"
    sub["time_diff"] = (sub[real_time_col] - ht).abs()
    nearest = sub.sort_values("time_diff").iloc[0]
    return float(nearest[real_wait_col]), nearest[real_time_col], "向下取整后最近整点补齐"

def is_core_project(project_name):
    if info_type_col is None or info_type_col not in project_info_lookup.columns:
        return True
    val = project_info_lookup.loc[project_name, info_type_col]
    if pd.isna(val): return True
    s = str(val)
    if ("餐" in s) or ("美食" in s) or ("购物" in s) or ("商店" in s):
        return False
    return True

def has_height_limit(project_name):
    if info_height_col is None or info_height_col not in project_info_lookup.columns:
        return False
    h = project_info_lookup.loc[project_name, info_height_col]
    if pd.isna(h): return False
    hs = str(h).strip()
    if hs == "" or hs.lower() == "nan" or hs in ["无", "不限", "0", "0.0"]: return False
    return True

def get_attr_name(project_name):
    if project_name not in project_info_lookup.index: return None
    attr_name = project_info_lookup.loc[project_name, "八维表对应项目名"]
    if pd.isna(attr_name): return None
    attr_name = str(attr_name).strip()
    if attr_name not in attr_lookup_df.index: return None
    return attr_name

def calc_first_question_x(candidate_project, done_projects, nf, ns):
    selected_projects = list(done_projects) + [candidate_project]
    attr_names = []
    for p in selected_projects:
        a = get_attr_name(p)
        if a is not None: attr_names.append(a)
    attr_names = list(dict.fromkeys(attr_names))
    if len(attr_names) == 0: return None
    sub = attr_lookup_df.loc[attr_names].copy()
    x = {}
    for dim in CORE_DIM_COLS:
        col = dim_col_map[dim]
        vals = pd.to_numeric(sub[col], errors="coerce")
        if isinstance(vals, pd.Series):
            x[dim] = float(vals.mean())
        else:
            x[dim] = float(vals)
    x["餐饮"] = float(z_food(nf))
    x["购物"] = float(z_shop(ns))
    return x

def calc_candidate_base(project_name, user_type, current_location, done_projects, current_dt, queue_deadline_dt):
    if project_name not in open_projects_today:
        return None, "当天未开放"
    if project_name not in project_info_lookup.index:
        return None, "不在项目主表"
    if not is_core_project(project_name):
        return None, "非核心项目（餐饮/购物）"
    if project_name in done_projects:
        return None, "已完成"
    if user_type == "亲子" and has_height_limit(project_name):
        return None, "亲子剔除身高限制项目"

    attr_name = get_attr_name(project_name)
    if attr_name is None:
        return None, "缺少八维向量"

    q_hist = hist_wait_dict.get(project_name, np.nan)
    if pd.isna(q_hist):
        return None, "缺少历史平均等待时间"
    q_hist_safe = max(float(q_hist), 1.0)

    distance_m, walk_time_min, crowd_factor = calc_walk_time_min(
        from_loc=current_location, to_project=project_name, user_type=user_type
    )
    if pd.isna(walk_time_min):
        return None, "步行时间无法计算"

    arrive_dt = current_dt + pd.Timedelta(minutes=float(walk_time_min))
    if arrive_dt >= queue_deadline_dt:
        return None, "到达时已超过21:30截止排队"

    q_real, used_time, real_source = get_real_wait(project_name, arrive_dt)
    if pd.isna(q_real):
        return None, "缺少实时等待时间"

    real_total_time = float(q_real) + float(walk_time_min)
    hist_total_time = q_hist_safe + float(walk_time_min)
    total_time_ratio = real_total_time / max(hist_total_time, 1.0)
    time_comfort_x9 = math.exp(-max(0.0, total_time_ratio - 1.0))

    x8 = calc_first_question_x(
        candidate_project=project_name, done_projects=done_projects,
        nf=NF_FOR_SATISFACTION, ns=NS_FOR_SATISFACTION
    )
    if x8 is None:
        return None, "八维计算失败"

    row = {
        "项目中文名": project_name,
        "八维表对应项目名": attr_name,
        "园区": str(project_info_lookup.loc[project_name, info_area_col]) if info_area_col in project_info_lookup.columns else "",
        "大类": str(project_info_lookup.loc[project_name, info_type_col]) if info_type_col in project_info_lookup.columns else "",
        "距离_m": round(distance_m, 1),
        "步行时间_min": round(walk_time_min, 2),
        "预计到达时间": arrive_dt.strftime("%H:%M:%S"),
        "实时等待_min": round(float(q_real), 1),
        "历史平均等待_min": round(float(q_hist), 1),
        "实时总时间_min": round(real_total_time, 1),
        "历史基准总时间_min": round(hist_total_time, 1),
        "总时间倍率_r": round(total_time_ratio, 4),
        "第九维时间舒适度_x9": round(time_comfort_x9, 6),
        "排队取整时间": used_time.strftime("%H:%M:%S") if used_time is not None else "",
        "实时数据来源": real_source,
        "拥挤系数": round(crowd_factor, 3),
    }
    if info_duration_col and info_duration_col in project_info_lookup.columns:
        val = project_info_lookup.loc[project_name, info_duration_col]
        row["时长_min"] = float(val) if pd.notna(val) else 0
    if info_height_col and info_height_col in project_info_lookup.columns:
        val = project_info_lookup.loc[project_name, info_height_col]
        row["身高限制_cm"] = str(val) if pd.notna(val) else ""
    for dim in DIM_COLS:
        row[f"x_{dim}"] = round(x8[dim], 6)
    return row, None

def finalize_dynamic_scores(rows, user_type):
    if not rows: return [], []
    mu = mu_dict[user_type]
    kk = K_dict[user_type]
    results = []
    for row in rows:
        exponent_8 = 0.0
        for dim in DIM_COLS:
            exponent_8 += kk[dim] * (float(row[f"x_{dim}"]) - mu[dim]) ** 2
        exponent_9 = K9 * (float(row["第九维时间舒适度_x9"]) - 1.0) ** 2
        u = math.exp(-(exponent_8 + exponent_9))
        row["八维满意度指数项"] = round(exponent_8, 6)
        row["第九维指数项"] = round(exponent_9, 6)
        row["动态满意度_Udyn"] = round(u, 6)
        results.append(row)
    results.sort(key=lambda x: x["动态满意度_Udyn"], reverse=True)
    for i, r in enumerate(results):
        r["动态满意度排名"] = i + 1
    top5 = results[:5]
    return results, top5

# --- 路线模拟 ---

def get_project_type(project_name):
    if project_name not in project_info_lookup.index: return ""
    if info_type_col is None or info_type_col not in project_info_lookup.columns: return ""
    val = project_info_lookup.loc[project_name, info_type_col]
    if pd.isna(val): return ""
    return str(val)

def is_show_project(project_name):
    s = get_project_type(project_name)
    return ("演出" in s) or ("表演" in s) or ("娱乐演出" in s)

def is_meet_project(project_name):
    s = get_project_type(project_name)
    return ("见面" in s) or ("迪士尼朋友" in s)

def get_duration_min(project_name):
    if is_meet_project(project_name): return 0.0
    if info_duration_col and info_duration_col in project_info_lookup.columns:
        val = project_info_lookup.loc[project_name, info_duration_col]
        if pd.notna(val):
            try: return max(float(val), 0.0)
            except: return 0.0
    return 0.0

def simulate_route_sequence(sequence, start_time, start_location, user_type,
                            lunch_done_init=False, dinner_done_init=False):
    t = start_time
    loc = start_location
    route_events = []
    total_walk = 0.0; total_wait = 0.0; total_play = 0.0; total_meal = 0.0
    lunch_done = lunch_done_init; dinner_done = dinner_done_init
    completed_projects = []; stop_reason = ""

    meal_duration = 60 if user_type == "亲子" else 45
    park_close_dt = time_str_to_datetime(TARGET_DATE, PARK_CLOSE_TIME)
    queue_deadline_dt = time_str_to_datetime(TARGET_DATE, QUEUE_DEADLINE)
    lunch_start = time_str_to_datetime(TARGET_DATE, LUNCH_WINDOW[0])
    lunch_end = time_str_to_datetime(TARGET_DATE, LUNCH_WINDOW[1])
    dinner_start = time_str_to_datetime(TARGET_DATE, DINNER_WINDOW[0])
    dinner_end = time_str_to_datetime(TARGET_DATE, DINNER_WINDOW[1])

    for step_idx, project_name in enumerate(sequence, start=1):
        ok, reason = True, ""
        if project_name not in open_projects_today:
            ok, reason = False, "当天未开放"
        elif project_name not in project_info_lookup.index:
            ok, reason = False, "不在项目主表"
        elif user_type == "亲子" and has_height_limit(project_name):
            ok, reason = False, "亲子游客不能参加有身高限制项目"
        elif project_name in completed_projects:
            ok, reason = False, "项目已完成"
        if not ok:
            stop_reason = f"第{step_idx}个项目【{project_name}】不可行：{reason}"
            break

        # 午餐插入
        if not lunch_done and lunch_start <= t <= lunch_end:
            end_t = t + pd.Timedelta(minutes=meal_duration)
            if end_t <= lunch_end:
                route_events.append({"步骤": "", "节点类型": "午餐", "项目名称": "就近餐厅午餐",
                    "开始时间": t.strftime("%H:%M:%S"), "结束时间": end_t.strftime("%H:%M:%S"),
                    "耗时_min": meal_duration, "说明": "午餐时间窗"})
                t = end_t; lunch_done = True; total_meal += meal_duration

        # 晚餐插入
        if not dinner_done and dinner_start <= t <= dinner_end:
            end_t = t + pd.Timedelta(minutes=meal_duration)
            if end_t <= dinner_end:
                route_events.append({"步骤": "", "节点类型": "晚餐", "项目名称": "就近餐厅晚餐",
                    "开始时间": t.strftime("%H:%M:%S"), "结束时间": end_t.strftime("%H:%M:%S"),
                    "耗时_min": meal_duration, "说明": "晚餐时间窗"})
                t = end_t; dinner_done = True; total_meal += meal_duration

        distance_m, walk_time_min, crowd_factor = calc_walk_time_min(
            from_loc=loc, to_project=project_name, user_type=user_type
        )
        if pd.isna(walk_time_min):
            stop_reason = f"第{step_idx}个项目步行时间无法计算"
            break

        walk_end = t + pd.Timedelta(minutes=float(walk_time_min))
        if walk_end >= queue_deadline_dt:
            stop_reason = f"第{step_idx}个项目到达时间不早于21:30，保留前{len(completed_projects)}个项目"
            break

        route_events.append({"步骤": step_idx, "节点类型": "步行",
            "项目名称": f"{loc} → {project_name}",
            "开始时间": t.strftime("%H:%M:%S"), "结束时间": walk_end.strftime("%H:%M:%S"),
            "耗时_min": round(float(walk_time_min), 2), "距离_m": round(float(distance_m), 1),
            "拥挤系数": round(float(crowd_factor), 3), "说明": "步行至目标项目"})
        t = walk_end; total_walk += float(walk_time_min)

        q_real, used_time, real_source = get_real_wait(project_name, t)
        if pd.isna(q_real):
            stop_reason = "缺少实时排队时间"
            break

        queue_end = t + pd.Timedelta(minutes=float(q_real))
        route_events.append({"步骤": step_idx, "节点类型": "排队", "项目名称": project_name,
            "开始时间": t.strftime("%H:%M:%S"), "结束时间": queue_end.strftime("%H:%M:%S"),
            "耗时_min": round(float(q_real), 1),
            "排队取整时间": used_time.strftime("%H:%M:%S") if used_time else "",
            "说明": "排队等待"})
        total_wait += float(q_real)

        duration_min = get_duration_min(project_name)
        play_end = queue_end + pd.Timedelta(minutes=float(duration_min))
        if play_end > park_close_dt:
            stop_reason = f"第{step_idx}个项目完成时间超闭园，保留前{len(completed_projects)}个项目"
            break

        if is_show_project(project_name): node_type = "观看演出"
        elif is_meet_project(project_name): node_type = "见面互动"
        else: node_type = "游玩"

        route_events.append({"步骤": step_idx, "节点类型": node_type, "项目名称": project_name,
            "开始时间": queue_end.strftime("%H:%M:%S"), "结束时间": play_end.strftime("%H:%M:%S"),
            "耗时_min": round(float(duration_min), 1), "说明": "核心项目体验"})
        total_play += float(duration_min)
        completed_projects.append(project_name)
        t = play_end; loc = project_name

    if len(completed_projects) == 0:
        return False, None, route_events, stop_reason if stop_reason else "未能完成任何项目"

    # 补晚餐
    if not dinner_done and dinner_start <= t <= dinner_end:
        end_t = t + pd.Timedelta(minutes=meal_duration)
        if end_t <= dinner_end:
            route_events.append({"步骤": "", "节点类型": "晚餐", "项目名称": "就近餐厅晚餐",
                "开始时间": t.strftime("%H:%M:%S"), "结束时间": end_t.strftime("%H:%M:%S"),
                "耗时_min": meal_duration, "说明": "补晚餐"})
            t = end_t; dinner_done = True; total_meal += meal_duration

    core_end_time = t
    shop_count = 0; shop_t = t
    while shop_count < MAX_SHOP_TIMES:
        if shop_t >= queue_deadline_dt: break
        shop_end = shop_t + pd.Timedelta(minutes=SHOP_TIME_MIN)
        if shop_end > park_close_dt: break
        shop_count += 1
        route_events.append({"步骤": "", "节点类型": "购物", "项目名称": f"就近商店购物{shop_count}",
            "开始时间": shop_t.strftime("%H:%M:%S"), "结束时间": shop_end.strftime("%H:%M:%S"),
            "耗时_min": SHOP_TIME_MIN, "说明": "空闲购物"})
        shop_t = shop_end

    t_after_shop = shop_t
    free_until_close = max((park_close_dt - t_after_shop).total_seconds() / 60, 0)
    total_route_time = (t_after_shop - start_time).total_seconds() / 60
    full_completed = len(completed_projects) == len(sequence)

    summary = {
        "原计划项目数": len(sequence),
        "实际完成项目数": len(completed_projects),
        "是否完成原路线": "是" if full_completed else "否",
        "原计划路线": " → ".join(sequence),
        "实际可执行路线": " → ".join(completed_projects),
        "未执行项目": "；".join([p for p in sequence if p not in completed_projects]) if not full_completed else "",
        "总耗时_min": round(total_route_time, 1),
        "总步行_min": round(total_walk, 1),
        "总排队_min": round(total_wait, 1),
        "总游玩_min": round(total_play, 1),
        "就餐_min": round(total_meal, 1),
        "购物_min": round(shop_count * SHOP_TIME_MIN, 1),
        "午餐是否已安排": "是" if lunch_done else "否",
        "晚餐是否已安排": "是" if dinner_done else "否",
        "购物次数": shop_count,
        "结束后至闭园空闲_min": round(free_until_close, 1),
        "降级说明": "" if full_completed else stop_reason,
        "detail_events": route_events,
    }
    return True, summary, route_events, stop_reason

# ============================================================
# 5. HTTP 请求处理（纯标准库，无需 Flask）
# ============================================================

# 读取 HTML 模板
with open(os.path.join(TEMPLATE_DIR, "index.html"), "r", encoding="utf-8") as f:
    HTML_TEMPLATE = f.read()

class DisneyRequestHandler(BaseHTTPRequestHandler):

    def log_message(self, format, *args):
        # 简洁日志
        print(f"  [{self.command}] {args[0]}")

    def _send_json(self, data, status=200):
        body = json.dumps(data, ensure_ascii=False, default=str).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def _send_html(self, html, status=200):
        body = html.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_error_json(self, message, status=400):
        self._send_json({"error": message}, status)

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path
        qs = parse_qs(parsed.query)

        if path == "/" or path == "/index.html":
            self._send_html(HTML_TEMPLATE)

        elif path == "/api/data-info":
            self._send_json({
                "target_date": TARGET_DATE,
                "date_type": DATE_TYPE,
                "open_projects_count": len(open_projects_today),
                "open_projects": sorted(list(open_projects_today)),
                "hist_wait_count": len(hist_wait_dict),
                "attr_count": len(attr_lookup_df),
                "crowd_col": str(crowd_col),
                "eta": ETA_TIME,
                "k9": round(K9, 6),
                "queue_deadline": QUEUE_DEADLINE,
                "park_close": PARK_CLOSE_TIME,
            })

        elif path == "/api/projects":
            all_names = sorted(project_info_full["项目中文名"].dropna().astype(str).unique().tolist())
            self._send_json({
                "all_projects": all_names,
                "open_projects": sorted(list(open_projects_today)),
                "location_options": [ENTRANCE_NAME, "米奇大街"] + all_names,
            })

        else:
            self._send_error_json("Not Found", 404)

    def do_POST(self):
        parsed = urlparse(self.path)
        path = parsed.path

        if path == "/api/calculate":
            # 读取请求体
            content_length = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(content_length) if content_length > 0 else b"{}"
            try:
                data = json.loads(body.decode("utf-8"))
            except json.JSONDecodeError:
                self._send_error_json("无效的 JSON 请求体", 400)
                return

            current_time = data.get("current_time", "09:35")
            user_type = data.get("user_type", "普通")
            current_location = data.get("current_location", ENTRANCE_NAME)
            done_projects = data.get("done_projects", [])

            if user_type not in mu_dict:
                self._send_error_json(f"无效的游客类型：{user_type}", 400)
                return

            current_dt = time_str_to_datetime(TARGET_DATE, current_time)
            queue_deadline_dt = time_str_to_datetime(TARGET_DATE, QUEUE_DEADLINE)
            if current_dt >= queue_deadline_dt:
                self._send_error_json("当前时间已不早于21:30，所有项目已截止排队", 400)
                return

            # ===== Phase 1: Dynamic Satisfaction =====
            candidate_projects = sorted(open_projects_today)
            rows = []
            exclude_records = []
            for project_name in candidate_projects:
                row, reason = calc_candidate_base(
                    project_name, user_type, current_location, done_projects,
                    current_dt, queue_deadline_dt
                )
                if row is not None:
                    rows.append(row)
                else:
                    exclude_records.append({"项目中文名": project_name, "剔除原因": reason})

            all_results, top5 = finalize_dynamic_scores(rows, user_type)

            # ===== Phase 2: Route Planning =====
            top5_projects = [r["项目中文名"] for r in top5]
            route_summaries = []
            route_details = []
            infeasible_count = 0

            if len(top5_projects) >= 3:
                combos = list(itertools.combinations(top5_projects, 3))
            elif len(top5_projects) >= 1:
                combos = [tuple(top5_projects)]
            else:
                combos = []

            for combo_idx, combo in enumerate(combos, start=1):
                best_summary = None; best_detail = None; best_order = None
                best_count = -1; best_time = float("inf")
                for seq in itertools.permutations(combo):
                    feasible, summary, detail_events, reason = simulate_route_sequence(
                        sequence=seq, start_time=current_dt,
                        start_location=current_location, user_type=user_type
                    )
                    if feasible:
                        cnt = int(summary["实际完成项目数"])
                        ttime = float(summary["总耗时_min"])
                        if cnt > best_count or (cnt == best_count and ttime < best_time):
                            best_count = cnt; best_time = ttime
                            best_summary = summary; best_detail = detail_events; best_order = seq
                    else:
                        infeasible_count += 1
                if best_summary is not None:
                    best_summary["组合编号"] = combo_idx
                    best_summary["项目组合"] = " + ".join(combo)
                    best_summary["最优访问顺序"] = " → ".join(best_order)
                    route_summaries.append(best_summary)
                    for ev in best_detail:
                        ev["组合编号"] = combo_idx
                        ev["项目组合"] = " + ".join(combo)
                        ev["最优访问顺序"] = " → ".join(best_order)
                    route_details.extend(best_detail)

            route_summaries.sort(key=lambda x: (x["实际完成项目数"], -x["总耗时_min"]), reverse=True)
            for i, s in enumerate(route_summaries):
                s["路线推荐排名"] = i + 1

            self._send_json({
                "params": {
                    "current_time": current_time,
                    "user_type": user_type,
                    "current_location": current_location,
                    "done_projects": done_projects,
                    "date_type": DATE_TYPE,
                    "target_date": TARGET_DATE,
                    "queue_deadline": QUEUE_DEADLINE,
                },
                "phase1": {
                    "all_results": all_results,
                    "top5": top5,
                    "exclude_records": exclude_records,
                    "total_candidates": len(candidate_projects),
                    "total_valid": len(all_results),
                },
                "phase2": {
                    "route_summaries": route_summaries,
                    "route_details": route_details,
                    "infeasible_count": infeasible_count,
                    "total_combos": len(combos),
                }
            })

        else:
            self._send_error_json("Not Found", 404)

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

# ============================================================
# 6. 启动服务（适配 Render 云端环境）
# ============================================================

if __name__ == "__main__":
    # Render 会通过环境变量 PORT 指定监听端口
    port = int(os.environ.get("PORT", 5000))
    server = HTTPServer(("0.0.0.0", port), DisneyRequestHandler)
    print("\n" + "=" * 60)
    print("  🏰 迪士尼乐园路径规划可视化系统")
    print(f"  🌐 服务已启动，监听端口 {port}")
    print("  ⚠️  关闭此窗口或按 Ctrl+C 停止服务")
    print("=" * 60 + "\n")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n服务已停止。")
        server.shutdown()
