# geo_utils.py - 地理计算工具函数

import uuid
import numpy as np
import openpyxl
import pandas as pd
import geopandas as gpd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
from matplotlib import font_manager
from pathlib import Path
from shapely.geometry import box
import contextily as ctx
from pyproj import Transformer

from city_data import CITIES
from check_land_sea import load_land_polygons, classify_coordinates, detect_col, LON_ALIASES, LAT_ALIASES

UPLOADS_DIR = Path(__file__).parent / "static" / "uploads"
UPLOADS_DIR.mkdir(parents=True, exist_ok=True)

# 注册中文字体
_CN_FONT_PATHS = [
    "/System/Library/Fonts/STHeiti Medium.ttc",
    "/System/Library/Fonts/STHeiti Light.ttc",
    "/System/Library/Fonts/PingFang.ttc",
    "C:/Windows/Fonts/msyh.ttc",
    "C:/Windows/Fonts/simhei.ttf",
]
for _p in _CN_FONT_PATHS:
    if Path(_p).exists():
        font_manager.fontManager.addfont(_p)
        _cn_font = font_manager.FontProperties(fname=_p).get_name()
        plt.rcParams["font.family"] = _cn_font
        break
plt.rcParams["axes.unicode_minus"] = False

SHEET_COLORS = [
    "#1a3a6e", "#4fc3f7", "#00bcd4", "#ffeb3b",
    "#ff9800", "#e53935", "#e91e8c", "#9c27b0",
    "#4caf50", "#ff5722", "#607d8b", "#795548",
    "#f06292", "#aed581", "#4dd0e1", "#ffd54f",
]


def generate_land_points(city_name: str, count: int) -> tuple:
    """
    生成指定城市 count 个陆地随机点，按区划最近中心分配，写入 Excel。
    返回 (excel_path, summary)
    summary: [{"sheet": name, "count": n}, ...]
    """
    if city_name not in CITIES:
        raise ValueError(f"不支持的城市：{city_name}")

    city = CITIES[city_name]
    lon_min, lon_max, lat_min, lat_max = city["bbox"]
    districts = city["districts"]

    # 加载陆地多边形（使用缓存）
    land = load_land_polygons()
    bbox_geom = box(lon_min, lat_min, lon_max, lat_max)
    land_clip = land[land.intersects(bbox_geom)].copy()

    rng = np.random.default_rng()

    # 拒绝采样：随机生成 → 过滤陆地
    collected_lon, collected_lat = [], []
    while len(collected_lon) < count:
        batch = max(count * 6, 5000)
        # 均匀随机 + 高斯混合制造不规则感
        lons = rng.uniform(lon_min, lon_max, batch)
        lats = rng.uniform(lat_min, lat_max, batch)
        # 高斯混合叠加
        for d in districts:
            n = batch // len(districts)
            sx = (lon_max - lon_min) * 0.08
            sy = (lat_max - lat_min) * 0.08
            lons = np.append(lons, rng.normal(d["lon"], sx, n))
            lats = np.append(lats, rng.normal(d["lat"], sy, n))

        in_range = (
            (lons >= lon_min) & (lons <= lon_max) &
            (lats >= lat_min) & (lats <= lat_max)
        )
        lons, lats = lons[in_range], lats[in_range]

        pts = gpd.GeoDataFrame(
            geometry=gpd.points_from_xy(lons, lats), crs="EPSG:4326"
        )
        joined = gpd.sjoin(pts, land_clip[["geometry"]], how="left", predicate="within")
        joined = joined[~joined.index.duplicated(keep="first")]
        mask = joined["index_right"].notna().values
        take = count - len(collected_lon)
        collected_lon.extend(lons[mask][:take].tolist())
        collected_lat.extend(lats[mask][:take].tolist())

    collected_lon = np.array(collected_lon[:count])
    collected_lat = np.array(collected_lat[:count])

    # 按最近区划中心分配（经度×cos(lat) 距离修正）
    centers = np.array([[d["lon"], d["lat"]] for d in districts])
    lat_rad = np.deg2rad(collected_lat)
    dx = (collected_lon[:, None] - centers[:, 0]) * np.cos(lat_rad)[:, None]
    dy = collected_lat[:, None] - centers[:, 1]
    nearest = np.argmin(np.sqrt(dx**2 + dy**2), axis=1)

    # 写入 Excel
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    summary = []
    for i, d in enumerate(districts):
        mask = nearest == i
        ws = wb.create_sheet(d["name"])
        ws.append(["number", "lon", "lat"])
        lons_d = collected_lon[mask]
        lats_d = collected_lat[mask]
        for j, (lo, la) in enumerate(zip(lons_d, lats_d), 1):
            ws.append([j, round(float(lo), 8), round(float(la), 8)])
        summary.append({"sheet": d["name"], "count": int(mask.sum())})

    fname = f"{city_name}_{count}_{uuid.uuid4().hex[:8]}.xlsx"
    out_path = UPLOADS_DIR / fname
    wb.save(out_path)
    return str(out_path), summary


def load_all_sheets(path: str, lon_col_arg=None, lat_col_arg=None) -> dict:
    """读取所有 Sheet，返回 {sheet_name: {df, lon, lat, xs, ys}}"""
    xl = pd.ExcelFile(path, engine="openpyxl")
    sheets = {}
    for name in xl.sheet_names:
        df = xl.parse(name)
        cols = list(df.columns)
        lon_col = lon_col_arg or detect_col(cols, LON_ALIASES, "lon")
        lat_col = lat_col_arg or detect_col(cols, LAT_ALIASES, "lat")
        if not lon_col or not lat_col:
            continue
        df = df.dropna(subset=[lon_col, lat_col])
        df = df[df[lon_col].between(-180, 180) & df[lat_col].between(-90, 90)]
        if df.empty:
            continue
        sheets[name] = {"df": df, "lon": lon_col, "lat": lat_col}
    return sheets


def wgs84_to_webmercator(lons, lats):
    transformer = Transformer.from_crs("EPSG:4326", "EPSG:3857", always_xy=True)
    return transformer.transform(lons, lats)


def generate_map(excel_path: str) -> str:
    """读取 Excel，生成卫星底图 PNG，返回相对于 static/ 的路径"""
    sheets = load_all_sheets(excel_path)
    if not sheets:
        raise ValueError("Excel 中没有可识别的经纬度数据")

    all_xs, all_ys = [], []
    for name, info in sheets.items():
        df = info["df"]
        xs, ys = wgs84_to_webmercator(df[info["lon"]].values, df[info["lat"]].values)
        info["xs"] = xs
        info["ys"] = ys
        all_xs.extend(xs)
        all_ys.extend(ys)

    x_min, x_max = min(all_xs), max(all_xs)
    y_min, y_max = min(all_ys), max(all_ys)
    x_pad = (x_max - x_min) * 0.08
    y_pad = (y_max - y_min) * 0.08

    fig, ax = plt.subplots(figsize=(20, 11), dpi=120)
    ax.set_xlim(x_min - x_pad, x_max + x_pad)
    ax.set_ylim(y_min - y_pad, y_max + y_pad)

    try:
        ctx.add_basemap(ax, crs="EPSG:3857",
                        source=ctx.providers.Esri.WorldImagery, zoom="auto")
    except Exception:
        try:
            ctx.add_basemap(ax, crs="EPSG:3857",
                            source=ctx.providers.OpenStreetMap.Mapnik, zoom="auto")
        except Exception:
            ax.set_facecolor("#1a1a2e")

    legend_handles = []
    for i, (name, info) in enumerate(sheets.items()):
        df = info["df"]
        xs, ys = info["xs"], info["ys"]
        color = SHEET_COLORS[i % len(SHEET_COLORS)]
        if "陆地或海洋" in df.columns:
            mask_land = df["陆地或海洋"] == "陆地"
            mask_ocean = df["陆地或海洋"] == "海洋"
            if mask_land.any():
                ax.scatter(xs[mask_land.values], ys[mask_land.values],
                           c=color, s=6, alpha=0.75, linewidths=0, zorder=3)
            if mask_ocean.any():
                ax.scatter(xs[mask_ocean.values], ys[mask_ocean.values],
                           c=color, s=16, alpha=0.9, marker="x", zorder=4)
        else:
            ax.scatter(xs, ys, c=color, s=6, alpha=0.75, linewidths=0, zorder=3)
        handle = Line2D([0], [0], marker="o", color="w",
                        markerfacecolor=color, markersize=8,
                        label=f"{name}  ({len(df):,})")
        legend_handles.append(handle)

    total = sum(len(v["df"]) for v in sheets.values())
    ax.set_title(f"经纬度分布图  |  共 {total:,} 个点",
                 fontsize=15, fontweight="bold", color="white",
                 pad=12, backgroundcolor="#1a1a2e")
    legend = ax.legend(handles=legend_handles, loc="lower left",
                       fontsize=8, framealpha=0.85,
                       facecolor="#1a1a2e", edgecolor="#444",
                       labelcolor="white", title="区划",
                       title_fontsize=9)
    legend.get_title().set_color("white")
    ax.set_axis_off()
    plt.tight_layout(pad=0)

    stem = Path(excel_path).stem
    out_name = f"{stem}_map.png"
    out_path = UPLOADS_DIR / out_name
    fig.savefig(out_path, dpi=120, bbox_inches="tight",
                facecolor="#1a1a2e", edgecolor="none")
    plt.close(fig)

    # 返回相对 static/ 的路径供 url_for 使用
    return f"uploads/{out_name}"
