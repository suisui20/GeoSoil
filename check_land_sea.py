#!/usr/bin/env python3
"""
check_land_sea.py - 核实经纬度是否在陆地或海洋上

数据源：Natural Earth ne_10m_land（1:10m 比例尺矢量多边形，精度接近 ArcGIS）
首次运行自动下载并缓存到 ~/.cache/land_polygons/，后续离线使用。

用法:
    python check_land_sea.py input.xlsx
    python check_land_sea.py input.xlsx --output result.xlsx
    python check_land_sea.py input.xlsx --lon-col 经度 --lat-col 纬度
"""

import argparse
import sys
import urllib.request
import zipfile
from pathlib import Path

import geopandas as gpd
import pandas as pd

LAND_URL = "https://naturalearth.s3.amazonaws.com/10m_physical/ne_10m_land.zip"
CACHE_DIR = Path.home() / ".cache" / "land_polygons" / "ne_10m_land"


def load_land_polygons() -> gpd.GeoDataFrame:
    shp_files = list(CACHE_DIR.glob("*.shp"))
    if shp_files:
        return gpd.read_file(shp_files[0])

    print("正在下载 Natural Earth ne_10m_land 数据（仅首次，约 3MB）...")
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    zip_path = CACHE_DIR / "ne_10m_land.zip"

    try:
        urllib.request.urlretrieve(LAND_URL, zip_path)
    except Exception as e:
        print(f"错误：下载失败：{e}", file=sys.stderr)
        print(f"请手动下载 {LAND_URL} 并解压到 {CACHE_DIR}", file=sys.stderr)
        sys.exit(1)

    with zipfile.ZipFile(zip_path, "r") as zf:
        zf.extractall(CACHE_DIR)
    zip_path.unlink()

    shp_files = list(CACHE_DIR.glob("*.shp"))
    if not shp_files:
        print("错误：解压后未找到 .shp 文件", file=sys.stderr)
        sys.exit(1)

    print("下载完成，已缓存到本地。")
    return gpd.read_file(shp_files[0])


def parse_args():
    parser = argparse.ArgumentParser(
        description="读取 Excel 文件中的经纬度，判断每个坐标点是否在陆地或海洋上，结果写回新列。"
    )
    parser.add_argument("input_file", help="输入 Excel 文件路径（.xlsx）")
    parser.add_argument("--output", help="输出文件路径，默认为 {原文件名}_result.xlsx")
    parser.add_argument("--lon-col", default="lon", help="经度列名（默认：lon）")
    parser.add_argument("--lat-col", default="lat", help="纬度列名（默认：lat）")
    return parser.parse_args()


# 自动从列名中识别经纬度列，支持多种命名约定
LON_ALIASES = ["lon", "longitude", "经度", "lng", "x"]
LAT_ALIASES = ["lat", "latitude", "纬度", "y"]

def detect_col(columns, aliases, prefer):
    if prefer in columns:
        return prefer
    for alias in aliases:
        if alias in columns:
            return alias
    return None


def resolve_output(input_path: Path, output_arg) -> Path:
    if output_arg:
        return Path(output_arg)
    return input_path.parent / (input_path.stem + "_result.xlsx")


def classify_coordinates(df: pd.DataFrame, lat_col: str, lon_col: str, land: gpd.GeoDataFrame) -> pd.Series:
    results = pd.Series("未知", index=df.index, dtype=str)

    lat = df[lat_col]
    lon = df[lon_col]

    mask_valid = (
        lat.notna()
        & lon.notna()
        & lat.between(-90, 90)
        & lon.between(-180, 180)
    )

    mask_has_value = lat.notna() & lon.notna()
    results[mask_has_value & ~mask_valid] = "无效坐标"

    if not mask_valid.any():
        return results

    valid_df = df[mask_valid][[lon_col, lat_col]].copy()
    points_gdf = gpd.GeoDataFrame(
        valid_df,
        geometry=gpd.points_from_xy(valid_df[lon_col], valid_df[lat_col]),
        crs="EPSG:4326",
    )

    joined = gpd.sjoin(points_gdf, land[["geometry"]], how="left", predicate="within")
    # 同一点可能匹配多个多边形，去重取第一条
    joined = joined[~joined.index.duplicated(keep="first")]

    is_land = joined["index_right"].notna()
    results[mask_valid] = is_land.map({True: "陆地", False: "海洋"})

    return results


def print_summary(results: pd.Series, output_path: Path):
    total = len(results)
    counts = results.value_counts()

    print(f"\n处理完成：共 {total} 行")
    for label in ["陆地", "海洋", "未知", "无效坐标"]:
        if label in counts:
            n = counts[label]
            print(f"  {label}: {n:>6}  ({n / total * 100:.1f}%)")
    print(f"\n结果已写入: {output_path}")


def main():
    args = parse_args()
    input_path = Path(args.input_file)

    if not input_path.exists():
        print(f"错误：文件不存在：{input_path}", file=sys.stderr)
        sys.exit(1)

    if input_path.suffix.lower() != ".xlsx":
        print(f"错误：仅支持 .xlsx 格式，当前文件：{input_path.suffix}", file=sys.stderr)
        sys.exit(1)

    print(f"读取文件：{input_path}")
    xl = pd.ExcelFile(input_path, engine="openpyxl")
    sheet_names = xl.sheet_names
    print(f"发现 {len(sheet_names)} 个 Sheet：{sheet_names}")

    land = load_land_polygons()
    print("正在分类坐标（使用 Natural Earth 1:10m 矢量数据）...\n")

    output_path = resolve_output(input_path, args.output)
    all_results = []
    total_rows = 0

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet in sheet_names:
            df = xl.parse(sheet)
            cols = list(df.columns)

            lon_col = detect_col(cols, LON_ALIASES, args.lon_col)
            lat_col = detect_col(cols, LAT_ALIASES, args.lat_col)

            if not lon_col or not lat_col:
                print(f"  [{sheet}] 跳过：找不到经纬度列（列名：{cols}）", file=sys.stderr)
                df.to_excel(writer, sheet_name=sheet, index=False)
                continue

            print(f"  [{sheet}] {len(df)} 行，经度={lon_col!r}，纬度={lat_col!r}")
            df["陆地或海洋"] = classify_coordinates(df, lat_col, lon_col, land)
            df.to_excel(writer, sheet_name=sheet, index=False)

            all_results.append(df["陆地或海洋"])
            total_rows += len(df)

    if all_results:
        combined = pd.concat(all_results)
        print_summary(combined, output_path)
        print(f"（共处理 {len(sheet_names)} 个 Sheet，合计 {total_rows} 行）")


if __name__ == "__main__":
    main()
