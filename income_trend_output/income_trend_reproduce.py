# -*- coding: utf-8 -*-
"""
收入趋势分析案例：用 Python 复现完整 Excel 案例中的 5 张图。

依赖：pandas、matplotlib、openpyxl
运行：python income_trend_reproduce.py

默认输入文件：
    E:\案例1：收入趋势分析\案例1：收入趋势分析原始数据.xlsx

输出：
    income_trend_output/
        01_年度收入趋势.png
        02_2020月度收入.png
        03_2020月度同比.png
        04_2020年11月各省同比.png
        05_四川省月度同比.png
        income_trend_dashboard.png
        analysis_tables.xlsx
"""

from pathlib import Path
import platform
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.ticker import FuncFormatter, PercentFormatter


# =========================
# 1. 文件与全局格式
# =========================
INPUT_FILE = Path(r"E:\案例1：收入趋势分析\案例1：收入趋势分析原始数据.xlsx")
OUTPUT_DIR = Path("income_trend_output")

# Excel 默认主题蓝，尽量贴近原案例的 Office 图表样式
BLUE = "#4472C4"
ORANGE = "#ED7D31"
GRAY = "#7F7F7F"
GRID = "#D9E2F3"


def set_chinese_font():
    """按常见系统字体顺序设置中文字体，避免中文标题/坐标轴乱码。"""
    candidates = [
        "Microsoft YaHei", "SimHei", "DengXian", "Noto Sans CJK SC",
        "Source Han Sans CN", "WenQuanYi Zen Hei", "Arial Unicode MS",
    ]
    available = {f.name for f in font_manager.fontManager.ttflist}
    for font in candidates:
        if font in available:
            plt.rcParams["font.sans-serif"] = [font]
            break
    plt.rcParams["axes.unicode_minus"] = False


def money_fmt(x, _):
    return f"{x:,.0f}"


def style_axes(ax, percent=False):
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#BFBFBF")
    ax.spines["bottom"].set_color("#BFBFBF")
    ax.tick_params(colors="#404040")
    ax.grid(axis="y", color=GRID, linewidth=0.8)
    ax.set_axisbelow(True)
    if percent:
        ax.yaxis.set_major_formatter(PercentFormatter(1.0, decimals=0))
    else:
        ax.yaxis.set_major_formatter(FuncFormatter(money_fmt))


def save_fig(fig, filename):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    fig.savefig(OUTPUT_DIR / filename, dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(fig)


# =========================
# 2. 读取原始数据并整理
# =========================
def load_data():
    if not INPUT_FILE.exists():
        raise FileNotFoundError(
            f"找不到输入文件：{INPUT_FILE}\n"
            "请把 INPUT_FILE 改成你的原始数据 Excel 路径。"
        )

    raw = pd.read_excel(INPUT_FILE, sheet_name=0)
    if raw.shape[1] < 4:
        raise ValueError("原始数据至少需要 4 列：年份、月份、省份、销售额。")

    # 不依赖 Excel 表头编码，直接按原始表前四列取数，更稳健。
    raw = raw.iloc[:, :4].copy()
    raw.columns = ["年份", "月份", "省份", "销售额"]
    raw["年份"] = pd.to_numeric(raw["年份"], errors="coerce")
    raw["月份"] = pd.to_numeric(raw["月份"], errors="coerce")
    raw["销售额"] = pd.to_numeric(raw["销售额"], errors="coerce").fillna(0)
    raw = raw.dropna(subset=["年份", "月份", "省份"])
    raw["年份"] = raw["年份"].astype(int)
    raw["月份"] = raw["月份"].astype(int)
    raw["省份"] = raw["省份"].astype(str)
    return raw


def yoy_like_excel(current, previous):
    """复现完整案例中的同比公式：(当期-上期)/当期。

    Excel 案例对当期为 0 的情况显示 0；缺少上期而当期有值时显示 100%。
    """
    current = current.fillna(0)
    previous = previous.fillna(0)
    result = pd.Series(0.0, index=current.index)
    mask = current.ne(0)
    result.loc[mask] = (current.loc[mask] - previous.loc[mask]) / current.loc[mask]
    return result


def build_tables(raw):
    yearly = raw.groupby("年份", as_index=False)["销售额"].sum().sort_values("年份")

    monthly_2020 = (
        raw.loc[raw["年份"].eq(2020)]
        .groupby("月份", as_index=False)["销售额"].sum()
        .set_index("月份")
        .reindex(range(1, 13), fill_value=0)
        .rename_axis("月份")
        .reset_index()
    )

    month_compare = (
        raw.loc[raw["年份"].isin([2019, 2020])]
        .pivot_table(index="月份", columns="年份", values="销售额", aggfunc="sum", fill_value=0)
        .reindex(range(1, 13), fill_value=0)
        .rename_axis("月份")
        .reset_index()
    )
    for year in [2019, 2020]:
        if year not in month_compare:
            month_compare[year] = 0
    month_compare["同比增长"] = yoy_like_excel(month_compare[2020], month_compare[2019])

    province_compare = (
        raw.loc[(raw["年份"].isin([2019, 2020])) & (raw["月份"].eq(11))]
        .pivot_table(index="省份", columns="年份", values="销售额", aggfunc="sum", fill_value=0)
    )
    for year in [2019, 2020]:
        if year not in province_compare:
            province_compare[year] = 0
    # 保留原始数据中省份首次出现的顺序，贴近 Excel 案例排序。
    province_order = raw["省份"].drop_duplicates().tolist()
    province_compare = province_compare.reindex(province_order).fillna(0).reset_index()
    province_compare["同比增长"] = yoy_like_excel(
        province_compare[2020], province_compare[2019]
    )

    sichuan = (
        raw.loc[(raw["省份"].eq("四川")) & (raw["年份"].isin([2019, 2020]))]
        .pivot_table(index="月份", columns="年份", values="销售额", aggfunc="sum", fill_value=0)
        .reindex(range(1, 13), fill_value=0)
        .rename_axis("月份")
        .reset_index()
    )
    for year in [2019, 2020]:
        if year not in sichuan:
            sichuan[year] = 0
    sichuan["同比增长"] = yoy_like_excel(sichuan[2020], sichuan[2019])

    return yearly, monthly_2020, month_compare, province_compare, sichuan


# =========================
# 3. 绘制 5 张图
# =========================
def plot_annual(yearly):
    fig, ax = plt.subplots(figsize=(15, 7.5))
    ax.plot(yearly["年份"], yearly["销售额"], color=BLUE, linewidth=2.2)
    ax.set_title("全部地区历年收入趋势", fontsize=15, pad=16)
    ax.set_xlabel("年份")
    ax.set_ylabel("销售额")
    ax.set_xticks(yearly["年份"])
    style_axes(ax)
    save_fig(fig, "01_年度收入趋势.png")


def plot_monthly(monthly_2020):
    fig, ax = plt.subplots(figsize=(15, 7.5))
    ax.plot(monthly_2020["月份"], monthly_2020["销售额"], color=BLUE, linewidth=2.2)
    ax.set_title("2020年各月收入趋势", fontsize=15, pad=16)
    ax.set_xlabel("月份")
    ax.set_ylabel("销售额")
    ax.set_xticks(range(1, 13))
    style_axes(ax)
    save_fig(fig, "02_2020月度收入.png")


def plot_month_yoy(month_compare):
    fig, ax = plt.subplots(figsize=(15, 7.5))
    ax.plot(month_compare["月份"], month_compare["同比增长"], color=BLUE, linewidth=2.2)
    ax.axhline(0, color="#A6A6A6", linewidth=0.8)
    ax.set_title("2020年各月同比增长率", fontsize=15, pad=16)
    ax.set_xlabel("月份")
    ax.set_ylabel("同比增长率")
    ax.set_xticks(range(1, 13))
    style_axes(ax, percent=True)
    save_fig(fig, "03_2020月度同比.png")


def plot_province_yoy(province_compare):
    fig, ax = plt.subplots(figsize=(15, 7.5))
    ax.bar(province_compare["省份"], province_compare["同比增长"], color=BLUE, width=0.72)
    ax.axhline(0, color="#A6A6A6", linewidth=0.8)
    ax.set_title("2020年11月各省同比增长率", fontsize=15, pad=16)
    ax.set_xlabel("省份")
    ax.set_ylabel("同比增长率")
    ax.tick_params(axis="x", labelrotation=45)
    style_axes(ax, percent=True)
    save_fig(fig, "04_2020年11月各省同比.png")


def plot_sichuan_yoy(sichuan):
    fig, ax = plt.subplots(figsize=(15, 7.5))
    ax.plot(sichuan["月份"], sichuan["同比增长"], color=BLUE, linewidth=2.2)
    ax.axhline(0, color="#A6A6A6", linewidth=0.8)
    ax.set_title("四川省2020年各月同比增长率", fontsize=15, pad=16)
    ax.set_xlabel("月份")
    ax.set_ylabel("同比增长率")
    ax.set_xticks(range(1, 13))
    style_axes(ax, percent=True)
    save_fig(fig, "05_四川省月度同比.png")


def plot_dashboard(tables):
    yearly, monthly_2020, month_compare, province_compare, sichuan = tables
    fig, axes = plt.subplots(3, 2, figsize=(18, 24), constrained_layout=True)
    specs = [
        (axes[0, 0], yearly["年份"], yearly["销售额"], "全部地区历年收入趋势", "销售额", False),
        (axes[0, 1], monthly_2020["月份"], monthly_2020["销售额"], "2020年各月收入趋势", "销售额", False),
        (axes[1, 0], month_compare["月份"], month_compare["同比增长"], "2020年各月同比增长率", "同比增长率", True),
        (axes[1, 1], province_compare["省份"], province_compare["同比增长"], "2020年11月各省同比增长率", "同比增长率", True),
        (axes[2, 0], sichuan["月份"], sichuan["同比增长"], "四川省2020年各月同比增长率", "同比增长率", True),
    ]
    for ax, x, y, title, ylabel, percent in specs:
        if title.startswith("2020年11月"):
            ax.bar(x, y, color=BLUE, width=0.72)
            ax.tick_params(axis="x", labelrotation=45)
        else:
            ax.plot(x, y, color=BLUE, linewidth=2.2)
        ax.axhline(0, color="#A6A6A6", linewidth=0.8) if percent else None
        ax.set_title(title, fontsize=14, pad=12)
        ax.set_xlabel("省份" if title.startswith("2020年11月") else ("年份" if "历年" in title else "月份"))
        ax.set_ylabel(ylabel)
        style_axes(ax, percent=percent)
    axes[0, 0].set_xticks(yearly["年份"])
    axes[0, 1].set_xticks(range(1, 13))
    axes[1, 0].set_xticks(range(1, 13))
    axes[2, 0].set_xticks(range(1, 13))
    axes[2, 1].axis("off")
    fig.suptitle("收入趋势分析", fontsize=20, y=1.01)
    save_fig(fig, "income_trend_dashboard.png")


def main():
    set_chinese_font()
    raw = load_data()
    tables = build_tables(raw)
    plot_annual(tables[0])
    plot_monthly(tables[1])
    plot_month_yoy(tables[2])
    plot_province_yoy(tables[3])
    plot_sichuan_yoy(tables[4])
    plot_dashboard(tables)
    with pd.ExcelWriter(OUTPUT_DIR / "analysis_tables.xlsx", engine="openpyxl") as writer:
        raw.to_excel(writer, sheet_name="原始数据", index=False)
        names = ["年度汇总", "2020月度汇总", "月度同比", "省份同比", "四川月度同比"]
        for name, table in zip(names, tables):
            table.to_excel(writer, sheet_name=name, index=False)
    print(f"完成。图表和汇总表已保存到：{OUTPUT_DIR.resolve()}")


if __name__ == "__main__":
    main()
