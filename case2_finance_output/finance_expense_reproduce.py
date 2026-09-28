# -*- coding: utf-8 -*-
"""案例2：财务费用分析——用 Python 复现完整 Excel 案例中的 3 张图。

默认使用“完整案例”中准备区的图表源数据，以保证与 Excel 图表完全一致。
如果希望完全从原始数据重新汇总，把 USE_COMPLETE_CASE_VALUES 改为 False。

依赖：pandas、matplotlib、openpyxl
运行：python finance_expense_reproduce.py
"""

from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.ticker import FuncFormatter, PercentFormatter


COMPLETE_FILE = Path(r"E:\案例2：财务费用分析\案例2：财务费用分析-完整案例.xlsx")
RAW_FILE = Path(r"E:\案例2：财务费用分析\案例2：财务费用分析-原始数据.xlsx")
OUTPUT_DIR = Path("case2_finance_output")

# True：读取完整案例中实际供图表引用的准备区，保证复现 Excel 图表。
# False：读取原始数据并重新汇总。
USE_COMPLETE_CASE_VALUES = True

BLUE = "#4472C4"  # Excel Office 默认蓝色
GRID = "#D9E2F3"


def set_chinese_font():
    candidates = [
        "Microsoft YaHei", "SimHei", "DengXian", "Noto Sans CJK SC",
        "Source Han Sans CN", "WenQuanYi Zen Hei", "Arial Unicode MS",
    ]
    available = {f.name for f in font_manager.fontManager.ttflist}
    for name in candidates:
        if name in available:
            plt.rcParams["font.sans-serif"] = [name]
            break
    plt.rcParams["axes.unicode_minus"] = False


def money_fmt(value, _):
    return f"{value:,.0f}"


def style_axes(ax, percent=False):
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#BFBFBF")
    ax.spines["bottom"].set_color("#BFBFBF")
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


def normalize_expense_name(value):
    return str(value).strip()


def build_from_complete_case():
    """读取完整案例准备区的 3 个图表数据块。"""
    prepared = pd.read_excel(COMPLETE_FILE, sheet_name=4, header=None)

    # Excel 行号：5:6，图表引用 B5:B6 / C5:C6。
    total = pd.DataFrame({
        "月份": prepared.iloc[4:6, 1].astype(str).tolist(),
        "金额": pd.to_numeric(prepared.iloc[4:6, 2]).tolist(),
    })

    # Excel 行号：19:21，图表引用 A19:A21 / E19:E21。
    period = pd.DataFrame({
        "费用类型": prepared.iloc[18:21, 0].map(normalize_expense_name).tolist(),
        "环比差额": pd.to_numeric(prepared.iloc[18:21, 4]).tolist(),
    })

    # Excel 行号：31:38，图表引用 A31:A38 / E31:E38。
    admin = pd.DataFrame({
        "管理费用项": prepared.iloc[30:38, 0].map(normalize_expense_name).tolist(),
        "环比增长": pd.to_numeric(prepared.iloc[30:38, 4]).tolist(),
    })
    return total, period, admin


def build_from_raw_data():
    """从原始数据汇总，保留完整案例中的类别和排序。"""
    total_raw = pd.read_excel(RAW_FILE, sheet_name=0)
    total_raw.columns = ["年份", "月份", "金额"]
    total = total_raw[["月份", "金额"]].copy()
    total["月份"] = total["月份"].map(lambda x: f"{int(x)}月")

    period_raw = pd.read_excel(RAW_FILE, sheet_name=1)
    period_raw = period_raw.iloc[:, :4].copy()
    period_raw.columns = ["年份", "月份", "费用类型", "金额"]
    period_raw["费用类型"] = period_raw["费用类型"].map(normalize_expense_name)
    period_pivot = period_raw.pivot_table(
        index="费用类型", columns="月份", values="金额", aggfunc="sum", fill_value=0
    )
    period = pd.DataFrame({
        "费用类型": period_pivot.index,
        "环比差额": period_pivot.get(7, 0) - period_pivot.get(6, 0),
    }).reset_index(drop=True)

    admin_raw = pd.read_excel(RAW_FILE, sheet_name=2)
    admin_raw = admin_raw.iloc[:, :4].copy()
    admin_raw.columns = ["年份", "月份", "管理费用项", "金额"]
    admin_raw["管理费用项"] = admin_raw["管理费用项"].map(normalize_expense_name)
    admin_pivot = admin_raw.pivot_table(
        index="管理费用项", columns="月份", values="金额", aggfunc="sum", fill_value=0
    )
    # 与完整案例相同的项目顺序。
    order = [
        "办公用品与设备", "办公租金", "差旅费", "工资及福利",
        "培训与发展", "其他费用", "折旧与摊销", "咨询费",
    ]
    admin_pivot = admin_pivot.reindex(order).fillna(0)
    june = admin_pivot.get(6, pd.Series(0, index=admin_pivot.index))
    july = admin_pivot.get(7, pd.Series(0, index=admin_pivot.index))
    admin = pd.DataFrame({
        "管理费用项": admin_pivot.index,
        "环比增长": ((july - june) / june.replace(0, pd.NA)).fillna(0).astype(float).values,
    })
    return total, period, admin


def plot_total(total):
    fig, ax = plt.subplots(figsize=(15, 7.5))
    ax.bar(total["月份"], total["金额"], color=BLUE, width=0.62)
    ax.set_title("总财务费用对比", fontsize=15, pad=16)
    ax.set_xlabel("月份")
    ax.set_ylabel("金额")
    style_axes(ax)
    save_fig(fig, "01_总财务费用对比.png")


def plot_period(period):
    fig, ax = plt.subplots(figsize=(15, 7.5))
    ax.bar(period["费用类型"], period["环比差额"], color=BLUE, width=0.62)
    ax.axhline(0, color="#A6A6A6", linewidth=0.8)
    ax.set_title("三大期间费用7月较6月增长情况", fontsize=15, pad=16)
    ax.set_xlabel("费用类型")
    ax.set_ylabel("环比差额")
    style_axes(ax)
    save_fig(fig, "02_三大期间费用环比差额.png")


def plot_admin(admin):
    fig, ax = plt.subplots(figsize=(15, 7.5))
    ax.bar(admin["管理费用项"], admin["环比增长"], color=BLUE, width=0.62)
    ax.axhline(0, color="#A6A6A6", linewidth=0.8)
    ax.set_title("管理费用7月环比增长分析", fontsize=15, pad=16)
    ax.set_xlabel("管理费用项")
    ax.set_ylabel("环比增长")
    ax.tick_params(axis="x", labelrotation=45)
    style_axes(ax, percent=True)
    save_fig(fig, "03_管理费用环比增长.png")


def plot_dashboard(total, period, admin):
    fig, axes = plt.subplots(3, 1, figsize=(16, 19))

    axes[0].bar(total["月份"], total["金额"], color=BLUE, width=0.62)
    axes[0].set_title("总财务费用对比", fontsize=14, pad=12)
    axes[0].set_xlabel("月份")
    axes[0].set_ylabel("金额")
    style_axes(axes[0])

    axes[1].bar(period["费用类型"], period["环比差额"], color=BLUE, width=0.62)
    axes[1].axhline(0, color="#A6A6A6", linewidth=0.8)
    axes[1].set_title("三大期间费用7月较6月增长情况", fontsize=14, pad=12)
    axes[1].set_xlabel("费用类型")
    axes[1].set_ylabel("环比差额")
    style_axes(axes[1])

    axes[2].bar(admin["管理费用项"], admin["环比增长"], color=BLUE, width=0.62)
    axes[2].axhline(0, color="#A6A6A6", linewidth=0.8)
    axes[2].set_title("管理费用7月环比增长分析", fontsize=14, pad=12)
    axes[2].set_xlabel("管理费用项")
    axes[2].set_ylabel("环比增长")
    axes[2].tick_params(axis="x", labelrotation=45)
    style_axes(axes[2], percent=True)

    fig.suptitle("财务费用分析", fontsize=20, y=0.995)
    fig.subplots_adjust(top=0.965, bottom=0.055, hspace=0.48)
    save_fig(fig, "finance_expense_dashboard.png")


def main():
    set_chinese_font()
    if not COMPLETE_FILE.exists() or not RAW_FILE.exists():
        raise FileNotFoundError("请检查 COMPLETE_FILE 和 RAW_FILE 中的 Excel 路径。")

    if USE_COMPLETE_CASE_VALUES:
        total, period, admin = build_from_complete_case()
        print("当前模式：完整案例准备区，优先保证与 Excel 图表一致。")
    else:
        total, period, admin = build_from_raw_data()
        print("当前模式：原始数据重算，结果可能与完整案例不同。")

    plot_total(total)
    plot_period(period)
    plot_admin(admin)
    plot_dashboard(total, period, admin)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUTPUT_DIR / "analysis_tables.xlsx", engine="openpyxl") as writer:
        total.to_excel(writer, sheet_name="总费用对比", index=False)
        period.to_excel(writer, sheet_name="三大期间费用", index=False)
        admin.to_excel(writer, sheet_name="管理费用环比", index=False)
    print(f"完成，输出目录：{OUTPUT_DIR.resolve()}")


if __name__ == "__main__":
    main()
