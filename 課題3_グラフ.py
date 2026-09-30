import csv
import math
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

BASE_DIR = Path(__file__).parent
CSV_PATH = BASE_DIR / "課題3.csv"
XLSX_PATH = BASE_DIR / "グラフ.xlsx"

# CSVを読み込む
with open(CSV_PATH, encoding="utf-8-sig", newline="") as f:
    rows = [{"名前": r["名前"], "所属": r["所属"], "スコア": int(r["スコア"])} for r in csv.DictReader(f)]

# ① 所属ごとの参加者数（登場順を維持）
dept_counts = {}
for r in rows:
    dept_counts[r["所属"]] = dept_counts.get(r["所属"], 0) + 1

# ② 参加者ごとの平均スコア
person_scores = {}
for r in rows:
    person_scores.setdefault(r["名前"], []).append(r["スコア"])
person_avg = {name: round(sum(v) / len(v), 1) for name, v in person_scores.items()}

# ③ スコア分布（スタージェスの公式でビン数の目安を出し、5点刻みの区間にする）
scores = [r["スコア"] for r in rows]
sturges = math.ceil(1 + math.log2(len(scores)))
low = min(scores) // 5 * 5
high = max(scores) // 5 * 5 + 5
width = max(5, math.ceil((high - low) / sturges / 5) * 5)
bins = []
for start in range(low, high, width):
    end = start + width - 1
    bins.append((f"{start}〜{end}点", sum(start <= s <= end for s in scores)))

# 書式
header_fill = PatternFill(fill_type="solid", start_color="4472C4", end_color="4472C4")
header_font = Font(bold=True, color="FFFFFF")
thin = Side(style="thin", color="999999")
border = Border(left=thin, right=thin, top=thin, bottom=thin)


def write_table(ws, headers, data, widths):
    """シートに表を書き込み、見出しの色・罫線・列幅を整える"""
    ws.append(headers)
    for row in data:
        ws.append(list(row))
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=len(headers)):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if cell.row == 1:
                cell.fill = header_fill
                cell.font = header_font
    for col, w in zip("ABC", widths):
        ws.column_dimensions[col].width = w


wb = Workbook()

# ① 円グラフ：所属ごとの参加者数
ws1 = wb.active
ws1.title = "所属別参加者数"
write_table(ws1, ["所属", "参加者数"], dept_counts.items(), [12, 12])

pie = PieChart()
pie.title = "所属ごとの参加者数"
pie.add_data(Reference(ws1, min_col=2, min_row=1, max_row=ws1.max_row), titles_from_data=True)
pie.set_categories(Reference(ws1, min_col=1, min_row=2, max_row=ws1.max_row))
pie.dataLabels = DataLabelList()
pie.dataLabels.showCatName = True   # 所属名
pie.dataLabels.showPercent = True   # パーセンテージ
pie.dataLabels.showVal = False
pie.dataLabels.showSerName = False
pie.dataLabels.showLeaderLines = True
pie.dataLabels.separator = "\n"
pie.dataLabels.position = "bestFit"
pie.legend.position = "r"
pie.legend.overlay = False
pie.width, pie.height = 16, 10
ws1.add_chart(pie, "D2")

# ② 棒グラフ：参加者ごとの平均スコア
ws2 = wb.create_sheet("参加者別平均スコア")
write_table(ws2, ["名前", "平均スコア"], person_avg.items(), [12, 12])
for (cell,) in ws2.iter_rows(min_row=2, min_col=2, max_col=2):
    cell.number_format = "0.0"

bar = BarChart()
bar.type = "col"
bar.title = "参加者ごとの平均スコア"
bar.x_axis.title = "名前"
bar.y_axis.title = "スコア（点）"
bar.add_data(Reference(ws2, min_col=2, min_row=1, max_row=ws2.max_row), titles_from_data=True)
bar.set_categories(Reference(ws2, min_col=1, min_row=2, max_row=ws2.max_row))
bar.y_axis.scaling.min = 0
bar.y_axis.scaling.max = 100
bar.y_axis.majorUnit = 10
bar.x_axis.delete = False
bar.y_axis.delete = False
bar.x_axis.tickLblSkip = 1  # 全員の名前を表示
bar.varyColors = False  # 全員同じ色にする
bar.y_axis.number_format = "0"
bar.legend.position = "r"
bar.legend.overlay = False
bar.gapWidth = 50
bar.width, bar.height = 30, 12
ws2.add_chart(bar, "D2")

# ③ ヒストグラム：スコアの分布
ws3 = wb.create_sheet("スコア分布")
write_table(ws3, ["スコア区間", "人数"], bins, [14, 10])

hist = BarChart()
hist.type = "col"
hist.title = f"スコアの分布（{width}点刻み・{len(bins)}区間）"
hist.x_axis.title = "スコア区間"
hist.y_axis.title = "人数（人）"
hist.add_data(Reference(ws3, min_col=2, min_row=1, max_row=ws3.max_row), titles_from_data=True)
hist.set_categories(Reference(ws3, min_col=1, min_row=2, max_row=ws3.max_row))
hist.gapWidth = 5  # 棒の隙間を詰めてヒストグラムらしくする
hist.y_axis.scaling.min = 0
hist.y_axis.majorUnit = 2
hist.x_axis.delete = False
hist.y_axis.delete = False
hist.dataLabels = DataLabelList()
hist.dataLabels.showVal = True  # 棒の上に人数だけを表示
hist.dataLabels.showCatName = False
hist.dataLabels.showSerName = False
hist.dataLabels.showLegendKey = False
hist.dataLabels.position = "outEnd"
hist.varyColors = False
hist.legend.position = "r"
hist.legend.overlay = False
hist.width, hist.height = 18, 10
ws3.add_chart(hist, "D2")

# タイトル・軸ラベルがグラフ本体に重ならないようにする
for chart in (pie, bar, hist):
    chart.title.overlay = False
for chart in (bar, hist):
    chart.x_axis.title.overlay = False
    chart.y_axis.title.overlay = False

wb.save(XLSX_PATH)

print(f"{XLSX_PATH.name} を作成しました")
print("所属別:", dept_counts)
print(f"ビン数: スタージェスの目安 {sturges} → {width}点刻み {len(bins)}区間")
print("分布:", bins)
