# 「一覧表」シート（自動生成・編集不可）を作成するファイル
# 全員×主要スキルの色分けマトリクスと、人員配置の穴チェック（属人化リスク）を出力する

from openpyxl.styles import Alignment, Font, PatternFill, Side, Border
from openpyxl.utils import get_column_letter

from src import config, gap_check

HEADER_FILL = PatternFill("solid", fgColor="FF305496")
HEADER_FONT = Font(color="FFFFFFFF", bold=True)
WARN_FILL = PatternFill("solid", fgColor="FFFFC7CE")
THIN_BORDER = Border(*(Side(style="thin", color="FFBFBFBF"),) * 4)
TITLE_FONT = Font(bold=True, size=13)


def build_list_sheet(wb, employees, skills, skill_items):
    if config.SHEET_LIST in wb.sheetnames:
        del wb[config.SHEET_LIST]
    ws = wb.create_sheet(config.SHEET_LIST)

    ws.cell(row=1, column=1, value="一覧表（自動生成・このシートは編集しないでください。次回実行時に上書きされます）").font = TITLE_FONT

    base_headers = ["社員番号", "氏名", "現在の所属", "班"]
    headers = base_headers + list(skill_items)
    header_row = 3
    for col, title in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col, value=title)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(col)].width = 10 if col > len(base_headers) else 14
    ws.freeze_panes = get_column_letter(len(base_headers) + 1) + str(header_row + 1)
    ws.row_dimensions[header_row].height = 40

    for r, emp in enumerate(employees, start=header_row + 1):
        emp_id = emp["社員番号"]
        ws.cell(row=r, column=1, value=emp_id).border = THIN_BORDER
        ws.cell(row=r, column=2, value=emp.get("氏名")).border = THIN_BORDER
        ws.cell(row=r, column=3, value=emp.get("現在の所属")).border = THIN_BORDER
        ws.cell(row=r, column=4, value=emp.get("班")).border = THIN_BORDER
        levels = skills.get(emp_id, {})
        for c, item in enumerate(skill_items, start=len(base_headers) + 1):
            level = levels.get(item, 0)
            cell = ws.cell(row=r, column=c, value=level)
            cell.alignment = Alignment(horizontal="center")
            cell.border = THIN_BORDER
            cell.fill = PatternFill("solid", fgColor=config.SKILL_FILL_COLORS.get(level, "FFFFFFFF"))

    last_data_row = header_row + max(len(employees), 1)

    # 凡例
    legend_row = last_data_row + 2
    ws.cell(row=legend_row, column=1, value="凡例：").font = Font(bold=True)
    for i, (level, color) in enumerate(config.SKILL_FILL_COLORS.items()):
        col = 2 + i
        cell = ws.cell(row=legend_row, column=col, value=f"{level}:{config.SKILL_LEVELS[level]}")
        cell.fill = PatternFill("solid", fgColor=color)
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(col)].width = 14

    # 属人化リスク（穴チェック）
    gap_title_row = legend_row + 3
    ws.cell(row=gap_title_row, column=1,
            value="■ 人員配置の穴チェック（属人化リスク：この作業ができる人が1人以下の項目）").font = TITLE_FONT

    gap_headers = ["作業項目", "Lv3以上人数", "該当者", "属人化リスク"]
    gap_header_row = gap_title_row + 1
    for col, title in enumerate(gap_headers, start=1):
        cell = ws.cell(row=gap_header_row, column=col, value=title)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")

    gap_rows = gap_check.compute_gap_warnings(employees, skills, skill_items)
    for r, row in enumerate(gap_rows, start=gap_header_row + 1):
        for c, header in enumerate(gap_headers, start=1):
            cell = ws.cell(row=r, column=c, value=row[header])
            cell.border = THIN_BORDER
            if row["属人化リスク"]:
                cell.fill = WARN_FILL

    ws.protection.sheet = True
    ws.protection.formatCells = False
    return ws
