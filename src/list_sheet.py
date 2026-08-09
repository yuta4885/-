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


def build_list_sheet(wb, employees, effective_skills, skill_items):
    """一覧表シートを作成する。effective_skillsは品質不具合による一時的な見直しを
    反映済みの実効レベル（skill_review.compute_effective_skillsの戻り値）を渡すこと"""
    if config.SHEET_LIST in wb.sheetnames:
        del wb[config.SHEET_LIST]
    ws = wb.create_sheet(config.SHEET_LIST)

    ws.cell(row=1, column=1, value="一覧表（自動生成・このシートは編集しないでください。次回実行時に上書きされます）").font = TITLE_FONT
    ws.cell(row=2, column=1,
            value="※品質不具合による「スキルレベルの一時的な見直し」を反映した実効レベルで集計しています。"
            ).font = Font(italic=True, size=9, color="FF808080")

    n_employees = max(len(employees), 1)
    n_items = max(len(skill_items), 1)
    total_qualified = sum(
        1
        for emp in employees
        for item in skill_items
        if effective_skills.get(emp["社員番号"], {}).get(item, 0) >= config.SKILL_MATRIX_WARNING_LEVEL
    )
    tahnoko_rate = total_qualified / (n_employees * n_items)
    rate_cell = ws.cell(row=3, column=1,
                         value=f"多能工化率（全員×全工程のうちLv3以上の割合）： {tahnoko_rate:.0%}")
    rate_cell.font = Font(bold=True, size=10, color="FF1F3864")

    base_headers = ["社員番号", "氏名", "現在の所属", "班"]
    extra_headers = ["対応可能工程数(Lv3+)"]
    headers = base_headers + list(skill_items) + extra_headers
    header_row = 5
    for col, title in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col, value=title)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        is_extra = col > len(base_headers) + len(skill_items)
        ws.column_dimensions[get_column_letter(col)].width = 14 if (col <= len(base_headers) or is_extra) else 10
    ws.freeze_panes = get_column_letter(len(base_headers) + 1) + str(header_row + 1)
    ws.row_dimensions[header_row].height = 40

    skill_count_col = len(base_headers) + len(skill_items) + 1
    for r, emp in enumerate(employees, start=header_row + 1):
        emp_id = emp["社員番号"]
        ws.cell(row=r, column=1, value=emp_id).border = THIN_BORDER
        ws.cell(row=r, column=2, value=emp.get("氏名")).border = THIN_BORDER
        ws.cell(row=r, column=3, value=emp.get("現在の所属")).border = THIN_BORDER
        ws.cell(row=r, column=4, value=emp.get("班")).border = THIN_BORDER
        levels = effective_skills.get(emp_id, {})
        for c, item in enumerate(skill_items, start=len(base_headers) + 1):
            level = levels.get(item, 0)
            cell = ws.cell(row=r, column=c, value=level)
            cell.alignment = Alignment(horizontal="center")
            cell.border = THIN_BORDER
            cell.fill = PatternFill("solid", fgColor=config.SKILL_FILL_COLORS.get(level, "FFFFFFFF"))

        qualified_count = sum(1 for lv in levels.values() if lv >= config.SKILL_MATRIX_WARNING_LEVEL)
        count_cell = ws.cell(row=r, column=skill_count_col, value=qualified_count)
        count_cell.alignment = Alignment(horizontal="center")
        count_cell.font = Font(bold=True)
        count_cell.border = THIN_BORDER

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

    gap_rows = gap_check.compute_gap_warnings(employees, effective_skills, skill_items)
    for r, row in enumerate(gap_rows, start=gap_header_row + 1):
        for c, header in enumerate(gap_headers, start=1):
            cell = ws.cell(row=r, column=c, value=row[header])
            cell.border = THIN_BORDER
            if row["属人化リスク"]:
                cell.fill = WARN_FILL

    ws.protection.sheet = True
    ws.protection.formatCells = False
    return ws
