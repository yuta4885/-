# 「個人シート」（自動生成・編集不可、社員1人につき1枚）を作成するファイル
# 基本情報／所属歴／資格・免許／教育歴／習熟度／品質不具合歴／育成方針を
# A4縦1枚に収まる体裁でまとめ、面談などにそのまま印刷して使えるようにする

from datetime import date

from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

from src import calculations, config

SECTION_FILL = PatternFill("solid", fgColor="FF305496")
SECTION_FONT = Font(color="FFFFFFFF", bold=True, size=10)
NAME_FONT = Font(bold=True, size=16)
SUB_FONT = Font(size=9, color="FF595959")
THIN_BORDER = Border(*(Side(style="thin", color="FFBFBFBF"),) * 4)
TABLE_HEADER_FILL = PatternFill("solid", fgColor="FFD9E1F2")

TOTAL_COLS = 8


def _section_title(ws, row, text):
    ws.cell(row=row, column=1, value=text).fill = SECTION_FILL
    for c in range(1, TOTAL_COLS + 1):
        ws.cell(row=row, column=c).fill = SECTION_FILL
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = SECTION_FONT
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=TOTAL_COLS)
    ws.row_dimensions[row].height = 16
    return row + 1


def _mini_table(ws, row, headers, records, empty_text="記録なし", date_cols=()):
    if not records:
        ws.cell(row=row, column=1, value=empty_text).font = SUB_FONT
        return row + 1

    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=col, value=header)
        cell.fill = TABLE_HEADER_FILL
        cell.font = Font(bold=True, size=8)
        cell.border = THIN_BORDER
    row += 1
    for record in records:
        for col, header in enumerate(headers, start=1):
            value = record.get(header, "")
            cell = ws.cell(row=row, column=col, value=value)
            cell.font = Font(size=8)
            cell.border = THIN_BORDER
            if header in date_cols and value not in (None, ""):
                cell.number_format = "yyyy/mm"
        row += 1
    return row


def compute_quality_summary(quality_records, emp_id):
    """品質不具合歴から直近1年・3年の件数と繰り返し傾向を計算する"""
    today = date.today()
    my_records = [r for r in quality_records if r.get("社員番号") == emp_id]

    def within_days(record, days):
        occurred = calculations.to_date(record.get("発生年月"))
        if occurred is None:
            return False
        return (today - occurred).days <= days

    count_1y = sum(1 for r in my_records if within_days(r, 365))
    count_3y = sum(1 for r in my_records if within_days(r, 365 * 3))

    category_counts = {}
    for r in my_records:
        if within_days(r, 365 * 3):
            cat = r.get("区分", "")
            category_counts[cat] = category_counts.get(cat, 0) + 1
    repeated = [cat for cat, cnt in category_counts.items() if cnt >= 2]

    if repeated:
        trend = "同じ区分の繰り返しあり：" + "、".join(repeated)
    elif my_records:
        trend = "同区分の繰り返しなし"
    else:
        trend = "記録なし"

    return {"直近1年件数": count_1y, "直近3年件数": count_3y, "傾向": trend}


def build_individual_sheet(wb, emp, assignments, licenses, external_edu, internal_edu,
                            quality, skills, skill_items):
    emp_id = emp["社員番号"]
    sheet_name = f"{config.INDIVIDUAL_SHEET_PREFIX}{emp_id}"
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)

    for c in range(1, TOTAL_COLS + 1):
        ws.column_dimensions[get_column_letter(c)].width = 11

    row = 1
    ws.cell(row=row, column=1, value=f"{emp.get('氏名', '')}（{emp.get('氏名カナ', '')}）").font = NAME_FONT
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
    ws.cell(row=row, column=6, value=f"社員番号: {emp_id}").font = SUB_FONT
    ws.merge_cells(start_row=row, start_column=6, end_row=row, end_column=8)
    row += 1
    ws.cell(row=row, column=1,
            value="※本シートは人員配置・育成計画のための資料です。人事評価・処遇決定には使用しません。").font = SUB_FONT
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    row += 2

    # 基本情報
    row = _section_title(ws, row, "基本情報")
    basic_pairs = [
        ("年齢", f"{calculations.years_between(emp.get('生年月'))}歳" if emp.get("生年月") else ""),
        ("性別", emp.get("性別", "")),
        ("雇用区分", emp.get("雇用区分", "")),
        ("勤続年数", f"{calculations.years_between(emp.get('入社年月'))}年" if emp.get("入社年月") else ""),
        ("現在の所属", emp.get("現在の所属", "")),
        ("班", emp.get("班", "")),
        ("役職", emp.get("役職", "")),
        ("等級", emp.get("等級", "")),
    ]
    for i, (label, value) in enumerate(basic_pairs):
        c = 1 + (i % 4) * 2
        r = row + i // 4
        ws.cell(row=r, column=c, value=label).font = Font(bold=True, size=8)
        ws.cell(row=r, column=c + 1, value=value).font = Font(size=8)
    row += (len(basic_pairs) + 3) // 4
    ws.cell(row=row, column=1, value="本人の希望：").font = Font(bold=True, size=8)
    ws.cell(row=row, column=2, value=emp.get("本人の希望（今後やりたい仕事）", "")).font = Font(size=8)
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=8)
    row += 2

    # 所属歴
    row = _section_title(ws, row, "所属歴")
    my_assignments = [a for a in assignments if a.get("社員番号") == emp_id]
    for a in my_assignments:
        a["期間"] = (
            f"{_fmt(a.get('開始年月'))}〜{_fmt(a.get('終了年月')) or '現在'}"
        )
    row = _mini_table(ws, row, ["期間", "事業所・部署", "班", "担当業務", "役割"], my_assignments)
    row += 1

    # 免許・資格
    row = _section_title(ws, row, "免許・資格")
    my_licenses = [
        {**lic, "有効期限表示": _fmt(lic.get("有効期限"))}
        for lic in licenses if lic.get("社員番号") == emp_id
    ]
    row = _mini_table(ws, row, ["資格名", "取得年月", "有効期限表示", "実務での使用状況"], my_licenses,
                       date_cols={"取得年月"})
    row += 1

    # 教育歴（社外・社内）
    row = _section_title(ws, row, "教育歴")
    my_external = [e for e in external_edu if e.get("社員番号") == emp_id]
    my_internal = [e for e in internal_edu if e.get("社員番号") == emp_id]
    ws.cell(row=row, column=1, value="社外教育").font = Font(bold=True, size=8, italic=True)
    row += 1
    row = _mini_table(ws, row, ["講習・研修名", "実施機関", "受講年月"], my_external, date_cols={"受講年月"})
    ws.cell(row=row, column=1, value="社内教育").font = Font(bold=True, size=8, italic=True)
    row += 1
    row = _mini_table(ws, row, ["教育名", "実施年月", "修了区分"], my_internal, date_cols={"実施年月"})
    row += 1

    # スキル習熟度
    row = _section_title(ws, row, "スキル習熟度")
    my_skills = skills.get(emp_id, {})
    cols_per_line = 4
    for i, item in enumerate(skill_items):
        level = my_skills.get(item, 0)
        c = 1 + (i % cols_per_line) * 2
        r = row + i // cols_per_line
        ws.cell(row=r, column=c, value=item).font = Font(size=7)
        level_cell = ws.cell(row=r, column=c + 1, value=level)
        level_cell.font = Font(size=8, bold=True)
        level_cell.alignment = Alignment(horizontal="center")
        level_cell.fill = PatternFill("solid", fgColor=config.SKILL_FILL_COLORS.get(level, "FFFFFFFF"))
    row += (len(skill_items) + cols_per_line - 1) // cols_per_line + 1

    # 品質不具合歴
    row = _section_title(ws, row, "品質不具合歴（育成・再発防止のための記録）")
    summary = compute_quality_summary(quality, emp_id)
    ws.cell(row=row, column=1,
            value=f"直近1年: {summary['直近1年件数']}件　直近3年: {summary['直近3年件数']}件　傾向: {summary['傾向']}"
            ).font = Font(size=8, bold=True)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    row += 1
    my_quality = [q for q in quality if q.get("社員番号") == emp_id]
    row = _mini_table(ws, row, ["発生年月", "区分", "原因", "再発防止として実施した教育・対策"], my_quality,
                       date_cols={"発生年月"})
    row += 1

    # 備考・今後の育成方針
    row = _section_title(ws, row, "備考")
    ws.cell(row=row, column=1, value=emp.get("備考", "") or "").font = Font(size=8)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    row += 2

    row = _section_title(ws, row, "今後の育成方針（面談時に記入）")
    for i in range(3):
        ws.cell(row=row + i, column=1, value="").border = THIN_BORDER
        ws.merge_cells(start_row=row + i, start_column=1, end_row=row + i, end_column=8)
        ws.row_dimensions[row + i].height = 16
    row += 3

    # 印刷設定：A4縦1枚に収まるように
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.4, bottom=0.4, header=0.2, footer=0.2)
    ws.print_area = f"A1:{get_column_letter(TOTAL_COLS)}{row}"
    ws.sheet_view.showGridLines = False
    ws.protection.sheet = True
    ws.protection.formatCells = False
    return ws


def _fmt(value):
    d = calculations.to_date(value)
    return d.strftime("%Y/%m") if d else ""


def remove_all_individual_sheets(wb):
    """既存の個人シートをすべて削除する（社員が減った場合に古いシートが残らないようにするため）"""
    for name in list(wb.sheetnames):
        if name.startswith(config.INDIVIDUAL_SHEET_PREFIX):
            del wb[name]


def build_all_individual_sheets(wb, data):
    remove_all_individual_sheets(wb)
    for emp in data["employees"]:
        build_individual_sheet(
            wb, emp,
            data["assignments"], data["licenses"], data["external_edu"],
            data["internal_edu"], data["quality"], data["skills"], data["skill_items"],
        )
