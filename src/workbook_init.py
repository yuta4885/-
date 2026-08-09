# Excelブックの「入力用シート」（設定・社員マスタ・所属歴・免許資格・
# 社外教育・社内教育・品質不具合歴・スキル習熟度）の雛形を作成するファイル
# 見出しの書式、日付の表示形式、ドロップダウン選択肢、自動計算の数式、
# 条件付き書式（色分け）をここでまとめて設定する

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule

from src import calculations, config

HEADER_FILL = PatternFill("solid", fgColor="FF305496")
HEADER_FONT = Font(color="FFFFFFFF", bold=True)
THIN_BORDER = Border(*(Side(style="thin", color="FFBFBFBF"),) * 4)
NOTE_FONT = Font(italic=True, color="FF808080", size=9)
DATE_FORMAT = "yyyy/mm"

WARN_FILL = PatternFill("solid", fgColor="FFFFC7CE")     # 赤：期限切れ
ALERT_FILL = PatternFill("solid", fgColor="FFFFD9B3")    # オレンジ：3か月以内
CAUTION_FILL = PatternFill("solid", fgColor="FFFFEB9C")  # 黄：6か月以内


def _write_headers(ws, headers):
    for col, title in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=title)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(col)].width = max(12, len(title) + 2)
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 28


def _write_rows(ws, headers, rows, date_headers=()):
    for r, row in enumerate(rows, start=2):
        for c, header in enumerate(headers, start=1):
            value = row.get(header)
            cell = ws.cell(row=r, column=c, value=value)
            cell.border = THIN_BORDER
            if header in date_headers and value not in (None, ""):
                cell.number_format = DATE_FORMAT


def _add_validation(ws, col_index, choices, max_row, allow_blank=True):
    dv = DataValidation(
        type="list",
        formula1='"' + ",".join(choices) + '"',
        allow_blank=allow_blank,
        showDropDown=False,
    )
    ws.add_data_validation(dv)
    col_letter = get_column_letter(col_index)
    dv.add(f"{col_letter}2:{col_letter}{max_row}")


def _header_index(headers, name):
    return headers.index(name) + 1


def _add_note(ws, text, row=None):
    # 1列目（社員番号など各シートのキー列）に書き込むと、次回読み込み時に
    # ダミーのデータ行として誤認識されるため、必ず2列目以降に書き込む
    row = row or (ws.max_row + 2)
    cell = ws.cell(row=row, column=2, value=text)
    cell.font = NOTE_FONT
    return row


def settings_target_table_rows(skill_items):
    """設定シート内の『等級別目標スキルレベル表』の行位置（ヘッダー行、等級の開始行・終了行）を返す。
    build_settings_sheetのレイアウトと必ず一致させること（個人詳細シートの数式生成でも使用）"""
    note_row = len(skill_items) + 3
    table_top = note_row + 2
    header_row = table_top + 1
    grade_row_start = header_row + 1
    grade_row_end = header_row + len(config.GRADE_CHOICES)
    return header_row, grade_row_start, grade_row_end


def build_settings_sheet(wb, skill_items, grade_targets=None):
    ws = wb.create_sheet(config.SHEET_SETTINGS)
    _write_headers(ws, config.SETTINGS_HEADERS)
    for i, item in enumerate(skill_items, start=2):
        ws.cell(row=i, column=1, value=item).border = THIN_BORDER
    ws.column_dimensions["A"].width = 24
    _add_note(
        ws,
        "※この列に作業項目を追加・削除すると、次回スクリプト実行時にスキル習熟度シートや一覧表に反映されます。",
        row=len(skill_items) + 3,
    )

    # 等級別目標スキルレベル表（GAP分析用の目標値。現場の実情に合わせて自由に書き換え可能）
    header_row, grade_row_start, grade_row_end = settings_target_table_rows(skill_items)
    table_top = header_row - 1
    title_cell = ws.cell(row=table_top, column=1, value=config.SETTINGS_TARGET_TITLE)
    title_cell.font = Font(bold=True, size=10, color="FF1F3864")
    ws.merge_cells(start_row=table_top, start_column=1, end_row=table_top,
                    end_column=len(skill_items) + 1)

    target_headers = ["等級"] + list(skill_items)
    for col, title in enumerate(target_headers, start=1):
        cell = ws.cell(row=header_row, column=col, value=title)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    grade_targets = grade_targets or {}
    max_row = header_row
    for r, grade in enumerate(config.GRADE_CHOICES, start=header_row + 1):
        ws.cell(row=r, column=1, value=grade).border = THIN_BORDER
        targets_for_grade = grade_targets.get(grade, {})
        for c, item in enumerate(skill_items, start=2):
            level = targets_for_grade.get(item, 0)
            cell = ws.cell(row=r, column=c, value=level)
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center")
            dv = DataValidation(type="whole", operator="between", formula1=0, formula2=5, allow_blank=True)
            ws.add_data_validation(dv)
            dv.add(cell.coordinate)
        max_row = r
    _add_note(ws, "※等級ごとの目標スキルレベル（0〜5）です。個人詳細シートのGAP分析で「実効レベルー目標」として使われます。",
              row=max_row + 2)
    return ws


def build_employee_sheet(wb, employees):
    headers = config.EMPLOYEE_HEADERS
    ws = wb.create_sheet(config.SHEET_EMPLOYEE)
    _write_headers(ws, headers)
    _write_rows(ws, headers, employees, date_headers={"生年月", "入社年月"})

    age_col = _header_index(headers, "年齢")
    birth_col = _header_index(headers, "生年月")
    tenure_col = _header_index(headers, "勤続年数")
    join_col = _header_index(headers, "入社年月")
    max_row = max(2, len(employees) + 1)

    for r in range(2, max_row + 1):
        birth_cell = f"{get_column_letter(birth_col)}{r}"
        join_cell = f"{get_column_letter(join_col)}{r}"
        ws.cell(row=r, column=age_col, value=calculations.age_formula(birth_cell)).border = THIN_BORDER
        ws.cell(row=r, column=tenure_col, value=calculations.tenure_formula(join_cell)).border = THIN_BORDER

    _add_validation(ws, _header_index(headers, "性別"), config.GENDER_CHOICES, max_row)
    _add_validation(ws, _header_index(headers, "雇用区分"), config.EMPLOYMENT_TYPE_CHOICES, max_row)
    _add_validation(ws, _header_index(headers, "現在の所属"), config.DEPARTMENT_CHOICES, max_row)
    _add_validation(ws, _header_index(headers, "班"), config.TEAM_CHOICES, max_row)
    _add_validation(ws, _header_index(headers, "等級"), config.GRADE_CHOICES, max_row)
    _add_note(ws, "※「年齢」「勤続年数」は自動計算です（生年月・入社年月から算出、直接入力しないでください）。")
    return ws


def build_assignment_sheet(wb, assignments):
    headers = config.ASSIGNMENT_HEADERS
    ws = wb.create_sheet(config.SHEET_ASSIGNMENT)
    _write_headers(ws, headers)
    _write_rows(ws, headers, assignments, date_headers={"開始年月", "終了年月"})

    period_col = _header_index(headers, "在籍期間（年）")
    start_col = _header_index(headers, "開始年月")
    end_col = _header_index(headers, "終了年月")
    max_row = max(2, len(assignments) + 1)

    for r in range(2, max_row + 1):
        start_cell = f"{get_column_letter(start_col)}{r}"
        end_cell = f"{get_column_letter(end_col)}{r}"
        ws.cell(row=r, column=period_col,
                value=calculations.assignment_period_formula(start_cell, end_cell)).border = THIN_BORDER

    _add_validation(ws, _header_index(headers, "事業所・部署"), config.DEPARTMENT_CHOICES, max_row)
    _add_validation(ws, _header_index(headers, "班"), config.TEAM_CHOICES, max_row)
    _add_validation(ws, _header_index(headers, "担当業務"), config.DUTY_CHOICES, max_row)
    _add_validation(ws, _header_index(headers, "役割"), config.ROLE_CHOICES, max_row)
    _add_note(ws, "※終了年月が空欄の行は「現在も在籍中」として扱われ、在籍期間は自動計算されます。")
    return ws


def build_license_sheet(wb, licenses):
    headers = config.LICENSE_HEADERS
    ws = wb.create_sheet(config.SHEET_LICENSE)
    _write_headers(ws, headers)
    _write_rows(ws, headers, licenses, date_headers={"取得年月", "有効期限"})

    renewal_col = _header_index(headers, "更新要否")
    expiry_col = _header_index(headers, "有効期限")
    max_row = max(2, len(licenses) + 1)

    for r in range(2, max_row + 1):
        expiry_cell = f"{get_column_letter(expiry_col)}{r}"
        ws.cell(row=r, column=renewal_col,
                value=calculations.license_renewal_formula(expiry_cell)).border = THIN_BORDER

    _add_validation(ws, _header_index(headers, "実務での使用状況"), config.LICENSE_USAGE_CHOICES, max_row)

    # 有効期限が近い／切れている行を3段階で色分け（条件付き書式）
    expiry_letter = get_column_letter(expiry_col)
    rng = f"{expiry_letter}2:{expiry_letter}{max_row}"
    ws.conditional_formatting.add(
        rng, FormulaRule(formula=[f'{expiry_letter}2<TODAY()'], fill=WARN_FILL))
    ws.conditional_formatting.add(
        rng, FormulaRule(formula=[f'AND({expiry_letter}2>=TODAY(),{expiry_letter}2<=TODAY()+{config.LICENSE_WARNING_DAYS})'],
                          fill=ALERT_FILL))
    ws.conditional_formatting.add(
        rng, FormulaRule(formula=[f'AND({expiry_letter}2>TODAY()+{config.LICENSE_WARNING_DAYS},'
                                   f'{expiry_letter}2<=TODAY()+{config.LICENSE_CAUTION_DAYS})'],
                          fill=CAUTION_FILL))
    _add_note(ws, "※「更新要否」は自動計算です。有効期限セルは赤＝期限切れ、オレンジ＝3か月以内、黄＝6か月以内を表します。")
    return ws


def build_external_edu_sheet(wb, rows):
    headers = config.EXTERNAL_EDU_HEADERS
    ws = wb.create_sheet(config.SHEET_EXTERNAL_EDU)
    _write_headers(ws, headers)
    _write_rows(ws, headers, rows, date_headers={"受講年月"})
    max_row = max(2, len(rows) + 1)
    _add_validation(ws, _header_index(headers, "修了証の有無"), config.CERTIFICATE_CHOICES, max_row)
    return ws


def build_internal_edu_sheet(wb, rows):
    headers = config.INTERNAL_EDU_HEADERS
    ws = wb.create_sheet(config.SHEET_INTERNAL_EDU)
    _write_headers(ws, headers)
    _write_rows(ws, headers, rows, date_headers={"実施年月"})
    max_row = max(2, len(rows) + 1)
    _add_validation(ws, _header_index(headers, "修了区分"), config.INTERNAL_EDU_STATUS_CHOICES, max_row)

    status_col = get_column_letter(_header_index(headers, "修了区分"))
    rng = f"{status_col}2:{status_col}{max_row}"
    ws.conditional_formatting.add(
        rng, CellIsRule(operator="equal", formula=['"未受講"'], fill=WARN_FILL))
    _add_note(ws, "※「未受講」の行は赤で強調表示されます。抽出結果シートでも一覧化されます。")
    return ws


def build_quality_sheet(wb, rows, skill_items):
    headers = config.QUALITY_HEADERS
    ws = wb.create_sheet(config.SHEET_QUALITY)
    ws.cell(row=1, column=1,
            value="※このシートは育成・再発防止のための記録です。個人を責めるためのものではありません。"
                  "「見直し対象スキル項目」を設定すると、再認定日まで該当スキルの評価が一時的に引き下げられ、"
                  "一覧表・抽出結果・個人詳細に反映されます。").font = Font(bold=True, color="FFC00000")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    ws.row_dimensions[1].height = 28
    for col, title in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=col, value=title)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(col)].width = max(12, len(title) + 2)
    ws.freeze_panes = "A3"
    for r, row in enumerate(rows, start=3):
        for c, header in enumerate(headers, start=1):
            value = row.get(header)
            cell = ws.cell(row=r, column=c, value=value)
            cell.border = THIN_BORDER
            if header in ("発生年月", "再認定日") and value not in (None, ""):
                cell.number_format = DATE_FORMAT
    max_row = max(3, len(rows) + 2)
    _add_validation(ws, _header_index(headers, "区分"), config.QUALITY_CATEGORY_CHOICES, max_row)
    _add_validation(ws, _header_index(headers, "原因"), config.QUALITY_CAUSE_CHOICES, max_row)
    _add_validation(ws, _header_index(headers, "再発の有無"), config.YES_NO_CHOICES, max_row)
    _add_validation(ws, _header_index(headers, "見直し対象スキル項目"), skill_items, max_row)

    review_level_col = _header_index(headers, "見直し後レベル")
    dv = DataValidation(type="whole", operator="between", formula1=0, formula2=5, allow_blank=True)
    ws.add_data_validation(dv)
    review_level_letter = get_column_letter(review_level_col)
    dv.add(f"{review_level_letter}3:{review_level_letter}{max_row}")

    _add_note(ws, "※直近1年・3年の件数集計や繰り返し傾向は、個人詳細シートに自動反映されます。", row=max_row + 2)
    _add_note(ws, "※「再認定日」が未入力または今日より先の日付の間は、見直し中として扱われます。",
              row=max_row + 3)
    return ws


def build_skill_sheet(wb, employees, skill_items, skills):
    headers = config.SKILL_FIXED_HEADERS + list(skill_items)
    ws = wb.create_sheet(config.SHEET_SKILL)
    _write_headers(ws, headers)

    for r, emp in enumerate(employees, start=2):
        emp_id = emp["社員番号"]
        ws.cell(row=r, column=1, value=emp_id).border = THIN_BORDER
        ws.cell(row=r, column=2, value=emp["氏名"]).border = THIN_BORDER
        emp_skills = skills.get(emp_id, {})
        for c, item in enumerate(skill_items, start=len(config.SKILL_FIXED_HEADERS) + 1):
            level = emp_skills.get(item, 0)
            cell = ws.cell(row=r, column=c, value=level)
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center")
            cell.fill = PatternFill("solid", fgColor=config.SKILL_FILL_COLORS.get(level, "FFFFFFFF"))

    max_row = max(2, len(employees) + 1)
    for c in range(len(config.SKILL_FIXED_HEADERS) + 1, len(headers) + 1):
        dv = DataValidation(type="whole", operator="between", formula1=0, formula2=5, allow_blank=True)
        ws.add_data_validation(dv)
        col_letter = get_column_letter(c)
        dv.add(f"{col_letter}2:{col_letter}{max_row}")
        # 数値に応じて色分け（値を直接入力し直したときにも追随するよう条件付き書式でも設定）
        for level, color in config.SKILL_FILL_COLORS.items():
            ws.conditional_formatting.add(
                f"{col_letter}2:{col_letter}{max_row}",
                CellIsRule(operator="equal", formula=[str(level)], fill=PatternFill("solid", fgColor=color)),
            )
    _add_note(ws, "※習熟度: 0=未経験 1=教育を受ければできる 2=指示を受ければ一人でできる "
                  "3=通常作業を一人で完結できる 4=異常対応もできる 5=人に教えられる",
              row=max_row + 2)
    return ws


def build_input_sheets(wb, data, skill_items):
    """入力用の全シートを作成する"""
    build_settings_sheet(wb, skill_items, data.get("grade_targets"))
    build_employee_sheet(wb, data["employees"])
    build_assignment_sheet(wb, data["assignments"])
    build_license_sheet(wb, data["licenses"])
    build_external_edu_sheet(wb, data["external_edu"])
    build_internal_edu_sheet(wb, data["internal_edu"])
    build_quality_sheet(wb, data["quality"], skill_items)
    build_skill_sheet(wb, data["employees"], skill_items, data["skills"])
