# 「個人詳細」シート（自動生成・編集不可、社員全員で1枚に共通化）を作成するファイル
#
# 社員ごとにシートを分けていた旧方式をやめ、プルダウンで社員番号を選ぶと
# XLOOKUP／FILTER関数でその人のデータだけが自動表示される単一シートに統合した。
# Excel 365 / Excel 2021以降が必要（XLOOKUP・FILTERを使用するため）。
#
# 表示するデータの「本来の値」は各入力シートのままとし、このシートは
# それを検索して映すだけ＝入力シートを直接編集すれば自動で反映される。
# 品質不具合による「スキルレベルの一時的な見直し」を反映した実効レベルは、
# effective_sheet.pyが書き出す隠しシート（_有効スキル(自動計算)）をXLOOKUPで参照する
#（判定は最後にスクリプトを実行した時点のものになる。詳細はREADME参照）。
#
# ※レーダーチャートは、Excelバージョンによる表示崩れリスクを避けるため、
#   グラフオブジェクトではなくセルの色分け・データバー（Excel標準機能）で表現している。

from openpyxl.formatting.rule import CellIsRule, DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.page import PageMargins

from src import config, effective_sheet, workbook_init

# ------------------------------------------------------------
# 配色
# ------------------------------------------------------------
NAVY = "FF1F3864"
NAVY_SOFT = "FF2E5395"
ACCENT_BAR = "FFED7D31"
BLUE = "FF2E5395"
GREEN = "FF548235"
ORANGE = "FFC55A11"
PURPLE = "FF7030A0"
TEAL = "FF1F7A6C"
AMBER = "FFBF8F00"
FOREST = "FF375623"

CARD_BG = "FFF7F9FC"
CARD_BORDER_COLOR = "FFD9E2F3"
WHITE = "FFFFFFFF"
GRAY_TEXT = "FF595959"

BADGE_COLORS = [BLUE, GREEN, ORANGE, PURPLE]
THIN_BORDER = Border(*(Side(style="thin", color=CARD_BORDER_COLOR),) * 4)

TOTAL_COLS = 12
RANGE_MAX = 1000  # 各入力シートを検索する範囲の下限（将来の増員を見込んだ余裕）
SEL = "$G$2"      # 選択中の社員番号セル（絶対参照。ヘッダー行のプルダウンセル）


def _fill(color):
    return PatternFill("solid", fgColor=color)


def _sheet_ref(name):
    return f"'{name}'"


def _col_letter(headers, name):
    return get_column_letter(headers.index(name) + 1)


def _xlookup_expr(sheet, headers, key_name, return_name, header_row=1, value_ref=SEL):
    """単一値を検索するXLOOKUP式（先頭の"="を含まない、他の式に埋め込むための断片）を組み立てる"""
    key_col = _col_letter(headers, key_name)
    ret_col = _col_letter(headers, return_name)
    ref = _sheet_ref(sheet)
    start = header_row + 1
    return (f'IFERROR(XLOOKUP({value_ref},{ref}!${key_col}${start}:${key_col}${RANGE_MAX},'
            f'{ref}!${ret_col}${start}:${ret_col}${RANGE_MAX}),"")')


def _xlookup(sheet, headers, key_name, return_name, header_row=1, value_ref=SEL):
    """単一値を検索するXLOOKUP数式（セルにそのまま入れられる完全な形）を組み立てる"""
    return "=" + _xlookup_expr(sheet, headers, key_name, return_name, header_row, value_ref)


def _filter_formula(sheet, headers, key_name, return_name, header_row=1, as_date=False, value_ref=SEL):
    """条件に一致する複数行を1列分だけ取り出すFILTER数式を組み立てる（下方向にスピルする）"""
    key_col = _col_letter(headers, key_name)
    ret_col = _col_letter(headers, return_name)
    ref = _sheet_ref(sheet)
    start = header_row + 1
    inner = (f'FILTER({ref}!${ret_col}${start}:${ret_col}${RANGE_MAX},'
             f'{ref}!${key_col}${start}:${key_col}${RANGE_MAX}={value_ref})')
    if as_date:
        inner = f'TEXT({inner},"yyyy/mm")'
    return f'=IFERROR({inner},"")'


def _card_header(ws, row, title, color, col_start=1, col_end=TOTAL_COLS):
    for c in range(col_start, col_end + 1):
        ws.cell(row=row, column=c).fill = _fill(color)
    cell = ws.cell(row=row, column=col_start, value=f"  {title}")
    cell.font = Font(color=WHITE, bold=True, size=10)
    cell.alignment = Alignment(vertical="center")
    ws.merge_cells(start_row=row, start_column=col_start, end_row=row, end_column=col_end)
    ws.row_dimensions[row].height = 18
    return row + 1


def _pill(ws, row, col_start, col_end, text, color):
    for c in range(col_start, col_end + 1):
        ws.cell(row=row, column=c).fill = _fill(color)
    cell = ws.cell(row=row, column=col_start, value=text)
    cell.font = Font(color=WHITE, bold=True, size=8.5)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=row, start_column=col_start, end_row=row, end_column=col_end)


def _stat_tile(ws, num_row, label_row, col_start, col_end, number_formula, label_text, accent_color):
    for c in range(col_start, col_end + 1):
        top_cell = ws.cell(row=num_row, column=c)
        top_cell.border = Border(top=Side(style="medium", color=accent_color))
        top_cell.fill = _fill(CARD_BG)
        ws.cell(row=label_row, column=c).fill = _fill(CARD_BG)
    num_cell = ws.cell(row=num_row, column=col_start, value=number_formula)
    num_cell.font = Font(bold=True, size=17, color=accent_color)
    num_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=num_row, start_column=col_start, end_row=num_row, end_column=col_end)
    ws.row_dimensions[num_row].height = 20

    label_cell = ws.cell(row=label_row, column=col_start, value=label_text)
    label_cell.font = Font(size=8, color=GRAY_TEXT)
    label_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=label_row, start_column=col_start, end_row=label_row, end_column=col_end)
    ws.row_dimensions[label_row].height = 12


def _add_filter_section(ws, row, sheet, headers, columns, header_row=1, date_cols=(), budget_rows=6):
    """FILTER関数で複数行を表示するセクションを追加する。
    数式が入るセル（スピルの起点）は結合セルにできない（結合するとスピルがブロックされる）ため、
    項目は列を空けながら配置し、値が長い場合は右隣の空セルへはみ出して表示されるようにする"""
    n = len(columns)
    span = max(TOTAL_COLS // n, 1)
    starts = [1 + i * span for i in range(n)]

    for c0, col_name in zip(starts, columns):
        c1 = min(c0 + span - 1, TOTAL_COLS)
        header_cell = ws.cell(row=row, column=c0, value=col_name)
        header_cell.font = Font(bold=True, size=7.5)
        header_cell.fill = _fill("FFD9E1F2")
        header_cell.border = THIN_BORDER
        if c1 > c0:
            ws.merge_cells(start_row=row, start_column=c0, end_row=row, end_column=c1)

        formula = _filter_formula(sheet, headers, "社員番号", col_name, header_row=header_row,
                                   as_date=(col_name in date_cols))
        data_cell = ws.cell(row=row + 1, column=c0, value=formula)
        data_cell.font = Font(size=7.5)
        data_cell.alignment = Alignment(vertical="top")

    # データ行（FILTERのスピル先を含む）を白背景にして見やすくする
    for r in range(row + 1, row + 1 + budget_rows):
        for c in range(1, TOTAL_COLS + 1):
            ws.cell(row=r, column=c).fill = _fill(WHITE)
    return row + 1 + budget_rows


def _card_body_border(ws, row_start, row_end, col_start=1, col_end=TOTAL_COLS):
    for r in range(row_start, row_end + 1):
        for c in range(col_start, col_end + 1):
            cell = ws.cell(row=r, column=c)
            if cell.fill.fgColor.rgb in (None, "00000000"):
                cell.fill = _fill(CARD_BG)
    for c in range(col_start, col_end + 1):
        ws.cell(row=row_start, column=c).border = Border(
            top=Side(style="thin", color=CARD_BORDER_COLOR),
            left=Side(style="thin", color=CARD_BORDER_COLOR) if c == col_start else None,
            right=Side(style="thin", color=CARD_BORDER_COLOR) if c == col_end else None,
        )
        ws.cell(row=row_end, column=c).border = Border(
            bottom=Side(style="thin", color=CARD_BORDER_COLOR),
            left=Side(style="thin", color=CARD_BORDER_COLOR) if c == col_start else None,
            right=Side(style="thin", color=CARD_BORDER_COLOR) if c == col_end else None,
        )
    for r in range(row_start, row_end + 1):
        ws.cell(row=r, column=col_start).border = Border(
            left=Side(style="thin", color=CARD_BORDER_COLOR),
            top=ws.cell(row=r, column=col_start).border.top,
            bottom=ws.cell(row=r, column=col_start).border.bottom,
        )
        ws.cell(row=r, column=col_end).border = Border(
            right=Side(style="thin", color=CARD_BORDER_COLOR),
            top=ws.cell(row=r, column=col_end).border.top,
            bottom=ws.cell(row=r, column=col_end).border.bottom,
        )


def build_individual_sheet(wb, data):
    employees = data["employees"]
    skill_items = data["skill_items"]
    EMP = config.EMPLOYEE_HEADERS
    ASSIGN = config.ASSIGNMENT_HEADERS
    LIC = config.LICENSE_HEADERS
    EXT_EDU = config.EXTERNAL_EDU_HEADERS
    INT_EDU = config.INTERNAL_EDU_HEADERS
    QUAL = config.QUALITY_HEADERS

    previous_selection = None
    if config.SHEET_INDIVIDUAL in wb.sheetnames:
        previous_selection = wb[config.SHEET_INDIVIDUAL]["G2"].value
        del wb[config.SHEET_INDIVIDUAL]
    ws = wb.create_sheet(config.SHEET_INDIVIDUAL)
    ws.sheet_view.showGridLines = False
    for c in range(1, TOTAL_COLS + 1):
        ws.column_dimensions[get_column_letter(c)].width = 8.6

    # 上部アクセントバー
    for c in range(1, TOTAL_COLS + 1):
        ws.cell(row=1, column=c).fill = _fill(ACCENT_BAR)
    ws.row_dimensions[1].height = 5

    # ヘッダー（タイトル＋社員番号セレクタ）
    header_row = 2
    ws.merge_cells(start_row=header_row, start_column=1, end_row=header_row, end_column=5)
    title_cell = ws.cell(row=header_row, column=1, value="個人詳細")
    title_cell.font = Font(bold=True, size=18, color=WHITE)
    title_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    ws.merge_cells(start_row=header_row, start_column=6, end_row=header_row, end_column=6)
    label_cell = ws.cell(row=header_row, column=6, value="社員番号")
    label_cell.font = Font(size=8, color="FFBFD0EC")
    label_cell.alignment = Alignment(horizontal="center", vertical="bottom")

    name_row = header_row + 1
    ws.merge_cells(start_row=name_row, start_column=9, end_row=name_row, end_column=12)
    name_formula = _xlookup(config.SHEET_EMPLOYEE, EMP, "社員番号", "氏名")
    name_cell = ws.cell(row=name_row, column=9, value=name_formula)
    name_cell.font = Font(bold=True, size=13, color=WHITE)
    name_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    for r in (header_row, name_row):
        for c in range(1, TOTAL_COLS + 1):
            if ws.cell(row=r, column=c).fill.fgColor.rgb in (None, "00000000"):
                ws.cell(row=r, column=c).fill = _fill(NAVY)
    ws.row_dimensions[header_row].height = 24
    ws.row_dimensions[name_row].height = 22

    # 社員番号セレクタ（プルダウン）
    valid_ids = {emp["社員番号"] for emp in employees}
    if previous_selection in valid_ids:
        default_id = previous_selection
    else:
        default_id = employees[0]["社員番号"] if employees else ""
    ws.cell(row=header_row, column=7, value=default_id)
    sel_cell = ws["G2"]
    sel_cell.font = Font(bold=True, size=12, color=NAVY)
    sel_cell.fill = _fill(WHITE)
    sel_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=header_row, start_column=7, end_row=header_row, end_column=8)
    dv = DataValidation(
        type="list",
        formula1=f"{_sheet_ref(config.SHEET_EMPLOYEE)}!$A$2:$A${RANGE_MAX}",
        allow_blank=False,
    )
    ws.add_data_validation(dv)
    dv.add("G2")

    # バッジ行（所属・班・役職・雇用区分・性別・等級）
    badge_row = name_row + 1
    spans = [(1, 4), (5, 8), (9, 12)]
    badge_fields_1 = [
        ("現在の所属", "所属未設定"), ("班", "班未設定"), ("役職", "役職未設定"),
    ]
    badge_fields_2 = [
        ("雇用区分", "雇用区分："), ("性別", "性別："), ("等級", "等級："),
    ]
    for (c0, c1), (field, empty_text), color in zip(spans, badge_fields_1, BADGE_COLORS):
        base = _xlookup_expr(config.SHEET_EMPLOYEE, EMP, "社員番号", field)
        formula = f'=IF({base}="","{empty_text}",{base})'
        _pill(ws, badge_row, c0, c1, formula, color)
    for (c0, c1), (field, prefix), color in zip(spans, badge_fields_2, BADGE_COLORS[::-1]):
        base = _xlookup_expr(config.SHEET_EMPLOYEE, EMP, "社員番号", field)
        formula = f'="{prefix}"&IF({base}="","―",{base})'
        _pill(ws, badge_row + 1, c0, c1, formula, color)
    ws.row_dimensions[badge_row].height = 15
    ws.row_dimensions[badge_row + 1].height = 15

    row = badge_row + 2
    note_cell = ws.cell(row=row, column=1,
                         value="※本シートは人員配置・育成計画のための資料です。人事評価・処遇決定には使用しません。"
                               "　社員番号セレクタ（右上）を変更すると表示が切り替わります。")
    note_cell.font = Font(size=7.5, italic=True, color=GRAY_TEXT)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=TOTAL_COLS)
    row += 2

    # ------------------------------------------------------------
    # サマリー（統計タイル＋習熟度分布）
    # ------------------------------------------------------------
    row = _card_header(ws, row, "サマリー", TEAL)
    tile_top = row
    cols = effective_sheet.columns_for(skill_items)
    eff_ref = _sheet_ref(config.SHEET_EFFECTIVE)
    first_eff_col = get_column_letter(cols[skill_items[0]]["effective_col"]) if skill_items else "B"
    last_eff_col = get_column_letter(cols[skill_items[-1]]["effective_col"]) if skill_items else "B"
    row_index_expr = f'MATCH({SEL},{eff_ref}!$A$2:$A${RANGE_MAX},0)'
    skill_row_array = f'INDEX({eff_ref}!${first_eff_col}$2:${last_eff_col}${RANGE_MAX},{row_index_expr},0)'

    age_expr = _xlookup_expr(config.SHEET_EMPLOYEE, EMP, "社員番号", "年齢")
    tenure_expr = _xlookup_expr(config.SHEET_EMPLOYEE, EMP, "社員番号", "勤続年数")
    tiles = [
        (f'=IF({age_expr}="","―",{age_expr}&"歳")', "年齢", BLUE),
        (f'=IF({tenure_expr}="","―",{tenure_expr}&"年")', "勤続年数", GREEN),
        (f'=COUNTIF({_sheet_ref(config.SHEET_LICENSE)}!$A:$A,{SEL})', "保有資格数", ORANGE),
        (f'=SUMPRODUCT(({skill_row_array}>=4)*1)', "Lv4以上のスキル数(実効)", PURPLE),
    ]
    positions = [(1, 4, tile_top), (5, 8, tile_top), (1, 4, tile_top + 2), (5, 8, tile_top + 2)]
    for (c0, c1, r0), (formula, label_text, color) in zip(positions, tiles):
        _stat_tile(ws, r0, r0 + 1, c0, c1, formula, label_text, color)

    # 習熟度分布（未経験／初級／中級／上級の件数。データバーで視覚化）
    dist_row = tile_top + 4
    dist_label_row = dist_row + 1
    dist_specs = [
        ("未経験", f'=SUMPRODUCT(({skill_row_array}=0)*1)'),
        ("初級(Lv1-2)", f'=SUMPRODUCT(({skill_row_array}>=1)*({skill_row_array}<=2))'),
        ("中級(Lv3)", f'=SUMPRODUCT(({skill_row_array}=3)*1)'),
        ("上級(Lv4-5)", f'=SUMPRODUCT(({skill_row_array}>=4)*1)'),
    ]
    dist_title = ws.cell(row=dist_row - 1, column=1, value="習熟度分布（全作業項目・実効レベル）")
    dist_title.font = Font(size=7.5, bold=True, color=GRAY_TEXT)
    dist_title.fill = _fill(CARD_BG)
    ws.merge_cells(start_row=dist_row - 1, start_column=1, end_row=dist_row - 1, end_column=TOTAL_COLS)
    for i, (label, formula) in enumerate(dist_specs):
        c0 = 1 + i * 3
        c1 = c0 + 2
        cell = ws.cell(row=dist_row, column=c0, value=formula)
        cell.font = Font(bold=True, size=9)
        cell.alignment = Alignment(horizontal="center")
        cell.fill = _fill(WHITE)
        if c1 > c0:
            ws.merge_cells(start_row=dist_row, start_column=c0, end_row=dist_row, end_column=c1)
        ws.conditional_formatting.add(
            cell.coordinate,
            DataBarRule(start_type="num", start_value=0, end_type="num", end_value=len(skill_items) or 1,
                        color="638EC6", showValue=True),
        )
        label_cell = ws.cell(row=dist_label_row, column=c0, value=label)
        label_cell.font = Font(size=6.5, color=GRAY_TEXT)
        label_cell.fill = _fill(CARD_BG)
        if c1 > c0:
            ws.merge_cells(start_row=dist_label_row, start_column=c0, end_row=dist_label_row, end_column=c1)

    row = dist_label_row + 1
    _card_body_border(ws, tile_top, row - 1)
    row += 1

    # 本人の希望
    row = _card_header(ws, row, "本人の希望（今後やりたい仕事）", BLUE)
    wish_formula = _xlookup(config.SHEET_EMPLOYEE, EMP, "社員番号", "本人の希望（今後やりたい仕事）")
    wish_cell = ws.cell(row=row, column=1, value=wish_formula)
    wish_cell.font = Font(size=8.5)
    wish_cell.fill = _fill(CARD_BG)
    wish_cell.alignment = Alignment(wrap_text=True, vertical="center")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=TOTAL_COLS)
    ws.row_dimensions[row].height = 16
    _card_body_border(ws, row, row)
    row += 1

    # 所属歴
    row = _card_header(ws, row, "所属歴（最大6件を表示）", NAVY_SOFT)
    body_top = row
    row = _add_filter_section(ws, row, config.SHEET_ASSIGNMENT, ASSIGN,
                               ["開始年月", "終了年月", "事業所・部署", "班", "担当業務", "役割"],
                               date_cols={"開始年月", "終了年月"}, budget_rows=6)
    _card_body_border(ws, body_top, row - 1)
    row += 1

    # 免許・資格
    row = _card_header(ws, row, "免許・資格（最大6件を表示）", GREEN)
    body_top = row
    row = _add_filter_section(ws, row, config.SHEET_LICENSE, LIC,
                               ["資格名", "取得年月", "有効期限", "実務での使用状況"],
                               date_cols={"取得年月", "有効期限"}, budget_rows=6)
    _card_body_border(ws, body_top, row - 1)
    row += 1

    # 教育歴
    row = _card_header(ws, row, "教育歴（社外・社内）", PURPLE)
    body_top = row
    sub_cell = ws.cell(row=row, column=1, value="社外教育（最大4件）")
    sub_cell.font = Font(bold=True, size=7.5, italic=True, color=GRAY_TEXT)
    sub_cell.fill = _fill(CARD_BG)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=TOTAL_COLS)
    row += 1
    row = _add_filter_section(ws, row, config.SHEET_EXTERNAL_EDU, EXT_EDU,
                               ["講習・研修名", "実施機関", "受講年月"],
                               date_cols={"受講年月"}, budget_rows=4)
    sub_cell = ws.cell(row=row, column=1, value="社内教育（最大9件）")
    sub_cell.font = Font(bold=True, size=7.5, italic=True, color=GRAY_TEXT)
    sub_cell.fill = _fill(CARD_BG)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=TOTAL_COLS)
    row += 1
    row = _add_filter_section(ws, row, config.SHEET_INTERNAL_EDU, INT_EDU,
                               ["教育名", "実施年月", "修了区分"],
                               date_cols={"実施年月"}, budget_rows=9)
    _card_body_border(ws, body_top, row - 1)
    row += 1

    # スキル習熟度×目標GAP（疑似レーダー：データバー＋GAP色分け）
    row = _card_header(ws, row, "スキル習熟度 × 目標GAP（実効レベル基準）", TEAL)
    body_top = row
    _, grade_row_start, grade_row_end = workbook_init.settings_target_table_rows(skill_items)
    settings_ref = _sheet_ref(config.SHEET_SETTINGS)
    grade_formula = _xlookup_expr(config.SHEET_EMPLOYEE, EMP, "社員番号", "等級")

    for i, item in enumerate(skill_items):
        block = i % 2
        r = body_top + i // 2
        c0 = 1 if block == 0 else 7
        eff_col_letter = get_column_letter(cols[item]["effective_col"])
        flag_col_letter = get_column_letter(cols[item]["flag_col"])

        flag_check = (f'IFERROR(XLOOKUP({SEL},{eff_ref}!$A$2:$A${RANGE_MAX},'
                       f'{eff_ref}!${flag_col_letter}$2:${flag_col_letter}${RANGE_MAX}),"")')
        name_formula = f'=IF({flag_check}="","{item}","⚠ {item}(見直し中)")'
        name_cell = ws.cell(row=r, column=c0, value=name_formula)
        name_cell.font = Font(size=6.8, color=GRAY_TEXT)
        name_cell.alignment = Alignment(vertical="center", indent=1)
        ws.merge_cells(start_row=r, start_column=c0, end_row=r, end_column=c0 + 3)

        level_formula = (f'=IFERROR(XLOOKUP({SEL},{eff_ref}!$A$2:$A${RANGE_MAX},'
                          f'{eff_ref}!${eff_col_letter}$2:${eff_col_letter}${RANGE_MAX}),0)')
        level_cell = ws.cell(row=r, column=c0 + 4, value=level_formula)
        level_cell.font = Font(size=8, bold=True)
        level_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.conditional_formatting.add(
            level_cell.coordinate,
            DataBarRule(start_type="num", start_value=0, end_type="num", end_value=5,
                        color="638EC6", showValue=True),
        )

        item_target_col = get_column_letter(2 + skill_items.index(item))
        target_formula = (f'XLOOKUP({grade_formula},{settings_ref}!$A${grade_row_start}:$A${grade_row_end},'
                           f'{settings_ref}!${item_target_col}${grade_row_start}:${item_target_col}${grade_row_end})')
        gap_cell = ws.cell(row=r, column=c0 + 5,
                            value=f'=IFERROR({level_cell.coordinate}-({target_formula}),"")')
        gap_cell.font = Font(size=8, bold=True)
        gap_cell.alignment = Alignment(horizontal="center", vertical="center")
        gap_cell.number_format = '+0;-0;±0'
        ws.conditional_formatting.add(
            gap_cell.coordinate,
            CellIsRule(operator="lessThan", formula=["0"], font=Font(color="FFC00000", bold=True)))
        ws.conditional_formatting.add(
            gap_cell.coordinate,
            CellIsRule(operator="greaterThanOrEqual", formula=["0"], font=Font(color="FF548235", bold=True)))

    n_rows = -(-len(skill_items) // 2) if skill_items else 0
    row = body_top + max(n_rows, 1)
    caption = ws.cell(row=row, column=1, value="※GAP＝実効レベルー等級別目標レベル（設定シートで編集可）。⚠は品質不具合による見直し中を示します。")
    caption.font = Font(size=6.5, italic=True, color=GRAY_TEXT)
    caption.fill = _fill(CARD_BG)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=TOTAL_COLS)
    row += 1
    _card_body_border(ws, body_top, row - 1)
    row += 1

    # 品質不具合歴
    row = _card_header(ws, row, "品質不具合歴（育成・再発防止のための記録／個人を責めるものではありません）", AMBER)
    body_top = row
    sum_cols = effective_sheet.summary_columns(skill_items)
    count_1y_letter = get_column_letter(sum_cols["count_1y"])
    count_3y_letter = get_column_letter(sum_cols["count_3y"])
    trend_letter = get_column_letter(sum_cols["trend"])
    count_1y_formula = (f'IFERROR(XLOOKUP({SEL},{eff_ref}!$A$2:$A${RANGE_MAX},'
                         f'{eff_ref}!${count_1y_letter}$2:${count_1y_letter}${RANGE_MAX}),0)')
    count_3y_formula = (f'IFERROR(XLOOKUP({SEL},{eff_ref}!$A$2:$A${RANGE_MAX},'
                         f'{eff_ref}!${count_3y_letter}$2:${count_3y_letter}${RANGE_MAX}),0)')
    trend_formula = (f'IFERROR(XLOOKUP({SEL},{eff_ref}!$A$2:$A${RANGE_MAX},'
                      f'{eff_ref}!${trend_letter}$2:${trend_letter}${RANGE_MAX}),"")')
    summary_cell = ws.cell(
        row=row, column=1,
        value=f'="直近1年: "&{count_1y_formula}&"件　直近3年: "&{count_3y_formula}&"件　傾向: "&{trend_formula}')
    summary_cell.font = Font(size=8, bold=True)
    summary_cell.fill = _fill(CARD_BG)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=TOTAL_COLS)
    row += 1
    row = _add_filter_section(ws, row, config.SHEET_QUALITY, QUAL,
                               ["発生年月", "区分", "原因", "再発防止として実施した教育・対策"],
                               header_row=2, date_cols={"発生年月"}, budget_rows=6)
    _card_body_border(ws, body_top, row - 1)
    row += 1

    # 備考
    row = _card_header(ws, row, "備考", NAVY_SOFT)
    body_top = row
    remark_formula = _xlookup(config.SHEET_EMPLOYEE, EMP, "社員番号", "備考")
    remark_cell = ws.cell(row=row, column=1, value=remark_formula)
    remark_cell.font = Font(size=8)
    remark_cell.fill = _fill(CARD_BG)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=TOTAL_COLS)
    _card_body_border(ws, body_top, row)
    row += 1

    # 今後の育成方針（印刷して手書き記入する欄）
    row = _card_header(ws, row, "今後の育成方針（面談時に記入）", FOREST)
    body_top = row
    for i in range(2):
        r = row + i
        for c in range(1, TOTAL_COLS + 1):
            ws.cell(row=r, column=c).border = Border(bottom=Side(style="dotted", color=CARD_BORDER_COLOR))
        ws.row_dimensions[r].height = 15
    row += 2
    _card_body_border(ws, body_top, row - 1)
    row += 1

    # 一括印刷についての案内
    hint_cell = ws.cell(row=row, column=1,
                         value="※全員分をまとめて印刷する方法はREADMEの「全員分をまとめて印刷する」を参照してください（初回のみマクロ設定が必要）。")
    hint_cell.font = Font(size=7, italic=True, color=GRAY_TEXT)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=TOTAL_COLS)
    row += 1

    # 印刷設定：A4縦1枚に収まるように
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_margins = PageMargins(left=0.35, right=0.35, top=0.3, bottom=0.3, header=0.2, footer=0.2)
    ws.print_area = f"A1:{get_column_letter(TOTAL_COLS)}{row}"
    return ws
