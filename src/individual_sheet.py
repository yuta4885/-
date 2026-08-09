# 「個人シート」（自動生成・編集不可、社員1人につき1枚）を作成するファイル
# ダッシュボード風のカードレイアウト（ヘッダーバー／バッジ／統計タイル／
# 習熟度分布バー／セクションカード）でまとめ、
# A4縦1枚に収まる体裁で、面談などにそのまま印刷して使えるようにする
#
# ※グラフオブジェクト（openpyxlのChart機能）は使わず、セルの色分けだけで
#   統計・分布を表現している。理由：Excelバージョンによってはグラフを含む
#   ブックが正しく開けない・崩れるケースがあるため、現場のExcelで確実に
#   同じ見た目になるセルベースの表現に統一した。

from datetime import date

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

from src import calculations, config

# ------------------------------------------------------------
# 配色（ダッシュボード風の配色パレット）
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
ZEBRA_FILL = "FFEFF3FA"
WHITE = "FFFFFFFF"
GRAY_TEXT = "FF595959"

BADGE_COLORS = [BLUE, GREEN, ORANGE, PURPLE]

THIN_BORDER = Border(*(Side(style="thin", color=CARD_BORDER_COLOR),) * 4)
CARD_OUTLINE = Border(
    left=Side(style="thin", color=CARD_BORDER_COLOR), right=Side(style="thin", color=CARD_BORDER_COLOR),
    top=Side(style="thin", color=CARD_BORDER_COLOR), bottom=Side(style="thin", color=CARD_BORDER_COLOR),
)

TOTAL_COLS = 12


def _fill(color):
    return PatternFill("solid", fgColor=color)


def _card_header(ws, row, title, color, col_start=1, col_end=TOTAL_COLS):
    for c in range(col_start, col_end + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = _fill(color)
    cell = ws.cell(row=row, column=col_start, value=f"  {title}")
    cell.font = Font(color=WHITE, bold=True, size=10)
    cell.alignment = Alignment(vertical="center")
    ws.merge_cells(start_row=row, start_column=col_start, end_row=row, end_column=col_end)
    ws.row_dimensions[row].height = 15.5
    return row + 1


def _card_body_border(ws, row_start, row_end, col_start=1, col_end=TOTAL_COLS):
    """カード本体に外枠＋薄い背景色を付けて、白い帯の中に浮いたカードのように見せる"""
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


def _pill(ws, row, col_start, col_end, text, color):
    for c in range(col_start, col_end + 1):
        ws.cell(row=row, column=c).fill = _fill(color)
    cell = ws.cell(row=row, column=col_start, value=text)
    cell.font = Font(color=WHITE, bold=True, size=8.5)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=row, start_column=col_start, end_row=row, end_column=col_end)


def _stat_tile(ws, num_row, label_row, col_start, col_end, number_text, label_text, accent_color):
    for c in range(col_start, col_end + 1):
        top_cell = ws.cell(row=num_row, column=c)
        top_cell.border = Border(top=Side(style="medium", color=accent_color))
        top_cell.fill = _fill(CARD_BG)
        ws.cell(row=label_row, column=c).fill = _fill(CARD_BG)
    num_cell = ws.cell(row=num_row, column=col_start, value=number_text)
    num_cell.font = Font(bold=True, size=17, color=accent_color)
    num_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=num_row, start_column=col_start, end_row=num_row, end_column=col_end)
    ws.row_dimensions[num_row].height = 20

    label_cell = ws.cell(row=label_row, column=col_start, value=label_text)
    label_cell.font = Font(size=8, color=GRAY_TEXT)
    label_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=label_row, start_column=col_start, end_row=label_row, end_column=col_end)
    ws.row_dimensions[label_row].height = 12


def _mini_table(ws, row, headers, records, empty_text="記録なし", date_cols=(), col_start=1, col_end=TOTAL_COLS):
    if not records:
        cell = ws.cell(row=row, column=col_start, value=empty_text)
        cell.font = Font(size=8, italic=True, color=GRAY_TEXT)
        cell.fill = _fill(CARD_BG)
        ws.merge_cells(start_row=row, start_column=col_start, end_row=row, end_column=col_end)
        return row + 1

    n_cols = col_end - col_start + 1
    widths = [n_cols // len(headers)] * len(headers)
    for i in range(n_cols - sum(widths)):
        widths[i] += 1

    col_positions = []
    c = col_start
    for w in widths:
        col_positions.append((c, c + w - 1))
        c += w

    for (c0, c1), header in zip(col_positions, headers):
        cell = ws.cell(row=row, column=c0, value=header)
        cell.fill = _fill("FFD9E1F2")
        cell.font = Font(bold=True, size=7.5)
        cell.border = THIN_BORDER
        if c1 > c0:
            ws.merge_cells(start_row=row, start_column=c0, end_row=row, end_column=c1)
    row += 1

    for i, record in enumerate(records):
        stripe = ZEBRA_FILL if i % 2 == 1 else WHITE
        for (c0, c1), header in zip(col_positions, headers):
            value = record.get(header, "")
            cell = ws.cell(row=row, column=c0, value=value)
            cell.font = Font(size=7.5)
            cell.fill = _fill(stripe)
            cell.border = THIN_BORDER
            if c1 > c0:
                ws.merge_cells(start_row=row, start_column=c0, end_row=row, end_column=c1)
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


def _skill_buckets(my_skills, skill_items):
    """習熟度分布バー用に、レベルを4段階のバケツに集計する"""
    buckets = {"未経験": 0, "初級(Lv1-2)": 0, "中級(Lv3)": 0, "上級(Lv4-5)": 0}
    for item in skill_items:
        level = my_skills.get(item, 0)
        if level == 0:
            buckets["未経験"] += 1
        elif level <= 2:
            buckets["初級(Lv1-2)"] += 1
        elif level == 3:
            buckets["中級(Lv3)"] += 1
        else:
            buckets["上級(Lv4-5)"] += 1
    return buckets


def _add_skill_distribution_bar(ws, row, col_start, col_end, buckets):
    """習熟度の分布を、セルを横に並べた色分けバーで表現する（グラフ機能を使わないため、
    どのExcelバージョンでも崩れずに同じ見た目になる）"""
    colors = ["FFD9D9D9", "FF9DC3E6", "FF5B9BD5", "FF1F5FA8"]
    total_cols = col_end - col_start + 1
    total_count = sum(buckets.values()) or 1

    # 件数に応じて各区分に列を配分する（最低1列は確保）
    raw_widths = [max(1, round(count / total_count * total_cols)) for count in buckets.values()]
    while sum(raw_widths) > total_cols:
        raw_widths[raw_widths.index(max(raw_widths))] -= 1
    while sum(raw_widths) < total_cols:
        raw_widths[raw_widths.index(max(raw_widths))] += 1

    c = col_start
    for (label, count), width, color in zip(buckets.items(), raw_widths, colors):
        c_end = c + width - 1
        cell = ws.cell(row=row, column=c, value=str(count) if count else "")
        cell.fill = _fill(color)
        cell.font = Font(bold=True, size=8, color=WHITE if color != "FFD9D9D9" else "FF595959")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        if c_end > c:
            ws.merge_cells(start_row=row, start_column=c, end_row=row, end_column=c_end)
        c = c_end + 1
    ws.row_dimensions[row].height = 16

    legend_row = row + 1
    c = col_start
    for (label, count), color in zip(buckets.items(), colors):
        chip = ws.cell(row=legend_row, column=c, value=" ")
        chip.fill = _fill(color)
        c += 1
        text_cell = ws.cell(row=legend_row, column=c, value=f"{label}:{count}")
        text_cell.font = Font(size=6.5, color=GRAY_TEXT)
        c += 2
    ws.row_dimensions[legend_row].height = 11


def build_individual_sheet(wb, emp, assignments, licenses, external_edu, internal_edu,
                            quality, skills, skill_items):
    emp_id = emp["社員番号"]
    sheet_name = f"{config.INDIVIDUAL_SHEET_PREFIX}{emp_id}"
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False

    for c in range(1, TOTAL_COLS + 1):
        ws.column_dimensions[get_column_letter(c)].width = 8.6

    # 上部アクセントバー
    for c in range(1, TOTAL_COLS + 1):
        ws.cell(row=1, column=c).fill = _fill(ACCENT_BAR)
    ws.row_dimensions[1].height = 5

    # ヘッダー（氏名・社員番号）
    header_row = 2
    ws.merge_cells(start_row=header_row, start_column=1, end_row=header_row + 1, end_column=8)
    name_cell = ws.cell(row=header_row, column=1,
                         value=f"{emp.get('氏名', '')}（{emp.get('氏名カナ', '')}）")
    name_cell.font = Font(bold=True, size=18, color=WHITE)
    name_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    ws.merge_cells(start_row=header_row, start_column=9, end_row=header_row, end_column=12)
    label_cell = ws.cell(row=header_row, column=9, value="社員番号")
    label_cell.font = Font(size=8, color="FFBFD0EC")
    label_cell.alignment = Alignment(horizontal="center", vertical="bottom")

    ws.merge_cells(start_row=header_row + 1, start_column=9, end_row=header_row + 1, end_column=12)
    id_cell = ws.cell(row=header_row + 1, column=9, value=emp_id)
    id_cell.font = Font(bold=True, size=14, color=WHITE)
    id_cell.alignment = Alignment(horizontal="center", vertical="top")

    for r in (header_row, header_row + 1):
        for c in range(1, TOTAL_COLS + 1):
            if ws.cell(row=r, column=c).fill.fgColor.rgb in (None, "00000000"):
                ws.cell(row=r, column=c).fill = _fill(NAVY)
    ws.row_dimensions[header_row].height = 24
    ws.row_dimensions[header_row + 1].height = 20

    # バッジ行（所属・班・役職・雇用区分・性別・等級を2段のピルバッジで表示）
    badge_row = header_row + 2
    spans = [(1, 4), (5, 8), (9, 12)]
    badges_line1 = [
        emp.get("現在の所属", "") or "所属未設定",
        emp.get("班", "") or "班未設定",
        emp.get("役職", "") or "役職未設定",
    ]
    badges_line2 = [
        f"雇用区分：{emp.get('雇用区分', '') or '―'}",
        f"性別：{emp.get('性別', '') or '―'}",
        f"等級：{emp.get('等級', '') or '―'}",
    ]
    for (c0, c1), text, color in zip(spans, badges_line1, BADGE_COLORS):
        _pill(ws, badge_row, c0, c1, text, color)
    for (c0, c1), text, color in zip(spans, badges_line2, BADGE_COLORS[::-1]):
        _pill(ws, badge_row + 1, c0, c1, text, color)
    ws.row_dimensions[badge_row].height = 15
    ws.row_dimensions[badge_row + 1].height = 15

    row = badge_row + 2
    note_cell = ws.cell(row=row, column=1,
                         value="※本シートは人員配置・育成計画のための資料です。人事評価・処遇決定には使用しません。")
    note_cell.font = Font(size=7.5, italic=True, color=GRAY_TEXT)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=TOTAL_COLS)
    row += 2

    # サマリー（統計タイル＋習熟度分布バー）
    row = _card_header(ws, row, "サマリー", TEAL)
    tile_top = row
    my_skills = skills.get(emp_id, {})
    n_licenses = sum(1 for lic in licenses if lic.get("社員番号") == emp_id)
    n_advanced = sum(1 for level in my_skills.values() if level >= 4)
    tiles = [
        (f"{calculations.years_between(emp.get('生年月'))}歳" if emp.get("生年月") else "―", "年齢", BLUE),
        (f"{calculations.years_between(emp.get('入社年月'))}年" if emp.get("入社年月") else "―", "勤続年数", GREEN),
        (str(n_licenses), "保有資格数", ORANGE),
        (str(n_advanced), "Lv4以上のスキル数", PURPLE),
    ]
    positions = [(1, 3), (4, 6), (7, 9), (10, 12)]
    for (c0, c1), (number_text, label_text, color) in zip(positions, tiles):
        _stat_tile(ws, tile_top, tile_top + 1, c0, c1, number_text, label_text, color)

    bar_title_row = tile_top + 2
    bar_title_cell = ws.cell(row=bar_title_row, column=1, value="習熟度分布（全作業項目）")
    bar_title_cell.font = Font(size=7.5, bold=True, color=GRAY_TEXT)
    bar_title_cell.fill = _fill(CARD_BG)
    ws.merge_cells(start_row=bar_title_row, start_column=1, end_row=bar_title_row, end_column=TOTAL_COLS)

    buckets = _skill_buckets(my_skills, skill_items)
    _add_skill_distribution_bar(ws, bar_title_row + 1, 1, TOTAL_COLS, buckets)

    row = bar_title_row + 3
    _card_body_border(ws, tile_top, row - 1)
    row += 1

    # 本人の希望
    row = _card_header(ws, row, "本人の希望（今後やりたい仕事）", BLUE)
    wish_cell = ws.cell(row=row, column=1, value=emp.get("本人の希望（今後やりたい仕事）", "") or "特になし")
    wish_cell.font = Font(size=8.5)
    wish_cell.alignment = Alignment(wrap_text=True, vertical="center")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=TOTAL_COLS)
    ws.row_dimensions[row].height = 16
    _card_body_border(ws, row, row)
    row += 1

    # 所属歴
    row = _card_header(ws, row, "所属歴", NAVY_SOFT)
    body_top = row
    my_assignments = [dict(a) for a in assignments if a.get("社員番号") == emp_id]
    for a in my_assignments:
        a["期間"] = f"{_fmt(a.get('開始年月'))}〜{_fmt(a.get('終了年月')) or '現在'}"
    row = _mini_table(ws, row, ["期間", "事業所・部署", "班", "担当業務", "役割"], my_assignments)
    _card_body_border(ws, body_top, row - 1)
    row += 1

    # 免許・資格
    row = _card_header(ws, row, "免許・資格", GREEN)
    body_top = row
    my_licenses = [
        {**lic, "有効期限表示": _fmt(lic.get("有効期限"))}
        for lic in licenses if lic.get("社員番号") == emp_id
    ]
    row = _mini_table(ws, row, ["資格名", "取得年月", "有効期限表示", "実務での使用状況"], my_licenses,
                       date_cols={"取得年月"})
    _card_body_border(ws, body_top, row - 1)
    row += 1

    # 教育歴
    row = _card_header(ws, row, "教育歴", PURPLE)
    body_top = row
    my_external = [e for e in external_edu if e.get("社員番号") == emp_id]
    my_internal = [e for e in internal_edu if e.get("社員番号") == emp_id]
    sub_cell = ws.cell(row=row, column=1, value="社外教育")
    sub_cell.font = Font(bold=True, size=7.5, italic=True, color=GRAY_TEXT)
    sub_cell.fill = _fill(CARD_BG)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=TOTAL_COLS)
    row += 1
    row = _mini_table(ws, row, ["講習・研修名", "実施機関", "受講年月"], my_external, date_cols={"受講年月"})
    sub_cell = ws.cell(row=row, column=1, value="社内教育")
    sub_cell.font = Font(bold=True, size=7.5, italic=True, color=GRAY_TEXT)
    sub_cell.fill = _fill(CARD_BG)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=TOTAL_COLS)
    row += 1
    row = _mini_table(ws, row, ["教育名", "実施年月", "修了区分"], my_internal, date_cols={"実施年月"})
    _card_body_border(ws, body_top, row - 1)
    row += 1

    # スキル習熟度
    row = _card_header(ws, row, "スキル習熟度", TEAL)
    body_top = row
    cols_per_line = 4
    chip_width = TOTAL_COLS // cols_per_line
    for i, item in enumerate(skill_items):
        level = my_skills.get(item, 0)
        c0 = 1 + (i % cols_per_line) * chip_width
        c1 = c0 + chip_width - 1
        r = body_top + i // cols_per_line
        label_cell = ws.cell(row=r, column=c0, value=item)
        label_cell.font = Font(size=6.8, color=GRAY_TEXT)
        label_cell.fill = _fill(WHITE)
        label_cell.alignment = Alignment(vertical="center", indent=1)
        if c1 - 1 > c0:
            ws.merge_cells(start_row=r, start_column=c0, end_row=r, end_column=c1 - 1)
        level_cell = ws.cell(row=r, column=c1, value=level)
        level_cell.font = Font(size=8, bold=True,
                                color=WHITE if level >= 3 else "FF404040")
        level_cell.alignment = Alignment(horizontal="center", vertical="center")
        level_cell.fill = _fill(config.SKILL_FILL_COLORS.get(level, "FFFFFFFF"))
        level_cell.border = THIN_BORDER
    row = body_top + -(-len(skill_items) // cols_per_line)
    _card_body_border(ws, body_top, row - 1)
    row += 1

    # 品質不具合歴
    row = _card_header(ws, row, "品質不具合歴（育成・再発防止のための記録／個人を責めるものではありません）", AMBER)
    body_top = row
    summary = compute_quality_summary(quality, emp_id)
    stat_specs = [
        (str(summary["直近1年件数"]), "直近1年件数", 1, 3),
        (str(summary["直近3年件数"]), "直近3年件数", 4, 6),
    ]
    for number_text, label_text, c0, c1 in stat_specs:
        _stat_tile(ws, row, row + 1, c0, c1, number_text, label_text, AMBER)
    trend_cell = ws.cell(row=row, column=7, value=f"傾向：{summary['傾向']}")
    trend_cell.font = Font(size=7.5, color=GRAY_TEXT)
    trend_cell.fill = _fill(CARD_BG)
    trend_cell.alignment = Alignment(wrap_text=True, vertical="center")
    ws.merge_cells(start_row=row, start_column=7, end_row=row + 1, end_column=TOTAL_COLS)
    row += 2
    my_quality = [q for q in quality if q.get("社員番号") == emp_id]
    row = _mini_table(ws, row, ["発生年月", "区分", "原因", "再発防止として実施した教育・対策"], my_quality,
                       date_cols={"発生年月"})
    _card_body_border(ws, body_top, row - 1)
    row += 1

    # 備考
    row = _card_header(ws, row, "備考", NAVY_SOFT)
    body_top = row
    remark_cell = ws.cell(row=row, column=1, value=emp.get("備考", "") or "特になし")
    remark_cell.font = Font(size=8)
    remark_cell.fill = _fill(CARD_BG)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=TOTAL_COLS)
    _card_body_border(ws, body_top, row)
    row += 1

    # 今後の育成方針
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

    # 印刷設定：A4縦1枚に収まるように
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_margins = PageMargins(left=0.35, right=0.35, top=0.3, bottom=0.3, header=0.2, footer=0.2)
    ws.print_area = f"A1:{get_column_letter(TOTAL_COLS)}{row}"
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
