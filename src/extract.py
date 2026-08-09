# 「抽出結果」シート（自動生成・編集不可）を作成するファイル
# あらかじめ用意した5つの条件で該当者を抽出し、それぞれ表にして出力する
#   1. 海外出荷 Lv4以上
#   2. 新人教育ができる人（Lv5）
#   3. フォークリフト保有かつ使用中
#   4. 安全教育 未受講者
#   5. 資格の有効期限が6か月以内（期限切れ含む）

from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from src import calculations, config

HEADER_FILL = PatternFill("solid", fgColor="FF305496")
HEADER_FONT = Font(color="FFFFFFFF", bold=True)
TITLE_FONT = Font(bold=True, size=12, color="FF1F4E79")
THIN_BORDER = Border(*(Side(style="thin", color="FFBFBFBF"),) * 4)
WARN_FILL = PatternFill("solid", fgColor="FFFFC7CE")


def _employee_name_map(employees):
    return {e["社員番号"]: e.get("氏名", "") for e in employees}


def preset_overseas_shipping(data, level=4):
    item = "海外出荷"
    if item not in data["skill_items"]:
        return [], f"「{item}」が作業項目に見つかりません（設定シートを確認してください）"
    names = _employee_name_map(data["employees"])
    rows = []
    for emp_id, levels in data["skills"].items():
        if levels.get(item, 0) >= level:
            rows.append({"社員番号": emp_id, "氏名": names.get(emp_id, ""), "習熟度": levels.get(item, 0)})
    return rows, None


def preset_trainer_level(data, item="新人教育", level=5):
    if item not in data["skill_items"]:
        return [], f"「{item}」が作業項目に見つかりません（設定シートを確認してください）"
    names = _employee_name_map(data["employees"])
    rows = []
    for emp_id, levels in data["skills"].items():
        if levels.get(item, 0) >= level:
            rows.append({"社員番号": emp_id, "氏名": names.get(emp_id, ""), "習熟度": levels.get(item, 0)})
    return rows, None


def preset_forklift_in_use(data):
    names = _employee_name_map(data["employees"])
    rows = []
    for lic in data["licenses"]:
        if "フォークリフト" in str(lic.get("資格名", "")) and lic.get("実務での使用状況") == "使用中":
            rows.append({
                "社員番号": lic["社員番号"],
                "氏名": names.get(lic["社員番号"], ""),
                "資格名": lic.get("資格名"),
                "実務での使用状況": lic.get("実務での使用状況"),
            })
    return rows, None


def preset_safety_training_incomplete(data, edu_name="安全教育"):
    names = _employee_name_map(data["employees"])
    latest_status = {}
    for rec in data["internal_edu"]:
        if rec.get("教育名") != edu_name:
            continue
        latest_status[rec["社員番号"]] = rec.get("修了区分")

    rows = []
    for emp_id, name in names.items():
        status = latest_status.get(emp_id, "未受講（記録なし）")
        if status in ("未受講", "未受講（記録なし）"):
            rows.append({"社員番号": emp_id, "氏名": name, "修了区分": status})
    return rows, None


def preset_license_expiring(data, days=None):
    days = days if days is not None else config.LICENSE_WARNING_DAYS
    names = _employee_name_map(data["employees"])
    rows = []
    for lic in data["licenses"]:
        remaining = calculations.days_until(lic.get("有効期限"))
        if remaining is None:
            continue
        if remaining <= days:
            status = "期限切れ" if remaining < 0 else "期限間近"
            rows.append({
                "社員番号": lic["社員番号"],
                "氏名": names.get(lic["社員番号"], ""),
                "資格名": lic.get("資格名"),
                "有効期限": lic.get("有効期限"),
                "残り日数": remaining,
                "状態": status,
            })
    rows.sort(key=lambda r: r["残り日数"])
    return rows, None


PRESETS = [
    ("① 海外出荷 Lv4以上", preset_overseas_shipping),
    ("② 新人教育ができる人（Lv5）", preset_trainer_level),
    ("③ フォークリフト保有かつ使用中", preset_forklift_in_use),
    ("④ 安全教育 未受講者", preset_safety_training_incomplete),
    ("⑤ 資格の有効期限が6か月以内（期限切れ含む）", preset_license_expiring),
]


def _write_table(ws, start_row, title, rows, empty_message=None):
    ws.cell(row=start_row, column=1, value=title).font = TITLE_FONT
    row = start_row + 1

    if empty_message:
        ws.cell(row=row, column=1, value=empty_message).font = Font(italic=True, color="FFC00000")
        return row + 2

    if not rows:
        ws.cell(row=row, column=1, value="該当者なし").font = Font(italic=True, color="FF808080")
        return row + 2

    headers = list(rows[0].keys())
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=col, value=header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        ws.column_dimensions[get_column_letter(col)].width = max(12, len(header) + 4)
    row += 1

    for record in rows:
        for col, header in enumerate(headers, start=1):
            value = record[header]
            cell = ws.cell(row=row, column=col, value=value)
            cell.border = THIN_BORDER
            if header == "有効期限" and value not in (None, ""):
                cell.number_format = "yyyy/mm/dd"
            if header == "状態" and value == "期限切れ":
                cell.fill = WARN_FILL
        row += 1

    return row + 2


def build_extract_sheet(wb, data):
    if config.SHEET_EXTRACT in wb.sheetnames:
        del wb[config.SHEET_EXTRACT]
    ws = wb.create_sheet(config.SHEET_EXTRACT)
    ws.cell(row=1, column=1,
            value="抽出結果（自動生成・このシートは編集しないでください。次回実行時に上書きされます）").font = Font(bold=True, size=13)

    row = 3
    for title, func in PRESETS:
        rows, empty_message = func(data)
        row = _write_table(ws, row, title, rows, empty_message)

    ws.protection.sheet = True
    ws.protection.formatCells = False
    return ws
