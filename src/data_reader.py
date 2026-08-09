# 既存のExcelブック（現場管理者が編集した入力シート）を読み込み、
# Python側で扱いやすい形（辞書のリスト）に変換するファイル
# ※年齢・勤続年数・在籍期間・更新要否は数式のセルなので読み込まず、
#   必要な箇所ではcalculations.pyで生年月などの元データから都度計算し直す

import openpyxl

from src import config


def _row_to_dict(ws, row_idx, headers, header_row=1):
    values = {}
    for col, header in enumerate(headers, start=1):
        values[header] = ws.cell(row=row_idx, column=col).value
    return values


def _read_rows(ws, headers, key_header, header_row=1):
    """key_header（通常は社員番号）が空でない行をすべて辞書として返す"""
    key_col = headers.index(key_header) + 1
    rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        key_value = ws.cell(row=r, column=key_col).value
        if key_value in (None, ""):
            continue
        rows.append(_row_to_dict(ws, r, headers, header_row))
    return rows


def read_settings(wb):
    """設定シートから作業項目一覧を読み込む"""
    if config.SHEET_SETTINGS not in wb.sheetnames:
        return list(config.DEFAULT_SKILL_ITEMS)
    ws = wb[config.SHEET_SETTINGS]
    items = []
    for r in range(2, ws.max_row + 1):
        value = ws.cell(row=r, column=1).value
        if value not in (None, ""):
            items.append(str(value).strip())
    return items or list(config.DEFAULT_SKILL_ITEMS)


def read_employees(wb):
    ws = wb[config.SHEET_EMPLOYEE]
    return _read_rows(ws, config.EMPLOYEE_HEADERS, "社員番号")


def read_assignments(wb):
    ws = wb[config.SHEET_ASSIGNMENT]
    return _read_rows(ws, config.ASSIGNMENT_HEADERS, "社員番号")


def read_licenses(wb):
    ws = wb[config.SHEET_LICENSE]
    return _read_rows(ws, config.LICENSE_HEADERS, "社員番号")


def read_external_edu(wb):
    ws = wb[config.SHEET_EXTERNAL_EDU]
    return _read_rows(ws, config.EXTERNAL_EDU_HEADERS, "社員番号")


def read_internal_edu(wb):
    ws = wb[config.SHEET_INTERNAL_EDU]
    return _read_rows(ws, config.INTERNAL_EDU_HEADERS, "社員番号")


def read_quality(wb):
    ws = wb[config.SHEET_QUALITY]
    # 品質不具合歴シートは1行目に注意書きがあるためヘッダーは2行目
    return _read_rows(ws, config.QUALITY_HEADERS, "社員番号", header_row=2)


def read_skills(wb, skill_items):
    """スキル習熟度シートを読み込む。列見出しが設定シートの項目と一致するものだけ拾う"""
    ws = wb[config.SHEET_SKILL]
    sheet_headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    id_col = sheet_headers.index("社員番号") + 1

    item_cols = {}
    for item in skill_items:
        if item in sheet_headers:
            item_cols[item] = sheet_headers.index(item) + 1

    skills = {}
    for r in range(2, ws.max_row + 1):
        emp_id = ws.cell(row=r, column=id_col).value
        if emp_id in (None, ""):
            continue
        levels = {}
        for item in skill_items:
            col = item_cols.get(item)
            value = ws.cell(row=r, column=col).value if col else None
            levels[item] = int(value) if isinstance(value, (int, float)) else 0
        skills[emp_id] = levels
    return skills


def load_all(path):
    """ブック全体を読み込み、各シートのデータをまとめて返す"""
    wb = openpyxl.load_workbook(path, data_only=False)
    skill_items = read_settings(wb)
    data = {
        "skill_items": skill_items,
        "employees": read_employees(wb),
        "assignments": read_assignments(wb),
        "licenses": read_licenses(wb),
        "external_edu": read_external_edu(wb),
        "internal_edu": read_internal_edu(wb),
        "quality": read_quality(wb),
        "skills": read_skills(wb, skill_items),
    }
    wb.close()
    return data
