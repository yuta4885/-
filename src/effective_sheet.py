# 個人詳細シートのXLOOKUP参照専用の隠しシートを作成するファイル
# 品質不具合歴による「スキルレベルの一時的な見直し」を反映した実効レベルと、
# 現在見直し中かどうかのフラグを、社員×作業項目の一覧として書き出す。
# 個人詳細シートは、このシートをXLOOKUPで参照するだけのシンプルな数式にできる
# （見直しの判定ロジック自体はPython側に閉じ込め、Excel式を複雑にしすぎないため）。
#
# 現場のExcelで開くたびに自動再計算されるものではなく、スクリプトを実行した
# 時点の判定結果である点に注意（README参照）。

from datetime import date

from openpyxl.styles import Font, PatternFill

from src import calculations, config, skill_review

HEADER_FILL = PatternFill("solid", fgColor="FF404040")
HEADER_FONT = Font(color="FFFFFFFF", bold=True, size=8)

FLAG_SUFFIX = "__見直し中レベル"

# 品質不具合の集計用の列名（社員番号・作業項目の列の後ろに追加する）
COL_COUNT_1Y = "__品質不具合_直近1年件数"
COL_COUNT_3Y = "__品質不具合_直近3年件数"
COL_TREND = "__品質不具合_傾向"


def columns_for(skill_items):
    """各作業項目の『実効レベル列』『見直し中フラグ列』の列番号（1始まり）を返す"""
    n_items = len(skill_items)
    cols = {}
    for i, item in enumerate(skill_items):
        cols[item] = {
            "effective_col": 2 + i,
            "flag_col": 2 + n_items + i,
        }
    return cols


def summary_columns(skill_items):
    """品質不具合の集計3列（直近1年件数・直近3年件数・傾向）の列番号（1始まり）を返す"""
    base = 2 + 2 * len(skill_items)
    return {"count_1y": base, "count_3y": base + 1, "trend": base + 2}


def _quality_summary(quality_records, emp_id):
    """品質不具合歴から、直近1年・3年の件数と繰り返し傾向を計算する"""
    today = date.today()
    my_records = [r for r in quality_records if r.get("社員番号") == emp_id]

    def within_days(record, days):
        occurred = calculations.to_date(record.get("発生年月"))
        return occurred is not None and (today - occurred).days <= days

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

    return count_1y, count_3y, trend


def build_effective_sheet(wb, employees, skills, quality, skill_items):
    """実効スキルレベル・見直し中フラグを計算し、隠しシートに書き出す"""
    if config.SHEET_EFFECTIVE in wb.sheetnames:
        del wb[config.SHEET_EFFECTIVE]
    ws = wb.create_sheet(config.SHEET_EFFECTIVE)

    effective, review_flags = skill_review.compute_effective_skills(skills, quality, skill_items)
    cols = columns_for(skill_items)
    sum_cols = summary_columns(skill_items)

    headers = (["社員番号"] + list(skill_items) + [f"{item}{FLAG_SUFFIX}" for item in skill_items]
               + [COL_COUNT_1Y, COL_COUNT_3Y, COL_TREND])
    for col, title in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=title)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT

    for r, emp in enumerate(employees, start=2):
        emp_id = emp["社員番号"]
        ws.cell(row=r, column=1, value=emp_id)
        emp_effective = effective.get(emp_id, {})
        for item in skill_items:
            ws.cell(row=r, column=cols[item]["effective_col"], value=emp_effective.get(item, 0))
            flag_value = review_flags.get((emp_id, item))
            ws.cell(row=r, column=cols[item]["flag_col"], value=flag_value if flag_value is not None else "")

        count_1y, count_3y, trend = _quality_summary(quality, emp_id)
        ws.cell(row=r, column=sum_cols["count_1y"], value=count_1y)
        ws.cell(row=r, column=sum_cols["count_3y"], value=count_3y)
        ws.cell(row=r, column=sum_cols["trend"], value=trend)

    ws.sheet_state = "hidden"
    return ws, effective, review_flags
