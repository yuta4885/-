# 人員配置の穴チェック（属人化リスク）を計算するファイル
# 作業項目ごとに「Lv3以上（一人前）の人数」を集計し、1人しかいない
# （その人が抜けると誰もできなくなる）項目を警告として拾い出す

from src import config


def compute_gap_warnings(employees, skills, skill_items):
    """作業項目ごとの一人前人数と属人化リスクを集計してリストで返す"""
    id_to_name = {e["社員番号"]: e["氏名"] for e in employees}

    results = []
    for item in skill_items:
        qualified_ids = [
            emp_id for emp_id, levels in skills.items()
            if levels.get(item, 0) >= config.SKILL_MATRIX_WARNING_LEVEL
        ]
        names = [id_to_name.get(emp_id, emp_id) for emp_id in qualified_ids]
        is_warning = len(qualified_ids) <= config.GAP_WARNING_HEADCOUNT
        results.append({
            "作業項目": item,
            "Lv3以上人数": len(qualified_ids),
            "該当者": "、".join(names) if names else "該当者なし",
            "属人化リスク": "★警告：この作業ができる人が1人以下です" if is_warning else "",
        })
    return results
