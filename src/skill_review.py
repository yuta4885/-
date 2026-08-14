# 品質不具合歴に記録された「スキルレベルの一時的な見直し」を、
# 実際のスキル評価（一覧表・抽出結果・個人詳細など）に反映するための共通ロジック
#
# 考え方：
#   ・スキル習熟度シートの値を「本来の評価」として保持したまま上書きはしない
#   ・品質不具合歴に「見直し対象スキル項目」「見直し後レベル」が入力され、
#     かつ「再認定日」が未到来（または未設定）の場合のみ、その項目の
#     実効レベルを「見直し後レベル」に一時的に引き下げて扱う
#   ・再認定日を過ぎれば、自動的に元のスキル習熟度シートの値に戻る
#   ・一覧表・抽出結果・個人詳細など、レベルを使うすべての集計はこの
#     「実効レベル」を使うことで、見直しの反映漏れを防ぐ

from datetime import date

from src import calculations


def is_review_active(record, today=None):
    """品質不具合歴の1行が『現在も見直し中（レベル引き下げ中）』かどうかを判定する"""
    today = today or date.today()
    item = record.get("見直し対象スキル項目")
    new_level = record.get("見直し後レベル")
    if not item or new_level in (None, ""):
        return False
    recert_date = calculations.to_date(record.get("再認定日"))
    if recert_date is None:
        return True  # 再認定日が未設定＝まだ見直し中として扱う
    return recert_date > today


def compute_effective_skills(skills, quality_records, skill_items):
    """
    品質不具合による見直しを反映した「実効スキルレベル」を計算する。

    戻り値:
      effective_skills: {社員番号: {作業項目: 実効レベル}}（skillsと同じ形。見直し中の項目だけ値が変わる）
      review_flags:      {(社員番号, 作業項目): 見直し後レベル}（現在見直し中のものだけ）
    """
    today = date.today()
    effective = {emp_id: dict(levels) for emp_id, levels in skills.items()}
    review_flags = {}

    for record in quality_records:
        if not is_review_active(record, today):
            continue
        emp_id = record.get("社員番号")
        item = record.get("見直し対象スキル項目")
        if emp_id not in effective or item not in skill_items:
            continue
        try:
            new_level = int(record.get("見直し後レベル"))
        except (TypeError, ValueError):
            continue

        key = (emp_id, item)
        if key in review_flags:
            new_level = min(new_level, review_flags[key])  # 複数記録がある場合は安全側（低い方）を採用
        review_flags[key] = new_level

    for (emp_id, item), level in review_flags.items():
        effective[emp_id][item] = level

    return effective, review_flags
