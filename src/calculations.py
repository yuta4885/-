# 年齢・勤続年数・在籍期間などの自動計算をまとめたファイル
# ・社員マスタ／所属歴シート上の表示用にはExcel数式を書き込み、
#   Excelを開くたびに自動で最新の日付を基準に再計算されるようにする
# ・スクリプト内部での集計（個人シート・抽出結果など）は
#   数式のキャッシュ値に頼らず、このファイルのPython関数で都度計算する

from datetime import date, datetime


def to_date(value):
    """セルの値（datetime／date／文字列／空欄）をdateに変換する。変換できなければNone"""
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y年%m月"):
            try:
                return datetime.strptime(value, fmt).date()
            except ValueError:
                continue
    return None


def years_between(start, end=None):
    """開始日から終了日（省略時は今日）までの満年数を計算する"""
    start = to_date(start)
    if start is None:
        return None
    end = to_date(end) if end is not None else date.today()
    years = end.year - start.year - ((end.month, end.day) < (start.month, start.day))
    return max(years, 0)


def days_until(target):
    """今日から対象日までの残り日数（過去日ならマイナス）"""
    target = to_date(target)
    if target is None:
        return None
    return (target - date.today()).days


def age_formula(birth_cell):
    """年齢セルに入れるExcel数式（生年月セルを参照）"""
    return f'=IF({birth_cell}="","",DATEDIF({birth_cell},TODAY(),"Y"))'


def tenure_formula(join_cell):
    """勤続年数セルに入れるExcel数式（入社年月セルを参照）"""
    return f'=IF({join_cell}="","",DATEDIF({join_cell},TODAY(),"Y"))'


def assignment_period_formula(start_cell, end_cell):
    """所属歴の在籍期間セルに入れるExcel数式。終了年月が空欄なら現在まで在籍中として計算"""
    return (
        f'=IF({start_cell}="","",'
        f'IF({end_cell}="",DATEDIF({start_cell},TODAY(),"Y"),'
        f'DATEDIF({start_cell},{end_cell},"Y")))'
    )


def license_renewal_formula(expiry_cell):
    """免許・資格の更新要否セルに入れるExcel数式（期限切れ／要更新＝3か月以内／そろそろ＝6か月以内／不要の4段階）"""
    return (
        f'=IF({expiry_cell}="","",'
        f'IF({expiry_cell}<TODAY(),"期限切れ",'
        f'IF({expiry_cell}<=TODAY()+90,"要更新",'
        f'IF({expiry_cell}<=TODAY()+180,"そろそろ","不要"))))'
    )
