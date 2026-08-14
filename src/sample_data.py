# サンプルデータ（架空の氏名10人分）を生成するファイル
# 初回実行時に、動作確認用としてこのデータでExcelブックを作成します

import random
from datetime import date, timedelta

from src import config

# 架空の氏名（10人分）。実在の人物とは関係ありません
SAMPLE_NAMES = [
    ("佐藤 大輔", "サトウ ダイスケ"),
    ("鈴木 美咲", "スズキ ミサキ"),
    ("高橋 健太", "タカハシ ケンタ"),
    ("田中 陽子", "タナカ ヨウコ"),
    ("伊藤 拓也", "イトウ タクヤ"),
    ("渡辺 花子", "ワタナベ ハナコ"),
    ("山本 翔太", "ヤマモト ショウタ"),
    ("中村 沙織", "ナカムラ サオリ"),
    ("小林 誠", "コバヤシ マコト"),
    ("加藤 千尋", "カトウ チヒロ"),
]

POSITIONS = ["一般", "主任", "班長", "係長"]
WISHES = [
    "海外出荷の担当を増やしたい", "フォークリフト作業を任せてほしい",
    "後輩の指導役をやってみたい", "ピッキングの精度を上げたい",
    "特になし", "品質管理の仕事に挑戦したい",
]


def _random_date(start_year, end_year, rng, no_future=False):
    """start_year〜end_yearの範囲で年月をランダムに生成する。
    no_future=Trueの場合、今日より未来の日付にならないよう調整する（入社年月・受講年月など向け）"""
    year = rng.randint(start_year, min(end_year, date.today().year) if no_future else end_year)
    today = date.today()
    if no_future and year == today.year:
        month = rng.randint(1, today.month)
    else:
        month = rng.randint(1, 12)
    return date(year, month, 1)


def generate_grade_targets(skill_items=None):
    """等級（G1〜G4）ごとの目標スキルレベルの初期値を生成する。
    等級が上がるほど期待レベルが上がる、という単純な既定値。
    現場の実情に合わせて『設定』シート上で自由に書き換えることを想定している"""
    skill_items = skill_items or config.DEFAULT_SKILL_ITEMS
    base_by_grade = {"G1": 1, "G2": 2, "G3": 3, "G4": 4}
    return {
        grade: {item: base for item in skill_items}
        for grade, base in base_by_grade.items()
    }


def generate_sample_data(skill_items=None, seed=2024):
    """サンプルデータ10人分を生成して辞書で返す"""
    rng = random.Random(seed)
    skill_items = skill_items or config.DEFAULT_SKILL_ITEMS

    employees = []
    assignments = []
    licenses = []
    external_edu = []
    internal_edu = []
    quality = []
    skills = {}

    today = date.today()

    for i, (name, kana) in enumerate(SAMPLE_NAMES, start=1):
        emp_id = f"E{i:03d}"
        birth = _random_date(today.year - 55, today.year - 20, rng)
        join = _random_date(max(birth.year + 18, today.year - 20), today.year, rng, no_future=True)

        employees.append({
            "社員番号": emp_id,
            "氏名": name,
            "氏名カナ": kana,
            "生年月": birth,
            "性別": rng.choice(config.GENDER_CHOICES),
            "雇用区分": rng.choice(config.EMPLOYMENT_TYPE_CHOICES),
            "入社年月": join,
            "現在の所属": rng.choice(config.DEPARTMENT_CHOICES),
            "班": rng.choice(config.TEAM_CHOICES),
            "役職": rng.choice(POSITIONS),
            "等級": rng.choice(config.GRADE_CHOICES),
            "本人の希望（今後やりたい仕事）": rng.choice(WISHES),
            "備考": "",
        })

        # 所属歴（1〜3件、入社から現在まで）
        n_assign = rng.randint(1, 3)
        cursor = join
        for j in range(n_assign):
            is_last = j == n_assign - 1
            span_months = rng.randint(6, 30)
            end = None if is_last else min(
                date(cursor.year + (cursor.month - 1 + span_months) // 12,
                     (cursor.month - 1 + span_months) % 12 + 1, 1),
                today,
            )
            assignments.append({
                "社員番号": emp_id,
                "開始年月": cursor,
                "終了年月": end,
                "事業所・部署": rng.choice(config.DEPARTMENT_CHOICES),
                "班": rng.choice(config.TEAM_CHOICES),
                "担当業務": rng.choice(config.DUTY_CHOICES),
                "役割": rng.choice(config.ROLE_CHOICES),
                "備考": "",
            })
            if end:
                cursor = end

        # 免許・資格（0〜3件）
        for lic_name in rng.sample(config.LICENSE_EXAMPLES, k=rng.randint(0, 3)):
            acquired = _random_date(join.year, today.year, rng, no_future=True)
            expiry = date(acquired.year + rng.choice([3, 5]), acquired.month, 1)
            licenses.append({
                "社員番号": emp_id,
                "資格名": lic_name,
                "取得年月": acquired,
                "有効期限": expiry,
                "実務での使用状況": rng.choice(config.LICENSE_USAGE_CHOICES),
                "備考": "",
            })

        # 社外教育（0〜2件）
        for _ in range(rng.randint(0, 2)):
            external_edu.append({
                "社員番号": emp_id,
                "講習・研修名": rng.choice(["物流品質セミナー", "安全運転講習", "リーダーシップ研修"]),
                "実施機関": rng.choice(["県トラック協会", "外部研修会社", "商工会議所"]),
                "受講年月": _random_date(join.year, today.year, rng, no_future=True),
                "日数": rng.choice([1, 1, 2, 3]),
                "修了証の有無": rng.choice(config.CERTIFICATE_CHOICES),
                "内容メモ": "",
            })

        # 社内教育（項目ごとに1件、未受講もあり得る）
        for edu_name in config.INTERNAL_EDU_EXAMPLES:
            status = rng.choices(
                config.INTERNAL_EDU_STATUS_CHOICES, weights=[75, 10, 15], k=1
            )[0]
            internal_edu.append({
                "社員番号": emp_id,
                "教育名": edu_name,
                "実施年月": _random_date(join.year, today.year, rng, no_future=True) if status != "未受講" else None,
                "実施者": rng.choice(["班長", "品質課", "外部講師"]) if status != "未受講" else "",
                "修了区分": status,
                "備考": "",
            })

        # スキル習熟度（作業項目ごとに0〜5、経験の少ない人は低めに偏らせる）
        base = rng.choice([1, 2, 2, 3])
        emp_skills = {}
        for item in skill_items:
            level = min(5, max(0, base + rng.randint(-1, 2)))
            emp_skills[item] = level
        skills[emp_id] = emp_skills

        # 品質不具合歴（0〜3件）。一部は「スキルレベルの一時的な見直し」のサンプルとして
        # 見直し対象スキル項目・見直し後レベル・再認定日を設定する
        for _ in range(rng.randint(0, 3)):
            occurred = _random_date(max(join.year, today.year - 3), today.year, rng, no_future=True)
            category = rng.choice(config.QUALITY_CATEGORY_CHOICES)
            record = {
                "社員番号": emp_id,
                "発生年月": occurred,
                "区分": category,
                "内容": f"{category}による出荷ミスが発生",
                "原因": rng.choice(config.QUALITY_CAUSE_CHOICES[:-1]),
                "再発防止として実施した教育・対策": rng.choice(["品質教育を再受講", "ダブルチェック手順を追加", "KY教育を実施"]),
                "再発の有無": rng.choices(config.YES_NO_CHOICES, weights=[15, 85], k=1)[0],
                "備考": "",
                "見直し対象スキル項目": None,
                "見直し後レベル": None,
                "再認定日": None,
            }
            # 習熟度Lv2以上の項目があれば、3件に1件ほどの割合で「見直し中」のサンプルを作る。
            # 再認定日は「今日」からの相対日数で決め、現在も有効なものと既に再認定済みのものを混ぜる
            # （動作確認時に見直し中フラグが実際に表示されるようにするため）
            candidate_items = [it for it in skill_items if emp_skills.get(it, 0) >= 2]
            if candidate_items and rng.random() < 0.35:
                target_item = rng.choice(candidate_items)
                record["見直し対象スキル項目"] = target_item
                record["見直し後レベル"] = max(0, emp_skills.get(target_item, 2) - 2)
                days_offset = rng.choice([-30, -10, 20, 45, 90])  # 負の値＝再認定済み（過去の日付）
                record["再認定日"] = today + timedelta(days=days_offset)
            quality.append(record)

    return {
        "employees": employees,
        "assignments": assignments,
        "licenses": licenses,
        "external_edu": external_edu,
        "internal_edu": internal_edu,
        "quality": quality,
        "skills": skills,
        "grade_targets": generate_grade_targets(skill_items),
    }
