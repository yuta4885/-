# 社員スキルマップ ツールのエントリーポイント
# 実行するたびに状況を判断して、次のどちらかを行う
#   ・output/社員スキルマップ.xlsx が無ければ → サンプルデータ入りで新規作成
#   ・すでにあれば                           → 入力内容から一覧表・抽出結果・個人シートを最新化
#
# 実行方法（Windows）: 「実行.bat」をダブルクリックするだけ
# 実行方法（コマンド）: python main.py

import openpyxl

from src import config, data_reader, extract, individual_sheet, list_sheet, sample_data, workbook_init


def init_workbook():
    """サンプルデータ10人分で新規にExcelブックを作成する（初回実行時）"""
    print("初回セットアップ：サンプルデータ10人分でExcelブックを新規作成します…")
    config.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    skill_items = list(config.DEFAULT_SKILL_ITEMS)
    data = sample_data.generate_sample_data(skill_items)
    data["skill_items"] = skill_items

    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # openpyxlが自動で作る空シートを削除

    workbook_init.build_input_sheets(wb, data, skill_items)
    list_sheet.build_list_sheet(wb, data["employees"], data["skills"], skill_items)
    extract.build_extract_sheet(wb, data)
    individual_sheet.build_all_individual_sheets(wb, data)

    wb.save(config.OUTPUT_FILE)
    print(f"作成しました：{config.OUTPUT_FILE}")
    print("「社員マスタ」などのシートに実際のデータを入力し、再度このスクリプトを実行すると")
    print("一覧表・抽出結果・個人シートが最新の内容に更新されます。")


def _warn_missing_skill_columns(skill_items):
    wb = openpyxl.load_workbook(config.OUTPUT_FILE, data_only=False)
    try:
        sheet_headers = [c.value for c in next(wb[config.SHEET_SKILL].iter_rows(min_row=1, max_row=1))]
    finally:
        wb.close()
    missing = [item for item in skill_items if item not in sheet_headers]
    if missing:
        print("※ 注意：「設定」シートに次の作業項目がありますが、「スキル習熟度」シートに列が見当たりません。")
        print("  スキル習熟度シートに同じ見出しの列を追加してから習熟度を入力してください：" + "、".join(missing))


def update_workbook():
    """既存のExcelブックの入力内容から、一覧表・抽出結果・個人シートを再生成する"""
    print("既存のExcelブックを読み込み、一覧表・抽出結果・個人シートを最新化します…")
    data = data_reader.load_all(config.OUTPUT_FILE)

    if not data["employees"]:
        print("※ 社員マスタにデータがありません。社員マスタシートに社員情報を入力してから再実行してください。")
        return

    _warn_missing_skill_columns(data["skill_items"])

    wb = openpyxl.load_workbook(config.OUTPUT_FILE, data_only=False)
    list_sheet.build_list_sheet(wb, data["employees"], data["skills"], data["skill_items"])
    extract.build_extract_sheet(wb, data)
    individual_sheet.build_all_individual_sheets(wb, data)
    wb.save(config.OUTPUT_FILE)

    print(f"更新しました：{config.OUTPUT_FILE}")


def main():
    if config.OUTPUT_FILE.exists():
        update_workbook()
    else:
        init_workbook()


if __name__ == "__main__":
    main()
