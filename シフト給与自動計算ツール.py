import pandas as pd
from pathlib import Path
from typing import List


def calculate_work_hours(df: pd.DataFrame) -> pd.DataFrame:

    """
    勤怠データから勤務時間（時間）を計算する関数。
    Excelの時刻型 / 文字列 / 数値 / 深夜跨ぎすべて対応。
    """

    # 日付を datetime に
    df["日付"] = pd.to_datetime(df["日付"]) #to_datetimeでdatetime型に変換

    # 出勤・退勤を datetime に変換
    df["出勤_dt"] = pd.to_datetime(
        df["日付"].astype(str) + " " + df["出勤"].astype(str), #一旦文字列にして、文字列からdatetime型に変換
        errors="coerce"
    )

    df["退勤_dt"] = pd.to_datetime(
        df["日付"].astype(str) + " " + df["退勤"].astype(str),
        errors="coerce"
    )

    # 日跨ぎ対応
    df.loc[df["退勤_dt"] < df["出勤_dt"], "退勤_dt"] += pd.Timedelta(days=1)

    # 勤務時間（時間）
    df["勤務時間"] = ((df["退勤_dt"] - df["出勤_dt"]).dt.total_seconds() / 3600).round(2)

    # 中間列削除
    df.drop(columns=["出勤_dt", "退勤_dt"], inplace=True)

    return df


def apply_wage_master(df: pd.DataFrame) -> pd.DataFrame:

    """
    名前に応じて時給を付与する
    """

    wage_map = {
    "田中": 1400,
    "鈴木": 1300,
    "佐藤": 1200,
    }

    df["時給"] = df["名前"].map(wage_map) #DF内の名前欄とPython内で指定した名前と時給を照らし合わせて、時給を割り当てる

    return df #戻り値を設定。時給を返す。


def calculate_salary(df: pd.DataFrame) -> pd.DataFrame:

    """
    勤務時間と時給から給与を計算する関数
    """

    # 給与計算（時間 × 時給）
    df["給与"] = df["勤務時間"] * df["時給"] #apply_wage_masterで算出した時給とcalculate_work_hoursで算出した勤務時間をかけて給与を算出

    return df #戻り値を設定。給与を返す。


def load_and_process_months(input_path: Path, months: List[int]) -> pd.DataFrame:

    """
    指定した複数月の勤怠シートを読み込み、
    勤務時間を計算して1つのDataFrameにまとめる。

    Parameters
    ----------
    input_path : Path
        勤怠Excelファイルが格納されたディレクトリパス
    months : List[int]
        対象とする月のリスト（例: [1, 2, 3]）

    Returns
    -------
    pd.DataFrame
    月別勤怠を結合した集計結果
    """

    monthly_dfs = [] #データフレームを格納するための空リスト作成

    for month in months: #該当月（1~3月）を一つずつ代入して処理
        sheet_name = f"{month}月" #該当月（1~3月）分のシートを確認
        df = pd.read_excel(input_path, sheet_name=sheet_name) #該当月（1~3月）のシートを読み込み

        df = calculate_work_hours(df) #該当シートの勤怠データから勤務時間を計算
        df = apply_wage_master(df) #時給を付与
        df = calculate_salary(df) #勤務時間×時給で自動的に給与を計算
        df["対象月"] = month #書き出し用。どの月の処理結果なのかを「対象月」としてわかりやすく記入。

        monthly_dfs.append(df) #計算処理が完了したデータフレームをリストに格納

    return pd.concat(monthly_dfs, ignore_index=True) #戻り値を設定。該当月（1~3月）の勤務時間が算出されたDFを縦に結合。それを返す。


def main():

    """実行用メイン処理"""

    # ===== 設定 =====
    year = 2026 #年単位を指定
    target_months = [1, 2, 3] #ターゲットしたい月を指定

    base_dir = Path(__file__).parent #現在のファイルをベースに指定
    input_path = base_dir / f"勤怠管理シート{year}.xlsx" #inputパスを指定、名前内に年数を入力、yearで指定した年数を代入。

    # ===== 処理 =====
    result_df = load_and_process_months(input_path, target_months) #シートの読み込み、計算までこの関数で処理（実際には計算はcalculate_work_hoursが処理している）

    # 必要な列のみ出力用に整理
    output_df = result_df[["日付", "名前", "出勤", "退勤", "勤務時間", "時給", "給与", "対象月"]] #勤務時間が算出されたDFが戻り値として返ってきているので、後は当てはめるだけ。

    # ===== 出力 =====
    output_path = base_dir / (f"勤怠管理_勤務時間付き_{year}_"f"{target_months[0]}-{target_months[-1]}月.xlsx") #シート名を指定。target_months[-1]で月が増えても最後の月を引っ張ってきてくれる。
    output_df.to_excel(output_path, index=False) #完成したDFをexcelファイルとして書き出し
    print("勤務時間の計算が完了しました") #完了報告

if __name__ == "__main__":
    main()
