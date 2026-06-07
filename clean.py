# -*- coding: utf-8 -*-
"""
顧客名簿クレンジングスクリプト
---------------------------------
汚いCSV (raw/顧客名簿_raw.csv) を読み込み、実務でよくある以下の問題を整形して
output/ にきれいなExcelと変更レポートを出力する。

対応している主な「汚れ」:
  - 前後の空白 / 全角スペース
  - 全角英数字・全角記号の混在        (NFKC正規化)
  - 会社名の表記ゆれ ((株)・㈱・（株）→ 株式会社 など)
  - メールアドレスの大文字小文字・全角＠
  - 電話番号のフォーマットばらつき     (区切り・全角・スペース)
  - 郵便番号の〒・ハイフン無し・全角
  - 都道府県の略記 (東京 → 東京都)
  - 登録日の表記ゆれ (和暦R6 / 令和6年 / 2024.1.5 等 → YYYY-MM-DD)
  - 売上金額の ¥・円・カンマ・全角数字・空欄
  - 顧客IDの重複行

使い方:
  python clean.py
"""

import re
import sys
import unicodedata
from datetime import date
from pathlib import Path

import pandas as pd

# Windowsのコンソールで日本語が文字化けしないようUTF-8で出力する
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

BASE = Path(__file__).resolve().parent
RAW = BASE / "raw" / "顧客名簿_raw.csv"
OUT_DIR = BASE / "output"
OUT_XLSX = OUT_DIR / "顧客名簿_clean.xlsx"
OUT_REPORT = OUT_DIR / "cleaning_report.txt"

# 変更内容を記録していくログ (レポート出力に使う)
change_log: list[str] = []


def nfkc(text):
    """全角英数字・記号を半角へ。Noneや非文字列はそのまま空文字に。"""
    if pd.isna(text):
        return ""
    return unicodedata.normalize("NFKC", str(text))


def strip_spaces(text: str) -> str:
    """前後の半角・全角スペースを除去し、連続スペースを1個に。"""
    text = text.replace("　", " ")  # 全角スペース → 半角
    return re.sub(r"\s+", " ", text).strip()


def clean_company(name) -> str:
    s = strip_spaces(nfkc(name))
    # NFKCで ㈱→(株) ㈲→(有) に分解されるので、そこから正式名称へ寄せる
    s = s.replace("(株)", "株式会社").replace("（株）", "株式会社")
    s = s.replace("(有)", "有限会社").replace("（有）", "有限会社")
    # 「○○（株）」のように末尾に付くケースも吸収
    s = re.sub(r"\s*株式会社\s*$", "株式会社", s) if s.endswith("株式会社") else s
    return s


def clean_email(addr) -> str:
    s = strip_spaces(nfkc(addr)).lower()
    return s


def clean_phone(num) -> str:
    """数字だけ取り出し、桁数から 03-1234-5678 形式へ整形。"""
    s = nfkc(num)
    digits = re.sub(r"\D", "", s)
    if not digits:
        return ""
    # 東京/大阪など市外局番2桁
    if digits.startswith(("03", "06")) and len(digits) == 10:
        return f"{digits[:2]}-{digits[2:6]}-{digits[6:]}"
    # 携帯 (070/080/090)
    if digits.startswith(("070", "080", "090")) and len(digits) == 11:
        return f"{digits[:3]}-{digits[3:7]}-{digits[7:]}"
    # その他10桁は 市外局番3桁 として整形 (052-123-4567 等)
    if len(digits) == 10:
        return f"{digits[:3]}-{digits[3:6]}-{digits[6:]}"
    return digits  # 想定外の桁数は数字のまま残す


def clean_postal(code) -> str:
    s = nfkc(code)
    digits = re.sub(r"\D", "", s)
    if len(digits) == 7:
        return f"{digits[:3]}-{digits[3:]}"
    return digits


PREF_SUFFIX = {
    "東京": "東京都",
    "大阪": "大阪府",
    "京都": "京都府",
    "愛知": "愛知県",
}


def clean_pref(pref) -> str:
    s = strip_spaces(nfkc(pref))
    return PREF_SUFFIX.get(s, s)


def clean_date(value) -> str:
    """様々な日付表記を ISO形式 (YYYY-MM-DD) に統一。失敗時は空文字。"""
    s = strip_spaces(nfkc(value))
    if not s:
        return ""

    # 和暦: 令和6年1月5日 / R6.1.5 / 令和6/1/5
    m = re.search(r"(?:令和|R)\s*(\d+)\D+(\d+)\D+(\d+)", s)
    if m:
        wareki_year, mo, d = (int(g) for g in m.groups())
        year = 2018 + wareki_year  # 令和元年 = 2019
        try:
            return date(year, mo, d).isoformat()
        except ValueError:
            return ""

    # 西暦: 2024/1/5 / 2024-01-05 / 2024.4.15 / 2024年1月5日
    m = re.search(r"(\d{4})\D+(\d{1,2})\D+(\d{1,2})", s)
    if m:
        y, mo, d = (int(g) for g in m.groups())
        try:
            return date(y, mo, d).isoformat()
        except ValueError:
            return ""
    return ""


def clean_amount(value):
    """¥ 円 カンマ 全角数字を除去し整数化。空欄は None。"""
    s = nfkc(value)
    digits = re.sub(r"[^\d]", "", s)
    if not digits:
        return None
    return int(digits)


def main():
    if not RAW.exists():
        raise SystemExit(f"入力ファイルが見つかりません: {RAW}")

    df = pd.read_csv(RAW, dtype=str, keep_default_na=False)
    n_in = len(df)
    change_log.append(f"入力行数: {n_in} 行")

    df["会社名"] = df["会社名"].map(clean_company)
    df["担当者名"] = df["担当者名"].map(lambda x: strip_spaces(nfkc(x)))
    df["メールアドレス"] = df["メールアドレス"].map(clean_email)
    df["電話番号"] = df["電話番号"].map(clean_phone)
    df["郵便番号"] = df["郵便番号"].map(clean_postal)
    df["都道府県"] = df["都道府県"].map(clean_pref)
    df["登録日"] = df["登録日"].map(clean_date)
    df["売上金額"] = df["売上金額"].map(clean_amount)

    # 欠損チェック (整形しても直らない「中身が無い」問題は報告する)
    empty_name = (df["担当者名"] == "").sum()
    empty_amount = df["売上金額"].isna().sum()
    bad_date = (df["登録日"] == "").sum()
    if empty_name:
        change_log.append(f"担当者名が空欄の行: {empty_name} 件 (要確認)")
    if empty_amount:
        change_log.append(f"売上金額が空欄/不明の行: {empty_amount} 件 (要確認)")
    if bad_date:
        change_log.append(f"日付を解釈できなかった行: {bad_date} 件 (要確認)")

    # 重複行の除去 (顧客IDをキーに、最初の1件を残す)
    dup_mask = df.duplicated(subset=["顧客ID"], keep="first")
    n_dup = int(dup_mask.sum())
    if n_dup:
        dup_ids = sorted(df.loc[dup_mask, "顧客ID"].unique().tolist())
        change_log.append(f"顧客ID重複により削除: {n_dup} 行 (ID: {', '.join(dup_ids)})")
    df = df[~dup_mask].reset_index(drop=True)

    change_log.append(f"出力行数: {len(df)} 行")

    OUT_DIR.mkdir(exist_ok=True)
    df.to_excel(OUT_XLSX, index=False)

    report = "顧客名簿 クレンジング結果レポート\n"
    report += "=" * 36 + "\n"
    report += "\n".join(f"- {line}" for line in change_log) + "\n"
    OUT_REPORT.write_text(report, encoding="utf-8")

    print(report)
    print(f"きれいなデータ: {OUT_XLSX}")
    print(f"変更レポート  : {OUT_REPORT}")


if __name__ == "__main__":
    main()
