"""
grade_summary.py - 成績集計スクリプト

【使い方】
    python grade_summary.py <CSVファイルパス>

【CSVファイルの形式】
    1行目: ヘッダー（氏名,国語,数学,英語,理科,社会 など）
    2行目以降: 生徒データ

【例: grades.csv】
    氏名,国語,数学,英語,理科,社会
    田中太郎,85,90,78,88,92
    山田花子,72,65,80,70,75
    ...

【出力】
    - 生徒ごとの合計点・平均点
    - 科目ごとの平均・最高・最低点
    - 成績段階分布
"""

import csv
import sys
from pathlib import Path


def load_csv(filepath: str) -> tuple[list[str], list[dict]]:
    """CSVを読み込み、ヘッダーと行データを返す。"""
    path = Path(filepath)
    if not path.exists():
        print(f"[エラー] ファイルが見つかりません: {filepath}")
        sys.exit(1)

    with open(path, encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        headers = reader.fieldnames or []
        rows = list(reader)

    return list(headers), rows


def summarize(headers: list[str], rows: list[dict]) -> None:
    """成績の集計・表示を行う。"""
    if not rows:
        print("データがありません。")
        return

    name_col = headers[0]
    score_cols = headers[1:]  # 氏名以外は得点科目とみなす

    # 数値変換
    for row in rows:
        for col in score_cols:
            try:
                row[col] = float(row[col])
            except (ValueError, KeyError):
                row[col] = 0.0

    print("=" * 60)
    print("  生徒別 成績一覧")
    print("=" * 60)

    # 生徒別集計
    header_line = f"{'氏名':<12}" + "".join(f"{c:>8}" for c in score_cols) + f"{'合計':>8}{'平均':>8}"
    print(header_line)
    print("-" * len(header_line))

    for row in rows:
        scores = [row[c] for c in score_cols]
        total = sum(scores)
        average = total / len(scores) if scores else 0
        line = f"{row[name_col]:<12}" + "".join(f"{s:>8.1f}" for s in scores)
        line += f"{total:>8.1f}{average:>8.1f}"
        print(line)

    print()
    print("=" * 60)
    print("  科目別 統計")
    print("=" * 60)

    for col in score_cols:
        scores = [row[col] for row in rows]
        avg = sum(scores) / len(scores)
        print(f"  {col:<10} 平均: {avg:5.1f}  最高: {max(scores):5.1f}  最低: {min(scores):5.1f}")

    print()
    print("=" * 60)
    print("  成績段階分布（全科目の平均点）")
    print("=" * 60)

    grades = {"A（90以上）": 0, "B（80〜89）": 0, "C（70〜79）": 0, "D（60〜69）": 0, "E（60未満）": 0}
    for row in rows:
        scores = [row[c] for c in score_cols]
        avg = sum(scores) / len(scores) if scores else 0
        if avg >= 90:
            grades["A（90以上）"] += 1
        elif avg >= 80:
            grades["B（80〜89）"] += 1
        elif avg >= 70:
            grades["C（70〜79）"] += 1
        elif avg >= 60:
            grades["D（60〜69）"] += 1
        else:
            grades["E（60未満）"] += 1

    for grade, count in grades.items():
        bar = "■" * count
        print(f"  {grade}: {count:3d}人  {bar}")

    print("=" * 60)


def main():
    if len(sys.argv) < 2:
        print("使い方: python grade_summary.py <CSVファイルパス>")
        print("例:     python grade_summary.py grades.csv")
        sys.exit(1)

    headers, rows = load_csv(sys.argv[1])
    summarize(headers, rows)


if __name__ == "__main__":
    main()
