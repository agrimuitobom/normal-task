# 教員業務サポートリポジトリ

教員の日常業務を効率化するためのテンプレート・スクリプト集です。

## ディレクトリ構成

```
normal-task/
├── README.md                    # このファイル
├── templates/                   # 文書テンプレート
│   ├── lesson_plan.md           # 授業計画書テンプレート
│   ├── notice_letter.md         # 保護者向け通知文テンプレート
│   ├── meeting_minutes.md       # 会議議事録テンプレート
│   └── approval_request.md      # 稟議書テンプレート
├── scripts/                     # 自動化スクリプト
│   └── grade_summary.py         # 成績集計スクリプト
└── docs/                        # マニュアル・業務メモ
    └── workflow.md              # 業務フロー説明
```

## 各ファイルの使い方

### テンプレート (`templates/`)
- `lesson_plan.md` — 授業計画書のひな形。科目・単元・目標・展開を記入して使用
- `notice_letter.md` — 保護者向けお便り・通知文のひな形
- `meeting_minutes.md` — 職員会議・学年会議の議事録ひな形
- `approval_request.md` — 稟議書のひな形。費用・見積もり比較・決裁欄付き

### スクリプト (`scripts/`)
- `grade_summary.py` — CSVファイルから生徒の成績を集計・統計表示するPythonスクリプト

### ドキュメント (`docs/`)
- `workflow.md` — 年間・月間の業務フローと管理方法の説明

## 利用方法

1. テンプレートを複製して使用してください
2. スクリプトはPython 3.8以上が必要です
3. ご要望があれば随時テンプレート・スクリプトを追加できます
