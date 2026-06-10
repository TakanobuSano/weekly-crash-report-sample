# weekly-crash-report-sample

Claude Code / Claude Cowork で **Crashlytics の週次レポート（PowerPoint）を自動生成**する仕組みの、サニタイズ済みサンプルです。
架空の電子書籍アプリ「BookShelf」を題材にしています。実在のアプリ・データとは関係ありません。

関連記事: [AIにスライドを丸ごと作らせない —— Claude Code × Crashlytics 週次レポート自動生成の設計
](https://qiita.com/4q_sano/items/5d188ee09e57ec6705cb)

## このサンプルの設計思想

このリポジトリで一番伝えたいのは、ツールの使い方ではなく **役割分担の設計** です。

> **AI にスライドを丸ごと作らせない。レイアウトはテンプレートに固定し、AI はデータ取得と JSON 作成に専念させる。**

LLM に「Crashlytics を見ていい感じのスライドを作って」と毎回任せると、回ごとに表がはみ出したり、配色やレイアウトがブレます。そこで責務を 2 つに分けています。

| 担当 | 役割 | 変わるもの |
| --- | --- | --- |
| プロンプト（AI） | データ取得・検証・JSON 作成 | 数字・文章（毎週変わる） |
| テンプレート（Python） | レイアウト・配色・グラフ描画 | 構造（固定。崩れない） |

この分担により、「見た目を直したい」ときはスクリプトだけ、「取得するデータを変えたい」ときはプロンプトだけ、と修正範囲が分離します。

## 構成

```
weekly-crash-report-sample/
├── README.md
├── prompts/
│   └── weekly_report_prompt.md   # 定期実行タスクに登録するプロンプト
└── scripts/
    ├── build_report.py           # レイアウトを固定したテンプレート本体
    └── report_data.sample.json   # JSON スキーマの実例（このサンプルの入力）
```

## 使い方

```bash
pip install python-pptx matplotlib
python scripts/build_report.py \
  --data scripts/report_data.sample.json \
  --out  weekly_report.pptx
```

`report_data.sample.json` を別の内容に差し替えれば、同じレイアウトで中身だけ変わります。
実運用では、この JSON を Claude（プロンプト）が Firebase Crashlytics MCP から生成します。

## テンプレートが引き受けている「判断」

単なる描画だけでなく、レポートの一貫性を保つロジックもテンプレート側に寄せています。

- 傾向の矢印（↑ ↓ →）は前週比の符号から自動決定（±5% 未満は横ばい）。矢印と符号の不一致が起きない
- 取得できない集計は `null` を渡すと「データなし」「—」と表示。推定で埋めない
- 新規 Issue（前週・前々週に無い）はグラフに `new!!`、テーブルに「新規」と自動表示
- セル内の行数を単語折り返しシミュレーションで見積もり、行高さを内容に合わせる
- 数値は 1,000 以上で自動カンマ区切り

「視覚的な差異は情報の差異に対応させる」——色やマークは装飾ではなく意味に対応させる、という原則で統一しています。

## ライセンス

MIT
