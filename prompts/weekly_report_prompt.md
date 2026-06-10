## Crashlytics 週次レポート自動生成（テンプレート運用版）

Firebase MCP で Crashlytics のデータを取得して JSON にまとめ、テンプレートスクリプトで PowerPoint を生成する。
**スライドのデザイン・レイアウトはスクリプト側で固定されているため、このタスクでは python-pptx のコードを書かないこと。**

### 前提条件
- 対象: iOS アプリ（Swift）、BookShelf! Reader
- Firebase プロジェクト: example-bookshelf-prod
- App ID: 1:000000000000:ios:xxxxxxxxxxxxxxxx
- タイムゾーン: Asia/Tokyo
- 対象期間: 直近で完了している ISO 週（月曜〜日曜、JST）= 実行日が属する週の「1つ前の週」。実行日からの相対日数ではなく曜日で固定することで、実行が月曜でも水曜でも常に同じ週を対象にし、実行タイミングのブレを吸収する。対象週は完了済みのため当日分の集計不足も起きない
  （例: 6/8(月)〜6/10(水) のいずれに実行しても、対象期間は 6/1(月)〜6/7(日)）
- 前週 = 対象期間のさらに1つ前の ISO 週（月曜〜日曜）
- 前々週 = 前週のさらに1つ前の ISO 週（月曜〜日曜）。シート3の3週推移グラフ用
- 最小サポート iOS: 16.7
- テンプレート: ~/Documents/weekly-crash-report/scripts/build_crashlytics_report.py（配置済み）
- サンプルデータ: ~/Documents/weekly-crash-report/scripts/report_data.sample.json（JSON スキーマの実例）

### フォルダアクセス
~/Documents/weekly-crash-report フォルダへのアクセスをリクエストすること（request_cowork_directory ツールを使用）。

### データ取得手順
1. Firebase MCP の crashlytics_get_report を使って以下を取得する:
   - topIssues（FATAL、今週 + 前週 + 前々週の3期間）
   - topIssues（NON-FATAL、今週 + 前週。4枚目KPIカード用に上位合算値のみ使用）
   - topVersions（FATAL、今週）
   - topOperatingSystems（FATAL、今週）
2. 上位 FATAL Issue（目安5〜10件）のサンプルイベントを crashlytics_batch_get_events で取得し、スタックトレースを抽出する
3. 前週データとの比較で各 Issue の前週比（%）を算出する

**データ検証ルール:**
- topVersions / topOperatingSystems の集計値が明らかに不完全な場合（極端な少数件しか返らない等）、サンプルイベントからの推定で数値を補完しない。該当チャートは JSON で `"available": false` とし、`reason` に理由を書く（スクリプトが「データなし」パネルを表示する）
- 4枚目KPI（NON-FATAL 件数）が取得できない場合は `kpi.fourth.current` を null にする（スクリプトが「データなし」カードを表示する）
- 検証できない推定値・参考値を JSON に入れないこと

### JSON 作成
report_data.sample.json と同じスキーマで output/latest/report_data.json を作成する。要点:

- `kpi`: fatal_events / affected_users / issue_count / fourth。current と previous は数値（previous から前週比はスクリプトが計算）。scope には集計範囲（例: 「上位10件合算」「上位10件合算・重複含む」）を必ず入れる
- `kpi.fourth`（4枚目のKPIカード）: label は「NON-FATAL 件数」、current / previous は NON-FATAL topIssues の上位合算値、scope は「上位10件合算」
- KPI カードの scope / note は全角16文字程度までを目安に簡潔にする（超えるとカード内で折り返し表示になる）
- `summary_lines`: 全体傾向のサマリー 3〜4 行。同じ内容を繰り返さない
- `trend_chart`: **今週の件数上位の FATAL Issue（最大7件、今週件数の降順）を必ず採用する**。前週との比較ができないことを理由に今週の上位 Issue をグラフから除外しないこと。labels は英語の短縮 Issue 名、week_before_last / last_week / this_week は各週の件数（前々週→前週→今週の3週分）。各週の topIssues に含まれない（ランク外で実数が取れない）週は null にする（グラフでは「—」と表示され、0件との誤読を防げる）。insight には3週の推移（増加が継続か一過性か等）の所見を1〜2行（日本語）。前々週・前週ともランク外から今週上位入りした新規 Issue は insight で必ず言及する
- `issues`: 各 Issue に id / user_summary（「○○画面で△△すると落ちる」のようなユーザー視点表現。クラス名・スタックトレースを含めない）/ affected_users / wow_change_pct（整数%。傾向矢印はスクリプトが符号から自動決定。前週 topIssues に無くて算出できない場合は null）/ new（前週ランク外から新たに上位入りした Issue は true を付与。傾向列に「新規」と表示される）/ ticket（不明なら null）/ status / tech（class_method, stack_summary, repro, notes の技術視点4項目）
- `version_chart` / `os_chart`: available, title（英語）, labels, values, note（日本語）。取得不可なら available: false + reason
- `actions`: top_priority / monitoring（各項目に title, detail, 対応する issues 番号の配列）, footnote（データソース等の参考情報 1 行）

### PowerPoint 生成
```
pip install python-pptx matplotlib   # 未導入の場合のみ
python ~/Documents/weekly-crash-report/scripts/build_crashlytics_report.py \
  --data ~/Documents/weekly-crash-report/output/latest/report_data.json \
  --out  ~/Documents/weekly-crash-report/output/latest/crashlytics_weekly_report.pptx \
  
```
- スクリプトは生成後にレイアウト検品（スライドはみ出しチェック）を自動実行する。WARN が出た場合は、該当セルの文章（stack_summary / repro / notes など）を短く要約し直して再実行する
- スクリプトのレイアウト・色・フォント設定を書き換えないこと

### 保存先
- latest: output/latest/crashlytics_weekly_report.pptx（上記コマンドで生成済み）
- archive: output/archive/crashlytics_weekly_report_YYYYMMDD.pptx にコピー（YYYYMMDD は対象期間の終了日。同名は上書き）
- report_data.json も output/archive/report_data_YYYYMMDD.json としてコピーする（再生成・差分確認用）
- 保存先フォルダが無ければ作成する
- 処理失敗時は output/latest/error.txt に理由を保存

### 完了報告
生成結果を簡潔に報告する（主要KPI、上位Issue、保存先パス）。「データなし」とした項目がある場合はその旨と理由も報告に含める。
