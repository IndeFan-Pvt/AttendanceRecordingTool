# 3分で使う GUI版

## 使うもの

- exe/generate_akanecco_shift_gui フォルダ
- 記載したい勤怠表の temp ファイル
  - 例: 【統一書式】あかねっこ2月_temp.xls
- 必要なら前月勤務表
  - 例: 【統一書式】あかねっこ1月.xls

## 事前確認

- Windows PC であること
- Microsoft Excel が入っていること

## 手順

### 1. GUI を起動する

次を開きます。

- exe/generate_akanecco_shift_gui/generate_akanecco_shift_gui.exe

### 2. GUI でファイルを選ぶ

最低限、以下を指定します。

1. 記載したい勤怠表 (.xls / .xlsx)
   - 例: 【統一書式】あかねっこ2月_temp.xls
2. 設定 JSON
   - exe 版では exe/generate_akanecco_shift_gui/akanecco_shift_config.json を使います

必要なら以下も指定します。

3. 前月勤務表 (任意)
   - 月初引継ぎを確実に使いたいとき
4. レポート保存先 (任意)

### 3. 実行する

- 生成実行 を押します

## 設定 JSON でよく使う職員フラグ

職員ごとの配慮条件は、自動ではなく設定 JSON で明示した職員だけに適用します。

- night_fairness_target
   - 夜勤回数の平等化対象にする
- weekend_fairness_target
   - 土日休系回数の平等化対象にする
- unit_shift_balance_target
   - 同じユニット内で早番・遅番回数の偏りを抑える対象にする
- preferred_four_day_streak_target
   - 4連勤が月1回程度に収まるよう配慮する対象にする
- require_standard_day
   - 通常の「日」を月1回以上入れたい対象にする

設定例:

```json
{
   "employee_id": "employee-001",
   "display_name": "職員A",
   "unit": "unit-a",
   "employment": "full",
   "allowed_shifts": ["早", "遅", "日", "夜", "休"],
   "night_fairness_target": true,
   "weekend_fairness_target": true,
   "unit_shift_balance_target": true,
   "preferred_four_day_streak_target": true,
   "require_standard_day": true
}
```

対象外にしたい場合は、そのフラグを false にします。

## 勤務表から直接変えられる項目

勤務表に見出し列を追加すれば、次の項目は JSON を開かなくても職員ごとに設定できます。

- 夜勤公平化対象
- 夜夜必須対象
- 夜夜必須回数
- 土日休公平化対象
- 個別連勤上限
- 4連勤許容回数
- 早遅平準化対象
- 4連勤配慮対象
- 日勤候補対象
- 休系回数指定
- 早番MAX
- 日勤MAX
- 遅番MAX
- 夜勤MAX
- 勤務可能一覧
- 曜日別勤務制限
- 日付別勤務制限
- 指定日の日勤増員

休系回数指定は、休だけではなく 休・特・夜休 を合計した回数です。

勤務可能一覧、曜日別勤務制限、日付別勤務制限は、次のように 1 セルへ書きます。

- 区切りは ; または改行
- キーと勤務記号は = または : でつなぐ
- 勤務記号どうしは / で区切る
- 空欄勤務を含めるときは 空欄 と書く

入力例:

```text
勤務可能一覧: 早/遅/日/夜/休
曜日別勤務制限: 金=早/遅/日/夜/休; 土=早/遅/日/夜/休; 日=早/遅/日/夜/休
日付別勤務制限: 5=休; 17=休; 21=早/日
```

勤務可能一覧は、その職員がその月に取り得る勤務記号の全体です。夜勤を含めるときは、夜休は自動補完されます。セルが空欄なら JSON 側の allowed_shifts をそのまま使います。

列見出しがある月は、その列の内容がその月の制限として使われます。セルが空欄なら、その月は曜日別または日付別の追加制限なしとして扱います。

指定日の日勤増員は、月次設定セルに次のように書きます。

- 区切りは ; または改行
- 1件ごとに 日付=勤務:人数 または 日付=勤務:最小-最大 と書く
- 同じ日に複数の勤務を指定するときは , で区切る

入力例:

```text
指定日の日勤増員: 5=日:1-2; 17=日:2
```

一方で、職員一覧そのものや行番号、ユニット名のような勤務表構造に関わる基本情報は、引き続き JSON 側で持ちます。

## 月ごとに設定を変えたいとき

月別に一部だけ変えたい場合は、period_overrides を使います。

- キーは YYYY-MM 形式にする
- その月だけ変えたい項目だけを書く
- employees を書く場合は、その月に使う職員一覧を丸ごと書く

例: 2026年1月だけ職員一覧を月別定義に差し替え、職員Aのフラグと固定休を変える

```json
{
   "period_overrides": {
      "2026-01": {
         "rules": {
            "max_consecutive_rest": 3
         },
         "employees": [
            {
               "employee_id": "employee-001",
               "display_name": "職員A",
               "unit": "unit-a",
               "employment": "full",
               "allowed_shifts": ["早", "遅", "日", "夜", "休"],
               "night_fairness_target": true,
               "weekend_fairness_target": false,
               "unit_shift_balance_target": true,
               "preferred_four_day_streak_target": false,
               "require_standard_day": true,
               "fixed_assignments": {
                  "3": "休",
                  "17": "休"
               }
            },
            {
               "employee_id": "employee-002",
               "display_name": "職員B",
               "unit": "unit-b",
               "employment": "part",
               "allowed_shifts": ["", "日", "休"],
               "night_fairness_target": false,
               "weekend_fairness_target": false,
               "unit_shift_balance_target": false,
               "preferred_four_day_streak_target": false,
               "require_standard_day": false,
               "fixed_assignments": {}
            }
         ]
      }
   }
}
```

職員を period_overrides で書くときは、employee_id を必ず合わせます。月別設定では、通常設定と同じフラグ名をそのまま使えます。employees を指定した月は、通常設定の employees 全体ではなく、その月の employees 配列が採用されます。

### 4. 結果を確認する

- 完了後に Excel を開く にチェックがあれば、勤怠表が開きます
- 完了後にレポートを開く にチェックがあれば、検証レポートが開きます

## うまくいかないとき

### Excel が開かない

- Excel がインストールされているか確認する

### ファイルが見つからない

- temp ファイルの場所を確認する
- 設定 JSON の場所を確認する
- 必要なら前月勤務表を指定する

### 生成に失敗する

- 実行ログ欄の内容を確認する
- 同月の参照元勤務表があるか確認する

## 一番簡単な使い方

1. GUI を起動
2. temp の xls を選択
3. 設定 JSON を選択
4. 必要なら前月勤務表を選択
5. 生成実行