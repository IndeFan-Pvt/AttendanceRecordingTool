# 3分で使う GUI版

## 使うもの

- generate_akanecco_shift_gui フォルダ
- 記載したい勤怠表の temp ファイル
  - 例: 【統一書式】あかねっこ2月_temp.xls
- 必要なら前月勤務表
  - 例: 【統一書式】あかねっこ1月.xls

## 事前確認

- Windows PC であること
- Microsoft Excel が入っていること

## 手順

### 1. GUI を起動する

次のどちらかを開きます。

- exe/generate_akanecco_shift_gui/run_generate_akanecco_shift_gui.bat
- exe/generate_akanecco_shift_gui/generate_akanecco_shift_gui.exe

### 2. GUI でファイルを選ぶ

最低限、以下を指定します。

1. 記載したい勤怠表 (.xls / .xlsx)
   - 例: 【統一書式】あかねっこ2月_temp.xls
2. 必要なら前月勤務表 (任意)
   - 月初引継ぎを確実に使いたいとき
3. 必要ならレポート保存先 (任意)
4. 設定 JSON
   - exe 版では exe/generate_akanecco_shift_gui/akanecco_shift_config.json を固定で使います
   - exe 版では手動変更できません

補足です。

- exe フォルダ配下のファイルは相対パスで表示されることがあります
- 非常勤の行は出力しない設定になっています

### 3. 実行する

- 生成実行 を押します
- 実行中は下部に進捗バーと進捗パーセンテージが表示されます

### 4. 結果を確認する

- 完了後に Excel を開く にチェックがあれば、勤怠表が開きます
- 完了後にレポートを開く は初期状態ではオフです
- 必要ならチェックを入れると、完了後に検証レポートが開きます

## うまくいかないとき

### Excel が開かない

- Excel がインストールされているか確認する

### ファイルが見つからない

- temp ファイルの場所を確認する
- exe 版では設定 JSON の場所を変更しない
- exe/generate_akanecco_shift_gui/akanecco_shift_config.json があるか確認する
- 必要なら前月勤務表を指定する

### 生成に失敗する

- 実行ログ欄の内容を確認する
- 同月の参照元勤務表があるか確認する

## 一番簡単な使い方

1. GUI を起動
2. temp の xls を選択
3. 必要なら前月勤務表を選択
4. 生成実行