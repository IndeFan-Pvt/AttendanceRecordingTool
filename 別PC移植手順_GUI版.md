# あかねっこ勤怠生成 GUI版 別PC移植手順

## 1. 前提

- 移植先 PC は Windows であること
- 移植先 PC に Microsoft Excel がインストールされていること
- Python のインストールは不要
- 使用する実行ファイルは GUI 版とする

対象フォルダ:

- exe/generate_akanecco_shift_gui

## 2. 移植するもの

別 PC にコピーするものは以下です。

### 2-1. アプリ本体

以下のフォルダを丸ごとコピーします。

- exe/generate_akanecco_shift_gui

この中には以下が含まれます。

- generate_akanecco_shift_gui.exe
- akanecco_shift_config.json
- run_generate_akanecco_shift_gui.bat
- _internal

### 2-2. 勤怠表データ

必要に応じて以下をコピーします。

- 記載したい勤怠表
  - 例: 【統一書式】あかねっこ2月_temp.xls
- 同月の参照元勤務表
  - 例: 【統一書式】あかねっこ2月.xls
- 前月勤務表
  - 例: 【統一書式】あかねっこ1月.xls

## 3. 推奨フォルダ構成

別 PC では以下のような構成を推奨します。

```text
任意の作業フォルダ/
  generate_akanecco_shift_gui/
    generate_akanecco_shift_gui.exe
    akanecco_shift_config.json
    run_generate_akanecco_shift_gui.bat
    _internal/
  勤怠データ/
    【統一書式】あかねっこ2月_temp.xls
    【統一書式】あかねっこ2月.xls
    【統一書式】あかねっこ1月.xls
```

ポイント:

- GUI 本体と勤怠データは別フォルダでもよい
- GUI 上でファイルを選択するため、厳密に同じパス構成でなくてもよい
- ただし、同月参照元と前月勤務表は分かりやすい場所に置くこと

## 4. 実行手順

### 4-1. GUI の起動

以下のどちらかで起動します。

- generate_akanecco_shift_gui/run_generate_akanecco_shift_gui.bat
- generate_akanecco_shift_gui/generate_akanecco_shift_gui.exe

### 4-2. GUI で指定する項目

GUI が開いたら以下を指定します。

1. 記載したい勤怠表 (.xls / .xlsx)
   - 例: 【統一書式】あかねっこ2月_temp.xls
2. 設定 JSON
   - 例: generate_akanecco_shift_gui/akanecco_shift_config.json
3. 前月勤務表 (任意)
   - 月初の勤務引継ぎを確実に反映したい場合に指定
   - 例: 【統一書式】あかねっこ1月.xls
4. レポート保存先 (任意)
   - 空欄でも可
   - 空欄の場合は対象ファイルの隣に *_validation.html を作成

必要に応じて以下も選択します。

- 完了後にレポートを開く
- 完了後に Excel を開く

最後に「生成実行」を押します。

## 5. 実運用時のおすすめ

- 最初はコピーしたテスト用ファイルで 1 回試す
- 本番用ファイルは必ずバックアップを取ってから実行する
- 前月引継ぎを使う月は、前月勤務表を GUI で明示指定する
- 同月の参照元勤務表も一緒に保管しておく

## 6. うまく動かないときの確認項目

### 6-1. GUI が起動しない

- Windows か確認する
- generate_akanecco_shift_gui フォルダを丸ごとコピーしたか確認する
- _internal フォルダが欠けていないか確認する

### 6-2. Excel が開けない、または書き込めない

- Microsoft Excel がインストールされているか確認する
- 対象の xls が別の Excel で開きっぱなしでないか確認する
- 保護ビューで開かれていないか確認する

### 6-3. 生成に失敗する

- 設定 JSON が正しいか確認する
- 対象ファイルが temp 側か確認する
- 必要なら前月勤務表を指定する
- 同月の参照元勤務表が存在するか確認する

### 6-4. 月初引継ぎが不安

- GUI の「前月勤務表 (任意)」で前月 xls を明示指定する
- 指定後に生成し、レポートの「月初への前月末勤務引継ぎ」を確認する

## 7. 移植時の注意

- Python は不要だが Excel は必要
- ネットワークドライブ上のファイルは保護ビューになることがある
- 日本語ファイル名を変更すると、運用手順が分かりにくくなることがある
- 設定 JSON の中身を変更する場合は、コピー後のファイルを編集する

## 8. 最低限の配布チェックリスト

- generate_akanecco_shift_gui フォルダを丸ごとコピーした
- generate_akanecco_shift_gui.exe がある
- akanecco_shift_config.json がある
- _internal フォルダがある
- 記載対象の temp xls を配置した
- 必要なら同月参照元 xls を配置した
- 必要なら前月 xls を配置した
- Excel インストール済み PC である

## 9. 推奨手順の要約

1. generate_akanecco_shift_gui フォルダを別 PC にコピーする
2. 勤怠データを別 PC にコピーする
3. bat または exe を起動する
4. GUI で対象 xls、設定 JSON、必要なら前月勤務表を選ぶ
5. 生成実行する
6. 必要なら生成後に Excel とレポートを確認する