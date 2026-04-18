# GUI 検証用コピー手順

## 位置づけ

この資料は、開発中や受け入れ確認時に GUI 版を安全に試すための内部向け手順です。

配布後の通常運用では使わず、元の temp ファイルを直接上書きしたくない場面でだけ使います。

## 目的

GUI 版を試すたびに元の temp ファイルを直接上書きしないよう、毎回検証用コピーを作ってから操作します。

この手順では、次の固定ファイルを使います。

- 検証用勤怠表: _inspect/gui_verify_target.xls
- 検証用レポート: _inspect/gui_verify_report.html

## 事前準備

ワークスペースのルートで次を実行します。

```powershell
powershell -ExecutionPolicy Bypass -File .\prepare_gui_verify_files.ps1
```

既定では、old フォルダ配下にある最新の *_temp.xls を元ファイルとして検証用コピーを作ります。

- 例: old/【統一書式】あかねっこ2月_temp.xls

別の temp ファイルで試したいときは、SourcePath を指定します。

```powershell
powershell -ExecutionPolicy Bypass -File .\prepare_gui_verify_files.ps1 -SourcePath "old/【統一書式】あかねっこ1月_temp.xls"
```

## GUI で指定するもの

GUI を開いたら、次を指定します。

1. 記載したい勤怠表
   - _inspect/gui_verify_target.xls
2. 設定 JSON
   - exe/generate_akanecco_shift_gui/shift_config.json
3. レポート保存先
   - _inspect/gui_verify_report.html
4. 必要なら前月勤務表
   - 例: old/【統一書式】あかねっこ1月.xls

## 実行時のおすすめ

- 完了後に Excel を開く は外しておく
- 完了後にレポートを開く は必要に応じて使う
- 本番用 temp を開いたままにしない

## 確認ポイント

- 生成後に _inspect/gui_verify_target.xls が更新されていること
- 生成後に _inspect/gui_verify_report.html が作られていること
- レポート内の対象ファイル名が gui_verify_target.xls になっていること

## 使い終わったら

次回も同じコマンドを実行すれば、検証用コピーを作り直し、前回のレポートも消した状態からやり直せます。