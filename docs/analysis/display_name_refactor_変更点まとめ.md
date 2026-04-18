# display_name refactor 変更点まとめ

作成日: 2026年4月13日

## 概要

- 職員の内部識別を `employee_id`、画面表示とレポート表示を `display_name` に分離しました。
- 旧 `name` は設定読込の後方互換のために残し、内部キーとしては使わない形へ整理しました。

## 実装変更

- `generate_akanecco_shift.py`
  - `EmployeeConfig` に `display_name` を正式追加
  - 設定読込で `display_name` を優先し、未指定時のみ `name` を使用
  - solver / schedule の内部キーは `employee_id` を使用
  - 検証メッセージ、集計、HTML レポート表示は `display_name` を使用
- `generate_akanecco_shift_gui.py`
  - GUI 側の `EmployeeConfig` 再構築でも `display_name` を保持
  - 配布版 GUI からの生成処理も動作確認済み
- `akanecco_shift_config.json`
  - 各職員に `display_name` を追加
  - 表示用の正規表記を設定へ明示
- `employee_id設計案.html`
  - 現在の実装状態に合わせて、`display_name` を正式採用した説明へ更新

## 検証結果

- Python 実行で生成と検証が成功
- `schedule` キーは `employee_id` のまま維持
- 検証出力キーとレポート表示は `display_name` を使用
- `issue_count = 0` を確認
- PyInstaller で CLI / GUI の exe を再ビルド済み
- 配布先 `exe` フォルダへ最新 build を反映済み
- 配布先 GUI を実際に起動し、対象ファイル入力と生成実行を確認済み

## 成果物

- 最新 build: `dist/`
- 配布用反映先: `exe/`
- 履歴保存先:
  - `archive/dist-history/dist_display_name_refactor_20260413`
  - `archive/exe-history/exe_display_name_refactor_20260413`

## 補足

- Frozen CLI の検証レポートは exe 配下に出力されます。
- GUI の再ビルド時は既存 `dist/generate_akanecco_shift_gui` があると上書き確認で止まるため、`PyInstaller -y` を使う運用が安全です。