# Changelog
All notable changes to the "excel-vba-sync" extension are documented here.

This file follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/)
and uses [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]
- バッチ生成（複数文字列の一括処理）を追加予定
- ログ出力の強化（`--verbose`）

## [1.1.0] - 2025-10-02
### Added
- PNG を **PngByteQRCode** で生成する実装（System.Drawing 非依存）
- 相対パスやファイル名のみを **ドキュメント直下**に解決する仕様
- 例外時の終了コードとメッセージ整備（`2=入力エラー`, `1=想定外エラー`）

### Changed
- 使い方メッセージの日本語化・改善
- README に Excel 連携の例を追記

## [1.0.0] - 2025-09-15
### Added
- 初版リリース。CLI から **SVG/PNG** 生成、ECC（L/M/Q/H）指定に対応