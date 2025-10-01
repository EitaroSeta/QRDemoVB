# QrDemoVB

VB.NET 製のシンプルな **QRコード生成 CLI** です。  
**QRCoder** を利用し、**PNG（PngByteQRCode）** / **SVG** をローカルで生成します。外部アクセス不要。

## 特徴
- 入力：`文字列（必須）` / `出力形式（PNG|SVG）` / `出力先ファイル名（必須）` / `ECC（L|M|Q|H, 任意）`
- PNGは **System.Drawing 非依存**（`PngByteQRCode` 使用）で .NET 6/7/8/9 でも安心
- 相対パスやファイル名のみを指定した場合は **ユーザーのドキュメント直下**に保存
- 単一EXE（self-contained / single-file）として発行可能

## 必要環境
- **.NET 8 SDK**（推奨）
- Windows / Linux / macOS（CLI版は可。PNGは System.Drawing 非依存の実装）
- QRCoder

## ビルド
```bash
dotnet build ./src/QrDemoVB

