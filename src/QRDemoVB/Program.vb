'
'   File: export_opened_vba.ps1
'   Description: QRコードを作成する
'   Author: Eitaro SETA
'   License: MIT License
'   Copyright (c) 2025 Eitaro SETA
'
'   Permission is hereby granted, free of charge, to any person obtaining a copy
'   of this software and associated documentation files (the "Software"), to deal
'   in the Software without restriction, including without limitation the rights
'   to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'   copies of the Software, and to permit persons to whom the Software is
'   furnished to do so, subject to the following conditions:
'
'   The above copyright notice and this permission notice shall be included in all
'   copies or substantial portions of the Software.
'
'   THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'   IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'   FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'   AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'   LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'   OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'   SOFTWARE.
'

Imports System
Imports System.IO
Imports QRCoder

Module Program
    ' 出力フォーマットの列挙型
    Enum OutFmt
        PNG
        SVG
    End Enum

    Sub Main(args As String())
        Try
            ' --- 引数チェック 必須は3つ---
            ' --- 引数: 0=文字列(必須), 1=フォーマット(svg|png)(必須), 2=出力パス(必須), 3=ECC(L|M|Q|H)(任意) ---
            If args.Length < 3 Then Throw New ArgumentException("引数が足りません。")

            ' 入力された値を取り出し＆整形
            Dim payload As String = args(0)                              ' QRコード化する中身
            Dim fmt As OutFmt = ParseFormat(args(1))                     ' "png"/"svg" を列挙型に変換
            Dim outPath As String = NormalizeOutputPath(args(2), fmt)    ' 相対パスなら「ドキュメント直下」へ設定
            ' 誤り訂正レベル（ECC）を変換。指定が無ければ Q を使う。
            Dim ecc As QRCodeGenerator.ECCLevel = ParseEcc(If(args.Length >= 4, args(3), "Q"))

            ' --- QR データ生成（UTF-8, 任意ECC）---
            Dim gen As New QRCodeGenerator()
            Dim data = gen.CreateQrCode(payload,
                                        ecc,
                                        forceUtf8:=True,
                                        utf8BOM:=False,
                                        eciMode:=QRCodeGenerator.EciMode.Utf8)

          ' --- 画像として出力 ---
          ' ppm = pixels per module（QRの1マスを何ピクセルで描くか）
          ' 数値を上げるほど画像が大きく・読み取りやすくなる（印刷用途は 12～20 目安）
            Dim ppm As Integer = 16 ' 画像サイズ（1マスのピクセル数）
            Select Case fmt
                Case OutFmt.SVG
                    ' SVG（テキストベクタ形式）：拡大縮小に強く、環境依存も少ない
                    Dim svgQr As New SvgQRCode(data)
                    Dim svgText As String = svgQr.GetGraphic(ppm)
                    File.WriteAllText(outPath, svgText, System.Text.Encoding.UTF8)

                Case OutFmt.PNG
                    ' PNG（ビットマップ）：PngByteQRCode を使うことで System.Drawing 依存を回避
                    ' （= .NET 6+ での非Windows制限を気にせず使える）
                    Dim pngQr As New PngByteQRCode(data)
                    Dim bytes As Byte() = pngQr.GetGraphic(ppm)
                    File.WriteAllBytes(outPath, bytes)
            End Select

            ' 正常終了メッセージ（呼び出し元がログで拾えるように標準出力へ）
            Console.WriteLine("OK: " & outPath)
        Catch ex As ArgumentException
            ' 想定内の入力エラー：使い方を表示し、終了コード 2 で返す
            Console.Error.WriteLine("ERROR: " & ex.Message)
            ShowUsage()
            Environment.Exit(2)
        Catch ex As Exception
            ' 想定外の例外（ファイル書き込み不可など）：スタックを含めて標準エラーへ
            Console.Error.WriteLine("ERROR: " & ex.ToString())
            Environment.Exit(1)
        End Try
    End Sub

    ' 文字列の "svg" / "png" を列挙型へ変換（日本語入力の "SVGのみ"/"PNGのみ" も許容）
    Private Function ParseFormat(s As String) As OutFmt
        Dim k = s.Trim().ToLowerInvariant()
        If k = "svg" OrElse k = "svgのみ" Then Return OutFmt.SVG
        If k = "png" OrElse k = "pngのみ" Then Return OutFmt.PNG
        Throw New ArgumentException("出力フォーマットは 'SVG' または 'PNG' を指定してください。")
    End Function

    ' 誤り訂正レベル（ECC）文字を QRCoder の列挙型へ変換
    ' L < M < Q < H の順で強くなる（強いほど冗長ビットが増え、QRのサイズも大きくなる）
    Private Function ParseEcc(s As String) As QRCodeGenerator.ECCLevel
        Dim k = s.Trim().ToUpperInvariant()
        Select Case k
            Case "L" : Return QRCodeGenerator.ECCLevel.L
            Case "M" : Return QRCodeGenerator.ECCLevel.M
            Case "Q" : Return QRCodeGenerator.ECCLevel.Q
            Case "H" : Return QRCodeGenerator.ECCLevel.H
            Case Else
                Throw New ArgumentException("誤り訂正レベルは L/M/Q/H のいずれかを指定してください。")
        End Select
    End Function

    ' 出力パスの標準化:
    ' - 相対パスやファイル名だけが渡された場合は「ドキュメント直下」に解決
    ' - 拡張子は指定フォーマット（SVG/PNG）に強制（不一致なら置き換え）
    ' - ファイル名に使えない文字は "_" に置換（簡易サニタイズ）
    ' - ※ フォルダは自動作成しない（存在しない場合は呼び出し側で対処 or 例外）
    Private Function NormalizeOutputPath(spec As String, fmt As OutFmt) As String
        ' 拡張子をフォーマットに合わせて強制（不一致なら差し替え）
        Dim requiredExt As String = If(fmt = OutFmt.SVG, ".svg", ".png")

        ' 入力文字列をトリム
        Dim p As String = spec.Trim()

        ' フォルダ指定が無い場合は「Document直下」を基準にする（フォルダは作成しない）
        If Not System.IO.Path.IsPathRooted(p) AndAlso Not p.StartsWith("\\") Then
            Dim docs As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            p = System.IO.Path.Combine(docs, p)  ' 例: C:\Users\<you>\Documents\ファイル名.png
        End If

        ' 無効文字の置換（簡易サニタイズ）
        Dim dir As String = System.IO.Path.GetDirectoryName(p)
        Dim name As String = System.IO.Path.GetFileName(p)
        For Each ch In System.IO.Path.GetInvalidFileNameChars()
            name = name.Replace(ch, "_"c)
        Next

        p = If(String.IsNullOrEmpty(dir), name, System.IO.Path.Combine(dir, name))

         ' 拡張子をフォーマットに合わせて強制（例: .txt が渡されたら .png/.svg に差し替え）
        Dim ext = System.IO.Path.GetExtension(p)
        If String.IsNullOrEmpty(ext) OrElse Not String.Equals(ext, requiredExt, StringComparison.OrdinalIgnoreCase) Then
            p = System.IO.Path.ChangeExtension(p, requiredExt)
        End If

        Return p
    End Function

    ' コンソールに使い方を整形して出力（入力ミス時など）
    Private Sub ShowUsage()
        Console.Error.WriteLine("")
        Console.Error.WriteLine("使い方:")
        Console.Error.WriteLine("  QrDemoVB.exe ""<QRコード化したい文字列>"" <SVG|PNG> ""<出力先フォルダ\ファイル名>"" 誤り訂正レベル[L|M|Q|H]")
        Console.Error.WriteLine("")
        Console.Error.WriteLine("例:")
        Console.Error.WriteLine("  QrDemoVB.exe ""端末ID=A111-222; 期限=20xx-12-31"" PNG ""C:\temp\ab001.png"" H")
        Console.Error.WriteLine("  QrDemoVB.exe ""https://example.com/日本語"" SVG ""label.svg""")
        Console.Error.WriteLine("")
    End Sub
End Module