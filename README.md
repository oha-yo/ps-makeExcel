# Convert-CsvToXlsx.ps1

PowerShell スクリプト `Convert-CsvToXlsx.ps1` は、CSV 形式または区切り付きテキストファイルを Excel（`.xlsx`）形式に変換するためのツールです。  
Shift_JIS や UTF-8 などのエンコーディングにも対応し、列ヘッダーは自動で通番を付与します。

## 特徴

- 入力ファイルの拡張子は自動的に `.xlsx` に変換
- Shift_JIS など任意の文字コードで読み込み可能
- 開始行（ヘッダー除去）や最大読み込み行数を指定可能
- 引用付きCSV (`"aaa","bbb"` 形式) のパースに対応
- EPPlus を使用して `.xlsx` を高速・安定に生成
- 出力先ファイルが既に開かれている場合はエラーメッセージを表示

## 必須要件

- PowerShell 7.x
- [EPPlus.dll](https://www.nuget.org/packages/EPPlus)  
  ※本スクリプトでは ImportExcel モジュールに同梱された DLL を使用します。

## スクリプトの構成

```bash
Convert-CsvToXlsx.ps1
