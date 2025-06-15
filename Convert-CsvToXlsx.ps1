param(
    [Parameter(Mandatory = $true)][string]$InputFile,
    [Parameter()][string]$Encoding = "Shift_JIS",
    [Parameter()][int]$StartRow = 2,
    [Parameter()][int]$MaxRows = 0,
    [Parameter()][string]$Separator = ",",
    [Parameter()][bool]$AddColumnNumbers = $false
)
# CSVを分割し配列にして返却する関数
function SplitCsvLine {
    param(
        [string]$line,
        [string]$Separator
    )
    $pattern = "$Separator(?=(?:[^""]*""[^""]*"")*[^""]*$)"
    # 行を分割
    $csvline = [regex]::Split($line, $pattern)
    $csvline = $csvline | ForEach-Object {
        <#
          Trim() → 前後の空白を削除
          -replace '^"(.*)"$', '$1' → 前後のダブルクォートを削除
          -replace '""', '"' → 二重クォートを単一のクォートに変換
        #>
        $_.Trim() -replace '^"(.*)"$', '$1' -replace '""', '"'
    }
    return $csvline
}

# EPPlus.dll の読み込み（ImportExcelモジュールから直接）
$epplusPath = ".\Modules\ImportExcel\7.8.10\EPPlus.dll"
# EPPlus.dllを PowerShell に読み込んで、型情報（.NETクラス）を使えるようにする。
Add-Type -Path $epplusPath
[Reflection.Assembly]::LoadFrom($epplusPath) | Out-Null

# 入力ファイルチェック
if (-not (Test-Path $InputFile)) {
    Write-Error "Input file does not exist: $InputFile"
    exit 1
}
Write-Debug "Separator: $Separator"

# 出力ファイル名を自動生成
$OutputFile = [System.IO.Path]::ChangeExtension($InputFile, "xlsx")

# ファイル読み取り（ストリーム）
$reader = [System.IO.StreamReader]::new($InputFile, [System.Text.Encoding]::GetEncoding($Encoding))

# 指定行までスキップ
$currentLineNumber = 0
while (-not $reader.EndOfStream -and $currentLineNumber -lt ($StartRow - 1)) {
    $reader.ReadLine() | Out-Null
    $currentLineNumber++
}

# 最初のデータ行を読み取り、列数を判定
$firstDataLine = $reader.ReadLine()
$currentLineNumber++

# カラム数をカウント
$columnCount = (SplitCsvLine -line $firstDataLine -Separator $Separator ).Count

# 通番でヘッダー作成
$headers = @(1..$columnCount | ForEach-Object { "$_" })

# データ行を List[string] で読み込み（1行目含む）
$linesToProcess = [System.Collections.Generic.List[string]]::new()
$linesToProcess.Add($firstDataLine)

$maxToRead = if ($MaxRows -gt 0) { $MaxRows - 1 } else { [int]::MaxValue }
while (-not $reader.EndOfStream -and $linesToProcess.Count -lt $maxToRead + 1) {
    $linesToProcess.Add($reader.ReadLine())

    if ($linesToProcess.Count % 50000 -eq 0) {
        Write-Debug "読み込み中: $($linesToProcess.Count) 行..."
    }
}
$reader.Close()
Write-Debug "InputFile読み込み完了: $($linesToProcess.Count) 行"

# Excelファイル作成
$package = New-Object OfficeOpenXml.ExcelPackage
$sheet = $package.Workbook.Worksheets.Add("Sheet1")

if ($AddColumnNumbers) {
    # ヘッダー出力（1行目に列通番をつける）
    for ($j = 0; $j -lt $headers.Length; $j++) {
        $sheet.Cells.Item(1, $j + 1).Value = $headers[$j]
    }
}

# データ出力（ヘッダーの有無を考慮）
$rowIndex = if ($AddColumnNumbers) { 2 } else { 1 }

foreach ($line in $linesToProcess) {
    $columns = SplitCsvLine -line $line -Separator $Separator
    #Write-Debug "columns: $($columns) "
    for ($colIndex = 0; $colIndex -lt $headers.Length; $colIndex++) {
        $value = if ($colIndex -lt $columns.Count) { $columns[$colIndex] } else { $null }
        $sheet.Cells.Item($rowIndex, $colIndex + 1).Value = $value
    }

    $rowIndex++
    if ($rowIndex % 50000 -eq 0) {
        Write-Debug "書き出し中: $rowIndex 行目..."
    }
}
Write-Debug "書き出し完了: $($rowIndex -1)行"
# オートフィット・保存
Write-Debug "ファイル保存中..."
$sheet.Cells.AutoFitColumns()
try {
    $package.SaveAs([System.IO.FileInfo]::new($OutputFile))
    Write-Debug "Excelファイル出力完了: $OutputFile"
} catch {
    Write-Error "ファイル保存時にエラーが発生しました: $($_.Exception.Message)"
    Write-Error "出力ファイルが既に開かれている可能性があります。閉じて再実行してください。"
    exit 1
}