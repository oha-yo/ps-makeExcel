param(
    [Parameter(Mandatory = $true)][string]$InputFile,
    [Parameter()][string]$Encoding = "Shift_JIS",
    [Parameter()][int]$StartRow = 1,
    [Parameter()][int]$MaxRows = 0,
    [Parameter()][string]$Separator = ",",
    [Parameter()][bool]$AddColumnNumbers = $false
)
class CsvSplitter {
    [string]$Pattern
    CsvSplitter([string]$Separator) {
        # セパレータをエスケープしてリテラルで扱えるようにする（いろんな文字をセパレータで対応できるように）
        $escaped = [Regex]::Escape($Separator)
        # 区切り文字がクォートの外側にある場合のみ一致させるための正規表現
        # 以下のような場合３カラムに分けるため
        # 例)　みかん,"1,200円","愛媛"
        $this.Pattern = "$escaped(?=(?:[^""]*""[^""]*"")*[^""]*$)"
    }

    [string[]] Split([string]$line) {
        return [regex]::Split($line, $this.Pattern)
    }

    [string[]] SplitAndClean([string]$line) {
        $csvline = $this.Split($line)
        $cleaned = @()
        foreach ($item in $csvline) {
            <# エクセルに読み込む際に以下ロジックでクリーン
                Trim() → 前後の空白を削除
                -replace '^"(.*)"$', '$1' → 前後のダブルクォートを削除
                -replace '""', '"' → 二重クォートを単一のクォートに変換
            #>
            $cleaned += $item.Trim() -replace '^"(.*)"$', '$1' -replace '""', '"'
        }
        return $cleaned
    }
}

# EPPlus.dll の読み込み（ImportExcelモジュールから直接）
# $MyInvocation.MyCommand.Path は、現在実行中のスクリプトファイルのフルパスを返却
# Split-Path -Parent を使うことで、その親フォルダのパスを取得
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
# EPPlus.dllのフルパスを作成
$epplusPath = Join-Path $scriptDir ".\Modules\ImportExcel\7.8.10\EPPlus.dll"

# EPPlus.dllを PowerShell に読み込んで、（.NETクラス）を使えるようにする。
# PowerShellが.NETクラスを認識できるようにする
Add-Type -Path $epplusPath
# 以下行はコメントアウトしても動作するようだけど「保険」で記載しておくのがよさそう、、
[Reflection.Assembly]::LoadFrom($epplusPath) | Out-Null

# 入力ファイルチェック
if (-not (Test-Path $InputFile)) {
    Write-Error "Input file does not exist: $InputFile"
    exit 1
}
Write-Debug "Separator: $Separator"
# 出力ファイル名を自動生成  $OutputFile は Excel保存とログ出力に使用
$OutputFile = [System.IO.Path]::ChangeExtension($InputFile, "xlsx")

if (Test-Path $OutputFile) {
    try {
        $stream = [System.IO.File]::Open($OutputFile, 'Open', 'ReadWrite', 'None')
        $stream.Close()
    } catch {
        Write-Error "出力ファイルが使用中です: $OutputFile"
        exit 1
    }
}

try {
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
    #$columnCount = (SplitCsvLine -line $firstDataLine -Separator $Separator ).Count
    $splitter = [CsvSplitter]::new($Separator)
    $allColumns = $splitter.Split($firstDataLine)
    $columnCount = $allColumns.Count

    # 通番でヘッダー作成
    $headers = @()
    foreach ($i in 1..$columnCount) {
        $headers += "$i"
    }

    # データ読み込み用配列の作成
    $linesToProcess = [System.Collections.Generic.List[string]]::new()

    # データ行を List[string] で読み込み（1行目含む）
    $linesToProcess.Add($firstDataLine)
    
    $maxToRead = if ($MaxRows -gt 0) { $MaxRows - 1 } else { [int]::MaxValue }
    
    while (-not $reader.EndOfStream -and $linesToProcess.Count -lt $maxToRead + 1) {
        $linesToProcess.Add($reader.ReadLine())
        if ($linesToProcess.Count % 1000 -eq 0) {
            Write-Debug "読み込み中: $($linesToProcess.Count) 行..."
        }
    }
} finally {
    if ($reader) { $reader.Dispose() }
}
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
    #$columns = SplitCsvLine -line $line -Separator $Separator
    $columns = $splitter.SplitAndClean($line)
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
    Write-Host "エクセル出力完了: $OutputFile"
} catch {
    Write-Error "ファイル保存時にエラーが発生しました: $($_.Exception.Message)"
    Write-Error "出力ファイルが既に開かれている可能性があります。閉じて再実行してください。"
    exit 1
}