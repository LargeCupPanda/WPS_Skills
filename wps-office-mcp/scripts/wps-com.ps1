# Input: Action 名称与 JSON 参数
# Output: WPS COM 调用结果 JSON
# Pos: Windows COM 桥接脚本。一旦我被修改，请更新我的头部注释（Updated: 2026-02-05 13:09:38 CST），以及所属文件夹的md。
# WPS COM Bridge - PowerShell script for WPS COM operations
# Full implementation for Excel, Word, PPT, and common conversions
# Usage: powershell -File wps-com.ps1 -Action <action> -Params <json>

param(
    [string]$Action,
    [string]$Params = "{}"
)

$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ==================== COM Object Getters ====================

function Get-WpsExcel {
    try { return [System.Runtime.InteropServices.Marshal]::GetActiveObject('Ket.Application') }
    catch { return $null }
}

function Get-WpsWord {
    try { return [System.Runtime.InteropServices.Marshal]::GetActiveObject('Kwps.Application') }
    catch { return $null }
}

function Get-WpsPpt {
    try { return [System.Runtime.InteropServices.Marshal]::GetActiveObject('Kwpp.Application') }
    catch { return $null }
}

function Output-Json($obj) {
    $obj | ConvertTo-Json -Depth 10 -Compress
}

function Convert-ColumnLetterToNumber([string]$letters) {
    if (-not $letters) { return $null }
    $letters = $letters.ToUpper().Trim()
    $sum = 0
    foreach ($ch in $letters.ToCharArray()) {
        if ($ch -lt 'A' -or $ch -gt 'Z') { continue }
        $sum = ($sum * 26) + ([int][char]$ch - [int][char]'A' + 1)
    }
    return $sum
}

function Convert-ColumnNumberToLetter([int]$col) {
    if ($col -le 0) { return $null }
    $letter = ""
    while ($col -gt 0) {
        $mod = ($col - 1) % 26
        $letter = [char](65 + $mod) + $letter
        $col = [Math]::Floor(($col - 1) / 26)
    }
    return $letter
}

function Convert-HexColorToRgbInt([string]$hex) {
    if (-not $hex) { return $null }
    $hex = $hex.Trim()
    if ($hex.StartsWith('#')) { $hex = $hex.Substring(1) }
    if ($hex.Length -eq 3) {
        $hex = ($hex[0] + $hex[0] + $hex[1] + $hex[1] + $hex[2] + $hex[2])
    }
    if ($hex.Length -ne 6) { return $null }
    $r = [Convert]::ToInt32($hex.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($hex.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($hex.Substring(4, 2), 16)
    return $r + ($g * 256) + ($b * 65536)
}

function Get-RangeFromAddress($workbook, [string]$address) {
    if ($address -match "^(?<sheet>[^!]+)!(?<range>.+)$") {
        $sheetName = $matches.sheet.Trim("'")
        return $workbook.Sheets.Item($sheetName).Range($matches.range)
    }
    return $workbook.ActiveSheet.Range($address)
}

function Get-AppTypeByExtension([string]$filePath) {
    $ext = [System.IO.Path]::GetExtension($filePath).ToLower().Trim('.')
    $wordExts = @('doc', 'docx', 'docm', 'dot', 'dotx', 'dotm', 'rtf', 'txt', 'html', 'htm', 'xml', 'wps', 'wpt')
    $excelExts = @('xls', 'xlsx', 'xlsm', 'xlsb', 'xlt', 'xltx', 'xltm', 'csv', 'html', 'htm', 'et', 'ett')
    $pptExts = @('ppt', 'pptx', 'pptm', 'pot', 'potx', 'potm', 'pps', 'ppsx', 'ppsm', 'dps', 'dpt')
    if ($wordExts -contains $ext) { return 'word' }
    if ($excelExts -contains $ext) { return 'excel' }
    if ($pptExts -contains $ext) { return 'ppt' }
    return $null
}

function Get-WordSaveFormat([string]$format) {
    if (-not $format) { return $null }
    $f = $format.ToLower().Trim('.')
    # Word WdSaveFormat: doc=0, docx=16, pdf=17, rtf=6, xps=18, html=8, txt=2, xml=11
    $map = @{ doc = 0; docx = 16; pdf = 17; rtf = 6; xps = 18; html = 8; htm = 8; txt = 2; xml = 11 }
    return $map[$f]
}

function Get-ExcelFileFormat([string]$format) {
    if (-not $format) { return $null }
    $f = $format.ToLower().Trim('.')
    # Excel XlFileFormat: xls=-4143, xlsx=51, xlsm=52, xlsb=50, csv=6, html=44, xlt=17, xltx=54, xltm=53
    $map = @{ xls = -4143; xlsx = 51; xlsm = 52; xlsb = 50; csv = 6; html = 44; htm = 44; xlt = 17; xltx = 54; xltm = 53 }
    return $map[$f]
}

function Get-PptSaveFormat([string]$format) {
    if (-not $format) { return $null }
    $f = $format.ToLower().Trim('.')
    # PowerPoint PpSaveAsFileType: ppt=1, pptx=24, pptm=25, pdf=32, png=18, jpg=17, gif=16, bmp=19
    $map = @{ ppt = 1; pptx = 24; pptm = 25; pdf = 32; png = 18; jpg = 17; jpeg = 17; gif = 16; bmp = 19 }
    return $map[$f]
}

try { $p = $Params | ConvertFrom-Json } catch { $p = @{} }

switch ($Action) {

    # ==================== Common ====================
    "ping" {
        Output-Json @{ success = $true; data = @{ message = "pong"; timestamp = [DateTimeOffset]::Now.ToUnixTimeMilliseconds() } }
    }

    "wireCheck" {
        Output-Json @{ success = $true; data = @{ message = "WPS MCP Bridge 已连接" } }
    }

    "getAppInfo" {
        $excel = Get-WpsExcel
        if ($null -ne $excel) {
            $hasSelection = $false
            try { $hasSelection = ($null -ne $excel.Selection) } catch {}
            Output-Json @{ success = $true; data = @{ appType = "excel"; appName = $excel.Name; hasSelection = $hasSelection } }
            exit
        }
        $word = Get-WpsWord
        if ($null -ne $word) {
            $hasSelection = $false
            try { $hasSelection = ($null -ne $word.Selection) } catch {}
            Output-Json @{ success = $true; data = @{ appType = "word"; appName = $word.Name; hasSelection = $hasSelection } }
            exit
        }
        $ppt = Get-WpsPpt
        if ($null -ne $ppt) {
            $hasSelection = $false
            try { $hasSelection = ($null -ne $ppt.ActiveWindow.Selection) } catch {}
            Output-Json @{ success = $true; data = @{ appType = "ppt"; appName = $ppt.Name; hasSelection = $hasSelection } }
            exit
        }
        Output-Json @{ success = $false; error = "No WPS application running" }
    }

    "getSelectedText" {
        $word = Get-WpsWord
        if ($null -ne $word) {
            try {
                $text = $word.Selection.Text
                Output-Json @{ success = $true; data = @{ text = $text.Trim() } }
            } catch {
                Output-Json @{ success = $false; error = "Word selection not available" }
            }
            exit
        }
        $excel = Get-WpsExcel
        if ($null -ne $excel) {
            try {
                $sel = $excel.Selection
                $text = if ($sel -is [string]) { $sel } else { $sel.Text }
                Output-Json @{ success = $true; data = @{ text = [string]$text } }
            } catch {
                Output-Json @{ success = $false; error = "Excel selection not available" }
            }
            exit
        }
        $ppt = Get-WpsPpt
        if ($null -ne $ppt) {
            try {
                $sel = $ppt.ActiveWindow.Selection
                if ($sel -and $sel.TextRange) {
                    $text = $sel.TextRange.Text
                    Output-Json @{ success = $true; data = @{ text = [string]$text } }
                } else {
                    Output-Json @{ success = $false; error = "PPT selection has no text" }
                }
            } catch {
                Output-Json @{ success = $false; error = "PPT selection not available" }
            }
            exit
        }
        Output-Json @{ success = $false; error = "No WPS application running" }
    }

    "setSelectedText" {
        $text = if ($null -ne $p.text) { [string]$p.text } else { "" }
        $word = Get-WpsWord
        if ($null -ne $word) {
            try {
                $word.Selection.Text = $text
                Output-Json @{ success = $true }
            } catch {
                Output-Json @{ success = $false; error = "Word selection not available" }
            }
            exit
        }
        $excel = Get-WpsExcel
        if ($null -ne $excel) {
            try {
                $sel = $excel.Selection
                $sel.Value2 = $text
                Output-Json @{ success = $true }
            } catch {
                Output-Json @{ success = $false; error = "Excel selection not available" }
            }
            exit
        }
        $ppt = Get-WpsPpt
        if ($null -ne $ppt) {
            try {
                $sel = $ppt.ActiveWindow.Selection
                if ($sel -and $sel.TextRange) {
                    $sel.TextRange.Text = $text
                    Output-Json @{ success = $true }
                } else {
                    Output-Json @{ success = $false; error = "PPT selection has no text" }
                }
            } catch {
                Output-Json @{ success = $false; error = "PPT selection not available" }
            }
            exit
        }
        Output-Json @{ success = $false; error = "No WPS application running" }
    }

    "save" {
        $excel = Get-WpsExcel
        if ($null -ne $excel -and $null -ne $excel.ActiveWorkbook) {
            $excel.ActiveWorkbook.Save()
            Output-Json @{ success = $true; app = "excel" }
            exit
        }
        $word = Get-WpsWord
        if ($null -ne $word -and $null -ne $word.ActiveDocument) {
            $word.ActiveDocument.Save()
            Output-Json @{ success = $true; app = "word" }
            exit
        }
        $ppt = Get-WpsPpt
        if ($null -ne $ppt -and $null -ne $ppt.ActivePresentation) {
            $ppt.ActivePresentation.Save()
            Output-Json @{ success = $true; app = "ppt" }
            exit
        }
        Output-Json @{ success = $false; error = "No active document" }
    }

    "saveAs" {
        $path = $p.path
        if (-not $path) { Output-Json @{ success = $false; error = "Path required" }; exit }
        $appType = if ($p.appType) { $p.appType } else { Get-AppTypeByExtension $path }
        if ($appType -eq 'excel') {
            $excel = Get-WpsExcel
            if ($null -eq $excel -or $null -eq $excel.ActiveWorkbook) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
            $format = Get-ExcelFileFormat $p.format
            if ($null -ne $format) { $excel.ActiveWorkbook.SaveAs($path, $format) }
            else { $excel.ActiveWorkbook.SaveAs($path) }
            Output-Json @{ success = $true; data = @{ path = $path; appType = "excel" } }
            exit
        }
        if ($appType -eq 'word') {
            $word = Get-WpsWord
            if ($null -eq $word -or $null -eq $word.ActiveDocument) { Output-Json @{ success = $false; error = "No active document" }; exit }
            $format = Get-WordSaveFormat $p.format
            if ($null -ne $format) { $word.ActiveDocument.SaveAs($path, $format) }
            else { $word.ActiveDocument.SaveAs($path) }
            Output-Json @{ success = $true; data = @{ path = $path; appType = "word" } }
            exit
        }
        if ($appType -eq 'ppt') {
            $ppt = Get-WpsPpt
            if ($null -eq $ppt -or $null -eq $ppt.ActivePresentation) { Output-Json @{ success = $false; error = "No active presentation" }; exit }
            $format = Get-PptSaveFormat $p.format
            if ($null -ne $format) { $ppt.ActivePresentation.SaveAs($path, $format) }
            else { $ppt.ActivePresentation.SaveAs($path) }
            Output-Json @{ success = $true; data = @{ path = $path; appType = "ppt" } }
            exit
        }
        Output-Json @{ success = $false; error = "Unknown app type" }
    }

    "openFile" {
        $path = $p.path
        if (-not $path) { Output-Json @{ success = $false; error = "Path required" }; exit }
        $appType = if ($p.appType) { $p.appType } else { Get-AppTypeByExtension $path }
        if ($appType -eq 'excel') {
            $excel = Get-WpsExcel
            if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
            $excel.Workbooks.Open($path)
            Output-Json @{ success = $true; data = @{ path = $path; appType = "excel" } }
            exit
        }
        if ($appType -eq 'word') {
            $word = Get-WpsWord
            if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
            $word.Documents.Open($path)
            Output-Json @{ success = $true; data = @{ path = $path; appType = "word" } }
            exit
        }
        if ($appType -eq 'ppt') {
            $ppt = Get-WpsPpt
            if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
            $ppt.Presentations.Open($path)
            Output-Json @{ success = $true; data = @{ path = $path; appType = "ppt" } }
            exit
        }
        Output-Json @{ success = $false; error = "Unknown app type" }
    }

    "createDocument" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $word.Documents.Add() | Out-Null
        Output-Json @{ success = $true }
    }

    "createPresentation" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $ppt.Presentations.Add() | Out-Null
        Output-Json @{ success = $true }
    }

    "convertToPDF" {
        $excel = Get-WpsExcel
        if ($null -ne $excel -and $null -ne $excel.ActiveWorkbook) {
            $wb = $excel.ActiveWorkbook
            $sourcePath = $wb.FullName
            $outputPath = if ($p.outputPath) { $p.outputPath } else { [System.IO.Path]::ChangeExtension($sourcePath, 'pdf') }
            $wb.ExportAsFixedFormat(0, $outputPath)
            Output-Json @{ success = $true; data = @{ sourcePath = $sourcePath; outputPath = $outputPath; appType = "excel" } }
            exit
        }
        $word = Get-WpsWord
        if ($null -ne $word -and $null -ne $word.ActiveDocument) {
            $doc = $word.ActiveDocument
            $sourcePath = $doc.FullName
            $outputPath = if ($p.outputPath) { $p.outputPath } else { [System.IO.Path]::ChangeExtension($sourcePath, 'pdf') }
            $doc.ExportAsFixedFormat($outputPath, 17)
            Output-Json @{ success = $true; data = @{ sourcePath = $sourcePath; outputPath = $outputPath; appType = "word" } }
            exit
        }
        $ppt = Get-WpsPpt
        if ($null -ne $ppt -and $null -ne $ppt.ActivePresentation) {
            $pres = $ppt.ActivePresentation
            $sourcePath = $pres.FullName
            $outputPath = if ($p.outputPath) { $p.outputPath } else { [System.IO.Path]::ChangeExtension($sourcePath, 'pdf') }
            $pres.SaveAs($outputPath, 32)
            Output-Json @{ success = $true; data = @{ sourcePath = $sourcePath; outputPath = $outputPath; appType = "ppt" } }
            exit
        }
        Output-Json @{ success = $false; error = "No active document" }
    }

    "convertFormat" {
        $targetFormat = $p.targetFormat
        if (-not $targetFormat) { Output-Json @{ success = $false; error = "targetFormat required" }; exit }
        $excel = Get-WpsExcel
        if ($null -ne $excel -and $null -ne $excel.ActiveWorkbook) {
            $wb = $excel.ActiveWorkbook
            $sourcePath = $wb.FullName
            $outputPath = if ($p.outputPath) { $p.outputPath } else { [System.IO.Path]::ChangeExtension($sourcePath, $targetFormat) }
            $format = Get-ExcelFileFormat $targetFormat
            if ($null -ne $format) { $wb.SaveAs($outputPath, $format) } else { $wb.SaveAs($outputPath) }
            Output-Json @{ success = $true; data = @{ sourcePath = $sourcePath; outputPath = $outputPath; appType = "excel"; targetFormat = $targetFormat } }
            exit
        }
        $word = Get-WpsWord
        if ($null -ne $word -and $null -ne $word.ActiveDocument) {
            $doc = $word.ActiveDocument
            $sourcePath = $doc.FullName
            $outputPath = if ($p.outputPath) { $p.outputPath } else { [System.IO.Path]::ChangeExtension($sourcePath, $targetFormat) }
            $format = Get-WordSaveFormat $targetFormat
            if ($null -ne $format) { $doc.SaveAs($outputPath, $format) } else { $doc.SaveAs($outputPath) }
            Output-Json @{ success = $true; data = @{ sourcePath = $sourcePath; outputPath = $outputPath; appType = "word"; targetFormat = $targetFormat } }
            exit
        }
        $ppt = Get-WpsPpt
        if ($null -ne $ppt -and $null -ne $ppt.ActivePresentation) {
            $pres = $ppt.ActivePresentation
            $sourcePath = $pres.FullName
            $outputPath = if ($p.outputPath) { $p.outputPath } else { [System.IO.Path]::ChangeExtension($sourcePath, $targetFormat) }
            $format = Get-PptSaveFormat $targetFormat
            if ($null -ne $format) { $pres.SaveAs($outputPath, $format) } else { $pres.SaveAs($outputPath) }
            Output-Json @{ success = $true; data = @{ sourcePath = $sourcePath; outputPath = $outputPath; appType = "ppt"; targetFormat = $targetFormat } }
            exit
        }
        Output-Json @{ success = $false; error = "No active document" }
    }

    # ==================== Excel Basic ====================
    "getActiveWorkbook" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $sheets = @()
        for ($i = 1; $i -le $wb.Sheets.Count; $i++) { $sheets += $wb.Sheets.Item($i).Name }
        Output-Json @{ success = $true; data = @{ name = $wb.Name; path = $wb.FullName; sheetCount = $wb.Sheets.Count; sheets = $sheets } }
    }

    "getCellValue" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $sheet = if ($p.sheet -is [int]) { $wb.Sheets.Item($p.sheet) } else { $wb.Sheets.Item($p.sheet) }
        $cell = $sheet.Cells.Item($p.row, $p.col)
        Output-Json @{ success = $true; data = @{ value = $cell.Value2; text = $cell.Text; formula = $cell.Formula } }
    }

    "setCellValue" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        $sheet = if ($p.sheet -is [int]) { $wb.Sheets.Item($p.sheet) } else { $wb.Sheets.Item($p.sheet) }
        $sheet.Cells.Item($p.row, $p.col).Value2 = $p.value
        Output-Json @{ success = $true }
    }

    "getRangeData" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        $sheet = if ($p.sheet -is [int]) { $wb.Sheets.Item($p.sheet) } else { $wb.Sheets.Item($p.sheet) }
        $range = $sheet.Range($p.range)
        $data = @()
        for ($r = 1; $r -le $range.Rows.Count; $r++) {
            $row = @()
            for ($c = 1; $c -le $range.Columns.Count; $c++) { $row += $range.Cells.Item($r, $c).Value2 }
            $data += ,@($row)
        }
        Output-Json @{ success = $true; data = @{ data = $data } }
    }

    "setRangeData" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        $sheet = if ($p.sheet -is [int]) { $wb.Sheets.Item($p.sheet) } else { $wb.Sheets.Item($p.sheet) }
        $range = $sheet.Range($p.range)
        for ($r = 0; $r -lt $p.data.Count; $r++) {
            for ($c = 0; $c -lt $p.data[$r].Count; $c++) { $range.Cells.Item($r + 1, $c + 1).Value2 = $p.data[$r][$c] }
        }
        Output-Json @{ success = $true }
    }

    "diagnoseFormula" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $sheet = $excel.ActiveSheet
        $cell = $sheet.Range($p.cell)
        $value = $cell.Value2
        $formula = $cell.Formula
        $errorType = $null
        $diagnosis = ""
        $suggestion = ""
        $precedents = @()
        if ($value -is [string] -and $value.StartsWith('#')) {
            $errorType = $value
            switch ($errorType) {
                "#REF!" { $diagnosis = "引用了不存在的单元格或区域"; $suggestion = "检查引用区域是否被删除或移动" }
                "#N/A" { $diagnosis = "查找函数未找到匹配值"; $suggestion = "确认查找值存在，或检查匹配条件" }
                "#VALUE!" { $diagnosis = "参数类型不正确或运算类型不匹配"; $suggestion = "检查函数参数类型和引用单元格" }
                "#NAME?" { $diagnosis = "函数名或名称拼写错误"; $suggestion = "检查函数名是否正确" }
                "#DIV/0!" { $diagnosis = "除数为零"; $suggestion = "检查除数单元格，避免除以零" }
                "#NUM!" { $diagnosis = "数值无效或超出范围"; $suggestion = "检查函数参数范围" }
                "#NULL!" { $diagnosis = "交集为空"; $suggestion = "检查引用区域的交集是否存在" }
                default { $diagnosis = "未知错误"; $suggestion = "检查公式与引用" }
            }
        }
        try {
            $refs = $cell.DirectPrecedents
            if ($refs -ne $null) {
                foreach ($area in $refs.Areas) { $precedents += $area.Address() }
            }
        } catch {}
        Output-Json @{ success = $true; data = @{ cell = $p.cell; formula = $formula; currentValue = $value; errorType = $errorType; diagnosis = $diagnosis; suggestion = $suggestion; precedents = $precedents } }
    }

    "cleanData" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $sheet = if ($p.sheet -is [int]) { $wb.Sheets.Item($p.sheet) } else { $wb.Sheets.Item($p.sheet) }
        $range = $sheet.Range($p.range)
        $opsResult = @()
        foreach ($op in $p.operations) {
            $success = $true
            $message = ""
            switch ($op) {
                "trim" {
                    foreach ($cell in $range) {
                        if ($cell.Value2 -is [string]) { $cell.Value2 = $cell.Value2.Trim() }
                    }
                    $message = "已去除前后空格"
                }
                "remove_duplicates" {
                    $colCount = $range.Columns.Count
                    $cols = @()
                    for ($i = 1; $i -le $colCount; $i++) { $cols += $i }
                    $range.RemoveDuplicates($cols, 1)
                    $message = "已删除重复行"
                }
                "unify_date" {
                    foreach ($cell in $range) {
                        try {
                            if ($cell.Value2) {
                                $dt = [DateTime]::FromOADate($cell.Value2)
                                $cell.Value2 = $dt.ToString("yyyy-MM-dd")
                            }
                        } catch {}
                    }
                    $message = "已统一日期格式"
                }
                "remove_empty_rows" {
                    for ($i = $range.Rows.Count; $i -ge 1; $i--) {
                        $row = $range.Rows.Item($i)
                        $isEmpty = $true
                        foreach ($cell in $row.Cells) { if ($cell.Value2) { $isEmpty = $false; break } }
                        if ($isEmpty) { $row.Delete() }
                    }
                    $message = "已删除空行"
                }
                default { $success = $false; $message = "不支持的操作" }
            }
            $opsResult += @{ operation = $op; success = $success; message = $message }
        }
        Output-Json @{ success = $true; data = @{ range = $p.range; operations = $opsResult; message = "cleanData completed" } }
    }

    "createPivotTable" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $sourceRange = Get-RangeFromAddress $wb $p.sourceRange
        $destSheet = if ($p.destinationSheet) { $wb.Sheets.Item($p.destinationSheet) } else { $excel.ActiveSheet }
        $destCell = $destSheet.Range($p.destinationCell)
        $pivotName = if ($p.tableName) { $p.tableName } else { "PivotTable" + [Guid]::NewGuid().ToString("N").Substring(0, 6) }
        $cache = $wb.PivotCaches().Create(1, $sourceRange)
        $table = $cache.CreatePivotTable($destCell, $pivotName)
        foreach ($fieldName in $p.rowFields) {
            $field = $table.PivotFields($fieldName)
            $field.Orientation = 1
            $field.Position = 1
        }
        if ($p.columnFields) {
            foreach ($fieldName in $p.columnFields) {
                $field = $table.PivotFields($fieldName)
                $field.Orientation = 2
                $field.Position = 1
            }
        }
        if ($p.filterFields) {
            foreach ($fieldName in $p.filterFields) {
                $field = $table.PivotFields($fieldName)
                $field.Orientation = 3
                $field.Position = 1
            }
        }
        foreach ($vf in $p.valueFields) {
            $field = $table.PivotFields($vf.field)
            $funcMap = @{ SUM = -4157; COUNT = -4112; AVERAGE = -4106; MAX = -4136; MIN = -4139 }
            $func = $funcMap[$vf.aggregation]
            if ($null -eq $func) { $func = -4157 }
            $table.AddDataField($field, $vf.field, $func) | Out-Null
        }
        Output-Json @{ success = $true; data = @{ pivotTableName = $pivotName; location = $destCell.Address(); rowCount = $table.RowRange.Rows.Count; columnCount = $table.TableRange1.Columns.Count } }
    }

    "updatePivotTable" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $sheet = $excel.ActiveSheet
        $table = $null
        if ($p.pivotTableName) {
            try { $table = $sheet.PivotTables($p.pivotTableName) } catch {}
        }
        if ($null -eq $table -and $p.pivotTableCell) {
            try { $table = $sheet.Range($p.pivotTableCell).PivotTable } catch {}
        }
        if ($null -eq $table) { Output-Json @{ success = $false; error = "PivotTable not found" }; exit }
        $ops = @()
        if ($p.addRowFields) {
            foreach ($f in $p.addRowFields) {
                $field = $table.PivotFields($f); $field.Orientation = 1; $field.Position = 1
                $ops += @{ operation = "addRowFields"; success = $true; message = $f }
            }
        }
        if ($p.removeRowFields) {
            foreach ($f in $p.removeRowFields) {
                $field = $table.PivotFields($f); $field.Orientation = 0
                $ops += @{ operation = "removeRowFields"; success = $true; message = $f }
            }
        }
        if ($p.addColumnFields) {
            foreach ($f in $p.addColumnFields) {
                $field = $table.PivotFields($f); $field.Orientation = 2; $field.Position = 1
                $ops += @{ operation = "addColumnFields"; success = $true; message = $f }
            }
        }
        if ($p.removeColumnFields) {
            foreach ($f in $p.removeColumnFields) {
                $field = $table.PivotFields($f); $field.Orientation = 0
                $ops += @{ operation = "removeColumnFields"; success = $true; message = $f }
            }
        }
        if ($p.addFilterFields) {
            foreach ($f in $p.addFilterFields) {
                $field = $table.PivotFields($f); $field.Orientation = 3; $field.Position = 1
                $ops += @{ operation = "addFilterFields"; success = $true; message = $f }
            }
        }
        if ($p.removeFilterFields) {
            foreach ($f in $p.removeFilterFields) {
                $field = $table.PivotFields($f); $field.Orientation = 0
                $ops += @{ operation = "removeFilterFields"; success = $true; message = $f }
            }
        }
        if ($p.addValueFields) {
            foreach ($vf in $p.addValueFields) {
                $field = $table.PivotFields($vf.field)
                $funcMap = @{ SUM = -4157; COUNT = -4112; AVERAGE = -4106; MAX = -4136; MIN = -4139 }
                $func = $funcMap[$vf.aggregation]
                if ($null -eq $func) { $func = -4157 }
                $table.AddDataField($field, $vf.field, $func) | Out-Null
                $ops += @{ operation = "addValueFields"; success = $true; message = $vf.field }
            }
        }
        if ($p.updateValueFields) {
            foreach ($vf in $p.updateValueFields) {
                $field = $table.PivotFields($vf.field)
                $funcMap = @{ SUM = -4157; COUNT = -4112; AVERAGE = -4106; MAX = -4136; MIN = -4139 }
                $func = $funcMap[$vf.aggregation]
                if ($null -eq $func) { $func = -4157 }
                $field.Function = $func
                $ops += @{ operation = "updateValueFields"; success = $true; message = $vf.field }
            }
        }
        if ($p.removeValueFields) {
            foreach ($vf in $p.removeValueFields) {
                try { $table.DataFields($vf).Orientation = 0 } catch {}
                $ops += @{ operation = "removeValueFields"; success = $true; message = $vf }
            }
        }
        if ($p.refresh) { try { $table.RefreshTable() | Out-Null; $ops += @{ operation = "refresh"; success = $true; message = "refreshed" } } catch {} }
        Output-Json @{ success = $true; data = @{ pivotTableName = $table.Name; operations = $ops } }
    }

    "setFormula" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        $sheet = if ($p.sheet -is [int]) { $wb.Sheets.Item($p.sheet) } else { $wb.Sheets.Item($p.sheet) }
        $range = if ($p.range) { $sheet.Range($p.range) } else { $sheet.Cells.Item($p.row, $p.col) }
        $range.Formula = $p.formula
        Output-Json @{ success = $true }
    }

    "setCellFormat" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        if ($p.numberFormat) { $range.NumberFormat = $p.numberFormat }
        Output-Json @{ success = $true; data = @{ range = $p.range; format = $p.numberFormat } }
    }

    "setCellStyle" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        if ($p.fontSize) { $range.Font.Size = $p.fontSize }
        if ($null -ne $p.bold) { $range.Font.Bold = [bool]$p.bold }
        if ($null -ne $p.italic) { $range.Font.Italic = [bool]$p.italic }
        if ($p.fontName) { $range.Font.Name = $p.fontName }
        if ($p.backgroundColor) {
            $bg = Convert-HexColorToRgbInt([string]$p.backgroundColor)
            if ($null -ne $bg) { $range.Interior.Color = $bg }
        }
        if ($p.fontColor) {
            $fc = Convert-HexColorToRgbInt([string]$p.fontColor)
            if ($null -ne $fc) { $range.Font.Color = $fc }
        }
        if ($p.horizontalAlignment) {
            $hAlignMap = @{ left = -4131; center = -4108; right = -4152 }
            $hAlign = $hAlignMap[$p.horizontalAlignment]
            if ($null -ne $hAlign) { $range.HorizontalAlignment = $hAlign }
        }
        if ($p.verticalAlignment) {
            $vAlignMap = @{ top = -4160; center = -4108; bottom = -4107 }
            $vAlign = $vAlignMap[$p.verticalAlignment]
            if ($null -ne $vAlign) { $range.VerticalAlignment = $vAlign }
        }
        if ($null -ne $p.border -and $p.border) {
            $range.Borders.LineStyle = 1
            if ($p.borderColor) {
                $bc = Convert-HexColorToRgbInt([string]$p.borderColor)
                if ($null -ne $bc) { $range.Borders.Color = $bc }
            }
        }
        Output-Json @{ success = $true; data = @{ range = $p.range } }
    }

    "setBorder" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $styleMap = @{ thin = 1; medium = 2; thick = 4; double = 6; none = 0 }
        $style = $styleMap[$p.style]
        if ($null -eq $style) { $style = 1 }
        $position = if ($p.position) { $p.position } else { "all" }
        $borders = @()
        if ($position -eq "all" -or $position -eq "outside") { $borders += 7, 8, 9, 10 }
        if ($position -eq "all" -or $position -eq "inside") { $borders += 11, 12 }
        if ($position -eq "left") { $borders += 7 }
        if ($position -eq "top") { $borders += 8 }
        if ($position -eq "bottom") { $borders += 9 }
        if ($position -eq "right") { $borders += 10 }
        $colorValue = $null
        if ($p.color) { $colorValue = Convert-HexColorToRgbInt([string]$p.color) }
        foreach ($b in $borders) {
            $border = $range.Borders.Item($b)
            if ($style -eq 0) {
                $border.LineStyle = -4142
            } else {
                $border.LineStyle = 1
                $border.Weight = $style
            }
            if ($null -ne $colorValue) { $border.Color = $colorValue }
        }
        Output-Json @{ success = $true; data = @{ range = $p.range; style = $p.style; position = $position } }
    }

    "copyFormat" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $sourceRange = $sheet.Range($p.source)
        $targetRange = $sheet.Range($p.target)
        $sourceRange.Copy()
        $targetRange.PasteSpecial(-4122)
        $excel.CutCopyMode = $false
        Output-Json @{ success = $true; data = @{ source = $p.source; target = $p.target } }
    }

    "clearFormats" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $range.ClearFormats()
        Output-Json @{ success = $true; data = @{ range = $p.range } }
    }

    "addConditionalFormat" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $formatType = if ($p.type) { $p.type } else { "cellValue" }
        if ($formatType -eq "cellValue") {
            $operatorMap = @{ greater = 5; greaterThan = 5; less = 6; lessThan = 6; equal = 3; notEqual = 4; greaterEqual = 7; greaterThanOrEqual = 7; lessEqual = 8; lessThanOrEqual = 8; between = 1 }
            $op = $operatorMap[$p.operator]
            if ($null -eq $op) { $op = 3 }
            $val1 = if ($null -ne $p.value1) { $p.value1 } else { $p.value }
            $val2 = $p.value2
            $cf = $range.FormatConditions.Add(1, $op, $val1, $val2)
            if ($null -ne $cf -and $p.backgroundColor) {
                $bg = Convert-HexColorToRgbInt([string]$p.backgroundColor)
                if ($null -ne $bg) { $cf.Interior.Color = $bg }
            }
            if ($null -ne $cf -and $p.fontColor) {
                $fc = Convert-HexColorToRgbInt([string]$p.fontColor)
                if ($null -ne $fc) { $cf.Font.Color = $fc }
            }
        } elseif ($formatType -eq "colorScale") {
            $scaleType = if ($p.colorScaleType) { $p.colorScaleType } else { 3 }
            $range.FormatConditions.AddColorScale($scaleType)
        } elseif ($formatType -eq "dataBar") {
            $range.FormatConditions.AddDatabar() | Out-Null
        }
        Output-Json @{ success = $true; data = @{ range = $p.range; type = $formatType } }
    }

    "removeConditionalFormat" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        if ($p.index) {
            $range.FormatConditions.Item([int]$p.index).Delete()
        } else {
            $range.FormatConditions.Delete()
        }
        Output-Json @{ success = $true; data = @{ range = $p.range } }
    }

    "getConditionalFormats" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $formats = @()
        $count = $range.FormatConditions.Count
        for ($i = 1; $i -le $count; $i++) {
            $cf = $range.FormatConditions.Item($i)
            $formats += @{ index = $i; type = $cf.Type }
        }
        Output-Json @{ success = $true; data = @{ range = $p.range; formats = $formats; count = $count } }
    }

    "addDataValidation" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $typeMap = @{ list = 3; whole = 1; decimal = 2; date = 4; time = 5; textLength = 6; custom = 7 }
        $validationType = $typeMap[$p.validationType]
        if ($null -eq $validationType) { $validationType = 3 }
        $range.Validation.Delete()
        if ($p.validationType -eq "list") {
            $listFormula = if ($p.formula1) { $p.formula1 } elseif ($p.list) { ($p.list -join ",") } else { "" }
            $range.Validation.Add($validationType, 1, 1, $listFormula)
            if ($null -ne $p.showDropdown -and -not [bool]$p.showDropdown) {
                $range.Validation.InCellDropdown = $false
            } else {
                $range.Validation.InCellDropdown = $true
            }
        } else {
            $operatorMap = @{ between = 1; notBetween = 2; equal = 3; notEqual = 4; greater = 5; less = 6; greaterEqual = 7; lessEqual = 8 }
            $op = $operatorMap[$p.operator]
            if ($null -eq $op) { $op = 1 }
            $range.Validation.Add($validationType, 1, $op, $p.formula1, $p.formula2)
        }
        if ($p.inputTitle -or $p.inputMessage) {
            $range.Validation.InputTitle = if ($p.inputTitle) { $p.inputTitle } else { "" }
            $range.Validation.InputMessage = if ($p.inputMessage) { $p.inputMessage } else { "" }
        }
        if ($p.errorTitle -or $p.errorMessage) {
            $range.Validation.ErrorTitle = if ($p.errorTitle) { $p.errorTitle } else { "" }
            $range.Validation.ErrorMessage = if ($p.errorMessage) { $p.errorMessage } else { "" }
        }
        Output-Json @{ success = $true; data = @{ range = $p.range; type = $p.validationType } }
    }

    "removeDataValidation" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $range.Validation.Delete()
        Output-Json @{ success = $true; data = @{ range = $p.range } }
    }

    "getDataValidations" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $validation = $range.Validation
        Output-Json @{ success = $true; data = @{ range = $p.range; type = $validation.Type; formula1 = $validation.Formula1; formula2 = $validation.Formula2; inputTitle = $validation.InputTitle; inputMessage = $validation.InputMessage } }
    }

    "getFormula" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $cell = $sheet.Range($p.cell)
        $formula = if ($cell.Formula) { $cell.Formula } else { "" }
        $formulaLocal = ""
        try { $formulaLocal = $cell.FormulaLocal } catch { $formulaLocal = "" }
        Output-Json @{ success = $true; data = @{ cell = $p.cell; formula = $formula; formulaLocal = $formulaLocal; hasFormula = ($formula -like "=*") } }
    }

    "setArrayFormula" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $range.FormulaArray = $p.formula
        Output-Json @{ success = $true; data = @{ range = $p.range; formula = $p.formula } }
    }

    "getCellInfo" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $cell = $sheet.Range($p.cell)
        $value = $cell.Value2
        $formula = if ($cell.Formula) { $cell.Formula } else { "" }
        $numberFormat = if ($cell.NumberFormat) { $cell.NumberFormat } else { "" }
        $fontName = if ($cell.Font.Name) { $cell.Font.Name } else { "" }
        $fontSize = if ($cell.Font.Size) { $cell.Font.Size } else { 0 }
        $bold = if ($null -ne $cell.Font.Bold) { [bool]$cell.Font.Bold } else { $false }
        $bgColor = if ($null -ne $cell.Interior.Color) { $cell.Interior.Color } else { 0 }
        Output-Json @{ success = $true; data = @{ cell = $p.cell; value = $value; formula = $formula; numberFormat = $numberFormat; font = @{ name = $fontName; size = $fontSize; bold = $bold }; backgroundColor = $bgColor } }
    }

    "refreshLinks" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $links = $wb.LinkSources(1)
        if ($links) {
            for ($i = 1; $i -le $links.Length; $i++) {
                $wb.UpdateLink($links[$i - 1], 1)
            }
            Output-Json @{ success = $true; data = @{ refreshed = $links.Length } }
        } else {
            Output-Json @{ success = $true; data = @{ refreshed = 0; message = "没有外部链接" } }
        }
    }

    "consolidate" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $destRange = $sheet.Range($p.destination)
        $funcMap = @{ sum = 9; count = 2; average = 1; max = 4; min = 5 }
        $func = $funcMap[$p.function]
        if ($null -eq $func) { $func = 9 }
        $destRange.Consolidate($p.sources, $func, $p.topRow, $p.leftColumn, $p.createLinks)
        Output-Json @{ success = $true; data = @{ destination = $p.destination; sources = $p.sources } }
    }

    "calculateSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        if ($p.all) {
            $excel.Calculate()
            Output-Json @{ success = $true; data = @{ calculated = "all" } }
        } else {
            $sheet = $excel.ActiveSheet
            $sheet.Calculate()
            Output-Json @{ success = $true; data = @{ calculated = $sheet.Name } }
        }
    }

    # ==================== Excel Advanced ====================
    "getExcelContext" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $sheet = $excel.ActiveSheet
        $usedRange = $sheet.UsedRange
        $headers = @()
        if ($usedRange.Rows.Count -gt 0) {
            $headerRow = $usedRange.Rows.Item(1)
            for ($i = 1; $i -le [Math]::Min($headerRow.Columns.Count, 26); $i++) {
                $headers += @{ column = [char](64 + $i); value = $headerRow.Cells.Item(1, $i).Value2 }
            }
        }
        $sheets = @(); for ($i = 1; $i -le $wb.Sheets.Count; $i++) { $sheets += $wb.Sheets.Item($i).Name }
        Output-Json @{ success = $true; data = @{
            workbookName = $wb.Name; currentSheet = $sheet.Name; allSheets = $sheets
            selectedCell = $excel.Selection.Address(); headers = $headers; usedRange = $usedRange.Address(); usedRangeAddress = $usedRange.Address()
        }}
    }

    "getContext" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $sheet = $excel.ActiveSheet
        $usedRange = $sheet.UsedRange
        $headers = @()
        if ($usedRange.Rows.Count -gt 0) {
            $headerRow = $usedRange.Rows.Item(1)
            for ($i = 1; $i -le [Math]::Min($headerRow.Columns.Count, 26); $i++) {
                $headers += @{ column = [char](64 + $i); value = $headerRow.Cells.Item(1, $i).Value2 }
            }
        }
        $sheets = @(); for ($i = 1; $i -le $wb.Sheets.Count; $i++) { $sheets += $wb.Sheets.Item($i).Name }
        Output-Json @{ success = $true; data = @{
            workbookName = $wb.Name; currentSheet = $sheet.Name; allSheets = $sheets
            selectedCell = $excel.Selection.Address(); headers = $headers; usedRangeAddress = $usedRange.Address(); usedRange = $usedRange.Address()
        }}
    }

    "sortRange" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $keyCol = $sheet.Range($p.keyColumn)
        $order = if ($p.order -eq "desc") { 2 } else { 1 }
        $range.Sort($keyCol, $order)
        Output-Json @{ success = $true }
    }

    "autoFilter" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        if ($p.criteria) {
            $range.AutoFilter($p.field, $p.criteria)
        } else {
            $range.AutoFilter()
        }
        Output-Json @{ success = $true }
    }

    "createChart" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.dataRange)
        $chartType = $p.chartType
        if ($null -eq $chartType) {
            $chartTypes = @{ column = 51; column_clustered = 51; column_stacked = 52; bar = 57; bar_clustered = 57; line = 4; line_markers = 65; pie = 5; doughnut = -4120; area = 1; scatter = -4169; radar = -4151 }
            $chartType = $chartTypes[$p.chartTypeName]
        }
        if ($null -eq $chartType) { $chartType = 51 }
        $chartTypeName = if ($p.chartTypeName) { $p.chartTypeName } else { $null }
        $left = if ($p.position -and $p.position.left) { $p.position.left } elseif ($p.left) { $p.left } else { $range.Left + $range.Width + 20 }
        $top = if ($p.position -and $p.position.top) { $p.position.top } elseif ($p.top) { $p.top } else { $range.Top }
        $width = if ($p.position -and $p.position.width) { $p.position.width } else { 400 }
        $height = if ($p.position -and $p.position.height) { $p.position.height } else { 300 }
        $chartObj = $sheet.ChartObjects().Add($left, $top, $width, $height)
        $chartObj.Chart.SetSourceData($range)
        $chartObj.Chart.ChartType = $chartType
        if ($p.title) { $chartObj.Chart.HasTitle = $true; $chartObj.Chart.ChartTitle.Text = $p.title }
        if ($null -ne $p.showLegend) { $chartObj.Chart.HasLegend = [bool]$p.showLegend }
        if ($p.showDataLabels) { $chartObj.Chart.ApplyDataLabels() }
        Output-Json @{ success = $true; data = @{ chartName = $chartObj.Name; chartIndex = $chartObj.Index; dataRange = $p.dataRange; chartType = $chartTypeName; position = @{ left = $left; top = $top; width = $width; height = $height } } }
    }

    "updateChart" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = if ($p.sheet) { $excel.ActiveWorkbook.Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        $chartObj = $null
        if ($p.chartName) {
            try { $chartObj = $sheet.ChartObjects($p.chartName) } catch {}
        }
        if ($null -eq $chartObj -and $p.chartIndex) {
            try { $chartObj = $sheet.ChartObjects().Item([int]$p.chartIndex) } catch {}
        }
        if ($null -eq $chartObj) { Output-Json @{ success = $false; error = "Chart not found" }; exit }
        $updated = @()
        if ($null -ne $p.title) {
            if ($p.title -eq "") { $chartObj.Chart.HasTitle = $false }
            else { $chartObj.Chart.HasTitle = $true; $chartObj.Chart.ChartTitle.Text = $p.title }
            $updated += "title"
        }
        if ($null -ne $p.chartType) {
            $chartObj.Chart.ChartType = $p.chartType
            $updated += "chartType"
        }
        if ($null -ne $p.showLegend) {
            $chartObj.Chart.HasLegend = [bool]$p.showLegend
            $updated += "showLegend"
        }
        if ($null -ne $p.legendPosition) {
            $legendMap = @{ bottom = -4107; top = -4160; left = -4131; right = -4152 }
            $pos = $legendMap[$p.legendPosition]
            if ($pos) { $chartObj.Chart.Legend.Position = $pos; $updated += "legendPosition" }
        }
        if ($null -ne $p.showDataLabels) {
            if ($p.showDataLabels) {
                $chartObj.Chart.ApplyDataLabels()
            } else {
                try {
                    $series = $chartObj.Chart.SeriesCollection()
                    for ($i = 1; $i -le $series.Count; $i++) {
                        try { $series.Item($i).DataLabels().Delete() } catch {}
                    }
                } catch {}
            }
            $updated += "showDataLabels"
        }
        if ($p.dataRange) {
            $chartObj.Chart.SetSourceData($sheet.Range($p.dataRange))
            $updated += "dataRange"
        }
        if ($p.colors -and $p.colors.Count -gt 0) {
            $series = $chartObj.Chart.SeriesCollection()
            for ($i = 1; $i -le $series.Count -and $i -le $p.colors.Count; $i++) {
                $color = Convert-HexColorToRgbInt($p.colors[$i - 1])
                if ($null -ne $color) { $series.Item($i).Format.Fill.ForeColor.RGB = $color }
            }
            $updated += "colors"
        }
        Output-Json @{ success = $true; data = @{ chartName = $chartObj.Name; updatedProperties = $updated } }
    }

    "removeDuplicates" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $originalCount = $range.Rows.Count
        $cols = @()
        if ($p.columns -and $p.columns.Count -gt 0) {
            foreach ($col in $p.columns) {
                if ($col -is [int]) { $cols += $col }
                else {
                    $num = Convert-ColumnLetterToNumber([string]$col)
                    if ($num) { $cols += $num }
                }
            }
        }
        if ($cols.Count -eq 0) { $cols = @(1) }
        $hasHeader = if ($null -ne $p.hasHeader) { [int]([bool]$p.hasHeader) } else { 1 }
        $range.RemoveDuplicates($cols, $hasHeader)
        $remainingCount = $range.Rows.Count
        $removedCount = $originalCount - $remainingCount
        Output-Json @{ success = $true; data = @{ originalCount = $originalCount; removedCount = $removedCount; remainingCount = $remainingCount } }
    }

    "createSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $sheet = $wb.Sheets.Add()
        if ($p.name) { $sheet.Name = $p.name }
        Output-Json @{ success = $true; data = @{ sheetName = $sheet.Name; sheetIndex = $sheet.Index } }
    }

    "deleteSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        $name = $sheet.Name
        $excel.DisplayAlerts = $false
        $sheet.Delete()
        $excel.DisplayAlerts = $true
        Output-Json @{ success = $true; data = @{ deletedSheet = $name } }
    }

    "renameSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        $oldName = $sheet.Name
        $sheet.Name = $p.newName
        Output-Json @{ success = $true; data = @{ oldName = $oldName; newName = $p.newName } }
    }

    "copySheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        if ($p.before) {
            $sheet.Copy($wb.Sheets.Item($p.before))
        } elseif ($p.after) {
            $sheet.Copy($null, $wb.Sheets.Item($p.after))
        } else {
            $sheet.Copy($null, $wb.Sheets.Item($wb.Sheets.Count))
        }
        Output-Json @{ success = $true; data = @{ copiedFrom = $sheet.Name } }
    }

    "getSheetList" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $sheets = @()
        for ($i = 1; $i -le $wb.Sheets.Count; $i++) {
            $sheet = $wb.Sheets.Item($i)
            $sheets += @{ name = $sheet.Name; index = $i; visible = $sheet.Visible }
        }
        Output-Json @{ success = $true; data = @{ sheets = $sheets; count = $sheets.Count; activeSheet = $excel.ActiveSheet.Name } }
    }

    "switchSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $sheet = $wb.Sheets.Item($p.sheet)
        $sheet.Activate()
        Output-Json @{ success = $true; data = @{ activeSheet = $sheet.Name } }
    }

    "moveSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        if ($p.before) {
            $sheet.Move($wb.Sheets.Item($p.before))
        } elseif ($p.after) {
            $sheet.Move($null, $wb.Sheets.Item($p.after))
        }
        Output-Json @{ success = $true; data = @{ movedSheet = $sheet.Name } }
    }

    "mergeCells" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $across = if ($null -ne $p.across) { [bool]$p.across } else { $false }
        $range.Merge($across)
        Output-Json @{ success = $true; data = @{ range = $p.range } }
    }

    "unmergeCells" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $range.UnMerge()
        Output-Json @{ success = $true; data = @{ range = $p.range } }
    }

    "setColumnWidth" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $col = $p.column
        if ($col -is [int]) { $col = Convert-ColumnNumberToLetter([int]$col) }
        $sheet.Range("$col:$col").ColumnWidth = $p.width
        Output-Json @{ success = $true; data = @{ column = $col; width = $p.width } }
    }

    "setRowHeight" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $sheet.Range("$($p.row):$($p.row)").RowHeight = $p.height
        Output-Json @{ success = $true; data = @{ row = $p.row; height = $p.height } }
    }

    "autoFitColumn" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        if ($p.range) {
            $sheet.Range($p.range).Columns.AutoFit()
        } elseif ($p.column) {
            $col = $p.column
            if ($col -is [int]) { $col = Convert-ColumnNumberToLetter([int]$col) }
            $sheet.Range("$col:$col").AutoFit()
        } else {
            $sheet.UsedRange.Columns.AutoFit()
        }
        Output-Json @{ success = $true; data = @{ message = "列宽已自动调整" } }
    }

    "autoFitRow" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        if ($p.range) {
            $sheet.Range($p.range).Rows.AutoFit()
        } elseif ($p.row) {
            $sheet.Range("$($p.row):$($p.row)").AutoFit()
        } else {
            $sheet.UsedRange.Rows.AutoFit()
        }
        Output-Json @{ success = $true; data = @{ message = "行高已自动调整" } }
    }

    "autoFitAll" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $range = if ($p.range) { $sheet.Range($p.range) } else { $sheet.UsedRange }
        $range.Columns.AutoFit()
        $range.Rows.AutoFit()
        Output-Json @{ success = $true; data = @{ message = "列宽行高已自动调整" } }
    }

    "setNumberFormat" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        if ($p.format) { $range.NumberFormat = $p.format }
        Output-Json @{ success = $true; data = @{ range = $p.range; format = $p.format } }
    }

    "wrapText" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $wrap = if ($null -ne $p.wrap) { [bool]$p.wrap } else { $true }
        $range.WrapText = $wrap
        Output-Json @{ success = $true; data = @{ range = $p.range; wrapText = $wrap } }
    }

    "setPrintArea" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        if ($p.range) { $sheet.PageSetup.PrintArea = $p.range } else { $sheet.PageSetup.PrintArea = "" }
        Output-Json @{ success = $true; data = @{ printArea = if ($p.range) { $p.range } else { "cleared" } } }
    }

    "getSelection" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sel = $excel.Selection
        if ($null -eq $sel) { Output-Json @{ success = $false; error = "No selection" }; exit }
        $addr = $sel.Address()
        Output-Json @{ success = $true; data = @{ address = $addr; rows = $sel.Rows.Count; columns = $sel.Columns.Count } }
    }

    "clearRange" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $clearType = if ($p.type) { $p.type } else { "all" }
        if ($clearType -eq "contents") { $range.ClearContents() }
        elseif ($clearType -eq "formats") { $range.ClearFormats() }
        elseif ($clearType -eq "comments") { $range.ClearComments() }
        else { $range.Clear() }
        Output-Json @{ success = $true; data = @{ range = $p.range; clearType = $clearType } }
    }

    "insertRows" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $startRow = if ($p.row) { [int]$p.row } elseif ($p.startRow) { [int]$p.startRow } else { $null }
        if ($null -eq $startRow) { Output-Json @{ success = $false; error = "row/startRow required" }; exit }
        $count = if ($p.count) { [int]$p.count } else { 1 }
        $endRow = $startRow + $count - 1
        $sheet.Range("$startRow:$endRow").Insert()
        Output-Json @{ success = $true; data = @{ insertedAt = $startRow; count = $count } }
    }

    "insertColumns" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $col = if ($p.column) { $p.column } elseif ($p.startColumn) { $p.startColumn } else { $null }
        if ($null -eq $col) { Output-Json @{ success = $false; error = "column/startColumn required" }; exit }
        if ($col -is [int]) { $col = Convert-ColumnNumberToLetter([int]$col) }
        $count = if ($p.count) { [int]$p.count } else { 1 }
        for ($i = 0; $i -lt $count; $i++) {
            $sheet.Range("$col:$col").Insert()
        }
        Output-Json @{ success = $true; data = @{ insertedAt = $col; count = $count } }
    }

    "deleteRows" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $startRow = if ($p.row) { [int]$p.row } elseif ($p.startRow) { [int]$p.startRow } else { $null }
        if ($null -eq $startRow) { Output-Json @{ success = $false; error = "row/startRow required" }; exit }
        $count = if ($p.count) { [int]$p.count } else { 1 }
        $endRow = $startRow + $count - 1
        $sheet.Range("$startRow:$endRow").Delete()
        Output-Json @{ success = $true; data = @{ deletedFrom = $startRow; count = $count } }
    }

    "deleteColumns" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $col = if ($p.column) { $p.column } elseif ($p.startColumn) { $p.startColumn } else { $null }
        if ($null -eq $col) { Output-Json @{ success = $false; error = "column/startColumn required" }; exit }
        if ($col -is [int]) { $col = Convert-ColumnNumberToLetter([int]$col) }
        $count = if ($p.count) { [int]$p.count } else { 1 }
        for ($i = 0; $i -lt $count; $i++) {
            $sheet.Range("$col:$col").Delete()
        }
        Output-Json @{ success = $true; data = @{ deletedFrom = $col; count = $count } }
    }

    "hideRows" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $rows = if ($p.rows) { $p.rows } elseif ($p.row) { @($p.row) } else { @() }
        if ($rows.Count -eq 0) { Output-Json @{ success = $false; error = "row/rows required" }; exit }
        foreach ($r in $rows) { $sheet.Range("$r:$r").Hidden = $true }
        Output-Json @{ success = $true; data = @{ hiddenRows = $rows } }
    }

    "hideColumns" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $cols = if ($p.columns) { $p.columns } elseif ($p.column) { @($p.column) } else { @() }
        if ($cols.Count -eq 0) { Output-Json @{ success = $false; error = "column/columns required" }; exit }
        $hidden = @()
        foreach ($c in $cols) {
            $col = $c
            if ($col -is [int]) { $col = Convert-ColumnNumberToLetter([int]$col) }
            $sheet.Range("$col:$col").Hidden = $true
            $hidden += $col
        }
        Output-Json @{ success = $true; data = @{ hiddenColumns = $hidden } }
    }

    "showRows" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $rows = if ($p.rows) { $p.rows } elseif ($p.row) { @($p.row) } else { @() }
        if ($rows.Count -eq 0) { Output-Json @{ success = $false; error = "row/rows required" }; exit }
        foreach ($r in $rows) { $sheet.Range("$r:$r").Hidden = $false }
        Output-Json @{ success = $true; data = @{ shownRows = $rows } }
    }

    "showColumns" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $cols = if ($p.columns) { $p.columns } elseif ($p.column) { @($p.column) } else { @() }
        if ($cols.Count -eq 0) { Output-Json @{ success = $false; error = "column/columns required" }; exit }
        $shown = @()
        foreach ($c in $cols) {
            $col = $c
            if ($col -is [int]) { $col = Convert-ColumnNumberToLetter([int]$col) }
            $sheet.Range("$col:$col").Hidden = $false
            $shown += $col
        }
        Output-Json @{ success = $true; data = @{ shownColumns = $shown } }
    }

    "groupRows" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        if (-not $p.startRow -or -not $p.endRow) { Output-Json @{ success = $false; error = "startRow/endRow required" }; exit }
        $sheet.Range("$($p.startRow):$($p.endRow)").Group()
        Output-Json @{ success = $true; data = @{ grouped = "$($p.startRow):$($p.endRow)" } }
    }

    "groupColumns" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        if (-not $p.startColumn -or -not $p.endColumn) { Output-Json @{ success = $false; error = "startColumn/endColumn required" }; exit }
        $startCol = $p.startColumn
        $endCol = $p.endColumn
        if ($startCol -is [int]) { $startCol = Convert-ColumnNumberToLetter([int]$startCol) }
        if ($endCol -is [int]) { $endCol = Convert-ColumnNumberToLetter([int]$endCol) }
        $sheet.Range("$startCol:$endCol").Group()
        Output-Json @{ success = $true; data = @{ grouped = "$startCol:$endCol" } }
    }

    "freezePanes" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        if ($p.cell) {
            $sheet.Range($p.cell).Select()
        } elseif ($p.row -and $p.column) {
            $col = $p.column
            if ($col -is [int]) { $col = Convert-ColumnNumberToLetter([int]$col) }
            $sheet.Range("$col$($p.row)").Select()
        }
        $excel.ActiveWindow.FreezePanes = $true
        Output-Json @{ success = $true; data = @{ message = "窗格已冻结" } }
    }

    "unfreezePanes" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $excel.ActiveWindow.FreezePanes = $false
        Output-Json @{ success = $true; data = @{ message = "窗格冻结已取消" } }
    }

    "findInSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $searchRange = if ($p.range) { $sheet.Range($p.range) } else { $sheet.UsedRange }
        $results = @()
        $lookAt = if ($p.matchCase) { 1 } else { 2 }
        $found = $searchRange.Find($p.searchText, $null, -4163, $lookAt)
        if ($found) {
            $firstAddr = $found.Address()
            do {
                $addr = $found.Address()
                $results += @{ address = ($addr -replace "\$", ""); value = $found.Value2 }
                $found = $searchRange.FindNext($found)
                if ($null -eq $found) { break }
                $currAddr = $found.Address()
            } while ($currAddr -ne $firstAddr)
        }
        Output-Json @{ success = $true; data = @{ searchText = $p.searchText; results = $results; count = $results.Count } }
    }

    "replaceInSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $searchRange = if ($p.range) { $sheet.Range($p.range) } else { $sheet.UsedRange }
        $lookAt = if ($p.matchCase) { 1 } else { 2 }
        $replaced = $searchRange.Replace($p.searchText, $p.replaceText, $lookAt)
        Output-Json @{ success = $true; data = @{ searchText = $p.searchText; replaceText = $p.replaceText; success = $replaced } }
    }

    "copyRange" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $range.Copy()
        Output-Json @{ success = $true; data = @{ range = $p.range; message = "已复制到剪贴板" } }
    }

    "pasteRange" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $destRange = $sheet.Range($p.destination)
        if ($p.pasteType -eq "values") {
            $destRange.PasteSpecial(-4163)
        } elseif ($p.pasteType -eq "formats") {
            $destRange.PasteSpecial(-4122)
        } elseif ($p.pasteType -eq "formulas") {
            $destRange.PasteSpecial(-4123)
        } else {
            $sheet.Paste($destRange)
        }
        Output-Json @{ success = $true; data = @{ destination = $p.destination } }
    }

    "fillSeries" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $startCell = $range.Cells.Item(1, 1)
        $startCell.Value2 = if ($null -ne $p.startValue) { $p.startValue } else { 1 }
        $typeMap = @{ linear = 0; growth = 1; date = 2; autoFill = 3 }
        $fillType = $typeMap[$p.type]
        if ($null -eq $fillType) { $fillType = 0 }
        $step = if ($null -ne $p.step) { $p.step } else { 1 }
        $range.DataSeries($null, -4132, $fillType, $step)
        Output-Json @{ success = $true; data = @{ range = $p.range; type = $p.type } }
    }

    "transpose" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $sourceRange = $sheet.Range($p.sourceRange)
        $destCell = if ($p.destinationCell) { $p.destinationCell } elseif ($p.targetCell) { $p.targetCell } else { $null }
        if ($null -eq $destCell) { Output-Json @{ success = $false; error = "destinationCell/targetCell required" }; exit }
        $destRange = $sheet.Range($destCell)
        $sourceRange.Copy()
        $destRange.PasteSpecial(-4163, -4142, $false, $true)
        $excel.CutCopyMode = $false
        Output-Json @{ success = $true; data = @{ source = $p.sourceRange; destination = $destCell } }
    }

    "textToColumns" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $delimiter = if ($p.delimiter) { $p.delimiter } else { "," }
        $tab = $false; $semicolon = $false; $comma = $false; $space = $false; $other = $false; $otherChar = $null
        if ($delimiter -eq "`t") { $tab = $true }
        elseif ($delimiter -eq ";") { $semicolon = $true }
        elseif ($delimiter -eq ",") { $comma = $true }
        elseif ($delimiter -eq " ") { $space = $true }
        else { $other = $true; $otherChar = $delimiter }
        $range.TextToColumns($null, 1, 1, $false, $tab, $semicolon, $comma, $space, $other, $otherChar)
        Output-Json @{ success = $true; data = @{ range = $p.range; delimiter = $delimiter } }
    }

    "subtotal" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $funcMap = @{ sum = 9; count = 2; average = 1; max = 4; min = 5 }
        $func = $funcMap[$p.function]
        if ($null -eq $func) { $func = 9 }
        $totalCols = $p.totalColumns
        if ($null -eq $totalCols) { $totalCols = $p.totalColumn }
        if ($totalCols -isnot [System.Array]) { $totalCols = @($totalCols) }
        $replace = if ($null -ne $p.replace) { [bool]$p.replace } else { $true }
        $range.Subtotal([int]$p.groupBy, $func, $totalCols, $replace, $false, $true)
        Output-Json @{ success = $true; data = @{ range = $p.range } }
    }

    "createNamedRange" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $addr = $range.Address()
        $wb.Names.Add($p.name, "=" + $sheet.Name + "!" + $addr)
        Output-Json @{ success = $true; data = @{ name = $p.name; range = $p.range } }
    }

    "deleteNamedRange" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $wb.Names.Item($p.name).Delete()
        Output-Json @{ success = $true; data = @{ deletedName = $p.name } }
    }

    "getNamedRanges" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $names = @()
        for ($i = 1; $i -le $wb.Names.Count; $i++) {
            $n = $wb.Names.Item($i)
            $names += @{ name = $n.Name; refersTo = $n.RefersTo; visible = $n.Visible }
        }
        Output-Json @{ success = $true; data = @{ names = $names; count = $names.Count } }
    }

    "addCellComment" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $cell = $sheet.Range($p.cell)
        if ($cell.Comment) { $cell.Comment.Delete() }
        $cell.AddComment($p.text)
        if ($p.visible) { $cell.Comment.Visible = $true }
        Output-Json @{ success = $true; data = @{ cell = $p.cell; text = $p.text } }
    }

    "deleteCellComment" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $cell = $sheet.Range($p.cell)
        if ($cell.Comment) { $cell.Comment.Delete() }
        Output-Json @{ success = $true; data = @{ cell = $p.cell } }
    }

    "getCellComments" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $comments = @()
        for ($i = 1; $i -le $sheet.Comments.Count; $i++) {
            $c = $sheet.Comments.Item($i)
            $addr = $c.Parent.Address()
            $comments += @{ cell = ($addr -replace "\$", ""); text = $c.Text(); author = if ($c.Author) { $c.Author } else { "" } }
        }
        Output-Json @{ success = $true; data = @{ comments = $comments; count = $comments.Count } }
    }

    "protectSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = if ($p.sheet) { $excel.ActiveWorkbook.Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        $password = if ($p.password) { $p.password } else { "" }
        $sheet.Protect($password, $p.drawingObjects, $p.contents, $p.scenarios)
        Output-Json @{ success = $true; data = @{ sheet = $sheet.Name; protected = $true } }
    }

    "unprotectSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = if ($p.sheet) { $excel.ActiveWorkbook.Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        $password = if ($p.password) { $p.password } else { "" }
        $sheet.Unprotect($password)
        Output-Json @{ success = $true; data = @{ sheet = $sheet.Name; protected = $false } }
    }

    "protectWorkbook" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; exit }
        $password = if ($p.password) { $p.password } else { "" }
        $structure = if ($null -ne $p.structure) { [bool]$p.structure } else { $true }
        $wb.Protect($password, $structure, $p.windows)
        Output-Json @{ success = $true; data = @{ workbook = $wb.Name; protected = $true } }
    }

    "insertExcelImage" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $left = if ($null -ne $p.left) { $p.left } else { 100 }
        $top = if ($null -ne $p.top) { $p.top } else { 100 }
        $width = if ($null -ne $p.width) { $p.width } else { -1 }
        $height = if ($null -ne $p.height) { $p.height } else { -1 }
        $pic = $sheet.Shapes.AddPicture($p.path, $false, $true, $left, $top, $width, $height)
        Output-Json @{ success = $true; data = @{ name = $pic.Name; path = $p.path } }
    }

    "setHyperlink" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.cell)
        $address = if ($p.address) { $p.address } else { "" }
        $subAddress = if ($p.subAddress) { $p.subAddress } else { "" }
        $screenTip = if ($p.screenTip) { $p.screenTip } else { "" }
        $textToDisplay = if ($p.textToDisplay) { $p.textToDisplay } else { "" }
        $sheet.Hyperlinks.Add($range, $address, $subAddress, $screenTip, $textToDisplay)
        Output-Json @{ success = $true; data = @{ cell = $p.cell; address = $address } }
    }

    "lockCells" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $locked = if ($null -ne $p.locked) { [bool]$p.locked } else { $true }
        $range.Locked = $locked
        Output-Json @{ success = $true; data = @{ range = $p.range; locked = $locked } }
    }

    "openWorkbook" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.Workbooks.Open($p.path, $p.updateLinks, $p.readOnly)
        Output-Json @{ success = $true; data = @{ name = $wb.Name; path = $wb.FullName; sheets = $wb.Sheets.Count } }
    }

    "getOpenWorkbooks" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $workbooks = @()
        for ($i = 1; $i -le $excel.Workbooks.Count; $i++) {
            $wb = $excel.Workbooks.Item($i)
            $workbooks += @{ name = $wb.Name; path = $wb.FullName; sheets = $wb.Sheets.Count; active = ($wb.Name -eq $excel.ActiveWorkbook.Name) }
        }
        Output-Json @{ success = $true; data = @{ workbooks = $workbooks; count = $workbooks.Count } }
    }

    "switchWorkbook" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.Workbooks.Item($(if ($p.name) { $p.name } else { $p.index }))
        $wb.Activate()
        Output-Json @{ success = $true; data = @{ name = $wb.Name; path = $wb.FullName } }
    }

    "closeWorkbook" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = if ($p.name) { $excel.Workbooks.Item($p.name) } else { $excel.ActiveWorkbook }
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No workbook" }; exit }
        $name = $wb.Name
        $saveChanges = if ($null -ne $p.saveChanges) { [bool]$p.saveChanges } else { $true }
        $wb.Close($saveChanges)
        Output-Json @{ success = $true; data = @{ closed = $name } }
    }

    "createWorkbook" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.Workbooks.Add()
        if ($p.name) { $wb.SaveAs($p.name) }
        Output-Json @{ success = $true; data = @{ name = $wb.Name; path = $wb.FullName; sheets = $wb.Sheets.Count } }
    }

    # ==================== Word ====================
    "getActiveDocument" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        Output-Json @{ success = $true; data = @{
            name = $doc.Name; path = $doc.FullName
            paragraphCount = $doc.Paragraphs.Count; wordCount = $doc.Words.Count; characterCount = $doc.Characters.Count
        }}
    }

    "generateTOC" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        $position = if ($p.position) { $p.position } else { "start" }
        $levels = if ($p.levels) { [int]$p.levels } else { 3 }
        $includePageNumbers = if ($null -ne $p.includePageNumbers) { [bool]$p.includePageNumbers } else { $true }
        $range = if ($position -eq "cursor") { $word.Selection.Range } else { $doc.Range(0, 0) }
        $doc.TablesOfContents.Add($range, $true, 1, $levels, $false, "", $true, $true, $includePageNumbers, $true)
        Output-Json @{ success = $true; data = @{ levels = $levels; position = $position; includePageNumbers = $includePageNumbers } }
    }

    "getDocumentText" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        $text = $doc.Content.Text
        if ($text.Length -gt 10000) { $text = $text.Substring(0, 10000) + "...(truncated)" }
        Output-Json @{ success = $true; data = @{ text = $text; length = $doc.Content.Text.Length } }
    }

    "getOpenDocuments" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $documents = @()
        for ($i = 1; $i -le $word.Documents.Count; $i++) {
            $docItem = $word.Documents.Item($i)
            $documents += @{ name = $docItem.Name; path = $docItem.FullName; paragraphs = $docItem.Paragraphs.Count; active = ($word.ActiveDocument.Name -eq $docItem.Name) }
        }
        Output-Json @{ success = $true; data = @{ documents = $documents; count = $documents.Count } }
    }

    "switchDocument" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $docItem = $word.Documents.Item($(if ($p.name) { $p.name } else { $p.index }))
        $docItem.Activate()
        Output-Json @{ success = $true; data = @{ name = $docItem.Name; path = $docItem.FullName } }
    }

    "openDocument" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        if (-not $p.path) { Output-Json @{ success = $false; error = "path required" }; exit }
        $docItem = $word.Documents.Open($p.path)
        Output-Json @{ success = $true; data = @{ name = $docItem.Name; path = $docItem.FullName; paragraphs = $docItem.Paragraphs.Count } }
    }

    "closeDocument" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $docItem = if ($p.name) { $word.Documents.Item($p.name) } else { $word.ActiveDocument }
        if ($null -eq $docItem) { Output-Json @{ success = $false; error = "No document" }; exit }
        $name = $docItem.Name
        $saveChanges = if ($null -ne $p.saveChanges) { [bool]$p.saveChanges } else { $true }
        $docItem.Close($saveChanges)
        Output-Json @{ success = $true; data = @{ closed = $name } }
    }

    "insertText" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        $position = if ($p.position) { $p.position } else { "cursor" }
        switch ($position) {
            "start" { $range = $doc.Range(0, 0); $range.InsertBefore($p.text) }
            "end" { $range = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1); $range.InsertAfter($p.text) }
            default { $word.Selection.TypeText($p.text) }
        }
        if ($p.style) {
            try { $word.Selection.Range.Style = $p.style } catch {}
        }
        Output-Json @{ success = $true; data = @{ position = $position; textLength = $p.text.Length } }
    }

    "setFont" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        $range = if ($p.range -eq "all") { $doc.Content } else { $word.Selection.Range }
        if ($p.fontName) { $range.Font.Name = $p.fontName }
        if ($p.fontSize) { $range.Font.Size = $p.fontSize }
        if ($null -ne $p.bold) { $range.Font.Bold = $p.bold }
        if ($null -ne $p.italic) { $range.Font.Italic = $p.italic }
        if ($null -ne $p.underline) { $range.Font.Underline = $p.underline }
        if ($p.color) {
            $colorMap = @{ red = 255; blue = 16711680; green = 65280; black = 0 }
            $colorValue = $colorMap[$p.color.ToLower()]
            if ($null -eq $colorValue) { $colorValue = Convert-HexColorToRgbInt([string]$p.color) }
            if ($null -ne $colorValue) { $range.Font.Color = $colorValue }
        }
        Output-Json @{ success = $true; data = @{ settings = @{ fontName = $p.fontName; fontSize = $p.fontSize; bold = $p.bold; italic = $p.italic; underline = $p.underline; color = $p.color } } }
    }

    "findReplace" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        $find = $doc.Content.Find
        $find.ClearFormatting()
        $find.Replacement.ClearFormatting()
        $find.Text = $p.findText
        $find.Replacement.Text = $p.replaceText
        $matchCase = if ($null -ne $p.matchCase) { [bool]$p.matchCase } else { $false }
        $matchWholeWord = if ($null -ne $p.matchWholeWord) { [bool]$p.matchWholeWord } else { $false }
        $replaceAll = if ($null -ne $p.replaceAll) { [bool]$p.replaceAll } else { $true }
        $replaceType = if ($replaceAll) { 2 } else { 1 }
        $result = $find.Execute($p.findText, $matchCase, $matchWholeWord, $false, $false, $false, $true, 1, $false, $p.replaceText, $replaceType)
        Output-Json @{ success = $true; data = @{ replaced = $result } }
    }

    "insertTable" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        $rows = if ($p.rows) { [int]$p.rows } else { 3 }
        $cols = if ($p.cols) { [int]$p.cols } else { 3 }
        $range = $word.Selection.Range
        $table = $doc.Tables.Add($range, $rows, $cols)
        if ($p.data) {
            for ($r = 0; $r -lt [Math]::Min($p.data.Count, $rows); $r++) {
                for ($c = 0; $c -lt [Math]::Min($p.data[$r].Count, $cols); $c++) {
                    $table.Cell($r + 1, $c + 1).Range.Text = [string]$p.data[$r][$c]
                }
            }
        }
        $table.Borders.Enable = $true
        Output-Json @{ success = $true }
    }

    "setParagraph" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        $range = if ($p.range -eq "all") { $doc.Content } else { $word.Selection.Range }
        $para = $range.ParagraphFormat
        if ($null -ne $p.alignment) {
            $alignMap = @{ left = 0; center = 1; right = 2; justify = 3 }
            $align = $alignMap[$p.alignment]
            if ($null -eq $align) { $align = 0 }
            $para.Alignment = $align
        }
        if ($p.lineSpacing) {
            $para.LineSpacingRule = 4
            $para.LineSpacing = [double]$p.lineSpacing * 12
        }
        if ($null -ne $p.spaceBefore) { $para.SpaceBefore = $p.spaceBefore }
        if ($null -ne $p.spaceAfter) { $para.SpaceAfter = $p.spaceAfter }
        if ($null -ne $p.firstLineIndent) { $para.FirstLineIndent = [double]$p.firstLineIndent * 28.35 }
        if ($null -ne $p.leftIndent) { $para.LeftIndent = [double]$p.leftIndent * 28.35 }
        if ($null -ne $p.rightIndent) { $para.RightIndent = [double]$p.rightIndent * 28.35 }
        Output-Json @{ success = $true }
    }

    "setPageSetup" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        $ps = $doc.PageSetup
        if ($null -ne $p.topMargin) { $ps.TopMargin = [double]$p.topMargin * 28.35 }
        if ($null -ne $p.bottomMargin) { $ps.BottomMargin = [double]$p.bottomMargin * 28.35 }
        if ($null -ne $p.leftMargin) { $ps.LeftMargin = [double]$p.leftMargin * 28.35 }
        if ($null -ne $p.rightMargin) { $ps.RightMargin = [double]$p.rightMargin * 28.35 }
        if ($null -ne $p.orientation) {
            $ps.Orientation = if ($p.orientation -eq "landscape") { 1 } else { 0 }
        }
        if ($null -ne $p.paperSize) {
            $sizeMap = @{ A4 = 7; A3 = 6; Letter = 1; Legal = 5 }
            $size = $sizeMap[$p.paperSize]
            if ($null -eq $size) { $size = 7 }
            $ps.PaperSize = $size
        }
        Output-Json @{ success = $true }
    }

    "insertPageBreak" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        $breakType = if ($p.type) { $p.type } else { "page" }
        $breakTypeMap = @{ page = 7; column = 8; section = 2; sectionContinuous = 3 }
        $bt = $breakTypeMap[$breakType]
        if ($null -eq $bt) { $bt = 7 }
        $word.Selection.InsertBreak($bt)
        Output-Json @{ success = $true; data = @{ type = $breakType } }
    }

    "insertHeader" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        $section = $doc.Sections.Item(1)
        $header = $section.Headers.Item(1)
        $header.Range.Text = if ($p.text) { $p.text } else { "" }
        if ($p.alignment) {
            $alignMap = @{ left = 0; center = 1; right = 2 }
            $header.Range.ParagraphFormat.Alignment = $alignMap[$p.alignment]
        }
        Output-Json @{ success = $true }
    }

    "insertFooter" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        $section = $doc.Sections.Item(1)
        $footer = $section.Footers.Item(1)
        $footer.Range.Text = if ($p.text) { $p.text } else { "" }
        if ($p.alignment) {
            $alignMap = @{ left = 0; center = 1; right = 2 }
            $footer.Range.ParagraphFormat.Alignment = $alignMap[$p.alignment]
        }
        if ($p.includePageNumber) {
            $footer.Range.InsertAfter(" - 第 ")
            $footer.Range.Fields.Add($footer.Range, -1, "PAGE", $false)
            $footer.Range.InsertAfter(" 页 ")
        }
        Output-Json @{ success = $true }
    }

    "insertHyperlink" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        $range = $word.Selection.Range
        $url = if ($p.url) { $p.url } else { $p.address }
        $text = if ($p.text) { $p.text } elseif ($p.displayText) { $p.displayText } else { $url }
        if ($range.Text -and $range.Text.Trim() -ne "") {
            $doc.Hyperlinks.Add($range, $url)
        } else {
            $range.Text = $text
            $doc.Hyperlinks.Add($doc.Range($range.Start, $range.Start + $text.Length), $url)
        }
        Output-Json @{ success = $true; data = @{ url = $url; text = $text } }
    }

    "insertBookmark" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if (-not $p.name) { Output-Json @{ success = $false; error = "name required" }; exit }
        $range = $word.Selection.Range
        $doc.Bookmarks.Add($p.name, $range)
        Output-Json @{ success = $true; data = @{ name = $p.name } }
    }

    "getBookmarks" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        $bookmarks = @()
        for ($i = 1; $i -le $doc.Bookmarks.Count; $i++) {
            $bm = $doc.Bookmarks.Item($i)
            $bookmarks += @{ name = $bm.Name; start = $bm.Start; end = $bm.End }
        }
        Output-Json @{ success = $true; data = @{ bookmarks = $bookmarks; count = $bookmarks.Count } }
    }

    "addComment" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        $text = if ($p.text) { $p.text } else { $p.comment }
        if (-not $text) { Output-Json @{ success = $false; error = "text required" }; exit }
        $range = $word.Selection.Range
        $doc.Comments.Add($range, $text)
        Output-Json @{ success = $true }
    }

    "getComments" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        $comments = @()
        for ($i = 1; $i -le $doc.Comments.Count; $i++) {
            $c = $doc.Comments.Item($i)
            $comments += @{ index = $i; text = $c.Range.Text; author = if ($c.Author) { $c.Author } else { "" }; date = if ($c.Date) { $c.Date.ToString() } else { "" } }
        }
        Output-Json @{ success = $true; data = @{ comments = $comments; count = $comments.Count } }
    }

    "getDocumentStats" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        $stats = @{ name = $doc.Name; path = $doc.FullName; pages = $doc.ComputeStatistics(2); words = $doc.ComputeStatistics(0); characters = $doc.ComputeStatistics(3); paragraphs = $doc.ComputeStatistics(4); lines = $doc.ComputeStatistics(1) }
        Output-Json @{ success = $true; data = $stats }
    }

    "insertImage" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        $path = if ($p.path) { $p.path } else { $p.filePath }
        if (-not $path) { Output-Json @{ success = $false; error = "path required" }; exit }
        $range = $word.Selection.Range
        $shape = $doc.InlineShapes.AddPicture($path, $false, $true, $range)
        if ($p.width) { $shape.Width = $p.width }
        if ($p.height) { $shape.Height = $p.height }
        if ($p.scale) { $shape.ScaleWidth = $p.scale; $shape.ScaleHeight = $p.scale }
        Output-Json @{ success = $true; data = @{ width = $shape.Width; height = $shape.Height } }
    }

    "applyStyle" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        if ($p.range -and $null -ne $p.range.start -and $null -ne $p.range.end) {
            $range = $doc.Range([int]$p.range.start, [int]$p.range.end)
        } else {
            $range = $word.Selection.Range
        }
        $range.Style = $p.styleName
        Output-Json @{ success = $true; data = @{ affectedText = $range.Text } }
    }

    # ==================== PPT ====================
    "getActivePresentation" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; exit }
        $slides = @()
        for ($i = 1; $i -le $pres.Slides.Count; $i++) {
            $slide = $pres.Slides.Item($i)
            $shapes = @()
            for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
                $shape = $slide.Shapes.Item($j)
                $text = ""
                try { if ($shape.HasTextFrame -and $shape.TextFrame.HasText) { $text = $shape.TextFrame.TextRange.Text.Substring(0, [Math]::Min(50, $shape.TextFrame.TextRange.Text.Length)) } } catch {}
                $shapes += @{ name = $shape.Name; type = $shape.Type; text = $text }
            }
            $slides += @{ index = $i; shapeCount = $slide.Shapes.Count; shapes = $shapes }
        }
        Output-Json @{ success = $true; data = @{ name = $pres.Name; path = $pres.FullName; slideCount = $pres.Slides.Count; slides = $slides } }
    }

    "openPresentation" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        if (-not $p.path) { Output-Json @{ success = $false; error = "path required" }; exit }
        $pres = $ppt.Presentations.Open($p.path)
        Output-Json @{ success = $true; data = @{ name = $pres.Name; path = $pres.FullName; slideCount = $pres.Slides.Count } }
    }

    "closePresentation" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = if ($p.name) { $ppt.Presentations.Item($p.name) } else { $ppt.ActivePresentation }
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No presentation" }; exit }
        $name = $pres.Name
        $pres.Close()
        Output-Json @{ success = $true; data = @{ closed = $name } }
    }

    "getOpenPresentations" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $presentations = @()
        for ($i = 1; $i -le $ppt.Presentations.Count; $i++) {
            $presItem = $ppt.Presentations.Item($i)
            $presentations += @{ name = $presItem.Name; path = $presItem.FullName; slideCount = $presItem.Slides.Count }
        }
        Output-Json @{ success = $true; data = @{ presentations = $presentations; count = $presentations.Count } }
    }

    "switchPresentation" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $presItem = $ppt.Presentations.Item($(if ($p.name) { $p.name } else { $p.index }))
        $presItem.Windows.Item(1).Activate()
        Output-Json @{ success = $true; data = @{ name = $presItem.Name } }
    }

    "deleteSlide" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; exit }
        $index = if ($p.index) { $p.index } else { $p.slideIndex }
        $pres.Slides.Item($index).Delete()
        Output-Json @{ success = $true; data = @{ deleted = $index } }
    }

    "duplicateSlide" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; exit }
        $index = if ($p.index) { $p.index } else { $p.slideIndex }
        if (-not $index) { $index = 1 }
        $slide = $pres.Slides.Item($index)
        $newSlide = $slide.Duplicate()
        $newIndex = $newSlide.Item(1).SlideIndex
        Output-Json @{ success = $true; data = @{ sourceIndex = $index; newIndex = $newIndex } }
    }

    "moveSlide" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; exit }
        $fromIndex = if ($p.from) { $p.from } else { $p.fromIndex }
        $toIndex = if ($p.to) { $p.to } else { $p.toIndex }
        $pres.Slides.Item($fromIndex).MoveTo($toIndex)
        Output-Json @{ success = $true; data = @{ from = $fromIndex; to = $toIndex } }
    }

    "getSlideCount" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; exit }
        Output-Json @{ success = $true; data = @{ count = $pres.Slides.Count } }
    }

    "getSlideInfo" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; exit }
        $index = if ($p.index) { $p.index } else { $p.slideIndex }
        if (-not $index) { $index = 1 }
        $slide = $pres.Slides.Item($index)
        $shapes = @()
        for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
            $shape = $slide.Shapes.Item($i)
            $shapes += @{ name = $shape.Name; type = $shape.Type; hasText = if ($shape.HasTextFrame) { $true } else { $false } }
        }
        Output-Json @{ success = $true; data = @{ index = $index; shapeCount = $slide.Shapes.Count; shapes = $shapes; layout = $slide.Layout } }
    }

    "switchSlide" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; exit }
        $index = if ($p.index) { $p.index } else { $p.slideIndex }
        $ppt.ActiveWindow.View.GotoSlide($index)
        Output-Json @{ success = $true; data = @{ currentSlide = $index } }
    }

    "setSlideLayout" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; exit }
        $index = if ($p.index) { $p.index } else { $p.slideIndex }
        if (-not $index) { $index = 1 }
        $layoutMap = @{ title = 1; titleContent = 2; blank = 12; twoColumn = 4; comparison = 5; titleOnly = 11 }
        $layout = $layoutMap[$p.layout]
        if ($null -eq $layout) { $layout = if ($p.layout) { $p.layout } else { 2 } }
        $pres.Slides.Item($index).Layout = $layout
        Output-Json @{ success = $true; data = @{ slideIndex = $index; layout = $layout } }
    }

    "getSlideNotes" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; exit }
        $index = if ($p.index) { $p.index } else { $p.slideIndex }
        if (-not $index) { $index = 1 }
        $slide = $pres.Slides.Item($index)
        $notes = ""
        try { $notes = $slide.NotesPage.Shapes.Item(2).TextFrame.TextRange.Text } catch { $notes = "" }
        Output-Json @{ success = $true; data = @{ slideIndex = $index; notes = $notes } }
    }

    "setSlideNotes" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; exit }
        $index = if ($p.index) { $p.index } else { $p.slideIndex }
        if (-not $index) { $index = 1 }
        $slide = $pres.Slides.Item($index)
        $slide.NotesPage.Shapes.Item(2).TextFrame.TextRange.Text = if ($p.notes) { $p.notes } else { "" }
        Output-Json @{ success = $true; data = @{ slideIndex = $index } }
    }

    "slide.add" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $layouts = @{ title = 1; title_content = 2; blank = 12; two_column = 3; comparison = 34 }
        $layoutKey = if ($p.layout) { $p.layout } else { "title_content" }
        $layoutType = $layouts[$layoutKey]
        if ($null -eq $layoutType) { $layoutType = 2 }
        $position = if ($p.position) { $p.position } else { $pres.Slides.Count + 1 }
        $slide = $pres.Slides.Add($position, $layoutType)
        if ($p.title -and $slide.Shapes.HasTitle) { $slide.Shapes.Title.TextFrame.TextRange.Text = $p.title }
        if ($p.content) {
            try {
                foreach ($shape in $slide.Shapes) {
                    try {
                        if ($shape.PlaceholderFormat.Type -eq 2) {
                            $shape.TextFrame.TextRange.Text = $p.content
                        }
                    } catch {}
                }
            } catch {}
        }
        Output-Json @{ success = $true; data = @{ slideIndex = $position; layout = $layoutKey } }
    }

    "addSlide" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $layouts = @{ title = 1; title_content = 2; blank = 12; two_column = 3; comparison = 34 }
        $layoutKey = if ($p.layout) { $p.layout } else { "title_content" }
        $layoutType = $layouts[$layoutKey]
        if ($null -eq $layoutType) { $layoutType = 2 }
        $position = if ($p.position) { $p.position } else { $pres.Slides.Count + 1 }
        $slide = $pres.Slides.Add($position, $layoutType)
        if ($p.title -and $slide.Shapes.HasTitle) { $slide.Shapes.Title.TextFrame.TextRange.Text = $p.title }
        if ($p.content) {
            try {
                foreach ($shape in $slide.Shapes) {
                    try {
                        if ($shape.PlaceholderFormat.Type -eq 2) {
                            $shape.TextFrame.TextRange.Text = $p.content
                        }
                    } catch {}
                }
            } catch {}
        }
        Output-Json @{ success = $true; data = @{ slideIndex = $position; layout = $layoutKey } }
    }

    "addTextBox" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { $ppt.ActiveWindow.Selection.SlideRange.SlideIndex }
        $slide = $pres.Slides.Item($slideIndex)
        $left = if ($p.left) { $p.left } else { 100 }
        $top = if ($p.top) { $p.top } else { 100 }
        $width = if ($p.width) { $p.width } else { 400 }
        $height = if ($p.height) { $p.height } else { 50 }
        $shape = $slide.Shapes.AddTextbox(1, $left, $top, $width, $height)
        $shape.TextFrame.TextRange.Text = $p.text
        if ($p.fontSize) { $shape.TextFrame.TextRange.Font.Size = $p.fontSize }
        if ($p.fontName) { $shape.TextFrame.TextRange.Font.Name = $p.fontName }
        Output-Json @{ success = $true; data = @{ shapeName = $shape.Name } }
    }

    "deleteTextBox" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        $shape.Delete()
        Output-Json @{ success = $true; data = @{ deleted = $(if ($p.name) { $p.name } else { $p.shapeIndex }) } }
    }

    "getTextBoxes" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $textBoxes = @()
        for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
            $shape = $slide.Shapes.Item($i)
            if ($shape.HasTextFrame) {
                $text = ""
                try { $text = $shape.TextFrame.TextRange.Text } catch {}
                $textBoxes += @{ name = $shape.Name; index = $i; text = $text; left = $shape.Left; top = $shape.Top; width = $shape.Width; height = $shape.Height }
            }
        }
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; textBoxes = $textBoxes; count = $textBoxes.Count } }
    }

    "setTextBoxText" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        $shape.TextFrame.TextRange.Text = if ($p.text) { $p.text } else { "" }
        Output-Json @{ success = $true; data = @{ name = $shape.Name; text = $p.text } }
    }

    "setTextBoxStyle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        $tr = $shape.TextFrame.TextRange
        if ($p.fontSize) { $tr.Font.Size = $p.fontSize }
        if ($p.fontName) { $tr.Font.Name = $p.fontName }
        if ($null -ne $p.bold) { $tr.Font.Bold = $p.bold }
        if ($null -ne $p.italic) { $tr.Font.Italic = $p.italic }
        if ($p.color) {
            $colorValue = Convert-HexColorToRgbInt([string]$p.color)
            if ($null -ne $colorValue) { $tr.Font.Color = $colorValue }
        }
        if ($p.alignment) {
            $alignMap = @{ left = 1; center = 2; right = 3 }
            $tr.ParagraphFormat.Alignment = $alignMap[$p.alignment]
        }
        Output-Json @{ success = $true; data = @{ name = $shape.Name } }
    }

    "setSlideTitle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        if ($slide.Shapes.HasTitle) { $slide.Shapes.Title.TextFrame.TextRange.Text = $p.title }
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; title = $p.title } }
    }

    "getSlideTitle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $title = ""
        if ($slide.Shapes.HasTitle) { $title = $slide.Shapes.Title.TextFrame.TextRange.Text }
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; title = $title } }
    }

    "setSlideSubtitle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
            $shape = $slide.Shapes.Item($i)
            $placeholderType = $null
            try { $placeholderType = $shape.PlaceholderFormat.Type } catch {}
            if ($placeholderType -eq 2) {
                $shape.TextFrame.TextRange.Text = if ($p.subtitle) { $p.subtitle } else { "" }
                Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; subtitle = $p.subtitle } }
                return
            }
        }
        Output-Json @{ success = $false; error = "Subtitle placeholder not found" }
    }

    "setSlideContent" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
            $shape = $slide.Shapes.Item($i)
            $placeholderType = $null
            try { $placeholderType = $shape.PlaceholderFormat.Type } catch {}
            if ($placeholderType -eq 7) {
                $shape.TextFrame.TextRange.Text = if ($p.content) { $p.content } else { "" }
                Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex } }
                return
            }
        }
        Output-Json @{ success = $false; error = "Content placeholder not found" }
    }

    "addShape" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shapeTypeMap = @{ rectangle = 1; oval = 9; triangle = 7; diamond = 4; pentagon = 51; hexagon = 52; arrow = 13; star = 12; heart = 21; cloud = 179 }
        $shapeType = $shapeTypeMap[$p.type]
        if ($null -eq $shapeType) { $shapeType = if ($p.type) { $p.type } else { 1 } }
        $left = if ($p.left) { $p.left } else { 100 }
        $top = if ($p.top) { $p.top } else { 100 }
        $width = if ($p.width) { $p.width } else { 100 }
        $height = if ($p.height) { $p.height } else { 100 }
        $shape = $slide.Shapes.AddShape($shapeType, $left, $top, $width, $height)
        if ($p.text) { $shape.TextFrame.TextRange.Text = $p.text }
        if ($p.fillColor) {
            $fillColor = Convert-HexColorToRgbInt([string]$p.fillColor)
            if ($null -ne $fillColor) { $shape.Fill.ForeColor.RGB = $fillColor }
        }
        Output-Json @{ success = $true; data = @{ name = $shape.Name; slideIndex = $slideIndex } }
    }

    "deleteShape" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        $shape.Delete()
        Output-Json @{ success = $true; data = @{ deleted = $(if ($p.name) { $p.name } else { $p.shapeIndex }) } }
    }

    "getShapes" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shapes = @()
        for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
            $shape = $slide.Shapes.Item($i)
            $shapes += @{ name = $shape.Name; index = $i; type = $shape.Type; left = $shape.Left; top = $shape.Top; width = $shape.Width; height = $shape.Height }
        }
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; shapes = $shapes; count = $shapes.Count } }
    }

    "setShapeStyle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        if ($p.fillColor) {
            $fillColor = Convert-HexColorToRgbInt([string]$p.fillColor)
            if ($null -ne $fillColor) { $shape.Fill.ForeColor.RGB = $fillColor }
        }
        if ($p.lineColor) {
            $lineColor = Convert-HexColorToRgbInt([string]$p.lineColor)
            if ($null -ne $lineColor) { $shape.Line.ForeColor.RGB = $lineColor }
        }
        if ($p.lineWidth) { $shape.Line.Weight = $p.lineWidth }
        Output-Json @{ success = $true; data = @{ name = $shape.Name } }
    }

    "setShapeText" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        $shape.TextFrame.TextRange.Text = if ($p.text) { $p.text } else { "" }
        Output-Json @{ success = $true; data = @{ name = $shape.Name; text = $p.text } }
    }

    "setShapePosition" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        if ($null -ne $p.left) { $shape.Left = $p.left }
        if ($null -ne $p.top) { $shape.Top = $p.top }
        if ($null -ne $p.width) { $shape.Width = $p.width }
        if ($null -ne $p.height) { $shape.Height = $p.height }
        Output-Json @{ success = $true; data = @{ name = $shape.Name; left = $shape.Left; top = $shape.Top; width = $shape.Width; height = $shape.Height } }
    }

    "insertPptImage" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $left = if ($p.left) { $p.left } else { 100 }
        $top = if ($p.top) { $p.top } else { 100 }
        $width = if ($null -ne $p.width) { $p.width } else { -1 }
        $height = if ($null -ne $p.height) { $p.height } else { -1 }
        $pic = $slide.Shapes.AddPicture($p.path, $false, $true, $left, $top, $width, $height)
        Output-Json @{ success = $true; data = @{ name = $pic.Name; path = $p.path } }
    }

    "deletePptImage" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        $shape.Delete()
        Output-Json @{ success = $true; data = @{ deleted = $(if ($p.name) { $p.name } else { $p.shapeIndex }) } }
    }

    "setImageStyle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        if ($null -ne $p.left) { $shape.Left = $p.left }
        if ($null -ne $p.top) { $shape.Top = $p.top }
        if ($null -ne $p.width) { $shape.Width = $p.width }
        if ($null -ne $p.height) { $shape.Height = $p.height }
        if ($null -ne $p.rotation) { $shape.Rotation = $p.rotation }
        Output-Json @{ success = $true; data = @{ name = $shape.Name } }
    }

    "insertPptTable" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $rows = if ($p.rows) { $p.rows } else { 3 }
        $cols = if ($p.cols) { $p.cols } else { 3 }
        $left = if ($p.left) { $p.left } else { 100 }
        $top = if ($p.top) { $p.top } else { 100 }
        $width = if ($p.width) { $p.width } else { 400 }
        $height = if ($p.height) { $p.height } else { 200 }
        $table = $slide.Shapes.AddTable($rows, $cols, $left, $top, $width, $height)
        Output-Json @{ success = $true; data = @{ name = $table.Name; rows = $rows; cols = $cols } }
    }

    "setPptTableCell" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $null
        $tableCount = 0
        $targetIndex = if ($p.tableIndex) { $p.tableIndex } else { 1 }
        if ($p.tableName) {
            $shape = $slide.Shapes.Item($p.tableName)
        } else {
            for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
                $s = $slide.Shapes.Item($i)
                if ($s.HasTable) {
                    $tableCount++
                    if ($tableCount -eq $targetIndex) { $shape = $s; break }
                }
            }
        }
        if ($null -eq $shape -or -not $shape.HasTable) { Output-Json @{ success = $false; error = "Table not found" }; exit }
        $cell = $shape.Table.Cell($p.row, $p.col)
        $textValue = if ($p.text) { $p.text } elseif ($p.value) { $p.value } else { "" }
        $cell.Shape.TextFrame.TextRange.Text = $textValue
        Output-Json @{ success = $true; data = @{ row = $p.row; col = $p.col; text = $textValue } }
    }

    "getPptTableCell" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.tableName) { $p.tableName } else { $p.tableIndex }))
        $cell = $shape.Table.Cell($p.row, $p.col)
        $value = $cell.Shape.TextFrame.TextRange.Text
        Output-Json @{ success = $true; data = @{ row = $p.row; col = $p.col; value = $value } }
    }

    "setPptTableStyle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.tableName) { $p.tableName } else { $p.tableIndex }))
        if ($null -ne $p.left) { $shape.Left = $p.left }
        if ($null -ne $p.top) { $shape.Top = $p.top }
        if ($null -ne $p.width) { $shape.Width = $p.width }
        if ($null -ne $p.height) { $shape.Height = $p.height }
        Output-Json @{ success = $true; data = @{ name = $shape.Name } }
    }

    "setPptTableCellStyle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $null
        $tableCount = 0
        $targetIndex = if ($p.tableIndex) { $p.tableIndex } else { 1 }
        if ($p.tableName) {
            $shape = $slide.Shapes.Item($p.tableName)
        } else {
            for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
                $s = $slide.Shapes.Item($i)
                if ($s.HasTable) {
                    $tableCount++
                    if ($tableCount -eq $targetIndex) { $shape = $s; break }
                }
            }
        }
        if ($null -eq $shape -or -not $shape.HasTable) { Output-Json @{ success = $false; error = "Table not found" }; exit }
        $cell = $shape.Table.Cell($p.row, $p.col)
        $cellShape = $cell.Shape
        if ($p.backgroundColor) {
            $bg = Convert-HexColorToRgbInt([string]$p.backgroundColor)
            if ($null -ne $bg) { $cellShape.Fill.Visible = $true; $cellShape.Fill.Solid(); $cellShape.Fill.ForeColor.RGB = $bg }
        }
        if ($p.fontColor) {
            $fc = Convert-HexColorToRgbInt([string]$p.fontColor)
            if ($null -ne $fc) { $cellShape.TextFrame.TextRange.Font.Color.RGB = $fc }
        }
        if ($p.fontSize) { $cellShape.TextFrame.TextRange.Font.Size = $p.fontSize }
        if ($null -ne $p.bold) { $cellShape.TextFrame.TextRange.Font.Bold = $p.bold }
        Output-Json @{ success = $true; data = @{ row = $p.row; col = $p.col } }
    }

    "setPptTableRowStyle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $null
        $tableCount = 0
        $targetIndex = if ($p.tableIndex) { $p.tableIndex } else { 1 }
        if ($p.tableName) {
            $shape = $slide.Shapes.Item($p.tableName)
        } else {
            for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
                $s = $slide.Shapes.Item($i)
                if ($s.HasTable) {
                    $tableCount++
                    if ($tableCount -eq $targetIndex) { $shape = $s; break }
                }
            }
        }
        if ($null -eq $shape -or -not $shape.HasTable) { Output-Json @{ success = $false; error = "Table not found" }; exit }
        $rowIndex = if ($p.row) { $p.row } else { 1 }
        $row = $shape.Table.Rows.Item($rowIndex)
        for ($c = 1; $c -le $row.Cells.Count; $c++) {
            $cellShape = $row.Cells.Item($c).Shape
            if ($p.backgroundColor) {
                $bg = Convert-HexColorToRgbInt([string]$p.backgroundColor)
                if ($null -ne $bg) { $cellShape.Fill.Visible = $true; $cellShape.Fill.Solid(); $cellShape.Fill.ForeColor.RGB = $bg }
            }
            if ($p.fontColor) {
                $fc = Convert-HexColorToRgbInt([string]$p.fontColor)
                if ($null -ne $fc) { $cellShape.TextFrame.TextRange.Font.Color.RGB = $fc }
            }
            if ($p.fontSize) { $cellShape.TextFrame.TextRange.Font.Size = $p.fontSize }
            if ($null -ne $p.bold) { $cellShape.TextFrame.TextRange.Font.Bold = $p.bold }
        }
        Output-Json @{ success = $true; data = @{ row = $rowIndex } }
    }

    "setShapeShadow" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        $shape.Shadow.Visible = $true
        if ($p.color) {
            $shadowColor = Convert-HexColorToRgbInt([string]$p.color)
            if ($null -ne $shadowColor) { $shape.Shadow.ForeColor.RGB = $shadowColor }
        }
        if ($null -ne $p.transparency) { $shape.Shadow.Transparency = $p.transparency }
        if ($null -ne $p.blur) { $shape.Shadow.Blur = $p.blur }
        if ($null -ne $p.offsetX) { $shape.Shadow.OffsetX = $p.offsetX }
        if ($null -ne $p.offsetY) { $shape.Shadow.OffsetY = $p.offsetY }
        Output-Json @{ success = $true; data = @{ name = $shape.Name } }
    }

    "setBackgroundGradient" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $bg = $slide.Background.Fill
        $bg.Visible = $true
        $bg.TwoColorGradient(1, 1)
        if ($p.color1) {
            $c1 = Convert-HexColorToRgbInt([string]$p.color1)
            if ($null -ne $c1) { $bg.ForeColor.RGB = $c1 }
        }
        if ($p.color2) {
            $c2 = Convert-HexColorToRgbInt([string]$p.color2)
            if ($null -ne $c2) { $bg.BackColor.RGB = $c2 }
        }
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex } }
    }

    "setShapeGradient" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        $shape.Fill.TwoColorGradient(1, 1)
        if ($p.color1) {
            $c1 = Convert-HexColorToRgbInt([string]$p.color1)
            if ($null -ne $c1) { $shape.Fill.ForeColor.RGB = $c1 }
        }
        if ($p.color2) {
            $c2 = Convert-HexColorToRgbInt([string]$p.color2)
            if ($null -ne $c2) { $shape.Fill.BackColor.RGB = $c2 }
        }
        Output-Json @{ success = $true; data = @{ name = $shape.Name } }
    }

    "setShapeBorder" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        if ($p.color) {
            $borderColor = Convert-HexColorToRgbInt([string]$p.color)
            if ($null -ne $borderColor) { $shape.Line.ForeColor.RGB = $borderColor }
        }
        if ($p.width) { $shape.Line.Weight = $p.width }
        Output-Json @{ success = $true; data = @{ name = $shape.Name } }
    }

    "setShapeTransparency" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        if ($null -ne $p.transparency) { $shape.Fill.Transparency = $p.transparency }
        Output-Json @{ success = $true; data = @{ name = $shape.Name } }
    }

    "setShapeRoundness" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        if ($null -ne $p.roundness) { $shape.Adjustments.Item(1) = $p.roundness }
        Output-Json @{ success = $true; data = @{ name = $shape.Name } }
    }

    "setShapeFullStyle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        if ($p.fillColor) { $shape.Fill.ForeColor.RGB = Convert-HexColorToRgbInt([string]$p.fillColor) }
        if ($p.lineColor) { $shape.Line.ForeColor.RGB = Convert-HexColorToRgbInt([string]$p.lineColor) }
        if ($p.lineWidth) { $shape.Line.Weight = $p.lineWidth }
        if ($null -ne $p.transparency) { $shape.Fill.Transparency = $p.transparency }
        Output-Json @{ success = $true; data = @{ name = $shape.Name } }
    }

    "alignShapes" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $alignMap = @{ left = 1; center = 2; right = 3; top = 4; middle = 5; bottom = 6 }
        $align = $alignMap[$p.alignment]
        if ($null -eq $align) { $align = 1 }
        $range = $slide.Shapes.Range($p.names)
        $range.Align($align, 1)
        Output-Json @{ success = $true; data = @{ alignment = $p.alignment } }
    }

    "distributeShapes" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $distMap = @{ horizontal = 0; vertical = 1 }
        $dist = $distMap[$p.direction]
        if ($null -eq $dist) { $dist = 0 }
        $range = $slide.Shapes.Range($p.names)
        $range.Distribute($dist, 1)
        Output-Json @{ success = $true; data = @{ direction = $p.direction } }
    }

    "groupShapes" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $range = $slide.Shapes.Range($p.names)
        $group = $range.Group()
        Output-Json @{ success = $true; data = @{ name = $group.Name } }
    }

    "duplicateShape" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        $dup = $shape.Duplicate()
        Output-Json @{ success = $true; data = @{ name = $dup.Name } }
    }

    "setShapeZOrder" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        $zMap = @{ front = 0; forward = 2; backward = 3; back = 1 }
        $z = $zMap[$p.zOrder]
        if ($null -eq $z) { $z = 0 }
        $shape.ZOrder($z)
        Output-Json @{ success = $true; data = @{ name = $shape.Name; zOrder = $p.zOrder } }
    }

    "insertPptChart" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $left = if ($p.left) { $p.left } else { 100 }
        $top = if ($p.top) { $p.top } else { 100 }
        $width = if ($p.width) { $p.width } else { 400 }
        $height = if ($p.height) { $p.height } else { 300 }
        $shape = $slide.Shapes.AddChart($p.type, $left, $top, $width, $height)
        Output-Json @{ success = $true; data = @{ name = $shape.Name } }
    }

    "addAnimation" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($p.shapeName)
        $slide.TimeLine.MainSequence.AddEffect($shape, $p.effect, 1)
        Output-Json @{ success = $true }
    }

    "setSlideTransition" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $slide.SlideShowTransition.EntryEffect = $p.effect
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; effect = $p.effect } }
    }

    "setSlideBackground" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        if ($p.color) {
            $colorValue = Convert-HexColorToRgbInt([string]$p.color)
            if ($null -ne $colorValue) { $slide.Background.Fill.ForeColor.RGB = $colorValue }
        }
        if ($p.imagePath) {
            $slide.FollowMasterBackground = $false
            $slide.Background.Fill.UserPicture($p.imagePath)
        }
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex } }
    }

    "slide.unifyFont" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $fontName = if ($p.fontName) { $p.fontName } else { "微软雅黑" }
        $count = 0
        $includeTitle = if ($null -ne $p.includeTitle) { [bool]$p.includeTitle } else { $true }
        $includeBody = if ($null -ne $p.includeBody) { [bool]$p.includeBody } else { $true }
        $start = if ($p.slideIndex) { [int]$p.slideIndex } else { 1 }
        $end = if ($p.slideIndex) { [int]$p.slideIndex } else { $pres.Slides.Count }
        for ($i = $start; $i -le $end; $i++) {
            $slide = $pres.Slides.Item($i)
            for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
                $shape = $slide.Shapes.Item($j)
                try {
                    if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                        $isTitle = $false
                        try { $isTitle = $shape.PlaceholderFormat.Type -eq 1 -or $shape.PlaceholderFormat.Type -eq 3 } catch {}
                        if (($isTitle -and $includeTitle) -or (-not $isTitle -and $includeBody)) {
                            $shape.TextFrame.TextRange.Font.Name = $fontName
                            $count++
                        }
                    }
                } catch {}
            }
        }
        $slideCount = if ($p.slideIndex) { 1 } else { $pres.Slides.Count }
        Output-Json @{ success = $true; data = @{ font = $fontName; count = $count; slideCount = $slideCount } }
    }

    "addConnector" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $left = if ($p.left) { $p.left } else { 100 }
        $top = if ($p.top) { $p.top } else { 100 }
        $width = if ($p.width) { $p.width } else { 200 }
        $height = if ($p.height) { $p.height } else { 0 }
        $connector = $slide.Shapes.AddConnector(1, $left, $top, $width, $height)
        Output-Json @{ success = $true; data = @{ name = $connector.Name } }
    }

    "addArrow" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $left = if ($p.left) { $p.left } else { 100 }
        $top = if ($p.top) { $p.top } else { 100 }
        $width = if ($p.width) { $p.width } else { 120 }
        $height = if ($p.height) { $p.height } else { 60 }
        $shape = $slide.Shapes.AddShape(13, $left, $top, $width, $height)
        Output-Json @{ success = $true; data = @{ name = $shape.Name } }
    }

    "applyColorScheme" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $schemeKey = if ($p.scheme) { $p.scheme } elseif ($p.colorScheme) { $p.colorScheme } else { "business" }
        $schemes = @{ business = @{ title = 0x2F5496; body = 0x333333 }; tech = @{ title = 0x00B0F0; body = 0x404040 }; creative = @{ title = 0xFF6B6B; body = 0x4A4A4A }; minimal = @{ title = 0x000000; body = 0x666666 } }
        $scheme = $schemes[$schemeKey]
        if ($null -eq $scheme) { $scheme = $schemes["business"] }
        $count = 0
        for ($i = 1; $i -le $pres.Slides.Count; $i++) {
            $slide = $pres.Slides.Item($i)
            for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
                $shape = $slide.Shapes.Item($j)
                try {
                    if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                        $textRange = $shape.TextFrame.TextRange
                        if ($textRange.Font.Size -ge 24) { $textRange.Font.Color.RGB = $scheme.title } else { $textRange.Font.Color.RGB = $scheme.body }
                        $count++
                    }
                } catch {}
            }
        }
        Output-Json @{ success = $true; data = @{ scheme = $schemeKey; count = $count } }
    }

    "autoBeautifySlide" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { $ppt.ActiveWindow.Selection.SlideRange.SlideIndex }
        $slide = $pres.Slides.Item($slideIndex)
        $schemes = @{ business = @{ title = 0x2F5496; body = 0x333333 }; tech = @{ title = 0x00B0F0; body = 0x404040 }; creative = @{ title = 0xFF6B6B; body = 0x4A4A4A }; minimal = @{ title = 0x000000; body = 0x666666 } }
        $schemeKey = if ($p.style) { $p.style } else { "business" }
        $scheme = $schemes[$schemeKey]
        if ($null -eq $scheme) { $scheme = $schemes["business"] }
        $count = 0
        for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
            $shape = $slide.Shapes.Item($j)
            try {
                if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                    $textRange = $shape.TextFrame.TextRange
                    if ($textRange.Font.Size -ge 24) { $textRange.Font.Color.RGB = $scheme.title } else { $textRange.Font.Color.RGB = $scheme.body }
                    $count++
                }
            } catch {}
        }
        Output-Json @{ success = $true; data = @{ style = $schemeKey; count = $count } }
    }

    "createKpiCards" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $cards = if ($p.cards) { $p.cards } else { @() }
        $startX = if ($p.startX) { $p.startX } else { 50 }
        $startY = if ($p.startY) { $p.startY } else { 350 }
        $cardWidth = if ($p.cardWidth) { $p.cardWidth } else { 150 }
        $cardHeight = if ($p.cardHeight) { $p.cardHeight } else { 80 }
        $gap = if ($p.gap) { $p.gap } else { 20 }

        $colors = @("#28a745", "#17a2b8", "#ffc107", "#dc3545", "#6f42c1", "#fd7e14")
        $created = @()

        for ($i = 0; $i -lt $cards.Count; $i++) {
            $card = $cards[$i]
            $x = $startX + ($i * ($cardWidth + $gap))
            $color = if ($card.color) { $card.color } else { $colors[$i % $colors.Count] }
            $shape = $slide.Shapes.AddShape(5, $x, $startY, $cardWidth, $cardHeight)
            $colorValue = Convert-HexColorToRgbInt([string]$color)
            if ($null -ne $colorValue) { $shape.Fill.Solid(); $shape.Fill.ForeColor.RGB = $colorValue }
            $shape.Line.Visible = $false
            $shape.Shadow.Visible = $true
            $shape.Shadow.Blur = 10
            $shape.Shadow.OffsetX = 3
            $shape.Shadow.OffsetY = 3
            $shape.Shadow.Transparency = 0.6
            $title = if ($card.title) { $card.title } else { "" }
            $value = if ($card.value) { $card.value } else { "" }
            $shape.TextFrame.TextRange.Text = ($title + "`n" + $value)
            $shape.TextFrame.TextRange.Font.Color.RGB = 16777215
            $shape.TextFrame.TextRange.Font.Size = 14
            $shape.TextFrame.TextRange.Font.Bold = $true
            $shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
            $created += $shape.Name
        }

        Output-Json @{ success = $true; data = @{ count = $created.Count; cards = $created } }
    }

    "createStyledTable" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $rows = if ($p.rows) { $p.rows } else { 5 }
        $cols = if ($p.cols) { $p.cols } else { 4 }
        $left = if ($p.left) { $p.left } else { 50 }
        $top = if ($p.top) { $p.top } else { 150 }
        $width = if ($p.width) { $p.width } else { 620 }
        $height = if ($p.height) { $p.height } else { 200 }

        $tableShape = $slide.Shapes.AddTable($rows, $cols, $left, $top, $width, $height)
        $table = $tableShape.Table

        $styleKey = if ($p.style) { $p.style } else { "business" }
        $styles = @{
            business = @{ header = "#1a365d"; headerText = "#ffffff"; oddRow = "#f8f9fa"; evenRow = "#ffffff" }
            tech = @{ header = "#0d47a1"; headerText = "#ffffff"; oddRow = "#e3f2fd"; evenRow = "#ffffff" }
            elegant = @{ header = "#37474f"; headerText = "#ffffff"; oddRow = "#eceff1"; evenRow = "#ffffff" }
        }
        $style = $styles[$styleKey]
        if ($null -eq $style) { $style = $styles["business"] }

        for ($c = 1; $c -le $cols; $c++) {
            $headerCell = $table.Cell(1, $c)
            $headerColor = Convert-HexColorToRgbInt([string]$style.header)
            if ($null -ne $headerColor) { $headerCell.Shape.Fill.Solid(); $headerCell.Shape.Fill.ForeColor.RGB = $headerColor }
            $headerTextColor = Convert-HexColorToRgbInt([string]$style.headerText)
            if ($null -ne $headerTextColor) { $headerCell.Shape.TextFrame.TextRange.Font.Color.RGB = $headerTextColor }
            $headerCell.Shape.TextFrame.TextRange.Font.Bold = $true
            $headerCell.Shape.TextFrame.TextRange.Font.Size = 12
        }

        for ($r = 2; $r -le $rows; $r++) {
            $rowColor = if ($r % 2 -eq 0) { $style.evenRow } else { $style.oddRow }
            $rowColorValue = Convert-HexColorToRgbInt([string]$rowColor)
            for ($c = 1; $c -le $cols; $c++) {
                $cell = $table.Cell($r, $c)
                if ($null -ne $rowColorValue) { $cell.Shape.Fill.Solid(); $cell.Shape.Fill.ForeColor.RGB = $rowColorValue }
                $cell.Shape.TextFrame.TextRange.Font.Size = 11
            }
        }

        if ($p.data) {
            for ($r = 0; $r -lt $p.data.Count; $r++) {
                $rowData = $p.data[$r]
                for ($c = 0; $c -lt $rowData.Count; $c++) {
                    if (($r + 1) -le $rows -and ($c + 1) -le $cols) {
                        $table.Cell($r + 1, $c + 1).Shape.TextFrame.TextRange.Text = [string]$rowData[$c]
                    }
                }
            }
        }

        Output-Json @{ success = $true; data = @{ name = $tableShape.Name; rows = $rows; cols = $cols } }
    }

    "addTitleDecoration" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $left = if ($p.left) { $p.left } else { 50 }
        $top = if ($p.top) { $p.top } else { 130 }
        $width = if ($p.width) { $p.width } else { 200 }
        $height = if ($p.height) { $p.height } else { 5 }
        $colorValue = Convert-HexColorToRgbInt([string]$(if ($p.color) { $p.color } else { "#1a365d" }))
        $bar = $slide.Shapes.AddShape(1, $left, $top, $width, $height)
        if ($null -ne $colorValue) { $bar.Fill.Solid(); $bar.Fill.ForeColor.RGB = $colorValue }
        $bar.Line.Visible = $false
        Output-Json @{ success = $true; data = @{ name = $bar.Name } }
    }

    "addPageIndicator" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $total = $pres.Slides.Count
        $current = $slideIndex
        $text = "$current / $total"
        $left = if ($p.left) { $p.left } else { 880 }
        $top = if ($p.top) { $p.top } else { 500 }
        $box = $slide.Shapes.AddTextbox(1, $left, $top, 80, 30)
        $box.TextFrame.TextRange.Text = $text
        $box.TextFrame.TextRange.Font.Size = 12
        $box.TextFrame.TextRange.Font.Color.RGB = if ($p.dark) { 0 } else { 16777215 }
        $box.TextFrame.TextRange.ParagraphFormat.Alignment = 3
        Output-Json @{ success = $true; data = @{ name = $box.Name; page = $text } }
    }

    "beautifyAllSlides" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $schemeKey = if ($p.style) { $p.style } else { "business" }
        $schemes = @{ business = @{ title = 0x2F5496; body = 0x333333 }; tech = @{ title = 0x00B0F0; body = 0x404040 }; creative = @{ title = 0xFF6B6B; body = 0x4A4A4A }; minimal = @{ title = 0x000000; body = 0x666666 } }
        $scheme = $schemes[$schemeKey]
        if ($null -eq $scheme) { $scheme = $schemes["business"] }
        $count = 0
        for ($i = 1; $i -le $pres.Slides.Count; $i++) {
            $slide = $pres.Slides.Item($i)
            for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
                $shape = $slide.Shapes.Item($j)
                try {
                    if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                        $textRange = $shape.TextFrame.TextRange
                        if ($textRange.Font.Size -ge 24) { $textRange.Font.Color.RGB = $scheme.title } else { $textRange.Font.Color.RGB = $scheme.body }
                        $count++
                    }
                } catch {}
            }
        }
        Output-Json @{ success = $true; data = @{ style = $schemeKey; count = $count } }
    }

    "createProgressBar" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $left = if ($p.left) { $p.left } else { 100 }
        $top = if ($p.top) { $p.top } else { 300 }
        $width = if ($p.width) { $p.width } else { 400 }
        $height = if ($p.height) { $p.height } else { 30 }
        $progress = if ($null -ne $p.progress) { [double]$p.progress } else { 0.5 }
        if ($progress -lt 0) { $progress = 0 }
        if ($progress -gt 1) { $progress = 1 }
        $label = if ($p.label) { $p.label } else { "" }

        $bgBar = $slide.Shapes.AddShape(5, $left, $top, $width, $height)
        $bgColor = Convert-HexColorToRgbInt([string]$(if ($p.bgColor) { $p.bgColor } else { "#e0e0e0" }))
        if ($null -ne $bgColor) { $bgBar.Fill.Solid(); $bgBar.Fill.ForeColor.RGB = $bgColor }
        $bgBar.Line.Visible = $false

        $progressWidth = $width * $progress
        $progressBar = $slide.Shapes.AddShape(5, $left, $top, $progressWidth, $height)
        $colorValue = Convert-HexColorToRgbInt([string]$(if ($p.color) { $p.color } else { "#28a745" }))
        if ($null -ne $colorValue) { $progressBar.Fill.Solid(); $progressBar.Fill.ForeColor.RGB = $colorValue }
        $progressBar.Line.Visible = $false

        $percentText = $slide.Shapes.AddTextbox(1, $left + $width + 10, $top, 60, $height)
        $percentText.TextFrame.TextRange.Text = "{0}%" -f ([Math]::Round($progress * 100))
        $percentText.TextFrame.TextRange.Font.Size = 14
        $percentText.TextFrame.TextRange.Font.Bold = $true

        if ($label) {
            $labelBox = $slide.Shapes.AddTextbox(1, $left, $top - 25, $width, 20)
            $labelBox.TextFrame.TextRange.Text = $label
            $labelBox.TextFrame.TextRange.Font.Size = 12
        }

        Output-Json @{ success = $true; data = @{ bgBar = $bgBar.Name; progressBar = $progressBar.Name; progress = $progress } }
    }

    "createGauge" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $centerX = if ($p.centerX) { $p.centerX } else { 360 }
        $centerY = if ($p.centerY) { $p.centerY } else { 300 }
        $radius = if ($p.radius) { $p.radius } else { 100 }
        $value = if ($null -ne $p.value) { [double]$p.value } else { 0.5 }
        if ($value -lt 0) { $value = 0 }
        if ($value -gt 1) { $value = 1 }
        $title = if ($p.title) { $p.title } else { "" }

        $outerCircle = $slide.Shapes.AddShape(9, $centerX - $radius, $centerY - $radius, $radius * 2, $radius * 2)
        $outerColor = Convert-HexColorToRgbInt([string]"#e0e0e0")
        if ($null -ne $outerColor) { $outerCircle.Fill.Solid(); $outerCircle.Fill.ForeColor.RGB = $outerColor }
        $outerCircle.Line.Visible = $false

        $innerRadius = $radius * 0.7
        $innerCircle = $slide.Shapes.AddShape(9, $centerX - $innerRadius, $centerY - $innerRadius, $innerRadius * 2, $innerRadius * 2)
        $innerColor = Convert-HexColorToRgbInt([string]"#ffffff")
        if ($null -ne $innerColor) { $innerCircle.Fill.Solid(); $innerCircle.Fill.ForeColor.RGB = $innerColor }
        $innerCircle.Line.Visible = $false

        $valueText = $slide.Shapes.AddTextbox(1, $centerX - 40, $centerY - 20, 80, 40)
        $valueText.TextFrame.TextRange.Text = "{0}%" -f ([Math]::Round($value * 100))
        $valueText.TextFrame.TextRange.Font.Size = 24
        $valueText.TextFrame.TextRange.Font.Bold = $true
        $valueText.TextFrame.TextRange.ParagraphFormat.Alignment = 2

        if ($title) {
            $titleBox = $slide.Shapes.AddTextbox(1, $centerX - 60, $centerY + $radius + 10, 120, 25)
            $titleBox.TextFrame.TextRange.Text = $title
            $titleBox.TextFrame.TextRange.Font.Size = 14
            $titleBox.TextFrame.TextRange.ParagraphFormat.Alignment = 2
        }

        Output-Json @{ success = $true; data = @{ gauge = $outerCircle.Name; value = $value } }
    }

    "createMiniCharts" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $items = if ($p.items) { $p.items } else { @() }
        $startX = if ($p.startX) { $p.startX } else { 50 }
        $startY = if ($p.startY) { $p.startY } else { 400 }
        $itemWidth = if ($p.itemWidth) { $p.itemWidth } else { 100 }
        $gap = if ($p.gap) { $p.gap } else { 30 }
        $colors = @("#28a745", "#17a2b8", "#ffc107", "#dc3545", "#6f42c1")
        $created = @()

        for ($i = 0; $i -lt $items.Count; $i++) {
            $item = $items[$i]
            $x = $startX + ($i * ($itemWidth + $gap))
            $color = if ($item.color) { $item.color } else { $colors[$i % $colors.Count] }
            $colorValue = Convert-HexColorToRgbInt([string]$color)

            $valueBox = $slide.Shapes.AddTextbox(1, $x, $startY, $itemWidth, 30)
            $valueBox.TextFrame.TextRange.Text = if ($item.value) { $item.value } else { "0" }
            $valueBox.TextFrame.TextRange.Font.Size = 20
            $valueBox.TextFrame.TextRange.Font.Bold = $true
            if ($null -ne $colorValue) { $valueBox.TextFrame.TextRange.Font.Color.RGB = $colorValue }

            $labelBox = $slide.Shapes.AddTextbox(1, $x, $startY + 30, $itemWidth, 20)
            $labelBox.TextFrame.TextRange.Text = if ($item.label) { $item.label } else { "" }
            $labelBox.TextFrame.TextRange.Font.Size = 11
            $labelBox.TextFrame.TextRange.Font.Color.RGB = 6710886

            if ($item.trend) {
                $trendBox = $slide.Shapes.AddTextbox(1, $x + $itemWidth - 30, $startY, 30, 20)
                $trendBox.TextFrame.TextRange.Text = if ($item.trend -eq "up") { "↑" } else { "↓" }
                $trendBox.TextFrame.TextRange.Font.Size = 14
                $trendColor = if ($item.trend -eq "up") { "#28a745" } else { "#dc3545" }
                $trendColorValue = Convert-HexColorToRgbInt([string]$trendColor)
                if ($null -ne $trendColorValue) { $trendBox.TextFrame.TextRange.Font.Color.RGB = $trendColorValue }
            }

            $created += @{ label = $item.label; value = $item.value }
        }

        Output-Json @{ success = $true; data = @{ count = $created.Count; items = $created } }
    }

    "createDonutChart" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $centerX = if ($p.centerX) { $p.centerX } else { 360 }
        $centerY = if ($p.centerY) { $p.centerY } else { 280 }
        $radius = if ($p.radius) { $p.radius } else { 80 }
        $value = if ($null -ne $p.value) { [double]$p.value } else { 0.75 }
        if ($value -lt 0) { $value = 0 }
        if ($value -gt 1) { $value = 1 }
        $title = if ($p.title) { $p.title } else { "" }
        $centerText = if ($p.centerText) { $p.centerText } else { "{0}%" -f ([Math]::Round($value * 100)) }

        $outerCircle = $slide.Shapes.AddShape(9, $centerX - $radius, $centerY - $radius, $radius * 2, $radius * 2)
        $outerColor = Convert-HexColorToRgbInt([string]$(if ($p.color) { $p.color } else { "#0d47a1" }))
        if ($null -ne $outerColor) { $outerCircle.Fill.Solid(); $outerCircle.Fill.ForeColor.RGB = $outerColor }
        $outerCircle.Line.Visible = $false

        $innerRadius = $radius * 0.6
        $innerCircle = $slide.Shapes.AddShape(9, $centerX - $innerRadius, $centerY - $innerRadius, $innerRadius * 2, $innerRadius * 2)
        $innerColor = Convert-HexColorToRgbInt([string]"#ffffff")
        if ($null -ne $innerColor) { $innerCircle.Fill.Solid(); $innerCircle.Fill.ForeColor.RGB = $innerColor }
        $innerCircle.Line.Visible = $false

        $textBox = $slide.Shapes.AddTextbox(1, $centerX - 40, $centerY - 15, 80, 30)
        $textBox.TextFrame.TextRange.Text = $centerText
        $textBox.TextFrame.TextRange.Font.Size = 18
        $textBox.TextFrame.TextRange.Font.Bold = $true
        $textBox.TextFrame.TextRange.ParagraphFormat.Alignment = 2

        if ($title) {
            $titleBox = $slide.Shapes.AddTextbox(1, $centerX - 60, $centerY + $radius + 15, 120, 25)
            $titleBox.TextFrame.TextRange.Text = $title
            $titleBox.TextFrame.TextRange.Font.Size = 12
            $titleBox.TextFrame.TextRange.ParagraphFormat.Alignment = 2
        }

        Output-Json @{ success = $true; data = @{ donut = $outerCircle.Name; value = $value } }
    }

    "autoLayout" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $layout = if ($p.layout) { $p.layout } else { "grid" }
        $margin = if ($p.margin) { $p.margin } else { 50 }
        $gap = if ($p.gap) { $p.gap } else { 20 }
        $slideWidth = 720
        $slideHeight = 540

        $shapes = @()
        for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
            $s = $slide.Shapes.Item($i)
            if ($s.Type -eq 1 -or $s.Type -eq 6 -or $s.Type -eq 17) { $shapes += $s }
        }

        if ($shapes.Count -eq 0) {
            Output-Json @{ success = $true; data = @{ message = "没有可排列的形状"; count = 0 } }
            break
        }

        if ($layout -eq "horizontal") {
            $totalWidth = 0
            foreach ($shape in $shapes) { $totalWidth += $shape.Width }
            $availWidth = $slideWidth - 2 * $margin
            $spacing = ($availWidth - $totalWidth) / ($shapes.Count + 1)
            $currentX = $margin + $spacing
            foreach ($shape in $shapes) {
                $shape.Left = $currentX
                $shape.Top = ($slideHeight - $shape.Height) / 2
                $currentX += $shape.Width + $spacing
            }
        } elseif ($layout -eq "vertical") {
            $currentY = $margin
            foreach ($shape in $shapes) {
                $shape.Left = ($slideWidth - $shape.Width) / 2
                $shape.Top = $currentY
                $currentY += $shape.Height + $gap
            }
        } else {
            $cols = if ($p.cols) { $p.cols } else { [Math]::Ceiling([Math]::Sqrt($shapes.Count)) }
            $rows = [Math]::Ceiling($shapes.Count / $cols)
            $cellWidth = ($slideWidth - 2 * $margin) / $cols
            $cellHeight = ($slideHeight - 2 * $margin) / $rows
            for ($idx = 0; $idx -lt $shapes.Count; $idx++) {
                $col = $idx % $cols
                $row = [Math]::Floor($idx / $cols)
                $shape = $shapes[$idx]
                $shape.Left = $margin + $col * $cellWidth + ($cellWidth - $shape.Width) / 2
                $shape.Top = $margin + $row * $cellHeight + ($cellHeight - $shape.Height) / 2
            }
        }

        Output-Json @{ success = $true; data = @{ layout = $layout; count = $shapes.Count } }
    }

    "smartDistribute" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $direction = if ($p.direction) { $p.direction } else { "horizontal" }
        $shapeNames = if ($p.shapes) { $p.shapes } else { @() }
        $shapes = @()

        if ($shapeNames.Count -gt 0) {
            foreach ($name in $shapeNames) {
                try { $shapes += $slide.Shapes.Item($name) } catch {}
            }
        } else {
            for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
                $s = $slide.Shapes.Item($i)
                if ($s.Type -eq 1 -or $s.Type -eq 6 -or $s.Type -eq 17) { $shapes += $s }
            }
        }

        if ($shapes.Count -lt 2) {
            Output-Json @{ success = $true; data = @{ message = "需要至少2个形状"; count = $shapes.Count } }
            break
        }

        $shapes = if ($direction -eq "horizontal") { $shapes | Sort-Object Left } else { $shapes | Sort-Object Top }

        if ($direction -eq "horizontal") {
            $minX = $shapes[0].Left
            $maxX = $shapes[$shapes.Count - 1].Left + $shapes[$shapes.Count - 1].Width
            $totalWidth = 0
            foreach ($shape in $shapes) { $totalWidth += $shape.Width }
            $spacing = ($maxX - $minX - $totalWidth) / ($shapes.Count - 1)
            $currentX = $minX
            foreach ($shape in $shapes) {
                $shape.Left = $currentX
                $currentX += $shape.Width + $spacing
            }
        } else {
            $minY = $shapes[0].Top
            $maxY = $shapes[$shapes.Count - 1].Top + $shapes[$shapes.Count - 1].Height
            $totalHeight = 0
            foreach ($shape in $shapes) { $totalHeight += $shape.Height }
            $spacing = ($maxY - $minY - $totalHeight) / ($shapes.Count - 1)
            $currentY = $minY
            foreach ($shape in $shapes) {
                $shape.Top = $currentY
                $currentY += $shape.Height + $spacing
            }
        }

        Output-Json @{ success = $true; data = @{ direction = $direction; count = $shapes.Count } }
    }

    "createGrid" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $rows = if ($p.rows) { $p.rows } else { 2 }
        $cols = if ($p.cols) { $p.cols } else { 3 }
        $margin = if ($p.margin) { $p.margin } else { 50 }
        $gap = if ($p.gap) { $p.gap } else { 15 }
        $slideWidth = 720
        $slideHeight = 540

        $cellWidth = ($slideWidth - 2 * $margin - ($cols - 1) * $gap) / $cols
        $cellHeight = ($slideHeight - 2 * $margin - ($rows - 1) * $gap) / $rows
        $cells = @()

        for ($r = 0; $r -lt $rows; $r++) {
            for ($c = 0; $c -lt $cols; $c++) {
                $left = $margin + $c * ($cellWidth + $gap)
                $top = $margin + $r * ($cellHeight + $gap)
                $cell = $slide.Shapes.AddShape(5, $left, $top, $cellWidth, $cellHeight)
                $cellColor = Convert-HexColorToRgbInt([string]$(if ($p.cellColor) { $p.cellColor } else { "#f8f9fa" }))
                $borderColor = Convert-HexColorToRgbInt([string]$(if ($p.borderColor) { $p.borderColor } else { "#dee2e6" }))
                if ($null -ne $cellColor) { $cell.Fill.Solid(); $cell.Fill.ForeColor.RGB = $cellColor }
                if ($null -ne $borderColor) { $cell.Line.ForeColor.RGB = $borderColor }
                $cell.Line.Weight = 1
                $cells += $cell.Name
            }
        }

        Output-Json @{ success = $true; data = @{ rows = $rows; cols = $cols; cells = $cells } }
    }

    "addAnimationPreset" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $preset = if ($p.preset) { $p.preset } else { "fadeIn" }
        $timeline = $slide.TimeLine
        $delay = 0
        $delayIncrement = if ($p.delayIncrement) { $p.delayIncrement } else { 0.3 }
        $presets = @{ fadeIn = @{ effect = 10; duration = 0.5 }; flyIn = @{ effect = 2; duration = 0.5 }; zoomIn = @{ effect = 53; duration = 0.4 }; wipeIn = @{ effect = 22; duration = 0.5 }; appear = @{ effect = 1; duration = 0 } }
        $config = $presets[$preset]
        if ($null -eq $config) { $config = $presets["fadeIn"] }
        $animatedCount = 0

        for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
            $shape = $slide.Shapes.Item($i)
            try {
                $effect = $timeline.MainSequence.AddEffect($shape, $config.effect, 0, 1)
                $effect.Timing.Duration = $config.duration
                $effect.Timing.TriggerDelayTime = $delay
                $delay += $delayIncrement
                $animatedCount++
            } catch {}
        }

        Output-Json @{ success = $true; data = @{ preset = $preset; animatedShapes = $animatedCount } }
    }

    "addEmphasisAnimation" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.shapeName) { $p.shapeName } else { $p.shapeIndex }))

        $effectType = if ($p.effect) { $p.effect } else { "pulse" }
        $effectMap = @{ pulse = 63; spin = 15; grow = 53; teeter = 28 }
        $effectId = if ($effectMap[$effectType]) { $effectMap[$effectType] } else { 63 }

        $timeline = $slide.TimeLine
        $effect = $timeline.MainSequence.AddEffect($shape, $effectId, 0, 2)
        $effect.Timing.Duration = if ($p.duration) { $p.duration } else { 0.5 }

        Output-Json @{ success = $true; data = @{ shape = $shape.Name; effect = $effectType } }
    }

    "createFlowChart" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $steps = if ($p.steps) { $p.steps } else { @("开始", "步骤1", "步骤2", "结束") }
        $direction = if ($p.direction) { $p.direction } else { "horizontal" }
        $startX = if ($p.startX) { $p.startX } else { 80 }
        $startY = if ($p.startY) { $p.startY } else { 250 }
        $boxWidth = if ($p.boxWidth) { $p.boxWidth } else { 100 }
        $boxHeight = if ($p.boxHeight) { $p.boxHeight } else { 50 }
        $gap = if ($p.gap) { $p.gap } else { 60 }
        $colors = @{ start = "#28a745"; process = "#0d47a1"; end = "#dc3545" }
        $shapes = @()

        for ($i = 0; $i -lt $steps.Count; $i++) {
            if ($direction -eq "horizontal") {
                $x = $startX + $i * ($boxWidth + $gap)
                $y = $startY
            } else {
                $x = $startX
                $y = $startY + $i * ($boxHeight + $gap)
            }

            $shapeType = 5
            $colorKey = "process"
            if ($i -eq 0) { $shapeType = 9; $colorKey = "start" }
            elseif ($i -eq $steps.Count - 1) { $shapeType = 9; $colorKey = "end" }

            $shape = $slide.Shapes.AddShape($shapeType, $x, $y, $boxWidth, $boxHeight)
            $colorValue = Convert-HexColorToRgbInt([string]$colors[$colorKey])
            if ($null -ne $colorValue) { $shape.Fill.Solid(); $shape.Fill.ForeColor.RGB = $colorValue }
            $shape.Line.Visible = $false
            $shape.TextFrame.TextRange.Text = $steps[$i]
            $shape.TextFrame.TextRange.Font.Color.RGB = 16777215
            $shape.TextFrame.TextRange.Font.Size = 12
            $shape.TextFrame.TextRange.Font.Bold = $true
            $shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
            $shapes += $shape.Name

            if ($i -lt $steps.Count - 1) {
                if ($direction -eq "horizontal") {
                    $arrowX1 = $x + $boxWidth
                    $arrowY1 = $y + $boxHeight / 2
                    $arrowX2 = $x + $boxWidth + $gap
                    $arrowY2 = $arrowY1
                } else {
                    $arrowX1 = $x + $boxWidth / 2
                    $arrowY1 = $y + $boxHeight
                    $arrowX2 = $arrowX1
                    $arrowY2 = $y + $boxHeight + $gap
                }
                $arrow = $slide.Shapes.AddLine($arrowX1, $arrowY1, $arrowX2, $arrowY2)
                $arrow.Line.EndArrowheadStyle = 2
                $arrow.Line.Weight = 2
                $arrowColor = Convert-HexColorToRgbInt([string]"#666666")
                if ($null -ne $arrowColor) { $arrow.Line.ForeColor.RGB = $arrowColor }
            }
        }

        Output-Json @{ success = $true; data = @{ steps = $steps.Count; shapes = $shapes } }
    }

    "createOrgChart" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $nodes = if ($p.nodes) { $p.nodes } else { @(@{ name = "CEO"; level = 0 }, @{ name = "技术总监"; level = 1 }, @{ name = "市场总监"; level = 1 }, @{ name = "财务总监"; level = 1 }) }
        $centerX = 360
        $startY = if ($p.startY) { $p.startY } else { 80 }
        $boxWidth = if ($p.boxWidth) { $p.boxWidth } else { 100 }
        $boxHeight = if ($p.boxHeight) { $p.boxHeight } else { 40 }
        $levelGap = if ($p.levelGap) { $p.levelGap } else { 80 }
        $levelColors = @("#1a365d", "#2d5a87", "#4a7ab0", "#6b9bd1")

        $levels = @{}
        foreach ($node in $nodes) {
            $level = if ($node.level -ne $null) { $node.level } else { 0 }
            if (-not $levels.ContainsKey($level)) { $levels[$level] = @() }
            $levels[$level] += $node
        }

        $levelKeys = $levels.Keys | Sort-Object
        $shapes = @()
        $levelIndex = 0
        foreach ($levelKey in $levelKeys) {
            $levelNodes = $levels[$levelKey]
            $levelY = $startY + $levelIndex * $levelGap
            $totalWidth = $levelNodes.Count * $boxWidth + ($levelNodes.Count - 1) * 30
            $startX = $centerX - $totalWidth / 2

            for ($i = 0; $i -lt $levelNodes.Count; $i++) {
                $x = $startX + $i * ($boxWidth + 30)
                $shape = $slide.Shapes.AddShape(5, $x, $levelY, $boxWidth, $boxHeight)
                $colorValue = Convert-HexColorToRgbInt([string]$levelColors[$levelIndex % $levelColors.Count])
                if ($null -ne $colorValue) { $shape.Fill.Solid(); $shape.Fill.ForeColor.RGB = $colorValue }
                $shape.Line.Visible = $false
                $shape.TextFrame.TextRange.Text = $levelNodes[$i].name
                $shape.TextFrame.TextRange.Font.Color.RGB = 16777215
                $shape.TextFrame.TextRange.Font.Size = 11
                $shape.TextFrame.TextRange.Font.Bold = $true
                $shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
                $shape.Shadow.Visible = $true
                $shape.Shadow.Blur = 5
                $shape.Shadow.OffsetX = 2
                $shape.Shadow.OffsetY = 2
                $shape.Shadow.Transparency = 0.7
                $shapes += $shape.Name
            }

            $levelIndex++
        }

        Output-Json @{ success = $true; data = @{ levels = $levelKeys.Count; totalNodes = $nodes.Count; shapes = $shapes } }
    }

    "createTimeline" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $events = if ($p.events) { $p.events } else { @(@{ date = "2020"; title = "成立" }, @{ date = "2021"; title = "融资" }, @{ date = "2022"; title = "扩张" }, @{ date = "2023"; title = "上市" }) }
        $startX = if ($p.startX) { $p.startX } else { 80 }
        $startY = if ($p.startY) { $p.startY } else { 280 }
        $lineLength = if ($p.lineLength) { $p.lineLength } else { 560 }

        $mainLine = $slide.Shapes.AddLine($startX, $startY, $startX + $lineLength, $startY)
        $lineColor = Convert-HexColorToRgbInt([string]$(if ($p.lineColor) { $p.lineColor } else { "#1a365d" }))
        if ($null -ne $lineColor) { $mainLine.Line.ForeColor.RGB = $lineColor }
        $mainLine.Line.Weight = 3

        $gap = if ($events.Count -gt 1) { $lineLength / ($events.Count - 1) } else { 0 }
        $shapes = @()

        for ($i = 0; $i -lt $events.Count; $i++) {
            $x = $startX + ($i * $gap)
            $above = ($i % 2 -eq 0)

            $dot = $slide.Shapes.AddShape(9, $x - 8, $startY - 8, 16, 16)
            $dotColor = Convert-HexColorToRgbInt([string]$(if ($p.dotColor) { $p.dotColor } else { "#0d47a1" }))
            if ($null -ne $dotColor) { $dot.Fill.Solid(); $dot.Fill.ForeColor.RGB = $dotColor }
            $dot.Line.Visible = $false

            $lineY = if ($above) { $startY - 50 } else { $startY + 50 }
            $connector = $slide.Shapes.AddLine($x, $startY, $x, $lineY)
            $connector.Line.Weight = 2
            $connectorColor = Convert-HexColorToRgbInt([string]"#cccccc")
            if ($null -ne $connectorColor) { $connector.Line.ForeColor.RGB = $connectorColor }

            $dateBox = $slide.Shapes.AddTextbox(1, $x - 30, $(if ($above) { $lineY - 45 } else { $lineY + 5 }), 60, 20)
            $dateBox.TextFrame.TextRange.Text = $events[$i].date
            $dateBox.TextFrame.TextRange.Font.Size = 12
            $dateBox.TextFrame.TextRange.Font.Bold = $true
            $dateBox.TextFrame.TextRange.ParagraphFormat.Alignment = 2

            $titleBox = $slide.Shapes.AddTextbox(1, $x - 40, $(if ($above) { $lineY - 25 } else { $lineY + 25 }), 80, 20)
            $titleBox.TextFrame.TextRange.Text = $events[$i].title
            $titleBox.TextFrame.TextRange.Font.Size = 10
            $titleColor = Convert-HexColorToRgbInt([string]"#666666")
            if ($null -ne $titleColor) { $titleBox.TextFrame.TextRange.Font.Color.RGB = $titleColor }
            $titleBox.TextFrame.TextRange.ParagraphFormat.Alignment = 2

            $shapes += $dot.Name
        }

        Output-Json @{ success = $true; data = @{ events = $events.Count; shapes = $shapes } }
    }

    "getSlideMaster" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $master = $pres.SlideMaster
        $masters = @()
        for ($i = 1; $i -le $master.Shapes.Count; $i++) {
            $shape = $master.Shapes.Item($i)
            $masters += @{ name = $shape.Name; type = $shape.Type; left = $shape.Left; top = $shape.Top; width = $shape.Width; height = $shape.Height }
        }
        Output-Json @{ success = $true; data = @{ shapeCount = $masters.Count; shapes = $masters } }
    }

    "setMasterBackground" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $master = $pres.SlideMaster
        if ($p.gradient) {
            $master.Background.Fill.TwoColorGradient(1, 1)
            $color1 = Convert-HexColorToRgbInt([string]$(if ($p.color1) { $p.color1 } else { "#1a365d" }))
            $color2 = Convert-HexColorToRgbInt([string]$(if ($p.color2) { $p.color2 } else { "#2d5a87" }))
            if ($null -ne $color1) { $master.Background.Fill.GradientStops.Item(1).Color.RGB = $color1 }
            if ($null -ne $color2) { $master.Background.Fill.GradientStops.Item(2).Color.RGB = $color2 }
        } else {
            $master.Background.Fill.Solid()
            $colorValue = Convert-HexColorToRgbInt([string]$(if ($p.color) { $p.color } else { "#ffffff" }))
            if ($null -ne $colorValue) { $master.Background.Fill.ForeColor.RGB = $colorValue }
        }
        Output-Json @{ success = $true; data = @{ message = "母版背景已更新" } }
    }

    "addMasterElement" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $master = $pres.SlideMaster

        $type = if ($p.type) { $p.type } else { "textbox" }
        $left = if ($p.left) { $p.left } else { 600 }
        $top = if ($p.top) { $p.top } else { 20 }
        $width = if ($p.width) { $p.width } else { 100 }
        $height = if ($p.height) { $p.height } else { 40 }

        if ($type -eq "textbox") {
            $shape = $master.Shapes.AddTextbox(1, $left, $top, $width, $height)
            $shape.TextFrame.TextRange.Text = if ($p.text) { $p.text } else { "Logo" }
            $shape.TextFrame.TextRange.Font.Size = if ($p.fontSize) { $p.fontSize } else { 14 }
            $shape.TextFrame.TextRange.Font.Bold = $true
        } elseif ($type -eq "shape") {
            $shape = $master.Shapes.AddShape($(if ($p.shapeType) { $p.shapeType } else { 5 }), $left, $top, $width, $height)
            $colorValue = Convert-HexColorToRgbInt([string]$(if ($p.color) { $p.color } else { "#1a365d" }))
            if ($null -ne $colorValue) { $shape.Fill.Solid(); $shape.Fill.ForeColor.RGB = $colorValue }
            $shape.Line.Visible = $false
        } else {
            Output-Json @{ success = $false; error = "Unsupported master element type" }
            break
        }

        Output-Json @{ success = $true; data = @{ name = $shape.Name; type = $type } }
    }

    "set3DRotation" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.shapeName) { $p.shapeName } else { $p.shapeIndex }))

        $shape.ThreeD.RotationX = if ($p.rotationX -ne $null) { $p.rotationX } else { 0 }
        $shape.ThreeD.RotationY = if ($p.rotationY -ne $null) { $p.rotationY } else { 0 }
        $shape.ThreeD.RotationZ = if ($p.rotationZ -ne $null) { $p.rotationZ } else { 0 }

        if ($p.preset) {
            $presets = @{ isometric = @{ x = 45; y = 45; z = 0 }; perspective = @{ x = 30; y = 30; z = 0 }; oblique = @{ x = 20; y = 60; z = 0 }; tiltLeft = @{ x = 0; y = -30; z = 0 }; tiltRight = @{ x = 0; y = 30; z = 0 } }
            $preset = $presets[$p.preset]
            if ($preset) {
                $shape.ThreeD.RotationX = $preset.x
                $shape.ThreeD.RotationY = $preset.y
                $shape.ThreeD.RotationZ = $preset.z
            }
        }

        Output-Json @{ success = $true; data = @{ shape = $shape.Name; rotationX = $shape.ThreeD.RotationX; rotationY = $shape.ThreeD.RotationY } }
    }

    "set3DDepth" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.shapeName) { $p.shapeName } else { $p.shapeIndex }))

        $depth = if ($p.depth) { $p.depth } else { 20 }
        $shape.ThreeD.Depth = $depth
        if ($p.depthColor) {
            $colorValue = Convert-HexColorToRgbInt([string]$p.depthColor)
            if ($null -ne $colorValue) { $shape.ThreeD.ExtrusionColor.RGB = $colorValue }
        }
        if ($p.lighting) {
            $lightingMap = @{ bright = 1; normal = 2; dim = 3; flat = 4 }
            $lightingValue = if ($lightingMap[$p.lighting]) { $lightingMap[$p.lighting] } else { 2 }
            $shape.ThreeD.PresetLighting = $lightingValue
        }

        Output-Json @{ success = $true; data = @{ shape = $shape.Name; depth = $depth } }
    }

    "set3DMaterial" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.shapeName) { $p.shapeName } else { $p.shapeIndex }))

        $material = if ($p.material) { $p.material } else { "plastic" }
        $materialMap = @{ matte = 1; plastic = 2; metal = 3; wireFrame = 4; glass = 5 }
        $materialValue = if ($materialMap[$material]) { $materialMap[$material] } else { 2 }
        $shape.ThreeD.PresetMaterial = $materialValue

        Output-Json @{ success = $true; data = @{ shape = $shape.Name; material = $material } }
    }

    "create3DText" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $text = if ($p.text) { $p.text } else { "3D文字" }
        $left = if ($p.left) { $p.left } else { 200 }
        $top = if ($p.top) { $p.top } else { 200 }
        $width = if ($p.width) { $p.width } else { 300 }
        $height = if ($p.height) { $p.height } else { 100 }
        $textBox = $slide.Shapes.AddTextbox(1, $left, $top, $width, $height)
        $textBox.TextFrame.TextRange.Text = $text
        $textBox.TextFrame.TextRange.Font.Size = if ($p.fontSize) { $p.fontSize } else { 48 }
        $textBox.TextFrame.TextRange.Font.Bold = $true
        $colorValue = Convert-HexColorToRgbInt([string]$(if ($p.color) { $p.color } else { "#1a365d" }))
        if ($null -ne $colorValue) { $textBox.TextFrame.TextRange.Font.Color.RGB = $colorValue }
        $textBox.TextFrame.TextRange.ParagraphFormat.Alignment = 2

        $textBox.ThreeD.RotationX = if ($p.rotationX -ne $null) { $p.rotationX } else { 15 }
        $textBox.ThreeD.RotationY = if ($p.rotationY -ne $null) { $p.rotationY } else { 30 }
        $textBox.ThreeD.Depth = if ($p.depth) { $p.depth } else { 10 }

        $textBox.Shadow.Visible = $true
        $textBox.Shadow.Blur = 8
        $textBox.Shadow.OffsetX = 5
        $textBox.Shadow.OffsetY = 5
        $textBox.Shadow.Transparency = 0.5

        Output-Json @{ success = $true; data = @{ name = $textBox.Name; text = $text } }
    }

    "setPptChartData" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.chartName) { $p.chartName } else { $p.chartIndex }))
        $chart = $shape.Chart
        $chart.ChartData.Activate()
        $dataSheet = $chart.ChartData.Workbook.Worksheets.Item(1)
        if ($p.data) {
            for ($r = 0; $r -lt $p.data.Count; $r++) {
                $rowData = $p.data[$r]
                for ($c = 0; $c -lt $rowData.Count; $c++) {
                    $dataSheet.Cells.Item($r + 1, $c + 1).Value2 = $rowData[$c]
                }
            }
        }
        Output-Json @{ success = $true; data = @{ chartName = $shape.Name } }
    }

    "setPptChartStyle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.chartName) { $p.chartName } else { $p.chartIndex }))
        $chart = $shape.Chart
        if ($p.title) {
            $chart.HasTitle = $true
            $chart.ChartTitle.Text = $p.title
        }
        if ($null -ne $p.hasLegend) {
            $chart.HasLegend = $p.hasLegend
        }
        Output-Json @{ success = $true; data = @{ chartName = $shape.Name } }
    }

    "removeAnimation" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $seq = $slide.TimeLine.MainSequence
        if ($p.index) {
            $seq.Item($p.index).Delete()
        } else {
            while ($seq.Count -gt 0) { $seq.Item(1).Delete() }
        }
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex } }
    }

    "getAnimations" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $seq = $slide.TimeLine.MainSequence
        $animations = @()
        for ($i = 1; $i -le $seq.Count; $i++) {
            $effect = $seq.Item($i)
            $animations += @{ index = $i; shapeName = $effect.Shape.Name; effectType = $effect.EffectType; duration = $effect.Timing.Duration }
        }
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; animations = $animations; count = $animations.Count } }
    }

    "setAnimationOrder" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $seq = $slide.TimeLine.MainSequence
        $effect = $seq.Item($p.from)
        $effect.MoveTo($p.to)
        Output-Json @{ success = $true; data = @{ from = $p.from; to = $p.to } }
    }

    "removeSlideTransition" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $slide.SlideShowTransition.EntryEffect = 0
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex } }
    }

    "applyTransitionToAll" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $transitionMap = @{ none = 0; fade = 1; push = 2; wipe = 3; split = 4; reveal = 5; random = 6 }
        $transition = if ($transitionMap[$p.transition]) { $transitionMap[$p.transition] } else { if ($p.transition) { $p.transition } else { 1 } }
        for ($i = 1; $i -le $pres.Slides.Count; $i++) {
            $slide = $pres.Slides.Item($i)
            $slide.SlideShowTransition.EntryEffect = $transition
            if ($p.duration) { $slide.SlideShowTransition.Duration = $p.duration }
        }
        Output-Json @{ success = $true; data = @{ transition = $p.transition; appliedTo = $pres.Slides.Count } }
    }

    "setBackgroundColor" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $colorValue = Convert-HexColorToRgbInt([string]$p.color)
        if ($p.applyToAll) {
            for ($i = 1; $i -le $pres.Slides.Count; $i++) {
                $slide = $pres.Slides.Item($i)
                $slide.FollowMasterBackground = $false
                $slide.Background.Fill.Solid()
                if ($null -ne $colorValue) { $slide.Background.Fill.ForeColor.RGB = $colorValue }
            }
            Output-Json @{ success = $true; data = @{ color = $p.color; appliedTo = $pres.Slides.Count } }
        } else {
            $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
            $slide = $pres.Slides.Item($slideIndex)
            $slide.FollowMasterBackground = $false
            $slide.Background.Fill.Solid()
            if ($null -ne $colorValue) { $slide.Background.Fill.ForeColor.RGB = $colorValue }
            Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; color = $p.color } }
        }
    }

    "setBackgroundImage" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $slide.FollowMasterBackground = $false
        $slide.Background.Fill.UserPicture($p.path)
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; path = $p.path } }
    }

    "addPptHyperlink" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.shapeName) { $p.shapeName } else { $p.shapeIndex }))
        $actionSettings = $shape.ActionSettings.Item(1)
        $actionSettings.Hyperlink.Address = if ($p.address) { $p.address } else { "" }
        if ($p.subAddress) { $actionSettings.Hyperlink.SubAddress = $p.subAddress }
        Output-Json @{ success = $true; data = @{ shapeName = $shape.Name; address = $p.address } }
    }

    "removePptHyperlink" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.shapeName) { $p.shapeName } else { $p.shapeIndex }))
        $shape.ActionSettings.Item(1).Hyperlink.Address = ""
        Output-Json @{ success = $true; data = @{ shapeName = $shape.Name } }
    }

    "setSlideNumber" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        if ($null -ne $p.visible) { $pres.SlideMaster.HeadersFooters.SlideNumber.Visible = $p.visible }
        Output-Json @{ success = $true; data = @{ visible = $p.visible } }
    }

    "setPptFooter" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $pres.SlideMaster.HeadersFooters.Footer.Visible = $true
        $pres.SlideMaster.HeadersFooters.Footer.Text = if ($p.text) { $p.text } else { "" }
        Output-Json @{ success = $true; data = @{ text = $p.text } }
    }

    "setPptDateTime" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $pres.SlideMaster.HeadersFooters.DateAndTime.Visible = if ($p.visible -ne $null) { $p.visible } else { $true }
        if ($p.useFixed) {
            $pres.SlideMaster.HeadersFooters.DateAndTime.UseFormat = $false
            $pres.SlideMaster.HeadersFooters.DateAndTime.Text = if ($p.text) { $p.text } else { "" }
        }
        Output-Json @{ success = $true; data = @{ visible = if ($p.visible -ne $null) { $p.visible } else { $true } } }
    }

    "findPptText" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $results = @()
        for ($i = 1; $i -le $pres.Slides.Count; $i++) {
            $slide = $pres.Slides.Item($i)
            for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
                $shape = $slide.Shapes.Item($j)
                try {
                    if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                        $text = $shape.TextFrame.TextRange.Text
                        if ($text -and $text.Contains($p.text)) {
                            $results += @{ slideIndex = $i; shapeName = $shape.Name; text = $text }
                        }
                    }
                } catch {}
            }
        }
        Output-Json @{ success = $true; data = @{ searchText = $p.text; results = $results; count = $results.Count } }
    }

    "replacePptText" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $count = 0
        for ($i = 1; $i -le $pres.Slides.Count; $i++) {
            $slide = $pres.Slides.Item($i)
            for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
                $shape = $slide.Shapes.Item($j)
                try {
                    if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                        $text = $shape.TextFrame.TextRange.Text
                        if ($text -and $text.Contains($p.findText)) {
                            $shape.TextFrame.TextRange.Text = $text.Replace($p.findText, $(if ($p.replaceText) { $p.replaceText } else { "" }))
                            $count++
                        }
                    }
                } catch {}
            }
        }
        Output-Json @{ success = $true; data = @{ findText = $p.findText; replaceText = $p.replaceText; count = $count } }
    }

    "startSlideShow" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $pres.SlideShowSettings.Run()
        Output-Json @{ success = $true }
    }

    "endSlideShow" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        if ($ppt.SlideShowWindows.Count -gt 0) { $ppt.SlideShowWindows.Item(1).View.Exit() }
        Output-Json @{ success = $true }
    }

    "unifyFont" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $fontName = if ($p.fontName) { $p.fontName } else { "微软雅黑" }
        $count = 0
        for ($i = 1; $i -le $pres.Slides.Count; $i++) {
            $slide = $pres.Slides.Item($i)
            for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
                $shape = $slide.Shapes.Item($j)
                try {
                    if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                        $shape.TextFrame.TextRange.Font.Name = $fontName
                        $count++
                    }
                } catch {}
            }
        }
        Output-Json @{ success = $true; data = @{ fontName = $fontName; count = $count } }
    }

    "slide.beautify" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $schemes = @{
            business = @{ title = 0x2F5496; body = 0x333333 }
            tech = @{ title = 0x00B0F0; body = 0x404040 }
            creative = @{ title = 0xFF6B6B; body = 0x4A4A4A }
            minimal = @{ title = 0x000000; body = 0x666666 }
        }
        $schemeKey = if ($p.style -and $p.style.colorScheme) { $p.style.colorScheme } elseif ($p.style) { $p.style } else { "business" }
        $scheme = $schemes[$schemeKey]
        if ($null -eq $scheme) { $scheme = $schemes["business"] }
        $fontName = if ($p.style -and $p.style.font) { $p.style.font } else { $null }
        $slideIndex = $p.slideIndex
        $slides = @()
        if ($slideIndex -eq "all") {
            for ($i = 1; $i -le $pres.Slides.Count; $i++) { $slides += $pres.Slides.Item($i) }
        } else {
            $target = if ($slideIndex) { [int]$slideIndex } else { $ppt.ActiveWindow.Selection.SlideRange.SlideIndex }
            $slides += $pres.Slides.Item($target)
        }
        $ops = @(
            @{ operation = "unify_font"; count = 0 },
            @{ operation = "apply_color_scheme"; count = 0 },
            @{ operation = "align"; count = 0 },
            @{ operation = "optimize_spacing"; count = 0 }
        )
        foreach ($slide in $slides) {
            for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
                $shape = $slide.Shapes.Item($j)
                try {
                    if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                        $textRange = $shape.TextFrame.TextRange
                        if ($textRange.Font.Size -ge 24) { $textRange.Font.Color.RGB = $scheme.title }
                        else { $textRange.Font.Color.RGB = $scheme.body }
                        $ops[1].count++
                        if ($fontName) { $textRange.Font.Name = $fontName; $ops[0].count++ }
                    }
                } catch {}
            }
        }
        $slideIndexOut = if ($slideIndex -eq "all") { "all" } else { ($slides | Select-Object -First 1).SlideIndex }
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndexOut; operations = $ops } }
    }

    "beautifySlide" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { $ppt.ActiveWindow.Selection.SlideRange.SlideIndex }
        $slide = $pres.Slides.Item($slideIndex)
        $schemes = @{
            business = @{ title = 0x2F5496; body = 0x333333 }
            tech = @{ title = 0x00B0F0; body = 0x404040 }
            creative = @{ title = 0xFF6B6B; body = 0x4A4A4A }
            minimal = @{ title = 0x000000; body = 0x666666 }
        }
        $scheme = $schemes[$p.style]
        if ($null -eq $scheme) { $scheme = $schemes["business"] }
        $count = 0
        for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
            $shape = $slide.Shapes.Item($j)
            try {
                if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                    $textRange = $shape.TextFrame.TextRange
                    if ($textRange.Font.Size -ge 24) { $textRange.Font.Color.RGB = $scheme.title }
                    else { $textRange.Font.Color.RGB = $scheme.body }
                    $count++
                }
            } catch {}
        }
        Output-Json @{ success = $true; data = @{ style = $p.style; count = $count } }
    }

    default {
        Output-Json @{ success = $false; error = "Unknown action: $Action" }
    }
}
