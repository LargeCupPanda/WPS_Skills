# WPS COM Bridge - PowerShell script for WPS COM operations
# Full implementation for Excel, Word, and PPT
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

try { $p = $Params | ConvertFrom-Json } catch { $p = @{} }

switch ($Action) {

    # ==================== Common ====================
    "ping" {
        Output-Json @{ success = $true; data = @{ message = "pong"; timestamp = [DateTimeOffset]::Now.ToUnixTimeMilliseconds() } }
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

    "setFormula" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $wb = $excel.ActiveWorkbook
        $sheet = if ($p.sheet -is [int]) { $wb.Sheets.Item($p.sheet) } else { $wb.Sheets.Item($p.sheet) }
        $sheet.Cells.Item($p.row, $p.col).Formula = $p.formula
        Output-Json @{ success = $true }
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
            selectedCell = $excel.Selection.Address(); headers = $headers; usedRange = $usedRange.Address()
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
        $chartTypes = @{ column = 51; bar = 57; line = 4; pie = 5; area = 1; scatter = -4169 }
        $chartType = $chartTypes[$p.chartType]
        if ($null -eq $chartType) { $chartType = 51 }
        $left = if ($p.left) { $p.left } else { $range.Left + $range.Width + 20 }
        $top = if ($p.top) { $p.top } else { $range.Top }
        $chartObj = $sheet.ChartObjects().Add($left, $top, 400, 300)
        $chartObj.Chart.SetSourceData($range)
        $chartObj.Chart.ChartType = $chartType
        if ($p.title) { $chartObj.Chart.HasTitle = $true; $chartObj.Chart.ChartTitle.Text = $p.title }
        Output-Json @{ success = $true; data = @{ chartName = $chartObj.Name } }
    }

    "removeDuplicates" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; exit }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $cols = if ($p.columns) { $p.columns } else { @(1) }
        $range.RemoveDuplicates($cols, 1)
        Output-Json @{ success = $true }
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

    "getDocumentText" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        $text = $doc.Content.Text
        if ($text.Length -gt 10000) { $text = $text.Substring(0, 10000) + "...(truncated)" }
        Output-Json @{ success = $true; data = @{ text = $text; length = $doc.Content.Text.Length } }
    }

    "insertText" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; exit }
        switch ($p.position) {
            "start" { $range = $doc.Range(0, 0); $range.InsertBefore($p.text) }
            "end" { $range = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1); $range.InsertAfter($p.text) }
            default { $word.Selection.TypeText($p.text) }
        }
        Output-Json @{ success = $true }
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
        Output-Json @{ success = $true }
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
        $replaceType = if ($p.replaceAll) { 2 } else { 1 }
        $result = $find.Execute($p.findText, $false, $false, $false, $false, $false, $true, 1, $false, $p.replaceText, $replaceType)
        Output-Json @{ success = $true; data = @{ replaced = $result } }
    }

    "insertTable" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $doc = $word.ActiveDocument
        $range = $word.Selection.Range
        $table = $doc.Tables.Add($range, $p.rows, $p.cols)
        if ($p.data) {
            for ($r = 0; $r -lt [Math]::Min($p.data.Count, $p.rows); $r++) {
                for ($c = 0; $c -lt [Math]::Min($p.data[$r].Count, $p.cols); $c++) {
                    $table.Cell($r + 1, $c + 1).Range.Text = [string]$p.data[$r][$c]
                }
            }
        }
        $table.Borders.Enable = $true
        Output-Json @{ success = $true }
    }

    "applyStyle" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; exit }
        $range = $word.Selection.Range
        $range.Style = $p.styleName
        Output-Json @{ success = $true }
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

    "addSlide" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $layouts = @{ title = 1; title_content = 2; blank = 12; two_column = 3 }
        $layoutType = $layouts[$p.layout]
        if ($null -eq $layoutType) { $layoutType = 2 }
        $position = if ($p.position) { $p.position } else { $pres.Slides.Count + 1 }
        $slide = $pres.Slides.Add($position, $layoutType)
        if ($p.title -and $slide.Shapes.HasTitle) { $slide.Shapes.Title.TextFrame.TextRange.Text = $p.title }
        Output-Json @{ success = $true; data = @{ slideIndex = $position } }
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

    "setSlideTitle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; exit }
        $pres = $ppt.ActivePresentation
        $slide = $pres.Slides.Item($p.slideIndex)
        if ($slide.Shapes.HasTitle) { $slide.Shapes.Title.TextFrame.TextRange.Text = $p.title }
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
