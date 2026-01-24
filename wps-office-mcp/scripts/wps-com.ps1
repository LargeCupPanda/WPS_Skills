# WPS COM Bridge - PowerShell script for WPS COM operations
# Usage: powershell -File wps-com.ps1 -Action <action> -Params <json>

param(
    [string]$Action,
    [string]$Params = "{}"
)

$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Get-WpsExcel {
    try {
        return [System.Runtime.InteropServices.Marshal]::GetActiveObject('Ket.Application')
    } catch {
        return $null
    }
}

function Get-WpsWord {
    try {
        return [System.Runtime.InteropServices.Marshal]::GetActiveObject('Kwps.Application')
    } catch {
        return $null
    }
}

function Get-WpsPpt {
    try {
        return [System.Runtime.InteropServices.Marshal]::GetActiveObject('Kwpp.Application')
    } catch {
        return $null
    }
}

function Output-Json($obj) {
    $obj | ConvertTo-Json -Depth 10 -Compress
}

try {
    $p = $Params | ConvertFrom-Json
} catch {
    $p = @{}
}

switch ($Action) {
    "ping" {
        Output-Json @{ success = $true; data = @{ message = "pong"; timestamp = [DateTimeOffset]::Now.ToUnixTimeMilliseconds() } }
    }

    "getActiveWorkbook" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) {
            Output-Json @{ success = $false; error = "WPS Excel not running" }
            exit
        }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) {
            Output-Json @{ success = $false; error = "No active workbook" }
            exit
        }
        $sheets = @()
        for ($i = 1; $i -le $wb.Sheets.Count; $i++) {
            $sheets += $wb.Sheets.Item($i).Name
        }
        Output-Json @{
            success = $true
            data = @{
                name = $wb.Name
                path = $wb.FullName
                sheetCount = $wb.Sheets.Count
                sheets = $sheets
            }
        }
    }

    "getCellValue" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) {
            Output-Json @{ success = $false; error = "WPS Excel not running" }
            exit
        }
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) {
            Output-Json @{ success = $false; error = "No active workbook" }
            exit
        }
        $sheet = if ($p.sheet -is [int]) { $wb.Sheets.Item($p.sheet) } else { $wb.Sheets.Item($p.sheet) }
        $cell = $sheet.Cells.Item($p.row, $p.col)
        Output-Json @{
            success = $true
            data = @{
                value = $cell.Value2
                text = $cell.Text
                formula = $cell.Formula
            }
        }
    }

    "setCellValue" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) {
            Output-Json @{ success = $false; error = "WPS Excel not running" }
            exit
        }
        $wb = $excel.ActiveWorkbook
        $sheet = if ($p.sheet -is [int]) { $wb.Sheets.Item($p.sheet) } else { $wb.Sheets.Item($p.sheet) }
        $sheet.Cells.Item($p.row, $p.col).Value2 = $p.value
        Output-Json @{ success = $true }
    }

    "getRangeData" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) {
            Output-Json @{ success = $false; error = "WPS Excel not running" }
            exit
        }
        $wb = $excel.ActiveWorkbook
        $sheet = if ($p.sheet -is [int]) { $wb.Sheets.Item($p.sheet) } else { $wb.Sheets.Item($p.sheet) }
        $range = $sheet.Range($p.range)
        $data = @()
        for ($r = 1; $r -le $range.Rows.Count; $r++) {
            $row = @()
            for ($c = 1; $c -le $range.Columns.Count; $c++) {
                $row += $range.Cells.Item($r, $c).Value2
            }
            $data += ,@($row)
        }
        Output-Json @{ success = $true; data = @{ data = $data } }
    }

    "setRangeData" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) {
            Output-Json @{ success = $false; error = "WPS Excel not running" }
            exit
        }
        $wb = $excel.ActiveWorkbook
        $sheet = if ($p.sheet -is [int]) { $wb.Sheets.Item($p.sheet) } else { $wb.Sheets.Item($p.sheet) }
        $range = $sheet.Range($p.range)
        for ($r = 0; $r -lt $p.data.Count; $r++) {
            for ($c = 0; $c -lt $p.data[$r].Count; $c++) {
                $range.Cells.Item($r + 1, $c + 1).Value2 = $p.data[$r][$c]
            }
        }
        Output-Json @{ success = $true }
    }

    "setFormula" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) {
            Output-Json @{ success = $false; error = "WPS Excel not running" }
            exit
        }
        $wb = $excel.ActiveWorkbook
        $sheet = if ($p.sheet -is [int]) { $wb.Sheets.Item($p.sheet) } else { $wb.Sheets.Item($p.sheet) }
        $sheet.Cells.Item($p.row, $p.col).Formula = $p.formula
        Output-Json @{ success = $true }
    }

    "save" {
        $excel = Get-WpsExcel
        if ($null -ne $excel -and $null -ne $excel.ActiveWorkbook) {
            $excel.ActiveWorkbook.Save()
            Output-Json @{ success = $true }
            exit
        }
        $word = Get-WpsWord
        if ($null -ne $word -and $null -ne $word.ActiveDocument) {
            $word.ActiveDocument.Save()
            Output-Json @{ success = $true }
            exit
        }
        Output-Json @{ success = $false; error = "No active document" }
    }

    default {
        Output-Json @{ success = $false; error = "Unknown action: $Action" }
    }
}
