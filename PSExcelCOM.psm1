<#
Functions for working with Excel COM object
#>

# Set column width 
# usage: PSSetColumnWidth -sheet sheet1 -column "A" -width 10
function PSSetColumnWidth($sheet, $column, $width) {
        $sheetWS = $WB.Worksheets.item("$sheet")
        $sheetWS.Columns("$column").ColumnWidth = $width
}

# Set valignment
# usage: PSSetVAlign -sheet sheet1 -column "A" -width '-4160'
function PSSetVAlign($sheet, $range, $Alignment) {
    $sheetWS = $WB.Worksheets.item("$sheet")
    $selection = $sheetWS.Range("$range")
    [void]$selection.select()
    $selection.Style.VerticalAlignment = $Alignment
}

# Set row height 
# usage: PSSetRowHeight -sheet sheet1 -rows "1" -height 10
function PSSetRowHeight($sheet, $rows, $height) {
    $sheetWS = $WB.Worksheets.item($sheet)
    $selection = $sheetWS.Rows("$rows")
    [void]$selection.select()
    $selection.RowHeight = "$height"
    $selection = $null
}

Function PSSaveWorkbook($file) {
    $WB.SaveAs($file)
    Write-host "Invoice excel sheet is done.." -ForegroundColor Yellow
}

# Set borderaround area 
# usage: PSSetBorderAround -sheet sheet1 -range "A1:b3" -criteria 1,1,1 (linestyle,weight,colorindex)
function PSSetBorderAround($sheet, $range, $linestyle, $weight, $Colorindex) {
    $sheetWS = $WB.Worksheets.item("$sheet")
    $selection = $sheetWS.Range($range)
    [void]$selection.select()
    $selection.BorderAround($linestyle,$weight,$Colorindex)
    $selection = $null
}

# Set background color on area 
# usage: PSSetbackgroundcolor -sheet sheet1 -range "a1:b3" -color 15
function PSSetbackgroundcolor($sheet, $range, $color) {
    $sheetWS = $WB.Worksheets.item("$sheet")
    $selection = $sheetWS.Range("$range")
    [void]$selection.select()
    $selection.Interior.ColorIndex = "$color"
    $selection = $null
}

# Set font on area 
# usage: PSSetfont -sheet sheet1 -range "a1:b3" -fname 'arial' -fsize 12 -color 0
function PSSetfont($sheet, $range, $fname, $fsize, $color) {
    $sheetWS = $WB.Worksheets.item("$sheet")
    $selection = $sheetWS.Range("$range")
    [void]$selection.select()
    $selection.Font.Name = "$fname"
    $selection.Font.Size = "$fsize"
    $selection.Font.ColorIndex = "$color"
}

# Set font Bold
# usage: PSSetfontBold -sheet sheet1 -range "a1:b3"
function PSSetFontBold($sheet, $range) {
    $sheetWS = $WB.Worksheets.item($sheet)
    $selection = $sheetWS.Range($range)
    [void]$selection.select()
    $selection.font.bold = $True
    $selection = $null
    
}

# Set cell content
# usage: PSSetcellcontent -sheet sheet1 -column 1 -row 2 -content "hello world"
function PSSetCellContent($sheet, $column, $row, $content) {
    $sheetWS = $WB.Worksheets.item($sheet)
    #$selection = $sheetWS.Range($range)
    #[void]$selection.select()
    $sheetWS.cells.item($column,$row) = $content
    
}

# Set  Active sheet
# usage: PSActiveSheet -sheet sheet1
function PSActiveSheet($sheet) {
    $active = $WB.ActiveSheet.name
    if ($active -ne $sheet) {
        Write-host "`nSheet active: $active, " -NoNewline -ForegroundColor Red
        Write-host "changing to $sheet`n" -ForegroundColor Yellow
        $sheetWS = $WB.Worksheets.item("$sheet")
        $sheetWS.activate()
    }else{
        Write-host "`nSheet active: $active`n" -ForegroundColor Green
    }
}

function PSMergecells($sheet, $range) {
    $sheetWS = $WB.Worksheets.item($sheet)
    $selection = $sheetWS.Range($range)
    [void]$selection.select()
    $selection.MergeCells = $True
}

