param (
    [Parameter(Mandatory=$true, Position=0)]
    [string]$folder
)

if (-not (Test-Path $folder)) {
    Write-Error "The folder path '$folder' does not exist."
    exit 1
}

# --- Word to PDF ---
$word = New-Object -ComObject Word.Application
$word.Visible = $false

Get-ChildItem -Path $folder -Filter *.docx | ForEach-Object {
    Write-Host "Converting Word: $($_.FullName)"
    $doc = $word.Documents.Open($_.FullName)
    $pdfPath = "$($_.DirectoryName)\$($_.BaseName).pdf"
    $doc.SaveAs([ref] $pdfPath, [ref] 17)  # 17 = wdFormatPDF
    $doc.Close()
}
$word.Quit()

# --- PowerPoint to PDF ---
$powerpoint = New-Object -ComObject PowerPoint.Application
# $powerpoint.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse

Get-ChildItem -Path $folder -Filter *.pptx | ForEach-Object {
    Write-Host "Converting PowerPoint: $($_.FullName)"
    $presentation = $powerpoint.Presentations.Open($_.FullName, $false, $false, $false)
    $pdfPath = "$($_.DirectoryName)\$($_.BaseName).pdf"
    $presentation.SaveAs($pdfPath, 32)  # 32 = ppSaveAsPDF
    $presentation.Close()
}
$powerpoint.Quit()

# --- Excel to PDF ---
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

Get-ChildItem -Path $folder -Filter *.xlsx | ForEach-Object {
    Write-Host "Converting Excel: $($_.FullName)"
    $workbook = $excel.Workbooks.Open($_.FullName)
    $pdfPath = "$($_.DirectoryName)\$($_.BaseName).pdf"
    try {
        $workbook.ExportAsFixedFormat(0, $pdfPath)  # 0 = xlTypePDF
    } catch {
        Write-Warning "Failed to convert $($_.FullName): $_"
    }
    $workbook.Close($false)
}

$excel.Quit()
