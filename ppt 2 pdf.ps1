# Define the folder path
$folderPath = "D:\0000 study spacd\05 SEM 6(winter)\01 IOT (A1)\01 CAT 1"

# Get all PPT files in the folder
$pptFiles = Get-ChildItem -Path $folderPath -Filter *.ppt*

# Create a COM object for PowerPoint
$powerPoint = New-Object -ComObject PowerPoint.Application
$powerPoint.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

# Convert each file
foreach ($pptFile in $pptFiles) {
    $presentation = $powerPoint.Presentations.Open($pptFile.FullName, $false, $false, $false)
    $pdfPath = [System.IO.Path]::ChangeExtension($pptFile.FullName, ".pdf")
    $presentation.SaveAs($pdfPath, [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF)
    $presentation.Close()
}

# Quit PowerPoint
$powerPoint.Quit()
Write-Host "Conversion completed!"
