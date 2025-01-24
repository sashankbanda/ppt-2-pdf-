# PowerPoint to PDF Conversion Script

This project provides a PowerShell script to convert all PowerPoint (.ppt or .pptx) files in a specified folder to PDF format. The script uses Microsoft PowerPoint's COM object to perform the conversion.

## Prerequisites

- **Windows OS** with PowerShell installed.
- **Microsoft Office PowerPoint** installed on your system.
- Ensure PowerPoint's security settings allow script automation.

## Installation

1. Clone this repository or download the script file `ConvertPPTtoPDF.ps1`.
2. Save the script file to your preferred location.

## Usage

1. Open PowerShell as Administrator.
2. Navigate to the directory containing the `ConvertPPTtoPDF.ps1` script.
3. Run the script using the following command:

   ```powershell
   .\ConvertPPTtoPDF.ps1
   ```

4. Enter the folder path containing the PowerPoint files when prompted or modify the script to include your desired folder path directly.

## Script Details

### Example Script

```powershell
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
```

### Output
- Converted PDF files will be saved in the same directory as the original `.ppt` or `.pptx` files.

## Notes

- Make sure the folder path is correct and accessible.
- If you encounter any permission issues, try running PowerShell as Administrator.
- The script overwrites existing PDF files with the same name, so back up your data if necessary.

## License

This project is licensed under the [MIT License](LICENSE).

## Contributions

Contributions are welcome! Feel free to submit a pull request or open an issue for any improvements or bug fixes.

---

### Author
Made with ❤️ Sashank banda.

---

### References
- [Microsoft PowerPoint COM Documentation](https://learn.microsoft.com/en-us/office/vba/api/overview/powerpoint)
