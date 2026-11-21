# Master LMS Submission Package Builder
$ErrorActionPreference = "Stop"

$baseDir = "c:\Users\deeps\Downloads\Appointment No Show Analysis"
$docsDir = Join-Path $baseDir "docs"
$outDir = Join-Path $baseDir "LMS_Submission_Documents"

if (Test-Path $outDir) {
    Remove-Item -Path $outDir -Recurse -Force
}
New-Item -ItemType Directory -Path $outDir -Force | Out-Null
Write-Output "Created LMS target directory: $outDir"

# 1. Initialize Word COM with ExportAsFixedFormat
Write-Output "Starting Microsoft Word COM automation..."
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

$wordDocs = @(
    @{ Src = "01_Project_Description.docx"; Out = "01_2500320140047_Deepshikha_Project_Understanding.pdf" },
    @{ Src = "02_HMW_Statements.docx"; Out = "02_2500320140047_Deepshikha_HMW.pdf" },
    @{ Src = "03_Needs_Analysis.docx"; Out = "03_2500320140047_Deepshikha_Needs_Analysis.pdf" },
    @{ Src = "05_BRD.docx"; Out = "05_2500320140047_Deepshikha_BRD.pdf" },
    @{ Src = "07_TRD.docx"; Out = "07_2500320140047_Deepshikha_TRD.pdf" },
    @{ Src = "09_ER_Diagram.docx"; Out = "09_2500320140047_Deepshikha_ER_Diagram.pdf" },
    @{ Src = "10_Dataset_Specification.docx"; Out = "10_2500320140047_Deepshikha_Data_Validation_Report.pdf" },
    @{ Src = "11_User_Stories.docx"; Out = "11_2500320140047_Deepshikha_User_Stories.pdf" },
    @{ Src = "12_UI_Framework.docx"; Out = "12_2500320140047_Deepshikha_UI_Framework.pdf" },
    @{ Src = "13_Test_Cases.docx"; Out = "13_2500320140047_Deepshikha_Test_Cases.pdf" },
    @{ Src = "2500320140047_Project_Document.docx"; Out = "14_2500320140047_Deepshikha_Project_Report.pdf" }
)

foreach ($item in $wordDocs) {
    [string]$srcPath = [System.IO.Path]::Combine($docsDir, $item.Src)
    [string]$targetPdf = [System.IO.Path]::Combine($outDir, $item.Out)
    Write-Output "Converting Word -> PDF: $($item.Src) -> $($item.Out)"
    $doc = $word.Documents.Open($srcPath, $false, $true)
    $doc.ExportAsFixedFormat($targetPdf, 17) # wdExportFormatPDF = 17
    $doc.Close(0) # wdDoNotSaveChanges = 0
}
$word.Quit()
Write-Output "Word conversions successfully finished."

# 2. Initialize PowerPoint COM
Write-Output "Starting Microsoft PowerPoint COM automation..."
$ppt = New-Object -ComObject PowerPoint.Application
[string]$pptSrc = [System.IO.Path]::Combine($docsDir, "2500320140047_Presentation.pptx")
[string]$pptTarget = [System.IO.Path]::Combine($outDir, "15_2500320140047_Deepshikha_Presentation.pdf")
Write-Output "Converting PPT -> PDF: 2500320140047_Presentation.pptx -> 15_2500320140047_Deepshikha_Presentation.pdf"
$pres = $ppt.Presentations.Open($pptSrc, [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
$pres.SaveAs($pptTarget, 32) # ppSaveAsPDF = 32
$pres.Close()
$ppt.Quit()
Write-Output "PowerPoint conversion successfully finished."

# 3. Copy Excel Files (kept as XLSX as specified in LMS guidelines)
Write-Output "Copying Excel files..."
Copy-Item (Join-Path $docsDir "04_Traceability_Matrix.xlsx") (Join-Path $outDir "04_2500320140047_Deepshikha_Traceability_Matrix.xlsx") -Force
Copy-Item (Join-Path $docsDir "08_Database_Design.xlsx") (Join-Path $outDir "08_2500320140047_Deepshikha_Database_Design.xlsx") -Force

# 4. Copy Source Code PDF & GitHub Details PDF
Write-Output "Copying Source Code & GitHub Details PDFs..."
Copy-Item (Join-Path $docsDir "2500320140047_Source_Code.pdf") (Join-Path $outDir "16_2500320140047_Deepshikha_Source_Code.pdf") -Force
Copy-Item (Join-Path $docsDir "2500320140047_GitHub_Streamlit_Details.pdf") (Join-Path $outDir "17_2500320140047_Deepshikha_GitHub_Streamlit_Details.pdf") -Force

# 5. Create Master LMS Submission ZIP archive
$zipAllDocs = Join-Path $baseDir "2500320140047_Deepshikha_LMS_Complete_Submission.zip"
if (Test-Path $zipAllDocs) { Remove-Item $zipAllDocs -Force }
Write-Output "Creating Complete LMS Documents ZIP: $zipAllDocs"
Compress-Archive -Path "$outDir\*" -DestinationPath $zipAllDocs -Force

# 6. Create Codebase ZIP (excluding venv, git, pycache)
$zipCodebase = Join-Path $baseDir "2500320140047_Deepshikha_Project_Codebase.zip"
if (Test-Path $zipCodebase) { Remove-Item $zipCodebase -Force }
Write-Output "Creating Project Codebase ZIP: $zipCodebase"

$codeFiles = Get-ChildItem -Path $baseDir -Exclude ".git", ".venv", "venv", "__pycache__", ".pytest_cache", "LMS_Submission_Documents", "*.zip"
Compress-Archive -Path $codeFiles.FullName -DestinationPath $zipCodebase -Force

Write-Output "=========================================================================="
Write-Output "LMS SUBMISSION PACKAGE GENERATED SUCCESSFULLY"
Write-Output "=========================================================================="
Write-Output "Destination Directory: $outDir"
Get-ChildItem -Path $outDir | Select-Object Name, @{Name="Size (KB)"; Expression={[math]::Round($_.Length / 1KB, 1)}} | Format-Table -AutoSize

Write-Output "ZIP Archives:"
Get-Item $zipAllDocs, $zipCodebase | Select-Object Name, @{Name="Size (MB)"; Expression={[math]::Round($_.Length / 1MB, 2)}} | Format-Table -AutoSize
