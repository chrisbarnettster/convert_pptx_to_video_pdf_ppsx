# Batch convert all .ppt/.pptx files encountered in folder and all its subfolders
# The produced PDF files are stored in the invocation folder
#
# Adapted from http://stackoverflow.com/questions/16534292/basic-powershell-batch-convert-word-docx-to-pdf
# Thanks to MFT, takabanana, ComFreek
#
# If PowerShell exits with an error, check if unsigned scripts are allowed in your system.
# You can allow them by calling PowerShell as an Administrator and typing
# ```
# Set-ExecutionPolicy Unrestricted
# ```
# Get invocation path
$curr_path = Split-Path -parent $MyInvocation.MyCommand.Path
# Create a Word object
$wrd_app = New-Object -ComObject Word.Application
# Get all objects of type .doc? in $curr_path and its subfolders
Get-ChildItem -Path $curr_path -Recurse -Filter *.doc? | ForEach-Object {
    Write-Host "Processing" $_.FullName "..."
    # Open it in Word
    $document = $wrd_app.Documents.Open($_.FullName)
    # Create a name for the PDF document; they are stored in the invocation folder!
    # If you want them to be created locally in the folders containing the source Word file, replace $curr_path with $_.DirectoryName
    $pdf_filename = "$($curr_path)\$($_.BaseName).pdf"
    # Save as PDF -- 17 is the literal value of `wdFormatPDF`
    $opt= [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatPDF
    $document.SaveAs($pdf_filename, $opt)
    # Close file
    $document.Close()
}
# Exit and release the PowerPoint object
$wrd_app.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wrd_app)
