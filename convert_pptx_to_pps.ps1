# Batch convert all .ppt/.pptx files encountered in folder and all its subfolders
# The produced PDF files are stored in the invocation folder
#
# Adapted from https://gist.github.com/mp4096/1a2279ec7b3dfec659f58e378ddd9aee
#
# If PowerShell exits with an error, check if unsigned scripts are allowed in your system.
# You can allow them by calling PowerShell as an Administrator and typing
# ```
# Set-ExecutionPolicy Unrestricted
# ```
# Get invocation path
$curr_path = Split-Path -parent $MyInvocation.MyCommand.Path
# Create a PowerPoint object
$ppt_app = New-Object -ComObject PowerPoint.Application
# Get all objects of type .ppt? in $curr_path and its subfolders
Get-ChildItem -Path $curr_path -Recurse -Filter *.ppt? | ForEach-Object {
    Write-Host "Processing" $_.FullName "..."
    # Open it in PowerPoint
    $document = $ppt_app.Presentations.Open($_.FullName)
    # Create a name for the PDF document; they are stored in the invocation folder!
    # If you want them to be created locally in the folders containing the source PowerPoint file, replace $curr_path with $_.DirectoryName
    $ppsx_filename = "$($curr_path)\$($_.BaseName).ppsx"
    # Save as PDF -- 17 is the literal value of `wdFormatPDF`
    $opt= [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsOpenXMLShow # https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
    $document.SaveAs($ppsx_filename, $opt)
    # Close PowerPoint file
    $document.Close()
}
# Exit and release the PowerPoint object
$ppt_app.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt_app)
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()