# Batch convert all .ppt/.pptx files encountered in folder and all its subfolders
# The produced MP4 files are stored in the invocation folder
#
# Adapted from https://gist.github.com/mp4096/1a2279ec7b3dfec659f58e378ddd9aee
#
# If PowerShell exits with an error, check if unsigned scripts are allowed in your system.
# You can allow them by calling PowerShell as an Administrator and typing
# ```
# Set-ExecutionPolicy Unrestricted
# ```
# Make sure PowerPoint is closed first or video names could be corrupted, ask the user to confirm. https://stackoverflow.com/questions/40889444/do-until-user-input-yes-no#:~:text=In%20PowerShell%20you%20have%20basically,for%20a%20yes%2Fno%20choice.&text=Use%20%2Dlike%20'y*'%20and,trailing%20characters%20in%20the%20response.
$title   = 'Close PowerPoint'
$msg     = 'Have you closed all PowerPoint Presentations?'
$options = '&Yes', '&No'
$default = 1  # 0=Yes, 1=No

do {
    $response = $Host.UI.PromptForChoice($title, $msg, $options, $default)
    if ($response -eq 1) {
        # Do nothing , code will resume after this
    }
} until ($response -eq 0)

# Get invocation path
$curr_path = Split-Path -parent $MyInvocation.MyCommand.Path
# Create a PowerPoint object
$ppt_app = New-Object -ComObject PowerPoint.Application
# Get all objects of type .ppt? in $curr_path and its subfolders
Get-ChildItem -Path $curr_path -Recurse -Filter *.ppt? | ForEach-Object {
    Write-Host "Processing" $_.FullName "..."
    # Open it in PowerPoint
    $document = $ppt_app.Presentations.Open($_.FullName)
    #$document.CreateVideoStatus
    # Create a name for the MP4 document; they are stored in the invocation folder!
    # If you want them to be created locally in the folders containing the source PowerPoint file, replace $curr_path with $_.DirectoryName
    $mp4_filename = "$($curr_path)\$($_.BaseName).mp4"
    # Before starting the conversion check nothing else is running
    # 0 None (nothing being processed, I'd assume)
    # 1 In Progress
    # 2 Queued
    # 3 Done
    # 4 Failed
    while ($document.CreateVideoStatus -eq "2") { # doesn't matter if the last one failed, as long as nothing is queued..
    # Wait a specific interval
    "Waiting on another video to process, before starting this conversion"
    $document.CreateVideoStatus
    Start-Sleep -Seconds 2
    }
    # Create the video Reference is https://docs.microsoft.com/en-us/office/vba/api/powerpoint.presentation.createvideo
    # Use the following hard-coded parameters UseTimingsAndNarration, DefaultSlideDuration,VertResolution,FramesPerSecond,Quality   $true,5,640,30,100
    $document.CreateVideo($mp4_filename,$true,5,640,30,100) # low res set for sharing during remote teaching.
    Start-Sleep -Seconds 2 # just wait for it start (using control values doesn't work well here as sometimes it is not in progress yet, so just wait
    # Wait until processing is complete (this takes quite some time)
    while ($document.CreateVideoStatus -eq "1") { # while in progress, wait
    ## Wait a specific interval
    "should be processing"
    $document.CreateVideoStatus
    Start-Sleep -Seconds 2
    }

    # Close PowerPoint file
    $document.Close()
}
# Exit and release the PowerPoint object
$ppt_app.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt_app)
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()