#
# Program Name: QB_Update.ps1
#
# Description: Powershell script to automate Quickbooks updates
#
# Original Author: adamef93  - https://github.com/adamef93/Automated-QB-Updates
#
# Editors: Queball1999
#
# Date Modified: 1st of June 2024
#

# Runs as admin
Write-Host "Checking for elevation... "  
$CurrentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent()) 
if (($CurrentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) -eq $false) 
{ 
    $ArgumentList = "-noprofile -noexit -file `"{0}`" -Path `"$Path`" -MaxStage $MaxStage" 
    If ($ValidateOnly) { $ArgumentList = $ArgumentList + " -ValidateOnly" } 
    If ($SkipValidation) { $ArgumentList = $ArgumentList + " -SkipValidation $SkipValidation" } 
    If ($Mode) { $ArgumentList = $ArgumentList + " -Mode $Mode" } 
    Write-Host "elevating" 
    Start-Process powershell.exe -Verb RunAs -ArgumentList ($ArgumentList -f ($myinvocation.MyCommand.Definition)) -Wait 
    Exit 
}  
write-host "in admin mode.."

# Check if the C:\IT\QB updates folder exists. If not, create it.
$folderPath = "C:\IT\QB updates"
if (-not (Test-Path -Path $folderPath)) {
    Write-Host "Creating folder: $folderPath"
    New-Item -ItemType Directory -Path $folderPath
}

# Disables UAC. The patches require UAC to run and this is included in preparation for automating the install with AutoIT
# Currently commenting out as it causes conflict with SIEM - Que 
#Set-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name "ConsentPromptBehaviorAdmin" -Value "0"

# Change this to your preferred directory. Keep the file name + extension
$destination = "C:\IT\QB Updates\qbwebpatch.exe" 
# Edit this array to download the versions you need. See README for steps
$downloads = @(
# QuickBooks Desktop Enterprise Solutions 24.0
"https://http-download.intuit.com/http.intuit/Downloads/2024/rnkpzeq9nUS_R6/Webpatch/en_qbwebpatch.exe"
# QuickBooks Desktop Enterprise Solutions 23.0
"https://http-download.intuit.com/http.intuit/Downloads/2023/nctqf0a84US_R12/Webpatch/en_qbwebpatch.exe"
# QuickBooks Desktop Enterprise Solutions 22.0
"https://http-download.intuit.com/http.intuit/Downloads/2022/dmknzyq5nUS_R15/Webpatch/en_qbwebpatch.exe"
)

$downloads | foreach {
    $args = "/silent", "/a"
    Start-BitsTransfer -Source $_ -Destination $destination 
    Unblock-File $destination
    Start-Process $destination -Wait -ArgumentList $args
    # This deletes the downloaded file upon completion of the loop. The update packages are overwritten as new ones are downloaded, but they can be large and don't need to stay       after installation
    Remove-Item -Path $destination -Force 
}

# This deletes the temp directory created by the patch installer, again to save space
Remove-Item -Path C:\Windows\Temp\qbwebpatch -Recurse -Force

# Edit this array for the versions you have installed
$quickbooks = @( 
# QuickBooks Enterprise Solutions 24.0
"C:\Program Files\Intuit\QuickBooks Enterprise Solutions 24.0\QBWEnterpriseProfessional.exe"
# QuickBooks Enterprise Solutions 23.0
"C:\Program Files\Intuit\QuickBooks Enterprise Solutions 23.0\QBWEnterpriseProfessional.exe"
# QuickBooks Enterprise Solutions 22.0
"C:\Program Files\Intuit\QuickBooks Enterprise Solutions 22.0\QBWEnterpriseProfessional.exe"
)
# This loop launches each version in the array above as admin so the patches will apply
# THIS WILL KILL ALL OPEN VERSIONS OF QUICKBOOKS SO MAKE SURE YOU ARE THE ONLY ONE USING THE SERVER
$quickbooks | foreach {
    Start-Process $_ -Verb runas
    # Pauses the script to give QB enough time to process the update 
    Start-Sleep -Seconds 20
    get-process qbw | foreach {$_.CloseMainWindow() | Out-Null} | stop-process –force # updated to qbw as newer versions only run 64 bit.
}

# Re-enables UAC
# Currently commenting out as it causes conflict with SIEM - Que 
#Set-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name "ConsentPromptBehaviorAdmin" -Value "5"
