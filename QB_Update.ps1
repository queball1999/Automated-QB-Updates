#
# Program Name: QB_Update.ps1
#
# Description: Powershell script to automate Quickbooks updates
#
# Author: Quynn Bell
#
# Original Author: adamef93  - https://github.com/adamef93/Automated-QB-Updates
#
# Date Modified: 16th of July 2024
#

# FIXME: - When we run quickbooks after the install, sometimes quickbooks can trigger the window that states we neeed to update and updates quickbooks. we need a way to detect this and "pause".
#        - Add UI animation methods for clicking through QB patch installer windows.


function Check-Elevation {
    Write-Host "Checking for elevation..."
    $CurrentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    if (($CurrentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) -eq $false) {
        $ArgumentList = "-noprofile -noexit -file `"{0}`" -Path `"$Path`" -MaxStage $MaxStage"
        If ($ValidateOnly) { $ArgumentList = $ArgumentList + " -ValidateOnly" }
        If ($SkipValidation) { $ArgumentList = $ArgumentList + " -SkipValidation $SkipValidation" }
        If ($Mode) { $ArgumentList = $ArgumentList + " -Mode $Mode" }
        Write-Host "elevating"
        Start-Process powershell.exe -Verb RunAs -ArgumentList ($ArgumentList -f ($myinvocation.MyCommand.Definition)) -Wait
        Exit
    }
    Write-Host "In admin mode..."
}

function Ensure-FolderExists {
    param (
        [string]$folderPath
    )
    if (-not (Test-Path -Path $folderPath)) {
        Write-Host "Creating folder: $folderPath"
        New-Item -ItemType Directory -Path $folderPath
    }
}

function Download-And-Install-Patch {
    param (
        [string[]]$downloads,
        [string]$destination
    )
    foreach ($url in $downloads) {
        $args = "/silent", "/a"
        Start-BitsTransfer -Source $url -Destination $destination
        Unblock-File $destination
        Start-Process $destination -Wait -ArgumentList $args
        Remove-Item -Path $destination -Force
    }
}

function Apply-QuickBooksPatches {
    param (
        [string[]]$quickbooks
    )
    foreach ($quickbooksPath in $quickbooks) {
        Start-Process $quickbooksPath -Verb runas
        Start-Sleep -Seconds 20
        Get-Process qbw | ForEach-Object {
            $_.CloseMainWindow() | Out-Null
            Stop-Process -Id $_.Id -Force
        }
    }
}

function Is-QuickBooksInstalled {
    param (
        [string]$quickbooksPath
    )
    return Test-Path -Path $quickbooksPath
}

# Main script
Check-Elevation

$folderPath = "C:\IT\QB updates"
Ensure-FolderExists -folderPath $folderPath

# Disables UAC. The patches require UAC to run and this is included in preparation for automating the install with AutoIT
# Currently commenting out as it causes conflict with SIEM - Que 
#Set-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name "ConsentPromptBehaviorAdmin" -Value "0"

$destination = "C:\IT\QB Updates\qbwebpatch.exe"
$downloads = @(
    # QuickBooks Desktop Enterprise Solutions 24.0
    "https://http-download.intuit.com/http.intuit/Downloads/2024/rnkpzeq9nUS_R6/Webpatch/en_qbwebpatch.exe",
    # QuickBooks Desktop Enterprise Solutions 23.0
    "https://http-download.intuit.com/http.intuit/Downloads/2023/nctqf0a84US_R12/Webpatch/en_qbwebpatch.exe",
    # QuickBooks Desktop Enterprise Solutions 22.0
    "https://http-download.intuit.com/http.intuit/Downloads/2022/dmknzyq5nUS_R15/Webpatch/en_qbwebpatch.exe",
    # QuickBooks Desktop Enterprise Solutions 21.0
    "https://http-download.intuit.com/http.intuit/Downloads/2021/eo2bf393iUS_R17/Webpatch/en_qbwebpatch.exe",
    # QuickBooks Desktop Enterprise Solutions 20.0
    "https://http-download.intuit.com/http.intuit/Downloads/2020/cveofqqkrsUS_R17/Webpatch/en_qbwebpatch.exe",
    # QuickBooks Desktop Enterprise Solutions 19.0
    "https://http-download.intuit.com/http.intuit/Downloads/2019/szxlidxcipUS_R17/Webpatch/en_qbwebpatch.exe",
    # QuickBooks Desktop Enterprise Solutions 18.0
    "https://http-download.intuit.com/http.intuit/Downloads/2018/qaammbxfvrUS_R17/Webpatch/en_qbwebpatch.exe"
    # QuickBooks Desktop Accountant 2024
    #"https://http-download.intuit.com/http.intuit/Downloads/2024/rnkpzeq9nUS_R6/Webpatch/qbwebpatch.exe",
    # QuickBooks Desktop Accountant 2023
    #"https://http-download.intuit.com/http.intuit/Downloads/2023/nctqf0a84US_R12/Webpatch/qbwebpatch.exe",
    # QuickBooks Desktop Accountant 2022
    #"https://http-download.intuit.com/http.intuit/Downloads/2022/dmknzyq5nUS_R15/Webpatch/qbwebpatch.exe",
    # QuickBooks Desktop Accountant 2021
    #"https://http-download.intuit.com/http.intuit/Downloads/2021/eo2bf393iUS_R17/Webpatch/qbwebpatch.exe",
    # QuickBooks Desktop Accountant 2020
    #"https://http-download.intuit.com/http.intuit/Downloads/2020/cveofqqkrsUS_R17/Webpatch/qbwebpatch.exe",
    # QuickBooks Desktop Accountant 2019
    #"https://http-download.intuit.com/http.intuit/Downloads/2019/szxlidxcipUS_R17/Webpatch/qbwebpatch.exe",
    # QuickBooks Desktop Accountant 2018
    #"https://http-download.intuit.com/http.intuit/Downloads/2018/qaammbxfvrUS_R17/Webpatch/qbwebpatch.exe",
    # QuickBooks Desktop Premier 2024
    #"https://http-download.intuit.com/http.intuit/Downloads/2024/rnkpzeq9nUS_R6/Webpatch/qbwebpatch.exe",
    # QuickBooks Desktop Premier 2023
    #"https://http-download.intuit.com/http.intuit/Downloads/2023/nctqf0a84US_R12/Webpatch/qbwebpatch.exe",
    # QuickBooks Desktop Premier 2022
    #"https://http-download.intuit.com/http.intuit/Downloads/2022/dmknzyq5nUS_R15/Webpatch/qbwebpatch.exe",
    # QuickBooks Desktop Premier 2021
    #"https://http-download.intuit.com/http.intuit/Downloads/2021/eo2bf393iUS_R17/Webpatch/qbwebpatch.exe",
    # QuickBooks Desktop Premier 2020
    #"https://http-download.intuit.com/http.intuit/Downloads/2020/cveofqqkrsUS_R17/Webpatch/qbwebpatch.exe",
    # QuickBooks Desktop Premier 2019
    #"https://http-download.intuit.com/http.intuit/Downloads/2019/szxlidxcipUS_R17/Webpatch/qbwebpatch.exe",
    # QuickBooks Desktop Premier 2018
    #"https://http-download.intuit.com/http.intuit/Downloads/2018/qaammbxfvrUS_R17/Webpatch/qbwebpatch.exe",
    # QuickBooks Desktop Pro 2024
    #"https://http-download.intuit.com/http.intuit/Downloads/2024/rnkpzeq9nUS_R6/Webpatch/qbwebpatch.exe",
    # QuickBooks Desktop Pro 2023
    #"https://http-download.intuit.com/http.intuit/Downloads/2023/nctqf0a84US_R12/Webpatch/qbwebpatch.exe",
    # QuickBooks Desktop Pro 2022
    #"https://http-download.intuit.com/http.intuit/Downloads/2022/dmknzyq5nUS_R15/Webpatch/qbwebpatch.exe",
    # QuickBooks Desktop Pro 2021
    #"https://http-download.intuit.com/http.intuit/Downloads/2021/eo2bf393iUS_R17/Webpatch/qbwebpatch.exe",
    # QuickBooks Desktop Pro 2020
    #"https://http-download.intuit.com/http.intuit/Downloads/2020/cveofqqkrsUS_R17/Webpatch/qbwebpatch.exe",
    # QuickBooks Desktop Pro 2019
    #"https://http-download.intuit.com/http.intuit/Downloads/2019/szxlidxcipUS_R17/Webpatch/qbwebpatch.exe",
    # QuickBooks Desktop Pro 2018
    #"https://http-download.intuit.com/http.intuit/Downloads/2018/qaammbxfvrUS_R17/Webpatch/qbwebpatch.exe"
)

$quickbooks = @(
    # QuickBooks Desktop Enterprise Solutions 24.0
    "C:\Program Files\Intuit\QuickBooks Enterprise Solutions 24.0\QBWEnterprise*.exe",
    # QuickBooks Desktop Enterprise Solutions 23.0
    "C:\Program Files\Intuit\QuickBooks Enterprise Solutions 23.0\QBWEnterprise*.exe",
    # QuickBooks Desktop Enterprise Solutions 22.0
    "C:\Program Files\Intuit\QuickBooks Enterprise Solutions 22.0\QBWEnterprise*.exe",
    # QuickBooks Desktop Enterprise Solutions 21.0
    "C:\Program Files (x86)\Intuit\QuickBooks Enterprise Solutions 21.0\QBWEnterprise*.exe",
    # QuickBooks Desktop Enterprise Solutions 20.0
    "C:\Program Files (x86)\Intuit\QuickBooks Enterprise Solutions 20.0\QBWEnterprise*.exe",
    # QuickBooks Desktop Enterprise Solutions 19.0
    "C:\Program Files (x86)\Intuit\QuickBooks Enterprise Solutions 19.0\QBWEnterprise*.exe",
    # QuickBooks Desktop Enterprise Solutions 18.0
    "C:\Program Files (x86)\Intuit\QuickBooks Enterprise Solutions 18.0\QBWEnterprise*.exe"
    # QuickBooks Desktop Premier 2024
    #"C:\Program Files\Intuit\QuickBooks 2024\QBWPremier*.exe"
    # QuickBooks Desktop Premier 2023
    #"C:\Program Files\Intuit\QuickBooks 2023\QBWPremier*.exe"
    # QuickBooks Desktop Premier 2022
    #"C:\Program Files\Intuit\QuickBooks 2022\QBWPremier*.exe"
    # QuickBooks Desktop Premier 2021
    #"C:\Program Files (x86)\Intuit\QuickBooks 2021\QBW32Premier*.exe"
    # QuickBooks Desktop Premier 2020
    #"C:\Program Files (x86)\Intuit\QuickBooks 2020\QBW32Premier*.exe"
    # QuickBooks Desktop Premier 2019
    #"C:\Program Files (x86)\Intuit\QuickBooks 2019\QBW32Premier*.exe"
    # QuickBooks Desktop Premier 2018
    #"C:\Program Files (x86)\Intuit\QuickBooks 2018\QBW32Premier*.exe"
)

foreach ($quickbooksWildcard in $quickbooks) {
    $resolvedPaths = Get-ChildItem -Path $quickbooksWildcard
    Write-Host "Resolved path: $resolvedPaths"
    if ($resolvedPaths) {
        foreach ($quickbooksPath in $resolvedPaths) {
            if (Is-QuickBooksInstalled -quickbooksPath $quickbooksPath) {
                $url = $downloads[$quickbooks.IndexOf($quickbooksWildcard)]
                Write-Host "Resolved URL: $url"
                Download-And-Install-Patch -downloads $url -destination $destination
                Apply-QuickBooksPatches -quickbooks $quickbooksPath
            } else {
                Write-Host "Skipping download and patch for $quickbooksPath as it is not installed."
            }
        }
    } else {
        Write-Host "No QuickBooks installations found for wildcard path: $quickbooksWildcard"
    }
}

# Re-enables UAC
# Currently commenting out as it causes conflict with SIEM - Que 
#Set-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name "ConsentPromptBehaviorAdmin" -Value "5"
