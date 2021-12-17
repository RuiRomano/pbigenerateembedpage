#Requires -Modules @{ ModuleName="MicrosoftPowerBIMgmt"; ModuleVersion="1.2.1026" }

param(
    [string]$outputPath = ".\output"
    ,
    [string]$embedHtmlFilePath = ".\embedPageTemplate.html"
    ,
    $reports = @(
        @{
            "workspaceId" = "cdee92d2-3ff9-43e2-9f71-0916e888ad27"
            ;
            "reportId" = "d471e78e-2251-4085-9a60-678c0d4b5dfa"
        }
        ,
        @{
            "workspaceId" = "9f19cb45-1182-4656-82b0-e7a049985835"
            ;
            "reportId" = "83b4e4ab-5f17-430c-a0f1-8f791cb14fad"
        }
    )
)

$ErrorActionPreference = "Stop"

$currentPath = (Split-Path $MyInvocation.MyCommand.Definition -Parent)

Set-Location $currentPath

# Ensure output folders

New-Item -ItemType Directory -Path $outputPath -ErrorAction SilentlyContinue | Out-Null

$htmlTemplate = Get-Content $embedHtmlFilePath

try {
    $accessToken = Get-PowerBIAccessToken -AsString
}
catch {
    Connect-PowerBIServiceAccount
    $accessToken = Get-PowerBIAccessToken -AsString 
}

$accessToken = $accessToken.Replace("Bearer ","").Trim()

Write-Host "Preparing HTML files on '$outputPath'"

foreach($report in $reports)
{
    $workspaceId = $report.workspaceId
    $reportId = $report.reportId

    $reportHtml = $htmlTemplate.Replace("[ACCESSTOKEN]","$accessToken").Replace("[REPORTID]",$reportId).Replace("[WORKSPACEID]",$workspaceId)

    $outputFilePath = "$outputPath\$reportId.html"

    $reportHtml | Out-File $outputFilePath

    Write-Host "HTML file ready: '$outputFilePath'"
}
