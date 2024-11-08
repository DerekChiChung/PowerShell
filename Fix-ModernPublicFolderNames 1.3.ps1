<#
    .SYNOPSIS
    Script to rename public folder names based on results in a CSV file.
    Version 1.3, 2024-11-07

    .DESCRIPTION	
    This script renames public folders based on validation results in a CSV file, specifically targeting folders 
    with `ResultType` equal to `SpecialCharacters`. The script replaces unsupported characters as follows:
    - Replaces "/" with "_"
    - Replaces "\" with "-"

    .PARAMETER ExportFolderNames
    Switch to export renamed folders to text files.

    .EXAMPLE
    Rename public folders based on validation results in a CSV file.

    .\Fix-ModernPublicFolderNames.ps1 

    .EXAMPLE
    Rename public folders and export list of renamed folders and folders with renaming errors.

    .\Fix-ModernPublicFolderNames.ps1 -ExportFolderNames

#>

[CmdletBinding()]
Param(
  [switch] $ExportFolderNames
)

Write-Host 'Reading from CSV file'

# Variables
$CSVPath = "ValidationResults.csv"
$ScriptDir = Split-Path -Path $script:MyInvocation.MyCommand.Path
$TimeStamp = $(Get-Date -Format 'yyyy-MM-dd HHmm')
$FileNameSuccess = 'RenamedFoldersSuccess'
$FileNameError = 'RenamedFoldersError'

# Read and filter CSV data
$CSVData = Import-Csv -Path $CSVPath | Where-Object { $_.ResultType -eq "SpecialCharacters" }

$UpdatedFolders = @()
$FoldersWithError = @()
$PublicFolderCount = ($CSVData | Measure-Object).Count

Write-Host ('Found {0} public folders with names containing unsupported characters' -f $PublicFolderCount)

$Count = 0

foreach ($Folder in $CSVData) { 
  # Display progress
  Write-Progress -Activity "Replace characters" -Status "Replace characters: $([math]::Round($(($Count/$PublicFolderCount)*100))) %" -PercentComplete (($Count/$PublicFolderCount)*100) -SecondsRemaining $($PublicFolderCount - $Count)
  
  # Original and new folder names
  $OriginalName = $Folder.FolderIdentity
  $NewPublicFolder = $OriginalName -replace "\\", "-" -replace "/", "_"
  $NewPublicFolder = $NewPublicFolder.Trim()
  
  try { 
    # Rename folder
    Set-PublicFolder -Identity $Folder.FolderEntryId -Name $NewPublicFolder -Confirm:$false -EA 'stop' 
    $UpdatedFolders += [PSCustomObject]@{OriginalName = $OriginalName; NewName = $NewPublicFolder}
  } 
  catch {
    # Log error
    Write-Host $Error[0].Exception.Message -ForegroundColor Yellow
    $FoldersWithError += $Folder
  }
  
  $Count++
}

if ($ExportFolderNames) {
  # Export renamed folders to file
  $OutputFile = Join-Path -Path $ScriptDir -ChildPath ('{0}-{1}.txt' -f $FileNameSuccess, $TimeStamp)
  $UpdatedFolders | ForEach-Object { "$($_.OriginalName) -> $($_.NewName)" } | Out-File -FilePath $OutputFile -Force -Confirm:$false
  
  # Export folders with errors to file
  $OutputFile = Join-Path -Path $ScriptDir -ChildPath ('{0}-{1}.txt' -f $FileNameError, $TimeStamp)
  $FoldersWithError | ForEach-Object { $_.FolderIdentity } | Out-File -FilePath $OutputFile -Force -Confirm:$false 
}

Write-Host ('Folders updated successfully: {0}' -f ($UpdatedFolders | Measure-Object).Count)
Write-Host ('Folders with update errors  : {0}' -f ($FoldersWithError | Measure-Object).Count)
