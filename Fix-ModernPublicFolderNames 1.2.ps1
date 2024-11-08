<#
    .SYNOPSIS
    Script to prepare on-premises modern public folder names for migration to modern public folders in Exchange Online 

    Originally from Thomas Stensitzki http://scripts.Granikos.eu
	
    .DESCRIPTION	
    This script renames modern public folder names by replacing unsupported characters. Specifically:
    - Replaces "/" with a underscore "_".
    - Replaces "\" with a hyphen "-".

    .NOTES 
    Requirements 
    - Windows Server Windows Server 2012R2 or newer
    - Exchange 2013 Management Shell or newer
    - Organization Management RBAC Management Role

    Revision History 
    -------------------------------------------------------------------------------- 
    1.0     Initial community release 
    1.1     Added logging for original and new folder names 
    1.2     Updated character replacements

    .PARAMETER ExportFolderNames
    Switch to export renamed folders to text files
    
    .EXAMPLE
    Rename and trim public folders

    .\Fix-ModernPublicFolderNames.ps1 

    .EXAMPLE
    Rename and trim public folders, export list of renamed folders and folders with renaming errors as text files

    .\Fix-ModernPublicFolderNames.ps1 -ExportFolderNames

#>

[CmdletBinding()]
Param(
  [switch] $ExportFolderNames
)

Write-Host 'Fetching Public Folders'

# Variables
$FolderScope = '\'
$ScriptDir = Split-Path -Path $script:MyInvocation.MyCommand.Path
$TimeStamp = $(Get-Date -Format 'yyyy-MM-dd HHmm')
$FileNameSuccess = 'RenamedFoldersSuccess'
$FileNameError = 'RenamedFoldersError'

# Fetch all public folders
$PublicFolders = Get-PublicFolder $FolderScope -Recurse -ResultSize Unlimited

# Filter Public Folders with unsupported characters
$FilteredFolders = $PublicFolders | Where-Object { ($_.Name -like "*\*") -or ($_.Name -like "*/ *") } 

$UpdatedFolders = @()
$FoldersWithError = @()
$PublicFolderCount = ($FilteredFolders | Measure-Object).Count

Write-Host ('Found {0} public folders with names containing unsupported characters' -f $PublicFolderCount)

$Count = 0

foreach ($PublicFolder in $FilteredFolders) { 
  # Display progress
  Write-Progress -Activity "Replace characters" -Status "Replace characters: $([math]::Round($(($Count/$PublicFolderCount)*100))) %" -PercentComplete (($Count/$PublicFolderCount)*100) -SecondsRemaining $($PublicFolderCount - $Count)
  
  # Generate new folder name with specific replacements
  $NewPublicFolder = $PublicFolder.Name -replace "\\", "-" -replace "/", "_"
  $NewPublicFolder = $NewPublicFolder.Trim()
  
  try { 
    # Rename folder
    Set-PublicFolder -Identity $PublicFolder.EntryId -Name $NewPublicFolder -Confirm:$false -EA 'stop' 
    $UpdatedFolders += [PSCustomObject]@{OriginalName = $PublicFolder.Name; NewName = $NewPublicFolder}
  } 
  catch {
    # Log error
    Write-Host $Error[0].Exception.Message -ForegroundColor Yellow
    $FoldersWithError += $PublicFolder
  }
  
  $Count++
}

if ($ExportFolderNames) {
  # Export renamed folders to file
  $OutputFile = Join-Path -Path $ScriptDir -ChildPath ('{0}-{1}.txt' -f $FileNameSuccess, $TimeStamp)
  $UpdatedFolders | ForEach-Object { "$($_.OriginalName) -> $($_.NewName)" } | Out-File -FilePath $OutputFile -Force -Confirm:$false
  
  # Export folders with errors to file
  $OutputFile = Join-Path -Path $ScriptDir -ChildPath ('{0}-{1}.txt' -f $FileNameError, $TimeStamp)
  $FoldersWithError | ForEach-Object { $_.Name } | Out-File -FilePath $OutputFile -Force -Confirm:$false 
}

Write-Host ('Folders updated successfully: {0}' -f ($UpdatedFolders | Measure-Object).Count)
Write-Host ('Folders with update errors  : {0}' -f ($FoldersWithError | Measure-Object).Count)
