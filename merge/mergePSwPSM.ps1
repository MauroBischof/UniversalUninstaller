$scriptPath = (Split-Path $PSScriptRoot -Parent) + "\universal-uninstaller.ps1"
$ModuleDir = (Split-Path $PSScriptRoot -Parent) + "\modules\"

$modulePaths = @(
  ( $ModuleDir + "gui\dyngui.psm1"),
  ( $ModuleDir + "functions\functions.psm1"),
  ( $ModuleDir + "guiDetailView\guiDetailView.psm1"),
  ( $ModuleDir + "validation\validation.psm1")
)

$outfile = $PSScriptRoot+'\merged.ps1'
Merge-ScriptWithModules -ScriptPath $scriptPath -ModulePaths $modulePaths | Out-File $outfile
& "$outfile"


function Merge-ScriptWithModules {
  param (
    [Parameter(Mandatory = $true)]
    [string]$ScriptPath,

    [Parameter(Mandatory = $true)]
    [string[]]$ModulePaths
  )

  # Skriptinhalt lesen
  $scriptContent = Get-Content -Path $ScriptPath | Where-Object {$_ -notlike "*Import-Module*" -and $_ -ne "Main" } | Out-String

  # Alle Modulinhalte lesen und zusammenführen
  $moduleContents = foreach ($modulePath in $ModulePaths) {
   Get-Content -Path $modulePath | Where-Object {$_ -notlike "*Export-ModuleMember*"} | Out-String
  }

  # Skriptinhalt und Modulinhalt zusammenführen
  $mergedContent = $scriptContent + "`n" + $moduleContents + "`nMain"

  return $mergedContent
}

