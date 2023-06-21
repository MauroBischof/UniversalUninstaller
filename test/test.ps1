

$global:PSSession = New-PSSession -ComputerName $targetComputer


if (!$PSSession) {
    Set-DisplayBoxText -displayBox $displaybox -text "No Connection could be established to $targetComputer with the current credentials, please provide alternative .. " -isError $true
    [PSCredential]$credential = Get-Credential
    $global:PSSession = New-PSSession -ComputerName $targetComputer -Credential $credential

}
if (!$PSSession) {
    Set-DisplayBoxText -displayBox $displaybox -text "No Connection could be established to $targetComputer with the current credentials, please provide alternative .. " -isError $true
    Enable-PsRemoting -ComputerName $targetComputer -credential $credential
    Start-Sleep -Seconds 10
    $global:PSSession = New-PSSession -ComputerName $targetComputer -Credential $credential
}
if ($PSSession) {

    Set-DisplayBoxText -displayBox $displaybox -text ("Successfully connected to $targetComputer.")
    Set-ConnectedTo -menuItem $formItems.menuStrip.menuStrip.Items[1] -text ("Connected to: $targetComputer")

    $InstalledApps = Invoke-Command -Session $PSSession -ScriptBlock ${function:Get-InstalledApp}
    Update-TableContent -tableContent $InstalledApps -table ($formItems.Table.Table)
}
else {
    Set-DisplayBoxText -displayBox $displaybox -text ($error[0].ErrorDetails) -isError $true
}