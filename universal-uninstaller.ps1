<#
.SYNOPSIS
Beschreibung des Skripts.

.DESCRIPTION
Ausführliche Beschreibung des Skripts.

.PARAMETER Parameter1
Beschreibung des Parameter1.

.PARAMETER Parameter2
Beschreibung des Parameter2.

.EXAMPLE
Beispiel für die Verwendung des Skripts.

.EXAMPLE
Ein weiteres Beispiel für die Verwendung des Skripts.

.NOTES
Zusätzliche Informationen zum Skript.

.LINK
Verweis auf weitere Informationen zum Skript.

# how to run
# .\universalSilentUninstaller.ps1 -AppName '' -AppVersion ''

#>

#region ******************** TOUCH ********************
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
Set-StrictMode -Version Latest
$ScriptDir = $PSScriptRoot
$ModuleDir = $ScriptDir + "\modules\"
#region ******************** TESTING ********************
$testing = $true
if ($testing) {
    $targetComputer = "HAMV21738"  # localhost HAMV21738
    $credential = "" #emea\rootmb  vboxuser
    #for testing

}
#endregion TESTING
#region ******************** MODULES ********************
$modules = @(
( $ModuleDir + "gui\dyngui.psm1" ),
( $ModuleDir + "functions\functions.psm1" ),
( $ModuleDir + "guiDetailView\guiDetailView.psm1" ),
( $ModuleDir + "validation\validation.psm1" )
)
#endregion MODULES
#endregion TOUCH

#region ******************** ADMIN ********************
# only change this region if necessary. e.g. add more modules or init-folders
function Main {
    Requirements -LogText "Requirements"
    Work -LogText "Work"
    Cleanup
}

Function Requirements {
    param([parameter(Mandatory = $true)][string]$LogText)
    try {
        foreach ($module in $modules) {
            Import-Module $module
        }
        #check if the init-folders exist - if not create them
        $InitFolders = @($ScriptDir)
        Foreach ($Folder in $InitFolders) { If (!(test-path $Folder)) { New-Item -ItemType Directory -Force -Path $Folder } }
    }
    catch {
        #Cleanup -exitCode 1
        "FATAL - initialisation not passed - Error at line " + $_.InvocationInfo.ScriptLineNumber + ": " + $_.Exception.Message | Out-File -FilePath ($MyInvocation.PSCommandPath + "_error.log") -Force
        exit
    }
}

# clear variables and console output, remove modules
Function Cleanup {
    try {
        Get-Module | Remove-Module -Force -ErrorAction SilentlyContinue
        Get-Job | Stop-Job | Remove-Job -Force -ErrorAction SilentlyContinue
        Get-PSSession | Remove-PSSession  -ErrorAction SilentlyContinue
        Get-Variable -Exclude exitCode, PWD, *Preference | Remove-Variable -Force -ErrorAction SilentlyContinue
        exit

    }
    catch {
        exit 1
    }
}
#endregion ADMIN

#region ******************** WORK ********************
# put your code into the try block, clone the function if needed.
Function Work {
    param(
        $LogText
    )
    try {
        #Funktioniert nicht
        $form, $formItems = Add-Form
        Add-FormAction -formItems $formItems -form $form
        $form.ShowDialog()

    }
    catch {
        $logMessage = "FAIL - $LogText - Error at line " + $_.InvocationInfo.ScriptLineNumber + ": " + $_.Exception.Message
        $logMessage | Out-File -FilePath ($MyInvocation.PSCommandPath + "_error.log") -Force
        Write-Host $logMessage
        #Cleanup -exitCode 1
    }
    finally {
        Cleanup
    }
}

function Add-Form {
    # Erstellen eines neuen Windows-Formulars
    $form = New-Object System.Windows.Forms.Form
    $form.Size = New-Object System.Drawing.Size(850, 600)
    $form.MinimumSize = $form.Size

    $form.Text = "Universal Uninstaller"
    $form.KeyPreview = $true
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    # Create the table layout panel
    $mainPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $mainPanel.Dock = 'Fill'
    #würd ich gern auf bottom und top beschränken
    $mainPanel.Padding = New-Object System.Windows.Forms.Padding(25)  # Set the padding here


    $formItems = @{
        ConnectArea  = New-ConnectArea -mainPanel $mainPanel
        OutputArea   = New-OutputArea -mainPanel $mainPanel
        CommandArea  = New-CommandArea -mainPanel $mainPanel
        TableActions = New-TableAction -mainPanel $mainPanel
        Table        = New-Table -mainPanel $mainPanel
        ExitButton   = New-ExitButton -mainPanel $mainPanel
        ProgressBar  = New-ProgressBar
        MenuStrip    = (New-MenuStrip).GetEnumerator() | Select-Object -Last 1

    }
    # Hinzufügen von ProgressBar und MenuStrip zum Formular
    foreach ($item in $formItems.GetEnumerator() | Where-Object { $_.Name -in "ProgressBar", "MenuStrip" }) {
        foreach ($formItem in $item.Value) {
            foreach ($item in $formItem.GetEnumerator()) {
                $form.Controls.Add($item.Value)
            }
        }
    }


    $form.Controls.Add($mainPanel)
    return $form, $formItems
}

function Add-FormAction {
    param (
        $formItems,
        $form
    )

    New-ExitButtonAction -ExitButton ($formItems.ExitButton.ExitButton) -form $form
    New-MenuStripAction -menuStrip ($formItems.menuStrip.menuStrip) -form $form

    if ((Test-PreRequirement -ouputTextBox ($formItems.OutputArea.ouputTextBox) -requiredVersion "5.1" -requireAdminRights $true -requiredPolicy "bypass")) {

        New-SearchBoxAction -filterTextBox ($formItems.TableActions.filterTextBox)
        New-TableCellDoubleClick  -table ($formItems.Table.Table)
        New-ConectButtonAction -ConectButton ($formItems.ConnectArea.ConnectButton)
        New-UninstallButtonAction -UninstallButton ($formItems.TableActions.UninstallButton)
        New-F5ButtonAction #-table ($formItems.Table.Table)
        New-PWSHCheckBoxChange -pwshCommands (Get-Command | Where-Object { $_.Module -like "Microsoft.Powershell*" } | Select-Object Name)
        New-InvokeButtonAction -InvokeButton ($formItems.CommandArea.invokeButton)
        New-TextBoxPopOutFormAction -popOutLabel ($formItems.OutputArea.popOutLabel)
    }
}


function New-ConectButtonAction {
    param(
        [parameter(Mandatory = $true)]$ConectButton
    )
    $global:ConnectTextlastEntries = New-Object System.Collections.Generic.List[string]

    $ConectButton.Add_Click({
            try {
                if (!$formItems.ConnectArea.HostnameBox.Text) {

                    #$targetComputer = ($formItems.ConnectArea.HostnameBox).Text
                    $displaybox = $formItems.OutputArea.ouputTextBox


                    Set-DisplayBoxText -displayBox $displaybox -text "Please wait while a connection to $targetComputer is established ..."
                    $global:PSSession = New-PSSession -ComputerName $targetComputer

                    if (!$PSSession) {
                        Set-DisplayBoxText -displayBox $displaybox -text "Please wait while a connection to $targetComputer is established ..."
                        $PSSession = New-PSSession -ComputerName $targetComputer -Credential $credential
                    }
                    if ($PSSession) {

                        Set-DisplayBoxText -displayBox $displaybox -text ("Successfully connected to $targetComputer.")
                        $InstalledApps = Invoke-Command -Session $PSSession -ScriptBlock ${function:Get-InstalledApp}
                        Update-TableContent -tableContent $InstalledApps -table ($formItems.Table.Table)
                    }
                    else {
                        Set-DisplayBoxText -displayBox $displaybox -text ($error[0].ErrorDetails)


                    }

                    $formItems.CommandArea.commandBox.Clear()
                    $formItems.TableActions.filterTextBox.Clear()
                    Set-LastEntrie -text $displaybox.Text -lastEntries $ConnectTextlastEntries -autoCompleteSource ($formItems.ConnectArea.HostnameBox.AutoCompleteCustomSource)
                }
            }
            catch {
                $logMessage = "FAIL - $LogText - Error at line " + $_.InvocationInfo.ScriptLineNumber + ": " + $_.Exception.Message
                Set-DisplayBoxText -displayBox $displaybox -text $logMessage
            }
        })
}


function New-SearchBoxAction {
    param (
        $filterTextBox
    )
    $filterTextBox.Add_TextChanged({
            $keyword = ($formItems.TableActions.filterTextBox).Text
            foreach ($row in ($formItems.Table.Table).Rows) {
                if ($row.IsNewRow) {
                    continue
                }
                $match = $false
                foreach ($cell in $row.Cells[0]) {
                    if ($cell.Value.ToString() -like "*$keyword*") {
                        $match = $true
                        break
                    }
                }
                $row.Visible = $match
            }
        })
}

function New-InvokeButtonAction {
    param (
        $InvokeButton
    )

    $global:CommandlastEntries = New-Object System.Collections.Generic.List[string]

    $InvokeButton.Add_Click({
            try {
                #funktioniert noch nicht 100%
                if ($PSSession -and $PSSession.Availability -eq "Available" -and ($formItems.CommandArea.commandBox.Text)) {

                    Set-LastEntrie -text $formItems.CommandArea.commandBox.Text -lastEntries $CommandlastEntries -autoCompleteSource $formItems.CommandArea.commandBox.AutoCompleteCustomSource

                    $exitcode = Invoke-CustomCommand `
                        -command (($formItems.CommandArea.commandBox).Text) `
                        -progressBar ($formItems.ProgressBar.ProgressBar) `
                        -type "c_command" `
                        -radioButtonPWSH (($formItems.CommandArea.radioButtonPWSH).Checked)

                    Set-DisplayBoxText -displayBox ($formItems.OutputArea.ouputTextBox) -text $exitcode


                }

            }
            catch {
                $logMessage = "FAIL - $LogText - Error at line " + $_.InvocationInfo.ScriptLineNumber + ": " + $_.Exception.Message
                Set-DisplayBoxText -displayBox ($formItems.OutputArea.ouputTextBox) -text $logMessage
            }


        })
}

function New-TableCellDoubleClick {
    param (
        $table
    )
    $table.Add_CellDoubleClick({
            Show-Detail -table ($formItems.Table.Table)
        })
}

function New-UninstallButtonAction {
    param(
        $UninstallButton
    )
    $UninstallButton.Add_Click({
            try {

                #funktioniert noch nicht 100%
                if ($PSSession -and $PSSession.Availability -eq "Available") {
                    $selectedRow = ($formItems.Table.Table).selectedRows[0]
                    $selectedItemName = $selectedRow.Cells["Name"].Value

                    if ($selectedItemName) {
                        $result = [System.Windows.Forms.MessageBox]::Show("Do you want to uninstall '$selectedItemName' ?", "Confirm", "YesNo", "Question")
                        if ($result -eq "Yes") {
                            #Invoke-KillProcess -displayName $selectedItemName -ErrorAction SilentlyContinue
                            $exitcode = Invoke-SilentUninstallString `
                                -command ($selectedRow.Cells["Uninstallstring"].Value) `
                                -progressBar ($formItems.ProgressBar.ProgressBar) `
                                -type ($selectedRow.Cells["Type"].Value)

                            #schöner, hier nicht korrekt ?
                            switch ($exitcode) {
                                124 { $text = ("The uninstaller ran into an timout. Please verify that " + ($selectedRow.Cells["Uninstallstring"].Value) + " is the correct unintallstring for the product $selectedItemName.") }
                                0 { $text = ("The product $selectedItemName was uninstalled successful.") }
                                1 { $text = ("Error while uninstalling $selectedItemName. Please verify that " + ($selectedRow.Cells["Uninstallstring"].Value) + " is the correct unintallstring for the product ") }
                                Default { $text = ("Unkown Status") }
                            }

                            $testApp = (Invoke-Command -Session $PSSession -ScriptBlock ${function:Get-InstalledApp} -ArgumentList $selectedItemName)

                            if ($exitcode -eq 0 -and !$testApp) {
                                [System.Windows.Forms.MessageBox]::Show("Uninstallation successful.", "Ok", "OK", "Information")
                                Set-DisplayBoxText -displayBox ($formItems.OutputArea.ouputTextBox) -text $text
                                Get-Table -table ($formItems.Table.Table)

                            }
                            else {
                                [System.Windows.Forms.MessageBox]::Show("Uninstallation not successful.", "Error", "OK", "Hand")
                                Set-DisplayBoxText -displayBox ($formItems.OutputArea.ouputTextBox) -text $text
                                Get-Table -table ($formItems.Table.Table)
                            }
                        }
                    }
                    else {
                        [System.Windows.Forms.MessageBox]::Show("No element selected.", "Warning", "OK", "Warning")
                    }
                }

            }
            catch {
                $logMessage = "FAIL - $LogText - Error at line " + $_.InvocationInfo.ScriptLineNumber + ": " + $_.Exception.Message
                Set-DisplayBoxText -displayBox ($formItems.OutputArea.ouputTextBox) -text $logMessage
            }
        })
}

function New-ExitButtonAction {
    param (
        $form,
        $ExitButton
    )
    $ExitButton.Add_Click({
            $form.Close()
            #Cleanup
        })
}

Function New-MenuStripAction {
    param (
        $menuStrip,
        $form
    )
    #exit
    $menuStrip.Items[0].DropDownItems[2].Add_Click({
            $form.Close()
        })

    <#   $menuStrip.Items[0].DropDownItems[2].Add_Click({
        $form.Close()
    })#>

    #donate
    $menuStrip.Items[0].DropDownItems[1].Add_Click({
            Start-Process "https://www.paypal.com/donate/?hosted_button_id=PP27RZLCAAKX2"
        })

    #about
    $menuStrip.Items[0].DropDownItems[0].Add_Click({
            Start-Process "https://google.com"
        })

}

function New-PWSHCheckBoxChange {
    param(
        $pwshCommands
    )
    $global:pwshCommands = $pwshCommands
    Set-PWSHCommandAutoComplete -autoCompleteSource ($formItems.CommandArea.commandBox.AutoCompleteCustomSource) -pwshCommands $pwshCommands
    ($formItems.CommandArea.radioButtonPWSH).add_CheckedChanged({

            if (($formItems.CommandArea.radioButtonPWSH).Checked) {
                Set-PWSHCommandAutoComplete -autoCompleteSource ($formItems.CommandArea.commandBox.AutoCompleteCustomSource) -pwshCommands $pwshCommands
            }
            else {
                ($formItems.CommandArea.commandBox.AutoCompleteCustomSource).Clear()
                ($formItems.CommandArea.commandBox.AutoCompleteCustomSource).AddRange($CommandlastEntries)
            }

        })


}

function New-TextBoxPopOutFormAction {
    param (
        $popOutLabel
    )
    # KeyDown-Ereignis der Form hinzufügen

    $popOutLabel.Add_Click({
            if ((($formItems.OutputArea.ouputTextBox).Text)) {
                New-TextBoxPopOutForm -text (($formItems.OutputArea.ouputTextBox).Text)
            }
        })
}

function New-F5ButtonAction {
    param (
        $table
    )
    # KeyDown-Ereignis der Form hinzufügen

    $form.Add_KeyDown({
            if ($_.KeyCode -eq "F5") {
                Get-Table -table ($formItems.Table.Table)
            }
        })
}

#endregion WORK
Main
