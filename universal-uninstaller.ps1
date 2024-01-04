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
$global:PSSession = $null

# Read the JSON content from the file
$configFile = "config.json"
#region ******************** TESTING ********************

#endregion TESTING
#region ******************** MODULES ********************
$modules = @(
( $ModuleDir + "gui\dyngui.psm1" ),
( $ModuleDir + "functions\functions.psm1" ),
( $ModuleDir + "guiDetailView\guiDetailView.psm1" ),
( $ModuleDir + "guiComputerView\guiComputerView.psm1" ),
( $ModuleDir + "validation\validation.psm1" )
)
#endregion MODULES
#endregion TOUCH

#region ******************** ADMIN ********************
# only change this region if necessary. e.g. add more modules or init-folders
function Main {
    Requirements
    Work
    Cleanup
}

Function Requirements {
    try {
        Start-Transcript -Path ($MyInvocation.PSCommandPath + ".log")

        foreach ($module in $modules) {
            Import-Module $module
        }

    }
    catch {
        "FATAL - initialisation not passed - Error at line " + $_.InvocationInfo.ScriptLineNumber + ": " + $_.Exception.Message | Out-File -FilePath ($MyInvocation.PSCommandPath + "_error.log") -Force
        Cleanup
    }
}

# clear variables and console output, remove modules
Function Cleanup {
    try {
        Stop-Transcript
        Get-Module | Remove-Module -Force -ErrorAction SilentlyContinue
        Get-Job | Remove-Job -Force -ErrorAction SilentlyContinue
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
        Write-Error ("FAIL - $LogText - Error at line " + $_.InvocationInfo.ScriptLineNumber + ": " + $_.Exception.Message)

    }
    finally {
        Cleanup
    }
}

function Add-Form {
    # Erstellen eines neuen Windows-Formulars
    $form = New-Object System.Windows.Forms.Form
    $form.Size = New-Object System.Drawing.Size(1200, 600)
    $form.MinimumSize = $form.Size
    $form.Icon = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new((Get-Base64)).GetHIcon()))


    $form.Text = "Universal Uninstaller"
    $form.KeyPreview = $true
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    # Create the table layout panel
    $mainPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $mainPanel.Dock = 'Fill'
    $mainPanel.Name = "mainPanel"

    $mainPanel.Padding = New-Object System.Windows.Forms.Padding(10, 25, 10, 25)


    $formItems = @{
        ConnectArea  = New-ConnectArea -mainPanel $mainPanel
        OutputArea   = New-OutputArea -mainPanel $mainPanel
        CommandArea  = New-CommandArea -mainPanel $mainPanel
        TableActions = New-TableAction -mainPanel $mainPanel
        Table        = New-Table -mainPanel $mainPanel
        Footer       = New-Footer -mainPanel $mainPanel


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


    New-ExitButtonAction -ExitButton ($formItems.Footer.ExitButton) -form $form
    New-MenuStripAction -menuStrip ($formItems.menuStrip.menuStrip) -form $form

    if ((Test-PreRequirement -ouputTextBox ($formItems.OutputArea.ouputTextBox) -requiredVersion "5.1" -requireAdminRights $true )) {

        New-SearchBoxAction -filterTextBox ($formItems.TableActions.filterTextBox)
        New-TableCellClick  -table ($formItems.Table.Table)
        New-TextBoxPopOutFormAction -popOutLabel ($formItems.OutputArea.popOutLabel)
        New-ConectButtonAction -ConectButton ($formItems.ConnectArea.ConnectButton)
        New-SaveButtonAction -SaveButton ($formItems.Footer.saveButton)
        New-UninstallButtonAction -UninstallButton ($formItems.TableActions.UninstallButton)
        New-F5ButtonAction
        New-PWSHCheckBoxChange
        New-InvokeButtonAction -InvokeButton ($formItems.CommandArea.invokeButton)
        New-CommandBoxPopOutFormAction


    }
}

function New-ConectButtonAction {
    param(
        [parameter(Mandatory = $true)]$ConectButton
    )
    $global:ConnectTextlastEntries = New-Object System.Collections.Generic.List[string]

    $ConectButton.Add_Click({
            try {
                $debugMode = Get-ConfigValue -configFile $configFile -configPart "debugMode"
                if ($debugMode) { $targetComputer = Get-ConfigValue -configFile $configFile -configPart "TargetComputer" }
                if ($formItems.ConnectArea.HostnameBox.Text -or $debugMode) {
                    if (!$debugMode) {$targetComputer = $formItems.ConnectArea.HostnameBox.Text}

                    $displaybox = $formItems.OutputArea.ouputTextBox
                    Get-PSSession | Remove-PSSession  -ErrorAction SilentlyContinue
                    Set-DisplayBoxText -displayBox $displaybox -text "Please wait while a connection to $targetComputer is established ..."

                    $global:PSSession = New-PSSession -ComputerName $targetComputer | Out-Null

                    if (!$PSSession) {
                        Set-DisplayBoxText -displayBox $displaybox -text "No connection could be established to with the current credentials. Please provide administrator credentials for $targetComputer" -isError $true
                        [PSCredential]$credential = Get-Credential
                        $global:PSSession = New-PSSession -ComputerName $targetComputer -Credential $credential
                    }
                    if (!$PSSession) {
                        Set-DisplayBoxText -displayBox $displaybox -text "Please wait while a connection to $targetComputer is established ..."
                        Enable-PsRemoting -ComputerName $targetComputer -credential $credential
                        $global:PSSession = New-PSSession -ComputerName $targetComputer -Credential $credential
                    }

                    if ($PSSession) {
                        Set-DisplayBoxText -displayBox $displaybox -text ("Successfully connected to $targetComputer.")
                        Set-ConnectedTo -menuItem $formItems.menuStrip.menuStrip.Items[1] -text ("Connected to: $targetComputer")
                        $InstalledApps = Invoke-Command -Session $PSSession -ScriptBlock ${function:Get-InstalledApp}
                        Update-TableContent -tableContent $InstalledApps -table ($formItems.Table.Table)
                        Start-SystemInfo
                    }
                    else { Set-DisplayBoxText -displayBox $displaybox -text ($error[0].ErrorDetails) -isError $true }

                    $formItems.CommandArea.commandBox.Clear()
                    $formItems.TableActions.filterTextBox.Clear()
                    Set-LastEntry -text $displaybox.Text -lastEntries $ConnectTextlastEntries -autoCompleteSource ($formItems.ConnectArea.HostnameBox.AutoCompleteCustomSource)
                }
            }
            catch {
                Set-DisplayBoxText -displayBox $displaybox -text $_.Exception.Message -isError $true
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

                if ($PSSession -and $PSSession.Availability -eq "Available" -and ($formItems.CommandArea.commandBox.Text -or $formItems.CommandArea.mlcommandBox.Text)) {

                    if ($formItems.CommandArea.commandBox.Text) {
                        $command = $formItems.CommandArea.commandBox.Text
                    }
                    elseif ($formItems.CommandArea.mlcommandBox.Text) {
                        $command = $formItems.CommandArea.mlcommandBox.Text
                    }


                    Set-LastEntry -text $formItems.CommandArea.commandBox.Text -lastEntries $CommandlastEntries -autoCompleteSource $formItems.CommandArea.commandBox.AutoCompleteCustomSource

                    $exitcode = Invoke-CustomCommand `
                        -command ($command) `
                        -progressBar ($formItems.ProgressBar.ProgressBar) `
                        -type "c_command" `
                        -radioButtonPWSH (($formItems.CommandArea.radioButtonPWSH).Checked)

                    Set-DisplayBoxText -displayBox ($formItems.OutputArea.ouputTextBox) -text $exitcode


                }

            }
            catch {
                Set-DisplayBoxText -displayBox ($formItems.OutputArea.ouputTextBox) -text $_.Exception.Message -isError $true
            }


        })
}

function New-UninstallButtonAction {
    param(
        $UninstallButton
    )
    $UninstallButton.Add_Click({
            try {
                if ($PSSession -and $PSSession.Availability -eq "Available") {
                    $selectedRow = ($formItems.Table.Table).selectedRows[0]
                    $selectedItemName = $selectedRow.Cells["Name"].Value

                    if ($selectedItemName) {
                        $result = [System.Windows.Forms.MessageBox]::Show("Do you want to uninstall '$selectedItemName' ?", "Confirm", "YesNo", "Question")
                        if ($result -eq "Yes") {
                            Invoke-KillProcess -displayName $selectedItemName -ErrorAction SilentlyContinue
                            $exitcode = Invoke-SilentUninstallString `
                                -command ($selectedRow.Cells["Uninstallstring"].Value) `
                                -progressBar ($formItems.ProgressBar.ProgressBar) `
                                -type ($selectedRow.Cells["Type"].Value)

                            switch ($exitcode) {
                                124 { $text = ("The uninstaller ran into a timout of" + (Get-ConfigValue -configFile $configFile -configPart "TargetComputer") + " seconds. This value can be changed in the config.json file. Please verify that " + ($selectedRow.Cells["Uninstallstring"].Value) + " is the correct uninstallstring for the product $selectedItemName.") }
                                0 { $text = ("The product $selectedItemName was uninstalled successful.") }
                                1 { $text = ("Error while uninstalling $selectedItemName. Please verify that " + ($selectedRow.Cells["Uninstallstring"].Value) + " is the correct unintallstring for the product ") }
                                Default { $text = ("Unkown Status") }
                            }

                            $testApp = (Invoke-Command -Session $PSSession -ScriptBlock ${function:Get-InstalledApp} -ArgumentList $selectedItemName)

                            if ($exitcode -eq 0 -and !$testApp) {
                                [System.Windows.Forms.MessageBox]::Show("Uninstallation successful.", "Ok", "OK", "Information")
                                Set-DisplayBoxText -displayBox ($formItems.OutputArea.ouputTextBox) -text $text

                                Get-Table -table ($formItems.Table.Table)
                                $formItems.TableActions.filterTextBox.Clear()

                            }
                            else {
                                [System.Windows.Forms.MessageBox]::Show("Uninstallation not successful.", "Error", "OK", "Hand")
                                Set-DisplayBoxText -displayBox ($formItems.OutputArea.ouputTextBox) -text $text -isError $true
                                $formItems.TableActions.filterTextBox.Clear()

                            }
                        }
                    }
                    else {
                        [System.Windows.Forms.MessageBox]::Show("No element selected.", "Warning", "OK", "Warning")
                    }

                }

            }
            catch {
                Set-DisplayBoxText -displayBox ($formItems.OutputArea.ouputTextBox) -text $_.Exception.Message -isError $true
            }
        })
}

Function New-MenuStripAction {
    param (
        $menuStrip
    )

    $menuStrip.Items["FileMenu"].DropDownItems["exitButton"].Add_Click({
            $form.Close()
        })

    $menuStrip.Items["FileMenu"].DropDownItems["aboutButton"].Add_Click({

        })

    $menuStrip.Items["FileMenu"].DropDownItems["donateButton"].Add_Click({

        })


    $menuStrip.Items["connectedTo"].Add_Click({

            if ($PSSession) {
                $systemInfo = Get-SystemInfo
                if ($systemInfo) {
                    New-ComputerViewForm -systemInfo $systemInfo -OutputArea ($formItems.OutputArea.detailPanel)
                }
            }
        })

    Set-ConnectedTo -menuItem $menuStrip.Items[1] -text ("Connected to: none")
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


function New-TextBoxPopOutFormAction {
    param (
        $popOutLabel,
        $icon
    )

    $popOutLabel.Add_Click({
            if ($formItems.OutputArea.ouputTextBox.Text) {
                New-TextBoxPopOutForm -text ($formItems.OutputArea.ouputTextBox.Text) -icon (Get-Base64)
            }
        })
}

function New-CommandBoxPopOutFormAction {
    param (
    )
    #Hübscher?
    $formitems.CommandArea.popOutLabel.Add_Click({
            $commandBox = $form.controls[2].controls[2].controls[0].controls[1]
            $mlcommandBox = $form.controls[2].controls[2].controls[0].controls[2]

            $commandBox.Visible = $false
            $mlcommandBox.Text = $commandBox.Text
            $commandBox.Clear()
            $mlcommandBox.Visible = $true
        })

    $formitems.CommandArea.mlPopOutLabel.Add_Click({
            $commandBox = $form.controls[2].controls[2].controls[0].controls[1]
            $mlcommandBox = $form.controls[2].controls[2].controls[0].controls[2]

            $mlcommandBox.Visible = $false
            $commandBox.Text = $mlcommandBox.Text
            $mlcommandBox.Clear()
            $commandBox.Visible = $true
        })
}

function New-PWSHCheckBoxChange {

    $commands = Get-Command
    $global:pwshCommands = $commands | Where-Object { $_.Module -like "Microsoft.Powershell*" } | Select-Object Name

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

function New-TableCellClick {
    param (
        $table
    )

    $table.Add_CellClick({

            New-DetailViewForm -table ($formItems.Table.Table) -icon (Get-Base64) -OutputArea ($formItems.OutputArea.detailPanel)

        })


}

function New-SaveButtonAction {
    param (
        $saveButton
    )
    $saveButton.Add_Click({
            try {
                $table = ($formItems.Table.Table)
                if ($PSSession -and $PSSession.Availability -eq "Available" -and ($table.Rows[0])) {
                    $csvPath = New-SaveFileDialog -table $table
                    if ($csvPath) {
                        Set-DisplayBoxText -displayBox ($formItems.OutputArea.ouputTextBox) -text ("Data saved to $csvPath" )
                    }
                }
            }
            catch {
                $logMessage = "FAIL - Error: " + $_.Exception.Message
                Set-DisplayBoxText -displayBox ($formItems.OutputArea.ouputTextBox) -text $logMessage -isError $true
            }
        })
}

function New-ExitButtonAction {
    param (
        $form,
        $ExitButton
    )

    $ExitButton.Add_Click({

            Close-Form -form $form

        })
}

function New-F5ButtonAction {
    param (
      )


    $form.Add_KeyDown({
            if ($_.KeyCode -eq "F5") {
                Get-Table -table ($formItems.Table.Table)
            }
        })
}

#endregion WORK
Main
