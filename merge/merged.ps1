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

<#
   .< help keyword>
   < help content>
   . . .
   #>

#region LOGGING

Set-StrictMode -Version Latest




function New-MenuStrip {
    param (
    )

    $menuStrip = New-Object System.Windows.Forms.MenuStrip

    $fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem
    $fileMenu.Text = "File"

    $donateButtonItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $donateButtonItem.Text = "Donate"

    $aboutButtonItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $aboutButtonItem.Text = "About"

    $exitButtonItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $exitButtonItem.Text = "Exit"

    $fileMenu.DropDownItems.Add($aboutButtonItem)
    $fileMenu.DropDownItems.Add($donateButtonItem)
    $fileMenu.DropDownItems.Add($exitButtonItem)

    $menuStrip.Items.Add($fileMenu)

    return  @{
        menuStrip = $menuStrip
    }
}

function New-ConnectArea {
    param(
        $mainPanel
    )

    # Create the top left panel
    $l0LeftPanel = New-Object System.Windows.Forms.Panel
    $l0LeftPanel.Dock = 'Top'
    #$l0LeftPanel.BackColor = 'LightGreen'
    $mainPanel.Controls.Add($l0LeftPanel, 0, 0)
    # Create the table layout panel
    $TableLayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $l0LeftPanel.Controls.Add($TableLayoutPanel)
    # Create the first label and text box
    $hostNameLabel = New-Object System.Windows.Forms.Label
    $hostNameLabel.Text = "Hostname:"
    $hostNameLabel.AutoSize = $true
    $hostNameBox = New-Object System.Windows.Forms.TextBox
    $hostNameBox.Dock = 'Fill'
    $hostNameBox.AutoCompleteMode = 'SuggestAppend' #SuggestAppend
    $hostNameBox.AutoCompleteSource = 'CustomSource'
    $hostNameBox.Size = New-Object System.Drawing.Size(200)
    # Set the autocomplete custom source
    $hostnameBoxAutoCompleteSource = New-Object System.Windows.Forms.AutoCompleteStringCollection
    $hostNameBox.AutoCompleteCustomSource = $hostnameBoxAutoCompleteSource
    # Erstellen einer SchaltflÃ¤che
    $connectButton = New-Object System.Windows.Forms.Button
    $connectButton.Text = "Connect"
    # Add the first label and text box to the table layout panel
    $TableLayoutPanel.Controls.Add($hostNameLabel, 0, 0)
    $TableLayoutPanel.Controls.Add($hostNameBox, 0, 1)
    $TableLayoutPanel.Controls.Add($connectButton, 0, 2)

    return @{
        hostNameBox   = $hostNameBox
        hostNameLabel = $hostNameLabel
        connectButton = $connectButton

    }


}

function New-OutputArea {
    param(
        $mainPanel
    )

    # Create the top panel
    $l0RightPanel = New-Object System.Windows.Forms.Panel
    $l0RightPanel.Dock = 'Top'
    #$l0RightPanel.BackColor = 'LightBlue'
    $mainPanel.Controls.Add($l0RightPanel, 1, 0)
    # Create the table layout panel
    $TableLayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $TableLayoutPanel.Dock = 'Fill'
    $l0RightPanel.Controls.Add($TableLayoutPanel)
    # Erstellen einer Anzeigefeld
    $ouputTextBox = New-Object System.Windows.Forms.RichTextBox
    $ouputTextBox.ReadOnly = $true
    $ouputTextBox.Multiline = $true
    $ouputTextBox.Dock = 'Fill'
    # Create a new Label control
    $ouputTextBoxLabel = New-Object System.Windows.Forms.Label
    $ouputTextBoxLabel.Text = "Output:"
    $ouputTextBoxLabel.AutoSize = $true
    # Create popOutLabel
    $popOutLabel = New-Object System.Windows.Forms.Label
    $popOutLabel.Text = "+"
    $popOutLabel.Font = New-Object System.Drawing.Font("System", 12)
    $popOutLabel.AutoSize = $true
    $popOutLabel.Cursor = [System.Windows.Forms.Cursors]::Hand
    $popOutLabel.Dock = 'Right'

    $ouputTextBox.Controls.Add($popOutLabel)
    # Create a tooltip for the text box
    $toolTip = New-Object System.Windows.Forms.ToolTip
    $toolTip.SetToolTip($popOutLabel, "Enlarge")
    $TableLayoutPanel.Controls.Add($ouputTextBoxLabel, 0, 0)
    $TableLayoutPanel.Controls.Add($ouputTextBox, 0, 1)

    return @{
        ouputTextBoxLabel = $ouputTextBoxLabel
        ouputTextBox      = $ouputTextBox
        popOutLabel       = $popOutLabel

    }


}

function New-TextBoxPopOutForm {
    param(
        $text
    )

    $scriptBlock = {
        Add-Type -AssemblyName System.Windows.Forms

        # Create a new form to show the enlarged text box
        $popOutForm = New-Object System.Windows.Forms.Form
        $popOutForm.AutoSize = $true
        #$popOutForm.Height = ($Height - 100)
        $popOutForm.Size = New-Object System.Drawing.Size(400, 300)
        #$popOutForm.MaximumSize = New-Object System.Drawing.Size(800, 600)
        $popOutForm.MinimumSize = New-Object System.Drawing.Size(300, 150)

        $enlargedTextBox = New-Object System.Windows.Forms.RichTextBox
        $enlargedTextBox.Text = $using:text
        $enlargedTextBox.Dock = "Fill"
        $enlargedTextBox.Multiline = $true
        $enlargedTextBox.ReadOnly = $true
        #$enlargedTextBox.Width = ($Width - 20)
        #$enlargedTextBox.Height = ($enlargedTextBox.GetLineFromCharIndex($enlargedTextBox.Text.Length) + 1) * $enlargedTextBox.Font.Height + $enlargedTextBox.Margin.Vertical
        #$enlargedTextBox.MaximumSize = New-Object System.Drawing.Size($Width, ($Height - 40))

        $popOutForm.Controls.Add($enlargedTextBox)

        $popOutForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
        $popOutForm.Add_KeyDown({
                if ($_.KeyCode -eq "Escape") {
                    $popOutForm.Close()
                }
            })
        $popOutForm.ShowDialog() | Out-Null
    }

    Start-Job -ScriptBlock $scriptBlock

}

function New-CommandArea {
    param(
        $mainPanel
    )

    # Create the bottom panel
    $l1RightPanel = New-Object System.Windows.Forms.Panel
    $l1RightPanel.Dock = 'Fill'
    $l1RightPanel.Size = New-Object System.Drawing.Size(0, 100)
    #$l1RightPanel.BackColor = 'LightGray'

    $mainPanel.Controls.Add($l1RightPanel, 1, 1)
    # Create the table layout panel
    $TableLayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $TableLayoutPanel.Dock = 'Fill'
    # create command box
    $commandBox = New-Object System.Windows.Forms.TextBox
    $commandBox.AutoCompleteMode = 'SuggestAppend' #SuggestAppend
    $commandBox.AutoCompleteSource = 'CustomSource'
    #$commandBox.Dock = 'Fill'
    #$commandBox.Multiline = $true
    $commandBox.Size = New-Object System.Drawing.Size(450)
    #$commandBox.MaximumSize = New-Object System.Drawing.Size(800)
    $commandBoxAutoCompleteSource = New-Object System.Windows.Forms.AutoCompleteStringCollection
    $commandBox.AutoCompleteCustomSource = $commandBoxAutoCompleteSource


    # Erzeuge ein FlowLayoutPanel
    $flowLayoutPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $flowLayoutPanel.Dock = "Fill"
    $flowLayoutPanel.FlowDirection = "LeftToRight"
    $flowLayoutPanel.Size = New-Object System.Drawing.Size(0, 25)

    # Create a new Label control
    $commandLabel = New-Object System.Windows.Forms.Label
    $commandLabel.Text = "Custom command:"
    $commandLabel.Anchor = "bottom"
    #$commandLabel.AutoSize = $true
    # Create the collection of radio buttons
    $radioButtonPWSH = New-Object System.Windows.Forms.RadioButton
    $radioButtonPWSH.Checked = $true
    $radioButtonPWSH.Text = "pwsh"
    $radioButtonCMD = New-Object System.Windows.Forms.RadioButton
    $radioButtonCMD.Checked = $false
    $radioButtonCMD.Text = "cmd"
    # Erstellen einer SchaltflÃ¤che
    $invokeButton = New-Object System.Windows.Forms.Button
    $invokeButton.Text = "Invoke"
    $invokeButton.Anchor = "Right, top"
    #AddingControles

    $flowLayoutPanel.Controls.Add($commandLabel)
    $flowLayoutPanel.Controls.Add($radioButtonPWSH)
    $flowLayoutPanel.Controls.Add($radioButtonCMD)
    $TableLayoutPanel.Controls.Add($flowLayoutPanel, 0, 0)
    $TableLayoutPanel.SetColumnSpan($flowLayoutPanel, 2);
    #$TableLayoutPanel.Controls.Add($commandLabel, 0, 0)
    #$TableLayoutPanel.Controls.Add($radioButtonPWSH, 1, 0)
    #$TableLayoutPanel.Controls.Add($radioButtonCMD, 1, 0)

    $TableLayoutPanel.Controls.Add($commandBox, 0, 1)
    $TableLayoutPanel.SetColumnSpan($commandBox, 2);
    $TableLayoutPanel.Controls.Add($invokeButton, 2, 1)


    $l1RightPanel.Controls.Add($TableLayoutPanel)


    return @{
        commandBox      = $commandBox
        radioButtonCMD  = $radioButtonCMD
        radioButtonPWSH = $radioButtonPWSH
        commandLabel    = $commandLabel
        invokeButton    = $invokeButton

    }

}

function New-TableAction {
    param(
        $mainPanel
    )

    #Create the bottom panel
    $l2LeftPanel = New-Object System.Windows.Forms.Panel
    $l2LeftPanel.Dock = 'top'
    #$l2LeftPanel.BackColor = 'Gray'
    $l2LeftPanel.Size = New-Object System.Drawing.Size(0, 50)
    $mainPanel.Controls.Add($l2LeftPanel, 0, 2)
    $mainPanel.SetColumnSpan($l2LeftPanel, 2);
    #$tableLayoutPanel.Controls.Add($l2l1rightPanel, 0, 2)

    # Create the table layout panel
    $TableLayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $TableLayoutPanel.Dock = 'Fill'
    # Create a new Label control
    $filterLabel = New-Object System.Windows.Forms.Label
    $filterLabel.Text = "Search:"
    $filterLabel.AutoSize = $true
    # Create a TextBox control for filtering
    $filterTextBox = New-Object System.Windows.Forms.TextBox
    $filterTextBox.Dock = 'Fill'
    $filterTextBox.Size = New-Object System.Drawing.Size(200)
    # Button a TextBox control for filtering
    $uninstallButton = New-Object System.Windows.Forms.Button
    $uninstallButton.Text = "Uninstall"
    $uninstallButton.Anchor = "Right, top"


    $TableLayoutPanel.Controls.Add($filterLabel, 0, 0)

    $TableLayoutPanel.Controls.Add($filterTextBox, 0, 1)
    $TableLayoutPanel.Controls.Add($uninstallButton, 3, 1)

    $TableLayoutPanel.SetColumnSpan($filterTextBox, 2);
    $l2LeftPanel.Controls.Add($TableLayoutPanel)

    return   @{
        filterTextBox   = $filterTextBox
        filterLabel     = $filterLabel
        uninstallButton = $uninstallButton
    }
}

function New-Table {
    param(
        $mainPanel
    )


    # Create the bottom panel
    $l3LeftPanel = New-Object System.Windows.Forms.Panel
    $l3LeftPanel.Dock = 'Fill'
    $l3LeftPanel.Size = New-Object System.Drawing.Size(700, 200)
    #$l3LeftPanel.BackColor = 'Red'

    $mainPanel.Controls.Add($l3LeftPanel, 0, 3)
    $mainPanel.SetColumnSpan($l3LeftPanel, 2);
    # Create the table layout panel
    $TableLayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $TableLayoutPanel.Dock = 'Fill'


    $l3LeftPanel.Controls.Add($TableLayoutPanel)
    # Erstellen einer Tabelle
    $table = New-Object System.Windows.Forms.DataGridView
    $table.Anchor = 'Top, Bottom, Left, Right'
    #$table.ScrollBars = "Vertical"
    $table.SelectionMode = "FullRowSelect"
    $table.AutoSizeColumnsMode = "Fill"

    $table.Anchor = "Top, Left, Bottom, Right"
    #$table.AutoSize = $true
    $table.MultiSelect = $false  # Nur eine Zeile auswÃ¤hlbar
    $table.ReadOnly = $true
    $table = Add-TableContent -table $table

    $TableLayoutPanel.Controls.Add($table, 0, 0)


    return   @{
        table = $table
    }

}
function Add-TableContent {
    param (
        $table
    )
    # HinzufÃ¼gen von Spalten zur Tabelle
    $colDisplayName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colDisplayName.HeaderText = "Name"
    $colDisplayName.Name = "Name"
    $colDisplayName.MinimumWidth = 200
    $colDisplayName.FillWeight = 200
    $table.Columns.Add($colDisplayName)  | Out-Null

    # HinzufÃ¼gen von Spalten zur Tabelle
    $colUninstallString = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colUninstallString.HeaderText = "Uninstallstring"
    $colUninstallString.Name = "Uninstallstring"
    #$column2.Width = 500
    $table.Columns.Add($colUninstallString)  | Out-Null
    $table.Columns["Uninstallstring"].Visible = $false

    # HinzufÃ¼gen von Spalten zur Tabelle
    $colVersion = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colVersion.HeaderText = "Version"
    $colVersion.Name = "Version"
    $colVersion.MinimumWidth = 150
    $colVersion.FillWeight = 100
    $table.Columns.Add($colVersion)  | Out-Null
    $table.Columns["Version"].Visible = $true

    # HinzufÃ¼gen von Spalten zur Tabelle
    $colPub = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colPub.HeaderText = "Publisher"
    $colPub.MinimumWidth = 160
    $colPub.FillWeight = 150
    $colPub.Name = "Publisher"
    $table.Columns.Add($colPub)  | Out-Null
    $table.Columns["Publisher"].Visible = $true
    $table.Columns["Publisher"].DisplayIndex = 2;


    # HinzufÃ¼gen von Spalten zur Tabelle
    $colType = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colType.HeaderText = "Type"
    $colType.MinimumWidth = 50
    $colType.FillWeight = 50
    $colType.Name = "Type"
    $table.Columns.Add($colType)  | Out-Null
    $table.Columns["Type"].Visible = $false

    # HinzufÃ¼gen von Spalten zur Tabelle
    $colContext = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colContext.HeaderText = "Context"
    $colContext.Name = "Context"
    $colContext.MinimumWidth = 50
    $colContext.FillWeight = 50
    $table.Columns.Add($colContext)  | Out-Null
    $table.Columns["Context"].Visible = $true

    return $table
}


function Update-TableContent {
    # Parameter help description
    param(
        $tableContent,
        $table
    )
    # Datenquelle leeren
    $table.Rows.Clear()

    foreach ($item in $tableContent) {
        if ($item.Type -ne "N/A") {
            $row = New-Object System.Windows.Forms.DataGridViewRow
            $row.CreateCells($table)
            $row.Cells[0].Value = $item.DisplayName
            $row.Cells[1].Value = $item.QuietUninstallString
            $row.Cells[2].Value = $item.DisplayVersion
            $row.Cells[3].Value = $item.Publisher
            $row.Cells[4].Value = $item.Type
            $row.Cells[5].Value = $item.Context
            $table.Rows.Add($row) | Out-Null
        }

    }

    # Datenquelle nach der Spalte "Name" sortieren
    $table.Sort($table.Columns[0], [System.ComponentModel.ListSortDirection]::Ascending)

}

function New-ExitButton {
    param(
        $mainPanel
    )


    # Create the bottom panel
    $l4RightPanel = New-Object System.Windows.Forms.Panel
    $l4RightPanel.Dock = 'Bottom'
    #$l4RightPanel.BackColor = 'green'
    $l4RightPanel.Size = New-Object System.Drawing.Size(0, 25)
    $mainPanel.Controls.Add($l4RightPanel, 1, 4)
    #$mainPanel.SetColumnSpan($l4RightPanel, 2)
    # Create the table layout panel
    $TableLayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $TableLayoutPanel.Dock = 'Fill'
    # Add Buttons
    $exitButton = New-Object System.Windows.Forms.Button
    $exitButton.Text = "Exit"
    $exitButton.Anchor = "right, bottom"

    $TableLayoutPanel.Controls.Add($exitButton, 0, 0)
    $l4RightPanel.Controls.Add($TableLayoutPanel)

    return   @{
        exitButton = $exitButton
    }

}


function New-ProgressBar {
    param (

    )
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Dock = "Bottom"
    $progressBar.MarqueeAnimationSpeed = 20

    return   @{
        progressBar = $progressBar
    }
}




# Show the form
 <#
   .< help keyword>
   < help content>
   . . .
   #>

#region LOGGING

Set-StrictMode -Version Latest


Function Get-InstalledApp {
    param(
        [parameter(Mandatory = $false)][string]$LogText,
        [Parameter(Position = 0)]$AppName
    )
    try {
        if (!$AppName) {
            $AppName = "*"
        }

        $AppVersion = $null

        $32BitUninstallObejct = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* `
        | Select-Object  DisplayName, DisplayVersion, Publisher, @{Name = "ProductID"; Expression = { $_.PSChildName } }, UninstallString, QuietUninstallString, Type, @{Name = "Context"; Expression = { "x32" } } `
        | Where-Object { if ($AppVersion) { $_.DisplayName -like $AppName -AND $_.DisplayVersion -like $AppVersion } else { $_.DisplayName -like $AppName } }

        $64BitUninstallObejct = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* `
        | Select-Object DisplayName, DisplayVersion, Publisher, @{Name = "ProductID"; Expression = { $_.PSChildName } }, UninstallString, QuietUninstallString, Type, @{Name = "Context"; Expression = { "x64" } } `
        | Where-Object { if ($AppVersion) { $_.DisplayName -like $AppName -AND $_.DisplayVersion -like $AppVersion } else { $_.DisplayName -like $AppName } }

        $UninstallObejctUser = Get-ItemProperty HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* `
        | Select-Object DisplayName, DisplayVersion, Publisher, @{Name = "ProductID"; Expression = { $_.PSChildName } }, UninstallString, QuietUninstallString, Type, @{Name = "Context"; Expression = { "User" } } `
        | Where-Object { if ($AppVersion) { $_.DisplayName -like $AppName -AND $_.DisplayVersion -like $AppVersion } else { $_.DisplayName -like $AppName } }


        if ($64BitUninstallObejct -or $32BitUninstallObejct -or $UninstallObejctUser) {


            $UninstallObejcts = (@($32BitUninstallObejct) + @($64BitUninstallObejct) + @($UninstallObejctUser))

            foreach ($object in $UninstallObejcts) {
                if ($object.QuietUninstallString) {
                    $object.Type = "Quiet"
                }
                elseif ($object.UninstallString -like "msiexec*" -AND $object.ProductID -like "{*}") {
                    $object.Type = "MSI"
                    $object.QuietUninstallString = "/qn /x " + '"' + $object.ProductID + '"' + " /norestart"
                }
                else {
                    $object.Type = "N/A"
                }


            }
            return $UninstallObejcts
        }
        else {

            #Cleanup -exitCode 1
        }
    }
    catch {
        $logMessage = "FAIL - $LogText - Error at line " + $_.InvocationInfo.ScriptLineNumber + ": " + $_.Exception.Message
        #Cleanup -exitCode 1
    }
}


function Invoke-KillProcess {
    param (
        $displayName
    )
    Invoke-Command -Session $PSSession -ScriptBlock `
    {(Get-Process | Where-Object {$using:displayName -like "*"+ $_.ProcessName +"*"} | Stop-Process -Force -Passthru)} -ErrorAction SilentlyContinue

}


Function Invoke-SilentUninstallString {
    param(
        $command,
        $progressBar,
        $type
    )
    try {
        if ($type -like "MSI") {

            $job = Invoke-Command -Session $PSSession -ScriptBlock `
            { (Start-Process -FilePath msiexec.exe -ArgumentList $using:command -Wait -Passthru -WindowStyle Hidden ) } -AsJob
        }
        elseif ($type -like "Quiet") {
            $regex = '"(.*?)"\s(.*)'
            $match = $command -split $regex

            $executable = $match[1]
            $argument = $match[2]

            $job = Invoke-Command -Session $PSSession -ScriptBlock `
            { (Start-Process -FilePath $using:executable -ArgumentList $using:argument -Wait -Passthru -WindowStyle Hidden ) } -AsJob

        }

        else {
            return 1
        }

        $progressBarStatus = Set-ProgressBar -progressBar $progressBar -job $job
        if ($progressBarStatus -ne 124) {
            return ($job.ChildJobs[0].Output[0].ExitCode)
        }
        else { return $progressBarStatus }
    }
    catch {
        $logMessage = "FAIL - Error at line " + $_.InvocationInfo.ScriptLineNumber + ": " + $_.Exception.Message
    }
}

function Invoke-CustomCommand {
    param (
        $command,
        $progressBar,
        $type,
        $radioButtonPWSH
    )

    if ($type -like "c_command") {
        if ($radioButtonPWSH -eq $true) {
            #$command = "'" + $command + "'"
            $job = Invoke-Command  -Session $PSSession -ScriptBlock `
            { (powershell.exe -WindowStyle hidden "$using:command" 2>&1 ) } -AsJob
        }
        else {
            $job = Invoke-Command  -Session $PSSession -ScriptBlock `
            { (cmd.exe /c "$using:command" 2>&1 ) } -AsJob
        }

        $progressBarStatus = Set-ProgressBar -progressBar $progressBar -job $job
        if ($progressBarStatus -ne 124) {
            $output = $job.ChildJobs[0].Output | Out-String

            $decodedString = [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::Default.GetBytes($output))

            return $decodedString
        }
        else { return $progressBarStatus }
    }
    else {
        return 1
    }

}

Function Set-ProgressBar {
    param (
        $job,
        $progressBar
    )
    $stopwatch = [Diagnostics.Stopwatch]::StartNew()

    while ($job.State -ne 'Completed') {
        $progressBar.Style = "Marquee"
        [System.Windows.Forms.Application]::DoEvents()

        #auf 60 erhÃ¶hren nach dem tests, timeout
        if ($stopwatch.Elapsed.TotalSeconds -gt 60) {
            $stopwatch.Stop()
            $progressBar.Style = "Continuous"
            Get-Job | Stop-Job | Remove-Job -Force -ErrorAction SilentlyContinue
            return 124
        }
    }

    $stopwatch.Stop()
    $progressBar.Style = "Continuous"

}


Function Get-Table {
    param( $table
    )

    #funktioniert noch nicht 100%
    if ($PSSession -and $PSSession.Availability -eq "Available") {

        $table.Rows.Clear()

        #($formItems.SearchBox.filterTextBox).Clear()
        $InstalledApps = Invoke-Command -Session $PSSession -ScriptBlock ${function:Get-InstalledApp}
        Update-TableContent -tableContent  $InstalledApps -table $table

    }
}

Function Set-DisplayBoxText {
    param (
        $displayBox,
        $text
    )
    $displayBox.Clear()
    $displayBox.AppendText($text)

}

function Set-LastEntrie {
    param (
        $text,
        $lastEntries,
        $autoCompleteSource
    )

    #$text = $text.Trim()

    # HinzufÃ¼gen der Eingabe zu den letzten Eingaben
    if (-not [string]::IsNullOrWhiteSpace($text) -and (-not $lastEntries.Contains($text))) {
        $lastEntries.Add($text)

        # Begrenzung der Anzahl der letzten Eingaben auf fÃ¼nf
        if ($lastEntries.Count -gt 10) {
            $lastEntries.RemoveAt(0)
        }

        Update-AutoCompleteSource -autoCompleteSource $autoCompleteSource -lastEntries $lastEntries
    }

}

function Set-PWSHCommandAutoComplete {
    param (
        $autoCompleteSource,
        $pwshCommands
    )
    #wird jedes mal neu aufgerufen?
    $filteredCommands = @();
    foreach ($item in $pwshCommands) {
        $filteredCommands += $item.Name
    }
    $autoCompleteSource.AddRange($filteredCommands)

}

function Update-AutoCompleteSource {
    param (
        $autoCompleteSource,
        $lastEntries
    )


    $autoCompleteSource.Clear()
    $autoCompleteSource.AddRange($lastEntries)

}

 <#
   .< help keyword>
   < help content>
   . . .
   #>

#region LOGGING


Set-StrictMode -Version Latest
Add-Type -AssemblyName System.Windows.Forms
function Show-Detail {
    param
    (
        $table
    )

    $selectedRow = $table.selectedRows[0]

    $scriptBlock = {
    Add-Type -AssemblyName System.Windows.Forms

    #$selectedRows = $table.selectedRows



    $detailsForm = New-Object System.Windows.Forms.Form
    $detailsForm.Size = New-Object System.Drawing.Size(400, 250)
    $detailsForm.MinimumSize = $detailsForm.Size
   #$detailsForm.MaximumSize = New-Object System.Drawing.Size(800, 600)

    $detailsForm.Text = "Details"
    $detailsForm.KeyPreview = $true
    $detailsForm.Dock = 'Fill'
    $detailsForm.Padding = New-Object System.Windows.Forms.Padding(10)  # Set the padding here

    $tableLayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $tableLayoutPanel.Dock = "Fill"
    $tableLayoutPanel.ColumnCount = 2
    $tableLayoutPanel.RowCount = 3

    $nameLabel = New-Object System.Windows.Forms.Label
    $nameLabel.Text = "Name"
    $nameLabel.Anchor = 'left'

    $tableLayoutPanel.Controls.Add($nameLabel, 0, 0)

    $nameTextBox = New-Object System.Windows.Forms.TextBox
    $nameTextBox.ReadOnly = $true
    $nameTextBox.Dock = "Fill"

    $nameTextBox.Size = New-Object System.Drawing.Size(($detailsForm.Size.Width - 150), 20)
    $nameTextBox.Text = $using:selectedRow.Cells["Name"].Value
    $tableLayoutPanel.Controls.Add($nameTextBox, 1, 0)

    $publisherLabel = New-Object System.Windows.Forms.Label
    $publisherLabel.Text = "Publisher"
    $publisherLabel.Anchor = 'left'
    $tableLayoutPanel.Controls.Add($publisherLabel, 0, 1)

    $publisherTextBox = New-Object System.Windows.Forms.TextBox
    $publisherTextBox.ReadOnly = $true
    $publisherTextBox.Dock = "Fill"
    $publisherTextBox.Text = $using:selectedRow.Cells["Publisher"].Value
    $publisherTextBox.Size = New-Object System.Drawing.Size(($detailsForm.Size.Width  - 150), 20)
    $tableLayoutPanel.Controls.Add($publisherTextBox, 1, 1)

    $versionLabel = New-Object System.Windows.Forms.Label
    $versionLabel.Text = "Version"
    $versionLabel.Anchor = 'left'
    $tableLayoutPanel.Controls.Add($versionLabel, 0, 2)

    $versionTextBox = New-Object System.Windows.Forms.TextBox
    $versionTextBox.ReadOnly = $true
    $versionTextBox.Dock = "Fill"
    $versionTextBox.Size = New-Object System.Drawing.Size(($detailsForm.Size.Width  - 150), 20)
    $versionTextBox.text = $using:selectedRow.Cells["Version"].Value
    $tableLayoutPanel.Controls.Add($versionTextBox, 1, 2)

    $stringLabel = New-Object System.Windows.Forms.Label
    $stringLabel.Text = "Command"
    $stringLabel.Anchor = 'left'
    $tableLayoutPanel.Controls.Add($stringLabel, 0, 3)

    $stringTextBox = New-Object System.Windows.Forms.TextBox
    $stringTextBox.ReadOnly = $true
    $stringTextBox.Dock = "Fill"
    $stringTextBox.Text = if ($using:selectedRow.Cells[4].Value -eq "MSI") {
        "msiexec " + $using:selectedRow.Cells["Uninstallstring"].Value
    }
    else { $using:selectedRow.Cells["Uninstallstring"].Value }
    $stringTextBox.Size = New-Object System.Drawing.Size(($detailsForm.Size.Width  - 150), 50)
    $stringTextBox.Multiline = $true
    $tableLayoutPanel.Controls.Add($stringTextBox, 1, 3)

    $stringLabel = New-Object System.Windows.Forms.Label
    $stringLabel.Text = "Context"
    $tableLayoutPanel.Controls.Add($stringLabel, 0, 4)

    $stringTextBox = New-Object System.Windows.Forms.TextBox
    $stringTextBox.ReadOnly = $true
    $stringTextBox.Dock = "Fill"
    $stringTextBox.Text = $using:selectedRow.Cells["Context"].Value
    $stringTextBox.Size = New-Object System.Drawing.Size(($detailsForm.Size.Width - 150), 20)
    $tableLayoutPanel.Controls.Add($stringTextBox, 1, 4)

    $detailsForm.Controls.Add($tableLayoutPanel)

    $detailsForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $detailsForm.Add_KeyDown({
            if ($_.KeyCode -eq "Escape") {
                $detailsForm.Close()
            }
        })

    $detailsForm.ShowDialog() | Out-Null


   }

   $job = Start-Job -ScriptBlock $scriptBlock

}


 <#
   .< help keyword>
   < help content>
   . . .
   #>

#region LOGGING


Set-StrictMode -Version Latest
function Test-PreRequirement {
    param (
        $ouputTextBox,
        $requiredVersion,
        $requireAdminRights,
        $requiredPolicy
    )
    if ((Test-IsRequiredPSVersion -requiredVersion $requiredVersion )) {
        if ((Test-IsAdminRole -requireAdminRights $requireAdminRights)) {
            if ((Test-IsCorrectExecutionPolicy -requiredPolicy $requiredPolicy)) {
                return $true
            }
            else {
                Set-DisplayBoxText -displayBox $ouputTextBox -text "Please set execution policy to $requiredPolicy"
                return $false
            }
        }
        else {
            Set-DisplayBoxText -displayBox $ouputTextBox -text "Please run this program as an administrator."
            return $false
        }
    }
    else {
        Set-DisplayBoxText -displayBox $ouputTextBox -text "Please install the powershell version $requiredVersion"
        return $false
    }
}

Function Test-IsCorrectExecutionPolicy {
    param (
        $requiredPolicy
    )
    if ((Get-ExecutionPolicy) -ne $requiredPolicy) {
        try {
            Set-ExecutionPolicy -ExecutionPolicy $requiredPolicy -Scope CurrentUser -Force
            return $true
        }
        catch {
            return $false
        }
    }
    else {
        return $true
    }
}

function Test-IsAdminRole {
    param ($requireAdminRights)

    if ($requireAdminRights) {
        $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
        return $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    else {
        $true
    }

}

function Test-IsRequiredPSVersion {
    param (
        $requiredVersion
    )
    # Erforderliche Mindestversion von PowerShell
    $requiredVersion = New-Object System.Version($requiredVersion)

    # Aktuelle Version von PowerShell
    $currentVersion = $PSVersionTable.PSVersion

    # ÃœberprÃ¼fen, ob die aktuelle Version grÃ¶ÃŸer oder gleich der erforderlichen Version ist
    if ($currentVersion -ge $requiredVersion) {
        return $true
    }
    else {
        return $false

    }

}



Main
