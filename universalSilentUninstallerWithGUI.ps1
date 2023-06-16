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
#param ( 
# [parameter(Mandatory = $false)][string]$AppName,
# [parameter(Mandatory = $false)][string]$AppVersion
#)
#$RepoDir = "C:\Program Files (x86)\LANDesk\LDClient\sdmcache\SWD\UniversalUninstaller" 
$ScriptDir = $PSScriptRoot
$ModuleDir = $ScriptDir + "\modules\" 
#region ******************** TESTING ********************
$testing = $true
if ($testing) {
    $targetComputer = "localhost" 
    #for testing
    # localhost HAMV21738

}
#endregion TESTING
#region ******************** MODULES ********************
$GuiModule = $ModuleDir + "gui\dyngui.psm1"
$FunctionsModule = $ModuleDir + "functions\functions.psm1"
$GuiDetailViewModule = $ModuleDir + "guiDetailView\guiDetailView.psm1"
$ValidationModule = $ModuleDir + "validation\validation.psm1"

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
        Import-Module $GuiModule
        Import-Module $ValidationModule
        Import-Module $FunctionsModule
        Import-Module $GuiDetailViewModule
        #check if the init-folders exist - if not create them
        $InitFolders = @($ScriptDir)
        Foreach ($Folder in $InitFolders) { If (!(test-path $Folder)) { New-Item -ItemType Directory -Force -Path $Folder } }
        #CreateLogEntry -logMessage "OK - $LogText" 
    }
    catch {
        Cleanup -exitCode 1
        #"FATAL - initialisation not passed - Error at line " + $_.InvocationInfo.ScriptLineNumber + ": " + $_.Exception.Message | Out-File -FilePath ($MyInvocation.PSCommandPath + ".log") -Force
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
        Add-FormActions -formItems $formItems -form $form
        $form.ShowDialog()
        
    }
    catch {
        $logMessage = "FAIL - $LogText - Error at line " + $_.InvocationInfo.ScriptLineNumber + ": " + $_.Exception.Message
        Write-host $logMessage 
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
    $form.Text = "Universal Uninstaller"
    $form.KeyPreview = $true
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    # Create the table layout panel
    $mainPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $mainPanel.Dock = 'Fill'
    #würd ich gern auf bottom und top beschränken
    $mainPanel.Padding = New-Object System.Windows.Forms.Padding(25)  # Set the padding here

    #Warum ?
    #alles noch sehr unschön hier
    $MenuStrip = New-MenuStrip 
    <# $ConnectArea  = New-ConnectArea -mainPanel $mainPanel
    $OutputArea   = New-OutputArea -mainPanel $mainPanel
    $CommandArea  = New-CommandArea -mainPanel $mainPanel
    $TableActions = New-TableActions -mainPanel $mainPanel
    $Table        = New-Table -mainPanel $mainPanel
    $ExitButton   = New-ExitButton -mainPanel $mainPanel#>

    $formItems = @{
        ConnectArea  = New-ConnectArea -mainPanel $mainPanel
        OutputArea   = New-OutputArea -mainPanel $mainPanel
        CommandArea  = New-CommandArea -mainPanel $mainPanel
        TableActions = New-TableActions -mainPanel $mainPanel
        Table        = New-Table -mainPanel $mainPanel
        ExitButton   = New-ExitButton -mainPanel $mainPanel
        ProgressBar  = New-ProgressBar
        MenuStrip    = $MenuStrip[2] #Warum ?
     
    }
    $formItems2 = @{
       
        ProgressBar = New-ProgressBar
        MenuStrip   = $MenuStrip[2] #Warum ?
     
    }
    # Hinzufügen von Schaltflächen zum Formular
    foreach ($item in $formItems2.GetEnumerator()) {
        foreach ($formItem in $item.Value) {
            foreach ($item in $formItem.GetEnumerator()) {
                $form.Controls.Add($item.Value)
            }
        }
    }
  
    $form.Controls.Add($mainPanel)
    #zum testen
    #$form.ShowDialog()

    return $form, $formItems
}

function Add-FormActions {
    param (
        $formItems,
        $form
    )

    New-ExitButtonAction -ExitButton ($formItems.ExitButton.ExitButton) -form $form -menuStrip ($formItems.menuStrip.menuStrip) 

    if ((Test-PreRequirements -ouputTextBox ($formItems.OutputArea.ouputTextBox) -requiredVersion "5.1" -requireAdminRights $true -requiredPolicy "bypass")) {
   
        New-SearchBoxAction -filterTextBox ($formItems.TableActions.filterTextBox) 
        New-TableAction  -table ($formItems.Table.Table)
        New-ConectButtonAction -ConectButton ($formItems.ConnectArea.ConnectButton) 
        New-UninstallButtonAction -UninstallButton ($formItems.TableActions.UninstallButton) 
        New-F5ButtonAction -form $form -table ($formItems.Table.Table)
        New-PWSHCheckBoxChange 
        New-InvokeButtonAction -InvokeButton ($formItems.CommandArea.invokeButton) 
        New-TextBoxPopOutFormAction -popOutLabel ($formItems.OutputArea.popOutLabel)
    }
}

<#
function Test-PreRequirements {
    param (
        $ouputTextBox
    )

    if ((Test-IsRequiredPSVersion)) {
        Set-DisplayBoxText -displayBox $ouputTextBox -text "Ok IsRequiredPSVersion"


        if ((Test-IsAdminRole)) {
            Set-DisplayBoxText -displayBox $ouputTextBox -text "Ok is admin"

            if ((Test-IsCorrectExecutionPolicy)) {
                Set-DisplayBoxText -displayBox $ouputTextBox -text "Ok IsCorrectExecutionPolicy"
        
            }
            else {
                Set-DisplayBoxText -displayBox $ouputTextBox -text "Please set IsCorrectExecutionPolicy"  
            }        
        }
        else {
            Set-DisplayBoxText -displayBox $ouputTextBox -text "Please run this program as an administrator."  
        }
    }
    else {
        Set-DisplayBoxText -displayBox $ouputTextBox -text "Please set IsRequiredPSVersion"  
    }

    
}#>

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
                        $PSSession = New-PSSession -ComputerName $targetComputer -Credential "emea\rootmb"
                    }
                    if ($PSSession) {
    
                        Set-DisplayBoxText -displayBox $displaybox -text ("Successfully connected to $targetComputer.")
                        $InstalledApps = Invoke-Command -Session $PSSession -ScriptBlock ${function:Get-InstalledApps}
                        Update-TableContent -tableContent $InstalledApps -table ($formItems.Table.Table)
                    }
                    else {
                        Set-DisplayBoxText -displayBox $displaybox -text ($error[0].ErrorDetails)
                    }

                  
                   
                    $formItems.CommandArea.commandBox.Clear()
                    $formItems.TableActions.filterTextBox.Clear()
                    Set-LastEntries -text $displaybox.Text -lastEntries $ConnectTextlastEntries -autoCompleteSource ($formItems.ConnectArea.HostnameBox.AutoCompleteCustomSource)
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
    
                    Set-LastEntries -text $formItems.CommandArea.commandBox.Text -lastEntries $CommandlastEntries -autoCompleteSource $formItems.CommandArea.commandBox.AutoCompleteCustomSource

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

function New-TableAction {
    param (
        $table
    )
    $table.Add_CellDoubleClick({
            Show-Details -table ($formItems.Table.Table)
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

                            $testApp = (Invoke-Command -Session $PSSession -ScriptBlock ${function:Get-InstalledApps} -ArgumentList $selectedItemName)
                      
                            if ($exitcode -eq 0 -and !$testApp) {
                                [System.Windows.Forms.MessageBox]::Show("Uninstallation successful.", "Ok", "OK", "Information")
                                Set-DisplayBoxText -displayBox ($formItems.ConnectArea.connectDisplayBox) -text $text
                                Get-Table -formItems $formItems
                      
                            }
                            else {
                                [System.Windows.Forms.MessageBox]::Show("Uninstallation not successful.", "Error", "OK", "Hand")
                                Set-DisplayBoxText -displayBox ($formItems.ConnectArea.connectDisplayBox) -text $text
                                Get-Table -formItems $formItems
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
                Set-DisplayBoxText -displayBox ($formItems.ConnectArea.connectDisplayBox) -text $logMessage
            }
        })   
}

function New-ExitButtonAction {
    param (
        $form,
        $ExitButton,
        $menuStrip
    )
    $ExitButton.Add_Click({ 
            $form.Close()
            #Cleanup
        })
    #Schöner!
    $menuStrip.Items[0].DropDownItems[0].Add_Click({ 
            $form.Close() 
            #Cleanup
        })
  
}

function New-PWSHCheckBoxChange {
    Set-PWSHCommandAutoComplete -autoCompleteSource ($formItems.CommandArea.commandBox.AutoCompleteCustomSource)
    ($formItems.CommandArea.radioButtonPWSH).add_CheckedChanged({ 

            if (($formItems.CommandArea.radioButtonPWSH).Checked) {
                Set-PWSHCommandAutoComplete -autoCompleteSource ($formItems.CommandArea.commandBox.AutoCompleteCustomSource)
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
        $form
       
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
