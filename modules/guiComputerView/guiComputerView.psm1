<#
   .< help keyword>
   < help content>
   . . .
   #>

#region LOGGING


Set-StrictMode -Version Latest
function New-ComputerViewForm {
    param
    (
        $systemInfo,
        $OutputArea
    )

  <#    $scriptBlock = {
        param($systemInfo) #>

        Add-Type -AssemblyName System.Windows.Forms

        $detailsForm = New-Object System.Windows.Forms.Form
        $detailsForm.Size = New-Object System.Drawing.Size(400, 250)
        $detailsForm.MinimumSize = $detailsForm.Size
        $detailsForm.Name = "detailsForm"
        #$icon = "resources\icon.ico"
        #$detailsForm.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($icon)

        $detailsForm.Text = "Details"
        $detailsForm.KeyPreview = $true
        $detailsForm.Dock = 'Fill'
        $detailsForm.Padding = New-Object System.Windows.Forms.Padding(10)  # Set the padding here

        $tableLayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
        $tableLayoutPanel.Dock = "Fill"
        $tableLayoutPanel.ColumnCount = 2
        $tableLayoutPanel.RowCount = 7


        $Hostlabel = New-Object System.Windows.Forms.Label
        $Hostlabel.Text = "Hostname"
        $Hostlabel.Anchor = 'left'
        $tableLayoutPanel.Controls.Add($Hostlabel, 0, 0)

        $HostTextBox = New-Object System.Windows.Forms.TextBox
        $HostTextBox.ReadOnly = $true
        $HostTextBox.Dock = "Fill"
        $HostTextBox.Size = New-Object System.Drawing.Size(($detailsForm.Size.Width - 150), 20)
        $HostTextBox.Text = $systemInfo.OSInfo.CSName
        $tableLayoutPanel.Controls.Add($HostTextBox, 1, 0)

        $ramLabel = New-Object System.Windows.Forms.Label
        $ramLabel.Text = "RAM"
        $ramLabel.Anchor = 'left'
        $tableLayoutPanel.Controls.Add($ramLabel, 0, 1)

        $ramTextBox = New-Object System.Windows.Forms.TextBox
        $ramTextBox.ReadOnly = $true
        $ramTextBox.Dock = "Fill"
        $ramTextBox.Size = New-Object System.Drawing.Size(($detailsForm.Size.Width - 150), 20)
        $ramTextBox.Text = $systemInfo.Memory
        $tableLayoutPanel.Controls.Add($ramTextBox, 1, 1)

        $cpuabel = New-Object System.Windows.Forms.Label
        $cpuabel.Text = "CPU"
        $cpuabel.Anchor = 'left'
        $tableLayoutPanel.Controls.Add($cpuabel, 0, 2)

        $cpuTextBox = New-Object System.Windows.Forms.TextBox
        $cpuTextBox.ReadOnly = $true
        $cpuTextBox.Dock = "Fill"
        $cpuTextBox.Size = New-Object System.Drawing.Size(($detailsForm.Size.Width - 150), 20)
        $cpuTextBox.Text = $systemInfo.cpu
        $tableLayoutPanel.Controls.Add($cpuTextBox, 1, 2)

        $osnamelabel = New-Object System.Windows.Forms.Label
        $osnamelabel.Text = "OS"
        $osnamelabel.Anchor = 'left'
        $tableLayoutPanel.Controls.Add($osnamelabel, 0, 3)

        $osnameTextBox = New-Object System.Windows.Forms.TextBox
        $osnameTextBox.ReadOnly = $true
        $osnameTextBox.Dock = "Fill"
        $osnameTextBox.Size = New-Object System.Drawing.Size(($detailsForm.Size.Width - 150), 20)
        $osnameTextBox.Text = $systemInfo.OSInfo.Caption
        $tableLayoutPanel.Controls.Add($osnameTextBox, 1, 3)

        $BuildNumberlabel = New-Object System.Windows.Forms.Label
        $BuildNumberlabel.Text = "Build"
        $BuildNumberlabel.Anchor = 'left'
        $tableLayoutPanel.Controls.Add($BuildNumberlabel, 0, 4)

        $BuildNumberTextBox = New-Object System.Windows.Forms.TextBox
        $BuildNumberTextBox.ReadOnly = $true
        $BuildNumberTextBox.Dock = "Fill"
        $BuildNumberTextBox.Size = New-Object System.Drawing.Size(($detailsForm.Size.Width - 150), 20)
        $BuildNumberTextBox.Text = $systemInfo.OSInfo.BuildNumber
        $tableLayoutPanel.Controls.Add($BuildNumberTextBox, 1, 4)

        $OSArchitectureabelabel = New-Object System.Windows.Forms.Label
        $OSArchitectureabelabel.Text = "Architecure"
        $OSArchitectureabelabel.Anchor = 'left'
        $tableLayoutPanel.Controls.Add($OSArchitectureabelabel, 0, 5)

        $OSArchitectureabeTextBox = New-Object System.Windows.Forms.TextBox
        $OSArchitectureabeTextBox.ReadOnly = $true
        $OSArchitectureabeTextBox.Dock = "Fill"
        $OSArchitectureabeTextBox.Size = New-Object System.Drawing.Size(($detailsForm.Size.Width - 150), 20)
        $OSArchitectureabeTextBox.Text = $systemInfo.OSInfo.OSArchitecture
        $tableLayoutPanel.Controls.Add($OSArchitectureabeTextBox, 1, 5)
 

        foreach ($item in $OutputArea.Controls)
        {
            $OutputArea.Controls.Remove($item)
        }
        $OutputArea.Controls.Add($tableLayoutPanel)


 <#       $detailsForm.Controls.Add($tableLayoutPanel)

        $detailsForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
        $detailsForm.Add_KeyDown({
                if ($_.KeyCode -eq "Escape") {
                    $detailsForm.Close()
                }
            })

        $detailsForm.ShowDialog() | Out-Null #>
    }
<#

    $newPowerShell = [PowerShell]::Create().AddScript($scriptBlock).AddArgument($systemInfo)
    $job = $newPowerShell.BeginInvoke()
    While (-Not $job.IsCompleted) {}
    $newPowerShell.EndInvoke($job)
    $newPowerShell.Dispose()
#>
#}

Export-ModuleMember -Function New-ComputerViewForm