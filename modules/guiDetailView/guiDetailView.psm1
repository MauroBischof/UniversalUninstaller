<#
   .< help keyword>
   < help content>
   . . .
   #>

#region LOGGING


Set-StrictMode -Version Latest
function New-DetailViewForm {
    param
    (
        $table,
        $OutputArea
    )


    $selectedRow = $table.selectedRows[0]
    Add-Type -AssemblyName System.Windows.Forms





    $detailsForm = New-Object System.Windows.Forms.Form
    $detailsForm.Size = New-Object System.Drawing.Size(400, 250)
    $detailsForm.MinimumSize = $detailsForm.Size
    $detailsForm.Name = "detailsForm"

    $detailsForm.Text = "Details"
    $detailsForm.KeyPreview = $true
    $detailsForm.Dock = 'Fill'
    $detailsForm.Padding = New-Object System.Windows.Forms.Padding(10) 

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
    $nameTextBox.Text = $selectedRow.Cells["Name"].Value
    $tableLayoutPanel.Controls.Add($nameTextBox, 1, 0)

    $publisherLabel = New-Object System.Windows.Forms.Label
    $publisherLabel.Text = "Publisher"
    $publisherLabel.Anchor = 'left'
    $tableLayoutPanel.Controls.Add($publisherLabel, 0, 1)

    $publisherTextBox = New-Object System.Windows.Forms.TextBox
    $publisherTextBox.ReadOnly = $true
    $publisherTextBox.Dock = "Fill"
    $publisherTextBox.Text = $selectedRow.Cells["Publisher"].Value
    $publisherTextBox.Size = New-Object System.Drawing.Size(($detailsForm.Size.Width - 150), 20)
    $tableLayoutPanel.Controls.Add($publisherTextBox, 1, 1)

    $versionLabel = New-Object System.Windows.Forms.Label
    $versionLabel.Text = "Version"
    $versionLabel.Anchor = 'left'
    $tableLayoutPanel.Controls.Add($versionLabel, 0, 2)

    $versionTextBox = New-Object System.Windows.Forms.TextBox
    $versionTextBox.ReadOnly = $true
    $versionTextBox.Dock = "Fill"
    $versionTextBox.Size = New-Object System.Drawing.Size(($detailsForm.Size.Width - 150), 20)
    $versionTextBox.text = $selectedRow.Cells["Version"].Value
    $tableLayoutPanel.Controls.Add($versionTextBox, 1, 2)

    $stringLabel = New-Object System.Windows.Forms.Label
    $stringLabel.Text = "Command"
    $stringLabel.Anchor = 'left'
    $tableLayoutPanel.Controls.Add($stringLabel, 0, 3)

    $stringTextBox = New-Object System.Windows.Forms.TextBox
    $stringTextBox.ReadOnly = $true
    $stringTextBox.Dock = "Fill"
    $stringTextBox.Text = if ($selectedRow.Cells[4].Value -eq "MSI") {
        "msiexec " + $selectedRow.Cells["Uninstallstring"].Value
    }
    else { $selectedRow.Cells["Uninstallstring"].Value }
    $stringTextBox.Size = New-Object System.Drawing.Size(($detailsForm.Size.Width - 150), 50)
    $stringTextBox.Multiline = $true
    $tableLayoutPanel.Controls.Add($stringTextBox, 1, 3)

    $stringLabel = New-Object System.Windows.Forms.Label
    $stringLabel.Text = "Context"
    $tableLayoutPanel.Controls.Add($stringLabel, 0, 4)

    $stringTextBox = New-Object System.Windows.Forms.TextBox
    $stringTextBox.ReadOnly = $true
    $stringTextBox.Dock = "Fill"
    $stringTextBox.Text = $selectedRow.Cells["Context"].Value
    $stringTextBox.Size = New-Object System.Drawing.Size(($detailsForm.Size.Width - 150), 20)
    $tableLayoutPanel.Controls.Add($stringTextBox, 1, 4)

    foreach ($item in $OutputArea.Controls)
    {
        $OutputArea.Controls.Remove($item)
    }
    $OutputArea.Controls.Add($tableLayoutPanel)



}


Export-ModuleMember -Function New-DetailViewForm