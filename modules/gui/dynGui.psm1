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
    $menuStrip.Name = "MenuStrip"
    $menuStrip.BackColor = 'AliceBlue'

    $fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem
    $fileMenu.Text = "File"
    $fileMenu.Name = "FileMenu"

    $connectedTo = New-Object System.Windows.Forms.ToolStripMenuItem
    #$connectedTo.Enabled = $false
    $connectedTo.Text = "connectedTo"
    $connectedTo.Name = "connectedTo"

    $connectedTo.Alignment = "Right"

    $donateButtonItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $donateButtonItem.Text = "Donate"
    $donateButtonItem.Name = "donateButton"
    $donateButtonItem.Visible = $False

    $aboutButtonItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $aboutButtonItem.Text = "About"
    $aboutButtonItem.Name = "aboutButton"
    $aboutButtonItem.Visible = $False

    $exitButtonItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $exitButtonItem.Text = "Exit"
    $exitButtonItem.Name = "exitButton"

    $fileMenu.DropDownItems.Add($aboutButtonItem)
    $fileMenu.DropDownItems.Add($donateButtonItem)
    $fileMenu.DropDownItems.Add($exitButtonItem)

    $menuStrip.Items.Add($fileMenu)
    $menuStrip.Items.Add($connectedTo)

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
    $l0LeftPanel.Name = "ConnectArea"
    $l0LeftPanel.Size = New-Object System.Drawing.Size(700, 80)
    #$l0LeftPanel.BackColor = 'Lightgray'
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
    # Erstellen einer Schaltfläche
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
    $00RightPanel = New-Object System.Windows.Forms.Panel
    $00RightPanel.Dock = 'Fill'
    $00RightPanel.Name = 'OutputArea'
    #$00RightPanel.Size = New-Object System.Drawing.Size(0, 80)

   # $00RightPanel.BackColor = 'LightBlue'
    #$mainPanel.Controls.Add($l00RightPanel, 1, 0)
    $mainPanel.SetRowSpan($00RightPanel, 5)
    $mainPanel.Controls.Add($00RightPanel, 2, 0)
    # Create the table layout panel
    $TableLayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel

    $TableLayoutPanel.Dock = 'Fill'
    $00RightPanel.Controls.Add($TableLayoutPanel)
    # Erstellen einer Anzeigefeld
    $ouputTextBox = New-Object System.Windows.Forms.RichTextBox
    $ouputTextBox.ReadOnly = $true
    $ouputTextBox.Multiline = $true
    $ouputTextBox.Dock = 'Fill'
    $ouputTextBox.Margin = '0,0,0,15'
    $ouputTextBox.Size = New-Object System.Drawing.Size(0, 220)
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

    $detailPanel = New-Object System.Windows.Forms.Panel
    $detailPanel.BackColor = 'lightgray'
    $detailPanel.Dock = 'Fill'
    $TableLayoutPanel.Controls.Add($detailPanel, 0, 4)

    return @{
        ouputTextBoxLabel = $ouputTextBoxLabel
        ouputTextBox      = $ouputTextBox
        popOutLabel       = $popOutLabel
        detailPanel       = $detailPanel
    }


}

function New-TextBoxPopOutForm {
    param(
        $text
    )

    $scriptBlock = {
        param($text)
        Add-Type -AssemblyName System.Windows.Forms

        # Create a new form to show the enlarged text box
        $popOutForm = New-Object System.Windows.Forms.Form
        $popOutForm.AutoSize = $true
        $popOutForm.Name = "popOutForm"
        $popOutForm.Size = New-Object System.Drawing.Size(600, 400)
        $popOutForm.MinimumSize = New-Object System.Drawing.Size(300, 150)
        #$icon = "resources\icon.ico"
        #$popOutForm.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($icon)
        $popOutForm.Text = "Output"

        $enlargedTextBox = New-Object System.Windows.Forms.RichTextBox
        $enlargedTextBox.Text = $text
        $enlargedTextBox.Dock = "Fill"
        $enlargedTextBox.Multiline = $true
        $enlargedTextBox.ReadOnly = $true

        $popOutForm.Controls.Add($enlargedTextBox)

        $popOutForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
        $popOutForm.Add_KeyDown({
                if ($_.KeyCode -eq "Escape") {
                    $popOutForm.Close()
                }
            })
        $popOutForm.ShowDialog() | Out-Null
    }




    $newPowerShell = [PowerShell]::Create().AddScript($scriptBlock).AddArgument($text)
    $job = $newPowerShell.BeginInvoke()
    While (-Not $job.IsCompleted) {}
    $newPowerShell.EndInvoke($job)
    $newPowerShell.Dispose()

}

function New-CommandArea {
    param(
        $mainPanel
    )

    # Create the bottom panel
    $l1MiddlePanel = New-Object System.Windows.Forms.Panel
    $l1MiddlePanel.Dock = 'Fill'
    $l1MiddlePanel.Name = 'CommandArea'
    $l1MiddlePanel.Size = New-Object System.Drawing.Size(0, 100)
    #$l1MiddlePanel.BackColor = 'LightGray'
    $mainPanel.Controls.Add($l1MiddlePanel, 0, 1)
    # Create the table layout panel
    $TableLayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $TableLayoutPanel.Dock = 'Fill'
    # create command box
    $commandBox = New-Object System.Windows.Forms.TextBox
    $commandBox.AutoCompleteMode = 'SuggestAppend' #SuggestAppend
    $commandBox.AutoCompleteSource = 'CustomSource'
    $commandBox.ScrollBars = "Vertical"
    $commandBox.Name = "commandBox"
    $commandBox.Size = New-Object System.Drawing.Size(450, 50)
    $commandBoxAutoCompleteSource = New-Object System.Windows.Forms.AutoCompleteStringCollection
    $commandBox.AutoCompleteCustomSource = $commandBoxAutoCompleteSource
    $popOutLabel = New-Object System.Windows.Forms.Label
    $popOutLabel.Text = "+"
    $popOutLabel.Font = New-Object System.Drawing.Font("System", 12)
    $popOutLabel.AutoSize = $true
    $popOutLabel.Cursor = [System.Windows.Forms.Cursors]::Hand
    $popOutLabel.Dock = 'Right'
    $popOutLabel.Name = 'popOutLabel'
    $commandBox.Controls.Add($popOutLabel)

    $mlCommandBox = New-Object System.Windows.Forms.TextBox
    $mlCommandBox.Multiline = $true
    $mlCommandBox.Visible = $false
    $mlCommandBox.ScrollBars = "Vertical"
    $mlCommandBox.Name = "mlCommandBox"
    $mlCommandBox.Size = New-Object System.Drawing.Size(465, 70)
    $mlPopOutLabel = New-Object System.Windows.Forms.Label
    $mlPopOutLabel.Text = "-"
    $mlPopOutLabel.Font = New-Object System.Drawing.Font("System", 13)
    $mlPopOutLabel.AutoSize = $true
    $mlPopOutLabel.Cursor = [System.Windows.Forms.Cursors]::Hand
    $mlPopOutLabel.Dock = 'Right'
    $mlPopOutLabel.Name = 'mlPopOutLabel'
    $mlCommandBox.Controls.Add($mlPopOutLabel)
    $toolTippl = New-Object System.Windows.Forms.ToolTip
    $toolTippl.SetToolTip($popOutLabel, "Enlarge")
    $toolTipmlpl = New-Object System.Windows.Forms.ToolTip
    $toolTipmlpl.SetToolTip($mlPopOutLabel, "Reduce")
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
    $radioButtonPWSH.Size = New-Object System.Drawing.Size(50, 25)
    $radioButtonCMD = New-Object System.Windows.Forms.RadioButton
    $radioButtonCMD.Checked = $false
    $radioButtonCMD.Text = "cmd"
    $radioButtonCMD.Size = New-Object System.Drawing.Size(50, 25)
    # Erstellen einer Schaltfläche
    $invokeButton = New-Object System.Windows.Forms.Button
    $invokeButton.Text = "Invoke"
    $invokeButton.Anchor = "Right, top"
    #AddingControles
    $flowLayoutPanel.Controls.Add($commandLabel)
    $flowLayoutPanel.Controls.Add($radioButtonPWSH)
    $flowLayoutPanel.Controls.Add($radioButtonCMD)
    $TableLayoutPanel.Controls.Add($flowLayoutPanel, 0, 0)
    $TableLayoutPanel.SetColumnSpan($flowLayoutPanel, 2);
    $TableLayoutPanel.Controls.Add($commandBox, 0, 1)
    $TableLayoutPanel.Controls.Add($mlcommandBox, 0, 1)
    $TableLayoutPanel.SetColumnSpan($commandBox, 2)
    $TableLayoutPanel.SetRowSpan($commandBox, 1)
    $TableLayoutPanel.Controls.Add($invokeButton, 2, 1)

    $l1MiddlePanel.Controls.Add($TableLayoutPanel)

    return @{
        commandBox      = $commandBox
        radioButtonCMD  = $radioButtonCMD
        radioButtonPWSH = $radioButtonPWSH
        commandLabel    = $commandLabel
        invokeButton    = $invokeButton
        popOutLabel     = $popOutLabel
        mlpopOutLabel   = $mlpopOutLabel
        mlcommandBox    = $mlcommandBox
    }
}

function New-TableAction {
    param(
        $mainPanel
    )

    #Create the bottom panel
    $l2LeftPanel = New-Object System.Windows.Forms.Panel
    $l2LeftPanel.Dock = 'top'
    $l2LeftPanel.Name = 'TableAction'
    #$l2LeftPanel.BackColor = 'Lightgray'
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
    $l3LeftPanel.Name = 'Table'
    $l3LeftPanel.MinimumSize = New-Object System.Drawing.Size(700, 230)
    #$l3LeftPanel.BackColor = 'Lightgray'

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
   
    #$table.Dock = "Fill"
    #$table.AutoSizeColumnsMode = "Fill"
    #$table.AutoSizeRowsMode = "AllCells"
    #$table.RowsDefaultCellStyle.WrapMode = "True"

    

    $table.Anchor = "Top, Left, Bottom, Right"
    #$table.AutoSize = $true
    $table.MultiSelect = $false  # Nur eine Zeile auswählbar
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
    # Hinzufügen von Spalten zur Tabelle
    $colDisplayName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colDisplayName.HeaderText = "Name"
    $colDisplayName.Name = "Name"
    $colDisplayName.MinimumWidth = 200
    $colDisplayName.FillWeight = 200
    $table.Columns.Add($colDisplayName)  | Out-Null

    # Hinzufügen von Spalten zur Tabelle
    $colUninstallString = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colUninstallString.HeaderText = "Uninstallstring"
    $colUninstallString.Name = "Uninstallstring"
    #$column2.Width = 500
    $table.Columns.Add($colUninstallString)  | Out-Null
    $table.Columns["Uninstallstring"].Visible = $false

    # Hinzufügen von Spalten zur Tabelle
    $colVersion = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colVersion.HeaderText = "Version"
    $colVersion.Name = "Version"
    $colVersion.MinimumWidth = 150
    $colVersion.FillWeight = 100
    $table.Columns.Add($colVersion)  | Out-Null
    $table.Columns["Version"].Visible = $true

    # Hinzufügen von Spalten zur Tabelle
    $colPub = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colPub.HeaderText = "Publisher"
    $colPub.MinimumWidth = 160
    $colPub.FillWeight = 150
    $colPub.Name = "Publisher"
    $table.Columns.Add($colPub)  | Out-Null
    $table.Columns["Publisher"].Visible = $true
    $table.Columns["Publisher"].DisplayIndex = 2;


    # Hinzufügen von Spalten zur Tabelle
    $colType = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colType.HeaderText = "Type"
    $colType.MinimumWidth = 50
    $colType.FillWeight = 50
    $colType.Name = "Type"
    $table.Columns.Add($colType)  | Out-Null
    $table.Columns["Type"].Visible = $false

    # Hinzufügen von Spalten zur Tabelle
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

function New-Footer {
    param(
        $mainPanel
    )


    # Create the bottom panel
    $l4MiddlePanel = New-Object System.Windows.Forms.Panel
    $l4MiddlePanel.Dock = 'Bottom'
    $l4MiddlePanel.Name = 'Footer'
    #$l4MiddlePanel.BackColor = 'green'
    $l4MiddlePanel.Size = New-Object System.Drawing.Size(0, 25)
    $mainPanel.Controls.Add($l4MiddlePanel, 0, 4)
    # Create the table layout panel
    $TableLayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $TableLayoutPanel.Dock = 'Fill'

    
    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Text = "Save"
    $saveButton.Anchor = "left, bottom"

    # Add Buttons
    $exitButton = New-Object System.Windows.Forms.Button
    $exitButton.Text = "Exit"
    $exitButton.Anchor = "right, bottom"

    $TableLayoutPanel.Controls.Add($exitButton, 1, 0)
    $TableLayoutPanel.Controls.Add($saveButton, 0, 0)
    $TableLayoutPanel.SetColumnSpan($l4MiddlePanel, 2)
    $l4MiddlePanel.Controls.Add($TableLayoutPanel)

    return   @{
        exitButton = $exitButton
        saveButton = $saveButton
    }

}

function New-ProgressBar {
    param (

    )
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Dock = "Bottom"
    $progressBar.Name = "ProgressBar"
    $progressBar.MarqueeAnimationSpeed = 20

    return   @{
        progressBar = $progressBar
    }
}




# Show the form
Export-ModuleMember -Function *
