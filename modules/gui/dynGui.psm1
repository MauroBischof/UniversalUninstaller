Set-StrictMode -Version Latest





function New-MenuStrip {
    param (     
    )
   
    $menuStrip = New-Object System.Windows.Forms.MenuStrip

    $fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem
    $fileMenu.Text = "File"

    $exitButtonItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $exitButtonItem.Text = "Exit"

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
    $hostNameBox.AutoCompleteMode = 'SuggestAppend'
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

    $Width = 500
    $Height = 400
    # Create a new form to show the enlarged text box
    $popOutForm = New-Object System.Windows.Forms.Form
    $popOutForm.AutoSize = $true
    $popOutForm.Height = ($Height - 100)
    $popOutForm.MaximumSize = New-Object System.Drawing.Size($Width, $Height)


    $enlargedTextBox = New-Object System.Windows.Forms.RichTextBox
    $enlargedTextBox.Text = $text
    $enlargedTextBox.Multiline = $true
    $enlargedTextBox.ReadOnly = $true
    $enlargedTextBox.Width = ($Width - 20)
    $enlargedTextBox.Height = ($enlargedTextBox.GetLineFromCharIndex($enlargedTextBox.Text.Length) + 1) * $enlargedTextBox.Font.Height + $enlargedTextBox.Margin.Vertical
    $enlargedTextBox.MaximumSize = New-Object System.Drawing.Size($Width, ($Height - 40))

    $popOutForm.Controls.Add($enlargedTextBox)

    # Legen Sie die maximale Größe fest, bis zu der sich die Form anpassen kann


    $popOutForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    # Show the pop-out form
    $popOutForm.ShowDialog()
    
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
    $commandBox.AutoCompleteMode = 'SuggestAppend'
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

function New-TableActions {
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
    #$progressBar.Visible = $false
  

    return   @{
        progressBar = $progressBar
    }
}




# Show the form
Export-ModuleMember -Function *
