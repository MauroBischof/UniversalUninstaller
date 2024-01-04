<#
   .< help keyword>
   < help content>
   . . .
   #>

#region LOGGING

Set-StrictMode -Version Latest

function Get-Base64 {

    $iconBase64 = 'iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAMAAACdt4HsAAAABGdBTUEAALGPC/xhBQAAAAFzUkdCAK7OHOkAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOp
gAABdwnLpRPAAAAkxQTFRFAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA////KI4PAAAAAMJ0Uk5TAAEPOWqiwdrj9f7AaTgRSJHD7v359Oni7cSSAiZ9yPf616x2WUEpISBad67Zx3gjJYnfzoJHGgQFG4bR+92I3POqUB
UXVdhxPLr8q0MJCkWv+LkHbeXNVwsMXdNnEpD2lB6ej6boZmzqoRZLp9JlxUxAWELejH9R62Q9bzcIn7LMH2tftIBEs3BJDc+Dm5OF4Jw6Y3rxjoe+KltgvfBosWIcXNDGVLUk2618YdQiuJqgcjJNT9
U+TpcOE53kBkqwMYTIEUSQAAAAAWJLR0TDimiOQgAAAAlwSFlzAAA7DgAAOw4BzLahgwAABCxJREFUWMOdl/9fU1UYx5/tgrs76gbbvQMZtPFlDHQ6hGUDJwxbUNhgE02mzFkpCtMwm7GAxCI1IgKhpn
wJdRSYZdEXw6ysnr+suzXbF3a3e3d+2Dn3Oa/3Z+d5zrnPPQ9A2iaRUnn5W2S0nBA5Lduav227VALCm0JZUKhSo5phNXSRhmW4YfGOEqVCGK0tLXtOh/ryisoqA1VtrKmmDLU7d5n0qNu9x6zNikvq9t
ajvsHy/L4XEs3WRmXTfhs50GzOtvjKFrS3HnwxjcMSx0u7WGxrz+SI5OVXOnSHXnXyzXd2udwdhynecHYfOYqvHevJtELP8RNI93ann/Se9OlOvZ4tSOY33PI307rhOE3O9Dmz8ZwfZ8+p+wfS8IdQ5s
9+WLQerfZgPZ7fpOA9jReo7H+vfWvwIsDbbdif4kX3STQJ4XtV+ktc/06AXE6O5BGdxi+IZ96NhmkoqHsvcWb4qLtPDA+SEYYeTQjAYTz1vhgewHkFx+JhqMSrH4jjufOwn3z4bFwqcx8TywOMM6a62H
AvtvaI5sHzEV6LLSBgvy6eB7jIBkqjgzJszRbBdDw4b+DHkd67u2MoFx6gVj0Y2YgJpiF+sD8ZlQrmwTHp7uK6Avw0butjp7xCeYDPcBpAWmhTxk3DM+SmVygPfluhFGZV5Y0JNmWqQgYeGstVs5CHFU
n5N1WhSqUv40szVheZg3ysTLamKHyOX/BvUghvwRashUwKiikyM8EnMI63QeY2QM4KBkYGtH1zJhKsQLE0yM9UQ84K8xoWsKgGclYw0oRHQKCCkUaQa6ohZ4WICzTLk84nBCjMskUgYwyQs0JkG7fiAu
SssIhL3FHeyScAQ1+SXsikEMImWCYuKw/vnNYHE171zQrWO+QuUMVt93j4Jia4kHSrSlW4ZyrenpJQMvKcwk0yk7Bp0YTCpbRpoTyXgad84fhTCVq4X6V70iGU59ZAxbOuYyWaVBVfqb8WzCe1MNkR/b
7uwQpnLnznKo5EB2stqZ82QTz42fuxe2szftMjnves4oPY0FzOLIjm4Vv9d8Zn43ay8lAsv3ZBvfz/g+J7/MEpju9cR1fCZ3SWdi+L4iU/MvU/JRp6fZqfRfDwS5A9nmToDpGWYa7fphfEGwJqS8o7rO
jHR78CbCxdEsKb8PFvqcaB83h/Q6J1CvB/4wBeebLZPrBOzp0VcN13tgfVj5+km1Fc9ul+r8nGm9d1dos0/Vz33B84Ge7MhHsWT2B9+E9e70ZvdPjGnvJe+pxPV922v/7OVJYo8kxoH6tKW/YNhFft+G
hZChmbZO1agOgbCvz7kvbZ2ugvWbGRlgdZY8TdqupGBrnSt80VGjdQ80bjPGVYDN2JlL6DIw8FVuCKrn8ixTeJFN90rPgutFwXWHz/54mUmrt1O1r+s0Wypea7FF/5/y9Z4TxUNDmqzgAAACV0RVh0ZG
F0ZTpjcmVhdGUAMjAyMC0wMi0xMlQxNzoyMzozNCswMDowMFHJSD8AAAAldEVYdGRhdGU6bW9kaWZ5ADIwMjAtMDItMTJUMTc6MjM6MzQrMDA6MDAglPCDAAAARnRFWHRzb2Z0d2FyZQBJbWFnZU1hZ2
ljayA2LjcuOC05IDIwMTktMDItMDEgUTE2IGh0dHA6Ly93d3cuaW1hZ2VtYWdpY2sub3JnQXviyAAAABh0RVh0VGh1bWI6OkRvY3VtZW50OjpQYWdlcwAxp/+7LwAAABh0RVh0VGh1bWI6OkltYWdlOj
poZWlnaHQANTEywNBQUQAAABd0RVh0VGh1bWI6OkltYWdlOjpXaWR0aAA1MTIcfAPcAAAAGXRFWHRUaHVtYjo6TWltZXR5cGUAaW1hZ2UvcG5nP7JWTgAAABd0RVh0VGh1bWI6Ok1UaW1lADE1ODE1Mj
gyMTS8W5itAAAAE3RFWHRUaHVtYjo6U2l6ZQAxNS42S0JC58sE1gAAAFB0RVh0VGh1bWI6OlVSSQBmaWxlOi8vLi91cGxvYWRzLzU2L0k4ZHFXamEvMjE1My9yb3VuZF9yZW1vdGVfZGVza3RvcF9pY2
9uXzEzMjc4MS5wbmd3Pbe6AAAAAElFTkSuQmCC'

    $iconBytes = [Convert]::FromBase64String($iconBase64)

    $stream = [System.IO.MemoryStream]::new($iconBytes, 0, $iconBytes.Length)
    return $stream


}

Function Get-ConfigValue{
    param (
        $configFile,
        $configPart
    )

    $config = ConvertFrom-Json (Get-Content -Raw -Path $configFile)
    return ($config.$configPart)
}

Function Set-DisplayBoxText {
    param (
        $displayBox,
        $text,
        $isError
    )
    $displayBox.Clear()
    $displayBox.AppendText($text)
    if ($isError) {
        Write-Error $text
        $displayBox.ForeColor = "Red"
    }
    else {
        $displayBox.ForeColor = "Black"
    }

}
function Set-ConnectedTo {
    param (
        $menuItem,
        $text
    )

    $menuItem.text = $text
}

function Enable-PsRemoting {
    param (
        $ComputerName,
        [PSCredential]$Credential
    )
    try {
        $session = New-CimSession -ComputerName $ComputerName -Credential $Credential
        Invoke-CimMethod -CimSession $session -Namespace "root/cimv2" -ClassName "Win32_Process" -MethodName "Create" -Arguments @{CommandLine = "powershell.exe -Command Enable-PSRemoting -Force -SkipNetworkProfileCheck" }
        $session | Remove-CimSession
        Start-Sleep -Seconds 5
    }
    catch {    }


}

Function Get-InstalledApp {
    param(
        [parameter(Mandatory = $false)][string]$LogText,
        [Parameter(Position = 0)]$AppName
    )
    if (!$AppName) {
        $AppName = "*"
    }

    $AppVersion = $null

    $UninstallObejct = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*
    $32BitUninstallObejct = $UninstallObejct | Select-Object  DisplayName, DisplayVersion, Publisher, @{Name = "ProductID"; Expression = { $_.PSChildName } }, UninstallString, QuietUninstallString, Type, @{Name = "Context"; Expression = { "x32" } } `
    | Where-Object { if ($AppVersion) { $_.DisplayName -like $AppName -AND $_.DisplayVersion -like $AppVersion } else { $_.DisplayName -like $AppName } }

    $UninstallObejct = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*
    $64BitUninstallObejct = $UninstallObejct | Select-Object DisplayName, DisplayVersion, Publisher, @{Name = "ProductID"; Expression = { $_.PSChildName } }, UninstallString, QuietUninstallString, Type, @{Name = "Context"; Expression = { "x64" } } `
    | Where-Object { if ($AppVersion) { $_.DisplayName -like $AppName -AND $_.DisplayVersion -like $AppVersion } else { $_.DisplayName -like $AppName } }

    $UninstallObejct = Get-ItemProperty HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*
    $UninstallObejctUser = $UninstallObejct  | Select-Object DisplayName, DisplayVersion, Publisher, @{Name = "ProductID"; Expression = { $_.PSChildName } }, UninstallString, QuietUninstallString, Type, @{Name = "Context"; Expression = { "User" } } `
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
    }

}

function Invoke-KillProcess {
    param (
        $displayName
    )

    Invoke-Command -Session $PSSession -ScriptBlock {
        $processes = Get-Process
        $processesToStop = $processes | Where-Object { $using:displayName -like "*" + $_.ProcessName + "*" }
        $processesToStop | Stop-Process -Force
    } -ErrorAction SilentlyContinue
}

function Get-SystemInfo {

    $systemInfo = Receive-Job -name "systemInfo" -Keep
    return $systemInfo
}
function Start-SystemInfo {    

    Invoke-SystemInfo
}

function Invoke-SystemInfo {
   
    $SystemInfo = Invoke-Command -Session $PSSession -ScriptBlock {

        $systemInfo = @{}

        $Memory = (Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property capacity -Sum).sum / 1gb
        $Memory = $Memory.ToString() + " GB"
        $systemInfo.Add("Memory", $Memory)

        $CPU = (Get-CimInstance -Class Win32_Processor | Select-Object -Property Name).Name 
        $systemInfo.Add("CPU", $cpu)

        $OSInfo = Get-CimInstance -Class Win32_OperatingSystem | Select-Object -Property Caption, CSName, BuildNumber, OSArchitecture
        $systemInfo.Add("OSInfo",  $OSInfo)

        $BiosInfo =  Get-CimInstance Win32_BIOS | Select-Object -Property Manufacturer, SerialNumber
        $systemInfo.Add("BiosInfo",  $BiosInfo)
    
        $Disk = (Get-CimInstance Win32_LogicalDisk | Where-Object {$_.DeviceID -like "C:"} | Measure-Object -Property Size -Sum).Sum / 1gb
        $Disk = [math]::Round($disk).ToString() + " GB"
        $systemInfo.Add("Disk",  $Disk)

        $IPInfo = Get-CimInstance win32_networkadapterconfiguration | Where-Object {$null -ne $_.IPAddress} | Select-Object MACAddress, IPAddress -First 1
        $systemInfo.Add("IPInfo",  $IPInfo)

        return $systemInfo

    } -AsJob -JobName "SystemInfo" 

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
            Remove-Job -Job $job
            return 1
        }

        $progressBarStatus = Set-ProgressBar -progressBar $progressBar -job $job
        if ($progressBarStatus -ne 124) {
            return ($job.ChildJobs[0].Output[0].ExitCode)
        }
        else {
            Remove-Job -Job $job
            return $progressBarStatus
        }
    }
    catch {
        #$logMessage = "FAIL - Error at line " + $_.InvocationInfo.ScriptLineNumber + ": " + $_.Exception.Message
    }
}

Function New-SaveFileDialog {
    param($table)

    # Open save file dialog
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV files (*.csv)|*.csv"
    $saveFileDialog.Title = "Save File"

    if ($saveFileDialog.ShowDialog() -eq 'OK') {
        $csvPath = $saveFileDialog.FileName
        # Export the data to a CSV file
        $headerRow = $table.Columns | ForEach-Object { $_.Name }
        $headerRow -join ";" | Out-File -Encoding UTF8 -FilePath $csvPath

        $rows = $table.Rows 
        $rows | ForEach-Object {
            $rowData = $_.Cells | ForEach-Object { $_.Value }
            $rowData -join ";" | Out-File -Append -Encoding UTF8 -FilePath $csvPath
        }
        return $csvPath
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
            $job = Invoke-Command  -Session $PSSession -ScriptBlock `
            { (powershell.exe -WindowStyle hidden "$using:command" 2>&1 ) } -AsJob

        }
        else {
            $job = Invoke-Command  -Session $PSSession -ScriptBlock `
            { (cmd.exe /c "$using:command" 2>&1 ) } -AsJob
        }

        $progressBarStatus = Set-ProgressBar -progressBar $progressBar -job $job

        if ($job.State -ne "Completed") {
            $errorMsg = $job.ChildJobs[0].Error
            Write-Host "Error executing command on $remoteComputer. Error: $errorMsg" -ForegroundColor Red
        }
        if ($progressBarStatus -ne 124) {
            $output = $job.ChildJobs[0].Output | Out-String

            $decodedString = [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::Default.GetBytes($output))

            Remove-Job -Job $job
            return $decodedString
        }
        else {
            Remove-Job -Job $job
            return $progressBarStatus
        }
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
        if ($stopwatch.Elapsed.TotalSeconds -gt (Get-ConfigValue -configFile $configFile -configPart "Timeout")) {
            $stopwatch.Stop()
            $progressBar.Style = "Continuous"
            Get-Job | Remove-Job -Force -ErrorAction SilentlyContinue
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
        $InstalledApps = Invoke-Command -Session $PSSession -ScriptBlock ${function:Get-InstalledApp}
        Update-TableContent -tableContent  $InstalledApps -table $table

    }
}

function Set-LastEntry {
    param (
        $text,
        $lastEntries,
        $autoCompleteSource
    )

    #$text = $text.Trim()

    # Hinzufügen der Eingabe zu den letzten Eingaben
    if (-not [string]::IsNullOrWhiteSpace($text) -and (-not $lastEntries.Contains($text))) {
        $lastEntries.Add($text)

        # Begrenzung der Anzahl der letzten Eingaben auf fünf
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

function Close-Form {
    param (
        $form
    )

    $form.Close()

}

function Get-CSVData {

    $openFileDialog = New-Object Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $result = $openFileDialog.ShowDialog()
    if ($result -eq [Windows.Forms.DialogResult]::OK) {
        $selectedFile = $openFileDialog.FileName
        if (Test-Path $selectedFile -PathType Leaf) {     
            $csvData = Import-Csv -Path $selectedFile -Delimiter $config.CSVDelimiter -Header 'Hostname', 'GUID'
            if ($csvData) { return $csvData }
        } 
    }      
}

Export-ModuleMember -Function *