<#
   .< help keyword>
   < help content>
   . . .
   #>

#region LOGGING

Set-StrictMode -Version Latest


Function Get-InstalledApps {
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
        Write-host $logMessage 
        #Cleanup -exitCode 1
    }
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
            { (Start-Process -FilePath msiexec.exe -NoNewWindow -ArgumentList "test" -Wait -Passthru) } -AsJob #$using:command
        }
        elseif ($type -like "Quiet") {           
            $regex = '"(.*?)"\s(.*)'
            $match = $command -split $regex 

            $executable = $match[1]
            $argument = $match[2]
         
            $job = Invoke-Command -Session $PSSession -ScriptBlock `
            { (Start-Process -FilePath $using:executable -NoNewWindow -ArgumentList $using:argument -Wait -Passthru) } -AsJob 

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
        Write-host $logMessage 
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
 
        #auf 60 erhöhren nach dem tests, timeout
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
    param( $table )

    #funktioniert noch nicht 100%
    if ($PSSession -and $PSSession.Availability -eq "Available") {
   
        $table.Rows.Clear()
        #($formItems.SearchBox.filterTextBox).Clear()
        $InstalledApps = Invoke-Command -Session $PSSession -ScriptBlock ${function:Get-InstalledApps}
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

function Set-LastEntries {
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
        $autoCompleteSource 
    )
    #wird jedes mal neu aufgerufen?
    $filteredCommands = @();
    $commands = Get-Command | Where-Object { $_.Module -like "Microsoft.Powershell*" } | Select-Object Name  
    foreach ($item in $commands) {
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

Export-ModuleMember -Function *