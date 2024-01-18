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
                Set-DisplayBoxText -displayBox $ouputTextBox -text "Please set execution policy to $requiredPolicy" -isError $true
                return $false
            }
        }
        else {
            Set-DisplayBoxText -displayBox $ouputTextBox -text "Please run this program as an administrator." -isError $true
            return $false
        }
    }
    else {
        Set-DisplayBoxText -displayBox $ouputTextBox -text "Please install the powershell version $requiredVersion" -isError $true
        return $false
    }
}

Function Test-IsCorrectExecutionPolicy {
    param (
        $requiredPolicy
    )
    if ($requiredPolicy) {
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

    # Überprüfen, ob die aktuelle Version größer oder gleich der erforderlichen Version ist
    if ($currentVersion -ge $requiredVersion) {
        return $true
    }
    else {
        return $false

    }

}


Export-ModuleMember -Function Test-PreRequirement