$TTErrorLogPreference = 'C:\Error.txt'
Function Get-TTSystemInfo {
    <#
    .SYNOPSIS
        Retrieves information about hardware and software from local or remote machine.
    .DESCRIPTION
        Get-SystenInfo uses WMI classes like Win32_OperatingSYstem or Win32_ComputerSystem to gather information from local or remote machine.
    .PARAMETER ComputerName
        You are allowed to put here up to 10 computer names.
    .PARAMETER ErrorLog
        Path to the place where the error log will be stored.
    .PARAMETER LogErrors
        It is a switch parameter to turn the log on or off.
    .EXAMPLE
        Get-Content U:\Temp\Computers.txt | Get-TTSystemInfo -Verbose
    .EXAMPLE
        Get-TTSystemInfo -ComputerName localhost -ErrorLog C:\ErrorLog.txt
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$True,
                    ValueFromPipeline=$True,
                    HelpMessage="Computer name or IP address")]
        [Alias('Hostname')]
        [ValidateCount(1,10)]
        [ValidateNotNullOrEmpty()]
        [string[]]$ComputerName,

        [string]$ErrorLog = $TTErrorLogPreference,

        [switch]$LogErrors
    )
    BEGIN {
        if ($LogErrors) {
            Write-Verbose "Error log: $ErrorLog"
            Try {
                Remove-Item -Path $ErrorLog -ErrorAction Stop -ErrorVariable ErrorVar
                Write-Warning "Previos log at $ErrorLog was removed"
            } Catch {
                Write-Warning $ErrorVar.message
            }
        } else {
            Write-Verbose "Error log is off"
        }
    }
    PROCESS {
        foreach ($Computer in $ComputerName) {
            Write-Verbose "Querying $Computer"
            Try {
                $ErrorStatus = $True
                $OS = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $Computer -ErrorAction Stop -ErrorVariable ErrorVar
            } Catch {
                $ErrorStatus = $False
                Write-Warning "$Computer FAILED"
                Write-Warning $ErrorVar.message
                If ($LogErrors) {
                    $Computer | Out-File -FilePath $ErrorLog -Append
                    $ErrorVar.message | Out-File -FilePath $ErrorLog -Append
                    Write-Warning "Logged to $ErrorLog"
                }
            }

            if ($ErrorStatus) {
                $Comp = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $Computer
                $Bios = Get-WmiObject -Class Win32_BIOS -ComputerName $Computer

                switch ($Comp.AdminPasswordStatus) {
                    1 {$AdminPassText = 'Disabled'}
                    2 {$AdminPassText = 'Enabled'}
                    3 {$AdminPassText = 'NA'}
                    4 {$AdminPassText = 'Unknown'}
                }
                $Props = @{
                    'ComputerName' = $Computer;
                    'OSVersion' = $OS.version;
                    'SPVersion' = $OS.servicepackmajorversion;
                    'BIOSSerial' = $Bios.serialnumber;
                    'Manufacturer' = $Comp.manufacturer;
                    'Model' = $Comp.model;
                    'AdminPassword' = $AdminPassText;
                    'Workgroup' = $Comp.workgroup
                }
                Write-Verbose "WMI queries completed"
                $Object = New-Object -TypeName psobject -Property $Props
                $Object.PSObject.TypeNames.Insert(0,'TTLab.SystemInfo')
                Write-Output $Object
            }
        }
    }
    END {}
}
Function Get-TTVolumeInfo {
    <#
    .SYNOPSIS
        Retrieves information about physical drives from local or remote machine.
    .DESCRIPTION
        Under the hood it uses Win32_Volume class to gather information from local or remote machine.
    .PARAMETER ComputerName
        You are allowed to put here up to 10 computer names.
    .PARAMETER ErrorLog
        Path to the place where the error log will be stored.
    .PARAMETER LogErrors
        It is a switch parameter to turn the log on or off.
    .EXAMPLE
        Get-Content U:\Temp\Computers.txt | Get-TTVolumeInfo -Verbose
    .EXAMPLE
        Get-TTVolumeInfo -ComputerName localhost -ErrorLog C:\ErrorLog.txt
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$True,
                    ValueFromPipeline=$True,
                    HelpMessage="Computer name or IP address")]
        [Alias('Hostname')]
        [ValidateCount(1,10)]
        [ValidateNotNullOrEmpty()]
        [string[]]$ComputerName,

        [string]$ErrorLog = $TTErrorLogPreference,

        [switch]$LogErrors
    )
    BEGIN {
        if ($LogErrors) {
            Write-Verbose "Error log: $ErrorLog"
            Try {
                Remove-Item -Path $ErrorLog -ErrorAction Stop -ErrorVariable ErrorVar
                Write-Warning "Previos log at $ErrorLog was removed"
            } Catch {
                Write-Warning $ErrorVar.message
            }
        } else {
            Write-Verbose "Error log is off"
        }
    }
    PROCESS {
        foreach ($Computer in $ComputerName) {
            Write-Verbose "Querying $Computer"
            Try {
                $ErrorStatus = $True
                $Volumes = Get-WmiObject -Class Win32_Volume -ComputerName $Computer -Filter "DriveType=3" -ErrorAction Stop -ErrorVariable ErrorVar
            } Catch {
                Write-Warning "$Computer FAILED"
                Write-Warning $ErrorVar.message
                $ErrorStatus = $False
                If ($LogErrors) {
                    $Computer | Out-File -FilePath $ErrorLog -Append
                    $ErrorVar.message | Out-File -FilePath $ErrorLog -Append
                    Write-Warning "Logged to $ErrorLog"
                }
            }
            if ($ErrorStatus) {
                foreach ($Volume in $Volumes) {

                    $Size="{0:N2}" -f ($Volume.capacity/1GB)
                    $Freespace="{0:N2}" -f ($Volume.Freespace/1GB)

                    $Hash = @{
                        'FreeSpace(GB)' = $Freespace;
                        'ComputerName' = $Volume.SystemName;
                        'Drive' = $Volume.Name;
                        'Size(GB)' = $Size;
                    }
                    Write-Verbose "WMI queries completed"
                    $Object = New-Object -TypeName psobject -Property $Hash
                    $Object.PSObject.TypeNames.Insert(0,'TTLab.VolumeInfo')
                    Write-Output $Object
                }
            }
        }
    }
    END {}
}
Function Get-TTServiceInfo {
    <#
    .SYNOPSIS
        Retrieves information about services from local or remote machine.
    .DESCRIPTION
        Get-ServiceInfo uses WMI classes like Win32_Service and Win32_Process to gather information from local or remote machine.
    .PARAMETER ComputerName
        You are allowed to put here up to 10 computer names.
    .PARAMETER ErrorLog
        Path to the place where the error log will be stored.
    .PARAMETER LogErrors
        It is a switch parameter to turn the log on or off.
    .EXAMPLE
        Get-Content U:\Temp\Computers.txt | Get-TTServiceInfo -Verbose
    .EXAMPLE
        Get-TTServiceInfo -ComputerName localhost -ErrorLog C:\ErrorLog.txt
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$True,
                    ValueFromPipeline=$True,
                    HelpMessage="Computer name or IP address")]
        [Alias('Hostname')]
        [ValidateCount(1,10)]
        [ValidateNotNullOrEmpty()]
        [string[]]$ComputerName,

        [string]$ErrorLog = $TTErrorLogPreference,

        [switch]$LogErrors
    )
    BEGIN {
        if ($LogErrors) {
            Write-Verbose "Error log: $ErrorLog"
            Try {
                Remove-Item -Path $ErrorLog -ErrorAction Stop -ErrorVariable ErrorVar
                Write-Warning "Previos log at $ErrorLog was removed"
            } Catch {
                Write-Warning $ErrorVar.message
            }
        } else {
            Write-Verbose "Error log is off"
        }
    }
    PROCESS {
        foreach ($Computer in $ComputerName) {
            Write-Verbose "Querying $Computer"
            Try {
                $ErrorStatus = $True
                $Services = Get-WmiObject -Class Win32_Service -ComputerName $Computer -Filter "State='Running'" -ErrorAction Stop -ErrorVariable ErrorVar
            } Catch {
                $ErrorStatus = $False
                Write-Warning "$Computer FAILED"
                Write-Warning $ErrorVar.message
                If ($LogErrors) {
                    $Computer | Out-File -FilePath $ErrorLog -Append
                    $ErrorVar.message | Out-File -FilePath $ErrorLog -Append
                    Write-Warning "Logged to $ErrorLog"
                }
            }
            if ($ErrorStatus) {
                foreach ($Service in $Services) {
                    $ProcessID = $Service.ProcessID
                    $Process = Get-WmiObject -Class Win32_Process -ComputerName $Computer -Filter "ProcessId=$ProcessID"

                    $Hash = @{
                        'ProcessName' = $Process.Name
                        'ServiceName' = $Service.Name
                        'DisplayName' = $Service.DisplayName
                        'ComputerName' = $Computer
                        'ThreadCount' = $Process.ThreadCount
                        'VM' = $Process.VirtualSize
                        'PeakPage' = $Process.PeakPageFileUsage
                    }
                    Write-Verbose "WMI queries completed"
                    $Object = New-Object -TypeName psobject -Property $Hash
                    $Object.PSObject.TypeNames.Insert(0,'TTLab.ServiceInfo')
                    Write-Output $Object
                }
            }
        }
    }
    END {}
}
Function Get-TTSystemInfo2 {
    <#
    .SYNOPSIS
        Retrieves information about hardware and software from local or remote machine.
    .DESCRIPTION
        Get-SystenInfo uses WMI classes like Win32_OperatingSYstem or Win32_ComputerSystem to gather information from local or remote machine.
    .PARAMETER ComputerName
        You are allowed to put here up to 10 computer names.
    .PARAMETER ErrorLog
        Path to the place where the error log will be stored.
    .PARAMETER LogErrors
        It is a switch parameter to turn the log on or off.
    .EXAMPLE
        Get-Content U:\Temp\Computers.txt | Get-TTSystemInfo2 -Verbose
    .EXAMPLE
        Get-TTSystemInfo2 -ComputerName localhost -ErrorLog C:\ErrorLog.txt
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$True,
                    ValueFromPipeline=$True,
                    HelpMessage="Computer name or IP address")]
        [Alias('Hostname')]
        [ValidateCount(1,10)]
        [ValidateNotNullOrEmpty()]
        [string[]]$ComputerName,

        [string]$ErrorLog = $TTErrorLogPreference,

        [switch]$LogErrors
    )
    BEGIN {
            if ($LogErrors) {
                Write-Verbose "Error log: $ErrorLog"
                Try {
                    Remove-Item -Path $ErrorLog -ErrorAction Stop -ErrorVariable ErrorVar
                    Write-Warning "Previos log at $ErrorLog was removed"
                } Catch {
                    Write-Warning $ErrorVar.message
                }
            } else {
                Write-Verbose "Error log is off"
            }
    }
    PROCESS {
        foreach ($Computer in $ComputerName) {
            Write-Verbose "Quering $Computer"
            Try {
                $ErrorStatus = $True
                $Win32_OS = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $Computer -ErrorAction Stop -ErrorVariable ErrorVar
            } Catch {
                $ErrorStatus = $False
                Write-Warning "$Computer FAILED"
                Write-Warning $ErrorVar.message
                if ($LogErrors) {
                    $Computer | Out-File -FilePath $ErrorLog -Append
                    $ErrorVar.message | Out-File -FilePath $ErrorLog -Append
                    Write-Warning "Logged to $ErrorLog"
                }
            }

            if ($ErrorStatus) {
                $Win32_CS = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $Computer
                $Hash = @{
                    OSType = $Win32_OS.Caption
                    ComputerName = $Win32_OS.PSComputerName
                    LastBootTime = $Win32_OS.ConvertToDateTime($Win32_OS.LastBootUpTime)
                    Manufacturer = $Win32_CS.Manufacturer
                    Model = $Win32_CS.Model
                }
                Write-Verbose "WMI queries completed"
                $Object = New-Object -TypeName PSobject -Property $Hash
                $Object.PSObject.TypeNames.Insert(0,'TTLab.SystemInfo2')
                Write-Output $Object
            }
        }
    }
    END {}
}

Export-ModuleMember -Variable TTErrorLogPreference
Export-ModuleMember -Function Get-TTSystemInfo
Export-ModuleMember -Function Get-TTVolumeInfo
Export-ModuleMember -Function Get-TTServiceInfo
Export-ModuleMember -Function Get-TTSystemInfo2