$DefaultErrorLogPath = 'C:\'

Function Get-TTSystemInfo {
    <#
    .SYNOPSIS

    Gets information about hardware and software from a local or remote machine.

    .DESCRIPTION

    The Get-TTSystemInfo cmdlet uses WMI classes to gather information about hardware and software from local or remote computers.
    
    .PARAMETER ComputerName

    Run cmdlet on the specified computers.

    .PARAMETER ErrorLogPath

    Specifies the path where the error log will be written. By default, it is C:\Error.txt.

    .PARAMETER LogError

    Indicates that this cmdlet will log errors. The path to the error log is specified by the -ErrorLogPath parameter.

    .PARAMETER WMIQuery

    Indicates that Get-WMIObject will be used instead of Get-CIMInstance.

    .EXAMPLE

    Get-Content U:\Temp\Computers.txt | Get-TTSystemInfo -Verbose

    .EXAMPLE

    Get-TTSystemInfo -ComputerName localhost -ErrorLog C:\ErrorLog.txt
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True,
                    ValueFromPipeline=$True,
                    ValueFromPipelineByPropertyName = $True)]
        [ValidateNotNullOrEmpty()]
        [string[]] $ComputerName,

        [string] $ErrorLogPath = $DefaultErrorLogPath,

        [switch] $LogError,

        [switch] $WMIQuery
    )
    BEGIN {
        if ($LogError) {
            $Time = (Get-Date)
            $Filename += 'Log'
            $Filename += '_'
            $Filename += "$($Time.Day)"
            $Filename += '_'
            $Filename += "$($Time.Month)"
            $Filename += '_'
            $Filename += "$($Time.Year)"
            $Filename += ".txt"

            $ErrorLogPath = (Join-Path -Path $ErrorLogPath -ChildPath $Filename)
        }
    }
    PROCESS {
        foreach ($Computer in $ComputerName) {
            try {
                $Status = $True
                if ($WMIQuery) {
                    $OS = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $Computer -ErrorAction Stop -ErrorVariable ErrorVar
                    $LastBootUpTime = $OS.ConvertToDateTime($OS.LastBootUpTime)
                } else {
                    $OS = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $Computer -ErrorAction Stop -ErrorVariable ErrorVar
                    $LastBootUpTime = $OS.LastBootUpTime
                }
            } catch {
                $Status = $False
                Write-Warning $ErrorVar.message
                if ($LogError) {
                    $Computer | Out-File -FilePath $ErrorLogPath -Append
                    $ErrorVar.message | Out-File -FilePath $ErrorLogPath -Append
                }
            }

            if ($Status) {
                if ($WMIQuery) {
                    $Comp = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $Computer
                    $Bios = Get-WmiObject -Class Win32_BIOS -ComputerName $Computer
                } else {
                    $Comp = Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $Computer
                    $Bios = Get-CimInstance -ClassName Win32_BIOS -ComputerName $Computer
                }
                switch ($Comp.AdminPasswordStatus) {
                    1 {$AdminPassText = 'Disabled'}
                    2 {$AdminPassText = 'Enabled'}
                    3 {$AdminPassText = 'NA'}
                    4 {$AdminPassText = 'Unknown'}
                }
                $Hash = @{
                    'ComputerName' = $Computer;
                    'OSVersion' = $OS.version;
                    'SPVersion' = $OS.servicepackmajorversion;
                    'BIOSSerial' = $Bios.serialnumber;
                    'Manufacturer' = $Comp.manufacturer;
                    'Model' = $Comp.model;
                    'AdminPassword' = $AdminPassText;
                    'Workgroup' = $Comp.workgroup
                    'LastBootTime' = $LastBootUpTime
                    'OSType' = $OS.caption
                }
                $Object = New-Object -TypeName psobject -Property $Hash
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

    Gets information about drives from local or remote computers.

    .DESCRIPTION

    The Get-TTVolumeInfo cmdlet uses the Win32_Volume class to gather information about drives from local or remote computers.

    .PARAMETER ComputerName

    Run cmdlet on the specified computers.

    .PARAMETER ErrorLogPath

    Specifies the path where the error log will be written. By default, it is C:\Error.txt.

    .PARAMETER LogError

    Indicates that this cmdlet will log errors. The path to the error log is specified by the -ErrorLog parameter.

    .PARAMETER WMIQuery

    Indicates that Get-WMIObject will be used instead of Get-CIMInstance.

    .EXAMPLE

    Get-Content U:\Temp\Computers.txt | Get-TTVolumeInfo -Verbose

    .EXAMPLE

    Get-TTVolumeInfo -ComputerName localhost -ErrorLog C:\ErrorLog.txt
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$True,
                    ValueFromPipeline=$True,
                    ValueFromPipelineByPropertyName = $True)]
        [ValidateNotNullOrEmpty()]
        [string[]] $ComputerName,

        [string] $ErrorLogPath = $DefaultErrorLogPath,

        [switch] $LogError,

        [switch] $WMIQuery
    )
    BEGIN {
        if ($LogError) {
            $Time = (Get-Date)
            $Filename += 'Log'
            $Filename += '_'
            $Filename += "$($Time.Day)"
            $Filename += '_'
            $Filename += "$($Time.Month)"
            $Filename += '_'
            $Filename += "$($Time.Year)"
            $Filename += ".txt"

            $ErrorLogPath = (Join-Path -Path $ErrorLogPath -ChildPath $Filename)
        }
    }
    PROCESS {
        foreach ($Computer in $ComputerName) {
            try {
                $Status = $True
                if ($WMIQuery) {
                    $Volumes = Get-WmiObject -Class Win32_Volume -ComputerName $Computer -Filter "DriveType=3" -ErrorAction Stop -ErrorVariable ErrorVar
                } else {
                    $Volumes = Get-CimInstance -ClassName Win32_Volume -ComputerName $Computer -Filter "DriveType=3" -ErrorAction Stop -ErrorVariable ErrorVar
                }
            } catch {
                Write-Warning $ErrorVar.message
                $Status = $False
                if ($LogError) {
                    $Computer | Out-File -FilePath $ErrorLogPath -Append
                    $ErrorVar.message | Out-File -FilePath $ErrorLogPath -Append
                }
            }
            if ($Status) {
                foreach ($Volume in $Volumes) {
                    $Size="{0:N2}" -f ($Volume.capacity/1GB)
                    $Freespace="{0:N2}" -f ($Volume.Freespace/1GB)

                    $Hash = @{
                        'FreeSpace(GB)' = $Freespace;
                        'ComputerName' = $Computer;
                        'Drive' = $Volume.Name;
                        'Size(GB)' = $Size;
                    }

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

    Gets information about services from local or remote computers.

    .DESCRIPTION

    The Get-TTServiceInfo cmdlet uses WMI classes to gather information about services from local or remote computers.

    .PARAMETER ComputerName

    Run cmdlet on the specified computers.

    .PARAMETER ErrorLogPath

    Specifies the path where the error log will be written. By default, it is C:\Error.txt.

    .PARAMETER LogError

    Indicates that this cmdlet will log errors. A path to the error log is specified by the -ErrorLog parameter.

    .PARAMETER WMIQuery

    Indicates that Get-WMIObject will be used instead of Get-CIMInstance.

    .EXAMPLE

    Get-Content U:\Temp\Computers.txt | Get-TTServiceInfo

    .EXAMPLE

    Get-TTServiceInfo -ComputerName localhost -ErrorLog C:\ErrorLog.txt
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$True,
                    ValueFromPipeline=$True,
                    ValueFromPipelineByPropertyName=$True)]
        [ValidateNotNullOrEmpty()]
        [string[]] $ComputerName,

        [string] $ErrorLogPath = $DefaultErrorLogPath,

        [switch] $LogError,

        [switch] $WMIQuery
    )
    BEGIN {
        if ($LogError) {
            $Time = (Get-Date)
            $Filename += 'Log'
            $Filename += '_'
            $Filename += "$($Time.Day)"
            $Filename += '_'
            $Filename += "$($Time.Month)"
            $Filename += '_'
            $Filename += "$($Time.Year)"
            $Filename += ".txt"

            $ErrorLogPath = (Join-Path -Path $ErrorLogPath -ChildPath $Filename)
        }
    }
    PROCESS {
        foreach ($Computer in $ComputerName) {
            try {
                $Status = $True
                if ($WMIQuery) {
                    $Services = Get-WmiObject -Class Win32_Service -ComputerName $Computer -Filter "State='Running'" -ErrorAction Stop -ErrorVariable ErrorVar
                } else {
                    $Services = Get-CimInstance -ClassName Win32_Service -ComputerName $Computer -Filter "State='Running'" -ErrorAction Stop -ErrorVariable ErrorVar
                }
            } catch {
                $Status = $False
                Write-Warning $ErrorVar.message
                if ($LogError) {
                    $Computer | Out-File -FilePath $ErrorLogPath -Append
                    $ErrorVar.message | Out-File -FilePath $ErrorLogPath -Append
                }
            }
            if ($Status) {
                foreach ($Service in $Services) {
                    $ProcessID = $Service.ProcessID
                    if ($WMIQuery) {
                        $Process = Get-WmiObject -Class Win32_Process -ComputerName $Computer -Filter "ProcessId=$ProcessID"
                    } else {
                        $Process = Get-CimInstance -ClassName Win32_Process -ComputerName $Computer -Filter "ProcessId=$ProcessID"
                    }

                    $Hash = @{
                        'ProcessName' = $Process.Name
                        'ServiceName' = $Service.Name
                        'DisplayName' = $Service.DisplayName
                        'ComputerName' = $Computer
                        'ThreadCount' = $Process.ThreadCount
                        'VM' = $Process.VirtualSize
                        'PeakPage' = $Process.PeakPageFileUsage
                    }
                    $Object = New-Object -TypeName psobject -Property $Hash
                    $Object.PSObject.TypeNames.Insert(0,'TTLab.ServiceInfo')
                    Write-Output $Object
                }
            }
        }
    }
    END {}
}

Function Get-TTSMBShare {
    <#
    .SYNOPSIS

    Gets a list of SMB shares on local or remote computers.

    .DESCRIPTION

    The Get-TTSMBShare cmdlet gets a list of SMB shares on local or remote computers.

    .PARAMETER ComputerName

    Run cmdlet on the specified computers.

    .PARAMETER ErrorLogPath

    Specifies the path where the error log will be written. By default, it is C:\Error.txt.

    .PARAMETER LogError

    Indicates that this cmdlet will log errors. A path to the error log is specified by the -ErrorLog parameter.

    .EXAMPLE

    Get-TTSMBShare -ComputerName localhost, $env:COMPUTERNAME

    .EXAMPLE

    Get-Content C:\PowerShellOutput\localhost.txt | Get-TTSMBShare
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True)]
        [ValidateNotNullOrEmpty()]
        [string[]] $ComputerName,

        [string] $ErrorLogPath = $DefaultErrorLogPath,

        [switch] $LogError
    )
    BEGIN {
        if ($LogError) {
            $Time = (Get-Date)
            $Filename += 'Log'
            $Filename += '_'
            $Filename += "$($Time.Day)"
            $Filename += '_'
            $Filename += "$($Time.Month)"
            $Filename += '_'
            $Filename += "$($Time.Year)"
            $Filename += ".txt"

            $ErrorLogPath = (Join-Path -Path $ErrorLogPath -ChildPath $Filename)
        }
    }
    PROCESS {
        foreach ($Computer in $ComputerName) {
            try {
                $Status = $True
                $Shares = Invoke-Command -ComputerName $Computer -ScriptBlock {Get-SmbShare} -ErrorAction Stop -ErrorVariable ErrorVar
            } catch {
                $Status = $False
                Write-Warning $ErrorVar.message
                if ($LogError) {
                    $Computer | Out-File -FilePath $ErrorLogPath -Append
                    $ErrorVar.message | Out-File -FilePath $ErrorLogPath -Append
                }
            }
            if ($Status) {
                foreach ($Share in $Shares) {
                    $Hash = @{
                        'ComputerName' = $Computer;
                        'Name' = $Share.Name;
                        'Description' = $Share.Description;
                        'Path' = $Share.Path
                    }
                    $Object = New-Object -TypeName psobject -Property $Hash
                    $Object.PSObject.TypeNames.Insert(0,'TTLab.RemoteSMBShare')
                    Write-Output $Object
                }
            }
         }
    }
    END {}
}

Function Get-TTProgram {
    <#
    .SYNOPSIS

    Gets a list of installed software on local or remote computers.

    .DESCRIPTION

    The Get-TTProgram cmdlet gets a list of installed software on local or remote computers.

    Before it starts to look for installed software it queries Win32_OperatingSystem class to check whether it is 32 or 64-bits architecture.
    Next, it retrieves a list from HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* for 32-bits systems or from
    HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* for 64-bits systems.

    .PARAMETER ComputerName

    Run cmdlet on the specified computers.

    .PARAMETER ErrorLogPath

    Specifies the path where the error log will be written. By default, it is C:\Error.txt.

    .PARAMETER LogError

    Indicates that this cmdlet will log errors. A path to the error log is specified by the -ErrorLog parameter.

    .EXAMPLE

    Get-TTProgram -ComputerName localhost
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$True,
                    ValueFromPipeline=$True,
                    ValueFromPipelineByPropertyName = $True)]
        [ValidateNotNullOrEmpty()]
        [string[]] $ComputerName,

        [string] $ErrorLogPath = $DefaultErrorLogPath,

        [switch] $LogError
    )
    BEGIN {
        if ($LogError) {
            $Time = (Get-Date)
            $Filename += 'Log'
            $Filename += '_'
            $Filename += "$($Time.Day)"
            $Filename += '_'
            $Filename += "$($Time.Month)"
            $Filename += '_'
            $Filename += "$($Time.Year)"
            $Filename += ".txt"

            $ErrorLogPath = (Join-Path -Path $ErrorLogPath -ChildPath $Filename)
        }
    }
    PROCESS {
        foreach ($Computer in $ComputerName) {
            try {
                $Status = $True
                $OSArchitecture = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $Computer -ErrorAction Stop -ErrorVariable ErrorVar | Select-Object -ExpandProperty OSArchitecture
            } catch {
                $Status = $False
                Write-Warning $ErrorVar.message
                if ($LogError) {
                    $Computer | Out-File -FilePath $ErrorLogPath -Append
                    $ErrorVar.message | Out-File -FilePath $ErrorLogPath -Append
                }
            }
            if ($Status) {
                if ($OSArchitecture.Substring(0,2) -eq 32) {
                    try {
                        $Programs = Invoke-Command -ComputerName $Computer -ScriptBlock {Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction Stop -ErrorVariable ErrorVar | Where-Object {$PSItem.DisplayName -gt $null}}
                    } catch {
                        Write-Warning $ErrorVar.message
                        if ($LogError) {
                            $Computer | Out-File -FilePath $ErrorLogPath -Append
                            $ErrorVar.message | Out-File -FilePath $ErrorLogPath -Append
                        }
                    }
                } else {
                    try {
                        $Programs = Invoke-Command -ComputerName $Computer -ScriptBlock {Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction Stop -ErrorVariable ErrorVar | Where-Object {$PSItem.DisplayName -gt $null}}
                    } catch {
                        Write-Warning $ErrorVar.message
                        if ($LogError) {
                            $Computer | Out-File -FilePath $ErrorLogPath -Append
                            $ErrorVar.message | Out-File -FilePath $ErrorLogPath -Append
                        }
                    }
                }
                foreach ($Program in $Programs) {
                    $Hash = @{
                        'ComputerName' = $Computer;
                        'Name' = $Program.DisplayName;
                        'Version' = $Program.DisplayVersion;
                        'Publisher' = $Program.Publisher
                    }
                    $Object = New-Object -TypeName psobject -Property $Hash
                    $Object.PSObject.TypeNames.Insert(0,'TTLab.Program')
                    Write-Output $Object
                }
            }
        }
    }
    END {}
}

Function Set-TTComputerState {
    <#
    .SYNOPSIS

    Performs the specified (LogOff, Restart, ShutDown, PowerOff) action on local or remote computers.

    .DESCRIPTION

    The Set-TTComputerState cmdlet performs the specified action (LogOff, Restart, ShutDown, PowerOff) on local or remote computers, which is set via -Action parameter.

    .PARAMETER ComputerName

    Run cmdlet on the specified computers.

    .PARAMETER ErrorLogPath

    Specifies a path where the error log will be written. By default, it is C:\Error.txt.

    .PARAMETER LogError

    Indicates that this cmdlet will log errors. A path to the error log is specified by the -ErrorLog parameter.

    .PARAMETER Action

    It accepts only one of the listed values: LogOff, Restart, ShutDown, PowerOff.

    .PARAMETER Force

    Indicates that any action specified by the -Action parameter will use -Force privileges.
    #>
    [CmdletBinding(SupportsShouldProcess = $True, ConfirmImpact = 'High')]
    Param (
        [Parameter(Mandatory=$True,
                    ValueFromPipeline=$True,
                    ValueFromPipelineByPropertyName = $True)]
        [ValidateNotNullOrEmpty()]
        [string[]] $ComputerName,

        [string] $ErrorLogPath = $DefaultErrorLogPath,

        [switch] $LogError,

        [switch] $Force,

        [Parameter(Mandatory=$True)]
        [ValidateSet("LogOff", "Restart", "ShutDown", "PowerOff")]
        [ValidateNotNullOrEmpty()]
        [string] $Action
    )
    BEGIN {
        if ($LogError) {
            $Time = (Get-Date)
            $Filename += 'Log'
            $Filename += '_'
            $Filename += "$($Time.Day)"
            $Filename += '_'
            $Filename += "$($Time.Month)"
            $Filename += '_'
            $Filename += "$($Time.Year)"
            $Filename += ".txt"

            $ErrorLogPath = (Join-Path -Path $ErrorLogPath -ChildPath $Filename)
        }

        switch ($Action) {
            "LogOff"    {$Attr = 0}
            "Restart"   {$Attr = 1}
            "ShutDown"  {$Attr = 2}
            "PowerOff"  {$Attr = 8}
        }
    }
    PROCESS {
        foreach ($Computer in $ComputerName) {
            try {
                $Status = $True
                $OS = Get-WmiObject -ComputerName $Computer -Class Win32_OperatingSYstem -ErrorAction Stop -ErrorVariable ErrorVar
            } catch {
                $Status = $False
                Write-Warning $ErrorVar.message
                if ($LogError) {
                    $Computer | Out-File -FilePath $ErrorLogPath -Append
                    $ErrorVar.message | Out-File -FilePath $ErrorLogPath -Append
                }
            }
            if ($Status) {
                if ($Force) {
                    if($PSCmdlet.ShouldProcess("Quering Win32_Shutdown method with $Action on $Computer and -Force parameter")) {
                        $OS.Win32Shutdown($Attr+4)
                    }
                } else {
                    if($PSCmdlet.ShouldProcess("Quering Win32_Shutdown method with $Action on $Computer")) {
                        $OS.Win32Shutdown($Attr)
                    }
                }
            }
        }
    }
    END {}
}

Function Get-TTNetworkInfo {
    <#
    .SYNOPSIS

    Gets basic network adapters information.

    .DESCRIPTION

    The Set-TTNetworkInfo cmdlet gets basic information such as Name, IP, MAC address from an active network adapter, from local or remote computers.

    .PARAMETER ComputerName

    Run cmdlet on the specified computers.

    .PARAMETER ErrorLogPath

    Specifies the path where the error log will be written. By default, it is C:\Error.txt.

    .PARAMETER LogError

    Indicates that this cmdlet will log errors. A path to the error log is specified by the -ErrorLog parameter.

    .PARAMETER WMIQuery

    Indicates that Get-WMIObject will be used instead of Get-CIMInstance.

    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$True,
                    ValueFromPipeline=$True,
                    ValueFromPipelineByPropertyName = $True)]
        [ValidateNotNullOrEmpty()]
        [string[]] $ComputerName,

        [string] $ErrorLogPath = $DefaultErrorLogPath,

        [switch] $LogError,

        [switch] $WMIQuery
    )
    BEGIN {
        if ($LogError) {
            $Time = (Get-Date)
            $Filename += 'Log'
            $Filename += '_'
            $Filename += "$($Time.Day)"
            $Filename += '_'
            $Filename += "$($Time.Month)"
            $Filename += '_'
            $Filename += "$($Time.Year)"
            $Filename += ".txt"

            $ErrorLogPath = (Join-Path -Path $ErrorLogPath -ChildPath $Filename)
        }
    }
    PROCESS {
        foreach ($Computer in $ComputerName) {
            try {
                $Status = $True
                if ($WMIQuery) {
                    $Adapters = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $Computer -ErrorAction Stop -ErrorVariable ErrorVar | Where-Object {$PSItem.IPEnabled -eq 'True'}
                } else {
                    $Adapters = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -ComputerName $Computer -ErrorAction Stop -ErrorVariable ErrorVar | Where-Object {$PSItem.IPEnabled -eq 'True'}
                }
            } catch {
                $Status = $False
                Write-Warning $ErrorVar.message
                if ($LogError) {
                    $Computer | Out-File -FilePath $ErrorLogPath -Append
                    $ErrorVar.message | Out-File -FilePath $ErrorLogPath -Append
                }
            }
            if ($Status) {
                foreach ($Adapter in $Adapters) {
                    $Hash = @{
                        'ComputerName' = $Computer;
                        'Name' = $Adapter.Description;
                        'DHCP Enabled' = $Adapter.DHCPEnabled;
                        'MAC' = $Adapter.MACAddress;
                        'IP' = $Adapter.IPAddress[0]
                    }
                    $Object = New-Object -TypeName psobject -Property $Hash
                    $Object.PSObject.TypeNames.Insert(0,'TTLab.NetworkAdapter')
                    Write-Output $Object
                }
            }
        }
    }
    END {}
}

Function Get-TTInfo {
    <#
    .SYNOPSIS

    Gets a huge amount of information about local or remote computers.

    .DESCRIPTION

    The Get-TTInfo cmdlet gets a huge amount of information about local or remote computers.
        'ComputerName' -    Hostname
        'OSVersion' -       Operating system version
        'SPVersion' -       Service pack version
        'BIOSSerial' -      BIOS serial number
        'Manufacturer' -    Device manufacturer
        'Model' -           Device model
        'AdminPassword' -   Admin password status
        'Workgroup' -       Workgroip
        'Volumes' -         List of disks
        'Services' -        List of running services
        'Shares' -          List of active shares
        'Programs' -        List of installed programs
        'Adapters' -        List of running network adapters

        It uses CIM cmdlet to query:
            Win32_OperatingSystem,
            Win32_ComputerSystem,
            Win32_BIOS,
            Win32_Volume,
            Win32_Service,
            Win32_Process,
            Win32_NetworkAdapterConfiguration classes. In addtion it uses two Invoke-Command cmdlets to run Get-ItemProperty cmdlet.

        IN THIS FORM IT IS HIGHLY INEFFECTIVE DUE TO THE TIME NEEDED TO COMPLETE THE FUNCTION :)

    .PARAMETER ComputerName

    Run cmdlet on the specified computers.

    .PARAMETER ErrorLogPath

    Specifies the path where the error log will be written. By default, it is C:\Error.txt.

    .PARAMETER LogError

    Indicates that this cmdlet will log errors. A path to the error log is specified by the -ErrorLog parameter.

    .EXAMPLE

    PS C:\WINDOWS\system32> Get-TTInfo -ComputerName localhost | Export-Clixml -Path C:\PowerShellOutput\massive.xml

    .EXAMPLE

    PS C:\WINDOWS\system32> Get-TTInfo -ComputerName localhost | Select-Object -ExpandProperty Adapters

    DHCP Enabled IP           Name                                                           MAC
    True 192.168.0.10 Marvell Yukon 88E8059 Family PCI-E Gigabit Ethernet Controller 00:24:45:45:45:45

    .EXAMPLE

    PS C:\WINDOWS\system32> Get-TTInfo -ComputerName localhost | Select-Object -ExpandProperty Programs

    Publisher                     Name                                                           Version
    Igor Pavlov                   7-Zip 18.01                                                    18.01
    Cisco Systems, Inc.           Cisco Packet Tracer 7.1.1 32Bit                                7.1.1.0131
    Dropbox, Inc.                 Dropbox                                                        53.4.67

    .EXAMPLE

    PS C:\WINDOWS\system32> Get-TTInfo -ComputerName localhost | Select-Object -ExpandProperty Shares

    Path                              Description     Name
    C:\WINDOWS                        Remote Admin    ADMIN$
    C:\                               Default share   C$
    D:\                               Default share   D$
    C:\WINDOWS\system32\spool\drivers Printer Drivers print$

    .EXAMPLE

    PS C:\WINDOWS\system32> Get-TTInfo -ComputerName localhost | Select-Object -ExpandProperty Services

    ProcessName : svchost.exe
    ServiceName : Appinfo
    PeakPage    : 91540
    VM          : 289193984
    ThreadCount : 49
    DisplayName : Application Information

    .EXAMPLE

    PS C:\WINDOWS\system32> Get-TTInfo -ComputerName localhost | Select-Object -ExpandProperty Volumes

    FreeSpace(GB) Drive                                             Size(GB)
    0.21          \\?\Volume{c019537a-0000-0000-0000-100000000000}\ 0.54
    165.68        C:\                                               199.90
    91.87         D:\                                               97.66
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$True,
                    ValueFromPipeline=$True,
                    ValueFromPipelineByPropertyName = $True)]
        [ValidateNotNullOrEmpty()]
        [string[]] $ComputerName,

        [string] $ErrorLogPath = $DefaultErrorLogPath,

        [switch] $LogError
    )
    BEGIN {
        if ($LogError) {
            $Time = (Get-Date)
            $Filename += 'Log'
            $Filename += '_'
            $Filename += "$($Time.Day)"
            $Filename += '_'
            $Filename += "$($Time.Month)"
            $Filename += '_'
            $Filename += "$($Time.Year)"
            $Filename += ".txt"

            $ErrorLogPath = (Join-Path -Path $ErrorLogPath -ChildPath $Filename)
        }
    }
    PROCESS {
        foreach ($Computer in $ComputerName) {
            try {
                $Status = $True
                $OS = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $Computer -ErrorAction Stop -ErrorVariable ErrorVar
                $Comp = Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $Computer
                $Bios = Get-CimInstance -ClassName Win32_BIOS -ComputerName $Computer
                $Volumes = Get-CimInstance -ClassName Win32_Volume -ComputerName $Computer -Filter "DriveType=3"
                $Services = Get-CimInstance -ClassName Win32_Service -ComputerName $Computer -Filter "State='Running'"
                $Shares = Invoke-Command -ComputerName $Computer -ScriptBlock {Get-SmbShare}
                $Adapters = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -ComputerName $Computer | Where-Object {$PSItem.IPEnabled -eq 'True'}
                $OSArchitecture = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $Computer -ErrorAction Stop -ErrorVariable ErrorVar | Select-Object -ExpandProperty OSArchitecture
            } catch {
                $Status = $False
                Write-Warning $ErrorVar.message
                if ($LogError) {
                    $Computer | Out-File -FilePath $ErrorLogPath -Append
                    $ErrorVar.message | Out-File -FilePath $ErrorLogPath -Append
                }
            }
            $VolumesArray = @()
            $ServicesArray = @()
            $SharesArray = @()
            $ProgramsArray = @()
            $AdaptersArray = @()
            if ($Status) {

                switch ($Comp.AdminPasswordStatus) {
                    1 {$AdminPassText = 'Disabled'}
                    2 {$AdminPassText = 'Enabled'}
                    3 {$AdminPassText = 'NA'}
                    4 {$AdminPassText = 'Unknown'}
                }
                
                $SystemHash = @{
                    'OSVersion' = $OS.version;
                    'SPVersion' = $OS.servicepackmajorversion;
                    'BIOSSerial' = $Bios.serialnumber;
                    'Manufacturer' = $Comp.manufacturer;
                    'Model' = $Comp.model;
                    'AdminPassword' = $AdminPassText;
                    'Workgroup' = $Comp.workgroup
                }
                $SystemObject = New-Object -TypeName psobject -Property $SystemHash

                foreach ($Volume in $Volumes) {

                    $Size="{0:N2}" -f ($Volume.capacity/1GB)
                    $Freespace="{0:N2}" -f ($Volume.Freespace/1GB)

                    $VolumeHash = @{
                        'FreeSpace(GB)' = $Freespace;
                        'Drive' = $Volume.Name;
                        'Size(GB)' = $Size;
                    }
                    $VolumeObject = New-Object -TypeName psobject -Property $VolumeHash
                    $VolumesArray += $VolumeObject
                }

                foreach ($Service in $Services) {
                    $ProcessID = $Service.ProcessID
                    $Process = Get-CimInstance -ClassName Win32_Process -ComputerName $Computer -Filter "ProcessId=$ProcessID"

                    $ServiceHash = @{
                        'ProcessName' = $Process.Name
                        'ServiceName' = $Service.Name
                        'DisplayName' = $Service.DisplayName
                        'ThreadCount' = $Process.ThreadCount
                        'VM' = $Process.VirtualSize
                        'PeakPage' = $Process.PeakPageFileUsage
                    }
                    $ServiceObject = New-Object -TypeName psobject -Property $ServiceHash
                    $ServicesArray += $ServiceObject
                }

                foreach ($Share in $Shares) {
                    $ShareHash = @{
                        'Name' = $Share.Name;
                        'Description' = $Share.Description;
                        'Path' = $Share.Path
                    }
                    $ShareObject = New-Object -TypeName psobject -Property $ShareHash
                    $SharesArray += $ShareObject
                }

                if ($OSArchitecture.Substring(0,2) -eq 32) {
                    $Programs = Invoke-Command -ComputerName $Computer -ScriptBlock {Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction Stop -ErrorVariable ErrorVar | Where-Object {$PSItem.DisplayName -gt $null}}
                } else {
                    $Programs = Invoke-Command -ComputerName $Computer -ScriptBlock {Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction Stop -ErrorVariable ErrorVar | Where-Object {$PSItem.DisplayName -gt $null}}
                }

                foreach ($Program in $Programs) {
                    $ProgramHash = @{
                        'Name' = $Program.DisplayName;
                        'Version' = $Program.DisplayVersion;
                        'Publisher' = $Program.Publisher
                    }
                    $ProgramObject = New-Object -TypeName psobject -Property $ProgramHash
                    $ProgramsArray += $ProgramObject
                }

                foreach ($Adapter in $Adapters) {
                    $AdapterHash = @{
                        'Name' = $Adapter.Description;
                        'DHCP Enabled' = $Adapter.DHCPEnabled;
                        'MAC' = $Adapter.MACAddress;
                        'IP' = $Adapter.IPAddress[0]
                    }
                    $AdapterObject = New-Object -TypeName psobject -Property $AdapterHash
                    $AdaptersArray += $AdapterObject
                }
            }

            $Hash = @{
                'ComputerName' = $Computer;
                'OSVersion' = $SystemObject.OSVersion;
                'SPVersion' = $SystemObject.SPVersion;
                'BIOSSerial' = $SystemObject.BIOSSerial;
                'Manufacturer' = $SystemObject.Manufacturer;
                'Model' = $SystemObject.Model;
                'AdminPassword' = $SystemObject.AdminPassword;
                'Workgroup' = $SystemObject.Workgroup;
                'Volumes' = $VolumesArray;
                'Services' = $ServicesArray;
                'Shares' = $SharesArray;
                'Programs'= $ProgramsArray;
                'Adapters' = $AdaptersArray;
            }
            $Object = New-Object -TypeName psobject -Property $Hash
            Write-Output $Object
        }
    }
    END {}
}

Function Get-TTAdminPasswordAge {
    <#
    .SYNOPSIS

    Gets information about active accounts on a local or remote machine.

    .DESCRIPTION

    The Get-TTAdminPasswordAge cmdlet gets information about active accounts on a local or remote machine. 
    First, it gets the names of members of the Administrator group, then it uses this names to call for an account object. 
    Finally, using object's properties it calculates password age.

    .PARAMETER ComputerName

    Run cmdlet on the specified computers.

    .PARAMETER ErrorLogPath

    Specifies the path where the error log will be written. By default, it is C:\Error.txt.

    .PARAMETER LogError

    Indicates that this cmdlet will log errors. A path to the error log is specified by the -ErrorLog parameter.

    .EXAMPLE

    PS C:\WINDOWS\system32> Get-TTAdminPasswordAge -ComputerName $env:COMPUTERNAME
    #>

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $True,
                    ValueFromPipeline = $True,
                    ValueFromPipelineByPropertyName = $True)]
        [ValidateNotNullOrEmpty()]
        [string[]] $ComputerName,

        [string] $ErrorLogPath = $DefaultErrorLogPath,

        [switch] $LogError
    )
    BEGIN{
        if ($LogError) {
            $Time = (Get-Date)
            $Filename += 'Log'
            $Filename += '_'
            $Filename += "$($Time.Day)"
            $Filename += '_'
            $Filename += "$($Time.Month)"
            $Filename += '_'
            $Filename += "$($Time.Year)"
            $Filename += ".txt"

            $ErrorLogPath = (Join-Path -Path $ErrorLogPath -ChildPath $Filename)
        }
    }

    PROCESS{
        foreach ($Computer in $ComputerName) {
            try {
                $Status = $True
                $AdminAccountsNames = Invoke-Command -ComputerName $Computer -ScriptBlock {Get-LocalGroupMember -SID S-1-5-32-544 |
                                        Where-Object {$PSItem.ObjectClass -eq 'User'} |
                                        Select-Object -ExpandProperty Name} -ErrorAction Stop -ErrorVariable ErrorVar
            } catch {
                $Status = $False
                Write-Warning $ErrorVar.message
                if ($LogError) {
                    $Computer | Out-File -FilePath $ErrorLogPath -Append
                    $ErrorVar.message | Out-File -FilePath $ErrorLogPath -Append
                }
            }

            if ($Status) {
                foreach ($AdminAccountName in $AdminAccountsNames) {
                    $Position = $AdminAccountName.IndexOf("\")
                    $Name = $AdminAccountName.Substring($Position+1)

                    $AdminAccount = Invoke-Command -ComputerName $Computer -ScriptBlock {Get-LocalUser | Where-Object {$PSItem.Enabled -eq $True -and $PSItem.Name -eq $Using:Name}}
                    if ($AdminAccount) {
                        $AdminPassLastSet = $AdminAccount | Select-Object -ExpandProperty PasswordLastSet
                        $Today = Get-Date
                        $PasswordAge = $Today - $AdminPassLastSet
                        $Hash = @{
                            'ComputerName' = $Computer;
                            'AccountName' = $AdminAccount.Name;
                            'PasswordAge' = $PasswordAge | Select-Object -ExpandProperty Days
                        }

                        $Object = New-Object -TypeName psobject -Property $Hash
                        Write-Output $Object
                    }
                }
            }
        }
    }
    END{}
}

Function Get-TTEventLog {
    <#
    .SYNOPSIS

    Gets specified event logs and writes it into a file.

    .DESCRIPTION

    The Get-TTEventlog cmdlet gets specified event log, writes it into a file and then remove it from the machine.

    .PARAMETER ComputerName

    Run cmdlet on the specified computers.

    .PARAMETER Path

    Specifies the path where the event log will be written. By default, it is C:\EventLog.txt.

    .PARAMETER LogName

    Gets the log provided with this parameter.

    .PARAMETER ErrorLogPath

    Specifies a path where the error log will be written. By default, it is C:\Error.txt.

    .PARAMETER LogError

    Indicates that this cmdlet will log errors. A path to the error log is specified by the -ErrorLog parameter.

    .PARAMETER ClearLog

    Specifies that log after writing to a file will be deleted from the machine.

    .EXAMPLE

    Get-TTEventLog -ComputerName $env:COMPUTERNAME -LogName Application -Path 'C:\Users\User\Desktop\log.txt'
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $True,
                    ValueFromPipeline = $True,
                    ValueFromPipelineByPropertyName = $True)]
        [ValidateNotNullOrEmpty()]
        [string[]] $ComputerName,

        [string] $ErrorLogPath = $DefaultErrorLogPath,

        [switch] $LogError,

        [Parameter(Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [string]$LogName,

        [switch]$ClearLog
    )
    BEGIN {
        if ($LogError) {
            $Time = (Get-Date)
            $Filename += 'Log'
            $Filename += '_'
            $Filename += "$($Time.Day)"
            $Filename += '_'
            $Filename += "$($Time.Month)"
            $Filename += '_'
            $Filename += "$($Time.Year)"
            $Filename += ".txt"

            $ErrorLogPath = (Join-Path -Path $ErrorLogPath -ChildPath $Filename)
        }
    }
    PROCESS {
        foreach ($Computer in $ComputerName) {
            try {
                $Status = $True
                $EventLogs = Get-EventLog -ComputerName $Computer -LogName $LogName -ErrorAction Stop -ErrorVariable ErrorVar | Format-List -Property '*'
            } catch {
                $Status = $False
                Write-Warning $ErrorVar.message
                if ($LogError) {
                    $Computer | Out-File -FilePath $ErrorLogPath -Append
                    $ErrorVar.message | Out-File -FilePath $ErrorLogPath -Append
                }
            }
            if ($Status) {
                try {
                    $Status = $True
                    $EventLogs | Out-File -FilePath $Path -Append -ErrorAction Stop -ErrorVariable ErrorVar
                } catch {
                    $Status = $False
                    Write-Warning $ErrorVar.message
                    if ($LogError) {
                        $Computer | Out-File -FilePath $ErrorLogPath -Append
                        $ErrorVar.message | Out-File -FilePath $ErrorLogPath -Append
                    }
                }
                if ($Status -and $ClearLog) {
                    try{
                        $Status = $True
                        Remove-EventLog -ComputerName $Computer -LogName $LogName -ErrorAction Stop -ErrorVariable ErrorVar
                    } catch {
                        $Status = $False
                        Write-Warning $ErrorVar.message
                        if ($LogError) {
                            $Computer | Out-File -FilePath $ErrorLogPath -Append
                            $ErrorVar.message | Out-File -FilePath $ErrorLogPath -Append
                        }
                    }
                }
            }
        }
    }
    END {}
}

Function Get-TTUptime {
    <#
    .SYNOPSIS

    Gets information about the last boot up time and uptime from a local or remote machine.

    .DESCRIPTION

    The Get-TTUptime cmdlet gets information about last boot up time and uptime from a local or remote machine.

    .PARAMETER ComputerName

    Run cmdlet on the specified computers.

    .PARAMETER ErrorLogPath

    Specifies the path where the error log will be written. By default, it is C:\Error.txt.

    .PARAMETER LogError

    Indicates that this cmdlet will log errors. A path to the error log is specified by the -ErrorLog parameter.

    .PARAMETER WMIQuery

    The switch parameter that indicates that Get-WMIObject will be used instead of Get-CIMInstance.

    .EXAMPLE

    PS C:\Windows\system32> Get-TTUptime -ComputerName $env:COMPUTERNAME, localhost

    ComputerName    Uptime     LastBootUpTime     
    ------------    ------     --------------     
    DESKTOP-VV4Q3OE 2d:19h:16m 22/09/2018 14:39:27
    localhost       2d:19h:16m 22/09/2018 14:39:27

    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $True,
                    ValueFromPipeline = $True,
                    ValueFromPipelineByPropertyName = $True)]
        [ValidateNotNullOrEmpty()]
        [string[]] $ComputerName,

        [string] $ErrorLogPath = $DefaultErrorLogPath,

        [switch] $LogError,

        [switch] $WMIQuery
    )
    BEGIN{
        if ($LogError) {
            $Time = (Get-Date)
            $Filename += 'Log'
            $Filename += '_'
            $Filename += "$($Time.Day)"
            $Filename += '_'
            $Filename += "$($Time.Month)"
            $Filename += '_'
            $Filename += "$($Time.Year)"
            $Filename += ".txt"

            $ErrorLogPath = (Join-Path -Path $ErrorLogPath -ChildPath $Filename)
        }
    }
    PROCESS{
        foreach ($Computer in $ComputerName) {
            try {
                $Status = $True
                if ($WMIQuery) {
                    $LastBootUpTime = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $Computer -ErrorAction Stop -ErrorVariable ErrorVar
                    $LastBootUpTime = $LastBootUpTime.ConvertToDateTime($LastBootUpTime.LastBootUpTime)
                } else {
                    $LastBootUpTime = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $Computer -ErrorAction Stop -ErrorVariable ErrorVar | 
                    Select-Object -ExpandProperty LastBootUpTime
                }
                $Today = Get-Date
            } catch {
                $Status = $False
                Write-Warning $ErrorVar.message
                if ($LogError) {
                    $Computer | Out-File -FilePath $ErrorLogPath -Append
                    $ErrorVar.message | Out-File -FilePath $ErrorLogPath -Append
                }
            }
            if ($Status) {
                $Uptime = $Today - $LastBootUpTime

                $Days = $Uptime.Days
                $Hours = $Uptime.Hours
                $Minutes = $Uptime.Minutes
                $UptimeString = "$Days`d:$Hours`h:$Minutes`m"

                $Hash = @{
                    'ComputerName' = $Computer;
                    'LastBootUpTime' = $LastBootUpTime;
                    'Uptime' = $UptimeString;
                }

                $Object = New-Object -TypeName psobject -Property $Hash
                $Object.PSObject.TypeNames.Insert(0,'TTLab.Uptime')
                Write-Output $Object
            }
        }
    }
    END{}
}

Function Get-TTRAM {
    <#
    .SYNOPSIS

    Gets information about installed RAM.

    .DESCRIPTION

    The Get-TTRAM cmdlet gets information about installed RAM or a local or remote machine.

    .PARAMETER ComputerName

    Run cmdlet on the specified computers.

    .PARAMETER ErrorLogPath

    Specifies the path where the error log will be written. By default, it is C:\Error.txt.

    .PARAMETER LogError

    Indicates that this cmdlet will log errors. A path to the error log is specified by the -ErrorLog parameter.

    .PARAMETER WMIQuery

    Indicates that Get-WMIObject will be used instead of Get-CIMInstance.
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $True,
                    ValueFromPipeline = $True,
                    ValueFromPipelineByPropertyName = $True)]
        [ValidateNotNullOrEmpty()]
        [string[]] $ComputerName,

        [string] $ErrorLogPath = $DefaultErrorLogPath,

        [switch] $LogError,

        [switch] $WMIQuery
    )
    BEGIN {
        if ($LogError) {
            $Time = (Get-Date)
            $Filename += 'Log'
            $Filename += '_'
            $Filename += "$($Time.Day)"
            $Filename += '_'
            $Filename += "$($Time.Month)"
            $Filename += '_'
            $Filename += "$($Time.Year)"
            $Filename += ".txt"

            $ErrorLogPath = (Join-Path -Path $ErrorLogPath -ChildPath $Filename)
        }
    }
    PROCESS {
        foreach ($Computer in $ComputerName) {
            try {
                $Status = $True
                if ($WMIQuery) {
                    $RAMs = Get-WMIObject -Class Win32_PhysicalMemory -ComputerName $Computer -ErrorAction Stop -ErrorVariable ErrorVar
                } else {
                    $RAMs = Get-CimInstance -ClassName Win32_PhysicalMemory -ComputerName $Computer -ErrorAction Stop -ErrorVariable ErrorVar
                }
            } catch {
                $Status = $false
                Write-Warning $ErrorVar.message
                if ($LogError) {
                    $Computer | Out-File -FilePath $ErrorLogPath -Append
                    $ErrorVar.message | Out-File -FilePath $ErrorLogPath -Append
                }
            }
            
            if ($Status) {
                foreach ($RAM in $RAMS) {
                    $Hash = @{
                            'Computer Name' = $RAM.PSComputerName;
                            'Bank Label' = $RAM.BankLabel;
                            'Capacity (GB)' = $RAM.Capacity / 1gb -as [int];
                            'Part Number' = $RAM.PartNumber;
                            'Serial Number' = $RAM.SerialNumber
                    }

                    $Object = New-Object -TypeName psobject -Property $Hash
                    $Object.PSObject.TypeNames.Insert(0,'TTLab.RAM')
                    Write-Output $Object
                }
            }
        }
    }
    END {}
}

Function Get-TTCPU {
    <#
    .SYNOPSIS

    Gets information about installed CPUs.

    .DESCRIPTION

    The Get-TTCPU cmdlet gets information about installed CPUs or a local or remote machine.

    .PARAMETER ComputerName

    Run cmdlet on the specified computers.

    .PARAMETER ErrorLogPath

    Specifies the path where the error log will be written. By default, it is C:\Error.txt.

    .PARAMETER LogError

    Indicates that this cmdlet will log errors. A path to the error log is specified by the -ErrorLog parameter.

    .PARAMETER WMIQuery

    Indicates that Get-WMIObject will be used instead of Get-CIMInstance.
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $True,
                    ValueFromPipeline = $True,
                    ValueFromPipelineByPropertyName = $True)]
        [ValidateNotNullOrEmpty()]
        [string[]] $ComputerName,

        [string] $ErrorLogPath = $DefaultErrorLogPath,

        [switch] $LogError,

        [switch] $WMIQuery
    )
    BEGIN {
        if ($LogError) {
            $Time = (Get-Date)
            $Filename += 'Log'
            $Filename += '_'
            $Filename += "$($Time.Day)"
            $Filename += '_'
            $Filename += "$($Time.Month)"
            $Filename += '_'
            $Filename += "$($Time.Year)"
            $Filename += ".txt"

            $ErrorLogPath = (Join-Path -Path $ErrorLogPath -ChildPath $Filename)
        }
    }
    
    PROCESS {
        foreach ($Computer in $ComputerName) {
            try {
                $Status = $True
                if ($WMIQuery) {
                    $CPUs = Get-WMIObject -Class Win32_Processor -ComputerName $Computer -ErrorAction Stop -ErrorVariable ErrorVar
                } else {
                    $CPUs = Get-CimInstance -ClassName Win32_Processor -ComputerName $Computer -ErrorAction Stop -ErrorVariable ErrorVar
                }
            } catch {
                $Status = $false
                Write-Warning $ErrorVar.message
                if ($LogError) {
                    $Computer | Out-File -FilePath $ErrorLogPath -Append
                    $ErrorVar.message | Out-File -FilePath $ErrorLogPath -Append
                }
            }

            if ($Status) {
                foreach ($CPU in $CPUs) {
                    $Hash = @{
                        'Computer Name' = $CPU.PSComputerName;
                        'CPU Name' = $CPU.Name;
                        'Current Voltage' = $CPU.CurrentVoltage;
                        'Current Clock Speed' = $CPu.CurrentClockSpeed;
                        'Cores' = $CPU.NumberOfCores;
                        'Logical CPUs' = $CPU.NumberOfLogicalProcessors;
                    }
                    $Object = New-Object -TypeName psobject -Property $Hash
                    $Object.PSObject.TypeNames.Insert(0,'TTLab.CPU')
                    Write-Output $Object
                }
            }
        }
    }

    END {}
}

Function Get-TTHDD {
    <#
    .SYNOPSIS

    Gets information about installed HDDs.

    .DESCRIPTION

    The Get-TTHDD cmdlet gets information about installed HDDs or a local or remote machine.

    .PARAMETER ComputerName

    Run cmdlet on the specified computers.

    .PARAMETER ErrorLogPath

    Specifies the path where the error log will be written. By default, it is C:\Error.txt.

    .PARAMETER LogError

    Indicates that this cmdlet will log errors. A path to the error log is specified by the -ErrorLog parameter.

    .PARAMETER WMIQuery

    Indicates that Get-WMIObject will be used instead of Get-CIMInstance.
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $True,
                    ValueFromPipeline = $True,
                    ValueFromPipelineByPropertyName = $True)]
        [ValidateNotNullOrEmpty()]
        [string[]] $ComputerName,

        [string] $ErrorLogPath = $DefaultErrorLogPath,

        [switch] $LogError,

        [switch] $WMIQuery
    )
    BEGIN {
        if ($LogError) {
            $Time = (Get-Date)
            $Filename += 'Log'
            $Filename += '_'
            $Filename += "$($Time.Day)"
            $Filename += '_'
            $Filename += "$($Time.Month)"
            $Filename += '_'
            $Filename += "$($Time.Year)"
            $Filename += ".txt"

            $ErrorLogPath = (Join-Path -Path $ErrorLogPath -ChildPath $Filename)
        }
    }

    PROCESS {
        foreach ($Computer in $ComputerName) {
            try {
                $Status = $True
                if ($WMIQuery) {
                    $HDDs = Get-WMIObject -Class Win32_DiskDrive -ComputerName $Computer -ErrorAction Stop -ErrorVariable ErrorVar
                } else {
                    $HDDs = Get-CimInstance -ClassName Win32_DiskDrive -ComputerName $Computer -ErrorAction Stop -ErrorVariable ErrorVar
                }
            } catch {
                $Status = $False
                Write-Warning $ErrorVar.message
                if ($LogError) {
                    $Computer | Out-File -FilePath $ErrorLogPath -Append
                    $ErrorVar.message | Out-File -FilePath $ErrorLogPath -Append
                }
            }

            if ($Status) {
                foreach ($HDD in $HDDs) {
                    $Hash = @{
                        'Computer Name' = $HDD.PSComputerName;
                        'Status' = $HDD.Status;
                        'Interface Type' = $HDD.InterfaceType;
                        'Size (GB)' = $HDD.Size / 1gb -as [int];
                        'Caption' = $HDD.Caption;
                        'Serial Number' = $HDD.SerialNumber;
                    }

                    $Object = New-Object -TypeName psobject -Property $Hash
                    $Object.PSObject.TypeNames.Insert(0,'TTLab.HDD')
                    Write-Output $Object
                }
            }
        }
    }
    
    END {}
}

Function Get-TTOu {
    <#
    .SYNOPSIS

    Gets information about the location of an Active Directory object, for example, a user or a computer.

    .DESCRIPTION

    The Get-TTOu cmdlet gets information about the location of an Active Directory object, for example, a user or a computer.

    .PARAMETER AccountName

    A name of an Active Directory object.

    .PARAMETER ErrorLogPath

    Specifies the path where the error log will be written. By default, it is C:\Error.txt.

    .PARAMETER LogError

    Indicates that this cmdlet will log errors. A path to the error log is specified by the -ErrorLog parameter.

    .PARAMETER ObjectType

    Specifies the type of an Active Directory object, whether it is a computer or a user.
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $True,
                    ValueFromPipeline = $True,
                    ValueFromPipelineByPropertyName = $True)]
        [ValidateNotNullOrEmpty()]
        [string[]] $AccountName,

        [string] $ErrorLogPath = $DefaultErrorLogPath,

        [switch] $LogError,

        [Parameter(Mandatory=$True)]
        [ValidateSet("User", "Computer")]
        [ValidateNotNullOrEmpty()]
        [string] $ObjectType
    )

    BEGIN {
        if ($LogError) {
            $Time = (Get-Date)
            $Filename += 'Log'
            $Filename += '_'
            $Filename += "$($Time.Day)"
            $Filename += '_'
            $Filename += "$($Time.Month)"
            $Filename += '_'
            $Filename += "$($Time.Year)"
            $Filename += ".txt"

            $ErrorLogPath = (Join-Path -Path $ErrorLogPath -ChildPath $Filename)
        }
    }

    PROCESS {
        foreach ($Name in $AccountName) {
            try {
                $Status = $True
                switch ($ObjectType) {
                    "User" {$Account = Get-ADUser -Identity $Name -ErrorAction Stop -ErrorVariable ErrorVar}
                    "Computer" {$Account = Get-ADComputer -Identity $Name -ErrorAction Stop -ErrorVariable ErrorVar}
                }
            } catch {
                $Status = $False
                Write-Warning $ErrorVar.message
                if ($LogError) {
                    $Computer | Out-File -FilePath $ErrorLogPath -Append
                    $ErrorVar.message | Out-File -FilePath $ErrorLogPath -Append
                }
            }
            if ($Status) {
                $DistinguishedName = $Account | Select-Object -ExpandProperty DistinguishedName
                $Array = $DistinguishedName.Split(',')
                for ($Size = $Array.Length; $Size -ge 0; $Size--) {
                    if($Array[$Size] -like "OU=*") {
                        Write-Output $Array[$Size]
                    }
                }
                Write-Output ""
            }
        }
    }
}
#Variables
Export-ModuleMember -Variable ErrorLogDefaultPath

#General Functions
Export-ModuleMember -Function Get-TTSystemInfo
Export-ModuleMember -Function Get-TTVolumeInfo
Export-ModuleMember -Function Get-TTServiceInfo
Export-ModuleMember -Function Get-TTSMBShare
Export-ModuleMember -Function Get-TTProgram
Export-ModuleMember -Function Set-TTComputerState
Export-ModuleMember -Function Get-TTNetworkInfo
Export-ModuleMember -Function Get-TTInfo
Export-ModuleMember -Function Get-TTAdminPasswordAge
Export-ModuleMember -Function Get-TTEventLog
Export-ModuleMember -Function Get-TTUptime
Export-ModuleMember -Function Get-TTRAM
Export-ModuleMember -Function Get-TTCPU
Export-ModuleMember -Function Get-TTHDD
Export-ModuleMember -Function Get-TTOu