$TTErrorLogPreference = 'C:\Error.txt'

Function Get-TTSystemInfo {
    <#
    .SYNOPSIS
        Gets information about hardware and software from a local or remote machine.
    .DESCRIPTION
        The Get-TTSystemInfo cmdlet uses WMI classes (Win32_OperatingSYstem and Win32_ComputerSystem) to gather information about hardware and software from a local or remote computer.
    .PARAMETER ComputerName
        Gets the information about hardware and software from the specified computers, up to ten machines are allowed.
    .PARAMETER ErrorLog
        Specifies a path where the error log will be stored. By default, it is C:\Error.txt.
    .PARAMETER LogErrors
        Indicates that this cmdlet will log errors. A path to the error log is specified by the -ErrorLog parameter.
    .EXAMPLE
        Get-Content U:\Temp\Computers.txt | Get-TTSystemInfo -Verbose
    .EXAMPLE
        Get-TTSystemInfo -ComputerName localhost -ErrorLog C:\ErrorLog.txt
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$True,
                    ValueFromPipeline=$True,
                    ValueFromPipelineByPropertyName = $True,
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
                $OS = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $Computer -ErrorAction Stop -ErrorVariable ErrorVar
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
                $Comp = Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $Computer
                $Bios = Get-CimInstance -ClassName Win32_BIOS -ComputerName $Computer

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
        Gets information about drives from a local or remote computer.
    .DESCRIPTION
        The Get-TTVolumeInfo cmdlet uses the Win32_Volume class to gather information about drives from a local or remote computer.
    .PARAMETER ComputerName
        Gets the information about volumes from the specified computers, up to ten machines are allowed.
    .PARAMETER ErrorLog
        Specifies a path where the error log will be stored. By default, it is C:\Error.txt.
    .PARAMETER LogErrors
        Indicates that this cmdlet will log errors. A path to the error log is specified by the -ErrorLog parameter.
    .EXAMPLE
        Get-Content U:\Temp\Computers.txt | Get-TTVolumeInfo -Verbose
    .EXAMPLE
        Get-TTVolumeInfo -ComputerName localhost -ErrorLog C:\ErrorLog.txt
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$True,
                    ValueFromPipeline=$True,
                    ValueFromPipelineByPropertyName = $True,
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
                $Volumes = Get-CimInstance -ClassName Win32_Volume -ComputerName $Computer -Filter "DriveType=3" -ErrorAction Stop -ErrorVariable ErrorVar
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
        Gets information about services from a local or remote computer.
    .DESCRIPTION
        The Get-TTServiceInfo cmdlet uses WMI classes (Win32_Service and Win32_Process) to gather information about services from a local or remote computer.
    .PARAMETER ComputerName
        Gets the information about services from the specified computers, up to ten machines are allowed.
    .PARAMETER ErrorLog
        Specifies a path where the error log will be stored. By default, it is C:\Error.txt.
    .PARAMETER LogErrors
        Indicates that this cmdlet will log errors. A path to the error log is specified by the -ErrorLog parameter.
    .EXAMPLE
        Get-Content U:\Temp\Computers.txt | Get-TTServiceInfo -Verbose
    .EXAMPLE
        Get-TTServiceInfo -ComputerName localhost -ErrorLog C:\ErrorLog.txt
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$True,
                    ValueFromPipeline=$True,
                    ValueFromPipelineByPropertyName=$True,
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
                $Services = Get-CimInstance -ClassName Win32_Service -ComputerName $Computer -Filter "State='Running'" -ErrorAction Stop -ErrorVariable ErrorVar
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
                    $Process = Get-CimInstance -ClassName Win32_Process -ComputerName $Computer -Filter "ProcessId=$ProcessID"

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
        Gets information about hardware and software from a local or remote computer.
    .DESCRIPTION
        The Get-TTSystenInfo cmdlet uses WMI classes (Win32_OperatingSystem and Win32_ComputerSystem) to gather information about hardware and software from a local or remote computer.
    .PARAMETER ComputerName
        Gets the information about hardware and software from the specified computers, up to ten machines are allowed.
    .PARAMETER ErrorLog
        Specifies a path where the error log will be stored. By default, it is C:\Error.txt.
    .PARAMETER LogErrors
        Indicates that this cmdlet will log errors. A path to the error log is specified by the -ErrorLog parameter.
    .EXAMPLE
        Get-Content U:\Temp\Computers.txt | Get-TTSystemInfo2 -Verbose
    .EXAMPLE
        Get-TTSystemInfo2 -ComputerName localhost -ErrorLog C:\ErrorLog.txt
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$True,
                    ValueFromPipeline=$True,
                    ValueFromPipelineByPropertyName=$True,
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
                $Win32_OS = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $Computer -ErrorAction Stop -ErrorVariable ErrorVar
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
                $Win32_CS = Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $Computer
                $Hash = @{
                    OSType = $Win32_OS.Caption
                    ComputerName = $Win32_OS.PSComputerName
                    LastBootTime = $Win32_OS.LastBootUpTime
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

Function Get-TTDBData {
    <#
    .SYNOPSIS
        Gets data from a database.
    .DESCRIPTION
        The Get-TTDBData cmdlet is used to queries information from a database.

        It is prepared to work with MS databases and others which supports OLEDB connection.
    .PARAMETER ConnectionString
        Specifies the connection string which should contain information how to connect to a database.
    .PARAMETER Query
        Specifies the actual SQL language query that will run.
    .PARAMETER IsSQLServer
        Indicates that we will query MS-SQL Server database.
    .EXAMPLE
        $ConnectionString = "server=localhost\SQLEXPRESS;database=inventory;trusted_connection=$True"

        $Query = "SELECT Something FROM Somewhere WHERE Something = Something"

        Get-TTDBData -ConnectionString $ConnectionString -Query $Query -IsSQLServer
    #>
    [CmdletBinding()]
    Param (
        [string]$ConnectionString,

        [string]$Query,

        [switch]$IsSQLServer
    )

    if ($IsSQLServer) {
        Write-Verbose "Attempting to create a SqlConnection"
        $Connection = New-Object -TypeName System.Data.SqlClient.SqlConnection
    } else {
        Write-Verbose "Attempting to create a OleDbConnection"
        $Connection = New-Object -TypeName System.Data.OleDb.OleDbConnection
    }

    $Connection.ConnectionString = $ConnectionString
    $Command = $Connection.CreateCommand()
    $Command.CommandText = $Query

    if ($IsSQLServer) {
        Write-Verbose "Creating SqlDataAdapter"
        $Adapter = New-Object -TypeName System.Data.SqlClient.SqlDataAdapter $Command
    } else {
        Write-Verbose "Creating OleDbDataAdapter"
        $Adapter = New-Object -TypeName System.Data.OleDb.OleDbDataAdapter $Command
    }

    $DataSet = New-Object -TypeName System.Data.DataSet
    $Adapter.Fill($DataSet)
    $DataSet.Tables[0]
    $Connection.Close()
}

Function Invoke-TTDBData {
    <#
    .SYNOPSIS
        Write data to a database.
    .DESCRIPTION
        The Invoke-TTDBData cmdlet is used to write data to a database.

        It is prepared to work with MS databases and others which supports OLEDB connection.
    .PARAMETER ConnectionString
        Specifies the connection string which should contain information how to connect to a database.
    .PARAMETER Query
        Specifies the actual SQL language query that will run.
    .PARAMETER IsSQLServer
        Indicates that we will query MS-SQL Server database.
    .EXAMPLE
        $ConnectionString = "server=localhost\SQLEXPRESS;database=inventory;trusted_connection=$True"

        $Query = "UPDATE Database SET Columns = Something, Columns = Something"

        Get-TTDBQuery -ConnectionString $ConnectionString -Query $Query
    #>
    [CmdletBinding(SupportsShouldProcess = $True, ConfirmImpact = 'Low')]
    Param (
        [string]$ConnectionString,

        [string]$Query,

        [switch]$IsSQLServer
    )

    if ($IsSQLServer) {
        Write-Verbose "Attempting to create a SqlConnection"
        $Connection = New-Object -TypeName System.Data.SqlClient.SqlConnection
    } else {
        Write-Verbose "Attempting to create a OleDbConnection"
        $Connection = New-Object -TypeName System.Data.OleDb.OleDbConnection
    }

    $Connection.ConnectionString = $ConnectionString
    $Command = $Connection.CreateCommand()
    $Command.CommandText = $Query

    if ($PSCmdlet.ShouldProcess($Query)) {
        Write-Verbose "Executing $Query"
        $Connection.Open()
        $Command.ExecuteNonQuery() | Out-Null
        $Connection.Close()
        Write-Verbose "Connection closed"
    }
}

Function Get-TTRemoteSMBShare {
    <#
    .SYNOPSIS
    Gets a list of SMB shares on a local or remote computer.
    .DESCRIPTION
    The Get-TTRemoteSMBShare cmdlet gets a list of SMB shares on a local or remote computer.
    It uses an Invoke-Command query to connect to a machine.
    .PARAMETER ComputerName
    Gets the information about SMB shares from the specified computers, up to 5 machines are allowed.
    .PARAMETER ErrorLog
    Specifies a path where the error log will be stored. By default, it is C:\Error.txt.
    .PARAMETER LogErrors
    Indicates that this cmdlet will log errors. A path to the error log is specified by the -ErrorLog parameter.
    .EXAMPLE
    Get-TTRemoteSMBShare -ComputerName localhost, localhost
    .EXAMPLE
    Get-Content C:\PowerShellOutput\localhost.txt | Get-TTRemoteSMBShare
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage="Computer name")]
        [Alias('Hostname')]
        [ValidateCount(1,5)]
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
            Try {
                $Status = $True
                Write-Verbose "Querying $Computer"
                $Shares = Invoke-Command -ComputerName $Computer -ScriptBlock {Get-SmbShare} -ErrorAction Stop -ErrorVariable ErrorVar
            } Catch {
                $Status = $False
                Write-Warning "$Computer FAILED"
                Write-Warning $ErrorVar.message
                If ($LogErrors) {
                    $Computer | Out-File -FilePath $ErrorLog -Append
                    $ErrorVar.message | Out-File -FilePath $ErrorLog -Append
                    Write-Warning "Logged to $ErrorLog"
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
                    Write-Verbose "WMI query completed"
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
    Gets a list of installed software on a local or remote computer.
    .DESCRIPTION
    The Get-TTProgram cmdlet gets a list of installed software on a local or remote computer.

    Before it starts to look for installed software it queries Win32_OperatingSystem class to check whether it is 32 or 64-bits architecture.
    Next, it retrieves a list from HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* for 32-bits systems or from
    HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* for 64-bits systems.

    As a final step, it creates an object called TTLab.Program which can be piped to another cmdlet.
    .PARAMETER ComputerName
    Gets the information about installed programs from the specified computers, up to ten machines are allowed.
    .PARAMETER ErrorLog
    Specifies a path where the error log will be stored. By default, it is C:\Error.txt.
    .PARAMETER LogErrors
    Indicates that this cmdlet will log errors. A path to the error log is specified by the -ErrorLog parameter.
    .EXAMPLE
    Get-TTProgram -ComputerName localhost
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$True,
                    ValueFromPipeline=$True,
                    ValueFromPipelineByPropertyName = $True,
                    HelpMessage="Computer name")]
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
            Try {
                $Status = $True
                Write-Verbose "Querying $Computer for OS architecture"
                $OSArchitecture = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $Computer -ErrorAction Stop -ErrorVariable ErrorVar | Select-Object -ExpandProperty OSArchitecture
            } Catch {
                $Status = $False
                Write-Warning "Querying $Computer for OS architecture FAILED"
                Write-Warning $ErrorVar.message
                If ($LogErrors) {
                    $Computer | Out-File -FilePath $ErrorLog -Append
                    $ErrorVar.message | Out-File -FilePath $ErrorLog -Append
                    Write-Warning "Logged to $ErrorLog"
                }
            }
            if ($Status) {
                if ($OSArchitecture.Substring(0,2) -eq 32) {
                    Try {
                        Write-Verbose "Querying $Computer x86"
                        $Programs = Invoke-Command -ComputerName $Computer -ScriptBlock {Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction Stop -ErrorVariable ErrorVar | Where-Object {$PSItem.DisplayName -gt $null}}
                    } Catch {
                        Write-Warning "$Computer FAILED"
                        Write-Warning $ErrorVar.message
                        If ($LogErrors) {
                            $Computer | Out-File -FilePath $ErrorLog -Append
                            $ErrorVar.message | Out-File -FilePath $ErrorLog -Append
                            Write-Warning "Logged to $ErrorLog"
                        }
                    }
                } else {
                    Try {
                        Write-Verbose "Querying $Computer x64"
                        $Programs = Invoke-Command -ComputerName $Computer -ScriptBlock {Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction Stop -ErrorVariable ErrorVar | Where-Object {$PSItem.DisplayName -gt $null}}
                    } Catch {
                        Write-Warning "$Computer FAILED"
                        Write-Warning $ErrorVar.message
                        If ($LogErrors) {
                            $Computer | Out-File -FilePath $ErrorLog -Append
                            $ErrorVar.message | Out-File -FilePath $ErrorLog -Append
                            Write-Warning "Logged to $ErrorLog"
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
                    Write-Verbose "WMI query completed. Creating object"
                    $Object = New-Object -TypeName psobject -Property $Hash
                    $Object.PSObject.TypeNames.Insert(0,'TTLab.Program')
                    Write-Output $Object
                }
            }
        }
    }
    END {}
}

Function Restart-TTComputer {
    <#
    .SYNOPSIS
    Reboots local or remote computer.
    .DESCRIPTION
    The Restart-TTComputer cmdlet reboots a local or remote computer.
    It uses both Invoke-CimMethod and Invoke-WmiMethod, second is launched when the first failed.
    .PARAMETER ComputerName
    Gets the information about installed programs from the specified computers, up to ten machines are allowed.
    .PARAMETER ErrorLog
    Specifies a path where the error log will be stored. By default, it is C:\Error.txt.
    .PARAMETER LogErrors
    Indicates that this cmdlet will log errors. A path to the error log is specified by the -ErrorLog parameter.
    #>
    [CmdletBinding(SupportsShouldProcess = $True, ConfirmImpact = 'High')]
    Param (
        [Parameter(Mandatory=$True,
                    ValueFromPipeline=$True,
                    ValueFromPipelineByPropertyName = $True,
                    HelpMessage="Computer name")]
        [Alias('Hostname')]
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
            Try {
                $Status = $True
                Write-Debug "Attempting to reboot $Computer via Invoke-CimMethod"
                Invoke-CimMethod -ClassName Win32_OperatingSystem -ComputerName $Computer -MethodName Reboot -ErrorAction Stop -ErrorVariable ErrorVar
            } Catch {
                $Status = $False
                Write-Warning "$Computer reboot FAILED (Cim)"
                Write-Warning $ErrorVar.message
                If ($LogErrors) {
                    $Computer | Out-File -FilePath $ErrorLog -Append
                    $ErrorVar.message | Out-File -FilePath $ErrorLog -Append
                    Write-Warning "Logged to $ErrorLog"
                }
            }
            If ($Status = $False) {
                Try {
                    Write-Debug "Attempting to reboot $Computer via Invoke-WmiMethod"
                    Invoke-WmiMethod -Class Win32_OperatingSystem -Name Reboot -ComputerName $Computer -ErrorAction Stop -ErrorVariable ErrorVar
                } Catch {
                    Write-Warning "$Computer reboot FAILED (Wmi)"
                    Write-Warning $ErrorVar.message
                    If ($LogErrors) {
                        $Computer | Out-File -FilePath $ErrorLog -Append
                        $ErrorVar.message | Out-File -FilePath $ErrorLog -Append
                        Write-Warning "Logged to $ErrorLog"
                    }
                }
            }
        }
    }
    END {}
}

Function Set-TTServicePassword {
        <#
    .SYNOPSIS
    Changes the service's startup password.
    .DESCRIPTION
    The Set-TTServicePassword cmdlet changes the service's startup password.
    .PARAMETER ServiceName
    Specifies the service names of services to be retrieved.
    .PARAMETER Password
    Specifies the new service's password.
    .PARAMETER ComputerName
    Gets the information about installed programs from the specified computers, up to ten machines are allowed.
    .PARAMETER ErrorLog
    Specifies a path where the error log will be stored. By default, it is C:\Error.txt.
    .PARAMETER LogErrors
    Indicates that this cmdlet will log errors. A path to the error log is specified by the -ErrorLog parameter.
    #>
    [CmdletBinding(SupportsShouldProcess = $True, ConfirmImpact = 'Medium')]
    Param (
        [Parameter(Mandatory=$True,
                    ValueFromPipeline=$True,
                    ValueFromPipelineByPropertyName = $True,
                    HelpMessage="Computer name")]
        [Alias('Hostname')]
        [ValidateNotNullOrEmpty()]
        [string[]]$ComputerName,

        [Parameter(Mandatory = $True)]
        [string]$ServiceName,

        [Parameter(Mandatory = $True)]
        [securestring]$Password,

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
            Try {
                $Status = $True
                Write-Debug "Querying $Computer"
                $Services = Get-WmiObject -ComputerName $Computer -Class Win32_Service -Filter "name='$ServiceName" -ErrorAction Stop -ErrorVariable ErrorVar
            } Catch {
                $Status = $False
                Write-Warning "Quering $Computer FAILED"
                Write-Warning $ErrorVar.message
                If ($LogErrors) {
                    $Computer | Out-File -FilePath $ErrorLog -Append
                    $ErrorVar.message | Out-File -FilePath $ErrorLog -Append
                    Write-Warning "Logged to $ErrorLog"
                }
            }
            if ($Status) {
                foreach ($Service in $Services) {
                    if ($PSCmdlet.ShouldProcess("$Service on $Computer")) {
                        $Service.change($null, $null, $null, $null, $null, $null, $null, $Password) | Out-Null
                    }
                }
            }
        }
    }
    END {}
}

Function Set-TTComputerState {
    <#
    .SYNOPSIS
    Performs the specified (LogOff, Restart, ShutDown, PowerOff) action on a local or remote machine.
    .DESCRIPTION
    The Set-TTComputerState cmdlet performs the specified action (LogOff, Restart, ShutDown, PowerOff) on a local or remote machine, which is set via -Action parameter.
    .PARAMETER ComputerName
    Performs the specified (LogOff, Restart, ShutDown, PowerOff) action on a local or remote machine, up to ten machines are allowed.
    .PARAMETER ErrorLog
    Specifies a path where the error log will be stored. By default, it is C:\Error.txt.
    .PARAMETER LogErrors
    Indicates that this cmdlet will log errors. A path to the error log is specified by the -ErrorLog parameter.
    .PARAMETER Action
    Accepts only one of the listed values: LogOff, Restart, ShutDown, PowerOff.
    .PARAMETER Force
    Indicates that any action specified by the -Action parameter will use -Force privileges.
    #>
    [CmdletBinding(SupportsShouldProcess = $True, ConfirmImpact = 'High')]
    Param (
        [Parameter(Mandatory=$True,
                    ValueFromPipeline=$True,
                    ValueFromPipelineByPropertyName = $True,
                    HelpMessage="Computer name")]
        [Alias('Hostname')]
        [ValidateNotNullOrEmpty()]
        [string[]]$ComputerName,

        [string]$ErrorLog = $TTErrorLogPreference,

        [switch]$LogErrors,

        [switch]$Force,

        [Parameter(Mandatory=$True)]
        [ValidateSet("LogOff", "Restart", "ShutDown", "PowerOff")]
        [ValidateNotNullOrEmpty()]
        [string]$Action
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

        switch ($Action) {
            "LogOff"    {$Attr = 0}
            "Restart"   {$Attr = 1}
            "ShutDown"  {$Attr = 2}
            "PowerOff"  {$Attr = 8}
        }
    }
    PROCESS {
        foreach ($Computer in $ComputerName) {
            Try {
                $Status = $True
                Write-Debug "Querying $Computer"
                $OS = Get-WmiObject -ComputerName $Computer -Class Win32_OperatingSYstem -ErrorAction Stop -ErrorVariable ErrorVar
            } Catch {
                $Status = $False
                Write-Warning "Quering $Computer FAILED"
                Write-Warning $ErrorVar.message
                If ($LogErrors) {
                    $Computer | Out-File -FilePath $ErrorLog -Append
                    $ErrorVar.message | Out-File -FilePath $ErrorLog -Append
                    Write-Warning "Logged to $ErrorLog"
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
    The Set-TTNetworkInfo cmdlet gets basic information such as Name, IP, MAC address from an active network adapter, from a local or remote machine.
    .PARAMETER ComputerName
    Gets basic network adapters information from a local or remote machine.
    .PARAMETER ErrorLog
    Specifies a path where the error log will be stored. By default, it is C:\Error.txt.
    .PARAMETER LogErrors
    Indicates that this cmdlet will log errors. A path to the error log is specified by the -ErrorLog parameter.
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$True,
                    ValueFromPipeline=$True,
                    ValueFromPipelineByPropertyName = $True,
                    HelpMessage="Computer name")]
        [Alias('Hostname')]
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
            Try {
                $Status = $True
                Write-Verbose "Querying $Computer for network active adapters"
                $Adapters = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -ComputerName $Computer -ErrorAction Stop -ErrorVariable ErrorVar | Where-Object {$PSItem.IPEnabled -eq 'True'}
            } Catch {
                $Status = $False
                Write-Warning "Querying $Computer for network active adapters FAILED"
                Write-Warning $ErrorVar.message
                If ($LogErrors) {
                    $Computer | Out-File -FilePath $ErrorLog -Append
                    $ErrorVar.message | Out-File -FilePath $ErrorLog -Append
                    Write-Warning "Logged to $ErrorLog"
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
                    Write-Verbose "WMI query completed. Creating object"
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
    Gets a huge amount of information about a local or remote machine.
    .DESCRIPTION
    The Get-TTInfo cmdlet gets a huge amount of information about a local or remote machine.
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
    Gets information from a local or remote machine.
    .PARAMETER ErrorLog
    Specifies a path where the error log will be stored. By default, it is C:\Error.txt.
    .PARAMETER LogErrors
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
                    ValueFromPipelineByPropertyName = $True,
                    HelpMessage="Computer name")]
        [Alias('Hostname')]
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
                $Status = $True
                $OS = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $Computer -ErrorAction Stop -ErrorVariable ErrorVar
                $Comp = Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $Computer
                $Bios = Get-CimInstance -ClassName Win32_BIOS -ComputerName $Computer
                $Volumes = Get-CimInstance -ClassName Win32_Volume -ComputerName $Computer -Filter "DriveType=3"
                $Services = Get-CimInstance -ClassName Win32_Service -ComputerName $Computer -Filter "State='Running'"
                $Shares = Invoke-Command -ComputerName $Computer -ScriptBlock {Get-SmbShare}
                $Adapters = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -ComputerName $Computer | Where-Object {$PSItem.IPEnabled -eq 'True'}
                $OSArchitecture = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $Computer -ErrorAction Stop -ErrorVariable ErrorVar | Select-Object -ExpandProperty OSArchitecture
            } Catch {
                $Status = $False
                Write-Warning "One of the command send to the $Computer FAILED"
                Write-Warning $ErrorVar.message
                If ($LogErrors) {
                    $Computer | Out-File -FilePath $ErrorLog -Append
                    $ErrorVar.message | Out-File -FilePath $ErrorLog -Append
                    Write-Warning "Logged to $ErrorLog"
                }
            }
            $VolumesArray = @()
            $ServicesArray = @()
            $SharesArray = @()
            $ProgramsArray = @()
            $AdaptersArray = @()
            If ($Status) {

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

            $MainHash = @{
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
            $MainObject = New-Object -TypeName psobject -Property $MainHash
            Write-Output $MainObject
        }
    }
    END {}
}

Function Export-TTHTML {
    [CmdletBinding(DefaultParameterSetName='Page',  RemotingCapability='None')]
    Param (
        [Parameter(ValueFromPipeline=$true)]
        [psobject]
        ${InputObject},

        [Parameter(Position=0)]
        [System.Object[]]
        ${Property},

        [Parameter(ParameterSetName='Page', Position=3)]
        [string[]]
        ${Body},

        [Parameter(ParameterSetName='Page', Position=1)]
        [string[]]
        ${Head},

        [Parameter(ParameterSetName='Page', Position=2)]
        [ValidateNotNullOrEmpty()]
        [string]
        ${Title},

        [ValidateNotNullOrEmpty()]
        [ValidateSet('Table','List')]
        [string]
        ${As},

        [Parameter(ParameterSetName='Page')]
        [Alias('cu','uri')]
        [ValidateNotNullOrEmpty()]
        [uri]
        ${CssUri},

        [ValidateNotNullOrEmpty()]
        [string[]]
        ${PostContent},

        [ValidateNotNullOrEmpty()]
        [string[]]
        ${PreContent},

        [Parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [string]
        $FilePath
    )


    BEGIN {
        Try {
            $outBuffer = $null
            if ($PSBoundParameters.TryGetValue('OutBuffer', [ref]$outBuffer))
            {
                $PSBoundParameters['OutBuffer'] = 1
            }
            $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand('Microsoft.PowerShell.Utility\ConvertTo-Html', [System.Management.Automation.CommandTypes]::Cmdlet)
            $PSBoundParameters.Remove('FilePath') | Out-Null
            $scriptCmd = {& $wrappedCmd @PSBoundParameters | Out-File $FilePath}
            $steppablePipeline = $scriptCmd.GetSteppablePipeline($myInvocation.CommandOrigin)
            $steppablePipeline.Begin($PSCmdlet)
        } Catch {
            Throw
        }
    }

    PROCESS {
        Try {
            $steppablePipeline.Process($_)
        } Catch {
            Throw
        }
    }

    END {
        Try {
            $steppablePipeline.End()
        } Catch {
            Throw
        }
}

}

Function Get-TTAdminPasswordAge {
    <#
    .SYNOPSIS
    Gets information about active accounts on a local or remote machine.
    .DESCRIPTION
    The Get-TTAdminPasswordAge cmdlet gets information about active accounts on a local or remote machine. First, it gets names of members of Administrator group, then it use
    this names to call for account object. Finally, using object's properties it calculates password age.
    .PARAMETER ComputerName
    Gets information from a local or remote machine.
    .PARAMETER ErrorLog
    Specifies a path where the error log will be stored. By default, it is C:\Error.txt.
    .PARAMETER LogErrors
    Indicates that this cmdlet will log errors. A path to the error log is specified by the -ErrorLog parameter.
    .EXAMPLE
    PS C:\WINDOWS\system32> Get-TTAdminPasswordAge -ComputerName $env:COMPUTERNAME
    #>

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $True,
                    ValueFromPipeline = $True,
                    ValueFromPipelineByPropertyName = $True)]
        [Alias("Hostname")]
        [ValidateNotNullOrEmpty()]
        [string[]]$ComputerName,

        [string]$ErrorLog = $TTErrorLogPreference,

        [switch]$LogErrors
    )
    BEGIN{
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

    PROCESS{
        foreach ($Computer in $ComputerName) {
            Write-Verbose "Querying $Computer"
            Try {
                $Status = $True
                $AdminAccountsNames = Invoke-Command -ComputerName $Computer -ScriptBlock {Get-LocalGroupMember -SID S-1-5-32-544 |
                                        Where-Object {$PSItem.ObjectClass -eq 'User'} |
                                        Select-Object -ExpandProperty Name} -ErrorAction Stop -ErrorVariable ErrorVar
            } Catch {
                $Status = $False
                Write-Warning "Querying $Computer FAILED"
                Write-Warning $ErrorVar.message
                If ($LogErrors) {
                    $Computer | Out-File -FilePath $ErrorLog -Append
                    $ErrorVar.message | Out-File -FilePath $ErrorLog -Append
                    Write-Warning "Logged to $ErrorLog"
                }
            }

            If ($Status) {
                foreach($AdminAccountName in $AdminAccountsNames) {
                    $Position = $AdminAccountName.IndexOf("\")
                    $Name = $AdminAccountName.Substring($Position+1)

                    $AdminAccount = Invoke-Command -ComputerName $Computer -ScriptBlock {Get-LocalUser | Where-Object {$PSItem.Enabled -eq "True" -and $PSItem.Name -eq $Using:Name}}
                    if ($AdminAccount){
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

#Variables
Export-ModuleMember -Variable TTErrorLogPreference

#General Functions
Export-ModuleMember -Function Get-TTSystemInfo
Export-ModuleMember -Function Get-TTVolumeInfo
Export-ModuleMember -Function Get-TTServiceInfo
Export-ModuleMember -Function Get-TTSystemInfo2
Export-ModuleMember -Function Get-TTRemoteSMBShare
Export-ModuleMember -Function Get-TTProgram
Export-ModuleMember -Function Restart-TTComputer
Export-ModuleMember -Function Set-TTServicePassword
Export-ModuleMember -Function Set-TTComputerState
Export-ModuleMember -Function Get-TTNetworkInfo
Export-ModuleMember -Function Get-TTInfo
Export-ModuleMember -Function Export-TTHTML
Export-ModuleMember -Function Get-TTAdminPasswordAge

#Database Functions
Export-ModuleMember -Function Get-TTDBData
Export-ModuleMember -Function Invoke-TTDBData