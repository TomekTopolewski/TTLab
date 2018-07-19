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
                $OS = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName localhost -ErrorAction Stop -ErrorVariable ErrorVar
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
        Gets information about physical drives from a local or remote computer.
    .DESCRIPTION
        The Get-TTVolumeInfo cmdlet uses the Win32_Volume class to gather information about physical drives from a local or remote computer.
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

        It is prepared to work with databases from MS and other which supports OLEDB connection.
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

        It is prepared to work with databases from MS and other which supports OLEDB connection.
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
        [string]$Password,

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

#Database Functions
Export-ModuleMember -Function Get-TTDBData
Export-ModuleMember -Function Invoke-TTDBData