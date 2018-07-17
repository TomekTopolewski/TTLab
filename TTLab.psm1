$TTErrorLogPreference = 'C:\Error.txt'
$TTConnectionString = "server=localhost\SQLEXPRESS;database=inventory;trusted_connection=$True"

Function Get-TTSystemInfo {
    <#
    .SYNOPSIS
        Retrieves information about hardware and software from local or remote machine.
    .DESCRIPTION
        Get-SystenInfo uses WMI classes like Win32_OperatingSYstem or Win32_ComputerSystem to gather information from local or remote machine.
    .PARAMETER ComputerName
        Up to 10 computer names are allowed.
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
        It uses Win32_Volume class under the hood to gather information from local or remote machine.
    .PARAMETER ComputerName
        Up to 10 computer names are allowed.
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
        Up to 10 computer names are allowed.
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
        Up to 5 computer names are allowed.
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
Function Get-TTDBData {
    <#
    .SYNOPSIS
    Function used to query information from a database. It is designed to work with MS-SQL databases and others via OleDB objects.
    .DESCRIPTION
    It uses .NET Framework object System.Data.SQLClient/OleDB.
    .PARAMETER ConnectionString
    Connection string should contain information about which database to connect and how to do it.
    .PARAMETER Query
    The actual SQL language query that will run.
    .PARAMETER IsSQLServer
    It is a switch parameter to choose between MS and other databases.
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
        $Connection = New-Object -TypeName System.Data.SqlClient.SqlConnection
    } else {
        $Connection = New-Object -TypeName System.Data.OleDb.OleDbConnection
    }

    $Connection.ConnectionString = $ConnectionString
    $Command = $Connection.CreateCommand()
    $Command.CommandText = $Query

    if ($IsSQLServer) {
        $Adapter = New-Object -TypeName System.Data.SqlClient.SqlDataAdapter $Command
    } else {
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
    Function used to work with data in database
    .PARAMETER ConnectionString
    Connection string should contain information about which database to connect and how to do it.
    .PARAMETER Query
    The actual SQL language query that will run.
    .PARAMETER IsSQLServer
    It is a switch parameter to choose between MS and other databases.
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
        $Connection = New-Object -TypeName System.Data.SqlClient.SqlConnection
    } else {
        $Connection = New-Object -TypeName System.Data.OleDb.OleDbConnection
    }

    $Connection.ConnectionString = $ConnectionString
    $Command = $Connection.CreateCommand()
    $Command.CommandText = $Query

    if ($PSCmdlet.ShouldProcess($Query)) {
        $Connection.Open()
        $Command.ExecuteNonQuery() | Out-Null
        $Connection.Close()
    }
}
Function Get-TTDBNames {
    <#
    .SYNOPSIS
    A sample function used to gather information from specific database.
    #>
    Get-TTDBData -ConnectionString $TTConnectionString -Query "SELECT computername FROM computers" -IsSQLServer
}
Function Set-TTDBInventory {
    <#
    .SYNOPSIS
    A sample function used to update computer list in database.
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $True, ValueFromPipeline = $True)]
        [Object]$InputObject
    )

    foreach ($Object in $InputObject) {
        $Query = "UPDATE
                    computers
                SET
                    OSversion = '$($Object.OSversion)',
                    SPversion = '$($Object.SPversion)',
                    Manufacturer = '$($Object.Manufacturer)',
                    Model = '$($Object.Model)'
                WHERE
                    computername = '$($Object.ComputerName)'"
        Invoke-TTDBData -ConnectionString $TTConnectionString -Query $Query -IsSQLServer
    }
}
Function Get-TTRemoteSMBShare {
    <#
    .SYNOPSIS
    Function returns a list of SMB shares
    .PARAMETER ComputerName
    Up to 5 computer names are allowed.
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
        HelpMessage="Computer name or IP address")]
        [Alias('Hostname')]
        [ValidateCount(1,5)]
        [ValidateNotNullOrEmpty()]
        [string[]]$ComputerName,

        [string]$ErrorLog = $TTErrorLogPreference
    )
    BEGIN {}
    PROCESS {
        foreach ($Computer in $ComputerName) {
            Try {
                Write-Verbose "Querying $Computer"
                $Shares = Invoke-Command -ComputerName $Computer -ScriptBlock {Get-SmbShare} -ErrorAction Stop -ErrorVariable ErrorVar
            } Catch {
                Write-Warning "$Computer FAILED"
                Write-Warning $ErrorVar.message
                $Computer | Out-File -FilePath $ErrorLog -Append
                $ErrorVar.message | Out-File -FilePath $ErrorLog -Append
                Write-Warning "Logged to $ErrorLog"
            }
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

#Database Functions
Export-ModuleMember -Function Get-TTDBData
Export-ModuleMember -Function Invoke-TTDBData