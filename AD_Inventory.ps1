# Creating a background process for the main script
if (-not $env:INVENTORY_CHILD_PROCESS) {
    Write-Host "Starting inventory in isolated child process..." -ForegroundColor Cyan
    
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "powershell.exe"
    $psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $psi
    $process.StartInfo.EnvironmentVariables["INVENTORY_CHILD_PROCESS"] = "true"

    [void]$process.Start()

    # Direct synchronous polling of the output buffer in real time
    while (-not $process.HasExited) {
        while (-not $process.StandardOutput.EndOfStream) {
            $line = $process.StandardOutput.ReadLine()
            if ($null -ne $line) {
                [System.Console]::WriteLine($line)
            }
        }
        Start-Sleep -Milliseconds 20
    }

    # Read the remaining output after the process exits
    while (-not $process.StandardOutput.EndOfStream) {
        $line = $process.StandardOutput.ReadLine()
        if ($null -ne $line) {
            [System.Console]::WriteLine($line)
        }
    }

    # Errors (if any)
    $err = $process.StandardError.ReadToEnd()
    if (-not [string]::IsNullOrWhiteSpace($err)) {
        Write-Host $err -ForegroundColor Red
    }

    $process.Dispose()

    Write-Host ""
    Write-Host "Child process finished. All hung DCOM/WMI runspaces purged." -ForegroundColor Green
    
    return
}

# Override Write-Host for PowerShell 5.1 with ANSI support (for colored log output)
function Write-Host {
    param(
        [Parameter(Position=0, ValueFromPipeline=$true)] $Object,
        [ConsoleColor] $ForegroundColor,
        [ConsoleColor] $BackgroundColor,
        [switch] $NoNewline
    )
    
    $esc = [char]27
    $colors = @{
        'Black'="${esc}[30m"; 'DarkBlue'="${esc}[34m"; 'DarkGreen'="${esc}[32m"; 'DarkCyan'="${esc}[36m"
        'DarkRed'="${esc}[31m"; 'DarkMagenta'="${esc}[35m"; 'DarkYellow'="${esc}[33m"; 'Gray'="${esc}[37m"
        'DarkGray'="${esc}[90m"; 'Blue'="${esc}[94m"; 'Green'="${esc}[92m"; 'Cyan'="${esc}[96m"
        'Red'="${esc}[91m"; 'Magenta'="${esc}[95m"; 'Yellow'="${esc}[93m"; 'White'="${esc}[97m"
    }

    $msg = if ($null -ne $Object) { $Object.ToString() } else { "" }

    if ($ForegroundColor -and $colors.ContainsKey($ForegroundColor.ToString())) {
        $colorCode = $colors[$ForegroundColor.ToString()]
        $msg = "${colorCode}${msg}${esc}[0m"
    }

    if ($NoNewline) {
        [System.Console]::Write($msg)
    } else {
        [System.Console]::WriteLine($msg)
    }

    [System.Console]::Out.Flush()
}

Import-Module ActiveDirectory
Import-Module ImportExcel -ErrorAction Stop

$ErrorActionPreference = "Stop"

$scriptDirectory = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$output = Join-Path $scriptDirectory "AD_Inventory.xlsx"
$RetryPartial = $false
$SuccessfulStatuses = @("OK", "FALLBACK")
$RetryStatuses = @("ERROR", "UNREACHABLE / OFFLINE", "UNREACHABLE / REMOTE ACCESS CLOSED", "SKIPPED_NON_WINDOWS/FIREWALL_BLOCKED")
$logFile = Join-Path $scriptDirectory "AD_Inventory.log"

$TestMode = $false
$Threads = 20
$JobTimeoutSeconds = 60
$JobStopGraceSeconds = 40

#
# If this many runspaces become hung
# within a single RunspacePool,
# create a new RunspacePool and continue processing.
#
$MaxAbandonedRunspaces = 15

#
# Preparation
#
New-Item -Path "C:\Temp" -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
Remove-Item $logFile -Force -ErrorAction SilentlyContinue

function Write-Log {
    param([string]$Text)

    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $Text"
    Add-Content -Path $logFile -Value $line
}

Write-Host ""
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host " AD Inventory START" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host ""

Write-Log "START INVENTORY"
Write-Log "Execution engine: RunspacePool"
Write-Log "Threads: $Threads"
Write-Log "Job timeout: $JobTimeoutSeconds seconds"
Write-Log "Stop grace: $JobStopGraceSeconds seconds"
Write-Log "Max abandoned runspaces per pool: $MaxAbandonedRunspaces"

#
# Worker script
#
$worker = @'
param(
    [Parameter(Mandatory = $true)]
    [string]$ComputerName,

    [Parameter(Mandatory = $true)]
    [System.Collections.Concurrent.ConcurrentQueue[object]]$PrecheckQueue
)

$ErrorActionPreference = "Stop"

#
# Result object
#
$result = [ordered]@{
    Computer    = $null
    Network     = [System.Collections.Generic.List[object]]::new()
    Users       = [System.Collections.Generic.List[object]]::new()
    Printers    = [System.Collections.Generic.List[object]]::new()
    Devices     = [System.Collections.Generic.List[object]]::new()
    Diagnostics = $null
    Error       = $null
    Log         = [System.Collections.Generic.List[string]]::new()
}

function Add-WorkerLog {
    param([string]$Message)

    [void]$result.Log.Add(
        "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $Message"
    )
}

#
# Additional ways to retrieve data: Remote Registry + WinRM
#
function Get-FallbackSystemInfo {
    param([string]$ComputerName)

    $fallback = [PSCustomObject]@{
        Method       = ""
        ComputerName = ""
        OS           = ""
        Build        = ""
        Manufacturer = ""
        Model        = ""
        CPU          = ""
        RAM_GB       = ""
        CurrentUser  = ""
        Error        = ""
    }

    #
    # 1. WINRM
    #
    try {
        $fallbackData = Invoke-Command -ComputerName $ComputerName -ScriptBlock {

            $system = Get-WmiObject Win32_ComputerSystem -ErrorAction Stop
            $os     = Get-WmiObject Win32_OperatingSystem -ErrorAction Stop
            $cpu    = Get-WmiObject Win32_Processor -ErrorAction Stop

            [PSCustomObject]@{
                ComputerName = $system.Name
                OS           = $os.Caption
                Build        = $os.BuildNumber
                Manufacturer = $system.Manufacturer
                Model        = $system.Model
                CPU          = (
                    $cpu |
                    Select-Object -First 1 |
                    Select-Object -ExpandProperty Name
                )
                RAM_GB       = [math]::Round(
                    $system.TotalPhysicalMemory / 1GB,
                    2
                )
                CurrentUser  = $system.UserName
            }

        } -ErrorAction Stop

        $fallback.Method       = "WinRM"
        $fallback.ComputerName = $fallbackData.ComputerName
        $fallback.OS           = $fallbackData.OS
        $fallback.Build        = $fallbackData.Build
        $fallback.Manufacturer = $fallbackData.Manufacturer
        $fallback.Model        = $fallbackData.Model
        $fallback.CPU          = $fallbackData.CPU
        $fallback.RAM_GB       = $fallbackData.RAM_GB
        $fallback.CurrentUser  = $fallbackData.CurrentUser

        return $fallback
    }
    catch {
        $winrmError = $_.Exception.Message
    }

    #
    # 2. REMOTE REGISTRY
    #
    try {

        $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey(
            [Microsoft.Win32.RegistryHive]::LocalMachine,
            $ComputerName
        )

        #
        # Computer name
        #
        try {

            $compKey = $reg.OpenSubKey(
                "SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName"
            )

            if ($compKey) {
                $fallback.ComputerName = $compKey.GetValue("ComputerName")
                $compKey.Close()
            }

        }
        catch {}

        if (-not $fallback.ComputerName) {
            $fallback.ComputerName = $ComputerName
        }

        #
        # OS
        #
        try {

            $osKey = $reg.OpenSubKey(
                "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
            )

            if ($osKey) {

                $fallback.OS = $osKey.GetValue("ProductName")
                $fallback.Build = $osKey.GetValue("CurrentBuildNumber")

                if (-not $fallback.Build) {
                    $fallback.Build = $osKey.GetValue("CurrentBuild")
                }

                $osKey.Close()
            }

        }
        catch {}

        #
        # CPU
        #
        try {

            $cpuKey = $reg.OpenSubKey(
                "HARDWARE\DESCRIPTION\System\CentralProcessor\0"
            )

            if ($cpuKey) {

                $fallback.CPU = $cpuKey.GetValue(
                    "ProcessorNameString"
                )

                $cpuKey.Close()
            }

        }
        catch {}

        #
        # Manufacturer / Model
        #
        try {

            $sysKey = $reg.OpenSubKey(
                "SYSTEM\CurrentControlSet\Control\SystemInformation"
            )

            if ($sysKey) {

                $fallback.Manufacturer = $sysKey.GetValue(
                    "SystemManufacturer"
                )

                $fallback.Model = $sysKey.GetValue(
                    "SystemProductName"
                )

                $sysKey.Close()
            }

        }
        catch {}

        $reg.Close()

        #
        # Check whether at least the OS or CPU was retrieved
        #
        if ($fallback.OS -or $fallback.CPU) {

            $fallback.Method = "RemoteRegistry"

            return $fallback
        }
        else {
            throw "Access denied to essential registry keys."
        }

    }
    catch {

        $fallback.Method = "FAILED"

        $fallback.Error = `
            "WinRM: $winrmError; RemoteRegistry: $($_.Exception.Message)"

        return $fallback
    }
}

#
# MAIN WORKER
#
$time = {
    Get-Date -Format "yyyy-MM-dd HH:mm:ss"
}

try {

    $target = $ComputerName

    [void]$result.Log.Add(
        "$(&$time) START $target"
    )

#
# ==========================================================
# PRECHECK
# ==========================================================
#
# 1. Ping check.
#
# 2. TCP 135/445/5985/5986 check.
#
# 3. If at least one TCP port is open:
#    consider the host available and continue collection.
#
# 4. If Ping fails + all TCP ports are closed:
#    consider the host unavailable.
#
# 5. If Ping succeeds + all TCP ports are closed:
#    the host responds to ICMP, but Windows management
#    endpoints are unavailable. Windows data collection is not started.
#

$isPingable = Test-Connection `
    -ComputerName $target `
    -Count 1 `
    -Quiet `
    -ErrorAction SilentlyContinue


#
# ==========================================================
# PRECHECK DIAGNOSTICS object
# ==========================================================
#

$precheckDiagnostics = [PSCustomObject]@{
    ComputerName   = $target
    DNS            = $false
    IP             = ""
    Ping           = [bool]$isPingable
    TCP135         = $false
    TCP445         = $false
    TCP5985        = $false
    TCP5986        = $false
    RemoteRegistry = "NotChecked"
}


#
# ==========================================================
# PING status
# ==========================================================
#

if ($isPingable) {

    [void]$result.Log.Add(
        "$(&$time) $target PING OK"
    )

    [void]$PrecheckQueue.Enqueue(
        [PSCustomObject]@{
            Computer = $target
            Type     = "PING_OK"
            Port     = $null
            Message  = "PRECHECK - PING succeeded. Checking TCP ports 135, 445, 5985, 5986..."
        }
    )
}
else {

    [void]$result.Log.Add(
        "$(&$time) $target PING FAILED. Checking TCP ports..."
    )

    [void]$PrecheckQueue.Enqueue(
        [PSCustomObject]@{
            Computer = $target
            Type     = "PING_FAILED"
            Port     = $null
            Message  = "PRECHECK - PING failed. Checking TCP ports 135, 445, 5985, 5986..."
        }
    )

    #
    # DNS check.
    #
    try {

        $addresses = [System.Net.Dns]::GetHostAddresses(
            $target
        )

        $ipv4 = @(
            $addresses |
                Where-Object {
                    $_.AddressFamily -eq
                    [System.Net.Sockets.AddressFamily]::InterNetwork
                }
        )

        if ($ipv4.Count -gt 0) {

            $precheckDiagnostics.DNS = $true

            $precheckDiagnostics.IP = (
                $ipv4 |
                    ForEach-Object {
                        $_.IPAddressToString
                    }
            ) -join ","
        }
    }
    catch {}
}


#
# ==========================================================
# TCP 135
# ==========================================================
#

try {

    $test = Test-NetConnection `
        -ComputerName $target `
        -Port 135 `
        -WarningAction SilentlyContinue `
        -ErrorAction SilentlyContinue

    $precheckDiagnostics.TCP135 = [bool]$test.TcpTestSucceeded
}
catch {}


#
# ==========================================================
# TCP 445
# ==========================================================
#

try {

    $test = Test-NetConnection `
        -ComputerName $target `
        -Port 445 `
        -WarningAction SilentlyContinue `
        -ErrorAction SilentlyContinue

    $precheckDiagnostics.TCP445 = [bool]$test.TcpTestSucceeded
}
catch {}


#
# ==========================================================
# TCP 5985
# ==========================================================
#

try {

    $test = Test-NetConnection `
        -ComputerName $target `
        -Port 5985 `
        -WarningAction SilentlyContinue `
        -ErrorAction SilentlyContinue

    $precheckDiagnostics.TCP5985 = [bool]$test.TcpTestSucceeded
}
catch {}


#
# ==========================================================
# TCP 5986
# ==========================================================
#

try {

    $test = Test-NetConnection `
        -ComputerName $target `
        -Port 5986 `
        -WarningAction SilentlyContinue `
        -ErrorAction SilentlyContinue

    $precheckDiagnostics.TCP5986 = [bool]$test.TcpTestSucceeded
}
catch {}


#
# ==========================================================
# TCP PRECHECK result
# ==========================================================
#

$tcpReachable =
    $precheckDiagnostics.TCP135 -or
    $precheckDiagnostics.TCP445 -or
    $precheckDiagnostics.TCP5985 -or
    $precheckDiagnostics.TCP5986


#
# Log complete TCP state.
#

[void]$result.Log.Add(
    "$(&$time) $target TCP135=$($precheckDiagnostics.TCP135) TCP445=$($precheckDiagnostics.TCP445) TCP5985=$($precheckDiagnostics.TCP5985) TCP5986=$($precheckDiagnostics.TCP5986)"
)


#
# ==========================================================
# If at least one TCP port is open
# ==========================================================
#

if ($tcpReachable) {

    #
    # Separate message for EVERY open TCP port.
    #

    if ($precheckDiagnostics.TCP135) {

        [void]$PrecheckQueue.Enqueue(
            [PSCustomObject]@{
                Computer = $target
                Type     = "TCP_OK"
                Port     = 135
                Message  = "PRECHECK - TCP port 135 is OPEN. Collection continues."
            }
        )

        [void]$result.Log.Add(
            "$(&$time) $target TCP port 135 OPEN"
        )
    }

    if ($precheckDiagnostics.TCP445) {

        [void]$PrecheckQueue.Enqueue(
            [PSCustomObject]@{
                Computer = $target
                Type     = "TCP_OK"
                Port     = 445
                Message  = "PRECHECK - TCP port 445 is OPEN. Collection continues."
            }
        )

        [void]$result.Log.Add(
            "$(&$time) $target TCP port 445 OPEN"
        )
    }

    if ($precheckDiagnostics.TCP5985) {

        [void]$PrecheckQueue.Enqueue(
            [PSCustomObject]@{
                Computer = $target
                Type     = "TCP_OK"
                Port     = 5985
                Message  = "PRECHECK - TCP port 5985 is OPEN. Collection continues."
            }
        )

        [void]$result.Log.Add(
            "$(&$time) $target TCP port 5985 OPEN"
        )
    }

    if ($precheckDiagnostics.TCP5986) {

        [void]$PrecheckQueue.Enqueue(
            [PSCustomObject]@{
                Computer = $target
                Type     = "TCP_OK"
                Port     = 5986
                Message  = "PRECHECK - TCP port 5986 is OPEN. Collection continues."
            }
        )

        [void]$result.Log.Add(
            "$(&$time) $target TCP port 5986 OPEN"
        )
    }


    #
    # Final message.
    #

    if ($isPingable) {

        [void]$result.Log.Add(
            "$(&$time) $target PING OK and at least one TCP endpoint is reachable. Continuing collection."
        )

        [void]$PrecheckQueue.Enqueue(
            [PSCustomObject]@{
                Computer = $target
                Type     = "PRECHECK_OK"
                Port     = $null
                Message  = "PRECHECK OK - PING succeeded and at least one TCP endpoint is reachable. Collection continues."
            }
        )
    }
    else {

        [void]$result.Log.Add(
            "$(&$time) $target PING FAILED, but TCP endpoint is reachable. Continuing collection."
        )

        [void]$PrecheckQueue.Enqueue(
            [PSCustomObject]@{
                Computer = $target
                Type     = "PRECHECK_OK"
                Port     = $null
                Message  = "PRECHECK OK - PING failed, but at least one TCP endpoint is reachable. Collection continues."
            }
        )
    }
}
else {

    #
    # ======================================================
    # If all TCP ports are closed
    # ======================================================
    #

    $result.Diagnostics = $precheckDiagnostics


    #
    # ------------------------------------------------------
    # PING failed + all TCP ports are closed
    # ------------------------------------------------------
    #

    if (-not $isPingable) {

        [void]$PrecheckQueue.Enqueue(
            [PSCustomObject]@{
                Computer = $target
                Type     = "PRECHECK_FAILED"
                Port     = $null
                Message  = "PRECHECK FAILED - PING failed and TCP ports 135, 445, 5985, 5986 are CLOSED. Host is unreachable/offline."
            }
        )

        [void]$result.Log.Add(
            "$(&$time) $target UNREACHABLE / OFFLINE. PING FAILED and all TCP endpoints are CLOSED."
        )

        $result.Computer = [PSCustomObject]@{
            ComputerName        = $target
            Status              = "UNREACHABLE / OFFLINE"
            CollectionMethod    = "Precheck"
            IP                  = $precheckDiagnostics.IP
            CurrentUser         = "N/A"
            CPU                 = "N/A"
            CPU_Cores           = $null
            CPU_Threads         = $null
            RAM_GB              = $null
            Manufacturer        = "Unavailable"
            Model               = "Unavailable"
            OS                  = "Host Unreachable"
            Build               = "N/A"
        }

        return [PSCustomObject]$result
    }


    #
    # ------------------------------------------------------
    # PING succeeded + all TCP ports are closed
    # ------------------------------------------------------
    #

    [void]$PrecheckQueue.Enqueue(
        [PSCustomObject]@{
            Computer = $target
            Type     = "PRECHECK_TCP_FAILED"
            Port     = $null
            Message  = "PRECHECK FAILED - PING succeeded, but TCP ports 135, 445, 5985, 5986 are CLOSED. Skipping collection."
        }
    )

    [void]$result.Log.Add(
        "$(&$time) $target PING OK, but all required TCP endpoints are CLOSED. Skipping collection."
    )

    $result.Computer = [PSCustomObject]@{
        ComputerName        = $target
        Status              = "UNREACHABLE / REMOTE ACCESS CLOSED"
        CollectionMethod    = "Precheck"
        IP                  = $precheckDiagnostics.IP
        CurrentUser         = "N/A"
        CPU                 = "N/A"
        CPU_Cores           = $null
        CPU_Threads         = $null
        RAM_GB              = $null
        Manufacturer        = "Unavailable"
        Model               = "Unavailable"
        OS                  = "Remote Access Unavailable"
        Build               = "N/A"
    }

    return [PSCustomObject]$result
}


#
# ==========================================================
# Check for LINUX / NON-WINDOWS
# ==========================================================
#

if (
    -not $precheckDiagnostics.TCP135 -and
    -not $precheckDiagnostics.TCP5985
) {

    [void]$result.Log.Add(
        "$(&$time) $target RPC 135 & WinRM 5985 CLOSED. Skipping Windows WMI/WinRM probes."
    )

    $result.Computer = [PSCustomObject]@{
        ComputerName        = $target
        Status              = "SKIPPED_NON_WINDOWS/FIREWALL_BLOCKED"
        CollectionMethod    = "Port_Check"
        IP                  = $precheckDiagnostics.IP
        CurrentUser         = "N/A"
        CPU                 = "N/A"
        CPU_Cores           = $null
        CPU_Threads         = $null
        RAM_GB              = $null
        Manufacturer        = "Non-Windows / Firewall blocked"
        Model               = "Generic Linux / NAS / Firewall blocked"
        OS                  = "Linux / Non-Windows Device / Firewall blocked"
        Build               = "N/A"
    }

    [void]$PrecheckQueue.Enqueue(
        [PSCustomObject]@{
            Computer = $target
            Type     = "NON_WINDOWS"
            Port     = $null
            Message  = "PRECHECK - TCP 135 and 5985 are CLOSED. Skipping Windows probes."
        }
    )

    return [PSCustomObject]$result
}

    #
    # SYSTEM
    #
    $wmiAvailable = $true
    $collectionMethod = "WMI"

    try {

        $system = Get-WmiObject `
            Win32_ComputerSystem `
            -ComputerName $target `
            -ErrorAction Stop

        [void]$result.Log.Add(
            "$(&$time) $target WMI SYSTEM OK"
        )
    }
    catch {

        $wmiAvailable = $false

        $wmiError = $_.Exception.Message

        [void]$result.Log.Add(
            "$(&$time) $target WMI SYSTEM FAILED: $wmiError"
        )

        #
        # DIAGNOSTICS
        #
        $diagnostics = [PSCustomObject]@{
            ComputerName   = $target
            DNS            = $false
            IP             = ""
            Ping           = $false
            TCP135         = $false
            TCP445         = $false
            TCP5985        = $false
            TCP5986        = $false
            RemoteRegistry = "Unknown"
        }

        #
        # DNS
        #
        try {

            $addresses = [System.Net.Dns]::GetHostAddresses(
                $target
            )

            $ipv4 = @(
                $addresses |
                    Where-Object {
                        $_.AddressFamily -eq
                        [System.Net.Sockets.AddressFamily]::InterNetwork
                    }
            )

            if ($ipv4.Count -gt 0) {

                $diagnostics.DNS = $true

                $diagnostics.IP = (
                    $ipv4 |
                        ForEach-Object {
                            $_.IPAddressToString
                        }
                ) -join ","
            }

        }
        catch {}

        #
        # PING
        #
        try {

            $ping = Test-Connection `
                -ComputerName $target `
                -Count 1 `
                -Quiet `
                -ErrorAction Stop

            $diagnostics.Ping = [bool]$ping
        }
        catch {}

        #
        # TCP 135
        #
        try {

            $test = Test-NetConnection `
                -ComputerName $target `
                -Port 135 `
                -WarningAction SilentlyContinue `
                -ErrorAction Stop

            $diagnostics.TCP135 = [bool]$test.TcpTestSucceeded
        }
        catch {}

        #
        # TCP 445
        #
        try {

            $test = Test-NetConnection `
                -ComputerName $target `
                -Port 445 `
                -WarningAction SilentlyContinue `
                -ErrorAction Stop

            $diagnostics.TCP445 = [bool]$test.TcpTestSucceeded
        }
        catch {}

        #
        # TCP 5985
        #
        try {

            $test = Test-NetConnection `
                -ComputerName $target `
                -Port 5985 `
                -WarningAction SilentlyContinue `
                -ErrorAction Stop

            $diagnostics.TCP5985 = [bool]$test.TcpTestSucceeded
        }
        catch {}

        #
        # TCP 5986
        #
        try {

            $test = Test-NetConnection `
                -ComputerName $target `
                -Port 5986 `
                -WarningAction SilentlyContinue `
                -ErrorAction Stop

            $diagnostics.TCP5986 = [bool]$test.TcpTestSucceeded
        }
        catch {}

        #
        # Remote Registry
        #
        try {

            $testRegistry = `
                [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey(
                    [Microsoft.Win32.RegistryHive]::LocalMachine,
                    $target
                )

            if ($testRegistry) {

                $diagnostics.RemoteRegistry = "Available"

                $testRegistry.Close()
            }

        }
        catch {

            $diagnostics.RemoteRegistry = "Unavailable"
        }

        $result.Diagnostics = $diagnostics

        if (-not $diagnostics.DNS) {

            [void]$result.Log.Add(
                "$(&$time) $target DNS FAILED"
            )
        }
        else {

            [void]$result.Log.Add(
                "$(&$time) $target DNS OK IP=$($diagnostics.IP)"
            )
        }

        #
        # FALLBACK START
        #
        [void]$result.Log.Add(
            "$(&$time) $target FALLBACK START"
        )

        $fallback = Get-FallbackSystemInfo `
            -ComputerName $target

        if ($fallback.Method -eq "WinRM") {

            [void]$result.Log.Add(
                "$(&$time) $target FALLBACK WINRM OK"
            )

            $system = [PSCustomObject]@{
                Name                = $fallback.ComputerName
                UserName            = $fallback.CurrentUser
                Manufacturer        = $fallback.Manufacturer
                Model               = $fallback.Model
                TotalPhysicalMemory = [int64](
                    $fallback.RAM_GB * 1GB
                )
            }

            $fallbackOS    = $fallback.OS
            $fallbackBuild = $fallback.Build
            $fallbackCPU   = $fallback.CPU

            $collectionMethod = "WinRM"
        }
        elseif ($fallback.Method -eq "RemoteRegistry") {

            [void]$result.Log.Add(
                "$(&$time) $target FALLBACK REMOTE REGISTRY OK"
            )

            $system = [PSCustomObject]@{
                Name                = $fallback.ComputerName
                UserName            = $null
                Manufacturer        = $fallback.Manufacturer
                Model               = $fallback.Model
                TotalPhysicalMemory = $null
            }

            $fallbackOS    = $fallback.OS
            $fallbackBuild = $fallback.Build
            $fallbackCPU   = $fallback.CPU

            $collectionMethod = "RemoteRegistry"
        }
        else {

            [void]$result.Log.Add(
                "$(&$time) $target FALLBACK FAILED: $($fallback.Error)"
            )

            throw `
                "WMI System failed: $wmiError; fallback failed: $($fallback.Error)"
        }
    }

    #
    # OS / CPU
    #
    if ($wmiAvailable) {

        $os = Get-WmiObject `
            Win32_OperatingSystem `
            -ComputerName $target `
            -ErrorAction Stop

        $cpu = Get-WmiObject `
            Win32_Processor `
            -ComputerName $target `
            -ErrorAction Stop
    }
    else {

        $os = [PSCustomObject]@{
            Caption     = $fallbackOS
            BuildNumber = $fallbackBuild
        }

        $cpu = [PSCustomObject]@{
            Name                      = $fallbackCPU
            NumberOfCores             = $null
            NumberOfLogicalProcessors = $null
        }
    }

    [void]$result.Log.Add(
        "$(&$time) $target CPU OK"
    )

    #
    # NETWORK
    #
    if ($wmiAvailable) {

        try {

            $adapters = Get-WmiObject `
                Win32_NetworkAdapterConfiguration `
                -ComputerName $target `
                -ErrorAction Stop |
                Where-Object {
                    $_.IPEnabled
                }

            foreach ($adapter in @($adapters)) {

                foreach ($ip in @($adapter.IPAddress)) {

                    if ($ip -match "^\d+\.") {

                        [void]$result.Network.Add(
                            [PSCustomObject]@{
                                ComputerName = $target
                                Adapter      = $adapter.Description
                                MAC          = $adapter.MACAddress
                                IP           = $ip
                                Gateway      = (
                                    $adapter.DefaultIPGateway -join ","
                                )
                                DHCP         = $adapter.DHCPEnabled
                            }
                        )
                    }
                }
            }

            [void]$result.Log.Add(
                "$(&$time) $target NETWORK OK"
            )
        }
        catch {

            [void]$result.Log.Add(
                "$(&$time) $target NETWORK FAILED $($_.Exception.Message)"
            )
        }
    }
    else {

        [void]$result.Log.Add(
            "$(&$time) $target NETWORK SKIPPED (WMI Unavailable)"
        )
    }

    #
    # USERS
    #
    if ($wmiAvailable) {
        try {
            # Get the current interactive user, if defined in $system
            $activeUser = $null
            if ($system -and $system.UserName) {
                $activeUser = $system.UserName.Trim()
            }

            # Exclude system profiles (SYSTEM, LocalService, NetworkService)
            $profiles = Get-WmiObject Win32_UserProfile -ComputerName $target -ErrorAction Stop |
                Where-Object { -not $_.Special -and $_.SID -match '^S-1-5-21-' }

            foreach ($prof in $profiles) {
                $accountName = $prof.SID
                try {
                    # Convert SID to DOMAIN\Username format
                    $sidObj = New-Object System.Security.Principal.SecurityIdentifier($prof.SID)
                    $accountName = $sidObj.Translate([System.Security.Principal.NTAccount]).Value
                }
                catch {
                    # If the domain is unavailable for resolution, keep the raw SID
                }

                #
                # Instead of $prof.Loaded, check the actual match with the logged-in user
                #
                $isRealCurrentUser = $false
                if ($activeUser -and $accountName) {
                    $isRealCurrentUser = ($accountName -eq $activeUser)
                }

                # Convert the last logon date (LastUseTime)
                $lastUse = $null
                if ($prof.LastUseTime) {
                    $lastUse = [System.Management.ManagementDateTimeConverter]::ToDateTime($prof.LastUseTime)
                }

                [void]$result.Users.Add(
                    [PSCustomObject]@{
                        ComputerName  = $target
                        User          = $accountName
                        CurrentUser   = $isRealCurrentUser   # True ONLY for the actually logged-in user
                        LastUseTime   = $lastUse
                        SID           = $prof.SID
                        LocalPath     = $prof.LocalPath
                    }
                )
            }

            [void]$result.Log.Add(
                "$(&$time) $target USERS OK (Count: $($profiles.Count))"
            )
        }
        catch {
            [void]$result.Log.Add(
                "$(&$time) $target USERS FAILED $($_.Exception.Message)"
            )
        }
    }
    else {
        [void]$result.Log.Add(
            "$(&$time) $target USERS SKIPPED (WMI Unavailable)"
        )
    }

    #
    # PRINTERS
    #
    if ($wmiAvailable) {

        try {

            $printers = Get-WmiObject `
                Win32_Printer `
                -ComputerName $target `
                -ErrorAction Stop

            $ports = Get-WmiObject `
                Win32_TCPIPPrinterPort `
                -ComputerName $target `
                -ErrorAction SilentlyContinue

            #
            # Create the port lookup table.
            #
            $portMap = @{}

            foreach ($port in @($ports)) {

                if ($port.Name) {
                    $portMap[$port.Name] = $port.HostAddress
                }
            }

            foreach ($printer in @($printers)) {

                if (
                    $printer.Name -match
                    "Fax|PDF|XPS|OneNote|Microsoft|Send To"
                ) {
                    continue
                }

                $printerIP = ""

                if (
                    $printer.PortName -and
                    $portMap.ContainsKey($printer.PortName)
                ) {
                    $printerIP = $portMap[$printer.PortName]
                }

                [void]$result.Printers.Add(
                    [PSCustomObject]@{
                        ComputerName = $target
                        Printer      = $printer.Name
                        Driver       = $printer.DriverName
                        Port         = $printer.PortName
                        IP           = $printerIP
                    }
                )
            }

            [void]$result.Log.Add(
                "$(&$time) $target PRINTERS OK"
            )
        }
        catch {

            [void]$result.Log.Add(
                "$(&$time) $target PRINTERS FAILED $($_.Exception.Message)"
            )
        }
    }
    else {

        [void]$result.Log.Add(
            "$(&$time) $target PRINTERS SKIPPED (WMI Unavailable)"
        )
    }

    #
    # PERIPHERAL DEVICES
    #
    if ($wmiAvailable) {

        try {

            $targetClasses = @(
                'Image',
                'SmartCardReader',
                'Ports',
                'POS',
                'Biometric'
            )

            $devices = Get-WmiObject `
                Win32_PnPEntity `
                -ComputerName $target `
                -ErrorAction Stop |
                Where-Object {
                    $_.PNPClass -in $targetClasses -and
                    $_.Present -eq $true -and
                    $_.Status -eq 'OK'
                }

            $badIDs = '^BTHENUM|^OXPCIEMF|^ACPI|^ROOT|^SWD|^LPT|^PNP05'

            $badNames = `
                'HP|Hewlett|LaserJet|OfficeJet|PageWide|WorkCentre|Kyocera|Xerox|Canon|Epson|Brother|WIA Driver|Microsoft|Standard Serial|Communications Port'

            $badRu = `
                '[\u0421\u0442\u0430\u043d\u0434\u0430\u0440\u0442\u043d\u044b\u0439|\u0423\u0441\u0442\u0440\u043e\u0439\u0441\u0442\u0432\u043e|\u041f\u043e\u0441\u043b\u0435\u0434\u043e\u0432\u0430\u0442\u0435\u043b\u044c\u043d\u044b\u0439|\u041f\u043e\u0440\u0442|\u041f\u043e\u0434\u0434\u0435\u0440\u0436\u043a\u0430|\u041c\u0430\u0439\u043a\u0440\u043e\u0441\u043e\u0444\u0442]'

            $badSvcs = `
                'kbdhid|mouhid|usbccgp|usbehci|usbhub|usbprint|vscdrv|serenum|parport'

            foreach ($dev in @($devices)) {

                if ($dev.DeviceID -match $badIDs) {
                    continue
                }

                if (
                    $dev.Name -match $badNames -or
                    $dev.Name -match $badRu -or
                    $dev.Manufacturer -match 'Microsoft|[\u041c\u0430\u0439\u043a\u0440\u043e\u0441\u043e\u0444\u0442]'
                ) {
                    continue
                }

                if ($dev.Service -match $badSvcs) {
                    continue
                }

                [void]$result.Devices.Add(
                    [PSCustomObject]@{
                        ComputerName = $target
                        Name         = $dev.Name
                        Manufacturer = $dev.Manufacturer
                        Class        = $dev.PNPClass
                        DeviceID     = $dev.DeviceID
                        Status       = $dev.Status
                    }
                )
            }

            [void]$result.Log.Add(
                "$(&$time) $target DEVICES OK"
            )
        }
        catch {

            [void]$result.Log.Add(
                "$(&$time) $target DEVICES FAILED $($_.Exception.Message)"
            )
        }
    }
    else {

        [void]$result.Log.Add(
            "$(&$time) $target DEVICES SKIPPED (WMI Unavailable)"
        )
    }

    #
    # COMPUTER INFORMATION
    #
    try {

        if (-not $collectionMethod) {
            $collectionMethod = "WMI"
        }

        $cpuName =
            if ($cpu) {

                if ($cpu -is [array]) {
                    $cpu[0].Name
                }
                else {
                    $cpu.Name
                }

            }
            else {
                $null
            }

        #
        # CPU cores / threads
        #
        $cpuCores = $null
        $cpuThreads = $null

        if ($wmiAvailable -and $cpu) {

            $cpuCores = 0
            $cpuThreads = 0

            foreach ($processor in @($cpu)) {

                if ($null -ne $processor.NumberOfCores) {
                    $cpuCores += [int]$processor.NumberOfCores
                }

                if ($null -ne $processor.NumberOfLogicalProcessors) {
                    $cpuThreads += [int]$processor.NumberOfLogicalProcessors
                }
            }

            if ($cpuCores -eq 0) {
                $cpuCores = $null
            }

            if ($cpuThreads -eq 0) {
                $cpuThreads = $null
            }
        }

        #
        # IP
        #
        $ipAddress = ""

        if (
            $result.Network.Count -gt 0
        ) {

            $ipList = [System.Collections.Generic.List[string]]::new()

            foreach ($networkItem in $result.Network) {

                if ($networkItem.IP) {
                    [void]$ipList.Add(
                        [string]$networkItem.IP
                    )
                }
            }

            if ($ipList.Count -gt 0) {
                $ipAddress = $ipList -join ","
            }
        }

        if (
            -not $ipAddress -and
            $result.Diagnostics -and
            $result.Diagnostics.IP
        ) {
            $ipAddress = $result.Diagnostics.IP
        }

        #
        # Computer name
        #
        $finalCompName =
            if ($system.Name) {
                $system.Name
            }
            else {
                $target
            }

        $result.Computer = [PSCustomObject]@{

            ComputerName = $finalCompName

            Status =
                if ($collectionMethod -eq "WMI") {
                    "OK"
                }
                elseif ($collectionMethod -eq "WinRM") {
                    "FALLBACK"
                }
                else {
                    "PARTIAL"
                }

            CollectionMethod = $collectionMethod

            IP = $ipAddress

            CurrentUser = $system.UserName

            CPU = $cpuName

            CPU_Cores = $cpuCores

            CPU_Threads = $cpuThreads

            RAM_GB =
                if ($system.TotalPhysicalMemory) {

                    [math]::Round(
                        $system.TotalPhysicalMemory / 1GB,
                        2
                    )

                }
                else {
                    $null
                }

            Manufacturer = $system.Manufacturer

            Model = $system.Model

            OS = $os.Caption

            Build = $os.BuildNumber
        }

        [void]$result.Log.Add(
            "$(&$time) $target COMPUTER DATA OK"
        )
    }
    catch {

        throw `
            "Computer data collection failed: $($_.Exception.Message)"
    }

    [void]$result.Log.Add(
        "$(&$time) FINISH $target"
    )
}
catch {

    $msg = $_.Exception.Message

    $result.Error = $msg

    [void]$result.Log.Add(
        "$(&$time) $target ERROR $msg"
    )

    $result.Computer = [PSCustomObject]@{
        ComputerName     = $ComputerName
        Status            = "ERROR"
        CollectionMethod = "FAILED"
        IP                = ""
        CurrentUser      = ""
        CPU              = ""
        CPU_Cores        = ""
        CPU_Threads      = ""
        RAM_GB           = ""
        Manufacturer     = ""
        Model            = ""
        OS               = ""
        Build            = ""
    }
}

#
# Runspace returns the result as an object.
#
[PSCustomObject]$result
'@

#
# Get computers from the AD domain
#
$computers = Get-ADComputer `
    -Filter * `
    -Properties DNSHostName, OperatingSystem, LastLogonDate

#
# Existing worksheets in the Excel workbook
#
$ExistingSheets = @{}
$ExistingComputerRows = @()
$ExistingNetworkRows = @()
$ExistingUserRows = @()
$ExistingPrinterRows = @()
$ExistingDeviceRows = @()
$ExistingDiagnosticsRows = @()
$ExistingErrorRows = @()

function Import-ExistingSheet {
    param([string]$Path, [string]$WorksheetName)
    if (-not (Test-Path -LiteralPath $Path)) { return @() }
    try {
        $rows = @(Import-Excel -Path $Path -WorksheetName $WorksheetName -ErrorAction Stop)
        return $rows
    }
    catch {
        Write-Host "WARNING: Cannot read worksheet ${WorksheetName}: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Log "WARNING: Cannot read worksheet ${WorksheetName}: $($_.Exception.Message)"
        return @()
    }
}

function Get-RowKey {
    param($Row)
    if ($null -eq $Row -or $null -eq $Row.ComputerName) { return "" }
    return ([string]$Row.ComputerName).Trim().ToUpperInvariant()
}

function Convert-ExcelSerialDate {
    param([object]$Value)

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [datetime]) {
        return $Value
    }

    if ($Value -is [double] -or
        $Value -is [decimal] -or
        $Value -is [float] -or
        $Value -is [int] -or
        $Value -is [long]) {

        try {
            $number = [double]$Value

            if ($number -ge 1 -and $number -lt 100000) {
                return [datetime]::FromOADate($number)
            }
        }
        catch {
        }
    }

    return $Value
}

if (Test-Path -LiteralPath $output) {
    Write-Host "Existing inventory found. Loading workbook..." -ForegroundColor Cyan
    Write-Log "Existing workbook found: $output"
    $ExistingComputerRows = @(Import-ExistingSheet $output "Computers")
    $ExistingNetworkRows = @(Import-ExistingSheet $output "Network")
    $ExistingUserRows = @(Import-ExistingSheet $output "Users")
    $ExistingPrinterRows = @(Import-ExistingSheet $output "Printers")
    $ExistingDeviceRows = @(Import-ExistingSheet $output "Devices")
    $ExistingDiagnosticsRows = @(Import-ExistingSheet $output "Diagnostics")
    $ExistingErrorRows = @(Import-ExistingSheet $output "Errors")
}

foreach ($row in @($ExistingUserRows)) {
    if ($null -eq $row) {
        continue
    }

    if ($row.PSObject.Properties['LastUseTime']) {
        $row.LastUseTime = Convert-ExcelSerialDate $row.LastUseTime
    }
}

$ExistingByComputer = @{}
foreach ($row in $ExistingComputerRows) {
    $key = Get-RowKey $row
    if ($key) { $ExistingByComputer[$key] = $row }
}

$computersToScan = @($computers)
$computers = @($computersToScan)

if ($TestMode) {
    $computers = $computers | Select-Object -First 30
}

$total = $computers.Count

Write-Host "Computers found: $total"
Write-Host "Threads: $Threads"
Write-Host "Job timeout: $JobTimeoutSeconds seconds"
Write-Host "Stop grace: $JobStopGraceSeconds seconds"
Write-Host "Max abandoned runspaces per pool: $MaxAbandonedRunspaces"
Write-Host "Execution engine: RunspacePool"
Write-Host ""

Write-Log "Computers found: $total"
Write-Log "Job timeout: $JobTimeoutSeconds seconds"
Write-Log "Stop grace: $JobStopGraceSeconds seconds"

#
# Results
#
$ComputerResults = [System.Collections.Generic.List[object]]::new()
$NetworkResults = [System.Collections.Generic.List[object]]::new()
$UserResults = [System.Collections.Generic.List[object]]::new()
$PrinterResults = [System.Collections.Generic.List[object]]::new()
$DeviceResults = [System.Collections.Generic.List[object]]::new()
$ErrorResults = [System.Collections.Generic.List[object]]::new()
$DiagnosticsResults = [System.Collections.Generic.List[object]]::new()

#
# Queue
#
$queue = New-Object System.Collections.Queue

foreach ($pc in $computers) {
    [void]$queue.Enqueue($pc)
}

#
# PRECHECK event queue.
#
# Workers write events here,
# the main thread outputs them to the console.
#
$precheckQueue = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()

#
# RunspacePool
#
function New-InventoryRunspacePool {

    Write-Host `
        "Creating RunspacePool..." `
        -ForegroundColor Cyan

    Write-Log `
        "Creating RunspacePool Min=1 Max=$Threads"

    $pool = [runspacefactory]::CreateRunspacePool(
        1,
        $Threads
    )

    $pool.Open()

    Write-Log "RunspacePool OPENED"

    return $pool
}

$runspacePool = New-InventoryRunspacePool

#
# Active runspace jobs
#
$active = New-Object System.Collections.ArrayList

$count = 0
$completed = 0

#
# All abandoned runspaces throughout the entire script execution.
#
$abandoned = 0

#
# Abandoned runspaces since the current RunspacePool was created.
#
$abandonedSincePoolRestart = 0

#
# RunspacePool generation (RunspacePool sequence number).
#
$poolGeneration = 1

Write-Log `
    "RunspacePool generation=$poolGeneration ACTIVE"

#
# START WORKER
#
function Start-InventoryWorker {

    param(
        [Parameter(Mandatory = $true)]
        $Computer
    )

    $script:count++

    $name = $Computer.Name

    Write-Host `
        "[$script:count/$total] START $name" `
        -ForegroundColor Cyan

    Write-Log `
        "START RUNSPACE $name PoolGeneration=$script:poolGeneration"

    try {

        $ps = [powershell]::Create()

        $ps.RunspacePool = $runspacePool

        [void]$ps.AddScript($worker)

        [void]$ps.AddParameter(
            "ComputerName",
            $name
        )

        [void]$ps.AddParameter(
            "PrecheckQueue",
            $precheckQueue
        )

        $asyncResult = $ps.BeginInvoke()

        $job = [PSCustomObject]@{
            Name            = $name
            PowerShell      = $ps
            AsyncResult     = $asyncResult
            Started         = Get-Date
            StopRequested   = $false
            StopStarted     = $null
            Abandoned       = $false
            PoolGeneration  = $script:poolGeneration
        }

        [void]$active.Add($job)

        Write-Log `
            "$name RUNSPACE STARTED PoolGeneration=$script:poolGeneration"

    }
    catch {

        $msg = $_.Exception.Message

        Write-Host `
            "FAIL START $name - $msg" `
            -ForegroundColor Red

        Write-Log `
            "$name RUNSPACE START FAILED: $msg"

        [void]$ErrorResults.Add(
            [PSCustomObject]@{
                ComputerName = $name
                Error        = "Failed to start runspace: $msg"
            }
        )
    }
}

#
# COMPLETE WORKER
#
function Complete-InventoryWorker {

    param(
        [Parameter(Mandatory = $true)]
        $Job
    )

    $name = $Job.Name
    $ps = $Job.PowerShell

    try {

        #
        # EndInvoke is called only after IsCompleted.
        #
        $output = $null

        try {

            $output = $ps.EndInvoke(
                $Job.AsyncResult
            )

        }
        catch {

            if ($Job.StopRequested) {

                $stopMessage = $_.Exception.Message

                [void]$ErrorResults.Add(
                    [PSCustomObject]@{
                        ComputerName = $name
                        Error        = "Timeout after $JobTimeoutSeconds seconds"
                    }
                )

                Write-Host `
                    "TIMEOUT COMPLETE $name" `
                    -ForegroundColor Yellow

                Write-Log `
                    "$name PIPELINE STOPPED: $stopMessage"

                Write-Log `
                    "$name TIMEOUT COMPLETE"

                return
            }

            throw
        }

        #
        # If the worker finishes after the watchdog,
        # this is still a timeout.
        #
        if ($Job.StopRequested) {

            [void]$ErrorResults.Add(
                [PSCustomObject]@{
                    ComputerName = $name
                    Error        = "Timeout after $JobTimeoutSeconds seconds"
                }
            )

            Write-Host `
                "TIMEOUT COMPLETE $name" `
                -ForegroundColor Yellow

            Write-Log `
                "$name PIPELINE COMPLETED AFTER STOP REQUEST"

            return
        }

        #
        # Worker without output.
        #
        if (
            $null -eq $output -or
            $output.Count -eq 0
        ) {

            $state = $ps.InvocationStateInfo.State

            $reason = ""

            if ($ps.InvocationStateInfo.Reason) {
                $reason = $ps.InvocationStateInfo.Reason.Exception.Message
            }

            throw `
                "Worker completed without output. State=$state Reason=$reason"
        }

        #
        # Worker returns exactly one result object.
        #
        $data = $output |
            Select-Object -Last 1

        if ($null -eq $data) {
            throw "Worker returned null result."
        }

        #
        # Structure check.
        #
        if (
            -not $data.PSObject.Properties["Computer"] -or
            -not $data.PSObject.Properties["Network"] -or
            -not $data.PSObject.Properties["Users"] -or
            -not $data.PSObject.Properties["Printers"] -or
            -not $data.PSObject.Properties["Devices"]
        ) {

            throw `
                "Invalid worker result object. Type=$($data.GetType().FullName)"
        }

        #
        # Worker log
        #
        if ($data.Log) {

            foreach ($line in @($data.Log)) {

                Add-Content `
                    -Path $logFile `
                    -Value $line
            }
        }

        #
        # Computer
        #
        if ($data.Computer) {

            [void]$ComputerResults.Add(
                $data.Computer
            )
        }

        #
        # Network
        #
        if (
            $data.Network -and
            $data.Network.Count -gt 0
        ) {

            foreach ($item in @($data.Network)) {
                [void]$NetworkResults.Add($item)
            }
        }

        #
        # Users
        #
        if (
            $data.Users -and
            $data.Users.Count -gt 0
        ) {

            foreach ($item in @($data.Users)) {
                [void]$UserResults.Add($item)
            }
        }

        #
        # Printers
        #
        if (
            $data.Printers -and
            $data.Printers.Count -gt 0
        ) {

            foreach ($item in @($data.Printers)) {
                [void]$PrinterResults.Add($item)
            }
        }

        #
        # Devices
        #
        if (
            $data.Devices -and
            $data.Devices.Count -gt 0
        ) {

            foreach ($item in @($data.Devices)) {
                [void]$DeviceResults.Add($item)
            }
        }

        #
        # Diagnostics
        #
        if ($data.Diagnostics) {

            [void]$DiagnosticsResults.Add(
                $data.Diagnostics
            )
        }

        #
        # Error
        #
        if ($data.Error) {

            [void]$ErrorResults.Add(
                [PSCustomObject]@{
                    ComputerName = $name
                    Error        = $data.Error
                }
            )

            Write-Host `
                "ERROR $name" `
                -ForegroundColor Red

            Write-Log `
                "$name WORKER ERROR: $($data.Error)"
        }
        elseif (
            $data.Computer -and
            $data.Computer.Status -eq "UNREACHABLE / OFFLINE"
        ) {

            Write-Log `
                "$name PRECHECK FAILED - DONE suppressed"

        }
        else {

            Write-Host `
                "DONE $name" `
                -ForegroundColor Green
        }

        Write-Log `
            "$name RUNSPACE COMPLETE State=$($ps.InvocationStateInfo.State) PoolGeneration=$($Job.PoolGeneration)"
    }
    catch {

        $msg = $_.Exception.Message

        $state = "Unknown"

        try {
            $state = $ps.InvocationStateInfo.State
        }
        catch {}

        $reason = ""

        try {
            if ($ps.InvocationStateInfo.Reason) {
                $reason = $ps.InvocationStateInfo.Reason.Exception.Message
            }
        }
        catch {}

        [void]$ErrorResults.Add(
            [PSCustomObject]@{
                ComputerName = $name
                Error        = "Result processing failed: $msg"
            }
        )

        Write-Host `
            "FAIL $name - result processing: $msg" `
            -ForegroundColor Red

        Write-Log `
            "$name RESULT PROCESSING FAILED: $msg"

        Write-Log `
            "$name RUNSPACE STATE=$state REASON=$reason"
    }
    finally {

        try {
            $ps.Dispose()
        }
        catch {}

        $script:completed++

        Write-Log `
            "COMPLETED $name ($script:completed/$total)"
    }
}

#
# STOP WORKER
#
function Stop-InventoryWorker {

    param(
        [Parameter(Mandatory = $true)]
        $Job
    )

    $name = $Job.Name
    $ps = $Job.PowerShell

    if ($Job.StopRequested) {
        return
    }

    $Job.StopRequested = $true
    $Job.StopStarted = Get-Date

    Write-Host `
        "TIMEOUT $name - requesting pipeline stop..." `
        -ForegroundColor Yellow

    Write-Log `
        "$name TIMEOUT after $JobTimeoutSeconds seconds"

    try {

        [void]$ps.BeginStop(
            $null,
            $null
        )

        Write-Log `
            "$name BeginStop() requested"
    }
    catch {

        $msg = $_.Exception.Message

        Write-Log `
            "$name BeginStop FAILED: $msg"

        [void]$ErrorResults.Add(
            [PSCustomObject]@{
                ComputerName = $name
                Error        = "Timeout after $JobTimeoutSeconds seconds; BeginStop failed: $msg"
            }
        )

        $Job.StopStarted = Get-Date
    }
}

#
# ABANDON WORKER
#
function Abandon-InventoryWorker {

    param(
        [Parameter(Mandatory = $true)]
        $Job
    )

    $name = $Job.Name

    if ($Job.Abandoned) {
        return
    }

    $Job.Abandoned = $true

    #
    # All abandoned runspaces throughout the entire script execution.
    #
    $script:abandoned++

    #
    # Abandoned runspaces since the current RunspacePool was created.
    #
    $script:abandonedSincePoolRestart++

    [void]$ErrorResults.Add(
        [PSCustomObject]@{
            ComputerName = $name
            Error        = "Timeout after $JobTimeoutSeconds seconds; pipeline did not stop within $JobStopGraceSeconds seconds"
        }
    )

    Write-Host `
        "ABANDON $name - pipeline did not stop" `
        -ForegroundColor Red

    Write-Log `
        "$name ABANDONED after $JobStopGraceSeconds second stop grace"

    Write-Log `
        "$name ABANDONED RUNSPACE MAY REMAIN BUSY"

    Write-Log `
        "Abandoned since current pool: $script:abandonedSincePoolRestart"

    Write-Log `
        "Total abandoned runspaces: $script:abandoned"
}

#
# RESTART RUNSPACE POOL
#
function Restart-InventoryRunspacePool {

    Write-Host ""
    Write-Host `
        "WARNING: abandoned runspace threshold reached for current pool. Creating a new RunspacePool..." `
        -ForegroundColor Magenta

    Write-Log `
        "RUNSPACE POOL RESTART requested. PoolGeneration=$script:poolGeneration AbandonedSinceRestart=$script:abandonedSincePoolRestart TotalAbandoned=$script:abandoned"

    #
    # The old pool is intentionally not closed to avoid hanging the entire script.
    #
    $oldPool = $script:runspacePool

    $script:runspacePool = New-InventoryRunspacePool

    $script:poolGeneration++

    #
    # The new pool starts its own abandoned counter.
    #
    $script:abandonedSincePoolRestart = 0

    Write-Log `
        "NEW RunspacePool created. PoolGeneration=$script:poolGeneration"

    Write-Log `
        "AbandonedSinceRestart reset to 0"

    $oldPool = $null
}

#
# MAIN QUEUE
#
while (
    $queue.Count -gt 0 -or
    $active.Count -gt 0
) {

    #
    # ==========================================================
    # PRECHECK EVENTS
    #
    # Workers asynchronously add events here.
    # The main thread outputs them to the console.
    # ==========================================================
    #

    $precheckEvent = $null

    while (
        $precheckQueue.TryDequeue(
            [ref]$precheckEvent
        )
    ) {

        switch ($precheckEvent.Type) {

            "PING_OK" {

                Write-Host `
                    "[$($precheckEvent.Computer)] $($precheckEvent.Message)" `
                    -ForegroundColor DarkCyan
            }

            "PING_FAILED" {

                Write-Host `
                    "[$($precheckEvent.Computer)] $($precheckEvent.Message)" `
                    -ForegroundColor Yellow
            }

            "TCP_OK" {

                Write-Host `
                    "[$($precheckEvent.Computer)] $($precheckEvent.Message)" `
                    -ForegroundColor DarkGreen
            }

            "PRECHECK_FAILED" {

                Write-Host `
                    "[$($precheckEvent.Computer)] $($precheckEvent.Message)" `
                    -ForegroundColor Red
            }
        }

        $precheckEvent = $null
    }

    #
    # If too many runspaces were abandoned
    # specifically in the current pool,
    # create a new pool.
    #
    if (
        $script:abandonedSincePoolRestart -ge
        $MaxAbandonedRunspaces
    ) {

        if (
            $queue.Count -gt 0 -and
            $active.Count -ge $Threads
        ) {

            Restart-InventoryRunspacePool
        }
    }

    #
    # Start workers up to the Threads limit.
    #
    while (
        $queue.Count -gt 0 -and
        $active.Count -lt $Threads
    ) {

        $computer = $queue.Dequeue()

        Start-InventoryWorker `
            -Computer $computer
    }

    #
    # Check active runspace jobs.
    #
    for (
        $i = $active.Count - 1;
        $i -ge 0;
        $i--
    ) {

        $job = $active[$i]
        $ps = $job.PowerShell

        #
        # Pipeline completed.
        #
        if ($job.AsyncResult.IsCompleted) {

            Write-Log `
                "$($job.Name) ASYNC COMPLETED"

            Complete-InventoryWorker `
                -Job $job

            [void]$active.RemoveAt($i)

            continue
        }

        #
        # Stop has already been requested.
        #
        if ($job.StopRequested) {

            if ($job.StopStarted) {

                $stopElapsed = (
                    (Get-Date) - $job.StopStarted
                ).TotalSeconds

                if (
                    $stopElapsed -ge
                    $JobStopGraceSeconds
                ) {

                    Abandon-InventoryWorker `
                        -Job $job

                    [void]$active.RemoveAt($i)

                    continue
                }
            }

            continue
        }

        #
        # Check the main timeout.
        #
        $elapsed = (
            (Get-Date) - $job.Started
        ).TotalSeconds

        if ($elapsed -ge $JobTimeoutSeconds) {

            Stop-InventoryWorker `
                -Job $job
        }
    }

    #
    # Short pause.
    #
    if (
        $queue.Count -gt 0 -or
        $active.Count -gt 0
    ) {

        Start-Sleep -Milliseconds 250
    }
}

$precheckEvent = $null

while (
    $precheckQueue.TryDequeue(
        [ref]$precheckEvent
    )
) {

    switch ($precheckEvent.Type) {

        "PING_OK" {

            Write-Host `
                "[$($precheckEvent.Computer)] $($precheckEvent.Message)" `
                -ForegroundColor DarkCyan
        }

        "PING_FAILED" {

            Write-Host `
                "[$($precheckEvent.Computer)] $($precheckEvent.Message)" `
                -ForegroundColor Yellow
        }

        "TCP_OK" {

            Write-Host `
                "[$($precheckEvent.Computer)] $($precheckEvent.Message)" `
                -ForegroundColor DarkGreen
        }

        "PRECHECK_FAILED" {

            Write-Host `
                "[$($precheckEvent.Computer)] $($precheckEvent.Message)" `
                -ForegroundColor Red
        }
    }

    $precheckEvent = $null
}

#
# Completion.
#
Write-Log "All non-abandoned workers completed."
Write-Log "Total abandoned runspaces: $abandoned"
Write-Log "Abandoned in current pool: $abandonedSincePoolRestart"
Write-Log "Current RunspacePool generation: $poolGeneration"

#
# If there are no abandoned runspaces in the CURRENT pool,
# it can be closed normally.
#
if ($abandonedSincePoolRestart -eq 0) {

    Write-Log `
        "No abandoned runspaces in current pool. Closing RunspacePool."

    try {
        $runspacePool.Close()
    }
    catch {}

    try {
        $runspacePool.Dispose()
    }
    catch {}

    Write-Log "Current RunspacePool CLOSED"
}
else {

    #
    # The current pool contains abandoned runspaces.
    # Close() is not called to avoid risking a hang.
    #
    Write-Log `
        "Abandoned runspaces detected in current pool ($abandonedSincePoolRestart). Current RunspacePool left without Close/Dispose to avoid blocking."
}

#
# Merge current results into existing inventory
#
function Test-UnavailableComputerStatus {
    param([object]$Row)

    if ($null -eq $Row) { return $false }

    $status = [string]$Row.Status

    return (
        $status -eq "UNREACHABLE / OFFLINE" -or
        $status -eq "UNREACHABLE / REMOTE ACCESS CLOSED" -or
        $status -eq "SKIPPED_NON_WINDOWS/FIREWALL_BLOCKED"
    )
}

function Get-InventoryAvailabilityValue {
    param([object]$Row)

    if ($null -eq $Row) { return "" }

    $property = $Row.PSObject.Properties |
        Where-Object {
            $_.Name -in @(
                "Availability",
                "AvailabilityStatus",
                "CurrentStatus"
            )
        } |
        Select-Object -First 1

    if ($property) {
        return [string]$property.Value
    }

    return ""
}

function Test-PreviouslyUnavailableComputer {
    param([object]$Row)

    if ($null -eq $Row) { return $false }

    $availability = Get-InventoryAvailabilityValue $Row

    if ($availability -like "*CURRENTLY UNAVAILABLE*") {
        return $true
    }

    return (Test-UnavailableComputerStatus $Row)
}

function Add-OfflineInventoryNote {
    param([object]$Row)

    if ($null -eq $Row) { return }

    $note = "CURRENTLY UNAVAILABLE - showing last collected information"

    $statusProperty = $Row.PSObject.Properties |
        Where-Object {
            $_.Name -eq "Availability"
        } |
        Select-Object -First 1

    if ($statusProperty) {
        $statusProperty.Value = $note
    }
    else {
        $Row | Add-Member `
            -NotePropertyName "Availability" `
            -NotePropertyValue $note `
            -Force
    }
}

function Clear-OfflineInventoryNote {
    param([object]$Row)

    if ($null -eq $Row) { return }

    $statusProperty = $Row.PSObject.Properties |
        Where-Object {
            $_.Name -eq "Availability"
        } |
        Select-Object -First 1

    if ($statusProperty) {
        $statusProperty.Value = ""
    }
    else {
        $Row | Add-Member `
            -NotePropertyName "Availability" `
            -NotePropertyValue "" `
            -Force
    }
}

function Merge-ComputerRows {
    param($OldRows, $NewRows)

    $map = @{}

    foreach ($row in @($OldRows)) {
        $key = Get-RowKey $row

        if ($key) {
            $map[$key] = $row
        }
    }

    foreach ($row in @($NewRows)) {
        $key = Get-RowKey $row

        if (-not $key) {
            continue
        }

        if (Test-UnavailableComputerStatus $row) {

            if ($map.ContainsKey($key)) {

                # A previously available PC is currently unavailable.
                # Leave a note about this in the table.
                if (-not (Test-PreviouslyUnavailableComputer $map[$key])) {
                    Add-OfflineInventoryNote $map[$key]
                }

                continue
            }

            # A PC that was not previously added to the table is unavailable.

            $map[$key] = $row
            continue
        }

        # PC is available.
		# Remove the unavailability note, if one was added.
		Clear-OfflineInventoryNote $row

		# Preserve manually entered comments from the existing workbook.
		if ($map.ContainsKey($key)) {
			$oldRow = $map[$key]

			$commentProperty = $oldRow.PSObject.Properties |
				Where-Object { $_.Name -eq "Комментарий" } |
				Select-Object -First 1

		if ($commentProperty) {
			$existingComment = $commentProperty.Value

			$newCommentProperty = $row.PSObject.Properties |
				Where-Object { $_.Name -eq "Комментарий" } |
				Select-Object -First 1

        if ($newCommentProperty) {
            $newCommentProperty.Value = $existingComment
        }
        else {
            $row | Add-Member `
                -NotePropertyName "Комментарий" `
                -NotePropertyValue $existingComment `
                -Force
        }
    }
}

$map[$key] = $row
    }

    return @(
        $map.Values |
            Sort-Object ComputerName
    )
}

function Merge-DetailRows {
    param(
        $OldRows,
        $NewRows,
        $ScannedComputerNames,
        [switch]$ReplaceUnavailableWithNew
    )

    $scanned = @{}

    foreach ($name in @($ScannedComputerNames)) {
        if ($name) {
            $scanned[([string]$name).Trim().ToUpperInvariant()] = $true
        }
    }

    $newKeys = @{}

    if ($ReplaceUnavailableWithNew) {
        foreach ($row in @($NewRows)) {
            $key = Get-RowKey $row
            if ($key) {
                $newKeys[$key] = $true
            }
        }
    }

    $out = [System.Collections.Generic.List[object]]::new()

    foreach ($row in @($OldRows)) {
        $key = Get-RowKey $row

        # If the PC is available: records are updated
        if ($scanned.ContainsKey($key)) {
            continue
        }

        # Diagnostics/Errors: if the PC is unavailable and the error has changed
        # replace the old record with the new one.
        # If the error is identical to the previous one, keep the old record.
        if ($ReplaceUnavailableWithNew -and $newKeys.ContainsKey($key)) {
            continue
        }

        # If the PC has disappeared from the domain, keep its record.
        [void]$out.Add($row)
    }

    foreach ($row in @($NewRows)) {
        [void]$out.Add($row)
    }

    return @($out)
}

$currentUserByComputer = @{}
foreach ($userRow in @($UserResults)) {
    if ($userRow.CurrentUser -eq $true -and $userRow.User) {
        $userComputerKey = ([string]$userRow.ComputerName).Trim().ToUpperInvariant()
        if ($userComputerKey) {
            $currentUserByComputer[$userComputerKey] = [string]$userRow.User
        }
    }
}

foreach ($computerRow in @($ComputerResults)) {
    if ($null -eq $computerRow) { continue }
    $computerKey = ([string]$computerRow.ComputerName).Trim().ToUpperInvariant()
    if ($computerKey -and $currentUserByComputer.ContainsKey($computerKey)) {
        $computerRow.CurrentUser = $currentUserByComputer[$computerKey]
    }
    elseif ($computerRow.PSObject.Properties['CurrentUser']) {
        $computerRow.CurrentUser = ''
    }
}

$unavailableComputerKeys = @{}

foreach ($computerRow in @($ComputerResults)) {
    if ($null -eq $computerRow) { continue }

    $computerKey = ([string]$computerRow.ComputerName).Trim().ToUpperInvariant()

    if ($computerKey -and (Test-UnavailableComputerStatus $computerRow)) {
        $unavailableComputerKeys[$computerKey] = $true
    }
}

# If the PC is available during the current poll, update the data in the table.
# If the computer is unavailable, but information was previously collected in Network/Users/Printers/Devices/
# Diagnostics/Errors, do not update the information.
$scannedNames = @(
    $computers |
        ForEach-Object { $_.Name } |
        Where-Object {
            $key = ([string]$_).Trim().ToUpperInvariant()
            -not $unavailableComputerKeys.ContainsKey($key)
        }
)
$ComputerResults = [System.Collections.Generic.List[object]](Merge-ComputerRows $ExistingComputerRows $ComputerResults)
$NetworkResults = [System.Collections.Generic.List[object]](Merge-DetailRows $ExistingNetworkRows $NetworkResults $scannedNames)
$UserResults = [System.Collections.Generic.List[object]](Merge-DetailRows $ExistingUserRows $UserResults $scannedNames)
$PrinterResults = [System.Collections.Generic.List[object]](Merge-DetailRows $ExistingPrinterRows $PrinterResults $scannedNames)
$DeviceResults = [System.Collections.Generic.List[object]](Merge-DetailRows $ExistingDeviceRows $DeviceResults $scannedNames)
$DiagnosticsResults = [System.Collections.Generic.List[object]](
    Merge-DetailRows $ExistingDiagnosticsRows $DiagnosticsResults $scannedNames -ReplaceUnavailableWithNew
)
$ErrorResults = [System.Collections.Generic.List[object]](
    Merge-DetailRows $ExistingErrorRows $ErrorResults $scannedNames -ReplaceUnavailableWithNew
)

Write-Host "Updating Excel workbook..." -ForegroundColor Cyan

foreach ($row in @($ComputerResults)) {

    if ($null -eq $row) {
        continue
    }

    if (-not ($row.PSObject.Properties.Name -contains "Availability")) {

        $row | Add-Member `
            -NotePropertyName "Availability" `
            -NotePropertyValue "" `
            -Force
    }
}

function Convert-ExcelCellValue {
    param([object] $Value)

    if ($null -eq $Value) { return $null }

    if ($Value -is [array]) {
        return ($Value -join ", ")
    }

    if ($Value -is [System.Collections.IEnumerable] -and
        $Value -isnot [string] -and
        $Value -isnot [hashtable]) {
        return (($Value | ForEach-Object { [string]$_ }) -join ", ")
    }

    return $Value
}

# Fix for the neutral style being lost during each table update
function Get-ExcelColorSignature {
    param([object] $Color)

    if ($null -eq $Color) {
        return ""
    }

    return @(
        [string]$Color.Rgb,
        [string]$Color.Indexed,
        [string]$Color.Theme,
        [string]$Color.Tint,
        [string]$Color.Auto
    ) -join "|"
}

function Get-ExcelStyleSignature {
    param([Parameter(Mandatory=$true)] $Style)

    $parts = [System.Collections.Generic.List[string]]::new()

    # Fill
    $parts.Add("Fill.Pattern=$([string]$Style.Fill.PatternType)")
    $parts.Add("Fill.Bg=$(Get-ExcelColorSignature $Style.Fill.BackgroundColor)")
    $parts.Add("Fill.Fg=$(Get-ExcelColorSignature $Style.Fill.ForegroundColor)")

    # Font
    $parts.Add("Font.Name=$([string]$Style.Font.Name)")
    $parts.Add("Font.Size=$([string]$Style.Font.Size)")
    $parts.Add("Font.Bold=$([string]$Style.Font.Bold)")
    $parts.Add("Font.Italic=$([string]$Style.Font.Italic)")
    $parts.Add("Font.UnderLine=$([string]$Style.Font.UnderLine)")
    $parts.Add("Font.Strike=$([string]$Style.Font.Strike)")
    $parts.Add("Font.Family=$([string]$Style.Font.Family)")
    $parts.Add("Font.Charset=$([string]$Style.Font.Charset)")
    $parts.Add("Font.Scheme=$([string]$Style.Font.Scheme)")
    $parts.Add("Font.VerticalAlign=$([string]$Style.Font.VerticalAlign)")
    $parts.Add("Font.Color=$(Get-ExcelColorSignature $Style.Font.Color)")

    # Number format
    $parts.Add("NumFmt=$([string]$Style.Numberformat.Format)")

    # Alignment
    $parts.Add("Align.Horizontal=$([string]$Style.HorizontalAlignment)")
    $parts.Add("Align.Vertical=$([string]$Style.VerticalAlignment)")
    $parts.Add("Align.Wrap=$([string]$Style.WrapText)")
    $parts.Add("Align.Shrink=$([string]$Style.ShrinkToFit)")
    $parts.Add("Align.Rotate=$([string]$Style.TextRotation)")
    $parts.Add("Align.Indent=$([string]$Style.Indent)")
    $parts.Add("Align.ReadingOrder=$([string]$Style.ReadingOrder)")

    # Borders
    foreach ($sideName in @("Left", "Right", "Top", "Bottom", "Diagonal")) {
        $side = $Style.Border.$sideName
        if ($null -eq $side) {
            continue
        }

        $parts.Add("Border.$sideName.Style=$([string]$side.Style)")
        $parts.Add("Border.$sideName.Color=$(Get-ExcelColorSignature $side.Color)")
    }

    # Protection
    $parts.Add("Protection.Locked=$([string]$Style.Protection.Locked)")
    $parts.Add("Protection.Hidden=$([string]$Style.Protection.Hidden)")

    return ($parts -join ([char]31))
}

function Set-ExcelColorFromSource {
    param(
        [Parameter(Mandatory=$false)] $TargetColor,
        [Parameter(Mandatory=$false)] $SourceColor
    )

    # Some EPPlus style components (for example, the default font color)
    # may legitimately have a $null value. This is not an error,
    # so there is nothing to copy in that case.
    if ($null -eq $SourceColor -or $null -eq $TargetColor) {
        return
    }

    $rgb = [string]$SourceColor.Rgb

    if (-not [string]::IsNullOrWhiteSpace($rgb)) {
        try {
            $argb = [Convert]::ToInt64($rgb, 16)
            $a = [byte](($argb -shr 24) -band 0xFF)
            $r = [byte](($argb -shr 16) -band 0xFF)
            $g = [byte](($argb -shr 8)  -band 0xFF)
            $b = [byte]($argb -band 0xFF)
            $TargetColor.SetColor([System.Drawing.Color]::FromArgb($a, $r, $g, $b))
            return
        }
        catch {
            # Move to other color representations in EPPlus.
        }
    }

    try {
        if ($null -ne $SourceColor.Indexed -and [string]$SourceColor.Indexed -ne "") {
            $TargetColor.SetColor([int]$SourceColor.Indexed)
            return
        }
    }
    catch {}

    try {
        if ($null -ne $SourceColor.Theme -and [string]$SourceColor.Theme -ne "") {
            $TargetColor.Theme = [int]$SourceColor.Theme
            $TargetColor.Tint = $SourceColor.Tint
            return
        }
    }
    catch {}
}

function Copy-ExcelStyleProperties {
    param(
        [Parameter(Mandatory=$true)] $SourceStyle,
        [Parameter(Mandatory=$true)] $TargetStyle
    )

    # Font
    foreach ($propertyName in @(
        "Name", "Size", "Bold", "Italic", "UnderLine", "Strike",
        "Family", "Charset", "Scheme", "VerticalAlign"
    )) {
        try {
            $TargetStyle.Font.$propertyName = $SourceStyle.Font.$propertyName
        }
        catch {}
    }
    Set-ExcelColorFromSource -TargetColor $TargetStyle.Font.Color -SourceColor $SourceStyle.Font.Color

    # Fill
    try { $TargetStyle.Fill.PatternType = $SourceStyle.Fill.PatternType } catch {}
    Set-ExcelColorFromSource -TargetColor $TargetStyle.Fill.BackgroundColor -SourceColor $SourceStyle.Fill.BackgroundColor
    Set-ExcelColorFromSource -TargetColor $TargetStyle.Fill.ForegroundColor -SourceColor $SourceStyle.Fill.ForegroundColor

    # Number format
    try { $TargetStyle.Numberformat.Format = $SourceStyle.Numberformat.Format } catch {}

    # Alignment
    foreach ($propertyName in @(
        "HorizontalAlignment", "VerticalAlignment", "WrapText",
        "ShrinkToFit", "TextRotation", "Indent", "ReadingOrder"
    )) {
        try {
            $TargetStyle.$propertyName = $SourceStyle.$propertyName
        }
        catch {}
    }

    # Borders
    foreach ($sideName in @("Left", "Right", "Top", "Bottom", "Diagonal")) {
        $sourceSide = $SourceStyle.Border.$sideName
        $targetSide = $TargetStyle.Border.$sideName

        if ($null -eq $sourceSide -or $null -eq $targetSide) {
            continue
        }

        try { $targetSide.Style = $sourceSide.Style } catch {}
        Set-ExcelColorFromSource -TargetColor $targetSide.Color -SourceColor $sourceSide.Color
    }

    # Protection
    foreach ($propertyName in @("Locked", "Hidden")) {
        try {
            $TargetStyle.Protection.$propertyName = $SourceStyle.Protection.$propertyName
        }
        catch {}
    }
}

function Get-NeutralNamedStyle {
    param([Parameter(Mandatory=$true)] $ExcelPackage)

    foreach ($namedStyle in $ExcelPackage.Workbook.Styles.NamedStyles) {
        try {
            if ([string]$namedStyle.Name -eq "NEUTRAL") {
                return $namedStyle
            }
        }
        catch {}
    }

    return $null
}

function Get-Neutral2NamedStyle {
    param([Parameter(Mandatory=$true)] $ExcelPackage)

    foreach ($namedStyle in $ExcelPackage.Workbook.Styles.NamedStyles) {
        try {
            if ([string]$namedStyle.Name -eq "NEUTRAL2") {
                return $namedStyle
            }
        }
        catch {}
    }

    return $null
}

function Ensure-Neutral2NamedStyle {
    param([Parameter(Mandatory=$true)] $ExcelPackage)

# NEUTRAL2 is intentionally kept as the permanent technical style name.
    $neutral2Style = Get-Neutral2NamedStyle -ExcelPackage $ExcelPackage

    if ($null -ne $neutral2Style) {
        return $neutral2Style
    }

    # First run only: copy the current 'Neutral' style (including any changes
    # manually made by the user in Excel) into a new named EPPlus style.
    $neutralStyle = Get-NeutralNamedStyle -ExcelPackage $ExcelPackage

    if ($null -eq $neutralStyle) {
        throw "Named style NEUTRAL was not found while creating NEUTRAL2."
    }

    $neutral2Style = $ExcelPackage.Workbook.Styles.CreateNamedStyle(
        "NEUTRAL2",
        $neutralStyle.Style
    )

    Write-Log "Created native named style NEUTRAL2 from current NEUTRAL."

    return $neutral2Style
}

function Get-NormalNamedStyle {
    param([Parameter(Mandatory=$true)] $ExcelPackage)

    # EPPlus/Excel may expose the built-in 'Normal' style
    # under a localized or English name. Prefer
    # the style whose StyleXfId is 0, because this is the style used by the table by default.
    foreach ($namedStyle in $ExcelPackage.Workbook.Styles.NamedStyles) {
        try {
            if ([int]$namedStyle.StyleXfId -eq 0) {
                return $namedStyle
            }
        }
        catch {}
    }

    foreach ($name in @("Normal", "Обычный")) {
        foreach ($namedStyle in $ExcelPackage.Workbook.Styles.NamedStyles) {
            try {
                if ([string]$namedStyle.Name -eq $name) {
                    return $namedStyle
                }
            }
            catch {}
        }
    }

    return $null
}

function Write-Neutral2StyleDiagnostics {
    param([Parameter(Mandatory=$true)] $ExcelPackage)

    try {
        $styles = $ExcelPackage.Workbook.Styles

        Write-Log "STYLE DIAGNOSTICS: CellStyleXfs=$($styles.CellStyleXfs.Count) CellXfs=$($styles.CellXfs.Count) NamedStyles=$($styles.NamedStyles.Count)"

        foreach ($namedStyle in $styles.NamedStyles) {
            try {
                $style = $namedStyle.Style
                Write-Log (
                    "STYLE DIAGNOSTICS: NamedStyle='{0}' StyleXfId={1} StyleId={2} FontId={3} FillId={4} FontColor={5} FillPattern={6} FillBg={7} FillFg={8}" -f `
                    [string]$namedStyle.Name,
                    [string]$namedStyle.StyleXfId,
                    [string]$style.Id,
                    [string]$style.Font.Id,
                    [string]$style.Fill.Id,
                    [string]$style.Font.Color.Rgb,
                    [string]$style.Fill.PatternType,
                    [string]$style.Fill.BackgroundColor.Rgb,
                    [string]$style.Fill.ForegroundColor.Rgb
                )
            }
            catch {
                Write-Log "STYLE DIAGNOSTICS: failed to inspect named style: $($_.Exception.Message)"
            }
        }
    }
    catch {
        Write-Log "STYLE DIAGNOSTICS: failed: $($_.Exception.Message)"
    }
}

function Ensure-Neutral2ForNewWorkbook {
    param([Parameter(Mandatory=$true)] $ExcelPackage)

    $styles = $ExcelPackage.Workbook.Styles

    $neutralStyle = Get-NeutralNamedStyle -ExcelPackage $ExcelPackage
    $neutral2Style = Get-Neutral2NamedStyle -ExcelPackage $ExcelPackage

    if ($null -eq $neutralStyle) {
        $normalStyle = Get-NormalNamedStyle -ExcelPackage $ExcelPackage

        if ($null -eq $normalStyle) {
            throw "The workbook Normal named style was not found while creating NEUTRAL."
        }

        Write-Log "Creating native named style NEUTRAL from Normal template: Name='$([string]$normalStyle.Name)' StyleXfId=$([string]$normalStyle.StyleXfId)"

        $neutralStyle = $styles.CreateNamedStyle(
            "NEUTRAL",
            $normalStyle.Style
        )

        if ($null -eq $neutralStyle) {
            throw "EPPlus returned a null named style while creating NEUTRAL."
        }

        try {
            $neutralStyle.Style.Fill.PatternType =
                [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        }
        catch {
            Write-Log "WARNING: Failed to set NEUTRAL fill pattern: $($_.Exception.Message)"
        }

        try {
            $neutralStyle.Style.Fill.BackgroundColor.SetColor(
                [System.Drawing.Color]::FromArgb(255, 255, 235, 156)
            )
        }
        catch {
            Write-Log "WARNING: Failed to set NEUTRAL background color: $($_.Exception.Message)"
        }

        try {
            $neutralStyle.Style.Fill.ForegroundColor.SetColor(
                [System.Drawing.Color]::FromArgb(255, 255, 235, 156)
            )
        }
        catch {
            Write-Log "WARNING: Failed to set NEUTRAL foreground color: $($_.Exception.Message)"
        }

        Write-Log "Created native named style NEUTRAL from Normal."
    }
    else {
        Write-Log "New workbook already contains NEUTRAL. Reusing existing named style."
    }

    if ($null -eq $neutral2Style) {
        Write-Log "Creating native named style NEUTRAL2 from NEUTRAL template."

        $neutral2Style = $styles.CreateNamedStyle(
            "NEUTRAL2",
            $neutralStyle.Style
        )

        if ($null -eq $neutral2Style) {
            throw "EPPlus returned a null named style while creating NEUTRAL2."
        }

        try {
            $neutral2Style.Style.Font.Color.SetColor(
                [System.Drawing.Color]::FromArgb(255, 156, 101, 0)
            )
        }
        catch {
            Write-Log "WARNING: Failed to set NEUTRAL2 font color: $($_.Exception.Message)"
        }

        Write-Log "Created native named style NEUTRAL2 from NEUTRAL."
    }
    else {
        Write-Log "New workbook already contains NEUTRAL2. Reusing existing named style."
    }

    try {
        $null = $neutralStyle.Style.Id
        $null = $neutralStyle.StyleXfId
        $null = $neutral2Style.Style.Id
        $null = $neutral2Style.StyleXfId
    }
    catch {}

    Write-Neutral2StyleDiagnostics -ExcelPackage $ExcelPackage

    return $neutral2Style
}

function Convert-NeutralCellsToNeutral2 {
    param([Parameter(Mandatory=$true)] $ExcelPackage)

    $convertedCount = 0

    foreach ($worksheet in $ExcelPackage.Workbook.Worksheets) {
        if ($null -eq $worksheet.Dimension) {
            continue
        }

        $maxRow = $worksheet.Dimension.End.Row
        $maxCol = $worksheet.Dimension.End.Column

        for ($row = 1; $row -le $maxRow; $row++) {
            for ($col = 1; $col -le $maxCol; $col++) {
                $cell = $worksheet.Cells[$row, $col]

                try {
                    if ([string]$cell.StyleName -eq "NEUTRAL") {
                        $cell.StyleName = "NEUTRAL2"
                        $convertedCount++
                    }
                }
                catch {
                    Write-Log "WARNING: Failed to convert $($worksheet.Name)!$($cell.Address) from NEUTRAL to NEUTRAL2: $($_.Exception.Message)"
                }
            }
        }
    }

    Write-Log "NEUTRAL -> NEUTRAL2 conversion complete: converted=$convertedCount"
}

function Get-Neutral2CellAddresses {
    param([Parameter(Mandatory=$true)] $ExcelPackage)

    $addresses = New-Object System.Collections.Generic.List[string]

    foreach ($worksheet in $ExcelPackage.Workbook.Worksheets) {
        if ($null -eq $worksheet.Dimension) {
            continue
        }

        $maxRow = $worksheet.Dimension.End.Row
        $maxCol = $worksheet.Dimension.End.Column

        for ($row = 1; $row -le $maxRow; $row++) {
            for ($col = 1; $col -le $maxCol; $col++) {
                $cell = $worksheet.Cells[$row, $col]

                try {
                    if ([string]$cell.StyleName -eq "NEUTRAL2") {
                        $addresses.Add("$($worksheet.Name)|$($cell.Address)")
                    }
                }
                catch {}
            }
        }
    }

    return $addresses.ToArray()
}

function Restore-Neutral2CellStyles {
    param(
        [Parameter(Mandatory=$true)] $ExcelPackage,
        [Parameter(Mandatory=$true)] [string[]] $Addresses
    )

    $restoredCount = 0

    foreach ($entry in $Addresses) {
        $separator = $entry.IndexOf('|')

        if ($separator -le 0) {
            continue
        }

        $worksheetName = $entry.Substring(0, $separator)
        $address = $entry.Substring($separator + 1)
        $worksheet = $ExcelPackage.Workbook.Worksheets[$worksheetName]

        if ($null -eq $worksheet) {
            continue
        }

        try {
            $cell = $worksheet.Cells[$address]

            $cell.StyleName = "NEUTRAL2"

            $restoredCount++
        }
        catch {
            Write-Log "WARNING: Failed to restore NEUTRAL2 to ${worksheetName}!${address}: $($_.Exception.Message)"
        }
    }

    Write-Log "NEUTRAL2 style restored before save: restored=$restoredCount total=$($Addresses.Count)"
}

function Update-WorksheetWithEPPlus {
    param(
        [Parameter(Mandatory=$true)] $ExcelPackage,
        [Parameter(Mandatory=$true)] [string] $WorksheetName,
        [Parameter(Mandatory=$true)] [object[]] $Rows,
        [string] $KeyColumn = "ComputerName"
    )

    $rowsArray = @($Rows)
    if ($rowsArray.Count -eq 0) { return }

    $worksheet = $ExcelPackage.Workbook.Worksheets[$WorksheetName]
    if ($null -eq $worksheet) {
        $worksheet = $ExcelPackage.Workbook.Worksheets.Add($WorksheetName)
    }

    $headers = @($rowsArray[0].PSObject.Properties.Name)
    $headerMap = @{}

    $maxColCheck = if ($worksheet.Dimension) {
        [Math]::Min($worksheet.Dimension.End.Column, 350)
    }
    else {
        350
    }

    $lastActualCol = 0

    for ($column = 1; $column -le $maxColCheck; $column++) {
        $headerValue = $worksheet.Cells[1, $column].Value

        if (
            $null -ne $headerValue -and
            [string]::IsNullOrWhiteSpace([string]$headerValue) -eq $false
        ) {
            $headerMap[[string]$headerValue] = $column
            $lastActualCol = $column
        }
    }

    foreach ($header in $headers) {
        if (-not $headerMap.ContainsKey($header)) {
            $lastActualCol++
            $headerMap[$header] = $lastActualCol
            $worksheet.Cells[1, $lastActualCol].Value = $header
        }
    }

    $existing = @{}

    $lastRow = if ($worksheet.Dimension) {
        $worksheet.Dimension.End.Row
    }
    else {
        1
    }

    if ($headerMap.ContainsKey($KeyColumn) -and $lastRow -ge 2) {
        $keyColumnNumber = $headerMap[$KeyColumn]

        for ($rowNumber = 2; $rowNumber -le $lastRow; $rowNumber++) {
            $existingValue = $worksheet.Cells[$rowNumber, $keyColumnNumber].Value

            if (
                $null -ne $existingValue -and
                [string]::IsNullOrWhiteSpace([string]$existingValue) -eq $false
            ) {
                $existing[([string]$existingValue).Trim().ToUpperInvariant()] = $rowNumber
            }
        }
    }

    $nextRow = [Math]::Max($lastRow + 1, 2)

    foreach ($item in $rowsArray) {
        $targetRow = $null
        $key = $null

        if ($headerMap.ContainsKey($KeyColumn)) {
            $keyProperty = $item.PSObject.Properties[$KeyColumn]

            if ($null -ne $keyProperty -and $null -ne $keyProperty.Value) {
                $key = ([string]$keyProperty.Value).Trim().ToUpperInvariant()

                if ($key -and $existing.ContainsKey($key)) {
                    $targetRow = $existing[$key]
                }
            }
        }

        if ($null -eq $targetRow) {
            $targetRow = $nextRow
            $nextRow++

            if ($key) {
                $existing[$key] = $targetRow
            }
        }

        foreach ($header in $headers) {
            $property = $item.PSObject.Properties[$header]

            if ($null -eq $property) {
                continue
            }

            $cell = $worksheet.Cells[$targetRow, $headerMap[$header]]
            $cell.Value = Convert-ExcelCellValue $property.Value
        }
    }

    if ($WorksheetName -eq "Users" -and $headerMap.ContainsKey("LastUseTime")) {
        $dateColumn = $headerMap["LastUseTime"]

        if ($nextRow -gt 2) {
            $worksheet.Cells[
                2,
                $dateColumn,
                ($nextRow - 1),
                $dateColumn
            ].Style.Numberformat.Format = "dd.MM.yyyy HH:mm"
        }
    }
}

if (-not (Test-Path -LiteralPath $output)) {
    Write-Host "Workbook does not exist. Creating a new workbook..." -ForegroundColor Yellow

    $firstSheet = $true
    foreach ($exportItem in @(
        @{ Name = "Computers"; Rows = @($ComputerResults) },
        @{ Name = "Network"; Rows = @($NetworkResults) },
        @{ Name = "Users"; Rows = @($UserResults) },
        @{ Name = "Printers"; Rows = @($PrinterResults) },
        @{ Name = "Devices"; Rows = @($DeviceResults) },
        @{ Name = "Diagnostics"; Rows = @($DiagnosticsResults) },
        @{ Name = "Errors"; Rows = @($ErrorResults) }
    )) {
        if (@($exportItem.Rows).Count -eq 0) { continue }

        $exportParams = @{
            Path = $output
            WorksheetName = $exportItem.Name
            AutoSize = $true
            FreezeTopRow = $true
        }

        if (-not $firstSheet) {
            $exportParams.Append = $true
        }

        @($exportItem.Rows) | Export-Excel @exportParams
        $firstSheet = $false
    }

    $newExcelPackage = Open-ExcelPackage -Path $output
    try {
        Ensure-Neutral2ForNewWorkbook -ExcelPackage $newExcelPackage | Out-Null
        $newExcelPackage.Save()
    }
    finally {
        if ($null -ne $newExcelPackage) {
            try { $newExcelPackage.Dispose() } catch {}
        }
    }
}
else {
    $backup = "$output.bak_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    Copy-Item -LiteralPath $output -Destination $backup -Force -ErrorAction Stop
    Write-Log "BACKUP CREATED: $backup"

    $excelPackage = Open-ExcelPackage -Path $output
    try {
        Ensure-Neutral2NamedStyle -ExcelPackage $excelPackage | Out-Null

        Convert-NeutralCellsToNeutral2 -ExcelPackage $excelPackage

        $neutralCellAddresses = @(
            Get-Neutral2CellAddresses -ExcelPackage $excelPackage
        )

        Write-Log "NEUTRAL2 cells captured before worksheet updates: count=$($neutralCellAddresses.Count)"

        Update-WorksheetWithEPPlus -ExcelPackage $excelPackage -WorksheetName "Computers"   -Rows @($ComputerResults)
        Update-WorksheetWithEPPlus -ExcelPackage $excelPackage -WorksheetName "Network"     -Rows @($NetworkResults)
        Update-WorksheetWithEPPlus -ExcelPackage $excelPackage -WorksheetName "Users"       -Rows @($UserResults)
        Update-WorksheetWithEPPlus -ExcelPackage $excelPackage -WorksheetName "Printers"    -Rows @($PrinterResults)
        Update-WorksheetWithEPPlus -ExcelPackage $excelPackage -WorksheetName "Devices"     -Rows @($DeviceResults)
        Update-WorksheetWithEPPlus -ExcelPackage $excelPackage -WorksheetName "Diagnostics" -Rows @($DiagnosticsResults)
        Update-WorksheetWithEPPlus -ExcelPackage $excelPackage -WorksheetName "Errors"      -Rows @($ErrorResults)

        Restore-Neutral2CellStyles -ExcelPackage $excelPackage -Addresses $neutralCellAddresses

        $excelPackage.Save()

        Write-Log "WORKBOOK UPDATED AND SAVED SUCCESSFULLY"
    }
    catch {
        Write-Log "ERROR SAVING EXCEL PACKAGE: $($_.Exception.Message)"
        if ($null -ne $excelPackage) { $excelPackage.Dispose() }
        throw $_
    }
    finally {
        if ($null -ne $excelPackage) {
            try { $excelPackage.Dispose() } catch {}
        }
    }
}

#
# Check that the file was actually created.
#

if (-not (Test-Path -LiteralPath $output)) {

    Write-Log `
        "ERROR: Excel file was not created: $output"

    throw "Excel export failed: file was not created."
}


Write-Log `
    "RESULT Computers=$($ComputerResults.Count) Network=$($NetworkResults.Count) Users=$($UserResults.Count) Printers=$($PrinterResults.Count) Devices=$($DeviceResults.Count) Diagnostics=$($DiagnosticsResults.Count) Errors=$($ErrorResults.Count)"

Write-Log `
    "ABANDONED RUNSPACES TOTAL=$abandoned"

Write-Log `
    "ABANDONED RUNSPACES CURRENT POOL=$abandonedSincePoolRestart"

Write-Log `
    "POOL GENERATION=$poolGeneration"

Write-Log "FINISHED"

Write-Host ""
Write-Host "====================================="
Write-Host " FINISHED "
Write-Host "====================================="
Write-Host ""

Write-Host "Results:"
Write-Host "  Computers:   $($ComputerResults.Count)"
Write-Host "  Network:     $($NetworkResults.Count)"
Write-Host "  Users:       $($UserResults.Count)"
Write-Host "  Printers:    $($PrinterResults.Count)"
Write-Host "  Devices:     $($DeviceResults.Count)"
Write-Host "  Diagnostics: $($DiagnosticsResults.Count)"
Write-Host "  Errors:      $($ErrorResults.Count)"
Write-Host "  Abandoned:   $abandoned"
Write-Host ""

Write-Host "Excel:"
Write-Host $output -ForegroundColor Green

Write-Host ""

Write-Host "Log:"
Write-Host $logFile -ForegroundColor Green

Write-Log "FINISHED. Cleaning up environment..."

# 1. Force console buffer flush
[System.Console]::Out.Flush()

# 2. Run garbage collection to release remaining EPPlus handles
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

# 3. Hard exit from the .NET process, forcibly terminating ALL remaining threads
[System.Environment]::Exit(0)
