# Windows 11 VM Performance Diagnostics Script
# This script collects various system metrics to help diagnose performance issues

function Write-Header {
    param (
        [string]$Title
    )
    Write-Host "`n========== $Title ==========" -ForegroundColor Cyan
}

function Get-SystemInfo {
    Write-Header "SYSTEM INFORMATION"
    $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
    $os = Get-CimInstance -ClassName Win32_OperatingSystem
    
    Write-Host "Computer Name: $($computerSystem.Name)"
    Write-Host "Windows Version: $($os.Caption) $($os.Version)"
    Write-Host "Total Physical Memory: $([math]::Round($computerSystem.TotalPhysicalMemory / 1GB, 2)) GB"
    Write-Host "Model: $($computerSystem.Model)"
    Write-Host "Manufacturer: $($computerSystem.Manufacturer)"
    Write-Host "Boot Time: $($os.LastBootUpTime)"
    Write-Host "Uptime: $((Get-Date) - $os.LastBootUpTime)"
}

function Get-CPUUsage {
    Write-Header "CPU USAGE (Current)"
    $cpuLoad = (Get-CimInstance -ClassName Win32_Processor | Measure-Object -Property LoadPercentage -Average).Average
    Write-Host "CPU Load: $cpuLoad%"

    # Get top 5 CPU-consuming processes
    Write-Host "`nTop CPU-Consuming Processes:"
    Get-Process | Sort-Object -Property CPU -Descending | Select-Object -First 5 -Property ProcessName, CPU, Id, WorkingSet | Format-Table -AutoSize
}

function Get-MemoryUsage {
    Write-Header "MEMORY USAGE"
    $os = Get-CimInstance -ClassName Win32_OperatingSystem
    $memoryUsed = $os.TotalVisibleMemorySize - $os.FreePhysicalMemory
    $memoryPct = [math]::Round(($memoryUsed / $os.TotalVisibleMemorySize) * 100, 2)
    
    Write-Host "Total Memory: $([math]::Round($os.TotalVisibleMemorySize / 1MB, 2)) GB"
    Write-Host "Used Memory: $([math]::Round($memoryUsed / 1MB, 2)) GB ($memoryPct%)"
    Write-Host "Free Memory: $([math]::Round($os.FreePhysicalMemory / 1MB, 2)) GB"
    
    # Get top 5 memory-consuming processes
    Write-Host "`nTop Memory-Consuming Processes:"
    Get-Process | Sort-Object -Property WorkingSet -Descending | Select-Object -First 5 -Property ProcessName, Id, @{Name="Memory Usage (MB)"; Expression={[math]::Round($_.WorkingSet / 1MB, 2)}}, CPU | Format-Table -AutoSize
}

function Get-DiskInfo {
    Write-Header "DISK PERFORMANCE"
    
    # Disk space information
    Write-Host "Disk Space:"
    Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3" | 
        Select-Object -Property DeviceID, 
            @{Name="Size (GB)"; Expression={[math]::Round($_.Size / 1GB, 2)}},
            @{Name="Free Space (GB)"; Expression={[math]::Round($_.FreeSpace / 1GB, 2)}},
            @{Name="Free (%)"; Expression={[math]::Round(($_.FreeSpace / $_.Size) * 100, 2)}} |
        Format-Table -AutoSize
    
    # Disk performance counters
    try {
        $diskCounters = Get-Counter -Counter "\PhysicalDisk(*)\Disk Reads/sec", "\PhysicalDisk(*)\Disk Writes/sec", "\PhysicalDisk(*)\Current Disk Queue Length" -ErrorAction SilentlyContinue
        if ($diskCounters) {
            Write-Host "`nDisk Performance Metrics:"
            $diskCounters.CounterSamples | ForEach-Object {
                Write-Host "$($_.Path): $([math]::Round($_.CookedValue, 2))"
            }
        }
    } catch {
        Write-Host "Could not retrieve disk performance counters." -ForegroundColor Yellow
    }
}

function Get-NetworkInfo {
    Write-Header "NETWORK INFORMATION"
    
    # Get network adapter information
    Write-Host "Network Adapters:"
    Get-NetAdapter | Where-Object Status -eq "Up" | Format-Table -Property Name, InterfaceDescription, Status, LinkSpeed -AutoSize
    
    # Get network statistics
    Write-Host "`nNetwork Statistics:"
    Get-NetAdapterStatistics | Where-Object Name -In (Get-NetAdapter | Where-Object Status -eq "Up").Name | 
        Format-Table -Property Name, ReceivedBytes, SentBytes, ReceivedPackets, SentPackets -AutoSize
}

function Get-StartupPrograms {
    Write-Header "STARTUP PROGRAMS"
    
    $startupItems = @()
    
    # Registry startup items (Current User)
    $regPathCU = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
    if (Test-Path $regPathCU) {
        Get-ItemProperty -Path $regPathCU | 
            Get-Member -MemberType NoteProperty | 
            Where-Object { $_.Name -notmatch "^PS" } | 
            ForEach-Object { 
                $startupItems += [PSCustomObject]@{
                    Name = $_.Name
                    Command = (Get-ItemProperty -Path $regPathCU -Name $_.Name).$($_.Name)
                    Location = "Registry (Current User)"
                }
            }
    }
    
    # Registry startup items (Local Machine)
    $regPathLM = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Run"
    if (Test-Path $regPathLM) {
        Get-ItemProperty -Path $regPathLM | 
            Get-Member -MemberType NoteProperty | 
            Where-Object { $_.Name -notmatch "^PS" } | 
            ForEach-Object { 
                $startupItems += [PSCustomObject]@{
                    Name = $_.Name
                    Command = (Get-ItemProperty -Path $regPathLM -Name $_.Name).$($_.Name)
                    Location = "Registry (Local Machine)"
                }
            }
    }
    
    # Startup folder items
    $startupFolderCU = [System.IO.Path]::Combine($env:APPDATA, "Microsoft\Windows\Start Menu\Programs\Startup")
    if (Test-Path $startupFolderCU) {
        Get-ChildItem -Path $startupFolderCU -File | ForEach-Object {
            $startupItems += [PSCustomObject]@{
                Name = $_.BaseName
                Command = $_.FullName
                Location = "Startup Folder (Current User)"
            }
        }
    }
    
    $startupFolderAU = [System.IO.Path]::Combine($env:ProgramData, "Microsoft\Windows\Start Menu\Programs\Startup")
    if (Test-Path $startupFolderAU) {
        Get-ChildItem -Path $startupFolderAU -File | ForEach-Object {
            $startupItems += [PSCustomObject]@{
                Name = $_.BaseName
                Command = $_.FullName
                Location = "Startup Folder (All Users)"
            }
        }
    }
    
    if ($startupItems.Count -gt 0) {
        $startupItems | Format-Table -Property Name, Location, Command -AutoSize -Wrap
    } else {
        Write-Host "No startup items found."
    }
}

function Get-RunningServices {
    Write-Header "TOP 15 RESOURCE-CONSUMING SERVICES"
    
    $runningServices = Get-WmiObject -Class Win32_Service | 
        Where-Object { $_.State -eq "Running" } | 
        Select-Object Name, DisplayName, ProcessId
    
    $serviceProcesses = @()
    
    foreach ($service in $runningServices) {
        $process = Get-Process -Id $service.ProcessId -ErrorAction SilentlyContinue
        if ($process) {
            $serviceProcesses += [PSCustomObject]@{
                ServiceName = $service.Name
                DisplayName = $service.DisplayName
                PID = $service.ProcessId
                CPU = $process.CPU
                Memory = [math]::Round($process.WorkingSet64 / 1MB, 2)
            }
        }
    }
    
    $serviceProcesses | Sort-Object -Property Memory -Descending | 
        Select-Object -First 15 | 
        Format-Table -Property ServiceName, DisplayName, PID, 
            @{Name="Memory (MB)"; Expression={$_.Memory}}, 
            @{Name="CPU Time"; Expression={$_.CPU}} -AutoSize -Wrap
}

function Get-HyperVConfig {
    Write-Header "HYPER-V VM CONFIGURATION"
    
    try {
        # Check if running in a Hyper-V VM
        $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
        if ($computerSystem.Model -like "*Virtual Machine*") {
            # Try to get VM configuration if Hyper-V modules are available
            if (Get-Module -ListAvailable -Name Hyper-V) {
                $vmName = $computerSystem.Name
                Write-Host "Virtual Machine Name: $vmName"
                
                # This will work only if run with admin privileges on the host
                try {
                    $vm = Get-VM -Name $vmName -ErrorAction SilentlyContinue
                    if ($vm) {
                        Write-Host "VM Configuration:"
                        Write-Host "  CPUs: $($vm.ProcessorCount)"
                        Write-Host "  Memory Assigned: $([math]::Round($vm.MemoryAssigned / 1GB, 2)) GB"
                        Write-Host "  Memory Startup: $([math]::Round($vm.MemoryStartup / 1GB, 2)) GB"
                        Write-Host "  Dynamic Memory Enabled: $($vm.DynamicMemoryEnabled)"
                        Write-Host "  Memory Minimum: $([math]::Round($vm.MemoryMinimum / 1GB, 2)) GB"
                        Write-Host "  Memory Maximum: $([math]::Round($vm.MemoryMaximum / 1GB, 2)) GB"
                    }
                } catch {
                    Write-Host "Unable to retrieve detailed Hyper-V VM configuration. Run this script on the host with administrator privileges for more information." -ForegroundColor Yellow
                }
            } else {
                Write-Host "Running in a virtual machine, but Hyper-V PowerShell module is not available." -ForegroundColor Yellow
            }
        } else {
            Write-Host "Not running in a Hyper-V virtual machine."
        }
    } catch {
        Write-Host "Error checking Hyper-V configuration: $_" -ForegroundColor Red
    }
}

function Get-EventLogErrors {
    Write-Header "RECENT SYSTEM AND APPLICATION ERRORS (Last 24 hours)"
    
    $yesterday = (Get-Date).AddDays(-1)
    
    # System log errors
    Write-Host "System Log Errors:" -ForegroundColor Yellow
    try {
        Get-WinEvent -FilterHashtable @{LogName='System'; Level=2; StartTime=$yesterday} -MaxEvents 10 -ErrorAction SilentlyContinue |
            Select-Object TimeCreated, Id, ProviderName, Message |
            Format-Table -Wrap
    } catch {
        Write-Host "No system errors found or unable to retrieve system errors."
    }
    
    # Application log errors
    Write-Host "Application Log Errors:" -ForegroundColor Yellow
    try {
        Get-WinEvent -FilterHashtable @{LogName='Application'; Level=2; StartTime=$yesterday} -MaxEvents 10 -ErrorAction SilentlyContinue |
            Select-Object TimeCreated, Id, ProviderName, Message |
            Format-Table -Wrap
    } catch {
        Write-Host "No application errors found or unable to retrieve application errors."
    }
}

function Get-WindowsUpdates {
    Write-Header "WINDOWS UPDATES HISTORY (Last 10)"
    
    try {
        $session = New-Object -ComObject "Microsoft.Update.Session"
        $searcher = $session.CreateUpdateSearcher()
        $historyCount = $searcher.GetTotalHistoryCount()
        
        if ($historyCount -gt 0) {
            $updates = $searcher.QueryHistory(0, $historyCount) | 
                Select-Object Title, Date, Description, 
                    @{Name="Operation"; Expression={
                        switch($_.Operation) {
                            1 {"Installation"}
                            2 {"Uninstallation"}
                            3 {"Other"}
                            default {"Unknown"}
                        }
                    }},
                    @{Name="Status"; Expression={
                        switch($_.ResultCode) {
                            0 {"Not Started"}
                            1 {"In Progress"}
                            2 {"Succeeded"}
                            3 {"Succeeded With Errors"}
                            4 {"Failed"}
                            5 {"Aborted"}
                            default {"Unknown"}
                        }
                    }}
            
            $updates | Sort-Object -Property Date -Descending | Select-Object -First 10 | Format-Table -Property Title, Date, Operation, Status -AutoSize -Wrap
        } else {
            Write-Host "No Windows Update history found."
        }
    } catch {
        Write-Host "Unable to retrieve Windows Update history: $_" -ForegroundColor Red
    }
}

function Get-ReliabilityMetrics {
    Write-Header "SYSTEM RELIABILITY METRICS"
    
    try {
        # First check if the namespace exists to avoid the error
        $namespaceExists = Get-CimInstance -Namespace "root\Microsoft\Windows" -ClassName "__NAMESPACE" -ErrorAction SilentlyContinue | 
            Where-Object { $_.Name -eq "Reliability" }
        
        if ($namespaceExists) {
            $reliability = Get-CimInstance -Namespace "root\Microsoft\Windows\Reliability" -ClassName Win32_ReliabilityStabilityMetrics -ErrorAction Stop | 
                Sort-Object -Property TimeGenerated -Descending | 
                Select-Object -First 7
            
            if ($reliability) {
                $reliability | Select-Object TimeGenerated, SystemStabilityIndex | Format-Table -AutoSize
                
                # Get recent reliability events
                Write-Host "`nRecent Reliability Events:"
                $events = Get-CimInstance -Namespace "root\Microsoft\Windows\Reliability" -ClassName Win32_ReliabilityRecords | 
                    Sort-Object -Property TimeGenerated -Descending | 
                    Select-Object -First 5
                    
                foreach ($event in $events) {
                    Write-Host "Time: $($event.TimeGenerated) | Type: $($event.EventType) | Source: $($event.SourceName)"
                    Write-Host "Message: $($event.Message)`n"
                }
            } else {
                Write-Host "No reliability metrics data available on this system."
            }
        } else {
            # Alternative method - use Get-WinEvent to retrieve reliability info
            Write-Host "The Reliability WMI namespace is not available on this system. Using Event Log instead."
            
            Write-Host "`nRecent Application Crashes:"
            try {
                Get-WinEvent -FilterHashtable @{LogName='Application'; Id=1000} -MaxEvents 5 -ErrorAction SilentlyContinue |
                    Select-Object TimeCreated, Id, ProviderName, Message |
                    Format-Table -Wrap
            } catch {
                Write-Host "No application crash data found."
            }
            
            Write-Host "`nRecent Reliability Issues:"
            try {
                # Try to get reliability monitor source events
                Get-WinEvent -FilterHashtable @{LogName='System'; Id=1001} -MaxEvents 5 -ErrorAction SilentlyContinue |
                    Select-Object TimeCreated, Id, ProviderName, Message |
                    Format-Table -Wrap
            } catch {
                Write-Host "No system reliability data found."
            }
        }
    } catch {
        Write-Host "Unable to retrieve reliability metrics. Error details: $_" -ForegroundColor Red
        
        # Fallback to basic system stability check
        Write-Host "`nPerforming basic system stability checks instead..."
        
        # Check for BSOD events
        try {
            Write-Host "`nRecent System Crashes (Blue Screens):"
            Get-WinEvent -FilterHashtable @{LogName='System'; Id=1001} -MaxEvents 5 -ErrorAction SilentlyContinue |
                Where-Object { $_.Message -match "crash|dump|blue screen" } |
                Select-Object TimeCreated, Id, ProviderName, Message |
                Format-Table -Wrap
        } catch {
            Write-Host "No system crash events found."
        }
        
        # Check for application crashes
        try {
            Write-Host "`nRecent Application Crashes:"
            Get-WinEvent -FilterHashtable @{LogName='Application'; Id=1000} -MaxEvents 5 -ErrorAction SilentlyContinue |
                Select-Object TimeCreated, Id, ProviderName, Message |
                Format-Table -Wrap
        } catch {
            Write-Host "No application crash events found."
        }
    }
}

# Execute the diagnostics functions
Clear-Host
Write-Host "Windows 11 VM Performance Diagnostics Report" -ForegroundColor Green
Write-Host "Generated on: $(Get-Date)" -ForegroundColor Green
Write-Host "----------------------------------------------" -ForegroundColor Green

Get-SystemInfo
Get-CPUUsage
Get-MemoryUsage
Get-DiskInfo
Get-NetworkInfo
Get-RunningServices
Get-StartupPrograms
Get-HyperVConfig
Get-EventLogErrors
Get-WindowsUpdates
Get-ReliabilityMetrics

Write-Host "`n=== Diagnostics Complete ===" -ForegroundColor Green