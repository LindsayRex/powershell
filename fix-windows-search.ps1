# Windows Search Repair Script
# Run this as Administrator to fix Windows Search issues
# Created: May 5, 2025

# Check for admin privileges
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "This script requires administrator privileges." -ForegroundColor Red
    Write-Host "Please right-click on PowerShell and select 'Run as administrator', then run this script again." -ForegroundColor Yellow
    Write-Host "Press any key to exit..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

Write-Host "Windows Search Repair Utility" -ForegroundColor Cyan
Write-Host "Running with administrator privileges" -ForegroundColor Green
Write-Host "----------------------------------------------"

# Stop dependent services first
Write-Host "Stopping dependent services..." -ForegroundColor Yellow
$servicesToStop = @("WSearch", "WSearchIdxPi")
foreach ($service in $servicesToStop) {
    try {
        if (Get-Service -Name $service -ErrorAction SilentlyContinue) {
            Stop-Service -Name $service -Force -ErrorAction Stop
            Write-Host "Service $service stopped successfully." -ForegroundColor Green
        }
    } catch {
        Write-Host "Could not stop service $service - $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Take ownership of search folders
Write-Host "Taking ownership of Windows Search folders..." -ForegroundColor Yellow
$searchPaths = @(
    "$env:ProgramData\Microsoft\Search",
    "$env:ProgramData\Microsoft\Search\Data",
    "$env:ProgramData\Microsoft\Search\Data\Applications",
    "$env:ProgramData\Microsoft\Search\Data\Applications\Windows"
)

foreach ($path in $searchPaths) {
    if (Test-Path $path) {
        try {
            # First take ownership using takeown command
            $takeownOutput = takeown.exe /f $path /r /d y 2>&1
            if ($takeownOutput -match "SUCCESS") {
                Write-Host "Successfully took ownership of $path" -ForegroundColor Green
            } else {
                Write-Host "Partial success taking ownership of $path" -ForegroundColor Yellow
            }
            
            # Grant full control using icacls
            $icaclsOutput = icacls.exe $path /grant Administrators:F /t /c 2>&1
            if ($icaclsOutput -match "processed") {
                Write-Host "Successfully granted permissions for $path" -ForegroundColor Green
            } else {
                Write-Host "Partial success granting permissions for $path" -ForegroundColor Yellow
            }
        } catch {
            Write-Host "Error modifying permissions on $path - $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# Rename/delete Windows Search database
Write-Host "Removing Windows Search database..." -ForegroundColor Yellow
$searchDB = "$env:ProgramData\Microsoft\Search\Data\Applications\Windows\Windows.edb"
if (Test-Path $searchDB) {
    try {
        # Try renaming first (safer than deleting)
        Rename-Item -Path $searchDB -NewName "Windows.edb.old" -Force -ErrorAction Stop
        Write-Host "Successfully renamed Windows Search database file." -ForegroundColor Green
    } catch {
        Write-Host "Could not rename Windows Search database. Trying to delete..." -ForegroundColor Yellow
        try {
            Remove-Item -Path $searchDB -Force -ErrorAction Stop
            Write-Host "Successfully deleted Windows Search database file." -ForegroundColor Green
        } catch {
            Write-Host "Could not delete Windows Search database - $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "Using alternative methods..." -ForegroundColor Yellow
            
            # Try to force close handles using handle.exe (SysInternals) if available
            try {
                if (Get-Command handle -ErrorAction Stop) {
                    Write-Host "Using Handle.exe to close open file handles..." -ForegroundColor Yellow
                    $handleOutput = handle.exe $searchDB -nobanner 2>&1
                    Write-Host $handleOutput
                }
            } catch {
                Write-Host "Handle.exe not found. You can install SysInternals tools if needed." -ForegroundColor Yellow
            }
        }
    }
} else {
    Write-Host "Windows Search database file not found at expected location." -ForegroundColor Yellow
}

# Reset the search service configuration in registry
Write-Host "Resetting Windows Search service configuration..." -ForegroundColor Yellow
try {
    $registryPath = "HKLM:\SYSTEM\CurrentControlSet\Services\WSearch"
    if (Test-Path $registryPath) {
        # Set to automatic startup
        Set-ItemProperty -Path $registryPath -Name "Start" -Value 2 -Type DWord -Force
        # Set delayed start
        Set-ItemProperty -Path $registryPath -Name "DelayedAutostart" -Value 1 -Type DWord -Force
        Write-Host "Registry configuration for Windows Search reset successfully." -ForegroundColor Green
    } else {
        Write-Host "Windows Search registry path not found." -ForegroundColor Red
    }
} catch {
    Write-Host "Error modifying registry - $($_.Exception.Message)" -ForegroundColor Red
}

# Restart the Windows Search service
Write-Host "Starting Windows Search service..." -ForegroundColor Yellow
try {
    Start-Service WSearch -ErrorAction Stop
    Write-Host "Windows Search service started successfully." -ForegroundColor Green
} catch {
    Write-Host "Could not start Windows Search service - $($_.Exception.Message)" -ForegroundColor Red
    
    # Try alternative approach - using SC command
    Write-Host "Trying alternative method to start service..." -ForegroundColor Yellow
    $scOutput = sc.exe start WSearch 2>&1
    if ($scOutput -match "SUCCESS") {
        Write-Host "Windows Search service started using SC command." -ForegroundColor Green
    } else {
        Write-Host "Failed to start Windows Search with SC command." -ForegroundColor Red
        Write-Host "A system restart is required to complete the repair." -ForegroundColor Yellow
    }
}

# Try to rebuild search index through Search and Indexing troubleshooter
Write-Host "Attempting to trigger Windows Search and Indexing troubleshooter..." -ForegroundColor Yellow
try {
    # Check if the troubleshooting module is available
    $troubleshootingModuleAvailable = $false
    try {
        if (Get-Command Get-TroubleshootingPack -ErrorAction Stop) {
            $troubleshootingModuleAvailable = $true
        }
    } catch {
        $troubleshootingModuleAvailable = $false
    }
    
    if ($troubleshootingModuleAvailable) {
        # Check if the Search troubleshooter path exists
        $troubleshooterPath = "C:\Windows\diagnostics\system\Search"
        if (Test-Path $troubleshooterPath -ErrorAction SilentlyContinue) {
            try {
                $searchTroubleshooter = Get-TroubleshootingPack -Path $troubleshooterPath -ErrorAction Stop
                if ($searchTroubleshooter) {
                    Write-Host "Running Windows Search troubleshooter in unattended mode..." -ForegroundColor Yellow
                    $results = Invoke-TroubleshootingPack -Pack $searchTroubleshooter -Unattended
                    Write-Host "Troubleshooter completed with status: $($results.Status)" -ForegroundColor Green
                }
            } catch {
                Write-Host "Error running troubleshooter - $($_.Exception.Message)" -ForegroundColor Red
            }
        } else {
            Write-Host "Windows Search troubleshooter not found at expected location." -ForegroundColor Yellow
            
            # Try to find it in alternate locations
            try {
                $altPath = Get-ChildItem -Path "C:\Windows\diagnostics\" -Recurse -Filter "*Search*" -Directory -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
                if ($altPath) {
                    Write-Host "Found alternative troubleshooter path: $altPath" -ForegroundColor Green
                    try {
                        $searchTroubleshooter = Get-TroubleshootingPack -Path $altPath -ErrorAction Stop
                        if ($searchTroubleshooter) {
                            $results = Invoke-TroubleshootingPack -Pack $searchTroubleshooter -Unattended
                            Write-Host "Troubleshooter completed with status: $($results.Status)" -ForegroundColor Green
                        }
                    } catch {
                        Write-Host "Failed to run alternative troubleshooter - $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
            } catch {
                Write-Host "Error searching for alternative troubleshooter path - $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    } else {
        # Fallback to using the Windows built-in troubleshooter launcher
        Write-Host "PowerShell troubleshooting module not available. Trying alternative approach..." -ForegroundColor Yellow
        if (Test-Path "C:\Windows\System32\msdt.exe") {
            Write-Host "Launching Windows troubleshooter using msdt.exe..." -ForegroundColor Yellow
            Start-Process "msdt.exe" -ArgumentList "/id SearchDiagnostic" -Wait -NoNewWindow
            Write-Host "Windows troubleshooter completed." -ForegroundColor Green
        } else {
            Write-Host "Windows troubleshooter tools not available." -ForegroundColor Yellow
        }
    }
} catch {
    Write-Host "Could not run troubleshooter - $($_.Exception.Message)" -ForegroundColor Red
}

# Final status check
Write-Host "`nChecking Windows Search service status..." -ForegroundColor Yellow
try {
    $searchService = Get-Service WSearch
    Write-Host "Windows Search service status: $($searchService.Status)" -ForegroundColor Cyan
    Write-Host "Windows Search service startup type: $($searchService.StartType)" -ForegroundColor Cyan
    
    if ($searchService.Status -eq "Running") {
        Write-Host "Windows Search repair completed successfully!" -ForegroundColor Green
    } else {
        Write-Host "Windows Search service is not running." -ForegroundColor Red
        Write-Host "A system restart is strongly recommended to complete the repair." -ForegroundColor Yellow
    }
} catch {
    Write-Host "Could not get service status - $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`nRepair process completed. If issues persist after restarting your computer," -ForegroundColor Yellow
Write-Host "consider disabling Windows Search service if you don't rely on it." -ForegroundColor Yellow

# Prompt to restart computer
$restartChoice = Read-Host "Would you like to restart your computer now to complete repairs? (Y/N)"
if ($restartChoice -eq "Y" -or $restartChoice -eq "y") {
    Write-Host "Restarting computer in 10 seconds..." -ForegroundColor Cyan
    Start-Sleep -Seconds 10
    Restart-Computer -Force
} else {
    Write-Host "Please restart your computer manually at your earliest convenience." -ForegroundColor Yellow
    Write-Host "Press any key to exit..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}