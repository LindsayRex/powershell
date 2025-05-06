# Windows 11 VM Performance Improvement Script
# Based on diagnostics report from May 5, 2025
# This script provides fixes for common VM performance issues

function Write-Header {
    param (
        [string]$Title
    )
    Write-Host "`n========== $Title ==========" -ForegroundColor Cyan
}

# Fix Windows Search issues
function Repair-WindowsSearch {
    Write-Header "REPAIRING WINDOWS SEARCH"
    
    Write-Host "Windows Search service is showing errors and recovery failures..." -ForegroundColor Yellow
    
    # Stop the service
    Write-Host "Stopping Windows Search service..."
    Stop-Service WSearch -Force -ErrorAction SilentlyContinue
    
    # Delete the search index - fix permissions issue by using Test-Path with -PathType Leaf
    $searchDataPath = "$env:ProgramData\Microsoft\Search\Data\Applications\Windows"
    $searchFile = "$searchDataPath\Windows.edb"
    
    # Check if we're running as admin
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
    
    if (-not $isAdmin) {
        Write-Host "Administrator privileges required to reset Windows Search index files." -ForegroundColor Red
        Write-Host "Please run this script as administrator to fully reset Windows Search." -ForegroundColor Red
    } else {
        # Check for files with proper error handling
        if (Test-Path $searchDataPath -ErrorAction SilentlyContinue) {
            Write-Host "Deleting Windows Search index files..."
            try {
                if (Test-Path $searchFile -PathType Leaf -ErrorAction SilentlyContinue) {
                    Remove-Item -Path $searchFile -Force -ErrorAction Stop
                }
                Get-ChildItem -Path $searchDataPath -Filter "Windows.*.log" -ErrorAction SilentlyContinue | 
                    ForEach-Object { Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue }
                Write-Host "Search index files deleted successfully." -ForegroundColor Green
            } catch {
                Write-Host "Unable to delete search index files: $($_.Exception.Message)" -ForegroundColor Red
                Write-Host "This may be due to the Windows Search service still having open handles to these files." -ForegroundColor Yellow
            }
        } else {
            Write-Host "Windows Search index path not found. The service may be using a different location." -ForegroundColor Yellow
        }
    }
    
    # Reset the service to delayed start - fix incorrect startup type
    Write-Host "Setting Windows Search to delayed start..."
    try {
        # Use registry method instead since Set-Service doesn't support delayed auto
        $path = "HKLM:\SYSTEM\CurrentControlSet\Services\WSearch"
        if (Test-Path $path) {
            Set-ItemProperty -Path $path -Name "Start" -Value 2 -ErrorAction Stop # 2 = Automatic
            Set-ItemProperty -Path $path -Name "DelayedAutostart" -Value 1 -ErrorAction Stop # 1 = Delayed
            Write-Host "Windows Search service set to delayed automatic start." -ForegroundColor Green
        } else {
            Write-Host "Windows Search registry path not found." -ForegroundColor Red
        }
    } catch {
        Write-Host "Unable to set Windows Search to delayed start: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Falling back to regular automatic start..." -ForegroundColor Yellow
        Set-Service WSearch -StartupType Automatic -ErrorAction SilentlyContinue
    }
    
    # Start the service
    Write-Host "Starting Windows Search service..."
    try {
        Start-Service WSearch -ErrorAction Stop
        Write-Host "Windows Search service started successfully." -ForegroundColor Green
    } catch {
        Write-Host "Unable to start Windows Search service: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Try restarting the computer for changes to take effect." -ForegroundColor Yellow
    }
    
    Write-Host "Windows Search repair procedure completed." -ForegroundColor Green
}

# Optimize virtual memory settings
function Optimize-VirtualMemory {
    Write-Header "OPTIMIZING VIRTUAL MEMORY"
    
    # Get system info
    $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
    $physicalMemory = [Math]::Round($computerSystem.TotalPhysicalMemory / 1GB, 0)
    
    # Calculate optimal pagefile size (min = 4GB, max = RAM size)
    $pageFileMin = 4096
    $pageFileMax = $physicalMemory * 1024
    
    Write-Host "Setting pagefile to Min: $pageFileMin MB, Max: $pageFileMax MB..."
    
    # Set pagefile settings
    $computersys = Get-WmiObject Win32_ComputerSystem -EnableAllPrivileges
    $computersys.AutomaticManagedPagefile = $false
    $computersys.Put() | Out-Null
    
    $pagefile = Get-WmiObject -Query "SELECT * FROM Win32_PageFileSetting WHERE Name='C:\\pagefile.sys'"
    if ($pagefile) {
        $pagefile.InitialSize = $pageFileMin
        $pagefile.MaximumSize = $pageFileMax
        $pagefile.Put() | Out-Null
    } else {
        $pagefileset = New-Object -TypeName System.Management.ManagementClass Win32_PageFileSetting
        $pagefileset.Name = "C:\pagefile.sys"
        $pagefileset.InitialSize = $pageFileMin
        $pagefileset.MaximumSize = $pageFileMax
        $pagefileset.Put() | Out-Null
    }
    
    Write-Host "Virtual memory settings optimized for VM performance." -ForegroundColor Green
    Write-Host "NOTE: A system restart will be required for these changes to take effect."
}

# Optimize High CPU usage from GoogleDriveFS
function Optimize-GoogleDriveFS {
    Write-Header "OPTIMIZING GOOGLE DRIVE FILE STREAM"
    
    Write-Host "GoogleDriveFS is using high CPU resources (1275+ CPU units)" -ForegroundColor Yellow
    
    # Stop GoogleDriveFS
    Get-Process -Name "GoogleDriveFS" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    
    # Find and modify the preferences file
    $prefsFile = "$env:LOCALAPPDATA\Google\DriveFS\Settings\Preferences"
    
    if (Test-Path $prefsFile) {
        Write-Host "Modifying Google Drive preferences to optimize performance..."
        
        # Backup the preferences file
        Copy-Item $prefsFile "$prefsFile.backup" -Force -ErrorAction SilentlyContinue
        
        # Read the file content
        $content = Get-Content $prefsFile -Raw -ErrorAction SilentlyContinue
        
        if ($content) {
            # Set performance optimizations for Drive File Stream
            $content = $content -replace '"bandwidthLimitKbps":\s*\d+', '"bandwidthLimitKbps": 1024'
            $content = $content -replace '"cacheMaxPercentOfDisk":\s*\d+', '"cacheMaxPercentOfDisk": 5'
            $content = $content -replace '"maxCacheSizeInGb":\s*\d+', '"maxCacheSizeInGb": 10'
            
            # Write back the modified content
            $content | Set-Content $prefsFile -Force -ErrorAction SilentlyContinue
            
            Write-Host "Google Drive File Stream settings optimized." -ForegroundColor Green
        } else {
            Write-Host "Unable to read Google Drive preferences file." -ForegroundColor Red
        }
    } else {
        Write-Host "Google Drive preferences file not found. Skipping optimization." -ForegroundColor Yellow
    }
    
    # Provide recommendation for startup
    Write-Host "Recommendation: Consider removing GoogleDriveFS from startup items or using it only when needed." -ForegroundColor Yellow
}

# Manage startup programs
function Manage-StartupPrograms {
    Write-Header "MANAGING STARTUP PROGRAMS"
    
    Write-Host "Your system has 20+ startup programs which can impact boot time and performance."
    Write-Host "Here are the non-essential startup items you might consider disabling:"
    
    $nonEssential = @(
        "GoogleDriveFS",
        "iZotopeProductPortal",
        "MicrosoftEdgeAutoLaunch",
        "NVIDIA Broadcast",
        "ProtonVPN",
        "SoundID Reference.exe",
        "Spotify",
        "Swiftpoint X1 Control Panel",
        "UA Connect",
        "Wargaming.net Game Center",
        "Corsair iCUE5 Software",
        "Overbridge Engine",
        "Pieces for Developers",
        "Pieces OS",
        "Roland Cloud Manager",
        "Mackie USB Driver Control Panel Autostart",
        "WavesLocalServer"
    )
    
    foreach ($item in $nonEssential) {
        Write-Host "  - $item" -ForegroundColor Yellow
    }
    
    Write-Host "`nTo disable these items:" -ForegroundColor White
    Write-Host "1. Press Win+R, type 'msconfig' and press Enter" -ForegroundColor White
    Write-Host "2. Go to the 'Startup' tab and click 'Open Task Manager'" -ForegroundColor White
    Write-Host "3. Disable non-essential startup items" -ForegroundColor White
    Write-Host "OR" -ForegroundColor White
    Write-Host "Run the following command to open startup items directly:" -ForegroundColor White
    Write-Host "  taskmgr.exe /0 /startup" -ForegroundColor White
}

# Optimize system services
function Optimize-SystemServices {
    Write-Header "OPTIMIZING SYSTEM SERVICES"
    
    Write-Host "Some system services can be optimized for better performance in a VM."
    
    # Superfetch/SysMain - often not needed in VMs
    Write-Host "Adjusting SysMain service (Superfetch)..."
    Set-Service SysMain -StartupType Disabled
    Stop-Service SysMain -Force -ErrorAction SilentlyContinue
    
    # Windows Search - if not needed
    Write-Host "Windows Search can be disabled if not needed. Not disabling automatically."
    # Set-Service WSearch -StartupType Disabled
    # Stop-Service WSearch -Force -ErrorAction SilentlyContinue
    
    Write-Host "Memory compression is useful for VMs - ensuring it's enabled..."
    Enable-MMAgent -MemoryCompression
    
    Write-Host "Services optimized for VM performance." -ForegroundColor Green
}

# Optimize VM-specific settings
function Optimize-VMSettings {
    Write-Header "OPTIMIZING VM-SPECIFIC SETTINGS"
    
    # Disable visual effects
    Write-Host "Optimizing for performance by reducing visual effects..."
    
    # Set performance options to "Adjust for best performance"
    $path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects"
    if (-not (Test-Path $path)) {
        New-Item -Path $path -Force | Out-Null
    }
    Set-ItemProperty -Path $path -Name "VisualFXSetting" -Value 2
    
    # Disable transparency
    $path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
    if (-not (Test-Path $path)) {
        New-Item -Path $path -Force | Out-Null
    }
    Set-ItemProperty -Path $path -Name "EnableTransparency" -Value 0
    
    # Disable lock screen
    $path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization"
    if (-not (Test-Path $path)) {
        New-Item -Path $path -Force | Out-Null
    }
    Set-ItemProperty -Path $path -Name "NoLockScreen" -Value 1
    
    # Disable Game DVR
    $path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\GameDVR"
    if (-not (Test-Path $path)) {
        New-Item -Path $path -Force | Out-Null
    }
    Set-ItemProperty -Path $path -Name "AllowGameDVR" -Value 0
    
    # Optimize power settings for performance
    Write-Host "Setting power plan to high performance..."
    powercfg.exe /setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c
    
    Write-Host "VM-specific performance settings optimized." -ForegroundColor Green
}

# Provide recommendations
function Show-Recommendations {
    Write-Header "PERFORMANCE RECOMMENDATIONS"
    
    Write-Host "Based on the diagnostic analysis, here are personalized recommendations:" -ForegroundColor Green
    
    Write-Host "1. Address GoogleDriveFS high resource usage:"
    Write-Host "   - Consider using the web version instead of the desktop app"
    Write-Host "   - Limit the number of files synced using selective sync"
    Write-Host "   - Update to the latest version (current: 107.0.3.0)"
    
    Write-Host "`n2. Reduce Visual Studio Code instances:"
    Write-Host "   - Multiple VS Code windows are using significant resources"
    Write-Host "   - Consider using workspace files to group projects in one window"
    
    Write-Host "`n3. Recent memory issues detected:"
    Write-Host "   - Memory diagnostic reported bad memory regions on March 8"
    Write-Host "   - A system crash (BSOD) occurred on March 8"
    Write-Host "   - Consider increasing VM memory allocation if possible"
    
    Write-Host "`n4. Windows Search is failing to recover:"
    Write-Host "   - Run the repair function in this script to reset"
    Write-Host "   - Or disable the service if you don't need search functionality"
    
    Write-Host "`n5. VirtualBox/QEMU optimization:"
    Write-Host "   - Check virtual machine settings and ensure proper resource allocation"
    Write-Host "   - Consider enabling virtualization extensions if supported"
    Write-Host "   - Ensure VM additions/tools are installed and up to date"
    
    Write-Host "`n6. Consider reducing startup items:"
    Write-Host "   - 20+ startup programs will slow down boot and consume resources"
    Write-Host "   - Disable non-essential startup items"
}

# Main execution
Clear-Host
Write-Host "Windows 11 VM Performance Improvement Tool" -ForegroundColor Green
Write-Host "Based on diagnostic report from May 5, 2025" -ForegroundColor Green
Write-Host "----------------------------------------------" -ForegroundColor Green

# Show menu
function Show-Menu {
    Write-Host "`nSelect an optimization to perform:" -ForegroundColor Cyan
    Write-Host "1: Repair Windows Search issues"
    Write-Host "2: Optimize virtual memory settings"
    Write-Host "3: Address GoogleDriveFS high CPU usage"
    Write-Host "4: Manage startup programs"
    Write-Host "5: Optimize system services"
    Write-Host "6: Apply VM-specific optimizations"
    Write-Host "7: Show performance recommendations"
    Write-Host "8: Run all optimizations (except startup management)"
    Write-Host "Q: Quit"
    
    $selection = Read-Host "`nEnter selection"
    
    switch ($selection) {
        "1" { Repair-WindowsSearch; pause; Show-Menu }
        "2" { Optimize-VirtualMemory; pause; Show-Menu }
        "3" { Optimize-GoogleDriveFS; pause; Show-Menu }
        "4" { Manage-StartupPrograms; pause; Show-Menu }
        "5" { Optimize-SystemServices; pause; Show-Menu }
        "6" { Optimize-VMSettings; pause; Show-Menu }
        "7" { Show-Recommendations; pause; Show-Menu }
        "8" {
            Repair-WindowsSearch
            Optimize-VirtualMemory
            Optimize-GoogleDriveFS
            Optimize-SystemServices
            Optimize-VMSettings
            Show-Recommendations
            Write-Host "`nAll optimizations completed!" -ForegroundColor Green
            pause
            Show-Menu
        }
        "Q" { return }
        "q" { return }
        default { 
            Write-Host "Invalid selection, please try again." -ForegroundColor Red
            Show-Menu
        }
    }
}

# Display initial recommendations
Show-Recommendations
Write-Host "`nThis script provides multiple optimization options to address performance issues." -ForegroundColor Yellow
Write-Host "IMPORTANT: Some changes require administrator privileges and may require a system restart." -ForegroundColor Yellow

# Show the menu
Show-Menu