# Windows 11 VM Performance Improvement Script
# Based on diagnostics report from May 5, 2025
# This script provides fixes for common VM performance issues

# Self-elevation mechanism
function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Restart script with admin rights if not already elevated
function Start-AdminSession {
    if (-not (Test-Admin)) {
        Write-Host "This script requires administrator privileges." -ForegroundColor Yellow
        Write-Host "Attempting to restart as administrator..." -ForegroundColor Yellow
        
        # Get the current script path
        $scriptPath = $MyInvocation.MyCommand.Definition
        
        # Start a new PowerShell process with elevated permissions
        Start-Process PowerShell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" -Verb RunAs
        
        # Exit the current non-elevated session
        exit
    }
}

# Call the self-elevation function at the start
Start-AdminSession

# Rest of the script continues below...
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
    try {
        Stop-Service WSearch -Force -ErrorAction Stop
        Write-Host "Windows Search service stopped successfully." -ForegroundColor Green
    } catch {
        Write-Host "Unable to stop Windows Search service: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "Continuing with repair procedure..." -ForegroundColor Yellow
    }
    
    # Delete the search index
    $searchDataPath = "$env:ProgramData\Microsoft\Search\Data\Applications\Windows"
    $searchFile = "$searchDataPath\Windows.edb"
    
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
            
            # Try to forcefully take ownership and set permissions
            try {
                Write-Host "Attempting to take ownership of search index files..." -ForegroundColor Yellow
                $acl = Get-Acl -Path $searchDataPath
                $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
                $fileSystemRights = [System.Security.AccessControl.FileSystemRights]::FullControl
                $type = [System.Security.AccessControl.AccessControlType]::Allow
                $fileSystemAccessRule = New-Object -TypeName System.Security.AccessControl.FileSystemAccessRule($identity, $fileSystemRights, $type)
                $acl.AddAccessRule($fileSystemAccessRule)
                Set-Acl -Path $searchDataPath -AclObject $acl -ErrorAction Stop
                
                # Try again to delete files
                if (Test-Path $searchFile -PathType Leaf -ErrorAction SilentlyContinue) {
                    Remove-Item -Path $searchFile -Force -ErrorAction SilentlyContinue
                }
            } catch {
                Write-Host "Unable to take ownership of search index files: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    } else {
        Write-Host "Windows Search index path not found. The service may be using a different location." -ForegroundColor Yellow
    }
    
    # Reset the service to delayed start
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
            Write-Host "Falling back to regular automatic start..." -ForegroundColor Yellow
            Set-Service WSearch -StartupType Automatic -ErrorAction SilentlyContinue
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
        
        # Try to rebuild search index through alternative method
        Write-Host "Attempting alternate Windows Search repair method..." -ForegroundColor Yellow
        try {
            # Use WMI to repair Windows Search
            $searchManager = New-Object -ComObject Microsoft.Search.Administration
            $catalogManager = $searchManager.GetCatalog("SystemIndex")
            $catalogManager.Reset()
            Write-Host "Windows Search index reset through Search Administration API." -ForegroundColor Green
            
            # Try to start the service again
            Start-Service WSearch -ErrorAction SilentlyContinue
        } catch {
            Write-Host "Alternative repair method failed: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "Try restarting the computer for changes to take effect." -ForegroundColor Yellow
        }
    }
    
    Write-Host "Windows Search repair procedure completed." -ForegroundColor Green
    Write-Host "NOTE: A system restart is recommended for changes to take full effect." -ForegroundColor Yellow
}

# Rest of the functions remain unchanged

# Main execution
Clear-Host
Write-Host "Windows 11 VM Performance Improvement Tool" -ForegroundColor Green
Write-Host "Based on diagnostic report from May 5, 2025" -ForegroundColor Green
Write-Host "Running with" -NoNewline
if (Test-Admin) {
    Write-Host " ADMINISTRATOR" -ForegroundColor Green -NoNewline
} else {
    Write-Host " STANDARD USER" -ForegroundColor Red -NoNewline
}
Write-Host " privileges"
Write-Host "----------------------------------------------" -ForegroundColor Green