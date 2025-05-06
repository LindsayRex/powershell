#
# AVX, AVX2, AVX-512F Hardware Intrinsics Checker (.NET Core / PowerShell 7+)
# This script checks if the current CPU and .NET runtime support these instruction sets.
#

Write-Host "Checking for AVX, AVX2, and AVX-512F instruction set support using .NET Hardware Intrinsics..." -ForegroundColor Cyan

# Show CPU model (Good for context)
try {
    $cpuInfo = Get-CimInstance Win32_Processor
    Write-Host "Detected CPU (from WMI/CIM): $($cpuInfo.Name)" -ForegroundColor Green
}
catch {
    Write-Warning "Could not retrieve CPU name using Get-CimInstance. Error: $($_.Exception.Message)"
}

function Test-AVXSupportWithIntrinsics {
    Write-Host "`n=== .NET Hardware Intrinsics Support Status ===" -ForegroundColor Cyan

    # Check hardware intrinsic support directly
    # These require .NET Core 3.0+ (PowerShell 7.0+)
    # For PowerShell 7.2+ Avx512F and other AVX-512 sets are more reliably exposed.
    $avxSupported = $false
    $avx2Supported = $false
    $avx512fSupported = $false # AVX-512 Foundation

    try {
        $avxSupported = [System.Runtime.Intrinsics.X86.Avx]::IsSupported
        $avx2Supported = [System.Runtime.Intrinsics.X86.Avx2]::IsSupported
        # Avx512F class is available in .NET 5+.
        # For .NET Core 3.1, you might need to check specific Avx512F.Base, etc.
        # but [System.Runtime.Intrinsics.X86.Avx512F]::IsSupported is the modern way for .NET 5+
        if ([System.Type]::GetType('System.Runtime.Intrinsics.X86.Avx512F')) {
            $avx512fSupported = [System.Runtime.Intrinsics.X86.Avx512F]::IsSupported
        } else {
            Write-Host "AVX-512F: .NET class not available (requires .NET 5+ for this specific check format)." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Warning "Error accessing .NET Intrinsics. This script requires PowerShell 7+ (ideally 7.2+ for best AVX-512 reporting) built on .NET Core 3.0+ / .NET 5+."
        Write-Warning "Error details: $($_.Exception.Message)"
        # Return early or with all false if intrinsics are not accessible
        return [PSCustomObject]@{
            AVX    = $false
            AVX2   = $false
            AVX512F = $false # Changed key name for clarity
            Error  = "Failed to access .NET Intrinsics. Ensure PowerShell 7+ / .NET Core 3.0+."
        }
    }

    # AVX
    Write-Host "AVX (Version 1) : " -NoNewline
    if ($avxSupported) {
        $exampleSumAvx = 1.5 + 2.5 # Example floating point
        Write-Host "SUPPORTED ✓ (Example: 1.5 + 2.5 = $exampleSumAvx)" -ForegroundColor Green
    } else {
        Write-Host "NOT SUPPORTED ✗" -ForegroundColor Red
    }

    # AVX2
    Write-Host "AVX2            : " -NoNewline
    if ($avx2Supported) {
        $exampleSumAvx2 = 2000000000 + 3000000000 # Example integer
        Write-Host "SUPPORTED ✓ (Example: 2B + 3B = $exampleSumAvx2)" -ForegroundColor Green
    } else {
        Write-Host "NOT SUPPORTED ✗" -ForegroundColor Red
    }

    # AVX-512F
    Write-Host "AVX-512F        : " -NoNewline # AVX-512 Foundation
    if ($avx512fSupported) {
        $exampleSumAvx512f = 3.14159 * 2.71828 # Example more complex float
        Write-Host "SUPPORTED ✓ (Example: PI * E approx = $($exampleSumAvx512f.ToString('F5')))" -ForegroundColor Green
    } else {
        Write-Host "NOT SUPPORTED ✗" -ForegroundColor Red
    }

    return [PSCustomObject]@{
        AVX    = $avxSupported
        AVX2   = $avx2Supported
        AVX512F = $avx512fSupported # Changed key name
    }
}

# Run and display summary
$results = Test-AVXSupportWithIntrinsics

Write-Host "`n--- Summary ---"
if ($results.Error) {
    Write-Host "An error occurred: $($results.Error)" -ForegroundColor Red
} else {
    Write-Host "AVX      : $($results.AVX)"
    Write-Host "AVX2     : $($results.AVX2)"
    Write-Host "AVX-512F : $($results.AVX512F)"
}

Write-Host "`nNote: This method relies on the .NET runtime's detection capabilities." -ForegroundColor Yellow
Write-Host "For this to be accurate in a VM, ensure your hypervisor (e.g., Proxmox) passes through"
Write-Host "CPU features correctly (e.g., CPU type 'host') and that the guest OS is compatible."