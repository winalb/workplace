

# Configuration
$intervalSeconds = 60       # How often to simulate activity (in seconds)
$jitterPercentage = 20      # Random variation in timing (+/- %)
$mouseMovePixels = 1        # How many pixels to move the mouse
$toggleDirection = $true    # For alternating mouse movement direction

# Import required assemblies
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public class UserInputSimulator {
    [DllImport("user32.dll")]
    public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

    [DllImport("user32.dll")]
    public static extern bool SetCursorPos(int x, int y);

    [DllImport("user32.dll")]
    public static extern bool GetCursorPos(out POINT lpPoint);

    [StructLayout(LayoutKind.Sequential)]
    public struct POINT {
        public int X;
        public int Y;
    }
}
'@

function Move-MouseSlightly {
    $point = New-Object UserInputSimulator+POINT
    [UserInputSimulator]::GetCursorPos([ref]$point)

    if ($script:toggleDirection) {
        $newX = $point.X + $script:mouseMovePixels
        $newY = $point.Y + $script:mouseMovePixels
    } else {
        $newX = $point.X - $script:mouseMovePixels
        $newY = $point.Y - $script:mouseMovePixels
    }

    [UserInputSimulator]::SetCursorPos($newX, $newY)
    $script:toggleDirection = !$script:toggleDirection
}

function Send-KeyPress {
    # Simulate pressing the Shift key (virtual key code 0x10)
    [UserInputSimulator]::keybd_event(0x10, 0, 0, [UIntPtr]::Zero)
    Start-Sleep -Milliseconds 50
    [UserInputSimulator]::keybd_event(0x10, 0, 0x0002, [UIntPtr]::Zero)
}

function Get-RandomizedInterval {
    $jitter = $script:intervalSeconds * ($script:jitterPercentage / 100)
    $randomJitter = Get-Random -Minimum (-$jitter) -Maximum $jitter
    return $script:intervalSeconds + $randomJitter
}

# Main execution loop
Write-Host "Preventing screen lock - Press Ctrl+C to stop"
Write-Host "Activity interval: $intervalSeconds seconds (簣$jitterPercentage%)"

try {
    while ($true) {
        # Alternate between mouse movement and key press
        if ((Get-Date).Second % 2 -eq 0) {
            Move-MouseSlightly
        } else {
            Send-KeyPress
        }

        $waitTime = Get-RandomizedInterval
        Start-Sleep -Seconds $waitTime
    }
}
finally {
    Write-Host "`nScript stopped. Screen lock behavior returned to normal."
}
