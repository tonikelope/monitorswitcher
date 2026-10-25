param(
    # Show a toast notification each time the mode changes. Off by default.
    [switch]$ShowNotifications
)

# Load notification module only when notifications are enabled.
if ($ShowNotifications) {
    Import-Module BurntToast
}

$global:showNotifications = [bool]$ShowNotifications
$global:icon = Join-Path $PSScriptRoot "MonitorDual.png"

$global:keyboardId = "VID_1532&PID_02A2" # Razer Ornata V3 X; verify with Get-PnpDevice -Class Keyboard
$global:monitorId = "\\.\DISPLAY1\Monitor0" # Get this from ControlMyMonitor
$global:monitor2Id = "\\.\DISPLAY2\Monitor0" # Get this from ControlMyMonitor
$global:modeA_monitor_output_id = 15 #HDMI, DP, ETC... Get this from ControlMyMonitor
$global:modeA_monitor2_output_id = 15 #HDMI, DP, ETC... Get this from ControlMyMonitor
$global:modeB_monitor_output_id = 17 #HDMI, DP, ETC... Get this from ControlMyMonitor
$global:modeB_monitor2_output_id = 17 #HDMI, DP, ETC... Get this from ControlMyMonitor


$global:modeA_msg = "MODE A ENABLED"
$global:modeB_msg = "MODE B ENABLED"
$global:controlMyMonitor = Join-Path $PSScriptRoot "ControlMyMonitor.exe"
$global:keyboardConnected = $false
$global:logName = "Application"
$global:source = "MonitorSwitcher"

# Reads Windows' "turn off display" timeout (VIDEOIDLE) from the active power plan,
# so the wake-up detection adapts automatically to whatever the system is set to.
# Picks the AC or DC value depending on how the machine is currently powered.
# Falls back to 900s if it can't be read or the display is set to never turn off.
function GetDisplaySleepThresholdSec {
    param([int]$Fallback = 900)
    try {
        $out = powercfg /q SCHEME_CURRENT SUB_VIDEO VIDEOIDLE 2>$null
        # The output lists several hex values (min, max, increment, AC index, DC
        # index). The two we want are the last ones: AC (second to last) and DC.
        $hex = [regex]::Matches(($out -join "`n"), '0x[0-9a-fA-F]+') | ForEach-Object { $_.Value }
        if ($hex.Count -lt 2) { return $Fallback }

        $onBattery = $false
        try {
            $bat = Get-CimInstance -ClassName Win32_Battery -ErrorAction SilentlyContinue
            if ($bat) { $onBattery = ($bat.BatteryStatus -ne 2) } # 2 = plugged in
        } catch {}

        $sec = if ($onBattery) { [Convert]::ToInt32($hex[-1], 16) } else { [Convert]::ToInt32($hex[-2], 16) }
        if ($sec -le 0) { return $Fallback } # 0 = "Never", so keep the fallback
        return $sec
    } catch {
        return $Fallback
    }
}

# When the screen sleeps by inactivity the monitor loses signal and jumps to the
# other input on its own. When the user comes back after being idle at least this
# long, the script reapplies the current mode so the monitor switches back to the
# right input. Read from the system so it matches the real display-off timeout.
$global:displaySleepThresholdSec = GetDisplaySleepThresholdSec
# How often (seconds) the loop wakes up to poll for a screen wake-up.
$global:pollIntervalSec = 2

# P/Invoke to read how long the user has been idle (no keyboard/mouse input).
Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class IdleInfo {
    [StructLayout(LayoutKind.Sequential)]
    struct LASTINPUTINFO {
        public uint cbSize;
        public uint dwTime;
    }
    [DllImport("user32.dll")]
    static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);
    public static uint GetIdleSeconds() {
        LASTINPUTINFO lii = new LASTINPUTINFO();
        lii.cbSize = (uint)Marshal.SizeOf(lii);
        if (!GetLastInputInfo(ref lii)) { return 0; }
        return (uint)(((uint)Environment.TickCount - lii.dwTime) / 1000);
    }
}
"@

try {
    # Create event source if it doesn't exist
	if (-not (Get-EventLog -LogName $global:logName | Where-Object { $_.Source -eq $global:source })) {
		New-EventLog -LogName $global:logName -Source $global:source
	}
} catch {
	Write-Host "Error creating event source"
}

# Read the monitor's current input (VCP feature 60). ControlMyMonitor is a GUI
# app, so Start-Process -Wait is required to get its exit code, which carries the
# value. Returns the exit code, or $null on error.
function GetMonitorInput {
    param([string]$MonitorId)
    try {
        $p = Start-Process -FilePath $global:controlMyMonitor -ArgumentList '/GetValue', $MonitorId, '60' -Wait -PassThru -WindowStyle Hidden
        return $p.ExitCode
    } catch {
        Write-Host "Error reading monitor input: $_"
        return $null
    }
}

# Apply the target inputs to both monitors, fire-and-forget for maximum speed.
#
# -SkipIfCurrent adds the safety control (read each input first and skip the ones
# already on target, to avoid a redundant switch / flicker). Reading is slow
# (~1s per monitor to launch ControlMyMonitor), so it's used ONLY where the
# flicker matters and latency doesn't: the initial setup (the monitor may already
# be on the target input) and the wake reapply (not interactive, monitor is
# powering on anyway). A real interactive KVM switch always changes the input
# (mode A <-> B), so there's nothing to skip and we write straight away, snappy.
# Reads happen BEFORE any write, so a read never overlaps a write on the bus.
function ApplyMode {
    param([int]$Target1, [int]$Target2, [switch]$SkipIfCurrent)
    if ($SkipIfCurrent) {
        $need1 = (GetMonitorInput $global:monitorId)  -ne $Target1
        $need2 = (GetMonitorInput $global:monitor2Id) -ne $Target2
    } else {
        $need1 = $true
        $need2 = $true
    }
    if ($need1) { & $global:controlMyMonitor /SetValue $global:monitorId 60 $Target1 }
    if ($need2) { & $global:controlMyMonitor /SetValue $global:monitor2Id 60 $Target2 }
}

function SwitchToModeA {
    param([switch]$Silent, [switch]$SkipIfCurrent)
    try {
        if ($global:showNotifications -and -not $Silent) {
            New-BurntToastNotification -AppLogo $global:icon -Text "MonitorSwitcher (by tonikelope)", $global:modeA_msg
        }
        ApplyMode $global:modeA_monitor_output_id $global:modeA_monitor2_output_id -SkipIfCurrent:$SkipIfCurrent
    } catch {
        Write-Host "Error switching to Mode A: $_"
        Write-EventLog -LogName $global:logName -Source $global:source -EntryType Error -EventId 1001 -Message "Error switching to Mode A: $_"
    }
}

function SwitchToModeB {
    param([switch]$Silent, [switch]$SkipIfCurrent)
    try {
        if ($global:showNotifications -and -not $Silent) {
            New-BurntToastNotification -AppLogo $global:icon -Text "MonitorSwitcher (by tonikelope)", $global:modeB_msg
        }
        ApplyMode $global:modeB_monitor_output_id $global:modeB_monitor2_output_id -SkipIfCurrent:$SkipIfCurrent
    } catch {
        Write-Host "Error switching to Mode B: $_"
        Write-EventLog -LogName $global:logName -Source $global:source -EntryType Error -EventId 1002 -Message "Error switching to Mode B: $_"
    }
}

function IsConnected {
    try {
        $dev = @(Get-PnpDevice -Class Keyboard | Where-Object {
            $_.InstanceId -like "*$global:keyboardId*" -and $_.Status -eq "OK"
        })
        return $dev.Count -gt 0
    } catch {
        Write-Host "Error checking if keyboard is connected: $_"
        Write-EventLog -LogName $global:logName -Source $global:source -EntryType Error -EventId 1003 -Message "Error checking if keyboard is connected: $_"
        return $false
    }
}

# Initial cleanup
try {
    Get-EventSubscriber | Where-Object { $_.SourceIdentifier -match "KeyboardEvent" } | Unregister-Event -ErrorAction SilentlyContinue
} catch {
    Write-Host "Error cleaning up previous events: $_"
    Write-EventLog -LogName $global:logName -Source $global:source -EntryType Error -EventId 1005 -Message "Error cleaning up previous events: $_"
}

# Initial setup
try {
    $global:keyboardConnected = IsConnected
    if ($global:keyboardConnected) {
        SwitchToModeA -SkipIfCurrent
    } else {
        SwitchToModeB -SkipIfCurrent
    }
} catch {
    Write-Host "Error setting up initial state: $_"
    Write-EventLog -LogName $global:logName -Source $global:source -EntryType Error -EventId 1006 -Message "Error setting up initial state: $_"
}

# Use event only as a signal
try {
    Register-WmiEvent -Class Win32_DeviceChangeEvent -SourceIdentifier "KeyboardEvent"
} catch {
    Write-Host "Error registering WMI event: $_"
    Write-EventLog -LogName $global:logName -Source $global:source -EntryType Error -EventId 1007 -Message "Error registering WMI event: $_"
    exit 1
}

Write-Host "Keyboard status monitor started. Waiting for events..."

$global:previousIdleSec = [IdleInfo]::GetIdleSeconds()

try {
    while ($true) {
        # Poll with a timeout so we can also detect display wake-ups (which don't
        # raise any device event). A keyboard event just makes this return sooner.
        try {
            Wait-Event -SourceIdentifier "KeyboardEvent" -Timeout $global:pollIntervalSec | Out-Null
        } catch {
            Write-Host "Error waiting for event: $_"
            Write-EventLog -LogName $global:logName -Source $global:source -EntryType Error -EventId 1008 -Message "Error waiting for event: $_"
            break
        }

        # Remove event to prevent accumulation (no-op on timeout).
        try {
            Remove-Event -SourceIdentifier "KeyboardEvent" -ErrorAction SilentlyContinue
        } catch {
            Write-Host "Error removing event: $_"
            Write-EventLog -LogName $global:logName -Source $global:source -EntryType Error -EventId 1009 -Message "Error removing event: $_"
        }

        # Detect a display wake-up: the user was idle long enough for the screen to
        # sleep (so the monitor jumped to the other input on its own) and is now
        # active again. Reapply the current mode silently to force the right input.
        try {
            $currentIdleSec = [IdleInfo]::GetIdleSeconds()
            if ($global:previousIdleSec -ge $global:displaySleepThresholdSec -and $currentIdleSec -lt $global:previousIdleSec) {
                Write-Host "DISPLAY WAKE DETECTED - reapplying current mode"
                if ($global:keyboardConnected) {
                    SwitchToModeA -Silent -SkipIfCurrent
                } else {
                    SwitchToModeB -Silent -SkipIfCurrent
                }
            }
            $global:previousIdleSec = $currentIdleSec
        } catch {
            Write-Host "Error handling display wake: $_"
            Write-EventLog -LogName $global:logName -Source $global:source -EntryType Error -EventId 1013 -Message "Error handling display wake: $_"
        }

        try {
            $newState = IsConnected
            if ($newState -ne $global:keyboardConnected) {
                Write-Host "STATE CHANGE DETECTED"
                $global:keyboardConnected = $newState
                if ($newState) {
                    SwitchToModeA
                } else {
                    SwitchToModeB
                }
            }
        } catch {
            Write-Host "Error checking or switching keyboard state: $_"
            Write-EventLog -LogName $global:logName -Source $global:source -EntryType Error -EventId 1010 -Message "Error checking or switching keyboard state: $_"
        }
    }
} catch {
    Write-Host "Error in main loop: $_"
    Write-EventLog -LogName $global:logName -Source $global:source -EntryType Error -EventId 1011 -Message "Error in main loop: $_"
} finally {
    try {
        Get-EventSubscriber | Where-Object { $_.SourceIdentifier -match "KeyboardEvent" } | Unregister-Event -ErrorAction SilentlyContinue
    } catch {
        Write-Host "Error cleaning up events at the end: $_"
        Write-EventLog -LogName $global:logName -Source $global:source -EntryType Error -EventId 1012 -Message "Error cleaning up events at the end: $_"
    }
}
