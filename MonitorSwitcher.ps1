[CmdletBinding()]
param(
    # Show a toast notification on startup and when the keyboard changes mode.
    [switch]$ShowNotifications
)

$script:keyboardId = 'VID_1532&PID_02A2' # Razer Ornata V3 X
$script:monitors = @(
    @{ Id = '\\.\DISPLAY1\Monitor0'; ModeA = 15; ModeB = 17 }
    @{ Id = '\\.\DISPLAY2\Monitor0'; ModeA = 15; ModeB = 17 }
)
$script:controlMyMonitor = Join-Path $PSScriptRoot 'ControlMyMonitor.exe'
$script:icon = Join-Path $PSScriptRoot 'MonitorDual.png'
$script:showNotifications = [bool]$ShowNotifications
$script:eventLogAvailable = $false
$script:keyboardQueryFailed = $false
$script:displayPower = $null
$script:lastWakeGeneration = 0L
$script:keyboardConnected = $null
$script:recovery = $null
$script:clock = [Diagnostics.Stopwatch]::StartNew()

function Start-DisplayPowerMonitor {
    if (-not ('MonitorSwitcher.DisplayPowerMonitor' -as [type])) {
        Add-Type -ErrorAction Stop -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace MonitorSwitcher {
    public sealed class DisplayPowerMonitor : IDisposable {
        private const uint DeviceNotifyCallback = 2;
        private const uint PowerSettingChange = 0x8013;
        private static readonly Guid SessionDisplayStatus = new Guid("2B84C20E-AD23-4DDF-93DB-05FFBD7EFCA5");

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        private delegate uint DeviceNotify(IntPtr context, uint type, IntPtr setting);

        [StructLayout(LayoutKind.Sequential)]
        private struct Subscription {
            public DeviceNotify Callback;
            public IntPtr Context;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct DisplaySetting {
            public Guid Id;
            public uint DataLength;
            public int State;
        }

        [DllImport("powrprof.dll", ExactSpelling = true)]
        private static extern uint PowerSettingRegisterNotification(ref Guid setting, uint flags,
            ref Subscription subscription, out IntPtr handle);

        [DllImport("powrprof.dll", ExactSpelling = true)]
        private static extern uint PowerSettingUnregisterNotification(IntPtr handle);

        public struct Snapshot {
            public int State;
            public long WakeGeneration;
        }

        private readonly object gate = new object();
        private readonly DeviceNotify callback;
        private IntPtr handle;
        private Snapshot snapshot = new Snapshot { State = -1 };
        private bool wasOff;

        public DisplayPowerMonitor() {
            callback = OnPowerSettingChanged;
            var subscription = new Subscription { Callback = callback };
            Guid setting = SessionDisplayStatus;
            uint result = PowerSettingRegisterNotification(ref setting, DeviceNotifyCallback, ref subscription, out handle);
            if (result != 0) { throw new Win32Exception((int)result); }
        }

        private uint OnPowerSettingChanged(IntPtr context, uint type, IntPtr setting) {
            if (type != PowerSettingChange || setting == IntPtr.Zero) { return 0; }
            var id = (Guid)Marshal.PtrToStructure(setting, typeof(Guid));
            uint length = unchecked((uint)Marshal.ReadInt32(setting, 16));
            if (id != SessionDisplayStatus || length != 4) { return 0; }
            var update = (DisplaySetting)Marshal.PtrToStructure(setting, typeof(DisplaySetting));
            if (update.State < 0 || update.State > 2) { return 0; }
            lock (gate) {
                snapshot.State = update.State;
                if (update.State == 0) { wasOff = true; }
                if (update.State == 1 && wasOff) {
                    wasOff = false;
                    snapshot.WakeGeneration++;
                }
            }
            return 0;
        }

        public Snapshot Read() {
            lock (gate) { return snapshot; }
        }

        public void Dispose() {
            if (handle != IntPtr.Zero) {
                uint result = PowerSettingUnregisterNotification(handle);
                if (result != 0) { throw new Win32Exception((int)result); }
                handle = IntPtr.Zero;
            }
            GC.KeepAlive(callback);
        }
    }
}
'@
    }
    return [MonitorSwitcher.DisplayPowerMonitor]::new()
}

function Get-DisplayStatus {
    if ($null -ne $script:displayPower) { return $script:displayPower.Read() }
    return @{ State = -1; WakeGeneration = 0L }
}

function Write-SwitcherWarning {
    param([int]$EventId, [string]$Message)
    Write-Warning $Message
    if ($script:eventLogAvailable) {
        Write-EventLog -LogName Application -Source MonitorSwitcher -EntryType Warning -EventId $EventId -Message $Message -ErrorAction SilentlyContinue
    }
}

function Get-KeyboardConnection {
    try {
        $devices = @(Get-CimInstance -ClassName Win32_PnPEntity -Filter "PNPClass = 'Keyboard' AND Present = TRUE" -ErrorAction Stop)
        $script:keyboardQueryFailed = $false
        return @($devices | Where-Object PNPDeviceID -Like "*$script:keyboardId*").Count -gt 0
    } catch {
        if (-not $script:keyboardQueryFailed) {
            Write-SwitcherWarning 1003 "Could not check keyboard presence: $_"
            $script:keyboardQueryFailed = $true
        }
        return $null
    }
}

function Invoke-ControlMyMonitor {
    param([string]$Arguments)

    $process = [Diagnostics.Process]::new()
    $process.StartInfo.FileName = $script:controlMyMonitor
    $process.StartInfo.Arguments = $Arguments
    $process.StartInfo.UseShellExecute = $false
    $process.StartInfo.CreateNoWindow = $true
    $process.StartInfo.WindowStyle = [Diagnostics.ProcessWindowStyle]::Hidden
    try {
        [void]$process.Start()
        if (-not $process.WaitForExit(2500)) {
            if (-not $process.HasExited) { $process.Kill() }
            [void]$process.WaitForExit(500)
            Write-Verbose "ControlMyMonitor timed out: $Arguments"
            return $null
        }
        return $process.ExitCode
    } catch {
        Write-Verbose "ControlMyMonitor failed ($Arguments): $_"
        return $null
    } finally {
        $process.Dispose()
    }
}

function Set-MonitorInput {
    param([string]$MonitorId, [int]$Target)
    $result = Invoke-ControlMyMonitor ('/SetValue "{0}" 60 {1}' -f $MonitorId, $Target)
    if ($null -eq $result -or $result -ne 0) {
        Write-SwitcherWarning 1001 "Input switch to $Target could not be confirmed for $MonitorId (exit: $result)."
    }
}

function Show-ModeNotification {
    param([bool]$KeyboardConnected)
    if ($script:showNotifications) {
        $mode = if ($KeyboardConnected) { 'A' } else { 'B' }
        try {
            New-BurntToastNotification -AppLogo $script:icon -Text 'MonitorSwitcher (by tonikelope)', "MODE $mode ENABLED" -ErrorAction Stop
        } catch {
            Write-SwitcherWarning 1014 "Could not show mode notification: $_"
        }
    }
}

function New-InputRecovery {
    param([long]$WakeGeneration)
    $now = $script:clock.ElapsedMilliseconds
    return @{
        WakeGeneration = $WakeGeneration
        Due = $now + 1200
        Deadline = $now + 12000
        Pending = @(
            foreach ($monitor in $script:monitors) {
                @{ Monitor = $monitor; SawAlternate = $false; Done = $false }
            }
        )
    }
}

function Test-RecoveryDisplay {
    param($Recovery)
    $display = Get-DisplayStatus
    return ($script:clock.ElapsedMilliseconds -lt $Recovery.Deadline -and
        $display.WakeGeneration -eq $Recovery.WakeGeneration -and
        ($display.State -eq 1 -or ($display.State -eq -1 -and $Recovery.WakeGeneration -eq 0)))
}

function Invoke-InputRecovery {
    param($Recovery, [bool]$KeyboardConnected)
    if ($script:clock.ElapsedMilliseconds -lt $Recovery.Due) { return }

    foreach ($pending in $Recovery.Pending) {
        if ($pending.Done) { continue }
        if (-not (Test-RecoveryDisplay $Recovery)) { return }
        $monitor = $pending.Monitor
        $target = if ($KeyboardConnected) { $monitor.ModeA } else { $monitor.ModeB }
        $alternate = if ($KeyboardConnected) { $monitor.ModeB } else { $monitor.ModeA }
        $inputCode = Invoke-ControlMyMonitor ('/GetValue "{0}" 60' -f $monitor.Id)
        if ($inputCode -eq $target) {
            $pending.Done = $true
        } elseif ($inputCode -eq $alternate) {
            # A valid alternate must survive a second read in a later pass.
            if ($pending.SawAlternate) {
                $currentKeyboard = Get-KeyboardConnection
                if ($null -eq $currentKeyboard -or $currentKeyboard -ne $KeyboardConnected) { return }
                if (-not (Test-RecoveryDisplay $Recovery)) { return }
                $pending.Done = $true
                Set-MonitorInput $monitor.Id $target
            }
            $pending.SawAlternate = $true
        } else {
            # A missing, failed or unconfigured read cannot authorize a write.
            $pending.SawAlternate = $false
        }
    }

    $Recovery.Due = $script:clock.ElapsedMilliseconds + 750
    if (@($Recovery.Pending | Where-Object { -not $_.Done }).Count -eq 0) {
        $script:recovery = $null
    }
}

function Update-MonitorSwitcher {
    $keyboard = Get-KeyboardConnection
    $keyboardChanged = $false
    if ($null -ne $keyboard -and $null -ne $script:keyboardConnected -and $keyboard -ne $script:keyboardConnected) {
        # Ignore a single transient USB re-enumeration or failed presence query.
        Start-Sleep -Milliseconds 250
        $confirmed = Get-KeyboardConnection
        if ($null -eq $confirmed -or $confirmed -ne $keyboard) { $keyboard = $null }
        else { $keyboardChanged = $true }
    }

    # Read display state after the debounce in case the same USB event woke it.
    $display = Get-DisplayStatus
    if ($display.WakeGeneration -ne $script:lastWakeGeneration) {
        $script:lastWakeGeneration = $display.WakeGeneration
        $script:recovery = New-InputRecovery $display.WakeGeneration
        Write-Verbose "Display wake $($display.WakeGeneration) detected."
    }

    if ($script:recovery -and $script:clock.ElapsedMilliseconds -ge $script:recovery.Deadline) {
        $pendingIds = @($script:recovery.Pending | Where-Object { -not $_.Done } | ForEach-Object { $_.Monitor.Id })
        Write-Verbose "Input recovery expired for $($pendingIds -join ', ')."
        $script:recovery = $null
    }
    if ($null -eq $keyboard) { return }

    if ($null -eq $script:keyboardConnected) {
        $script:keyboardConnected = $keyboard
        if ($display.WakeGeneration -eq 0) { $script:recovery = New-InputRecovery 0 }
        Show-ModeNotification $keyboard
    } elseif ($keyboardChanged) {
        $script:keyboardConnected = $keyboard
        if ($script:recovery -and $script:recovery.WakeGeneration -gt 0) {
            # Keep one guarded recovery for the new mode. Old readings no longer apply.
            foreach ($pending in $script:recovery.Pending) {
                $pending.SawAlternate = $false
                $pending.Done = $false
            }
        } else {
            $script:recovery = $null
            foreach ($monitor in $script:monitors) {
                if ($monitor.ModeA -eq $monitor.ModeB) { continue }
                $target = if ($keyboard) { $monitor.ModeA } else { $monitor.ModeB }
                Set-MonitorInput $monitor.Id $target
            }
        }
        Show-ModeNotification $keyboard
    }

    if ($script:recovery) { Invoke-InputRecovery $script:recovery $keyboard }
}

function Start-MonitorSwitcher {
    $created = $false
    $instance = [Threading.Mutex]::new($true, 'Local\MonitorSwitcher', [ref]$created)
    $deviceEvent = 'MonitorSwitcher.DeviceChange'
    try {
        if (-not $created) { return }
        if (-not (Test-Path -LiteralPath $script:controlMyMonitor -PathType Leaf)) {
            throw "ControlMyMonitor was not found: $script:controlMyMonitor"
        }
        if (Get-Command New-EventLog -ErrorAction SilentlyContinue) {
            try {
                if (-not [Diagnostics.EventLog]::SourceExists('MonitorSwitcher')) {
                    New-EventLog -LogName Application -Source MonitorSwitcher -ErrorAction Stop
                }
                $script:eventLogAvailable = $true
            } catch { Write-Verbose "Application log source unavailable: $_" }
        }
        if ($script:showNotifications) {
            try { Import-Module BurntToast -ErrorAction Stop }
            catch {
                $script:showNotifications = $false
                Write-SwitcherWarning 1014 "BurntToast is unavailable: $_"
            }
        }
        try { $script:displayPower = Start-DisplayPowerMonitor }
        catch { Write-SwitcherWarning 1013 "Display power notifications unavailable: $_" }

        try {
            Register-CimIndicationEvent -Query 'SELECT * FROM Win32_DeviceChangeEvent WHERE EventType = 2 OR EventType = 3' -SourceIdentifier $deviceEvent -ErrorAction Stop | Out-Null
        } catch { Write-SwitcherWarning 1007 "Device change notifications unavailable. Polling keyboard presence: $_" }

        while ($true) {
            Update-MonitorSwitcher
            $timeout = if ($script:recovery) { 1 } else { 2 }
            Wait-Event -SourceIdentifier $deviceEvent -Timeout $timeout | Out-Null
            Remove-Event -SourceIdentifier $deviceEvent -ErrorAction SilentlyContinue
        }
    } finally {
        if ($created) {
            Unregister-Event -SourceIdentifier $deviceEvent -ErrorAction SilentlyContinue
            Remove-Event -SourceIdentifier $deviceEvent -ErrorAction SilentlyContinue
            if ($script:displayPower) {
                try { $script:displayPower.Dispose() }
                catch { Write-SwitcherWarning 1013 "Could not release display power notifications: $_" }
            }
            $instance.ReleaseMutex()
        }
        $instance.Dispose()
    }
}

Start-MonitorSwitcher
