# Load notification module
Import-Module BurntToast

$global:keyboardId = "VID_046D&PID_C31C" # Use this in PowerShell -> Get-PnpDevice -Class "USB" | Select-Object Name, InstanceId
$global:monitorId = "\\.\DISPLAY1\Monitor0" # Get this from ControlMyMonitor
$global:modeA_monitor_output_id = 15 #HDMI, DP, ETC... Get this from ControlMyMonitor
$global:modeB_monitor_output_id = 17 #HDMI, DP, ETC... Get this from ControlMyMonitor


$global:modeA_msg = "MODE A ENABLED"
$global:modeB_msg = "MODE B ENABLED"
$global:controlMyMonitor = Join-Path $PSScriptRoot "ControlMyMonitor.exe"
$global:keyboardConnected = $false
$global:logName = "Application"
$global:source = "MonitorSwitcher"


# Create event source if it doesn't exist
if (-not (Get-EventLog -LogName $global:logName | Where-Object { $_.Source -eq $global:source })) {
    New-EventLog -LogName $global:logName -Source $global:source
}

function SwitchToModeA {
    try {
        ActivateMonitor
        New-BurntToastNotification -Text "MonitorSwitcher (by tonikelope)", $global:modeA_msg
        & $global:controlMyMonitor /SetValue $global:monitorId 60 $global:modeA_monitor_output_id
    } catch {
        Write-Host "Error switching to Mode A: $_"
        Write-EventLog -LogName $global:logName -Source $global:source -EntryType Error -EventId 1001 -Message "Error switching to Mode A: $_"
    }
}

function SwitchToModeB {
    try {
        New-BurntToastNotification -Text "MonitorSwitcher (by tonikelope)", $global:modeB_msg
        & $global:controlMyMonitor /SetValue $global:monitorId 60 $global:modeB_monitor_output_id
    } catch {
        Write-Host "Error switching to Mode B: $_"
        Write-EventLog -LogName $global:logName -Source $global:source -EntryType Error -EventId 1002 -Message "Error switching to Mode B: $_"
    }
}

function IsConnected {
    try {
        $dev = Get-PnpDevice -Class Keyboard | Where-Object { 
            $_.InstanceId -like "*$global:keyboardId*" -and $_.Status -eq "OK" 
        }
        return $dev -ne $null
    } catch {
        Write-Host "Error checking if keyboard is connected: $_"
        Write-EventLog -LogName $global:logName -Source $global:source -EntryType Error -EventId 1003 -Message "Error checking if keyboard is connected: $_"
        return $false
    }
}

function ActivateMonitor {
    try {
        (Add-Type '[DllImport("user32.dll")]public static extern int SendMessage(int hWnd, int hMsg, int wParam, int lParam);' -Name a -Pas)::SendMessage(-1,0x0112,0xF170,-1)
    } catch {
        Write-Host "Error activating monitor: $_"
        Write-EventLog -LogName $global:logName -Source $global:source -EntryType Error -EventId 1004 -Message "Error activating monitor: $_"
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
        SwitchToModeA
    } else {
        SwitchToModeB
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

try {
    while ($true) {
        try {
            Wait-Event -SourceIdentifier "KeyboardEvent" | Out-Null
        } catch {
            Write-Host "Error waiting for event: $_"
            Write-EventLog -LogName $global:logName -Source $global:source -EntryType Error -EventId 1008 -Message "Error waiting for event: $_"
            break
        }

        # Remove event to prevent accumulation
        try {
            Remove-Event -SourceIdentifier "KeyboardEvent"
        } catch {
            Write-Host "Error removing event: $_"
            Write-EventLog -LogName $global:logName -Source $global:source -EntryType Error -EventId 1009 -Message "Error removing event: $_"
        }

        Write-Host "Keyboard event received..."

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
