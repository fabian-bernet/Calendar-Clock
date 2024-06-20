Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.Office.Interop.Outlook

$form = New-Object System.Windows.Forms.Form
$form.Text = "Calendar Clock"
$form.StartPosition = "CenterScreen"
$form.Topmost = $true
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog

$labelTime = New-Object System.Windows.Forms.Label
$labelTime.AutoSize = $true
$labelTime.Font = New-Object System.Drawing.Font("Arial", 24, [System.Drawing.FontStyle]::Bold)
$labelTime.Location = New-Object System.Drawing.Point(10, 10)
$form.Controls.Add($labelTime)

$labelAppointment = New-Object System.Windows.Forms.Label
$labelAppointment.AutoSize = $true
$labelAppointment.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Regular)
$labelAppointment.ForeColor = [System.Drawing.Color]::Red
$labelAppointment.Location = New-Object System.Drawing.Point(10, 50)
$labelAppointment.MaximumSize = New-Object System.Drawing.Size(380, 0)
$form.Controls.Add($labelAppointment)

$global:nextAppointment = $null

function Update-Clock {
    $labelTime.Text = Get-Date -Format "HH:mm:ss:ff"
    if ($global:nextAppointment -ne $null) {
        $now = Get-Date
        $timeUntilAppointment = $global:nextAppointment.Start - $now
        if ($timeUntilAppointment.TotalSeconds -gt 0) {
            $remainingTimeFormatted = "{0:hh\:mm\:ss}" -f [timespan]::fromseconds($timeUntilAppointment.TotalSeconds)
            $labelAppointment.Text = "$remainingTimeFormatted remaining: $($global:nextAppointment.Subject)"
        } else {
            $labelAppointment.Text = "Now: $($global:nextAppointment.Subject)"
        }
    }
}

function Update-Appointment {
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        $calendar = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderCalendar)
        $appointments = $calendar.Items
        $appointments.Sort("[Start]")
        $appointments.IncludeRecurrences = $true
        $now = Get-Date
        $filter = "[Start] >= '" + $now.ToString("g") + "'"
        $global:nextAppointment = $appointments.Find($filter)
        
        if ($global:nextAppointment -eq $null) {
            $labelAppointment.Text = "No upcoming events."
        }
    } catch {
        $labelAppointment.Text = "Error when retrieving the appointment."
    }
}

$timerClock = New-Object System.Windows.Forms.Timer
$timerClock.Interval = 100
$timerClock.Add_Tick({
    Update-Clock
})
$timerClock.Start()

$timerAppointment = New-Object System.Windows.Forms.Timer
$timerAppointment.Interval = 60000
$timerAppointment.Add_Tick({
    Update-Appointment
})
$timerAppointment.Start()

Update-Appointment

$form.Add_Closing({
    $timerClock.Stop()
    $timerAppointment.Stop()
})

$form.AutoSize = $true
$form.AutoSizeMode = "GrowAndShrink"
$form.ShowDialog() | Out-Null
