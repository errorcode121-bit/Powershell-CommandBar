Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Globals
$global:currentJob = $null
$global:lastReceive = $null
$global:history = @()

# Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "PowerShell CommandBar"
$form.Width = 820
$form.Height = 560
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(25,25,25)
$form.ForeColor = [System.Drawing.Color]::White
$form.Font = New-Object System.Drawing.Font("Consolas",10)

# Input box (single line)
$inputBox = New-Object System.Windows.Forms.TextBox
$inputBox.Multiline = $false
$inputBox.Width = 620
$inputBox.Height = 28
$inputBox.Top = 12
$inputBox.Left = 12
$inputBox.BackColor = [System.Drawing.Color]::FromArgb(18,18,18)
$inputBox.ForeColor = [System.Drawing.Color]::White
$inputBox.BorderStyle = "FixedSingle"

# Run button
$runButton = New-Object System.Windows.Forms.Button
$runButton.Text = "Run"
$runButton.Width = 80
$runButton.Height = 28
$runButton.Top = 12
$runButton.Left = 640

# Stop button
$stopButton = New-Object System.Windows.Forms.Button
$stopButton.Text = "Stop"
$stopButton.Width = 80
$stopButton.Height = 28
$stopButton.Top = 12
$stopButton.Left = 728

# Output box (multiline console)
$outputBox = New-Object System.Windows.Forms.TextBox
$outputBox.Multiline = $true
$outputBox.ScrollBars = "Vertical"
$outputBox.WordWrap = $false
$outputBox.ReadOnly = $true
$outputBox.Width = 796
$outputBox.Height = 380
$outputBox.Top = 48
$outputBox.Left = 12
$outputBox.BackColor = [System.Drawing.Color]::FromArgb(10,10,10)
$outputBox.ForeColor = [System.Drawing.Color]::White
$outputBox.Font = New-Object System.Drawing.Font("Consolas",10)

# History list
$historyLabel = New-Object System.Windows.Forms.Label
$historyLabel.Text = "History"
$historyLabel.Top = 436
$historyLabel.Left = 12
$historyLabel.AutoSize = $true
$historyLabel.ForeColor = [System.Drawing.Color]::White

$historyList = New-Object System.Windows.Forms.ListBox
$historyList.Width = 796
$historyList.Height = 72
$historyList.Top = 456
$historyList.Left = 12
$historyList.BackColor = [System.Drawing.Color]::FromArgb(18,18,18)
$historyList.ForeColor = [System.Drawing.Color]::White
$historyList.Font = New-Object System.Drawing.Font("Consolas",9)

# Buttons for extra actions: Clear, Save, Load, Quit
$clearButton = New-Object System.Windows.Forms.Button
$clearButton.Text = "Clear"
$clearButton.Width = 80
$clearButton.Height = 26
$clearButton.Top = 12
$clearButton.Left = 560

$saveButton = New-Object System.Windows.Forms.Button
$saveButton.Text = "Save"
$saveButton.Width = 80
$saveButton.Height = 26
$saveButton.Top = 12
$saveButton.Left = 472

$loadButton = New-Object System.Windows.Forms.Button
$loadButton.Text = "Open"
$loadButton.Width = 80
$loadButton.Height = 26
$loadButton.Top = 12
$loadButton.Left = 384

$quitButton = New-Object System.Windows.Forms.Button
$quitButton.Text = "Quit"
$quitButton.Width = 80
$quitButton.Height = 26
$quitButton.Top = 12
$quitButton.Left = 296

# Timer to poll job output
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 500
$timer.Enabled = $false

# Helper: append text with newline and auto-scroll
function Append-Output {
    param($text)
    if ($null -eq $text) { return }
    $outputBox.AppendText([string]::Concat($text, [Environment]::NewLine))
    $outputBox.SelectionStart = $outputBox.Text.Length
    $outputBox.ScrollToCaret()
}

# Start a job to run the command/script
function Start-CommandJob {
    param($cmd)
    if ($global:currentJob -ne $null -and $global:currentJob.State -eq "Running") {
        Append-Output "A job is already running. Stop it first or wait."
        return
    }
    Append-Output "=== Starting job ==="
    $safeCmd = $cmd

    $job = Start-Job -ScriptBlock {
        param($c)
        try {
            # run the command and capture all output + errors
            $out = Invoke-Expression $c 2>&1
            if ($out -ne $null) {
                $out
            }
        } catch {
            $_ | Out-String
        }
    } -ArgumentList $safeCmd

    $global:currentJob = $job
    $global:lastReceive = 0
    $timer.Enabled = $true
    $global:history += $cmd
    $historyList.Items.Insert(0,$cmd)
}

function Poll-Job {
    if ($global:currentJob -eq $null) { return }
    try {
        $job = $global:currentJob
        # retrieve any output so far
        $out = Receive-Job -Job $job -Keep -ErrorAction SilentlyContinue
        if ($out -ne $null -and $out.Count -gt 0) {
            foreach ($line in $out) { Append-Output $line }
        }
        if ($job.State -ne 'Running') {
            Append-Output "=== Job finished: $($job.State) ==="
            # collect any final output/errors
            $final = Receive-Job -Job $job -Keep -ErrorAction SilentlyContinue
            if ($final -ne $null) { foreach ($l in $final) { Append-Output $l } }
            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
            $global:currentJob = $null
            $timer.Enabled = $false
        }
    } catch {
        Append-Output "Error while polling job: $($_.Exception.Message)"
        $timer.Enabled = $false
        $global:currentJob = $null
    }
}

function Stop-CommandJob {
    if ($global:currentJob -eq $null) {
        Append-Output "No job to stop."
        return
    }
    try {
        Stop-Job -Job $global:currentJob -Force -ErrorAction SilentlyContinue
        Append-Output "Stopping job..."
        Poll-Job
    } catch {
        Append-Output "Failed to stop job: $($_.Exception.Message)"
    }
}

# Events
$runButton.Add_Click({
    $cmd = $inputBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($cmd)) { return }
    Start-CommandJob -cmd $cmd
})

$inputBox.Add_KeyDown({
    param($s, $e)
    if ($e.KeyCode -eq "Enter") {
        $e.SuppressKeyPress = $true
        $runButton.PerformClick()
    }
})

$stopButton.Add_Click({ Stop-CommandJob })

$timer.Add_Tick({ Poll-Job })

$clearButton.Add_Click({ $outputBox.Clear() })
$quitButton.Add_Click({ 
    if ($global:currentJob -ne $null) { Stop-CommandJob }
    $form.Close() 
})

# Save (save inputBox contents to file)
$saveButton.Add_Click({
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "PowerShell Script|*.ps1|Text File|*.txt|All Files|*.*"
    $sfd.FileName = "script.ps1"
    if ($sfd.ShowDialog() -eq "OK") {
        [System.IO.File]::WriteAllText($sfd.FileName, $inputBox.Text)
        Append-Output "Saved to $($sfd.FileName)"
    }
})

# Load (open file into inputBox)
$loadButton.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "PowerShell Script|*.ps1;*.psm1;*.psm1;*.psd1|Text File|*.txt|All Files|*.*"
    if ($ofd.ShowDialog() -eq "OK") {
        $inputBox.Text = [System.IO.File]::ReadAllText($ofd.FileName)
        Append-Output "Loaded $($ofd.FileName)"
    }
})

# History double-click to paste into input
$historyList.Add_DoubleClick({
    if ($historyList.SelectedItem) {
        $inputBox.Text = $historyList.SelectedItem
    }
})

# Keyboard shortcuts
$form.Add_KeyDown({
    param($s,$e)
    if ($e.Control -and $e.KeyCode -eq "S") { $saveButton.PerformClick(); $e.Handled = $true }
    elseif ($e.Control -and $e.KeyCode -eq "O") { $loadButton.PerformClick(); $e.Handled = $true }
    elseif ($e.Control -and $e.KeyCode -eq "K") { $clearButton.PerformClick(); $e.Handled = $true }
    elseif ($e.Control -and $e.KeyCode -eq "Q") { $quitButton.PerformClick(); $e.Handled = $true }
})

# Add controls
$form.Controls.AddRange(@($inputBox,$runButton,$stopButton,$outputBox,$historyLabel,$historyList,$clearButton,$saveButton,$loadButton,$quitButton))

# Make Enter run when typing (AcceptButton)
$form.AcceptButton = $runButton

# Show
$form.ShowDialog() | Out-Null
