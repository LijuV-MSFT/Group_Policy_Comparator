Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = "Stop"

$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$CoreScript = Join-Path $ScriptRoot "Compare-GPOs.ps1"

if (-not (Test-Path -LiteralPath $CoreScript -PathType Leaf)) {
    [System.Windows.Forms.MessageBox]::Show(
        "Could not find Compare-GPOs.ps1 in:`r`n$ScriptRoot",
        "Missing Core Script",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    return
}

. $CoreScript

function Select-Folder {
    param(
        [string]$Description
    )

    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = $Description
    $dialog.ShowNewFolderButton = $true

    $result = $dialog.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.SelectedPath
    }

    return $null
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Group Policy Comparator"
$form.Size = New-Object System.Drawing.Size(1100, 720)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(900, 600)

$font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Font = $font

$labelGpo1 = New-Object System.Windows.Forms.Label
$labelGpo1.Text = "GPO 1 Backup Folder"
$labelGpo1.Location = New-Object System.Drawing.Point(15, 20)
$labelGpo1.Size = New-Object System.Drawing.Size(160, 22)

$textGpo1 = New-Object System.Windows.Forms.TextBox
$textGpo1.Location = New-Object System.Drawing.Point(180, 18)
$textGpo1.Size = New-Object System.Drawing.Size(760, 24)
$textGpo1.Anchor = "Top,Left,Right"

$buttonGpo1 = New-Object System.Windows.Forms.Button
$buttonGpo1.Text = "Browse..."
$buttonGpo1.Location = New-Object System.Drawing.Point(950, 16)
$buttonGpo1.Size = New-Object System.Drawing.Size(110, 28)
$buttonGpo1.Anchor = "Top,Right"

$labelGpo2 = New-Object System.Windows.Forms.Label
$labelGpo2.Text = "GPO 2 Backup Folder"
$labelGpo2.Location = New-Object System.Drawing.Point(15, 60)
$labelGpo2.Size = New-Object System.Drawing.Size(160, 22)

$textGpo2 = New-Object System.Windows.Forms.TextBox
$textGpo2.Location = New-Object System.Drawing.Point(180, 58)
$textGpo2.Size = New-Object System.Drawing.Size(760, 24)
$textGpo2.Anchor = "Top,Left,Right"

$buttonGpo2 = New-Object System.Windows.Forms.Button
$buttonGpo2.Text = "Browse..."
$buttonGpo2.Location = New-Object System.Drawing.Point(950, 56)
$buttonGpo2.Size = New-Object System.Drawing.Size(110, 28)
$buttonGpo2.Anchor = "Top,Right"

$labelOutput = New-Object System.Windows.Forms.Label
$labelOutput.Text = "Output Folder"
$labelOutput.Location = New-Object System.Drawing.Point(15, 100)
$labelOutput.Size = New-Object System.Drawing.Size(160, 22)

$textOutput = New-Object System.Windows.Forms.TextBox
$textOutput.Location = New-Object System.Drawing.Point(180, 98)
$textOutput.Size = New-Object System.Drawing.Size(760, 24)
$textOutput.Anchor = "Top,Left,Right"
$textOutput.Text = "C:\Temp\GPO_Comparison"

$buttonOutput = New-Object System.Windows.Forms.Button
$buttonOutput.Text = "Browse..."
$buttonOutput.Location = New-Object System.Drawing.Point(950, 96)
$buttonOutput.Size = New-Object System.Drawing.Size(110, 28)
$buttonOutput.Anchor = "Top,Right"

$checkIncludeSame = New-Object System.Windows.Forms.CheckBox
$checkIncludeSame.Text = "Include settings that are the same"
$checkIncludeSame.Location = New-Object System.Drawing.Point(180, 135)
$checkIncludeSame.Size = New-Object System.Drawing.Size(260, 24)
$checkIncludeSame.Checked = $true

$buttonCompare = New-Object System.Windows.Forms.Button
$buttonCompare.Text = "Compare GPOs"
$buttonCompare.Location = New-Object System.Drawing.Point(15, 170)
$buttonCompare.Size = New-Object System.Drawing.Size(145, 34)

$buttonOpenCsv = New-Object System.Windows.Forms.Button
$buttonOpenCsv.Text = "Open CSV"
$buttonOpenCsv.Location = New-Object System.Drawing.Point(170, 170)
$buttonOpenCsv.Size = New-Object System.Drawing.Size(110, 34)
$buttonOpenCsv.Enabled = $false

$buttonOpenHtml = New-Object System.Windows.Forms.Button
$buttonOpenHtml.Text = "Open HTML"
$buttonOpenHtml.Location = New-Object System.Drawing.Point(290, 170)
$buttonOpenHtml.Size = New-Object System.Drawing.Size(110, 34)
$buttonOpenHtml.Enabled = $false

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Select two GPO backup folders to begin."
$statusLabel.Location = New-Object System.Drawing.Point(420, 178)
$statusLabel.Size = New-Object System.Drawing.Size(640, 22)
$statusLabel.Anchor = "Top,Left,Right"

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(15, 220)
$grid.Size = New-Object System.Drawing.Size(1045, 400)
$grid.Anchor = "Top,Bottom,Left,Right"
$grid.ReadOnly = $true
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.SelectionMode = "FullRowSelect"
$grid.AutoSizeColumnsMode = "DisplayedCells"
$grid.AutoGenerateColumns = $true

$summaryLabel = New-Object System.Windows.Forms.Label
$summaryLabel.Text = ""
$summaryLabel.Location = New-Object System.Drawing.Point(15, 635)
$summaryLabel.Size = New-Object System.Drawing.Size(1045, 25)
$summaryLabel.Anchor = "Bottom,Left,Right"

$script:lastCsvPath = $null
$script:lastHtmlPath = $null

$buttonGpo1.Add_Click({
    $folder = Select-Folder -Description "Select the first GPO backup folder"
    if ($folder) {
        $textGpo1.Text = $folder
    }
})

$buttonGpo2.Add_Click({
    $folder = Select-Folder -Description "Select the second GPO backup folder"
    if ($folder) {
        $textGpo2.Text = $folder
    }
})

$buttonOutput.Add_Click({
    $folder = Select-Folder -Description "Select the output folder for CSV and HTML reports"
    if ($folder) {
        $textOutput.Text = $folder
    }
})

$buttonCompare.Add_Click({

    try {
        $buttonCompare.Enabled = $false
        $buttonOpenCsv.Enabled = $false
        $buttonOpenHtml.Enabled = $false
        $grid.DataSource = $null
        $summaryLabel.Text = ""
        $statusLabel.Text = "Running comparison..."

        if ([string]::IsNullOrWhiteSpace($textGpo1.Text)) {
            throw "Select the first GPO backup folder."
        }

        if ([string]::IsNullOrWhiteSpace($textGpo2.Text)) {
            throw "Select the second GPO backup folder."
        }

        if ([string]::IsNullOrWhiteSpace($textOutput.Text)) {
            throw "Select an output folder."
        }

        $result = Invoke-GpoComparison `
            -Gpo1BackupFolder $textGpo1.Text `
            -Gpo2BackupFolder $textGpo2.Text `
            -OutputFolder $textOutput.Text `
            -IncludeSame:$checkIncludeSame.Checked

        $script:lastCsvPath = $result.CsvOutputPath
        $script:lastHtmlPath = $result.HtmlOutputPath

        $grid.DataSource = @($result.Differences)

        $totalCount = @($result.Differences).Count
        $addedCount = @($result.Differences | Where-Object DifferenceType -eq "Added").Count
        $removedCount = @($result.Differences | Where-Object DifferenceType -eq "Removed").Count
        $changedCount = @($result.Differences | Where-Object DifferenceType -eq "Changed").Count
        $sameCount = @($result.Differences | Where-Object DifferenceType -eq "Same").Count

        $summaryLabel.Text = "Compared '$($result.Gpo1Name)' to '$($result.Gpo2Name)' | Total: $totalCount | Added: $addedCount | Removed: $removedCount | Changed: $changedCount | Same: $sameCount"
        $statusLabel.Text = "Comparison complete."

        $buttonOpenCsv.Enabled = $true
        $buttonOpenHtml.Enabled = $true
    }
    catch {
        $statusLabel.Text = "Comparison failed."
        [System.Windows.Forms.MessageBox]::Show(
            $_.Exception.Message,
            "Comparison Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
    finally {
        $buttonCompare.Enabled = $true
    }
})

$buttonOpenCsv.Add_Click({
    if ($script:lastCsvPath -and (Test-Path -LiteralPath $script:lastCsvPath)) {
        Start-Process $script:lastCsvPath
    }
})

$buttonOpenHtml.Add_Click({
    if ($script:lastHtmlPath -and (Test-Path -LiteralPath $script:lastHtmlPath)) {
        Start-Process $script:lastHtmlPath
    }
})

$form.Controls.AddRange(@(
    $labelGpo1,
    $textGpo1,
    $buttonGpo1,
    $labelGpo2,
    $textGpo2,
    $buttonGpo2,
    $labelOutput,
    $textOutput,
    $buttonOutput,
    $checkIncludeSame,
    $buttonCompare,
    $buttonOpenCsv,
    $buttonOpenHtml,
    $statusLabel,
    $grid,
    $summaryLabel
))

[void]$form.ShowDialog()
