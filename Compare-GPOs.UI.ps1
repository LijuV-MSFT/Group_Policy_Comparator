Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Data

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

function Initialize-ModernFolderPicker {
    if ("ModernFolderPicker.FileDialog" -as [type]) {
        return
    }

    $source = @"
using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace ModernFolderPicker
{
    [Flags]
    public enum FileOpenOptions : uint
    {
        OverwritePrompt        = 0x00000002,
        StrictFileTypes        = 0x00000004,
        NoChangeDir            = 0x00000008,
        PickFolders            = 0x00000020,
        ForceFileSystem        = 0x00000040,
        AllNonStorageItems     = 0x00000080,
        NoValidate             = 0x00000100,
        AllowMultiSelect       = 0x00000200,
        PathMustExist          = 0x00000800,
        FileMustExist          = 0x00001000,
        CreatePrompt           = 0x00002000,
        ShareAware             = 0x00004000,
        NoReadOnlyReturn       = 0x00008000,
        NoTestFileCreate       = 0x00010000,
        HideMruPlaces          = 0x00020000,
        HidePinnedPlaces       = 0x00040000,
        NoDereferenceLinks     = 0x00100000,
        OkButtonNeedsInteraction = 0x00200000,
        DontAddToRecent        = 0x02000000,
        ForceShowHidden        = 0x10000000,
        DefaultNoMiniMode      = 0x20000000,
        ForcePreviewPaneOn     = 0x40000000,
        SupportStreamableItems = 0x80000000
    }

    public enum Sigdn : uint
    {
        FileSysPath = 0x80058000
    }

    [ComImport]
    [Guid("42f85136-db7e-439c-85f1-e4075d135fc8")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IFileDialog
    {
        [PreserveSig]
        int Show(IntPtr parent);

        void SetFileTypes(uint cFileTypes, IntPtr rgFilterSpec);
        void SetFileTypeIndex(uint iFileType);
        void GetFileTypeIndex(out uint piFileType);
        void Advise(IntPtr pfde, out uint pdwCookie);
        void Unadvise(uint dwCookie);
        void SetOptions(FileOpenOptions fos);
        void GetOptions(out FileOpenOptions pfos);
        void SetDefaultFolder(IShellItem psi);
        void SetFolder(IShellItem psi);
        void GetFolder(out IShellItem ppsi);
        void GetCurrentSelection(out IShellItem ppsi);
        void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string pszName);
        void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);
        void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
        void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string pszText);
        void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
        void GetResult(out IShellItem ppsi);
        void AddPlace(IShellItem psi, uint fdap);
        void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);
        void Close(int hr);
        void SetClientGuid(ref Guid guid);
        void ClearClientData();
        void SetFilter(IntPtr pFilter);
    }

    [ComImport]
    [Guid("d57c7288-d4ad-4768-be02-9d969532d960")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IFileOpenDialog : IFileDialog
    {
        [PreserveSig]
        new int Show(IntPtr parent);

        new void SetFileTypes(uint cFileTypes, IntPtr rgFilterSpec);
        new void SetFileTypeIndex(uint iFileType);
        new void GetFileTypeIndex(out uint piFileType);
        new void Advise(IntPtr pfde, out uint pdwCookie);
        new void Unadvise(uint dwCookie);
        new void SetOptions(FileOpenOptions fos);
        new void GetOptions(out FileOpenOptions pfos);
        new void SetDefaultFolder(IShellItem psi);
        new void SetFolder(IShellItem psi);
        new void GetFolder(out IShellItem ppsi);
        new void GetCurrentSelection(out IShellItem ppsi);
        new void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string pszName);
        new void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);
        new void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
        new void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string pszText);
        new void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
        new void GetResult(out IShellItem ppsi);
        new void AddPlace(IShellItem psi, uint fdap);
        new void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);
        new void Close(int hr);
        new void SetClientGuid(ref Guid guid);
        new void ClearClientData();
        new void SetFilter(IntPtr pFilter);

        void GetResults(IntPtr ppsai);
        void GetSelectedItems(IntPtr ppsai);
    }

    [ComImport]
    [Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IShellItem
    {
        void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppv);
        void GetParent(out IShellItem ppsi);
        void GetDisplayName(Sigdn sigdnName, out IntPtr ppszName);
        void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
        void Compare(IShellItem psi, uint hint, out int piOrder);
    }

    [ComImport]
    [Guid("DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7")]
    public class FileOpenDialogRCW
    {
    }

    public static class NativeMethods
    {
        [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        public static extern void SHCreateItemFromParsingName(
            [MarshalAs(UnmanagedType.LPWStr)] string pszPath,
            IntPtr pbc,
            ref Guid riid,
            out IShellItem ppv);
    }

    public static class FileDialog
    {
        public static string SelectFolder(string title, string initialFolder, bool forcePreviewPane)
        {
            IFileOpenDialog dialog = (IFileOpenDialog)new FileOpenDialogRCW();

            try
            {
                FileOpenOptions options;
                dialog.GetOptions(out options);

                options |= FileOpenOptions.PickFolders;
                options |= FileOpenOptions.ForceFileSystem;
                options |= FileOpenOptions.PathMustExist;

                if (forcePreviewPane)
                {
                    options |= FileOpenOptions.ForcePreviewPaneOn;
                }

                dialog.SetOptions(options);
                dialog.SetTitle(title);
                dialog.SetOkButtonLabel("Select Folder");

                if (!String.IsNullOrWhiteSpace(initialFolder) &&
                    System.IO.Directory.Exists(initialFolder))
                {
                    Guid shellItemGuid = typeof(IShellItem).GUID;
                    IShellItem initialShellItem;
                    NativeMethods.SHCreateItemFromParsingName(
                        initialFolder,
                        IntPtr.Zero,
                        ref shellItemGuid,
                        out initialShellItem);

                    dialog.SetDefaultFolder(initialShellItem);
                }

                int result = dialog.Show(IntPtr.Zero);

                const int HRESULT_CANCELLED = unchecked((int)0x800704C7);

                if (result == HRESULT_CANCELLED)
                {
                    return null;
                }

                if (result != 0)
                {
                    Marshal.ThrowExceptionForHR(result);
                }

                IShellItem selectedItem;
                dialog.GetResult(out selectedItem);

                IntPtr pathPointer;
                selectedItem.GetDisplayName(Sigdn.FileSysPath, out pathPointer);

                try
                {
                    return Marshal.PtrToStringUni(pathPointer);
                }
                finally
                {
                    Marshal.FreeCoTaskMem(pathPointer);
                }
            }
            finally
            {
                if (dialog != null && Marshal.IsComObject(dialog))
                {
                    Marshal.FinalReleaseComObject(dialog);
                }
            }
        }
    }
}
"@

    Add-Type -TypeDefinition $source -Language CSharp
}

function Select-Folder {
    param(
        [Parameter(Mandatory)]
        [string]$Description,

        [string]$InitialFolder,

        [switch]$EnablePreviewPane
    )

    Initialize-ModernFolderPicker

    return [ModernFolderPicker.FileDialog]::SelectFolder(
        $Description,
        $InitialFolder,
        $EnablePreviewPane.IsPresent
    )
}

function Get-FirstExistingFolder {
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string[]]$CandidatePath
    )

    foreach ($path in @($CandidatePath)) {

        if ([string]::IsNullOrWhiteSpace($path)) {
            continue
        }

        if (Test-Path -LiteralPath $path -PathType Container) {
            return $path
        }
    }

    return $null
}

function Get-GridDisplayColumns {
    param(
        [Parameter(Mandatory)]
        [string]$Gpo1Name,

        [Parameter(Mandatory)]
        [string]$Gpo2Name
    )

    @(
        "PolicyScope",
        "PolicyContainer",
        "SettingCategory",
        "SettingType",
        "SettingName",
        "Path",
        "DifferenceType",
        "$Gpo1Name Value",
        "$Gpo2Name Value"
    )
}

function Convert-ObjectsToDataTable {
    param(
        [AllowNull()]
        [object[]]$Rows,

        [Parameter(Mandatory)]
        [string[]]$Columns
    )

    $dataTable = New-Object System.Data.DataTable

    foreach ($column in $Columns) {
        [void]$dataTable.Columns.Add($column, [string])
    }

    foreach ($row in @($Rows)) {
        $dataRow = $dataTable.NewRow()

        foreach ($column in $Columns) {
            $property = $row.PSObject.Properties[$column]

            if ($property) {
                $dataRow[$column] = [string]$property.Value
            }
            else {
                $dataRow[$column] = ""
            }
        }

        [void]$dataTable.Rows.Add($dataRow)
    }

    # Important: prevent PowerShell from enumerating the DataTable rows
    return ,$dataTable
}

function Refresh-GridFromCsv {
    if (-not $script:lastCsvPath -or -not (Test-Path -LiteralPath $script:lastCsvPath -PathType Leaf)) {
        return
    }

    if (-not $script:gridDisplayColumns) {
        return
    }

    $rows = @(Import-Csv -LiteralPath $script:lastCsvPath)

    if ($checkOnlyShowDifferences.Checked) {
        $rows = @($rows | Where-Object { $_.DifferenceType -ne "Same" })
    }

    [System.Data.DataTable]$dataTable = Convert-ObjectsToDataTable `
        -Rows $rows `
        -Columns $script:gridDisplayColumns

    $grid.DataSource = $null
    $grid.DataSource = $dataTable

    $script:visibleRowCount = @($rows).Count

    Format-Grid -Grid $grid
    Apply-GridRowColors -Grid $grid
    Update-SummaryLabel
}

function Format-Grid {
    param(
        [Parameter(Mandatory)]
        [System.Windows.Forms.DataGridView]$Grid
    )

    $Grid.AutoGenerateColumns = $true
    $Grid.AutoSizeColumnsMode = "DisplayedCells"
    $Grid.AutoSizeRowsMode = "DisplayedCells"
    $Grid.DefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::True
    $Grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $Grid.RowHeadersVisible = $false

    foreach ($column in $Grid.Columns) {
        $column.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic

        if ($column.Name -like "* Value") {
            $column.DefaultCellStyle.Font = New-Object System.Drawing.Font("Consolas", 9)
        }

        if ($column.Name -eq "Path") {
            $column.DefaultCellStyle.Font = New-Object System.Drawing.Font("Consolas", 9)
        }
    }
}

function Apply-GridRowColors {
    param(
        [Parameter(Mandatory)]
        [System.Windows.Forms.DataGridView]$Grid
    )

    if (-not $Grid.Columns.Contains("DifferenceType")) {
        return
    }

    foreach ($row in $Grid.Rows) {
        if ($row.IsNewRow) {
            continue
        }

        $differenceType = [string]$row.Cells["DifferenceType"].Value

        switch ($differenceType) {
            "Added" {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightGreen
            }
            "Removed" {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::Salmon
            }
            "Changed" {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightSkyBlue
            }
            default {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::White
            }
        }
    }
}

function Update-SummaryLabel {
    if (-not $script:lastResult) {
        return
    }

    $totalCount = @($script:lastResult.Differences).Count
    $addedCount = @($script:lastResult.Differences | Where-Object DifferenceType -eq "Added").Count
    $removedCount = @($script:lastResult.Differences | Where-Object DifferenceType -eq "Removed").Count
    $changedCount = @($script:lastResult.Differences | Where-Object DifferenceType -eq "Changed").Count
    $sameCount = @($script:lastResult.Differences | Where-Object DifferenceType -eq "Same").Count

    $visibleCount = if ($null -ne $script:visibleRowCount) {
        $script:visibleRowCount
    }
    else {
        $totalCount
    }

    $summaryLabel.Text = "Compared '$($script:lastResult.Gpo1Name)' to '$($script:lastResult.Gpo2Name)' | Total: $totalCount | Visible: $visibleCount | Added: $addedCount | Removed: $removedCount | Changed: $changedCount | Same: $sameCount"
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Group Policy Comparator"
$form.Size = New-Object System.Drawing.Size(1200, 780)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(1000, 650)

$font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Font = $font

$labelGpo1 = New-Object System.Windows.Forms.Label
$labelGpo1.Text = "GPO 1 Backup Folder"
$labelGpo1.Location = New-Object System.Drawing.Point(15, 20)
$labelGpo1.Size = New-Object System.Drawing.Size(160, 22)

$textGpo1 = New-Object System.Windows.Forms.TextBox
$textGpo1.Location = New-Object System.Drawing.Point(180, 18)
$textGpo1.Size = New-Object System.Drawing.Size(870, 24)
$textGpo1.Anchor = "Top,Left,Right"

$buttonGpo1 = New-Object System.Windows.Forms.Button
$buttonGpo1.Text = "Browse..."
$buttonGpo1.Location = New-Object System.Drawing.Point(1060, 16)
$buttonGpo1.Size = New-Object System.Drawing.Size(110, 28)
$buttonGpo1.Anchor = "Top,Right"

$labelGpo2 = New-Object System.Windows.Forms.Label
$labelGpo2.Text = "GPO 2 Backup Folder"
$labelGpo2.Location = New-Object System.Drawing.Point(15, 60)
$labelGpo2.Size = New-Object System.Drawing.Size(160, 22)

$textGpo2 = New-Object System.Windows.Forms.TextBox
$textGpo2.Location = New-Object System.Drawing.Point(180, 58)
$textGpo2.Size = New-Object System.Drawing.Size(870, 24)
$textGpo2.Anchor = "Top,Left,Right"

$buttonGpo2 = New-Object System.Windows.Forms.Button
$buttonGpo2.Text = "Browse..."
$buttonGpo2.Location = New-Object System.Drawing.Point(1060, 56)
$buttonGpo2.Size = New-Object System.Drawing.Size(110, 28)
$buttonGpo2.Anchor = "Top,Right"

$labelOutput = New-Object System.Windows.Forms.Label
$labelOutput.Text = "Output Folder"
$labelOutput.Location = New-Object System.Drawing.Point(15, 100)
$labelOutput.Size = New-Object System.Drawing.Size(160, 22)

$textOutput = New-Object System.Windows.Forms.TextBox
$textOutput.Location = New-Object System.Drawing.Point(180, 98)
$textOutput.Size = New-Object System.Drawing.Size(870, 24)
$textOutput.Anchor = "Top,Left,Right"
$textOutput.Text = "C:\Temp\GPO_Comparison"

$buttonOutput = New-Object System.Windows.Forms.Button
$buttonOutput.Text = "Browse..."
$buttonOutput.Location = New-Object System.Drawing.Point(1060, 96)
$buttonOutput.Size = New-Object System.Drawing.Size(110, 28)
$buttonOutput.Anchor = "Top,Right"

$checkIncludeSame = New-Object System.Windows.Forms.CheckBox
$checkIncludeSame.Text = "Include settings that are the same in generated reports"
$checkIncludeSame.Location = New-Object System.Drawing.Point(180, 135)
$checkIncludeSame.Size = New-Object System.Drawing.Size(340, 24)
$checkIncludeSame.Checked = $true

$checkOnlyShowDifferences = New-Object System.Windows.Forms.CheckBox
$checkOnlyShowDifferences.Text = "Only show differences"
$checkOnlyShowDifferences.Location = New-Object System.Drawing.Point(540, 135)
$checkOnlyShowDifferences.Size = New-Object System.Drawing.Size(220, 24)
$checkOnlyShowDifferences.Checked = $false

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
$statusLabel.Size = New-Object System.Drawing.Size(750, 22)
$statusLabel.Anchor = "Top,Left,Right"

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(15, 220)
$grid.Size = New-Object System.Drawing.Size(1155, 455)
$grid.Anchor = "Top,Bottom,Left,Right"
$grid.ReadOnly = $true
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.SelectionMode = "FullRowSelect"
$grid.MultiSelect = $true
$grid.AutoGenerateColumns = $true
$grid.AutoSizeColumnsMode = "DisplayedCells"
$grid.BackgroundColor = [System.Drawing.Color]::White
$grid.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$grid.ClipboardCopyMode = [System.Windows.Forms.DataGridViewClipboardCopyMode]::EnableAlwaysIncludeHeaderText

$summaryLabel = New-Object System.Windows.Forms.Label
$summaryLabel.Text = ""
$summaryLabel.Location = New-Object System.Drawing.Point(15, 695)
$summaryLabel.Size = New-Object System.Drawing.Size(1155, 25)
$summaryLabel.Anchor = "Bottom,Left,Right"

$script:lastCsvPath = $null
$script:lastHtmlPath = $null
$script:lastResult = $null
$script:gridDisplayColumns = $null
$script:visibleRowCount = 0

$buttonGpo1.Add_Click({
    try {
        $initialFolder = Get-FirstExistingFolder -CandidatePath @(
            $textGpo1.Text
            $textOutput.Text
            [Environment]::GetFolderPath("Desktop")
        )

        $folder = Select-Folder `
            -Description "Select the first GPO backup folder" `
            -InitialFolder $initialFolder `
            -EnablePreviewPane

        if (-not [string]::IsNullOrWhiteSpace($folder)) {
            $textGpo1.Text = $folder
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            $_.Exception.Message,
            "Folder Selection Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
})

$buttonGpo2.Add_Click({
    try {
        $gpo1ParentFolder = $null

        if (-not [string]::IsNullOrWhiteSpace($textGpo1.Text) -and
            (Test-Path -LiteralPath $textGpo1.Text -PathType Container)) {

            $gpo1ParentFolder = Split-Path -Parent $textGpo1.Text
        }

        $initialFolder = Get-FirstExistingFolder -CandidatePath @(
            $textGpo2.Text
            $gpo1ParentFolder
            $textOutput.Text
            [Environment]::GetFolderPath("Desktop")
        )

        $folder = Select-Folder `
            -Description "Select the second GPO backup folder" `
            -InitialFolder $initialFolder `
            -EnablePreviewPane

        if (-not [string]::IsNullOrWhiteSpace($folder)) {
            $textGpo2.Text = $folder
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            $_.Exception.Message,
            "Folder Selection Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
})

$buttonOutput.Add_Click({
    try {
        $initialFolder = Get-FirstExistingFolder -CandidatePath @(
            $textOutput.Text
            [Environment]::GetFolderPath("Desktop")
        )

        $folder = Select-Folder `
            -Description "Select the output folder for CSV and HTML reports" `
            -InitialFolder $initialFolder

        if (-not [string]::IsNullOrWhiteSpace($folder)) {
            $textOutput.Text = $folder
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            $_.Exception.Message,
            "Folder Selection Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
})

$checkOnlyShowDifferences.Add_CheckedChanged({
    if ($script:lastCsvPath) {
        Refresh-GridFromCsv
    }
})

$grid.Add_DataBindingComplete({
    Format-Grid -Grid $grid
    Apply-GridRowColors -Grid $grid
})

$buttonCompare.Add_Click({

    try {
        $buttonCompare.Enabled = $false
        $buttonOpenCsv.Enabled = $false
        $buttonOpenHtml.Enabled = $false

        $grid.DataSource = $null

        $script:lastResult = $null
        $script:lastCsvPath = $null
        $script:lastHtmlPath = $null
        $script:gridDisplayColumns = $null
        $script:visibleRowCount = 0

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

        $script:lastResult = $result
        $script:lastCsvPath = $result.CsvOutputPath
        $script:lastHtmlPath = $result.HtmlOutputPath

        $script:gridDisplayColumns = Get-GridDisplayColumns `
            -Gpo1Name $result.Gpo1Name `
            -Gpo2Name $result.Gpo2Name

        Refresh-GridFromCsv

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
    $checkOnlyShowDifferences,
    $buttonCompare,
    $buttonOpenCsv,
    $buttonOpenHtml,
    $statusLabel,
    $grid,
    $summaryLabel
))

[void]$form.ShowDialog()
