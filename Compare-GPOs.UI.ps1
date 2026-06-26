Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Data
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

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
        OverwritePrompt          = 0x00000002,
        StrictFileTypes          = 0x00000004,
        NoChangeDir              = 0x00000008,
        PickFolders              = 0x00000020,
        ForceFileSystem          = 0x00000040,
        AllNonStorageItems       = 0x00000080,
        NoValidate               = 0x00000100,
        AllowMultiSelect         = 0x00000200,
        PathMustExist            = 0x00000800,
        FileMustExist            = 0x00001000,
        CreatePrompt             = 0x00002000,
        ShareAware               = 0x00004000,
        NoReadOnlyReturn         = 0x00008000,
        NoTestFileCreate         = 0x00010000,
        HideMruPlaces            = 0x00020000,
        HidePinnedPlaces         = 0x00040000,
        NoDereferenceLinks       = 0x00100000,
        OkButtonNeedsInteraction = 0x00200000,
        DontAddToRecent          = 0x02000000,
        ForceShowHidden          = 0x10000000,
        DefaultNoMiniMode        = 0x20000000,
        ForcePreviewPaneOn       = 0x40000000,
        SupportStreamableItems   = 0x80000000
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

function Select-WorkbookFile {
    param(
        [string]$InitialDirectory
    )

    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = "Select Windows11PolicySettings25H2.xlsx"
    $dialog.Filter = "Excel workbook (*.xlsx)|*.xlsx|All files (*.*)|*.*"
    $dialog.CheckFileExists = $true
    $dialog.CheckPathExists = $true
    $dialog.Multiselect = $false

    if (-not [string]::IsNullOrWhiteSpace($InitialDirectory) -and
        (Test-Path -LiteralPath $InitialDirectory -PathType Container)) {
        $dialog.InitialDirectory = $InitialDirectory
    }

    $result = $dialog.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.FileName
    }

    return $null
}

function Get-DefaultPolicyInfoWorkbookPath {
    $candidatePaths = @(
        (Join-Path $ScriptRoot "Windows11PolicySettings25H2.xlsx"),
        (Join-Path $ScriptRoot "Windows11PolicySettings25H2(1).xlsx"),
        (Join-Path (Get-Location).Path "Windows11PolicySettings25H2.xlsx"),
        (Join-Path (Get-Location).Path "Windows11PolicySettings25H2(1).xlsx")
    )

    foreach ($path in $candidatePaths) {
        if (Test-Path -LiteralPath $path -PathType Leaf) {
            return $path
        }
    }

    $wildcardMatches = @(
        Get-ChildItem -LiteralPath $ScriptRoot -Filter "Windows11PolicySettings25H2*.xlsx" -File -ErrorAction SilentlyContinue
    )

    if ($wildcardMatches.Count -gt 0) {
        return $wildcardMatches[0].FullName
    }

    return ""
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

function Normalize-PolicyLookupKey {
    param(
        [AllowNull()]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return ""
    }

    return (($Value -replace '[^A-Za-z0-9]', '').ToLowerInvariant())
}

function Get-ZipEntryText {
    param(
        [Parameter(Mandatory)]
        [System.IO.Compression.ZipArchive]$Zip,

        [Parameter(Mandatory)]
        [string]$EntryName
    )

    $entry = $Zip.GetEntry($EntryName)

    if (-not $entry) {
        return $null
    }

    $stream = $entry.Open()
    $reader = New-Object System.IO.StreamReader($stream)

    try {
        return $reader.ReadToEnd()
    }
    finally {
        $reader.Dispose()
        $stream.Dispose()
    }
}

function Get-OpenXmlNamespaceManager {
    param(
        [Parameter(Mandatory)]
        [xml]$Xml
    )

    $namespaceManager = New-Object System.Xml.XmlNamespaceManager($Xml.NameTable)
    $namespaceManager.AddNamespace("main", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
    $namespaceManager.AddNamespace("rel", "http://schemas.openxmlformats.org/package/2006/relationships")
    $namespaceManager.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

    return $namespaceManager
}

function Convert-ExcelColumnNameToNumber {
    param(
        [Parameter(Mandatory)]
        [string]$ColumnName
    )

    $sum = 0

    foreach ($character in $ColumnName.ToUpperInvariant().ToCharArray()) {
        $sum *= 26
        $sum += ([int][char]$character - [int][char]'A' + 1)
    }

    return $sum
}

function Get-ExcelColumnNumberFromCellReference {
    param(
        [Parameter(Mandatory)]
        [string]$CellReference
    )

    $columnName = ($CellReference -replace '[0-9]', '')
    return Convert-ExcelColumnNameToNumber -ColumnName $columnName
}

function Get-ExcelCellText {
    param(
        [Parameter(Mandatory)]
        [System.Xml.XmlElement]$Cell,

        [Parameter(Mandatory)]
        [string[]]$SharedStrings
    )

    $cellType = [string]$Cell.GetAttribute("t")
    $valueNode = $Cell.SelectSingleNode("./*[local-name()='v']")

    if ($cellType -eq "inlineStr") {
        $textNodes = $Cell.SelectNodes(".//main:t", (Get-OpenXmlNamespaceManager -Xml $Cell.OwnerDocument))
        return (($textNodes | ForEach-Object { $_.InnerText }) -join "")
    }

    if (-not $valueNode) {
        return ""
    }

    $rawValue = [string]$valueNode.InnerText

    if ($cellType -eq "s") {
        $index = 0

        if ([int]::TryParse($rawValue, [ref]$index) -and $index -ge 0 -and $index -lt $SharedStrings.Count) {
            return [string]$SharedStrings[$index]
        }

        return ""
    }

    return $rawValue
}

function Import-XlsxWorksheetRows {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$WorksheetName
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Policy information workbook does not exist: $Path"
    }

    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)

    try {
        $workbookText = Get-ZipEntryText -Zip $zip -EntryName "xl/workbook.xml"
        $relsText = Get-ZipEntryText -Zip $zip -EntryName "xl/_rels/workbook.xml.rels"

        if (-not $workbookText -or -not $relsText) {
            throw "The workbook is missing required Open XML parts."
        }

        [xml]$workbookXml = $workbookText
        [xml]$relsXml = $relsText

        $workbookNs = Get-OpenXmlNamespaceManager -Xml $workbookXml
        $relsNs = Get-OpenXmlNamespaceManager -Xml $relsXml

        $sheetNode = $workbookXml.SelectSingleNode("./*[local-name()='v']")

        if (-not $sheetNode) {
            throw "Worksheet '$WorksheetName' was not found in workbook '$Path'."
        }

        $relationshipId = $sheetNode.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

        if ([string]::IsNullOrWhiteSpace($relationshipId)) {
            throw "Worksheet '$WorksheetName' does not have a relationship ID."
        }

        $relationshipNode = $relsXml.SelectSingleNode("./*[local-name()='v']")

        if (-not $relationshipNode) {
            throw "Could not resolve worksheet relationship '$relationshipId'."
        }

        $target = [string]$relationshipNode.Target

        if ($target.StartsWith("/")) {
            $worksheetEntryName = $target.TrimStart("/")
        }
        else {
            $worksheetEntryName = "xl/$target"
        }

        $worksheetText = Get-ZipEntryText -Zip $zip -EntryName $worksheetEntryName

        if (-not $worksheetText) {
            throw "Could not read worksheet XML part '$worksheetEntryName'."
        }

        $sharedStrings = @()
        $sharedStringsText = Get-ZipEntryText -Zip $zip -EntryName "xl/sharedStrings.xml"

        if ($sharedStringsText) {
            [xml]$sharedStringsXml = $sharedStringsText
            $sharedStringsNs = Get-OpenXmlNamespaceManager -Xml $sharedStringsXml

            $sharedStrings = @(
                $sharedStringsXml.SelectNodes("//main:si", $sharedStringsNs) |
                    ForEach-Object {
                        ($_.SelectNodes(".//main:t", $sharedStringsNs) | ForEach-Object { $_.InnerText }) -join ""
                    }
            )
        }

        [xml]$worksheetXml = $worksheetText
        $worksheetNs = Get-OpenXmlNamespaceManager -Xml $worksheetXml

        $rowNodes = @($worksheetXml.SelectNodes("//main:sheetData/main:row", $worksheetNs))

        if ($rowNodes.Count -lt 2) {
            return @()
        }

        $rowMaps = @{}

        foreach ($rowNode in $rowNodes) {
            $rowNumber = [int]$rowNode.GetAttribute("r")
            $cellMap = @{}

            foreach ($cellNode in @($rowNode.SelectNodes("./main:c", $worksheetNs))) {
                $cellReference = [string]$cellNode.GetAttribute("r")

                if ([string]::IsNullOrWhiteSpace($cellReference)) {
                    continue
                }

                $columnNumber = Get-ExcelColumnNumberFromCellReference -CellReference $cellReference
                $cellMap[$columnNumber] = Get-ExcelCellText -Cell $cellNode -SharedStrings $sharedStrings
            }

            $rowMaps[$rowNumber] = $cellMap
        }

        $headerRowNumber = ($rowMaps.Keys | Sort-Object | Select-Object -First 1)
        $headerMap = $rowMaps[$headerRowNumber]
        $headers = @{}

        foreach ($columnNumber in ($headerMap.Keys | Sort-Object)) {
            $header = [string]$headerMap[$columnNumber]

            if (-not [string]::IsNullOrWhiteSpace($header)) {
                $headers[$columnNumber] = $header.Trim()
            }
        }

        $rows = New-Object System.Collections.Generic.List[object]

        foreach ($rowNumber in ($rowMaps.Keys | Sort-Object)) {
            if ($rowNumber -le $headerRowNumber) {
                continue
            }

            $cellMap = $rowMaps[$rowNumber]
            $object = [ordered]@{}
            $hasValue = $false

            foreach ($columnNumber in ($headers.Keys | Sort-Object)) {
                $header = $headers[$columnNumber]
                $value = ""

                if ($cellMap.ContainsKey($columnNumber)) {
                    $value = [string]$cellMap[$columnNumber]
                }

                if (-not [string]::IsNullOrWhiteSpace($value)) {
                    $hasValue = $true
                }

                $object[$header] = $value
            }

            if ($hasValue) {
                $rows.Add([pscustomobject]$object)
            }
        }

        return @($rows)
    }
    finally {
        $zip.Dispose()
    }
}

function Import-WorksheetRows {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$WorksheetName
    )

    return @(Import-XlsxWorksheetRows -Path $Path -WorksheetName $WorksheetName)
}

function Import-PolicyInfoWorkbook {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $null
    }

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return $null
    }

    $adminRows = @(Import-WorksheetRows -Path $Path -WorksheetName "Administrative Templates")
    $securityRows = @(Import-WorksheetRows -Path $Path -WorksheetName "Security")

    $adminLookup = @{}
    $securityLookup = @{}

    foreach ($row in $adminRows) {
        $name = [string]$row.'Policy Setting Name'
        $key = Normalize-PolicyLookupKey -Value $name

        if ($key -and -not $adminLookup.ContainsKey($key)) {
            $adminLookup[$key] = [pscustomobject]@{
                SourceType          = "Administrative Templates"
                LookupName          = $name
                PolicyPath          = [string]$row.'Policy Path'
                RegistryInformation = [string]$row.'Registry Information'
                HelpText            = [string]$row.'Help Text'
            }
        }
    }

    foreach ($row in $securityRows) {
        $name = [string]$row.'Policy Name'
        $key = Normalize-PolicyLookupKey -Value $name

        if ($key -and -not $securityLookup.ContainsKey($key)) {
            $securityLookup[$key] = [pscustomobject]@{
                SourceType          = "Security"
                LookupName          = $name
                PolicyPath          = [string]$row.'Policy Path'
                RegistryInformation = [string]$row.'Registry Settings'
                HelpText            = [string]$row.'Help Text'
            }
        }
    }

    return [pscustomobject]@{
        Path     = $Path
        Admin    = $adminLookup
        Security = $securityLookup
    }
}

function Get-PolicyInfoForGridRow {
    param(
        [Parameter(Mandatory)]
        [System.Windows.Forms.DataGridViewRow]$Row
    )

    if (-not $script:policyInfoLookup) {
        return $null
    }

    $policyContainer = [string]$Row.Cells["PolicyContainer"].Value
    $settingCategory = [string]$Row.Cells["SettingCategory"].Value
    $settingName = [string]$Row.Cells["SettingName"].Value
    $path = [string]$Row.Cells["Path"].Value

    $keyCandidates = @(
        (Normalize-PolicyLookupKey -Value $settingName),
        (Normalize-PolicyLookupKey -Value ($path -split '\\' | Select-Object -Last 1))
    ) | Where-Object { $_ } | Select-Object -Unique

    if ($policyContainer -eq "Policies" -and $settingCategory -eq "Administrative Templates") {
        foreach ($key in $keyCandidates) {
            if ($script:policyInfoLookup.Admin.ContainsKey($key)) {
                return $script:policyInfoLookup.Admin[$key]
            }
        }
    }

    if ($policyContainer -eq "Policies" -and $settingCategory -eq "Security Settings") {
        foreach ($key in $keyCandidates) {
            if ($script:policyInfoLookup.Security.ContainsKey($key)) {
                return $script:policyInfoLookup.Security[$key]
            }
        }
    }

    return $null
}

function Clear-PolicyInfoPane {
    $textInfoSource.Text = ""
    $textInfoLookupName.Text = ""
    $textInfoPolicyPath.Text = ""
    $textInfoRegistry.Text = ""
    $textInfoHelp.Text = ""
}

function Update-PolicyInfoPaneFromSelection {
    if (-not $policyInfoPanel.Visible) {
        return
    }

    Clear-PolicyInfoPane

    if (-not $script:policyInfoLookup) {
        $textInfoHelp.Text = "Policy information workbook has not been loaded. Select Windows11PolicySettings25H2.xlsx and run the comparison again. The workbook must contain sheets named 'Administrative Templates' and 'Security'."
        return
    }

    if (-not $grid.CurrentRow) {
        return
    }

    $info = Get-PolicyInfoForGridRow -Row $grid.CurrentRow

    if (-not $info) {
        $settingName = [string]$grid.CurrentRow.Cells["SettingName"].Value
        $settingCategory = [string]$grid.CurrentRow.Cells["SettingCategory"].Value

        $textInfoHelp.Text = "No policy information match was found for setting '$settingName' in category '$settingCategory'."
        return
    }

    $textInfoSource.Text = $info.SourceType
    $textInfoLookupName.Text = $info.LookupName
    $textInfoPolicyPath.Text = $info.PolicyPath
    $textInfoRegistry.Text = $info.RegistryInformation
    $textInfoHelp.Text = $info.HelpText
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
        "Property",
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

    return ,$dataTable
}

function Get-SelectedSettingContainerFilter {
    if (-not $comboSettingFilter) {
        return "All settings"
    }

    if ([string]::IsNullOrWhiteSpace([string]$comboSettingFilter.SelectedItem)) {
        return "All settings"
    }

    return [string]$comboSettingFilter.SelectedItem
}

function Get-SelectedPolicyInfoDisplay {
    if (-not $comboPolicyInfoDisplay) {
        return "Hide"
    }

    if ([string]::IsNullOrWhiteSpace([string]$comboPolicyInfoDisplay.SelectedItem)) {
        return "Hide"
    }

    return [string]$comboPolicyInfoDisplay.SelectedItem
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

    $settingFilter = Get-SelectedSettingContainerFilter

    switch ($settingFilter) {
        "Policy settings only" {
            $rows = @($rows | Where-Object { $_.PolicyContainer -eq "Policies" })
        }

        "Preference settings only" {
            $rows = @($rows | Where-Object { $_.PolicyContainer -eq "Preferences" })
        }

        default {
            # All settings
        }
    }

    [System.Data.DataTable]$dataTable = Convert-ObjectsToDataTable `
        -Rows $rows `
        -Columns $script:gridDisplayColumns

    $grid.SuspendLayout()

    try {
        $grid.DataSource = $null
        $grid.DataSource = $dataTable
    }
    finally {
        $grid.ResumeLayout()
    }

    $script:visibleRowCount = @($rows).Count

    Format-Grid -Grid $grid
    Apply-GridRowColors -Grid $grid
    Update-SummaryLabel
    Update-PolicyInfoPaneFromSelection
}

function Format-Grid {
    param(
        [Parameter(Mandatory)]
        [System.Windows.Forms.DataGridView]$Grid
    )

    $Grid.SuspendLayout()

    try {
        $Grid.AutoGenerateColumns = $true
        $Grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
        $Grid.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::None

        $Grid.DefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::False
        $Grid.ColumnHeadersDefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::False

        $Grid.RowTemplate.Height = 24
        $Grid.RowHeadersVisible = $false
        $Grid.AllowUserToResizeColumns = $true
        $Grid.AllowUserToResizeRows = $false
        $Grid.ScrollBars = [System.Windows.Forms.ScrollBars]::Both

        $Grid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
        $Grid.ColumnHeadersHeight = 24

        $Grid.DefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $Grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font(
            "Segoe UI",
            9,
            [System.Drawing.FontStyle]::Bold
        )

        foreach ($column in $Grid.Columns) {
            $column.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
            $column.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
            $column.MinimumWidth = 50

            switch ($column.Name) {
                "PolicyScope" {
                    $column.Width = 90
                }

                "PolicyContainer" {
                    $column.Width = 120
                }

                "SettingCategory" {
                    $column.Width = 160
                }

                "SettingType" {
                    $column.Width = 160
                }

                "SettingName" {
                    $column.Width = 300
                }

                "Path" {
                    $column.Width = 500
                    $column.DefaultCellStyle.Font = New-Object System.Drawing.Font("Consolas", 9)
                }

                "Property" {
                    $column.Width = 130
                }

                "DifferenceType" {
                    $column.Width = 120
                }

                default {
                    if ($column.Name -like "* Value") {
                        $column.Width = 260
                        $column.DefaultCellStyle.Font = New-Object System.Drawing.Font("Consolas", 9)
                    }
                    else {
                        $column.Width = 140
                    }
                }
            }
        }

        foreach ($row in $Grid.Rows) {
            if (-not $row.IsNewRow) {
                $row.Height = 24
            }
        }
    }
    finally {
        $Grid.ResumeLayout()
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

    $Grid.SuspendLayout()

    try {
        foreach ($row in $Grid.Rows) {
            if ($row.IsNewRow) {
                continue
            }

            $row.Height = 24
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
    finally {
        $Grid.ResumeLayout()
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

    $settingFilter = Get-SelectedSettingContainerFilter

    $summaryLabel.Text = "Compared '$($script:lastResult.Gpo1Name)' to '$($script:lastResult.Gpo2Name)' | Filter: $settingFilter | Total: $totalCount | Visible: $visibleCount | Added: $addedCount | Removed: $removedCount | Changed: $changedCount | Same: $sameCount"
}

function Update-DataPaneLayout {
    $availableWidth = [Math]::Max(300, $form.ClientSize.Width - 30)
    $summaryHeight = 25
    $summaryY = $form.ClientSize.Height - 35
    $policyInfoHeight = 190
    $gap = 6

    $summaryLabel.Location = New-Object System.Drawing.Point(15, $summaryY)
    $summaryLabel.Size = New-Object System.Drawing.Size($availableWidth, $summaryHeight)

    if ($policyInfoPanel.Visible) {
        $policyInfoY = $summaryY - $policyInfoHeight - $gap

        $policyInfoPanel.Location = New-Object System.Drawing.Point(15, $policyInfoY)
        $policyInfoPanel.Size = New-Object System.Drawing.Size($availableWidth, $policyInfoHeight)

        $gridHeight = [Math]::Max(160, $policyInfoY - $grid.Top - $gap)
        $grid.Size = New-Object System.Drawing.Size($availableWidth, $gridHeight)
    }
    else {
        $gridHeight = [Math]::Max(200, $summaryY - $grid.Top - $gap)
        $grid.Size = New-Object System.Drawing.Size($availableWidth, $gridHeight)
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Group Policy Comparator"
$form.Size = New-Object System.Drawing.Size(1200, 820)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(1000, 700)

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

$labelPolicyInfoWorkbook = New-Object System.Windows.Forms.Label
$labelPolicyInfoWorkbook.Text = "Policy Info Workbook"
$labelPolicyInfoWorkbook.Location = New-Object System.Drawing.Point(15, 140)
$labelPolicyInfoWorkbook.Size = New-Object System.Drawing.Size(160, 22)

$textPolicyInfoWorkbook = New-Object System.Windows.Forms.TextBox
$textPolicyInfoWorkbook.Location = New-Object System.Drawing.Point(180, 138)
$textPolicyInfoWorkbook.Size = New-Object System.Drawing.Size(870, 24)
$textPolicyInfoWorkbook.Anchor = "Top,Left,Right"
$textPolicyInfoWorkbook.Text = Get-DefaultPolicyInfoWorkbookPath

$buttonPolicyInfoWorkbook = New-Object System.Windows.Forms.Button
$buttonPolicyInfoWorkbook.Text = "Browse..."
$buttonPolicyInfoWorkbook.Location = New-Object System.Drawing.Point(1060, 136)
$buttonPolicyInfoWorkbook.Size = New-Object System.Drawing.Size(110, 28)
$buttonPolicyInfoWorkbook.Anchor = "Top,Right"

$checkIncludeSame = New-Object System.Windows.Forms.CheckBox
$checkIncludeSame.Text = "Include equal values in generated reports"
$checkIncludeSame.Location = New-Object System.Drawing.Point(180, 175)
$checkIncludeSame.Size = New-Object System.Drawing.Size(280, 24)
$checkIncludeSame.Checked = $true

$checkOnlyShowDifferences = New-Object System.Windows.Forms.CheckBox
$checkOnlyShowDifferences.Text = "Only show differences"
$checkOnlyShowDifferences.Location = New-Object System.Drawing.Point(470, 175)
$checkOnlyShowDifferences.Size = New-Object System.Drawing.Size(160, 24)
$checkOnlyShowDifferences.Checked = $false

$labelSettingFilter = New-Object System.Windows.Forms.Label
$labelSettingFilter.Text = "Show:"
$labelSettingFilter.Location = New-Object System.Drawing.Point(645, 178)
$labelSettingFilter.Size = New-Object System.Drawing.Size(45, 20)

$comboSettingFilter = New-Object System.Windows.Forms.ComboBox
$comboSettingFilter.Location = New-Object System.Drawing.Point(695, 174)
$comboSettingFilter.Size = New-Object System.Drawing.Size(190, 24)
$comboSettingFilter.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
[void]$comboSettingFilter.Items.Add("All settings")
[void]$comboSettingFilter.Items.Add("Policy settings only")
[void]$comboSettingFilter.Items.Add("Preference settings only")
$comboSettingFilter.SelectedIndex = 0

$labelPolicyInfoDisplay = New-Object System.Windows.Forms.Label
$labelPolicyInfoDisplay.Text = "Policy info:"
$labelPolicyInfoDisplay.Location = New-Object System.Drawing.Point(900, 178)
$labelPolicyInfoDisplay.Size = New-Object System.Drawing.Size(75, 20)

$comboPolicyInfoDisplay = New-Object System.Windows.Forms.ComboBox
$comboPolicyInfoDisplay.Location = New-Object System.Drawing.Point(980, 174)
$comboPolicyInfoDisplay.Size = New-Object System.Drawing.Size(150, 24)
$comboPolicyInfoDisplay.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$comboPolicyInfoDisplay.Anchor = "Top,Right"

[void]$comboPolicyInfoDisplay.Items.Add("Hide")
[void]$comboPolicyInfoDisplay.Items.Add("Show")
$comboPolicyInfoDisplay.SelectedIndex = 0

$buttonCompare = New-Object System.Windows.Forms.Button
$buttonCompare.Text = "Compare GPOs"
$buttonCompare.Location = New-Object System.Drawing.Point(15, 210)
$buttonCompare.Size = New-Object System.Drawing.Size(145, 34)

$buttonOpenCsv = New-Object System.Windows.Forms.Button
$buttonOpenCsv.Text = "Open CSV"
$buttonOpenCsv.Location = New-Object System.Drawing.Point(170, 210)
$buttonOpenCsv.Size = New-Object System.Drawing.Size(110, 34)
$buttonOpenCsv.Enabled = $false

$buttonOpenHtml = New-Object System.Windows.Forms.Button
$buttonOpenHtml.Text = "Open HTML"
$buttonOpenHtml.Location = New-Object System.Drawing.Point(290, 210)
$buttonOpenHtml.Size = New-Object System.Drawing.Size(110, 34)
$buttonOpenHtml.Enabled = $false

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Select two GPO backup folders to begin."
$statusLabel.Location = New-Object System.Drawing.Point(420, 218)
$statusLabel.Size = New-Object System.Drawing.Size(750, 22)
$statusLabel.Anchor = "Top,Left,Right"

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(15, 260)
$grid.Size = New-Object System.Drawing.Size(
    ($form.ClientSize.Width - 30),
    ($form.ClientSize.Height - 310)
)
$grid.Anchor = "Top,Bottom,Left,Right"
$grid.ReadOnly = $true
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.SelectionMode = "FullRowSelect"
$grid.MultiSelect = $true
$grid.AutoGenerateColumns = $true
$grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
$grid.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::None
$grid.DefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::False
$grid.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
$grid.BackgroundColor = [System.Drawing.Color]::White
$grid.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$grid.ClipboardCopyMode = [System.Windows.Forms.DataGridViewClipboardCopyMode]::EnableAlwaysIncludeHeaderText

$policyInfoPanel = New-Object System.Windows.Forms.Panel
$policyInfoPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$policyInfoPanel.Visible = $false
$policyInfoPanel.Anchor = "Left,Right,Bottom"

$labelInfoSource = New-Object System.Windows.Forms.Label
$labelInfoSource.Text = "Source"
$labelInfoSource.Location = New-Object System.Drawing.Point(8, 8)
$labelInfoSource.Size = New-Object System.Drawing.Size(90, 20)

$textInfoSource = New-Object System.Windows.Forms.TextBox
$textInfoSource.Location = New-Object System.Drawing.Point(105, 6)
$textInfoSource.Size = New-Object System.Drawing.Size(180, 22)
$textInfoSource.ReadOnly = $true

$labelInfoLookupName = New-Object System.Windows.Forms.Label
$labelInfoLookupName.Text = "Matched Name"
$labelInfoLookupName.Location = New-Object System.Drawing.Point(300, 8)
$labelInfoLookupName.Size = New-Object System.Drawing.Size(95, 20)

$textInfoLookupName = New-Object System.Windows.Forms.TextBox
$textInfoLookupName.Location = New-Object System.Drawing.Point(400, 6)
$textInfoLookupName.Size = New-Object System.Drawing.Size(745, 22)
$textInfoLookupName.Anchor = "Top,Left,Right"
$textInfoLookupName.ReadOnly = $true

$labelInfoPolicyPath = New-Object System.Windows.Forms.Label
$labelInfoPolicyPath.Text = "Policy Path"
$labelInfoPolicyPath.Location = New-Object System.Drawing.Point(8, 38)
$labelInfoPolicyPath.Size = New-Object System.Drawing.Size(90, 20)

$textInfoPolicyPath = New-Object System.Windows.Forms.TextBox
$textInfoPolicyPath.Location = New-Object System.Drawing.Point(105, 36)
$textInfoPolicyPath.Size = New-Object System.Drawing.Size(1040, 22)
$textInfoPolicyPath.Anchor = "Top,Left,Right"
$textInfoPolicyPath.ReadOnly = $true

$labelInfoRegistry = New-Object System.Windows.Forms.Label
$labelInfoRegistry.Text = "Registry"
$labelInfoRegistry.Location = New-Object System.Drawing.Point(8, 68)
$labelInfoRegistry.Size = New-Object System.Drawing.Size(90, 20)

$textInfoRegistry = New-Object System.Windows.Forms.TextBox
$textInfoRegistry.Location = New-Object System.Drawing.Point(105, 66)
$textInfoRegistry.Size = New-Object System.Drawing.Size(1040, 44)
$textInfoRegistry.Anchor = "Top,Left,Right"
$textInfoRegistry.Multiline = $true
$textInfoRegistry.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$textInfoRegistry.ReadOnly = $true

$labelInfoHelp = New-Object System.Windows.Forms.Label
$labelInfoHelp.Text = "Help Text"
$labelInfoHelp.Location = New-Object System.Drawing.Point(8, 120)
$labelInfoHelp.Size = New-Object System.Drawing.Size(90, 20)

$textInfoHelp = New-Object System.Windows.Forms.TextBox
$textInfoHelp.Location = New-Object System.Drawing.Point(105, 118)
$textInfoHelp.Size = New-Object System.Drawing.Size(1040, 62)
$textInfoHelp.Anchor = "Top,Bottom,Left,Right"
$textInfoHelp.Multiline = $true
$textInfoHelp.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$textInfoHelp.ReadOnly = $true

$policyInfoPanel.Controls.AddRange(@(
    $labelInfoSource,
    $textInfoSource,
    $labelInfoLookupName,
    $textInfoLookupName,
    $labelInfoPolicyPath,
    $textInfoPolicyPath,
    $labelInfoRegistry,
    $textInfoRegistry,
    $labelInfoHelp,
    $textInfoHelp
))

$summaryLabel = New-Object System.Windows.Forms.Label
$summaryLabel.Text = ""
$summaryLabel.Location = New-Object System.Drawing.Point(15, ($form.ClientSize.Height - 35))
$summaryLabel.Size = New-Object System.Drawing.Size(($form.ClientSize.Width - 30), 25)
$summaryLabel.Anchor = "Bottom,Left,Right"

$form.Add_Resize({
    Update-DataPaneLayout
})

$script:lastCsvPath = $null
$script:lastHtmlPath = $null
$script:lastResult = $null
$script:gridDisplayColumns = $null
$script:visibleRowCount = 0
$script:policyInfoLookup = $null

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

$buttonPolicyInfoWorkbook.Add_Click({
    try {
        $initialFolder = Get-FirstExistingFolder -CandidatePath @(
            (Split-Path -Parent $textPolicyInfoWorkbook.Text)
            $ScriptRoot
            [Environment]::GetFolderPath("Desktop")
        )

        $workbookPath = Select-WorkbookFile -InitialDirectory $initialFolder

        if (-not [string]::IsNullOrWhiteSpace($workbookPath)) {
            $textPolicyInfoWorkbook.Text = $workbookPath
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            $_.Exception.Message,
            "Workbook Selection Error",
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

$comboSettingFilter.Add_SelectedIndexChanged({
    if ($script:lastCsvPath) {
        Refresh-GridFromCsv
    }
})

$comboPolicyInfoDisplay.Add_SelectedIndexChanged({
    $policyInfoPanel.Visible = ((Get-SelectedPolicyInfoDisplay) -eq "Show")
    Update-DataPaneLayout

    if ((Get-SelectedPolicyInfoDisplay) -eq "Show") {
        try {
            if (-not $script:policyInfoLookup) {
                if (-not [string]::IsNullOrWhiteSpace($textPolicyInfoWorkbook.Text) -and
                    (Test-Path -LiteralPath $textPolicyInfoWorkbook.Text -PathType Leaf)) {

                    $script:policyInfoLookup = Import-PolicyInfoWorkbook -Path $textPolicyInfoWorkbook.Text
                }
            }

            Update-PolicyInfoPaneFromSelection
        }
        catch {
            Clear-PolicyInfoPane
            $textInfoHelp.Text = "Unable to load policy information workbook: $($_.Exception.Message)"
        }
    }
    else {
        Clear-PolicyInfoPane
    }
})

$grid.Add_DataBindingComplete({
    Format-Grid -Grid $grid
    Apply-GridRowColors -Grid $grid
})

$grid.Add_SelectionChanged({
    Update-PolicyInfoPaneFromSelection
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
        $script:policyInfoLookup = $null

        Clear-PolicyInfoPane
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

		if ((Get-SelectedPolicyInfoDisplay) -eq "Show") {
			if ([string]::IsNullOrWhiteSpace($textPolicyInfoWorkbook.Text)) {
				throw "Policy info is set to Show, but no policy information workbook was selected."
			}

			if (-not (Test-Path -LiteralPath $textPolicyInfoWorkbook.Text -PathType Leaf)) {
				throw "Policy info is set to Show, but the selected workbook does not exist: $($textPolicyInfoWorkbook.Text)"
			}

			$statusLabel.Text = "Loading policy information workbook..."
			$script:policyInfoLookup = Import-PolicyInfoWorkbook -Path $textPolicyInfoWorkbook.Text
		}
		else {
			$script:policyInfoLookup = $null
		}

        $statusLabel.Text = "Running comparison..."

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
    $labelPolicyInfoWorkbook,
    $textPolicyInfoWorkbook,
    $buttonPolicyInfoWorkbook,
    $checkIncludeSame,
    $checkOnlyShowDifferences,
    $labelSettingFilter,
    $comboSettingFilter,
    $labelPolicyInfoDisplay,
    $comboPolicyInfoDisplay,
    $buttonCompare,
    $buttonOpenCsv,
    $buttonOpenHtml,
    $statusLabel,
    $grid,
    $policyInfoPanel,
    $summaryLabel
))

Update-DataPaneLayout

[void]$form.ShowDialog()
