<#
.SYNOPSIS
Compares two GPO backup folders by reading their gpreport.xml files and exports CSV and HTML difference reports.

.DESCRIPTION
This script contains the core comparison engine for Group Policy backup comparison.

It can be dot-sourced from a UI script:

    . .\Compare-GPOs.ps1

Then called with:

    Invoke-GpoComparison `
        -Gpo1BackupFolder "C:\GPOBackups\GPO1" `
        -Gpo2BackupFolder "C:\GPOBackups\GPO2" `
        -OutputFolder "C:\Temp\GPO_Comparison" `
        -IncludeSame

.NOTES
Expected input:
- Each selected GPO backup folder should contain, or contain beneath it, a gpreport.xml file.
#>

function Get-PolicyScopeSortOrder {
    param([string]$PolicyScope)

    switch ($PolicyScope) {
        "Computer" { 1 }
        "User"     { 2 }
        default    { 99 }
    }
}

function Get-PolicyContainerSortOrder {
    param([string]$PolicyContainer)

    switch ($PolicyContainer) {
        "Policies"    { 1 }
        "Preferences" { 2 }
        default       { 99 }
    }
}

function Get-SettingCategorySortOrder {
    param(
        [string]$PolicyContainer,
        [string]$SettingCategory
    )

    if ($PolicyContainer -eq "Policies") {
        switch ($SettingCategory) {
            "Scripts"                  { 1 }
            "Security Settings"        { 2 }
            "Administrative Templates" { 3 }
            default                    { 99 }
        }
    }
    elseif ($PolicyContainer -eq "Preferences") {
        switch ($SettingCategory) {
            "Registry Settings" { 1 }
            default             { 99 }
        }
    }
    else {
        99
    }
}

function Get-SettingTypeSortOrder {
    param(
        [string]$SettingCategory,
        [string]$SettingType
    )

    if ($SettingCategory -eq "Security Settings") {
        switch ($SettingType) {
            "Account policies"             { 1 }
            "Audit policy"                 { 2 }
            "Advanced Audit Configuration" { 3 }
            "User Rights"                  { 4 }
            "Security options"             { 5 }
            "Event Log"                    { 6 }
            "Restricted Groups"            { 7 }
            default                        { 99 }
        }
    }
    else {
        99
    }
}

function Get-FirstSettingValue {
    param(
        $Object
    )

    if ($null -ne $Object.SettingString -and $Object.SettingString -ne "") {
        return $Object.SettingString
    }

    if ($null -ne $Object.SettingNumber -and $Object.SettingNumber -ne "") {
        return $Object.SettingNumber
    }

    if ($null -ne $Object.SettingBoolean -and $Object.SettingBoolean -ne "") {
        return $Object.SettingBoolean
    }

    return $null
}

function ConvertTo-SemicolonList {
    param(
        [AllowNull()]
        $InputObject,

        [Parameter(Mandatory)]
        [scriptblock]$ValueScript
    )

    @(
        $InputObject | ForEach-Object {
            & $ValueScript $_
        } | Where-Object {
            $_
        } | Sort-Object
    ) -join ';'
}

function ConvertTo-AdvancedAuditDisplayValue {
    param(
        [AllowNull()]
        [object]$SettingValue
    )

    switch ("$SettingValue") {
        "0" { "No Auditing" }
        "1" { "Success" }
        "2" { "Failure" }
        "3" { "Success and Failure" }
        default { $SettingValue }
    }
}

function Resolve-GpoReportXmlPath {
    param(
        [Parameter(Mandatory)]
        [string]$FolderPath
    )

    if (-not (Test-Path -LiteralPath $FolderPath -PathType Container)) {
        throw "Folder does not exist: $FolderPath"
    }

    $directPath = Join-Path $FolderPath "gpreport.xml"

    if (Test-Path -LiteralPath $directPath -PathType Leaf) {
        return $directPath
    }

    $matches = @(Get-ChildItem -LiteralPath $FolderPath -Recurse -Filter "gpreport.xml" -File -ErrorAction SilentlyContinue)

    if ($matches.Count -eq 0) {
        throw "No gpreport.xml file found under: $FolderPath"
    }

    if ($matches.Count -gt 1) {
        throw "Multiple gpreport.xml files found under: $FolderPath. Select the specific GPO backup folder that contains the gpreport.xml file."
    }

    return $matches[0].FullName
}

function Get-GpoExtractedObjects {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    [xml]$GPOXMLData = Get-Content -LiteralPath $Path

    $GPOName = $GPOXMLData.GPO.Name

    $AllScripts = @()

    $AllAccountPolicies = @()
    $AllAuditPolicies = @()
    $AllAdvancedAuditSettings = @()
    $AllUserRights = @()
    $AllSecurityOptions = @()
    $AllEventLogSettings = @()
    $AllRestrictedGroups = @()

    $AllPolicySettings = @()
    $AllRegistrySettings = @()

    foreach ($PolicyScopeName in @("Computer", "User")) {

        $PolicyScopeNode = $GPOXMLData.GPO.$PolicyScopeName

        foreach ($extension in @($PolicyScopeNode.ExtensionData)) {

            switch ($extension.Name) {

                "Scripts" {

                    $Scripts = foreach ($script in @($extension.Extension.Script)) {
                        [pscustomobject]@{
                            PolicyScope     = $PolicyScopeName
                            PolicyContainer = "Policies"
                            SettingCategory = "Scripts"
                            SettingType     = $script.Type
                            Name            = $script.Command
                            Command         = $script.Command
                            Type            = $script.Type
                            Order           = $script.Order
                            RunOrder        = $script.RunOrder
                        }
                    }

                    $AllScripts += $Scripts
                }

                "Security" {

                    $AccountPolicies = foreach ($account in @($extension.Extension.Account)) {
                        [pscustomobject]@{
                            PolicyScope     = $PolicyScopeName
                            PolicyContainer = "Policies"
                            SettingCategory = "Security Settings"
                            SettingType     = "Account policies"
                            Name            = $account.Name
                            Type            = $account.Type
                            SettingValue    = Get-FirstSettingValue -Object $account
                        }
                    }

                    $AuditPolicies = foreach ($audit in @($extension.Extension.Audit)) {
                        [pscustomobject]@{
                            PolicyScope     = $PolicyScopeName
                            PolicyContainer = "Policies"
                            SettingCategory = "Security Settings"
                            SettingType     = "Audit policy"
                            Name            = $audit.Name
                            SuccessAttempts = $audit.SuccessAttempts
                            FailureAttempts = $audit.FailureAttempts
                        }
                    }

                    $UserRights = foreach ($right in @($extension.Extension.UserRightsAssignment)) {
                        [pscustomobject]@{
                            PolicyScope     = $PolicyScopeName
                            PolicyContainer = "Policies"
                            SettingCategory = "Security Settings"
                            SettingType     = "User Rights"
                            Name            = $right.Name
                            Member          = ConvertTo-SemicolonList -InputObject $right.Member -ValueScript {
                                param($x)
                                $x.SID.'#text'
                            }
                        }
                    }

                    $SecurityOptions = foreach ($setting in @($extension.Extension.SecurityOptions)) {

                        $display = $setting.Display

                        $displayValue = if ($display.DisplayString) {
                            $display.DisplayString
                        }
                        elseif ($null -ne $display.DisplayNumber -and $display.DisplayNumber -ne "") {
                            $display.DisplayNumber
                        }
                        elseif ($null -ne $display.DisplayBoolean -and $display.DisplayBoolean -ne "") {
                            $display.DisplayBoolean
                        }
                        elseif ($display.DisplayFields) {
                            @(
                                $display.DisplayFields.Field | ForEach-Object {
                                    "$($_.Name):$($_.Value)"
                                }
                            ) -join '; '
                        }
                        else {
                            $null
                        }

                        $friendlyName = if ($display.Name) {
                            $display.Name
                        }
                        elseif ($setting.KeyName) {
                            Split-Path -Path $setting.KeyName -Leaf
                        }
                        elseif ($setting.SystemAccessPolicyName) {
                            $setting.SystemAccessPolicyName
                        }
                        else {
                            $null
                        }

                        [pscustomobject]@{
                            PolicyScope            = $PolicyScopeName
                            PolicyContainer        = "Policies"
                            SettingCategory        = "Security Settings"
                            SettingType            = "Security options"
                            Name                   = $friendlyName
                            Units                  = $display.Units
                            DisplayValue           = $displayValue
                            KeyName                = $setting.KeyName
                            SettingValue           = Get-FirstSettingValue -Object $setting
                            SystemAccessPolicyName = $setting.SystemAccessPolicyName
                        }
                    }

                    $EventLogSettings = foreach ($eventLog in @($extension.Extension.EventLog)) {
                        [pscustomobject]@{
                            PolicyScope     = $PolicyScopeName
                            PolicyContainer = "Policies"
                            SettingCategory = "Security Settings"
                            SettingType     = "Event Log"
                            Name            = $eventLog.Name
                            Log             = $eventLog.Log
                            SettingValue    = Get-FirstSettingValue -Object $eventLog
                        }
                    }

                    $RestrictedGroups = foreach ($group in @($extension.Extension.RestrictedGroups)) {

                        $groupName = $group.GroupName.Name.'#text'

                        $members = ConvertTo-SemicolonList -InputObject $group.Member -ValueScript {
                            param($x)
                            $x.Name.'#text'
                        }

                        $memberOf = ConvertTo-SemicolonList -InputObject $group.Memberof -ValueScript {
                            param($x)
                            $x.Name.'#text'
                        }

                        [pscustomobject]@{
                            PolicyScope     = $PolicyScopeName
                            PolicyContainer = "Policies"
                            SettingCategory = "Security Settings"
                            SettingType     = "Restricted Groups"
                            Name            = $groupName
                            Member          = $members
                            MemberOf        = $memberOf
                        }
                    }

                    $AllAccountPolicies += $AccountPolicies
                    $AllAuditPolicies += $AuditPolicies
                    $AllUserRights += $UserRights
                    $AllSecurityOptions += $SecurityOptions
                    $AllEventLogSettings += $EventLogSettings
                    $AllRestrictedGroups += $RestrictedGroups
                }

                "Advanced Audit Configuration" {

                    $AdvancedAuditSettings = foreach ($auditSetting in @($extension.Extension.AuditSetting)) {
                        [pscustomobject]@{
                            PolicyScope     = $PolicyScopeName
                            PolicyContainer = "Policies"
                            SettingCategory = "Security Settings"
                            SettingType     = "Advanced Audit Configuration"
                            Name            = $auditSetting.SubcategoryName
                            PolicyTarget    = $auditSetting.PolicyTarget
                            SubcategoryName = $auditSetting.SubcategoryName
                            SubcategoryGuid = $auditSetting.SubcategoryGuid
                            SettingValue    = $auditSetting.SettingValue
                            DisplayValue    = ConvertTo-AdvancedAuditDisplayValue -SettingValue $auditSetting.SettingValue
                        }
                    }

                    $AllAdvancedAuditSettings += $AdvancedAuditSettings
                }

                "Registry" {

                    $PolicySettings = foreach ($policy in @($extension.Extension.Policy)) {
                        [pscustomobject]@{
                            PolicyScope     = $PolicyScopeName
                            PolicyContainer = "Policies"
                            SettingCategory = "Administrative Templates"
                            SettingType     = "Registry settings"
                            Name            = $policy.Name
                            State           = $policy.State
                            Category        = $policy.Category
                            Supported       = $policy.Supported
                            Explain         = ($policy.Explain -replace '\s+', ' ').Trim()
                        }
                    }

                    $AllPolicySettings += $PolicySettings
                }

                "Windows Registry" {

                    $RegistrySettings = foreach ($registry in @($extension.Extension.RegistrySettings.Registry)) {
                        [pscustomobject]@{
                            PolicyScope     = $PolicyScopeName
                            PolicyContainer = "Preferences"
                            SettingCategory = "Registry Settings"
                            SettingType     = "Registry settings"
                            Name            = $registry.name
                            GPOSettingOrder = $registry.GPOSettingOrder
                            RemovePolicy    = $registry.removePolicy
                            Action          = $registry.Properties.action
                            Hive            = $registry.Properties.hive
                            Key             = $registry.Properties.key
                            ValueName       = $registry.Properties.name
                            Type            = $registry.Properties.type
                            Value           = $registry.Properties.value
                        }
                    }

                    $AllRegistrySettings += $RegistrySettings
                }
            }
        }
    }

    [pscustomobject]@{
        Name                  = $GPOName
        Scripts               = $AllScripts
        AccountPolicies       = $AllAccountPolicies
        AuditPolicies         = $AllAuditPolicies
        AdvancedAuditSettings = $AllAdvancedAuditSettings
        UserRights            = $AllUserRights
        SecurityOptions       = $AllSecurityOptions
        EventLogSettings      = $AllEventLogSettings
        RestrictedGroups      = $AllRestrictedGroups
        PolicySettings        = $AllPolicySettings
        RegistrySettings      = $AllRegistrySettings
    }
}

function Compare-ObjectSet {
    param(
        [AllowNull()]
        [AllowEmptyCollection()]
        [object[]]$ReferenceObject,

        [AllowNull()]
        [AllowEmptyCollection()]
        [object[]]$DifferenceObject,

        [Parameter(Mandatory)]
        [scriptblock]$KeyScript,

        [Parameter(Mandatory)]
        [string[]]$CompareProperties,

        [Parameter(Mandatory)]
        [string]$ObjectType,

        [Parameter(Mandatory)]
        [string]$Gpo1ValueColumn,

        [Parameter(Mandatory)]
        [string]$Gpo2ValueColumn,

        [switch]$IncludeSame
    )

    $ReferenceObject  = @($ReferenceObject)
    $DifferenceObject = @($DifferenceObject)

    $refHash = @{}
    $difHash = @{}

    foreach ($item in $ReferenceObject) {
        $key = & $KeyScript $item
        if ($key) {
            $refHash[$key] = $item
        }
    }

    foreach ($item in $DifferenceObject) {
        $key = & $KeyScript $item
        if ($key) {
            $difHash[$key] = $item
        }
    }

    $allKeys = @($refHash.Keys + $difHash.Keys) |
        Where-Object { $_ } |
        Sort-Object -Unique

    foreach ($key in $allKeys) {

        $ref = $refHash[$key]
        $dif = $difHash[$key]

        $metadataSource = if ($dif) { $dif } else { $ref }

        $settingName = if ($metadataSource.Name) {
            $metadataSource.Name
        }
        elseif ($metadataSource.ValueName) {
            $metadataSource.ValueName
        }
        else {
            $null
        }

        $displayValue = if ($metadataSource.DisplayValue) {
            $metadataSource.DisplayValue
        }
        else {
            $null
        }

        if (-not $ref) {

            foreach ($property in $CompareProperties) {

                $row = [ordered]@{
                    PolicyScope              = $dif.PolicyScope
                    PolicyContainer          = $dif.PolicyContainer
                    SettingCategory          = $dif.SettingCategory
                    SettingType              = $dif.SettingType
                    SettingName              = $settingName
                    Path                     = $key
                    DisplayValue             = $displayValue
                    Property                 = $property
                    DifferenceType           = "Added"
                    PolicyScopeSortOrder     = Get-PolicyScopeSortOrder -PolicyScope $dif.PolicyScope
                    PolicyContainerSortOrder = Get-PolicyContainerSortOrder -PolicyContainer $dif.PolicyContainer
                    SettingCategorySortOrder = Get-SettingCategorySortOrder -PolicyContainer $dif.PolicyContainer -SettingCategory $dif.SettingCategory
                    SettingTypeSortOrder     = Get-SettingTypeSortOrder -SettingCategory $dif.SettingCategory -SettingType $dif.SettingType
                }

                $row[$Gpo1ValueColumn] = $null
                $row[$Gpo2ValueColumn] = $dif.$property
                $row["ObjectType"]     = $ObjectType

                [pscustomobject]$row
            }

            continue
        }

        if (-not $dif) {

            foreach ($property in $CompareProperties) {

                $row = [ordered]@{
                    PolicyScope              = $ref.PolicyScope
                    PolicyContainer          = $ref.PolicyContainer
                    SettingCategory          = $ref.SettingCategory
                    SettingType              = $ref.SettingType
                    SettingName              = $settingName
                    Path                     = $key
                    DisplayValue             = $displayValue
                    Property                 = $property
                    DifferenceType           = "Removed"
                    PolicyScopeSortOrder     = Get-PolicyScopeSortOrder -PolicyScope $ref.PolicyScope
                    PolicyContainerSortOrder = Get-PolicyContainerSortOrder -PolicyContainer $ref.PolicyContainer
                    SettingCategorySortOrder = Get-SettingCategorySortOrder -PolicyContainer $ref.PolicyContainer -SettingCategory $ref.SettingCategory
                    SettingTypeSortOrder     = Get-SettingTypeSortOrder -SettingCategory $ref.SettingCategory -SettingType $ref.SettingType
                }

                $row[$Gpo1ValueColumn] = $ref.$property
                $row[$Gpo2ValueColumn] = $null
                $row["ObjectType"]     = $ObjectType

                [pscustomobject]$row
            }

            continue
        }

        foreach ($property in $CompareProperties) {

            $refValue = $ref.$property
            $difValue = $dif.$property
            $reportSource = if ($ref) { $ref } else { $dif }

            if ("$refValue" -ne "$difValue") {

                $row = [ordered]@{
                    PolicyScope              = $reportSource.PolicyScope
                    PolicyContainer          = $reportSource.PolicyContainer
                    SettingCategory          = $reportSource.SettingCategory
                    SettingType              = $reportSource.SettingType
                    SettingName              = $settingName
                    Path                     = $key
                    DisplayValue             = $displayValue
                    Property                 = $property
                    DifferenceType           = "Changed"
                    PolicyScopeSortOrder     = Get-PolicyScopeSortOrder -PolicyScope $reportSource.PolicyScope
                    PolicyContainerSortOrder = Get-PolicyContainerSortOrder -PolicyContainer $reportSource.PolicyContainer
                    SettingCategorySortOrder = Get-SettingCategorySortOrder -PolicyContainer $reportSource.PolicyContainer -SettingCategory $reportSource.SettingCategory
                    SettingTypeSortOrder     = Get-SettingTypeSortOrder -SettingCategory $reportSource.SettingCategory -SettingType $reportSource.SettingType
                }

                $row[$Gpo1ValueColumn] = $refValue
                $row[$Gpo2ValueColumn] = $difValue
                $row["ObjectType"]     = $ObjectType

                [pscustomobject]$row
            }
            elseif ($IncludeSame) {

                $row = [ordered]@{
                    PolicyScope              = $reportSource.PolicyScope
                    PolicyContainer          = $reportSource.PolicyContainer
                    SettingCategory          = $reportSource.SettingCategory
                    SettingType              = $reportSource.SettingType
                    SettingName              = $settingName
                    Path                     = $key
                    DisplayValue             = $displayValue
                    Property                 = $property
                    DifferenceType           = "Same"
                    PolicyScopeSortOrder     = Get-PolicyScopeSortOrder -PolicyScope $reportSource.PolicyScope
                    PolicyContainerSortOrder = Get-PolicyContainerSortOrder -PolicyContainer $reportSource.PolicyContainer
                    SettingCategorySortOrder = Get-SettingCategorySortOrder -PolicyContainer $reportSource.PolicyContainer -SettingCategory $reportSource.SettingCategory
                    SettingTypeSortOrder     = Get-SettingTypeSortOrder -SettingCategory $reportSource.SettingCategory -SettingType $reportSource.SettingType
                }

                $row[$Gpo1ValueColumn] = $refValue
                $row[$Gpo2ValueColumn] = $difValue
                $row["ObjectType"]     = $ObjectType

                [pscustomobject]$row
            }
        }
    }
}

function ConvertTo-HtmlEncodedText {
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return ""
    }

    [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Get-DifferenceCssClass {
    param(
        [string]$DifferenceType
    )

    switch ($DifferenceType) {
        "Added"   { "diff-added" }
        "Removed" { "diff-removed" }
        "Changed" { "diff-changed" }
        "Same"    { "diff-same" }
        default   { "diff-same" }
    }
}

function Get-DifferencePrefix {
    param(
        [string]$DifferenceType
    )

    switch ($DifferenceType) {
        "Added"   { "[+] " }
        "Removed" { "[-] " }
        "Changed" { "[#] " }
        "Same"    { "" }
        default   { "" }
    }
}

function Export-GpoDifferenceHtml {
    param(
        [Parameter(Mandatory)]
        [object[]]$Differences,

        [Parameter(Mandatory)]
        [string[]]$Columns,

        [Parameter(Mandatory)]
        [string]$Gpo1Name,

        [Parameter(Mandatory)]
        [string]$Gpo2Name,

        [Parameter(Mandatory)]
        [string]$Path
    )

    $generatedOn = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    $totalCount   = @($Differences).Count
    $addedCount   = @($Differences | Where-Object DifferenceType -eq "Added").Count
    $removedCount = @($Differences | Where-Object DifferenceType -eq "Removed").Count
    $changedCount = @($Differences | Where-Object DifferenceType -eq "Changed").Count
    $sameCount    = @($Differences | Where-Object DifferenceType -eq "Same").Count

    $html = New-Object System.Text.StringBuilder

    [void]$html.AppendLine("<!DOCTYPE html>")
    [void]$html.AppendLine("<html lang='en-US'>")
    [void]$html.AppendLine("<head>")
    [void]$html.AppendLine("<meta charset='utf-8'>")
    [void]$html.AppendLine("<title>GPO Difference Report</title>")
    [void]$html.AppendLine("<style>")
    [void]$html.AppendLine(@"
body {
    background-color: #ffffff;
    color: #000000;
    font-family: 'Segoe UI', Arial, sans-serif;
    font-size: 13px;
    margin: 0;
    padding: 0;
}

.header {
    border-bottom: 1px solid #999999;
    padding: 16px 20px;
    background-color: #ffffff;
}

.header h1 {
    margin: 0 0 8px 0;
    color: #333333;
    font-size: 22px;
}

.header .meta {
    color: #333333;
    line-height: 1.5em;
}

.summary {
    display: flex;
    gap: 12px;
    padding: 12px 20px;
    border-bottom: 1px solid #cccccc;
    background: #f7f7f7;
    flex-wrap: wrap;
}

.summary-card {
    border: 1px solid #cccccc;
    background: #ffffff;
    padding: 8px 12px;
    min-width: 130px;
}

.summary-card .number {
    font-size: 20px;
    font-weight: bold;
}

.summary-card .label {
    color: #555555;
}

.legend {
    padding: 10px 20px;
    border-bottom: 1px solid #cccccc;
}

.legend span {
    display: inline-block;
    margin-right: 16px;
    padding: 4px 8px;
    border: 1px solid #bbbbbb;
}

.diff-added {
    background-color: lightgreen;
}

.diff-removed {
    background-color: salmon;
}

.diff-changed {
    background-color: lightskyblue;
}

.diff-same {
    background-color: #ffffff;
}

.scope {
    margin: 12px 20px;
    border: 1px solid #bbbbbb;
}

.scope > summary {
    background-color: #fef7d6;
    padding: 8px 10px;
    font-weight: bold;
    cursor: pointer;
}

.container {
    margin: 8px 12px;
    border: 1px solid #bbbbbb;
}

.container > summary {
    background-color: #a0bacb;
    padding: 8px 10px;
    font-weight: bold;
    cursor: pointer;
}

.category {
    margin: 8px 12px;
    border: 1px solid #bbbbbb;
}

.category > summary {
    background-color: #c0d2de;
    padding: 8px 10px;
    font-weight: bold;
    cursor: pointer;
}

.settingtype {
    margin: 8px 12px;
    border: 1px solid #bbbbbb;
}

.settingtype > summary {
    background-color: #d9e3ea;
    padding: 8px 10px;
    font-weight: bold;
    cursor: pointer;
}

table {
    border-collapse: collapse;
    table-layout: fixed;
    width: 100%;
    font-size: 12px;
}

th {
    border-bottom: 1px solid #999999;
    background: #e8e8e8;
    text-align: left;
    padding: 6px;
}

td {
    border-bottom: 1px solid #dddddd;
    padding: 6px;
    vertical-align: top;
    word-break: break-word;
}

.value-column {
    font-family: Consolas, monospace;
}

.path-column {
    font-family: Consolas, monospace;
}

.footer {
    padding: 12px 20px;
    color: #555555;
    font-size: 12px;
}
"@)
    [void]$html.AppendLine("</style>")
    [void]$html.AppendLine("</head>")
    [void]$html.AppendLine("<body>")

    [void]$html.AppendLine("<div class='header'>")
    [void]$html.AppendLine("<h1>GPO Difference Report</h1>")
    [void]$html.AppendLine("<div class='meta'>")
    [void]$html.AppendLine("<strong>Reference GPO:</strong> $(ConvertTo-HtmlEncodedText $Gpo1Name)<br>")
    [void]$html.AppendLine("<strong>Comparison GPO:</strong> $(ConvertTo-HtmlEncodedText $Gpo2Name)<br>")
    [void]$html.AppendLine("<strong>Generated:</strong> $(ConvertTo-HtmlEncodedText $generatedOn)")
    [void]$html.AppendLine("</div>")
    [void]$html.AppendLine("</div>")

    [void]$html.AppendLine("<div class='summary'>")
    [void]$html.AppendLine("<div class='summary-card'><div class='number'>$totalCount</div><div class='label'>Total rows</div></div>")
    [void]$html.AppendLine("<div class='summary-card diff-added'><div class='number'>$addedCount</div><div class='label'>Added</div></div>")
    [void]$html.AppendLine("<div class='summary-card diff-removed'><div class='number'>$removedCount</div><div class='label'>Removed</div></div>")
    [void]$html.AppendLine("<div class='summary-card diff-changed'><div class='number'>$changedCount</div><div class='label'>Changed</div></div>")
    [void]$html.AppendLine("<div class='summary-card diff-same'><div class='number'>$sameCount</div><div class='label'>Same</div></div>")
    [void]$html.AppendLine("</div>")

    [void]$html.AppendLine("<div class='legend'>")
    [void]$html.AppendLine("<span class='diff-added'>[+] Added</span>")
    [void]$html.AppendLine("<span class='diff-removed'>[-] Removed</span>")
    [void]$html.AppendLine("<span class='diff-changed'>[#] Changed</span>")
    [void]$html.AppendLine("<span class='diff-same'>Same</span>")
    [void]$html.AppendLine("</div>")

    $groupedByScope = $Differences |
        Sort-Object `
            PolicyScopeSortOrder,
            PolicyContainerSortOrder,
            SettingCategorySortOrder,
            SettingTypeSortOrder,
            SettingType,
            Path,
            Property |
        Group-Object PolicyScope

    foreach ($scopeGroup in $groupedByScope) {

        [void]$html.AppendLine("<details class='scope' open>")
        [void]$html.AppendLine("<summary>$(ConvertTo-HtmlEncodedText $scopeGroup.Name) Configuration</summary>")

        $containerGroups = $scopeGroup.Group |
            Sort-Object PolicyContainerSortOrder |
            Group-Object PolicyContainer

        foreach ($containerGroup in $containerGroups) {

            [void]$html.AppendLine("<details class='container' open>")
            [void]$html.AppendLine("<summary>$(ConvertTo-HtmlEncodedText $containerGroup.Name)</summary>")

            $categoryGroups = $containerGroup.Group |
                Sort-Object SettingCategorySortOrder |
                Group-Object SettingCategory

            foreach ($categoryGroup in $categoryGroups) {

                [void]$html.AppendLine("<details class='category' open>")
                [void]$html.AppendLine("<summary>$(ConvertTo-HtmlEncodedText $categoryGroup.Name)</summary>")

                $typeGroups = $categoryGroup.Group |
                    Sort-Object SettingTypeSortOrder, SettingType, Path, Property |
                    Group-Object SettingType

                foreach ($typeGroup in $typeGroups) {

                    [void]$html.AppendLine("<details class='settingtype' open>")
                    [void]$html.AppendLine("<summary>$(ConvertTo-HtmlEncodedText $typeGroup.Name)</summary>")

                    [void]$html.AppendLine("<table>")
                    [void]$html.AppendLine("<thead>")
                    [void]$html.AppendLine("<tr>")

                    foreach ($column in $Columns) {
                        [void]$html.AppendLine("<th>$(ConvertTo-HtmlEncodedText $column)</th>")
                    }

                    [void]$html.AppendLine("</tr>")
                    [void]$html.AppendLine("</thead>")
                    [void]$html.AppendLine("<tbody>")

                    foreach ($row in $typeGroup.Group) {

                        $cssClass = Get-DifferenceCssClass -DifferenceType $row.DifferenceType
                        $prefix = Get-DifferencePrefix -DifferenceType $row.DifferenceType

                        [void]$html.AppendLine("<tr class='$cssClass'>")

                        foreach ($column in $Columns) {

                            $value = $row.PSObject.Properties[$column].Value

                            if ($column -eq "DifferenceType") {
                                $value = "$prefix$value"
                            }

                            $tdClass = if ($column -eq "Path") {
                                "path-column"
                            }
                            elseif ($column -like "* Value") {
                                "value-column"
                            }
                            else {
                                ""
                            }

                            [void]$html.AppendLine("<td class='$tdClass'>$(ConvertTo-HtmlEncodedText $value)</td>")
                        }

                        [void]$html.AppendLine("</tr>")
                    }

                    [void]$html.AppendLine("</tbody>")
                    [void]$html.AppendLine("</table>")
                    [void]$html.AppendLine("</details>")
                }

                [void]$html.AppendLine("</details>")
            }

            [void]$html.AppendLine("</details>")
        }

        [void]$html.AppendLine("</details>")
    }

    [void]$html.AppendLine("<div class='footer'>Generated by Compare-GPOs.ps1</div>")
    [void]$html.AppendLine("</body>")
    [void]$html.AppendLine("</html>")

    $html.ToString() | Set-Content -LiteralPath $Path -Encoding UTF8
}

function Invoke-GpoComparison {
    param(
        [Parameter(Mandatory)]
        [string]$Gpo1BackupFolder,

        [Parameter(Mandatory)]
        [string]$Gpo2BackupFolder,

        [Parameter(Mandatory)]
        [string]$OutputFolder,

        [switch]$IncludeSame
    )

    if (-not (Test-Path -LiteralPath $OutputFolder -PathType Container)) {
        New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
    }

    $Gpo1Path = Resolve-GpoReportXmlPath -FolderPath $Gpo1BackupFolder
    $Gpo2Path = Resolve-GpoReportXmlPath -FolderPath $Gpo2BackupFolder

    $Gpo1 = Get-GpoExtractedObjects -Path $Gpo1Path
    $Gpo2 = Get-GpoExtractedObjects -Path $Gpo2Path

    $Gpo1ValueColumn = "$($Gpo1.Name) Value"
    $Gpo2ValueColumn = "$($Gpo2.Name) Value"

    $commonCompareParameters = @{
        Gpo1ValueColumn = $Gpo1ValueColumn
        Gpo2ValueColumn = $Gpo2ValueColumn
    }

    if ($IncludeSame) {
        $commonCompareParameters.IncludeSame = $true
    }

    $AllDifferences = @(

        Compare-ObjectSet `
            -ReferenceObject $Gpo1.Scripts `
            -DifferenceObject $Gpo2.Scripts `
            -KeyScript { param($x) "$($x.Type)\$($x.Order)\$($x.Command)" } `
            -CompareProperties @("Command", "Type", "Order", "RunOrder") `
            -ObjectType "Scripts" `
            @commonCompareParameters

        Compare-ObjectSet `
            -ReferenceObject $Gpo1.AccountPolicies `
            -DifferenceObject $Gpo2.AccountPolicies `
            -KeyScript { param($x) "$($x.Type)\$($x.Name)" } `
            -CompareProperties @("SettingValue") `
            -ObjectType "AccountPolicies" `
            @commonCompareParameters

        Compare-ObjectSet `
            -ReferenceObject $Gpo1.AuditPolicies `
            -DifferenceObject $Gpo2.AuditPolicies `
            -KeyScript { param($x) "$($x.Name)" } `
            -CompareProperties @("SuccessAttempts", "FailureAttempts") `
            -ObjectType "AuditPolicies" `
            @commonCompareParameters

        Compare-ObjectSet `
            -ReferenceObject $Gpo1.AdvancedAuditSettings `
            -DifferenceObject $Gpo2.AdvancedAuditSettings `
            -KeyScript { param($x) "$($x.PolicyTarget)\$($x.SubcategoryName)" } `
            -CompareProperties @("SettingValue") `
            -ObjectType "AdvancedAuditSettings" `
            @commonCompareParameters

        Compare-ObjectSet `
            -ReferenceObject $Gpo1.UserRights `
            -DifferenceObject $Gpo2.UserRights `
            -KeyScript { param($x) "$($x.Name)" } `
            -CompareProperties @("Member") `
            -ObjectType "UserRights" `
            @commonCompareParameters

        Compare-ObjectSet `
            -ReferenceObject $Gpo1.SecurityOptions `
            -DifferenceObject $Gpo2.SecurityOptions `
            -KeyScript {
                param($x)

                $identity = if ($x.KeyName) {
                    $x.KeyName
                }
                elseif ($x.SystemAccessPolicyName) {
                    $x.SystemAccessPolicyName
                }
                else {
                    $x.Name
                }

                "$identity"
            } `
            -CompareProperties @("SettingValue") `
            -ObjectType "SecurityOptions" `
            @commonCompareParameters

        Compare-ObjectSet `
            -ReferenceObject $Gpo1.EventLogSettings `
            -DifferenceObject $Gpo2.EventLogSettings `
            -KeyScript { param($x) "$($x.Log)\$($x.Name)" } `
            -CompareProperties @("SettingValue") `
            -ObjectType "EventLogSettings" `
            @commonCompareParameters

        Compare-ObjectSet `
            -ReferenceObject $Gpo1.RestrictedGroups `
            -DifferenceObject $Gpo2.RestrictedGroups `
            -KeyScript { param($x) "$($x.Name)" } `
            -CompareProperties @("Member", "MemberOf") `
            -ObjectType "RestrictedGroups" `
            @commonCompareParameters

        Compare-ObjectSet `
            -ReferenceObject $Gpo1.PolicySettings `
            -DifferenceObject $Gpo2.PolicySettings `
            -KeyScript { param($x) "$($x.Category)\$($x.Name)" } `
            -CompareProperties @("State") `
            -ObjectType "PolicySettings" `
            @commonCompareParameters

        Compare-ObjectSet `
            -ReferenceObject $Gpo1.RegistrySettings `
            -DifferenceObject $Gpo2.RegistrySettings `
            -KeyScript { param($x) "$($x.Hive)\$($x.Key)\$($x.ValueName)" } `
            -CompareProperties @("Action", "Type", "Value", "RemovePolicy") `
            -ObjectType "RegistrySettings" `
            @commonCompareParameters
    )

    $OutputColumns = @(
        "PolicyScope",
        "PolicyContainer",
        "SettingCategory",
        "SettingType",
        "SettingName",
        "Path",
        "DisplayValue",
        "Property",
        "DifferenceType",
        $Gpo1ValueColumn,
        $Gpo2ValueColumn
    )

    $SortedDifferences = $AllDifferences |
        Sort-Object `
            PolicyScopeSortOrder,
            PolicyContainerSortOrder,
            SettingCategorySortOrder,
            SettingTypeSortOrder,
            SettingType,
            Path,
            Property

    $safeGpo1Name = $Gpo1.Name -replace '[\\/:*?"<>|]', '_'
    $safeGpo2Name = $Gpo2.Name -replace '[\\/:*?"<>|]', '_'

    $CsvOutputPath  = Join-Path $OutputFolder "$safeGpo1Name-vs-$safeGpo2Name-GPO_Differences.csv"
    $HtmlOutputPath = Join-Path $OutputFolder "$safeGpo1Name-vs-$safeGpo2Name-GPO_Differences.html"

    $SortedDifferences |
        Select-Object $OutputColumns |
        Export-Csv $CsvOutputPath -NoTypeInformation -Encoding UTF8

    Export-GpoDifferenceHtml `
        -Differences $SortedDifferences `
        -Columns $OutputColumns `
        -Gpo1Name $Gpo1.Name `
        -Gpo2Name $Gpo2.Name `
        -Path $HtmlOutputPath

    [pscustomobject]@{
        Gpo1Name       = $Gpo1.Name
        Gpo2Name       = $Gpo2.Name
        Gpo1XmlPath    = $Gpo1Path
        Gpo2XmlPath    = $Gpo2Path
        CsvOutputPath  = $CsvOutputPath
        HtmlOutputPath = $HtmlOutputPath
        Differences    = $SortedDifferences
    }
}
