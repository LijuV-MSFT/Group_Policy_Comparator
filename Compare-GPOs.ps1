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
                            SettingCategory = "Windows Settings"
                            SettingType     = "Scripts"
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

                    ###############################
                    # Account Policies           #
                    ###############################

                    $AccountPolicies = foreach ($account in @($extension.Extension.Account)) {

                        $settingValue = if ($null -ne $account.SettingBoolean -and $account.SettingBoolean -ne "") {
                            $account.SettingBoolean
                        }
                        elseif ($null -ne $account.SettingNumber -and $account.SettingNumber -ne "") {
                            $account.SettingNumber
                        }
                        elseif ($null -ne $account.SettingString -and $account.SettingString -ne "") {
                            $account.SettingString
                        }
                        else {
                            $null
                        }

                        [pscustomobject]@{
                            PolicyScope     = $PolicyScopeName
                            SettingCategory = "Security settings"
                            SettingType     = "Account policies"
                            Name            = $account.Name
                            Type            = $account.Type
                            SettingValue    = $settingValue
                        }
                    }

                    ###############################
                    # Audit Policy               #
                    ###############################

                    $AuditPolicies = foreach ($audit in @($extension.Extension.Audit)) {
                        [pscustomobject]@{
                            PolicyScope     = $PolicyScopeName
                            SettingCategory = "Security settings"
                            SettingType     = "Audit policy"
                            Name            = $audit.Name
                            SuccessAttempts = $audit.SuccessAttempts
                            FailureAttempts = $audit.FailureAttempts
                        }
                    }

                    ###############################
                    # User Rights Assignment     #
                    ###############################

                    $UserRights = foreach ($right in @($extension.Extension.UserRightsAssignment)) {
                        [pscustomobject]@{
                            PolicyScope     = $PolicyScopeName
                            SettingCategory = "Security settings"
                            SettingType     = "User Rights"
                            Name            = $right.Name
                            Member          = @(
                                $right.Member | ForEach-Object {
                                    $_.SID.'#text'
                                } | Where-Object {
                                    $_
                                } | Sort-Object
                            ) -join ';'
                        }
                    }

                    ###############################
                    # Security Options           #
                    ###############################

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

                        $settingValue = if ($null -ne $setting.SettingString -and $setting.SettingString -ne "") {
                            $setting.SettingString
                        }
                        elseif ($null -ne $setting.SettingNumber -and $setting.SettingNumber -ne "") {
                            $setting.SettingNumber
                        }
                        elseif ($null -ne $setting.SettingBoolean -and $setting.SettingBoolean -ne "") {
                            $setting.SettingBoolean
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
                            SettingCategory        = "Security settings"
                            SettingType            = "Security options"
                            Name                   = $friendlyName
                            Units                  = $display.Units
                            DisplayValue           = $displayValue
                            KeyName                = $setting.KeyName
                            SettingValue           = $settingValue
                            SystemAccessPolicyName = $setting.SystemAccessPolicyName
                        }
                    }

                    ###############################
                    # Event Log                  #
                    ###############################

                    $EventLogSettings = foreach ($eventLog in @($extension.Extension.EventLog)) {

                        $settingValue = if ($null -ne $eventLog.SettingNumber -and $eventLog.SettingNumber -ne "") {
                            $eventLog.SettingNumber
                        }
                        elseif ($null -ne $eventLog.SettingBoolean -and $eventLog.SettingBoolean -ne "") {
                            $eventLog.SettingBoolean
                        }
                        elseif ($null -ne $eventLog.SettingString -and $eventLog.SettingString -ne "") {
                            $eventLog.SettingString
                        }
                        else {
                            $null
                        }

                        [pscustomobject]@{
                            PolicyScope     = $PolicyScopeName
                            SettingCategory = "Security settings"
                            SettingType     = "Event Log"
                            Name            = $eventLog.Name
                            Log             = $eventLog.Log
                            SettingValue    = $settingValue
                        }
                    }

                    ###############################
                    # Restricted Groups          #
                    ###############################

                    $RestrictedGroups = foreach ($group in @($extension.Extension.RestrictedGroups)) {

                        $groupName = $group.GroupName.Name.'#text'

                        $members = @(
                            $group.Member | ForEach-Object {
                                $_.Name.'#text'
                            } | Where-Object {
                                $_
                            } | Sort-Object
                        ) -join ';'

                        $memberOf = @(
                            $group.Memberof | ForEach-Object {
                                $_.Name.'#text'
                            } | Where-Object {
                                $_
                            } | Sort-Object
                        ) -join ';'

                        [pscustomobject]@{
                            PolicyScope     = $PolicyScopeName
                            SettingCategory = "Security settings"
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

                "Registry" {

                    $PolicySettings = foreach ($policy in @($extension.Extension.Policy)) {
                        [pscustomobject]@{
                            PolicyScope     = $PolicyScopeName
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
                            SettingCategory = "Preferences"
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
        Name              = $GPOName
        Scripts           = $AllScripts
        AccountPolicies   = $AllAccountPolicies
        AuditPolicies     = $AllAuditPolicies
        UserRights        = $AllUserRights
        SecurityOptions   = $AllSecurityOptions
        EventLogSettings  = $AllEventLogSettings
        RestrictedGroups  = $AllRestrictedGroups
        PolicySettings    = $AllPolicySettings
        RegistrySettings  = $AllRegistrySettings
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
                    PolicyScope     = $dif.PolicyScope
                    SettingCategory = $dif.SettingCategory
                    SettingType     = $dif.SettingType
                    SettingName     = $settingName
                    Setting         = $key
                    DisplayValue    = $displayValue
                    Property        = $property
                    DifferenceType  = "Added"
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
                    PolicyScope     = $ref.PolicyScope
                    SettingCategory = $ref.SettingCategory
                    SettingType     = $ref.SettingType
                    SettingName     = $settingName
                    Setting         = $key
                    DisplayValue    = $displayValue
                    Property        = $property
                    DifferenceType  = "Removed"
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

            if ("$refValue" -ne "$difValue") {

                $row = [ordered]@{
                    PolicyScope     = if ($ref.PolicyScope) { $ref.PolicyScope } else { $dif.PolicyScope }
                    SettingCategory = if ($ref.SettingCategory) { $ref.SettingCategory } else { $dif.SettingCategory }
                    SettingType     = if ($ref.SettingType) { $ref.SettingType } else { $dif.SettingType }
                    SettingName     = $settingName
                    Setting         = $key
                    DisplayValue    = $displayValue
                    Property        = $property
                    DifferenceType  = "Changed"
                }

                $row[$Gpo1ValueColumn] = $refValue
                $row[$Gpo2ValueColumn] = $difValue
                $row["ObjectType"]     = $ObjectType

                [pscustomobject]$row
            }
            elseif ($IncludeSame) {

                $row = [ordered]@{
                    PolicyScope     = if ($ref.PolicyScope) { $ref.PolicyScope } else { $dif.PolicyScope }
                    SettingCategory = if ($ref.SettingCategory) { $ref.SettingCategory } else { $dif.SettingCategory }
                    SettingType     = if ($ref.SettingType) { $ref.SettingType } else { $dif.SettingType }
                    SettingName     = $settingName
                    Setting         = $key
                    DisplayValue    = $displayValue
                    Property        = $property
                    DifferenceType  = "Same"
                }

                $row[$Gpo1ValueColumn] = $refValue
                $row[$Gpo2ValueColumn] = $difValue
                $row["ObjectType"]     = $ObjectType

                [pscustomobject]$row
            }
        }
    }
}

Clear-Host

$Gpo1 = Get-GpoExtractedObjects -Path "C:\Temp\GPO_Comparison\GPO_Comparison_GPO1.xml"
$Gpo2 = Get-GpoExtractedObjects -Path "C:\Temp\GPO_Comparison\GPO_Comparison_GPO2.xml"

$Gpo1ValueColumn = "$($Gpo1.Name) Value"
$Gpo2ValueColumn = "$($Gpo2.Name) Value"

$AllDifferences = @(

    Compare-ObjectSet `
        -ReferenceObject $Gpo1.Scripts `
        -DifferenceObject $Gpo2.Scripts `
        -KeyScript { param($x) "$($x.Type)\$($x.Order)\$($x.Command)" } `
        -CompareProperties @(
            "Command",
            "Type",
            "Order",
            "RunOrder"
        ) `
        -ObjectType "Scripts" `
        -Gpo1ValueColumn $Gpo1ValueColumn `
        -Gpo2ValueColumn $Gpo2ValueColumn `
        -IncludeSame

    Compare-ObjectSet `
        -ReferenceObject $Gpo1.AccountPolicies `
        -DifferenceObject $Gpo2.AccountPolicies `
        -KeyScript { param($x) "$($x.Type)\$($x.Name)" } `
        -CompareProperties @(
            "SettingValue"
        ) `
        -ObjectType "AccountPolicies" `
        -Gpo1ValueColumn $Gpo1ValueColumn `
        -Gpo2ValueColumn $Gpo2ValueColumn `
        -IncludeSame

    Compare-ObjectSet `
        -ReferenceObject $Gpo1.AuditPolicies `
        -DifferenceObject $Gpo2.AuditPolicies `
        -KeyScript { param($x) "$($x.Name)" } `
        -CompareProperties @(
            "SuccessAttempts",
            "FailureAttempts"
        ) `
        -ObjectType "AuditPolicies" `
        -Gpo1ValueColumn $Gpo1ValueColumn `
        -Gpo2ValueColumn $Gpo2ValueColumn `
        -IncludeSame

    Compare-ObjectSet `
        -ReferenceObject $Gpo1.UserRights `
        -DifferenceObject $Gpo2.UserRights `
        -KeyScript { param($x) "$($x.Name)" } `
        -CompareProperties @(
            "Member"
        ) `
        -ObjectType "UserRights" `
        -Gpo1ValueColumn $Gpo1ValueColumn `
        -Gpo2ValueColumn $Gpo2ValueColumn `
        -IncludeSame

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
        -CompareProperties @(
            "SettingValue"
        ) `
        -ObjectType "SecurityOptions" `
        -Gpo1ValueColumn $Gpo1ValueColumn `
        -Gpo2ValueColumn $Gpo2ValueColumn `
        -IncludeSame

    Compare-ObjectSet `
        -ReferenceObject $Gpo1.EventLogSettings `
        -DifferenceObject $Gpo2.EventLogSettings `
        -KeyScript { param($x) "$($x.Log)\$($x.Name)" } `
        -CompareProperties @(
            "SettingValue"
        ) `
        -ObjectType "EventLogSettings" `
        -Gpo1ValueColumn $Gpo1ValueColumn `
        -Gpo2ValueColumn $Gpo2ValueColumn `
        -IncludeSame

    Compare-ObjectSet `
        -ReferenceObject $Gpo1.RestrictedGroups `
        -DifferenceObject $Gpo2.RestrictedGroups `
        -KeyScript { param($x) "$($x.Name)" } `
        -CompareProperties @(
            "Member",
            "MemberOf"
        ) `
        -ObjectType "RestrictedGroups" `
        -Gpo1ValueColumn $Gpo1ValueColumn `
        -Gpo2ValueColumn $Gpo2ValueColumn `
        -IncludeSame

    Compare-ObjectSet `
        -ReferenceObject $Gpo1.PolicySettings `
        -DifferenceObject $Gpo2.PolicySettings `
        -KeyScript { param($x) "$($x.Category)\$($x.Name)" } `
        -CompareProperties @(
            "State"
        ) `
        -ObjectType "PolicySettings" `
        -Gpo1ValueColumn $Gpo1ValueColumn `
        -Gpo2ValueColumn $Gpo2ValueColumn `
        -IncludeSame

    Compare-ObjectSet `
        -ReferenceObject $Gpo1.RegistrySettings `
        -DifferenceObject $Gpo2.RegistrySettings `
        -KeyScript { param($x) "$($x.Hive)\$($x.Key)\$($x.ValueName)" } `
        -CompareProperties @(
            "Action",
            "Type",
            "Value",
            "RemovePolicy"
        ) `
        -ObjectType "RegistrySettings" `
        -Gpo1ValueColumn $Gpo1ValueColumn `
        -Gpo2ValueColumn $Gpo2ValueColumn `
        -IncludeSame
)

$OutputColumns = @(
    "PolicyScope",
    "SettingCategory",
    "SettingType",
    "SettingName",
    "Setting",
    "DisplayValue",
    "Property",
    "DifferenceType",
    $Gpo1ValueColumn,
    $Gpo2ValueColumn
)

$AllDifferences |
    Sort-Object PolicyScope, SettingCategory, SettingType, Setting, Property |
    Format-Table $OutputColumns -AutoSize

$AllDifferences |
    Sort-Object PolicyScope, SettingCategory, SettingType, Setting, Property |
    Select-Object $OutputColumns |
    Export-Csv "C:\Temp\GPO_Comparison\GPO_Differences.csv" -NoTypeInformation -Encoding UTF8
