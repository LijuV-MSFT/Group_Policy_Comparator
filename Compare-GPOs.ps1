function Get-GpoExtractedObjects {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    [xml]$GPOXMLData = Get-Content -LiteralPath $Path

    $GPOName = $GPOXMLData.GPO.Name

    $AllUserRights = @()
    $AllSecurityOptions = @()
    $AllPolicySettings = @()
    $AllRegistrySettings = @()

    foreach ($sectionName in @("Computer", "User")) {

        $sectionNode = $GPOXMLData.GPO.$sectionName

        foreach ($extension in @($sectionNode.ExtensionData)) {

            switch ($extension.Name) {

                "Security" {

                    $UserRights = foreach ($right in @($extension.Extension.UserRightsAssignment)) {
                        [pscustomobject]@{
                            Section         = $sectionName
                            SettingCategory = "Security settings"
                            SettingType     = "User Rights"
                            Name            = $right.Name
                            Member          = @($right.Member | ForEach-Object {
                                $_.SID.'#text'
                            }) -join ';'
                        }
                    }

                    $SecurityOptions = foreach ($setting in @($extension.Extension.SecurityOptions)) {

                        $display = $setting.Display

                        $displayValue = if ($display.DisplayString) {
                            $display.DisplayString
                        }
                        elseif ($null -ne $display.DisplayNumber) {
                            $display.DisplayNumber
                        }
                        elseif ($null -ne $display.DisplayBoolean) {
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

                        $settingValue = if ($setting.SettingString) {
                            $setting.SettingString
                        }
                        elseif ($null -ne $setting.SettingNumber) {
                            $setting.SettingNumber
                        }

                        [pscustomobject]@{
                            Section                = $sectionName
                            SettingCategory        = "Security settings"
                            SettingType            = "Security options"
                            Name                   = $display.Name
                            Units                  = $display.Units
                            DisplayValue           = $displayValue
                            KeyName                = $setting.KeyName
                            SettingValue           = $settingValue
                            SystemAccessPolicyName = $setting.SystemAccessPolicyName
                        }
                    }

                    $AllUserRights += $UserRights
                    $AllSecurityOptions += $SecurityOptions
                }

                "Registry" {

                    $PolicySettings = foreach ($policy in @($extension.Extension.Policy)) {
                        [pscustomobject]@{
                            Section         = $sectionName
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
                            Section         = $sectionName
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
        Name             = $GPOName
        UserRights       = $AllUserRights
        SecurityOptions  = $AllSecurityOptions
        PolicySettings   = $AllPolicySettings
        RegistrySettings = $AllRegistrySettings
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

        if (-not $ref) {

            foreach ($property in $CompareProperties) {
                [pscustomobject]@{
                    ObjectType      = $ObjectType
                    Section         = $dif.Section
                    SettingCategory = $dif.SettingCategory
                    SettingType     = $dif.SettingType
                    DifferenceType  = "Added"
                    SettingKey      = $key
                    Property        = $property
                    ReferenceValue  = $null
                    DifferenceValue = $dif.$property
                }
            }

            continue
        }

        if (-not $dif) {

            foreach ($property in $CompareProperties) {
                [pscustomobject]@{
                    ObjectType      = $ObjectType
                    Section         = $ref.Section
                    SettingCategory = $ref.SettingCategory
                    SettingType     = $ref.SettingType
                    DifferenceType  = "Removed"
                    SettingKey      = $key
                    Property        = $property
                    ReferenceValue  = $ref.$property
                    DifferenceValue = $null
                }
            }

            continue
        }

        foreach ($property in $CompareProperties) {

            $refValue = $ref.$property
            $difValue = $dif.$property

            if ("$refValue" -ne "$difValue") {
                [pscustomobject]@{
                    ObjectType      = $ObjectType
                    Section         = if ($ref.Section) { $ref.Section } else { $dif.Section }
                    SettingCategory = if ($ref.SettingCategory) { $ref.SettingCategory } else { $dif.SettingCategory }
                    SettingType     = if ($ref.SettingType) { $ref.SettingType } else { $dif.SettingType }
                    DifferenceType  = "Changed"
                    SettingKey      = $key
                    Property        = $property
                    ReferenceValue  = $refValue
                    DifferenceValue = $difValue
                }
            }
            elseif ($IncludeSame) {
                [pscustomobject]@{
                    ObjectType      = $ObjectType
                    Section         = if ($ref.Section) { $ref.Section } else { $dif.Section }
                    SettingCategory = if ($ref.SettingCategory) { $ref.SettingCategory } else { $dif.SettingCategory }
                    SettingType     = if ($ref.SettingType) { $ref.SettingType } else { $dif.SettingType }
                    DifferenceType  = "Same"
                    SettingKey      = $key
                    Property        = $property
                    ReferenceValue  = $refValue
                    DifferenceValue = $difValue
                }
            }
        }
    }
}

Clear-Host

$Gpo1 = Get-GpoExtractedObjects -Path "C:\Temp\GPO_Comparison\GPO_Comparison_GPO1.xml"
$Gpo2 = Get-GpoExtractedObjects -Path "C:\Temp\GPO_Comparison\GPO_Comparison_GPO2.xml"

$AllDifferences = @(
    Compare-ObjectSet `
        -ReferenceObject $Gpo1.UserRights `
        -DifferenceObject $Gpo2.UserRights `
        -KeyScript { param($x) "$($x.Section)\$($x.Name)" } `
        -CompareProperties @("Member") `
        -ObjectType "UserRights"

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

            "$($x.Section)\$identity"
        } `
        -CompareProperties @(
            "Name",
            "Units",
            "DisplayValue",
            "SettingValue",
            "SystemAccessPolicyName"
        ) `
        -ObjectType "SecurityOptions"

    Compare-ObjectSet `
        -ReferenceObject $Gpo1.PolicySettings `
        -DifferenceObject $Gpo2.PolicySettings `
        -KeyScript { param($x) "$($x.Section)\$($x.Category)\$($x.Name)" } `
        -CompareProperties @("State") `
        -ObjectType "PolicySettings"

    Compare-ObjectSet `
        -ReferenceObject $Gpo1.RegistrySettings `
        -DifferenceObject $Gpo2.RegistrySettings `
        -KeyScript { param($x) "$($x.Section)\$($x.Hive)\$($x.Key)\$($x.ValueName)" } `
        -CompareProperties @(
            "Action",
            "Type",
            "Value",
            "RemovePolicy"
        ) `
        -ObjectType "RegistrySettings"
)

$AllDifferences |
    Sort-Object SettingCategory, SettingType, Section, SettingKey, Property |
    Format-Table `
        SettingCategory,
        SettingType,
        Section,
        DifferenceType,
        SettingKey,
        Property,
        ReferenceValue,
        DifferenceValue `
        -AutoSize
