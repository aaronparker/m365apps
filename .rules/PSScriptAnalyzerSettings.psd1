@{
    CustomRulePath      = @(
        ".rules/LowercaseKeyword.psm1"
    )
    IncludeDefaultRules = $true
    Severity            = @("Error", "Warning")
    IncludeRules        = @(
        "Measure-LowercaseKeyword"
    )
}
