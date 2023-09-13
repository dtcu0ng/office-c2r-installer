function gatherInfomation() {
    # todo: use number selector to generate config
    $SourcePath = Read-Host -Prompt "Enter the source path, if you want to "
    $OfficeClientEdition = Read-Host -Prompt "Enter the office client edition (32 or 64)"
    $Channel = Read-Host -Prompt "Enter the update channel"
    $ProductIDs = Read-Host -Prompt "Enter the product IDs (comma-separated)"
    $LanguageIDs = Read-Host -Prompt "Enter the language IDs (comma-separated)"
    $ExcludeAppIDs = Read-Host -Prompt "Enter the app IDs to exclude (comma-separated)"
    $UpdatePath = Read-Host -Prompt "Enter the update path"
    $UpdatesEnabled = [bool]::Parse((Read-Host -Prompt "Enable updates? (true or false)"))
    $DisplayLevel = Read-Host -Prompt "Enter the display level"
    $AcceptEULA = [bool]::Parse((Read-Host -Prompt "Accept EULA? (true or false)"))
    $ProductIDs = $ProductIDs.Split(",")
    $LanguageIDs = $LanguageIDs.Split(",")
    $ExcludeAppIDs = $ExcludeAppIDs.Split(",")
    generateConfig
}

function generateConfig() {
    # Create XML configuration
    $config = @"
    <Configuration>
    <Add OfficeClientEdition="$OfficeClientEdition" Channel="$Channel">
        $(foreach ($product in $ProductIDs) {
            "<Product ID=`"$product`">"
            foreach ($language in $LanguageIDs) {
                "<Language ID=`"$language`" />"
            }
            foreach ($excludeApp in $ExcludeAppIDs) {
                "<ExcludeApp ID=`"$excludeApp`" />"
            }
            "</Product>"
        })
    </Add>
    <Updates Enabled="$UpdatesEnabled" UpdatePath="$UpdatePath" />
    <Display Level="$DisplayLevel" AcceptEULA="$AcceptEULA" />
    </Configuration>
"@

    # Save configuration to file
    $config | Out-File -FilePath .\configuration.xml -Encoding utf8
}
gatherInfomation

