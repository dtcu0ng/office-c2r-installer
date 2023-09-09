function header() {
    Clear-Host
    Write-Host "office-c2r-installer (WIP)"
    Write-Host "Version 0.0.1 beta 0 (non-working preview)"
    Write-Host "https://github.com/dtcu0ng/office-c2r-installer"
}

function customConfigOnline() {
    Clear-Host
    $configPath = Read-Host "Please enter path of the configuration file you want to install. Type q to go back"
    if ($configPath -eq "q") {
        main
    } elseif (!(Test-Path -Path $configPath))  {
        Write-Host "Configuration file not found. Please double-check your configuration file and try again."
    } else {
        Write-Host "Starting Office setup with a specified configuration: $configPath"
        Start-Process -FilePath "./files/setup.exe" -ArgumentList "/configure $configPath"
    }
}

function preConfigOnline() {
    preConfigMenu
}

function customConfigOffline() {
    Clear-Host
    $configPath = Read-Host "Please enter path of the configuration file you want to install. Type q to go back"
    if ($configPath -eq "q") {
        main
    } elseif (!(Test-Path -Path $configPath))  {
        Write-Host "Configuration file not found. Please double-check your configuration file and try again."
    } else {
        $installerPath = Read-Host "Please specify your Office installer path. Type q to go back to main menu."
        if ($installPath -eq "q") {
            main
        } elseif (!(Test-Path -Path $installerPath)) {
            Write-Host "Installer path not found. Please double check your installer path and try again."
        } else {
            Write-Host "WIP"
        }
    }
}

function preConfigOffline() {
    Clear-Host
    $installerPath = Read-Host "Please specify your Office installer path. Type q to go back to main menu."
    if ($installerPath -eq "q") {
        main
    } elseif (!(Test-Path -Path $configPath)) {
        Write-Host "Configuration file not found. Please double-check your configuration file and try again."
    } else {
        preConfigMenu
    }

}

# WIP, will fix it later

function configGenerator() {
#     gatherInfomation() {
# $SourcePath = Read-Host -Prompt "Enter the source path"
# $OfficeClientEdition = Read-Host -Prompt "Enter the office client edition (32 or 64)"
# $Channel = Read-Host -Prompt "Enter the update channel"
# $ProductIDs = Read-Host -Prompt "Enter the product IDs (comma-separated)"
# $LanguageIDs = Read-Host -Prompt "Enter the language IDs (comma-separated)"
# $ExcludeAppIDs = Read-Host -Prompt "Enter the app IDs to exclude (comma-separated)"
# $UpdatePath = Read-Host -Prompt "Enter the update path"
# $UpdatesEnabled = [bool]::Parse((Read-Host -Prompt "Enable updates? (true or false)"))
# $DisplayLevel = Read-Host -Prompt "Enter the display level"
# $AcceptEULA = [bool]::Parse((Read-Host -Prompt "Accept EULA? (true or false)"))
# $ProductIDs = $ProductIDs.Split(",")
# $LanguageIDs = $LanguageIDs.Split(",")
# $ExcludeAppIDs = $ExcludeAppIDs.Split(",")
# }

# generateConfig() {
# # Create XML configuration
# $config = @"
# <Configuration>
#     <Add OfficeClientEdition="$OfficeClientEdition" Channel="$Channel">
#         $(foreach ($product in $ProductIDs) {
#             "<Product ID=`"$product`">"
#             foreach ($language in $LanguageIDs) {
#                 "<Language ID=`"$language`" />"
#             }
#             foreach ($excludeApp in $ExcludeAppIDs) {
#                 "<ExcludeApp ID=`"$excludeApp`" />"
#             }
#             "</Product>"
#         })
#     </Add>
#     <Updates Enabled="$UpdatesEnabled" UpdatePath="$UpdatePath" />
#     <Display Level="$DisplayLevel" AcceptEULA="$AcceptEULA" />
# </Configuration>
# "@

# # Save configuration to file
# $config | Out-File -FilePath .\configuration.xml -Encoding utf8
# }
Write-Host "WIP"
}

function c2rExtractor {
    # todo...
    Write-Host "WIP"
}

function preConfigMenu {
# todo: list of pre-configured configuration, can select, use en-us as main language.
# two type of configuration are available - full or mininal (word, excel, powerpoint,...)
    Write-Host "WIP"
}

# codes in below this line is pretty simple, just the selector and the bootstrap code in main()
# should commented out all the code need to fix or not working before push to Github.
# this prevent CI failed to run because known not working feature.

# todo: integrate CI to fix syntax.

function mainSelector {
    param (
        [string]$Title = 'office-c2r-installer Selector'
    )
    Write-Host "================================================================="
    Write-Host "[1]: Install Office with specified configuration (online)"
    Write-Host "[2]: Install Office with pre-configured configuration (online)"
    Write-Host "================================================================="
    Write-Host "[3]: Install Office with specified configuration (offline)"
    Write-Host "[4]: Install Office with pre-configured configuration (offline)"
    Write-Host "================================================================="
    Write-Host "[G]: Generate configuration"
    Write-Host "[Q]: Quit"
    Write-Host "================================================================="
}

function main() {
    header
    mainSelector –Title 'office-c2r-installer Selector'
    $selection = Read-Host "Please use keyboard to make a selection"
    
    switch ($selection) {
        '1' { customConfigOnline }
        '2' { preConfigOnline }
        '3' { customConfigOffline }
        '4' { preConfigOffline }
        'g' { configGenerator }
        'q' { return }
    }
}

# run the script
main
