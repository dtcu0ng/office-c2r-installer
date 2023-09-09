function header() {
    Write-Host "office-c2r-installer (WIP)"
    Write-Host "Version 0.0.1 alpha 2 (partial working preview - UNSTABLE)"
    Write-Host "https://github.com/dtcu0ng/office-c2r-installer"
}

function languageSelector() {
    #todo (low priority): multilingual script, support Vietnamese and English
}
function customConfigOnline() {
    Clear-Host
    $configPath = Read-Host "Please enter path of the configuration file you want to install. Type q to go back"
    if ($configPath -eq "q") {
        main
    } elseif (!(Test-Path -Path $configPath))  {
        Clear-Host
        Write-Host "Configuration file not found. Please double-check your configuration file and try again."
    } else {
        Clear-Host
        Write-Host "Starting Office setup with a specified configuration: $configPath"
        Write-Host "Please do not close any window or process during installation."
        Start-Process -FilePath "./files/setup.exe" -ArgumentList "/configure $configPath" -Wait
        Clear-Host
        Write-Host "Setup has been exited, please check that your Office installation is successfully installed."
        Write-Host "If not, please try again. If the problem persists, please open an issue on Github"
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
# Create XML configuration (WIP)
    #$SourcePath = Read-Host -Prompt "Enter the source path"
    $OfficeClientEdition = Read-Host -Prompt "Enter the office client edition (32 or 64)"
    $channel = Read-Host -Prompt "Enter the update channel"
    $productIDs = Read-Host -Prompt "Enter the product IDs (comma-separated)"
    $languageIDs = Read-Host -Prompt "Enter the language IDs (comma-separated)"
    $excludeAppIDs = Read-Host -Prompt "Enter the app IDs to exclude (comma-separated)"
    #$updatePath = Read-Host -Prompt "Enter the update path"
    $updatesEnabled = [bool]::Parse((Read-Host -Prompt "Enable updates? (true or false)"))
    $DisplayLevel = Read-Host -Prompt "Enter the display level"
    $AcceptEULA = [bool]::Parse((Read-Host -Prompt "Accept EULA? (true or false)"))
    $ProductIDs = $ProductIDs.Split(",")
    $LanguageIDs = $LanguageIDs.Split(",")
    $ExcludeAppIDs = $ExcludeAppIDs.Split(",")

    $configurationName = Read-Host "Enter your configuration name (default is configuration)"
        if ($configurationName -eq "") {
            $configurationName = "configuration"
        }

    $generatedConfig = @"
<Configuration ID="$(New-Guid)">
    <Add OfficeClientEdition="$OfficeClientEdition" Channel="$channel">
    $(foreach ($product in $productIDs) {
        "   <Product ID=`"$product`">`n"
        foreach ($language in $languageIDs) {
        "       <Language ID=`"$language`" />`n"
        }
        foreach ($excludeApp in $excludeAppIDs) {
            "   <ExcludeApp ID=`"$excludeApp`" />`n"
            }
            "</Product>"
        })
            </Add>
            <Updates Enabled="$updatesEnabled" />
            <RemoveMSI />
            <Display Level="$displayLevel" AcceptEULA="$acceptEULA" />
    </Configuration>
"@
    $generatedConfig | Out-File -FilePath .\$configurationName.xml -Encoding utf8
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

function main(){
    Clear-Host
    header
    mainSelector -Title "office-c2r-installer Selector"
    $selection = Read-Host "Please use keyboard to make a selection."
        
    switch ($selection) {
        '1' { customConfigOnline }
        '2' { preConfigOnline }
        '3' { customConfigOffline }
        '4' { preConfigOffline }
        'g' { configGenerator }
        'q' { return }
    }
}

main
