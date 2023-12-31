#Requires -Version 5.1
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function header() {
    Write-Output "office-c2r-installer (WIP)"
    Write-Output "Version 0.0.1 alpha 2 (partial working preview - UNSTABLE)"
    Write-Output "https://github.com/dtcu0ng/office-c2r-installer"
}
# codes in below this line is pretty simple, just the selector and the bootstrap code in main()
# should commented out all the code need to fix or not working before push to Github.
# this prevent CI failed to run because known not working feature.

# todo: integrate CI to fix syntax.

function mainSelector {
    Write-Output "================================================================="
    Write-Output "[1]: Install Office with specified configuration (online)"
    Write-Output "[2]: Install Office with pre-configured configuration (online)"
    Write-Output "================================================================="
    Write-Output "[3]: Install Office with specified configuration (offline)"
    Write-Output "[4]: Install Office with pre-configured configuration (offline)"
    Write-Output "================================================================="
    Write-Output "[G]: Generate configuration"
    Write-Output "[Q]: Quit"
    Write-Output "================================================================="
}

function languageSelector() {
    Write-Output "Language selector:"
    Write-Output "(1) English"
    Write-Output "(2) Vietnamese"
    $mainLanguage = Read-Host "Please choose your preferred language:"
    if ($mainLanguage -eq "") {
        Write-Output "Please select a language"
    } else {
        mainSelector
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
}
function customConfigOnline() {
    Clear-Host
    $configPath = Read-Host "Please enter path of the configuration file you want to install. Type q to go back"
    if ($configPath -eq "q") {
        main
    } elseif (!(Test-Path -Path $configPath))  {
        Clear-Host
        Write-Output "Configuration file not found. Please double-check your configuration file and try again."
    } else {
        Clear-Host
        Write-Output "Starting Office setup with a specified configuration: $configPath"
        Write-Output "Please do not close any window or process during installation."
        Start-Process -FilePath "./files/setup.exe" -ArgumentList "/configure $configPath" -Wait
        Clear-Host
        Write-Output "Setup has been exited, please check that your Office installation is successfully installed."
        Write-Output "If not, please try again. If the problem persists, please open an issue on Github"
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
        Write-Output "Configuration file not found. Please double-check your configuration file and try again."
    } else {
        $installerPath = Read-Host "Please specify your Office installer path. Type q to go back to main menu."
        if ($installPath -eq "q") {
            main
        } elseif (!(Test-Path -Path $installerPath)) {
            Write-Output "Installer path not found. Please double check your installer path and try again."
        } else {
            Write-Output "WIP"
        }
    }
}

function preConfigOffline() {
    Clear-Host
    $installerPath = Read-Host "Please specify your Office installer path. Type q to go back to main menu."
    if ($installerPath -eq "q") {
        main
    } elseif (!(Test-Path -Path $configPath)) {
        Write-Output "Configuration file not found. Please double-check your configuration file and try again."
    } else {
        preConfigMenu
    }

}

# WIP, will fix it later

function configGenerator() {
# Create XML configuration (WIP)
    #$SourcePath = Read-Host -Prompt "Enter the source path"
    Write-Output "Do you want to install Office via your local source file?"
    $SelSourcePath = Read-Host -Prompt "If you want to, just specify the source path here. If you want to install via Office CDN, press 'o'"
    if ($SelSourcePath -eq "o") {
        $SelSourcePath = ""
        Write-Output "Will install via Office CDN."
    }
    Write-Output "Please select architecture for your Office installation."
    Write-Output "[1]: 64bit (x64)"
    Write-Output "[2]: 32bit (x86)"
    $confGenArch = Read-Host "Please use keyboard to make a selection"
    switch ($confGenArch) {
        '1' { $OfficeClientEdition = 64 }
        '2' { $OfficeClientEdition = 32 }
    }
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

    $configurationName = Read-Host "Enter your configuration name (default is Configuration)"
        if ($configurationName -eq "") {
            $configurationName = "Configuration"
        }

    $generatedConfig = @"
<Configuration ID="$(New-Guid)">
    <Add OfficeClientEdition="$OfficeClientEdition" Channel="$channel" MigrateArch="$migrateArch">
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
    Write-Output "WIP"
}

function preConfigMenu {
# todo: list of pre-configured configuration, can select, use en-us as main language.
# two type of configuration are available - full or mininal (word, excel, powerpoint,...)
    Write-Output "WIP"
}
function main(){
    Clear-Host
    header
    languageSelector
}
main
