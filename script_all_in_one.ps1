function header() {
    # todo: make header
}

function customConfigOnline() {
    Write-Host "[1]: Install Office with specified configuration (online) 1"

}

function preConfigOnline() {
    Write-Host "[1]: Install Office with specified configuration (online) 2"
}

function customConfigOffline() {
    Write-Host "[1]: Install Office with specified configuration (online)3"
}

function preConfigOffline() {
    Write-Host "[1]: Install Office with specified configuration (online)4"
}

function configGenerator() {
    
}

function c2rExtractor {
    # todo...
}

function mainSelector {
    param (
        [string]$Title = 'office-c2r-installer Selector'
    )
    Clear-Host
    Write-Host "==================== office-c2r-installer ======================="
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

mainSelector –Title 'office-c2r-installer Selector'
$selection = Read-Host "Please use keyboard to make a selection"

switch ($selection) {
    '1' { customConfigOnline }
    '2' { preConfigOnline }
    '3' { customConfigOffline }
    '4' { preConfigOffline }
    'g' { generateConfig }
    'q' { return }
}
