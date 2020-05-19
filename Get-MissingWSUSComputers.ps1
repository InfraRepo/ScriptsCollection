<#
.SYNOPSIS
    Create report of AD computers which don't show up in WSUS console. Also finds computers which are disabled in AD but still exist in WSUS console.
.DESCRIPTION
    Gets computer list from AD and compares it against computer list in WSUS.

    Results are exported to files and, optionally, can be displayed on the screen in a grid form.
    You can set the scope with -Scope parameter to limit results to only servers or only computers.
.EXAMPLE
    PS C:\> .\Get-MissingWSUSComputers.ps1 -Verbose -Scope Servers
    Only check for server objects from Active Directory, saves results to files by default.

.EXAMPLE
    PS C:\> .\Get-MissingWSUSComputers.ps1 -Grid
    Additionally show grid with all results.
.INPUTS
    None.
.OUTPUTS
    CSV files with computer list - from AD and from WSUS.
.LINK
#>

[cmdletbinding()]
Param(
    [string]$WSUSServer = 'localhost',
    [int]$WSUSPort = 8530,
    [bool]$UseSSL = $false,
    [ValidateSet('All','Servers','Computers')]
    [string]$ScopeLimit = 'All',
    [string]$ExportDir = 'Output',
    [string]$ExportMissingFromWSUS = 'export-MissingFromWSUS.csv',
    [string]$ExportStaleInWSUS = 'export-StaleInWSUS.csv',
    [switch]$Grid
)
Function Get-MissingFromWSUS($ADComputers,$WSUSComputers){
    Write-verbose "Checking misssing from WSUS"
    switch ($ScopeLimit) {
        'Servers' {
            write-verbose "Checking servers"
            $ADComputers = $ADComputers | Where-Object {$_.OperatingSystem -match 'server'}
        }
        'Computers' {
            write-verbose "Checking regular computers"
            $ADComputers = $ADComputers | Where-Object {$_.OperatingSystem -notmatch 'server'}
        }
    }
    $result = $ADComputers | Where-Object { $WSUSComputers.FullDomainName -notcontains $_.DNSHostName } | Sort-Object
    $result
}

Function Get-StaleInWSUS($ADComputers,$WSUSComputers){
    $result = $WSUSComputers | Where-Object { $ADComputers.DNSHostName -notcontains $_.FullDomainName } | Sort-Object
    $result
}

# define output directory path
$ExportDirPath = Join-Path $PSScriptRoot $ExportDir

# create output directory if doesn't exist
If(!(test-path $ExportDirPath)){
    New-Item -ItemType Directory -Force -Path $ExportDirPath
}

try{
    #Load necessary assemblies
    Write-Verbose 'Loading assembly...'
    $AssemblyLoaded = [reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration")
    If($AssemblyLoaded){
        Write-Verbose '  Done.'
    }
    else{
        throw '  Assembly could not be loaded!'
    }

    #Connect to WSUS
    Write-Verbose 'Connecting to WSUS...'
    $WSUSConnection = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer($WSUSServer, $UseSSL, $WSUSPort)
    Write-Verbose '  Done.'

    #Get WSUS Computers
    Write-Verbose 'Getting WSUS computer list...'
    $WSUSComputers = $WSUSConnection.GetComputerTargets()
    Write-Verbose '  Done.'

    #Get AD computers
    Write-Verbose 'Getting AD computer list...'
    $ADComputers = Get-ADComputer -Filter {OperatingSystem -like '*windows*'} -Properties DNSHostName,OperatingSystem
    Write-Verbose '  Done.'

    #Filter computers
    $EnabledInAD = $ADComputers | Where-Object {$_.Enabled -eq $true}
    $DisabledInAD = $ADComputers | Where-Object {$_.Enabled -eq $true}

    #Search for objects
    Write-Verbose 'Comparing objects...'
    $MissingFromWSUS = Get-MissingFromWSUS $EnabledInAD $WSUSComputers
    $StaleInWSUS = Get-StaleInWSUS $DisabledInAD $WSUSComputers
    Write-Verbose '  Done.'

    #Export data
    Write-Verbose 'Exporting data...'
    $MissingFromWSUS | Select-Object Name | Export-CSV -Path (Join-Path $ExportDirPath $ExportMissingFromWSUS) -Encoding UTF8 -NoTypeInformation
    $StaleInWSUS | Select-Object FullDomainName,IPAddress,LastSyncTime,LastSyncResult,LastReportedStatusTime | Export-CSV -Path (Join-Path $ExportDirPath $ExportStaleInWSUS) -Encoding UTF8 -NoTypeInformation
    Write-Verbose '  Done.'

    #Show results on the screen as a grid
    If($Grid){
        $WSUSComputers | Select-Object FullDomainName,IPAddress,LastSyncTime,LastSyncResult,LastReportedStatusTime | Out-GridView -Title "All WSUS Computers: [$(($WSUSComputers | Measure-Object).Count)]"
        $ADComputers | Out-GridView -Title "All ADComputers: [$(($ADComputers | Measure-Object).Count)]"
        $MissingFromWSUS | Out-GridView -Title "Hosts not found in WSUS: [$(($MissingFromWSUS | Measure-Object).Count)]"
        $StaleInWSUS | Out-GridView -Title "Hosts found in WSUS and account disabled in AD: [$(($StaleInWSUS | Measure-Object).Count)]"
    }
}
catch{
    Write-Error "  Failed! Error message: $($_.Exception.Message)"
}