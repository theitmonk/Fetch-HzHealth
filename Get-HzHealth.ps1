param(
    [Parameter(Mandatory = $true)][string]$vCenter,
    [Parameter(Mandatory = $true)][string]$Broker,
    [Parameter(Mandatory = $true)][string]$OndrivePath
)

$hvServerData = $null
$viServerData = $null
$hcPathHome = Split-Path -Parent $MyInvocation.MyCommand.Definition
$hcPathOneDrive = $OndrivePath
$greenPicUrl = "https://raw.githubusercontent.com/theitmonk/Get-HzHealth/main/img/green.png"
$redPicUrl = "https://raw.githubusercontent.com/theitmonk/Get-HzHealth/main/img/red.png"
$unknownPicUrl = "https://raw.githubusercontent.com/theitmonk/Get-HzHealth/main/img/unknown.png"

$statusRedBgColor = " style='background-color:Red;'"
$statusRedFontColor = " style='color:Red'"
$statusGreenBgColor = " style='background-color:Green'"
$statusGreenFontColor = " style='color:Green'"
$dekstopGreenStates = "AVAILABLE", "CONNECTED", "DISCONNECTED", "MAINTENANCE", "UNASSIGNED_USER_CONNECTED", "UNASSIGNED_USER_DISCONNECTED", "PROVISIONED"
$generalGreenStates = "Available", "Connected", "True", "OK"
$numValidationSet = "Hosts", "DataStores"
$checkArray = "ConnBroker", "Domain", "Composer", "EventDb", "Desktops", "Desktops", "vCenter", "Hosts", "Datastores"
$global:tempNodeResult = @{}

function getVsphereHc {
    Param(
        [Parameter(Mandatory = $true)][System.Object]$ViObject
    )

    $viHealth = @{}
    $vcenterStatus = Invoke-WebRequest "https://$($ViObject.Name)/ui"

    # ESXi host status
    $viHealth.Hosts = Get-VMHost -Server $ViObject | Select-Object `
        Name, `
    @{N = "Status"; E = { $_.ConnectionState } }, `
    @{N = "Cpu Usage %"; E = { [math]::Round($_.CpuUsageMhz / $_.CpuTotalMhz * 100) } }, `
    @{N = "Ram Usage %"; E = { [math]::Round($_.MemoryUsageGB / $_.MemoryTotalGB * 100) } } 
    
    # DataStore status
    $viHealth.DataStores = Get-Datastore  -Name VDI* -Server $ViObject | Select-Object  `
        Name, `
    @{N = "Status"; E = { $_.State } }, `
    @{N = "Used Space %"; E = { [math]::Round(($_.CapacityGB - $_.FreeSpaceGB) / $_.CapacityGB * 100) } } `
    | Sort-Object name

    # vCenter status
    $viHealth.VCenter = $ViObject | Select-Object  `
    @{N = "Name"; E = { $_.Name } }, `
    @{N = "Status"; E = { $_.IsConnected } }, `
    @{N = "Version"; E = { $_.Version } }, `
    @{N = "StatusCode"; E = { $vcenterStatus.StatusCode } }

    return $viHealth
}

function getHorizonHc {

    Param(
        [Parameter(Mandatory = $true)][System.Object]$HvObject
    )

    $hvData = $HvObject.ExtensionData
    $hvHealth = @{}

    # Connection Server status
    $hvHealth.ConnBroker = $hvData.ConnectionServerHealth.ConnectionServerHealth_List() | Select-Object `
        Name, `
        Status, `
        Version, `
        Build, `
    @{N = "ReplicationStatus"; E = { $_.ReplicationStatus.Status } }
    
    # Domain status    
    $hvDomain = ($hvData.ADDomainHealth.ADDomainHealth_List())[1]
    $hvHealth.Domain = $hvDomain.ConnectionServerState | Select-Object `
    @{N = "Domain"; E = { $hvDomain.DnsName } }, `
    @{N = "Status"; E = { $_.Status } }, `
    @{N = "CB"; E = { $_.ConnectionServerName } }, `
    @{N = "Trust"; E = { $_.TrustRelationship } }

    # Event status
    $hvHealth.EventDb = ($hvData.EventDatabaseHealth.EventDatabaseHealth_Get()).data | Select-Object `
        ServerName, `
    @{N = "Status"; E = { $_.State } }, `
        DatabaseName, `
        Error
    
    # Composer status
    $hvComposer = $hvData.ViewComposerHealth.ViewComposerHealth_List()
    $hvHealth.Composer = $hvComposer.ConnectionServerData | Select-Object `
    @{N = "CP"; E = { $hvComposer.ServerName } }, `
    @{N = "Status"; E = { $_.Status } }, `
    @{N = "CB"; E = { $_.Name } }, `
    @{N = "Thump"; E = { $_.ThumbprintAccepted } }, `
    @{N = "Version"; E = { $hvComposer.data.version } }


    # Desktop status1
    $hvHealth.DesktopsData = (Get-HVMachineSummary -HvServer $HvObject).Base | Group-Object BasicState  | Select-Object `
    @{N = "Status"; E = { $_.Name } }, `
    @{N = "ShellCount"; E = { $_.Count } } | Sort-Object Status

    # Desktop status1
    $totalDekstops = ( $($hvHealth.DesktopsData).ShellCount | Measure-Object -Sum ).Sum
    $probDekstops = ( $($hvHealth.DesktopsData | Where-Object { $_.Status -notin $dekstopGreenStates }).Shellcount | Measure-Object -Sum ).Sum
    if ($probDekstops -ge 5 ) {
        $desktopsStatus = "ActionRequired"
    }
    else {
        $desktopsStatus = "OK"
    }
    $hvHealth.Desktops = "" | Select-Object  `
    @{N = "TotalDesktops"; E = { $totalDekstops } }, `
    @{N = "Status"; E = { $desktopsStatus } }, `
    @{N = "PrDesktops"; E = { $probDekstops } } 

    return $hvHealth

}
function styleIt {
    param (
        [Parameter(Mandatory = $true)][System.Object] $Object,
        [Parameter(Mandatory = $false)][bool] $NumValidation = $false,
        [Parameter(Mandatory = $false)][string] $Type = $null
    )
    $tempObject = $Object | ConvertTo-Html -Fragment
    $fcnResult = $tempObject | ForEach-Object {
        $statusRegex = "(?<=<td>[^<,/,>]+</td><td>)\w+(?=</td>)"
        $Matches = $null;
        $temp = $_ -match $statusRegex
        if ($Matches.Values -in $generalGreenStates ) {
            $statusBgColor = $statusGreenBgColor
        }
        else {
            $statusBgColor = $statusRedBgColor
        }
        if ($numValidation -or $numValidationSet -in $Type) {
            $numBgColor = $statusRedBgColor
        }
        $_ -replace "(?<=<tr><td>[^<,/,>]+</td><td)(?=>\D+</td>)", $statusBgColor -replace "(?<=<td)(?=>(8[5-9]|[9][0-9]|100|404)</td>)", $numBgColor
    }
    return $fcnResult
}

function findNodeHealth {
    param (
        [Parameter(Mandatory = $true)][array]$Object
    )
    $result = $true
    if ($htmlObject -match $statusRedBgColor) {
        $result = $result -and $false
    }
    return $result
}
function createHtmlBody {
    param (
        [Parameter(Mandatory = $true)][System.Object]$Object
    )
    $objectBody = $null
    $Object.Keys | Sort-Object | ForEach-Object {
        $objectHead = $null
        $keyObject = "$($_)Data" 
        Set-Variable -Name $keyObject -Value (styleIt -Object $($Object.$_) -Type $_)
        $styledData = Get-Variable -Name $keyObject -ValueOnly
        if ($_ -ne "DesktopsData") {
            $objectHead = "<h2>$_</h2>"
        }
        $objectContent = "" + $styledData
        $objectBody += $objectHead + $objectContent

        $tempNodeResult.$_ = findNodeHealth $styledData
    }
    return $objectBody
}

function createErrorTable {
    param (
        [Parameter(Mandatory = $true)][ValidateSet("Broker", "vCenter")][string]$Type,
        [Parameter(Mandatory = $true)][string]$Name
    )
    $ErrorTable = "" | Select-Object `
    @{N = "Name"; E = { $Name } }, `
    @{N = "Status"; E = { "404" } }, `
    @{N = "Discription"; E = { "$Type Unavailable" } }
    $body = "<h1 $statusRedFontColor>Attention required! $Type unreachable</h1>" + (styleIt $ErrorTable -numValidation $true)
    return $body
}

$vmaCred = Import-Clixml "$hcPathHome\vmacred.xml"

$hvServerData = Connect-HVServer -Server $Broker -Credential $vmaCred -ErrorVariable "hvError"
if ($null -ne $hvServerData) {
    $hvData = getHorizonHc -HvObject $hvServerData
    $hvBody = createHtmlBody -Object $hvData
}
else {
    $hvBody = createErrorTable -Type Broker -Name $Broker
    $tempNodeResult.ConnBroker = $false
    $tempNodeResult.Domain = $tempNodeResult.EventDb = $tempNodeResult.Desktops = $tempNodeResult.Composer = "Unknown"
}

$viServerData = Connect-VIServer -Server $vCenter -Credential $vmaCred -ErrorVariable "viError"
if ($null -ne $viServerData) {
    $viData = getVsphereHc -ViObject $viServerData
    $viBody = createHtmlBody -Object $viData

}
else {
    $viBody = createErrorTable -Type vCenter -Name $vCenter
    $tempNodeResult.Hosts = $tempNodeResult.Datastores = "Unknown"
    $tempNodeResult.vCenter = $false
}

# HTML Report
$fBody = $hvBody + $viBody
$shortTime = Get-Date -DisplayHint DateTime -Format "MMddyyHHmm"
$longTime = Get-Date -DisplayHint DateTime
if ($fBody -match 'color:red' -or $tempNodeResult.ContainsValue($false)) {
    $gStatus = "ActionRequired!!"
    $hcStatus = "<b $statusRedFontColor>$gStatus</b>"
    $tempNodeResult.GlobalStatus = $false
}
else {
    $gStatus = "Green"
    $hcStatus = "<b $statusGreenFontColor>$gStatus</b>"
    $tempNodeResult.GlobalStatus = $true 
}
$tempNodeResult.Subject = "VDI3 HC : $longTime Status: $gStatus"
$tempNodeResult.GlobalStatusMessage = $gStatus
$body = "<h1> VDI3 HC : $longTime Status: $hcStatus </h1>" + $fBody
$fResult = ConvertTo-Html -Body $body -Head $Header -Title "VDI3 HC $shortTime "
$eMailSubject = "Daily VDI Farm Health check | $gStatus"
$eMailSubject | Out-File "$hcPathOneDrive\$shortTime-subject.txt"
$fResult | Out-File "$hcPathOneDrive\$shortTime.html"

# Teams adaptive card JSON 
$reportJson = Get-content "$hcPathHome\report.json"
$reportJson = $reportJson -replace "<<time>>", $longTime -replace "<<datetime>>", $shortTime -replace "<<GenaralStatusMessage>>", $gStatus
if ($tempNodeResult.GlobalStatus) {
    $rData = "good"
}
else {
    $rData = "attention"
}
$reportJson = $reportJson -replace "<<GenaralStatusColor>>", $rData

$checkArray | ForEach-Object {
    switch ($tempNodeResult.$_) {
        "True" { $replaceString = "$greenPicUrl" } 
        "False" { $replaceString = "$redPicUrl" }
        "Unknown" { $replaceString = "$unknownPicUrl" }
    }
    $reportJson = $reportJson -replace "<<$_>>", $replaceString
}
$reportJson | Out-File "$hcPathOneDrive\json\$shortTime.txt"

Disconnect-VIServer -Confirm:$false -ErrorAction SilentlyContinue
Disconnect-HVServer -Confirm:$false -ErrorAction SilentlyContinue

