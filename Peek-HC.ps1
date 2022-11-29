
$vmacred = Import-Clixml C:\Scripts\Peek-HC\vmacred.xml
$lccred = Import-Clixml C:\Scripts\Peek-HC\lacred.xml
$hvServerData = $null
$vma = $null
$statusRed = " style='background-color:Red;'"
$statusRed1 = " style='color:Red'"
$statusGreen = " style='background-color:Green'"
$statusGreen1 = " style='color:Green'"
# Connect vCenter and Broker

# vCenter

function getVsphereHc {
    # Parameter help description
    Param(
        [Parameter(Mandatory=$true)][System.Object] $vma
    )
    $vmhc = @{}
    $vmas = Invoke-WebRequest "https://$($vma.Name)/ui"
    $vmhc.hosts = Get-VMHost | select Name,ConnectionState,@{N="Cpu Usage %";E={[math]::Round($_.CpuUsageMhz/$_.CpuTotalMhz*100)}},@{N="Ram Usage %";E={[math]::Round($_.MemoryUsageGB/$_.MemoryTotalGB*100)}} 
    $vmhc.datastores = Get-Datastore VDI* | select  Name,State,@{N="Used Space %";E={[math]::Round(($_.CapacityGB-$_.FreeSpaceGB)/$_.CapacityGB*100)}} | sort name
    $vmhc.vcenter = $vma | select @{N="Name";E={$_.Name}},@{N="Status";E={$_.IsConnected}},@{N="Version";E={$_.Version}},@{N="StatusCode";E={$vmas.StatusCode}}

    return $vmhc
}

function getHorizonHc{

    Param(
        [Parameter(Mandatory=$true)][System.Object] $hvs
    )

    $hvdata = $hvs.ExtensionData
    $hvhc = @{}
    $hvhc.ConnBroker = $hvdata.ConnectionServerHealth.ConnectionServerHealth_List() | select Name,Status,Version,Build,@{N="ReplicationStatus";E={$_.ReplicationStatus.Status}}
    $hvd = ($hvdata.ADDomainHealth.ADDomainHealth_List())[1]
    $hvhc.Domain = $hvd.ConnectionServerState | select @{N="Domain";E={$hvd.DnsName}},@{N="Status";E={$_.Status}},@{N="CB";E={$_.ConnectionServerName}},@{N="Trust";E={$_.TrustRelationship}}
    $hvhc.Event = ($hvdata.EventDatabaseHealth.EventDatabaseHealth_Get()).data | select ServerName,State,DatabaseName,Error
    $cmp = $hvdata.ViewComposerHealth.ViewComposerHealth_List()
    $hvhc.Composer = $cmp.ConnectionServerData | select @{N="CP";E={$cmp.ServerName}},@{N="Status";E={$_.Status}},@{N="CB";E={$_.Name}},@{N="Thump";E={$_.ThumbprintAccepted}},@{N="Version";E={$cpn.data.version}}
    $hvhc.Desktops = (Get-HVMachineSummary).Base | group BasicState  | select @{N="State";E={$_.Name}},@{N="ShellCount";E={$_.Count}} 
    $hvhc.ShellCount = ($($hvhc.Desktops).ShellCount | Measure-Object -Sum).Sum
    
    return $hvhc

}

function styleIt {
    param (
        [Parameter(Mandatory=$true)][System.Object] $Object,
        [Parameter(Mandatory=$false)][bool] $numValidation = $false
    )
    $tempObject = $Object | ConvertTo-Html -Fragment
    $fcnResult = $tempObject | % {
        $statusRegex = "(?<=<td>.+</td><td>)\w+(?=</td>)"
        $Matches=$null;
        $temp = $_ -match $statusRegex
        if ($Matches.Values -in "Available","Connected","True","OK"  ){
            $statusBgColor = $statusGreen
        }
        else{
            $statusBgColor = $statusRed
        }
        if($numValidation){
            $numBgColor = $statusRed
        }
        $_ -replace "(?<=<tr><td>[^<,/,>]+</td><td)(?=>\D+</td>)", $statusBgColor -replace "(?<=<td)(?=>(8[5-9]|[9][0-9]|100|404)</td>)", $numBgColor
    }
   return $fcnResult
}

function findNodeHealth {
    param (
        [Parameter(Mandatory=$true)][array]$htmlObject,
        [Parameter(Mandatory=$false)][System.Object]$desktopData = $null
    )
    $result = $true
    if($null -ne $desktopData){
        $probShellCount = (($desktopData | ? {$_.State -notin "Available","Disconnected","Connected","Maintenance"} | select ShellCount).Shellcount | measure -Sum).Sum
        if($probShellCount -gt 15){
            $result = $result -and $false
        }
    }
    if($htmlObject -match $statusRed){
        $result = $result -and $false
    }
    else{
        $result = $result -and $true
    }
    return $result
}

$Header = @"
<style>
table {
    font-family: arial, sans-serif;
    border-collapse: collapse;
    font-size:12px;
    width: 50%;
    margin-left: auto;
    margin-right: auto;
  }
  
  td, th {
    border: 1px solid #dddddd;
    text-align: center;
    padding: 5px;
  }
 th {
    background-color: blue;
    color:white;
 }
 h1,h2,html,table{
    text-align : center;
 }
</style>
"@
# 
# 
    $tempNodeResult=@{}
    $hvServerData = Connect-HVServer -Server port142vms -Credential $vmacred -ErrorAction SilentlyContinue
    if($hvServerData -ne $null){
        $hvData = getHorizonHc -hvs $hvServerData
        $connData = styleIt $($hvData.ConnBroker)
        $domainData = styleIt $($hvData.Domain)
        $eventData = styleIt $($hvData.Event)
        $composerData = styleIt $($hvData.Composer)
        $sortedDesktopData = $hvData.Desktops | sort State
        $desktopData = styleIt $($sortedDesktopData)
        
        $tempNodeResult.ConnServer = findNodeHealth $connData
        $tempNodeResult.Domain = findNodeHealth $domainData
        $tempNodeResult.EventDb = findNodeHealth $eventData
        $tempNodeResult.Composer = findNodeHealth $composerData
        $tempNodeResult.Desktops = findNodeHealth $desktopData -desktopData $($hvData.Desktops)
        $desktopStatusColor = $null
        if(!$tempNodeResult.Desktops){
            $desktopStatusColor = $statusRed1
        }

        $hbody = "<h2>Connection Servers</h2>" + "" + $connData + "" + "<h2>Domain</h2>" + "" + $domainData +"" + "<h2>Composer</h2>" + "" + $composerData + "<h2>EventDb</h2>" + "" + $eventdata + "<h2$desktopStatusColor>Desktops</h2>" + "" + $desktopData
    }
    else{
        $h = "" | select @{N="Name";E={"PORT142VMS"}},@{N="Status";E={"404"}},@{N="Discription";E={"CB Server Unavailable"}}
        $hbody = "<h1 $statusRed1>!Attention required!! Broker unreachable!!</h1>" + (styleIt $h -numValidation $true)
        $tempNodeResult.ConnServer = $false
        $tempNodeResult.Domain =  $tempNodeResult.Domain = $tempNodeResult.EventDb = $tempNodeResult.Desktops =  $tempNodeResult.Composer = "Unknown"
    }



    $vma = Connect-VIServer -Server port140vma -Credential $vmacred -ErrorAction SilentlyContinue

    if($null -ne $vma){
        $viData = getVsphereHc -vma $vma
        $hostData = styleIt $($viData.hosts) -numValidation $true
        $vcData = styleIt $($viData.vcenter) 
        $storeData = styleIt $($viData.datastores) -numValidation $true
    
        $tempNodeResult.vCenter = findNodeHealth $vcData
        $tempNodeResult.Hosts = findNodeHealth $hostData
        $tempNodeResult.Datastores = findNodeHealth $storeData
        
        $vbody = "<h2>vCenter</h2>" + "" + $vcdata + "" + "<h2>Hosts</h2>" + "" + $hostdata +"" + "<h2>Datastores</h2>" + "" + $storedata    
    }
    else {
        $v = "" | select @{N="Name";E={"PORT140VMA"}},@{N="Status";E={"404"}},@{N="Discription";E={"vCenter Unavailable"}}
        $vbody = "<h1 $statusRed1>!Attention required!! vCenter unreachable!!</h1>" + (styleIt $v -numValidation $true)
        $tempNodeResult.Hosts = $tempNodeResult.Datastores = $false
        $tempNodeResult.vCenter = "Unknown"
    }

$body = $hbody+$vbody
$d1 = get-date -DisplayHint DateTime -Format "MMddyyHHmm"
$d2 = get-date -DisplayHint DateTime
if($body -match 'color:red'){
    $s = "ActionRequired!!"
    $hcstatus = "<b $statusRed1 >$s</b>"
    $tempNodeResult.GlobalStatus = $false
}
else{
    $s = "Green"
    $hcstatus = "<b $statusGreen1 >$s</b>"
    $tempNodeResult.GlobalStatus = $true 

}
$tempNodeResult.Subject = "VDI3 HC : $d2 Status: $s"
$tempNodeResult.GlobalStatusMessage = $s 

$body = "<h1> VDI3 HC : $d2 Status: $hcstatus </h1>" + $body
$fresult = ConvertTo-Html -Body $body -Head $Header -Title "VDI3 HC $d1 "
$eMailSubject = "Daily VDI Farm Health check | $s"
$eMailSubject | Out-File "D:\TVG_VDI_HC\OneDrive - kyndryl\TVG VDI HC\$d1-subject.txt"
$fresult | Out-File "D:\TVG_VDI_HC\OneDrive - kyndryl\TVG VDI HC\$d1.html"
# $tempNodeResult | ConvertTo-Json | Out-File "D:\TVG_VDI_HC\OneDrive - kyndryl\TVG VDI HC\Subject\$d1.json"

$reportJson = Get-content C:\Scripts\Peek-HC\report.json

$reportJson = $reportJson -replace "<<time>>", $d2
$reportJson = $reportJson -replace "<<datetime>>", $d1
$reportJson = $reportJson -replace "<<GenaralStatusMessage>>", $s
if($tempNodeResult.GlobalStatus){
    $rData = "good"
}
else {
    $rData = "attention"
}
$reportJson = $reportJson -replace "<<GenaralStatusColor>>", $rData

$checkArray = "ConnServer","Domain","Composer","EventDb","Desktops","vCenter","Hosts","Datastores"
$checkArray | % {
    switch ($tempNodeResult.$_) {
        "True" { $replaceString = "<<StatusColorGreen>>" } 
        "False" {$replaceString = "<<StatusColorRed>>"}
        "Unknown" {$replaceString = "<<StatusColorUnknown>>"}
    }
    $reportJson = $reportJson -replace "<<$_>>", $replaceString
}
$reportJson | out-File "D:\TVG_VDI_HC\OneDrive - kyndryl\TVG VDI HC\json\$d1.txt"

Disconnect-HVServer -Confirm:$false
Disconnect-VIServer -Confirm:$false