param (
     [string]$InputIP = "0.0.0.0"  ,
     [string]$LOGGING = "NO"
)
#$LOGGING = 'YES'
#$inputIp = "holiday"

$ScriptVersion = " -- Version: 1.2"

# COMMON coding
CLS
$InformationPreference = "Continue"
$WarningPreference = "Continue"
$ErrorActionPreference = "Stop"

$Node = " -- Node: " + $env:COMPUTERNAME
$d = Get-Date
$Datum = " -- Date: " + $d.ToShortDateString()
$Tijd = " -- Time: " + $d.ToShortTimeString()

$myname = $MyInvocation.MyCommand.Name
$p = $myname.Split(".")
$process = $p[0]
$FullScriptName = $MyInvocation.MyCommand.Definition
$mypath = $FullScriptName.Replace($MyName, "")

$LocalInitVar = $mypath + "InitVar.PS1"
& "$LocalInitVar" -JSON Silent
if (!$ADHC_InitSuccessfull) {
    # Write-Warning "YES"
    throw $ADHC_InitError
}
$MessageList = @()
$Global:ResultObject = [PSCustomObject] [ordered] @{MessageList = $Messagelist;
                                            IPinput = $InputIP;
                                            IPaddress = "n/a";
                                            MACaddress = "n/a";
                                            IPcached = $false;
                                            IPpingable = $false;
                                            ItsMe = $false;
                                            Status = "Unknown";
                                            StatusCode = -1}
function AddMessage ([string]$level, [string]$msg) {
    $msgentry = [PSCustomObject] [ordered] @{Level = $level;
                                             Message = $msg}
    $global:ResultObject.MessageList += $msgentry
    # Write-Host $msg
    
    return  
}


if ($LOGGING -eq "YES") {$log = $true} else {$log = $false}

$Scriptmsg = "Directory " + $mypath + " -- PowerShell script " + $MyName + $ScriptVersion + $Datum + $Tijd +$Node
AddMessage "I" $Scriptmsg                                           

if ($log) {
    $thisdate = Get-Date
    AddMessage "I" "==> START $thisdate"
    AddMessage "I" "Host = $inputIP, Logging = $LOGGING" 
}

# END OF COMMON CODING

$scripterror = $false

# Determine your own attributes

try {
    $Global:ResultObject.itsme = $false
    if ($log) {
        AddMessage "I" "Get IP/Mac address from myself"
    }
    $ComputerName = $env:computername
    $OrgSettings = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $ComputerName -EA Stop | ? { $_.DNSDomain -eq "fritz.box" }
    $myip = $OrgSettings.IPAddress[0]

    $a = Get-NetAdapter | ? {$_.Name -eq "Wi-Fi"}
    $mymac = $a.MacAddress.Replace("-",":")
}
catch {    
    $scripterror = $true
    $errortext = $error[0]
    $scripterrormsg = "Getting IP/MAC address failed - $errortext"
    if ($log) {
       AddMessage "E" "$scripterrormsg"
    }
    
}

# PING first, ARP -A after that

if (!$scripterror) {
    try {
        if ($log) {
            AddMessage "I" "Perform PING"
        }
        try {
            $ipnum = $null
            if ($inputip -match "\w+") {
                $ipx = [System.Net.Dns]::GetHostAddresses($InputIp.Trim())
                if ($ipx) {
                    write-host "ipx"
                    $Ipnum = $ipx.IPAddressToSTring
                    $Global:ResultObject.IPaddress = $ipnum                    
                }
                else {
                    write-host "no ipx"
                    $Global:ResultObject.IPaddress = $inputip  
                }
            }
            else {
                $Global:ResultObject.IPaddress = $inputip
            }
            if ($myip -eq $Global:ResultObject.IPaddress) {
                $Global:ResultObject.itsme = $true
            }
            $ping = Test-Connection -COmputerName $Global:ResultObject.IPaddress -Count 1
            $Global:ResultObject.IPpingable = $true
            
        }
        catch {
            $Global:ResultObject.IPpingable = $false
            $Global:ResultObject.IPaddress = $inputip
        }
    }
    catch {
        $scripterror = $true
        $errortext = $error[0]
        
        $scripterrormsg = "Ping failed - $errortext"
        if ($log) {
           AddMessage "E" "$scripterrormsg"
        }

    }

    
}

# Get ARP table

if (!$scripterror) {

    try { 
        if ($log) {
            AddMessage "I" "Get ARP -A info"
        }
        $Global:ResultObject.IPcached = $false
        $arpa = (arp -a) 
        foreach ($line in $arpa) {
            # Write-Warning "Line: $line"
            $words =  $line.TrimStart() -split '\s+'
            $thisIP = $words[0].Trim()
            if ($thisIP -eq $Global:ResultObject.IPaddress) {
                $thismac = $words[1] 
                # Write-Warning "ThisMac: $thisMac"
                if (!(($thisMac -eq "---") -or ($thisMac -eq "Address") -or ($thisMac -eq $null) -or ($thisMac -eq "ff-ff-ff-ff-ff-ff") -or ($thisMac -eq "static"))) {
                    $Global:ResultObject.MACaddress = $thisMac.Replace("-",":")
                    $Global:ResultObject.IPcached = $true
                    break
                }
               
            }
        }
        if ($Global:ResultObject.ItsMe) {
            if ($Global:ResultObject.IPcached) {        # if found in cache (not likely) then check correctness of MAC
                If ($thismac -ne $mymac) {
                    $scripterror = $true
                    $scripterrormsg = "For IP $myip ARP -A reports MACaddress $thismac while it should be $mymac"
                    AddMessage "E" "$scripterrormsg"
                   
                }
            }
           $Global:ResultObject.MACaddress = $mymac
        }
    }
    catch {
        $scripterror = $true
        $errortext = $error[0]
        
        $scripterrormsg = "Getting ARP -A info failed - $errortext"
        if ($log) {
           AddMessage "E" "$scripterrormsg"
        }
    }
    
}



if ($log) {
    $thisdate = Get-Date
    AddMessage "I" "==> END $thisdate"
}

$Global:ResultObject.Status = "InActive"
$Global:ResultObject.StatusCode = 0
If ($scripterror) {
    $Global:ResultObject.Status = "Error"
    $Global:ResultObject.StatusCode = 12
}
else {
    if ($Global:ResultObject.IPcached) {
        $Global:ResultObject.Status = "Cached"
        $Global:ResultObject.StatusCode = 3
    }  
    if ($Global:ResultObject.IPpingable) {
        $Global:ResultObject.Status = "Pingable"
        $Global:ResultObject.StatusCode = 6
    }
    if (($Global:ResultObject.IPpingable) -and ($Global:ResultObject.IPcached)) {
        $Global:ResultObject.Status = "Cached, Pingable"
        $Global:ResultObject.StatusCode = 9
    }
}      

return $Global:ResultObject



