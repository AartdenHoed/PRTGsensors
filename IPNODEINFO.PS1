param (
     [string]$InputIP = "iets"  ,
     [string]$LOGGING = "YES"
)
#$inputip = "192.168.178.201"

$NumericAddress = $inputip -match "[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}"

$MyHost = $inputip.ToUpper()

$ScriptVersion = " -- Version: 1.4.1"

# COMMON coding

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
$InitObj = & "$LocalInitVar" "OBJECT"

if ($Initobj.AbEnd) {
    # Write-Warning "YES"
    throw "INIT script $LocalInitVar Failed"

}
$MessageList = @()
$ResultObject = [PSCustomObject] [ordered] @{MessageList = $Messagelist;
                                            IPinput = $InputIP;
                                            IPaddress = "n/a";
                                            MACaddress = "n/a";
                                            IPcached = $false;
                                            IPpingable = $false;
                                            UpDown = "Unknown";
                                            ItsMe = $false;
                                            Status = "Unknown";
                                            StatusCode = -1}
function AddMessage ([string]$level, [string]$msg, [object]$obj) {
    $msgentry = [PSCustomObject] [ordered] @{Level = $level;
                                             Message = $msg}
    $Obj.MessageList += $msgentry
    # Write-Host $msg
    
    return  
}


if ($LOGGING -eq "YES") {$log = $true} else {$log = $false}


if ($log) {
    $Scriptmsg = "Directory " + $mypath + " -- PowerShell script " + $MyName + $ScriptVersion + $Datum + $Tijd +$Node
    AddMessage "I" $Scriptmsg $ResultObject     

    foreach ($entry in $InitObj.MessageList){
        $lvl = $entry.Level
        $msg = $entry.Message
        AddMessage $lvl $msg $ResultObject
    }                                
    $thisdate = Get-Date
    AddMessage "I" "==> START $thisdate" $ResultObject
    AddMessage "I" "Host = $inputIP, Logging = $LOGGING" $ResultObject
}

# END OF COMMON CODING

# Determine your own attributes (try max 5 times)

for ($i=1; $i -le 5; $i++) {

    try {
        $scripterror = $false
        $ResultObject.itsme = $false
        if ($log) {
            AddMessage "I" "Get IP/Mac address from myself (attempt nr. $i)" $ResultObject
        }
        $ComputerName = $env:computername
        $OrgSettings = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $ComputerName -EA Stop |  ? { $_.DNSDomain -eq "fritz.box" }
        $myip = $OrgSettings.IPAddress[0]

        $a = Get-NetAdapter | ? {$_.Name -like "Wi-Fi*"}
        $mymac = $a.MacAddress.Replace("-",":")
        
        # leave loop if succesfull
        break

    }
    catch {    
        $scripterror = $true
        $errortext = $error[0]
        $scripterrormsg = "Getting IP/MAC address failed - $errortext"
        if ($log) {
           AddMessage "E" "$scripterrormsg" $ResultObject
        }
        # wait before retry
        Start-Sleep -s 3
    
    }

}

$Updown = "UnKnown"

# PING first, ARP -A after that

if (!$scripterror) {
    try {
        if ($log) {
            AddMessage "I" "Perform PING" $ResultObject
        }
        try {
            $ipnum = $null
            if (!$NumericAddress) {
                $ipx = [System.Net.Dns]::GetHostAddresses($InputIp.Trim())
                if ($ipx) {
                    # write-host "ipx"
                    $Ipnum = $ipx.IPAddressToSTring
                    if ($ipnum.count) {
                        $ResultObject.IPaddress = $ipnum[0] 
                    }
                    else {
                        $ResultObject.IPaddress = $ipnum
                    }                   
                }
                else {
                    # write-host "no ipx"
                    $ResultObject.IPaddress = "Not Found"  
                }
            }
            else {
                $ResultObject.IPaddress = $inputip
            }
            if ($myip -eq $ResultObject.IPaddress) {
                $ResultObject.itsme = $true
            }
            $ping = Test-Connection -COmputerName $ResultObject.IPaddress -Count 1
            $ResultObject.IPpingable = $true
            $Updown = "Up"
            
        }
        catch {
            $ResultObject.IPpingable = $false
            $ResultObject.IPaddress = $inputip
        }
    }
    catch {
        $scripterror = $true
        $errortext = $error[0]
        
        $scripterrormsg = "Ping failed - $errortext"
        if ($log) {
           AddMessage "E" "$scripterrormsg" $ResultObject
        }

    }

    
}

# Get ARP table

if (!$scripterror) {

    try { 
        if ($log) {
            AddMessage "I" "Get ARP -A info" $ResultObject
        }
        $ResultObject.IPcached = $false
        $arpa = (arp -a) 
        foreach ($line in $arpa) {
            # Write-Warning "Line: $line"
            $words =  $line.TrimStart() -split '\s+'
            $thisIP = $words[0].Trim()
            if ($thisIP -eq $ResultObject.IPaddress) {
                $thismac = $words[1] 
                # Write-Warning "ThisMac: $thisMac"
                if (!(($thisMac -eq "---") -or ($thisMac -eq "Address") -or ($thisMac -eq $null) -or ($thisMac -eq "ff-ff-ff-ff-ff-ff") -or ($thisMac -eq "static"))) {
                    $ResultObject.MACaddress = $thisMac.Replace("-",":")
                    $ResultObject.IPcached = $true
                    break
                }
               
            }
        }
        if ($ResultObject.ItsMe) {
            if ($ResultObject.IPcached) {        # if found in cache (not likely) then check correctness of MAC
                If ($thismac -ne $mymac) {
                    $scripterror = $true
                    $scripterrormsg = "For IP $myip ARP -A reports MACaddress $thismac while it should be $mymac"
                    AddMessage "E" "$scripterrormsg" $ResultObject
                   
                }
            }
           $ResultObject.MACaddress = $mymac
        }
    }
    catch {
        $scripterror = $true
        $errortext = $error[0]
        
        $scripterrormsg = "Getting ARP -A info failed - $errortext"
        if ($log) {
           AddMessage "E" "$scripterrormsg" $ResultObject
        }
    }
    
}

# Determine Stopped/Started

if (!$scripterror) {
    try {
        
        if (!$NumericAddress) {
            # Host with name => get bootfile
            
            $str = $ADHC_BootTime.Split("\")
            $dir = $ADHC_OutputDirectory + $str[0]
            $bootfile = $ADHC_OutputDirectory + $ADHC_BootTime.Replace($ADHC_Computer, $myHost)
            $lt = Test-Path $bootfile
            if (!$lt) {
               $UpDown = "File not Found"
            }
        
            # Read bootfile
            $bootrec = Get-Content $bootfile
            if (!$bootrec) {
                $UpDown = "Record Not Found" 
            }
            else {
                $bootsplit = $bootrec.Split("|")
                $starttime = [datetime]::ParseExact($bootsplit[1],"dd-MM-yyyy HH:mm:ss",$null)
                $stoptime = [datetime]::ParseExact($bootsplit[2],"dd-MM-yyyy HH:mm:ss",$null)

                $now = Get-Date
                            
                $diff = NEW-TIMESPAN –Start $stoptime –End $now

                if ($diff.TotalMinutes -gt 10) {
                    $UpDown = "Down"    
                } 
                else {
                    if ($diff.TotalMinutes -gt 6) {
                        $UpDown = "Stopping?"
                    }
                    else {
                        $Updown = "Up"
                    }

                }
            
                if ($log) {
                    $bt = $starttime.ToString()
                    $st = $stoptime.ToString()
                    AddMessage "I" "==> Last boottime = $bt - Last endtime = $st - Node = $Updown" $ResultObject
                }
            }
        }
        else {
            # Numeric IP Address
            AddMessage "I" "==> Host name is IP address $inputip - Up/Down = $Updown" $ResultObject
            
        }
        $ResultObject.UpDown = $Updown
       
    }
    catch {
        if ($log) {
            AddMessage "E" "==> Getting boottime failed for $myHost" $ResultObject
          
        }
        $scripterror = $true
        $errortext = $error[0]
        $scripterrormsg = "==> Getting boottime failed for $myHost - $errortext"
    }
}


if ($log) {
    $thisdate = Get-Date
    AddMessage "I" "==> END $thisdate" $ResultObject
}

$ResultObject.Status = "InActive"
$ResultObject.StatusCode = 0
If ($scripterror) {
    $ResultObject.Status = "Error"
    $ResultObject.StatusCode = 12
}
else {
    if (($ResultObject.IPcached) -and (!$ResultObject.IPpingable)) {
        $ResultObject.Status = "Cached, not Pingable"
        $ResultObject.StatusCode = 3
    }  
    if (($ResultObject.IPpingable) -and (!$ResultObject.IPcached)) {
        $ResultObject.Status = "Not Cached, but Pingable"
        $ResultObject.StatusCode = 6
    }
    if (($ResultObject.IPpingable) -and ($ResultObject.IPcached)) {
        $ResultObject.Status = "Cached and Pingable"
        $ResultObject.StatusCode = 9
    }
}

#$ResultObject.MessageList  
#$ResultObject 

return $ResultObject



