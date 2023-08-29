param (
     [string]$InputIP = "iets"  ,
     [string]$LOGGING = "YES"
)
#$inputip = "192.168.178.201"

$NumericAddress = $inputip -match "[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}"

$MyHost = $inputip.ToUpper()

$ScriptVersion = " -- Version: 1.4"

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
                                            UpDown = "Unknown";
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

# Determine your own attributes (try max 5 times)

for ($i=1; $i -le 5; $i++) {

    try {
        $scripterror = $false
        $Global:ResultObject.itsme = $false
        if ($log) {
            AddMessage "I" "Get IP/Mac address from myself (attempt nr. $i)"
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
           AddMessage "E" "$scripterrormsg"
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
            AddMessage "I" "Perform PING"
        }
        try {
            $ipnum = $null
            if (!$NumericAddress) {
                $ipx = [System.Net.Dns]::GetHostAddresses($InputIp.Trim())
                if ($ipx) {
                    # write-host "ipx"
                    $Ipnum = $ipx.IPAddressToSTring
                    if ($ipnum.count) {
                        $Global:ResultObject.IPaddress = $ipnum[0] 
                    }
                    else {
                        $Global:ResultObject.IPaddress = $ipnum
                    }                   
                }
                else {
                    # write-host "no ipx"
                    $Global:ResultObject.IPaddress = "Not Found"  
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
            $Updown = "Up"
            
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
                    AddMessage "I" "==> Last boottime = $bt - Last endtime = $st - Node = $Updown"
                }
            }
        }
        else {
            # Numeric IP Address
            AddMessage "I" "==> Host name is IP address $inputip - Up/Down = $Updown" 
            
        }
        $Global:ResultObject.UpDown = $Updown
       
    }
    catch {
        if ($log) {
            AddMessage "E" "==> Getting boottime failed for $myHost"
          
        }
        $scripterror = $true
        $errortext = $error[0]
        $scripterrormsg = "==> Getting boottime failed for $myHost - $errortext"
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
    if (($Global:ResultObject.IPcached) -and (!$Global:ResultObject.IPpingable)) {
        $Global:ResultObject.Status = "Cached, not Pingable"
        $Global:ResultObject.StatusCode = 3
    }  
    if (($Global:ResultObject.IPpingable) -and (!$Global:ResultObject.IPcached)) {
        $Global:ResultObject.Status = "Not Cached, but Pingable"
        $Global:ResultObject.StatusCode = 6
    }
    if (($Global:ResultObject.IPpingable) -and ($Global:ResultObject.IPcached)) {
        $Global:ResultObject.Status = "Cached and Pingable"
        $Global:ResultObject.StatusCode = 9
    }
}

#$Global:ResultObject.MessageList  
#$Global:ResultObject 

return $Global:ResultObject



