param (
    [string]$LOGGING = "NO", 
    [string]$myHost  = "????" ,
    [int]$sensorid = 77 
)
# $LOGGING = 'YES'
# $myHost = "hoesto"

$myhost = $myhost.ToUpper()

$ScriptVersion = " -- Version: 1.0"

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
& "$LocalInitVar"

if (!$ADHC_InitSuccessfull) {
    # Write-Warning "YES"
    throw $ADHC_InitError
}

if ($LOGGING -eq "YES") {$log = $true} else {$log = $false}

if ($log) {
    $dir = $ADHC_OutputDirectory + $ADHC_PRTGlogs
    # Write-Host $dir
    if (!(Test-Path $dir)) {
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
        # write-Host "Not"
    }
    if (!($sensorid -match "\d+")) {
        $sensorid = 99999
    }
    $uniqueid = $sensorid.ToString("00000")
    $logfile = $dir + $process + $uniqueid + ".log" 

    $Scriptmsg = "Directory " + $mypath + " -- PowerShell script " + $MyName + $ScriptVersion + $Datum + $Tijd +$Node
    Set-Content $logfile $Scriptmsg 

    $thisdate = Get-Date
    Add-Content $logfile "==> START $thisdate"
}

$scripterror = $false

function WriteXmlToScreen ([xml]$xml)
{
    # Function to write XML to log for PRTG
    $StringWriter = New-Object System.IO.StringWriter;
    $XmlWriter = New-Object System.Xml.XmlTextWriter $StringWriter;
    $XmlWriter.Formatting = "indented";
    $xmlWriter.QuoteChar = '"'
    $xml.WriteTo($XmlWriter);
    $XmlWriter.Flush();
    $StringWriter.Flush();
    Write-Host $StringWriter.ToString();
}

# END OF COMMON CODING
# Get Node status

if (!$scripterror) {
    try {
        if ($log) {
            Add-Content $logfile "==> Get NODE status for $myhost"
        }
        $nstat = & $ADHC_NodeInfoScript "$myHost"  "$LOGGING" 

    }
    Catch {
        if ($log) {
            Add-Content $logfile "==> Getting NODE status failed for $myhost"
        }
        $scripterror = $true
        $errortext = $error[0]
        $scripterrormsg = "Getting NODE status failed for $myhost - $errortext"

    }
    finally {
        if ($log) {
            foreach ($m in $nstat.MessageList) {
                $lvl = $m.Level
                $msg = $m.Message
                Add-COntent $logfile "($lvl) - $msg"

            }
        }
    }

}

$VpnScript = $mypath + "VpnCheck2.PS1"

if (!$scripterror) {
    try {
        # VPN info of machine
        if ($log) {
            Add-Content $logfile "==> Get VPN info from machine $myHost"
        }
        $invokable = $true
        if ($myHost -eq $ADHC_Computer.ToUpper()) {
            $VpnInfo = & $VpnScript 
        }
        else {
            try {
                
                $myjob = Invoke-Command -ComputerName $myhost `
                    -FilePath $vpnscript  -Credential $ADHC_Credentials `
                    -JobName VPNJob  -AsJob
                # write-host "Wait"
                $myjob | Wait-Job -Timeout 90 | Out-Null
                if ($myjob) { 
                    $mystate = $myjob.state
                } 
                else {
                    $mystate = "Unknown"
                }
                if ($log) {
                    $mj = $myjob.Name
                    Add-Content $logfile "==> Remote job $mj ended with status $mystate"
                }
                                
                # Write-host $mystate
                if ($mystate -eq "Completed") {
                    # write-host "YES"
                    $VpnInfo = (Receive-Job -Name VPNJob)
                }
                else {
                    #write-host "NO"
                    $invokable = $false
                }
                # write-host "Stop"
                $myjob | Stop-Job | Out-Null
                # write-host "Remove"
                $myjob | Remove-Job | Out-null
            }
            catch {
                # write-host "Catch"
                $invokable = $false
            }
            finally {
                # Write-Host $nodeisup
            }
        }
    }
    catch {
        if ($log) {
            Add-Content $logfile "==> Getting VPN info failed for $myHost"
          
        }
        $scripterror = $true
        $errortext = $error[0]
        $scripterrormsg = "==> Getting VPN info failed for $myHost - $errortext"
    }
}
if (!$scripterror) {
    try {
        # 
        if ($log) {
            Add-Content $logfile "==> Process VPN info"
        }
        # Init VPN info file if not existent
        $str = $ADHC_VpnInfo.Split("\")
        $dir = $ADHC_OutputDirectory + $str[0] + "\" + $str[1]
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
        $VpnInfoFile = $ADHC_OutputDirectory + $ADHC_VpnInfo.Replace($ADHC_Computer, $myHost)
        $dr = Test-Path $VpnInfoFile
        if (!$dr) {
            Set-Content $VpnInfoFile "$MyHost|INIT" -force
        }
        
        if (!$invokable) {
            # Node not invokable, get info from file
            if ($log) {
                Add-Content $logfile "==> Node is down, get info from dataset"
            }
            $Vpnline = Get-Content $VpnInfoFile
            $split = $vpnline.Split("|")       
            $letter = $split[1]
            
            if ($letter -ne "INIT") { 
                $VPNip = $split[2]
                $VPNinfo = $split[3]
                $VPNstatus = $split[4]                              
                $timestamp = [datetime]::ParseExact($split[5].Trim(),"dd-MM-yyyy HH:mm:ss",$null)
                $obj = [PSCustomObject] [ordered] @{Machine = $myhost; 
                                                    Letter = $letter;
                                                    VPNip = $VPNip;
                                                    VPNinfo = $VPNinfo;
                                                    Timestamp = $timestamp;
                                                    Status = $VPNstatus }
            }
            else {
                # Status INIT, and no realtime info available

                $obj = [PSCustomObject] [ordered] @{Machine = $myhost; 
                                                    Letter = $letter;
                                                    VPNip = "Unknown";
                                                    VPNinfo = "No realtime info available, and status is INIT";
                                                    Timestamp = Get-Date;
                                                    Status = 1 }

            }
        }

        else {
            # Node is UP, take real time info and write it tot dataset
            if ($log) {
                Add-Content $logfile "==> Node is up, get realtime info and write is to dataset"
            }         
            $timestamp = Get-Date
                
            $obj = [PSCustomObject] [ordered] @{Machine = $myhost; 
                                                    Letter = "NOINIT";
                                                    VPNip = $VPNinfo.IpAddress;
                                                    VPNinfo = $VPNinfo.Info;
                                                    Timestamp = $timestamp;
                                                    Status = $VPNinfo.Status }
                
            $record = $myhost + "|" + $obj.Letter + "|" + $obj.VPNip + "|" + $obj.VPNinfo + "|" + $obj.Status + "|" + $timestamp.ToString("dd-MM-yyyy HH:mm:ss")
            Set-Content $VpnInfofile $record
                        
        }

    }   
    catch {
        if ($log) {
            Add-Content $logfile "==> Processing VPN info failed for $myHost"
          
        }
        $scripterror = $true
        $errortext = $error[0]
        $scripterrormsg = "==> Processing VPN info failed for $myHost - $errortext"
    }
}

if ($log) {
    Add-Content $logfile "==> Create XML"
}

[xml]$xmldoc = New-Object System.Xml.XmlDocument
$decl = $xmldoc.CreateXmlDeclaration('1.0','Windows-1252',$null)

[void]$xmldoc.AppendChild($decl)

$PRTG = $xmldoc.CreateElement('PRTG')

# Overall VPN status
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
$NotifyChanged = $xmldoc.CreateElement('NotifyChanged')
$ValueLookup =  $xmldoc.CreateElement('ValueLookup')

$Channel.InnerText = "VPN status"
$Unit.InnerText = "Custom"
$Mode.Innertext = "Absolute"
$ValueLookup.Innertext = 'VpnStatus'

if ($scripterror) {
    $Value.Innertext = "12"
} 
else { 
   $Value.Innertext = $obj.Status
}

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
[void]$Result.AppendChild($NotifyChanged)
[void]$Result.AppendChild($ValueLookup)
[void]$Result.AppendChild($Mode)
    
[void]$PRTG.AppendChild($Result)

# Node status
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
$NotifyChanged = $xmldoc.CreateElement('NotifyChanged')
$ValueLookup =  $xmldoc.CreateElement('ValueLookup')

$Channel.InnerText = "Node status"
$Unit.InnerText = "Custom"
$Mode.Innertext = "Absolute"
$ValueLookup.Innertext = 'NodeStatus'

if ($invokable) {
    $Value.Innertext = $nstat.StatusCode + 1
    $livestat = $nstat.Status + ", Invokable"
} 
else { 
   $Value.Innertext = $nstat.StatusCode
   $livestat = $nstat.Status + ", Not Invokable"
}

#if ($nodeisup) {
#    $Value.Innertext = "0"
#    $livestat = "UP"
#} 
#else { 
#   $Value.Innertext = "1"
#   $livestat = "DOWN"
#}

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
[void]$Result.AppendChild($NotifyChanged)
[void]$Result.AppendChild($ValueLookup)
[void]$Result.AppendChild($Mode)
    
[void]$PRTG.AppendChild($Result)
  

# Add error block

$ErrorValue = $xmldoc.CreateElement('Error')
$ErrorText = $xmldoc.CreateElement('Text')

if ($scripterror) {
    $Errorvalue.InnerText = "1"
    $ErrorText.InnerText = $scripterrormsg + " *** Scriptversion=$scriptversion *** "
}
else {
    $ErrorValue.InnerText = "0"
    $i = $obj.VPNip
    $m = $obj.VPNinfo
    $t = $obj.Timestamp
    
    $message = "IP Address =  $i *** Info =  $m *** Timestamp: $t *** Script $scriptversion"
    $ErrorText.InnerText = $message
} 
[void]$PRTG.AppendChild($ErrorValue)
[void]$PRTG.AppendChild($ErrorText)
    
[void]$xmldoc.Appendchild($PRTG)

if ($log) {
    Add-Content $logfile "==> Write XML"
}

WriteXmlToScreen $xmldoc

if ($log) {
    $thisdate = Get-Date
    Add-Content $logfile "==> END $thisdate"
}

# <SPAN class=green>85.203.44.142</SPAN> 
# <P>Uw IP adres is veilig. Websites kunnen deze niet gebruiken om achter uw identiteit te komen. </P>

# <SPAN class=red>45.148.141.15</SPAN> 
# <P>Uw IP adres is momenteel zichtbaar. Begin uw online anonimiteit terug te eisen met een VPN. </P>


