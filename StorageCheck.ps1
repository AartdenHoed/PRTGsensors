param (
    [string]$LOGGING = "NO", 
    [string]$myHost  = "????" ,
    [int]$sensorid = 77 
)
# $LOGGING = 'YES'
# $myHost = "holiday"

$myhost = $myhost.ToUpper()
.1
$ScriptVersion = " -- Version: 3.4" 

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
$InitObj = & "$LocalInitVar" "OBJECT"

if ($Initobj.AbEnd) {
    # Write-Warning "YES"
    throw "INIT script $LocalInitVar Failed"

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

    foreach ($entry in $InitObj.MessageList){
        $lvl = $entry.Level
        $msg = $entry.Message
        Add-COntent $logfile "($lvl) - $msg"
    }

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

$duration = 0
if (!$scripterror) {
    try {
        # Storage info of machine
        if ($log) {
            Add-Content $logfile "==> Get storage info from machine $myHost"
        }
        $invokable = $true
        if ($myHost -eq $ADHC_Computer.ToUpper()) {
            $begin = Get-Date
                    
            $DriveInfo = get-WmiObject win32_logicaldisk | Where-Object {($_.DriveType -eq "3") }   # Only fixed disks

            $end = Get-Date
            $duration = ($end - $begin).seconds

        }
        else {
            try {
                $b = Get-Date
                $myjob = Invoke-Command -ComputerName $myhost `
                    -ScriptBlock { get-WmiObject win32_logicaldisk | Where-Object {($_.DriveType -eq "3")} } -Credential $ADHC_Credentials `
                    -JobName StorageJob  -AsJob
                
                $myjob | Wait-Job -Timeout 150 | Out-Null
                $e = Get-Date
                if ($myjob) { 
                    $mystate = $myjob.state
                    $begin = $myjob.PSBeginTime
                    $end = $myjob.PSEndTime
                    $duration = ($end - $begin).seconds
                    if ($duration -lt 0 ) {
                        $duration = ($e - $b).seconds
                    }
                } 
                else {
                    $mystate = "Unknown"
                    $duration = ($e - $b).seconds
                }
                if ($log) {
                    $mj = $myjob.Name
                    Add-Content $logfile "==> Remote job $mj ended with status $mystate"
                }
                                
                # Write-host $mystate
                if ($mystate -eq "Completed") {
                    #write-host "YES"
                    $DriveInfo = (Receive-Job -Name StorageJob)
                }
                else {
                    #write-host "NO"
                    $invokable = $false
                }
                
                $myjob | Stop-Job | Out-Null
                $myjob | Remove-Job | Out-null
            }
            catch {
                $invokable = $false
            }
            finally {
                # Write-Host $nodeisup
            }
        }
    }
    catch {
        if ($log) {
            Add-Content $logfile "==> Getting storage info failed for $myHost"
          
        }
        $scripterror = $true
        $errortext = $error[0]
        $scripterrormsg = "==> Getting storage info failed for $myHost - $errortext"
    }
}
if (!$scripterror) {
    try {
        # 
        if ($log) {
            Add-Content $logfile "==> Process storage info"
        }
        # Init storgae info file if not existent
        $str = $ADHC_DriveInfo.Split("\")
        $dir = $ADHC_OutputDirectory + $str[0] + "\" + $str[1]
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
        $DriveInfoFile = $ADHC_OutputDirectory + $ADHC_DriveInfo.Replace($ADHC_Computer, $myHost)
        $dr = Test-Path $DriveInfoFile
        if (!$dr) {
            Set-Content $DriveInfoFile "$MyHost|INIT" -force
        }
        $drivelist = @()
        if (!$invokable) {
            # Node not invokable, get info from file
            if ($log) {
                Add-Content $logfile "==> Node is down, get info from dataset"
            }
            $drivelines = Get-Content $DriveInfoFile
           
            
            foreach ($entry in $drivelines) {
                $split = $entry.Split("|")
                $letter = $split[1]
                if ($letter -eq "INIT") { break }
                $name = $split[2]
                $totalGB = [float]$split[3]
                $freeGB = [float]$split[4]
                
                $timestamp = [datetime]::ParseExact($split[5].Trim(),"dd-MM-yyyy HH:mm:ss",$null)
                $obj = [PSCustomObject] [ordered] @{Machine = $myhost; 
                                                    Letter = $letter;
                                                    Label = $name;
                                                    TotalGB = $totalGB;
                                                    FreeGB = $freeGB;
                                                    Timestamp = $timestamp
                                                    OK = 0 }
                $drivelist += $obj
                
            }
        }
        else {
            # Node is UP, take real time info and write it tot dataset
            if ($log) {
                Add-Content $logfile "==> Node is up, get realtime info and write is to dataset"
            }
            $firstrec = $true
            foreach ($drive in $DriveInfo){ 
                if ($drive.VolumeName) {
                    $name = $drive.VolumeName
                }
                else {
                    $name = "n/a"
                }
                $totalbytes = $drive.Size
                $freebytes = $drive.FreeSpace
                $factor = 1024*1024*1024
                $freeGB = [math]::Round($freebytes / $factor ,2)
                $totalGB = [math]::Round($totalbytes / $factor ,2)
                $timestamp = Get-Date
                
                $obj = [PSCustomObject] [ordered] @{Machine = $myhost; 
                                                    Letter = $drive.DeviceID;
                                                    Label = $name;
                                                    TotalGB = $totalGB;
                                                    FreeGB = $freeGB;
                                                    Timestamp = $timestamp;
                                                    OK = 0}
                $drivelist += $obj
                $record = $myhost + "|" + $drive.DeviceID + "|" + $name + "|" + $totalGB + "|" + $freeGB + "|" + $timestamp.ToString("dd-MM-yyyy HH:mm:ss")
                if ($firstrec) {
                    Set-Content $DriveInfofile $record
                    # Write-host "First"
                    $firstrec = $false
                }
                else { 
                    Add-Content $DriveInfofile $record
                }

            }
        }

    }   
    catch {
        if ($log) {
            Add-Content $logfile "==> Processing storage info failed for $myHost"
          
        }
        $scripterror = $true
        $errortext = $error[0]
        $scripterrormsg = "==> Processing storage info failed for $myHost - $errortext"
    }
}

if (!$scripterror) {
    try {
        # 
        if ($log) {
            Add-Content $logfile "==> Evaluate storage info"
        }
        $nrofdrives = 0
        $nrofwarning = 0
        $nrofcritical = 0
        # Evaluate info
        $overallscore = 2
        foreach ($result in $drivelist) {
            $nrofdrives += 1
            $freeabs = $Result.FreeGB
            $freerel = [math]::Round( $Result.FreeGB * 100 / $result.TotalGB, 2)
            if (($freeabs -gt 3) -and ($freerel -gt 5)) { $result.ok = 1} 
            if (($freeabs -gt 5) -and ($freerel -gt 10)) { $result.ok = 2} 
            $Overallscore = [math]::Min($Overallscore, $result.ok)
            if ($result.ok -eq 0) { $nrofcritical += 1 }
            if ($result.ok -eq 1) { $nrofwarning  += 1 } 
            $mytime = $result.TimeStamp.ToString("dd-MM-yyyy HH:mm:ss")
        } 
        
    }
catch {
        if ($log) {
            Add-Content $logfile "==> Evaluating storage info failed for $myHost"
          
        }
        $scripterror = $true
        $errortext = $error[0]
        $scripterrormsg = "==> Evaluating storage info failed for $myHost - $errortext"
    }
}


#$drivelist | Out-GridView
#exit

if ($log) {
    Add-Content $logfile "==> Create XML"
}

[xml]$xmldoc = New-Object System.Xml.XmlDocument
$decl = $xmldoc.CreateXmlDeclaration('1.0','Windows-1252',$null)

[void]$xmldoc.AppendChild($decl)

$PRTG = $xmldoc.CreateElement('PRTG')

# Overall storage status (Primary Channel)
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
$NotifyChanged = $xmldoc.CreateElement('NotifyChanged')
$ValueLookup =  $xmldoc.CreateElement('ValueLookup')

$Channel.InnerText = "Overall storage status"
$Unit.InnerText = "Custom"
$Mode.Innertext = "Absolute"
$ValueLookup.Innertext = 'OverallDriveStatus'

if ($scripterror) {
    $Value.Innertext = "12"
} 
else { 
   $Value.Innertext = "$Overallscore"
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
    $online = "realtime info"
} 
else { 
   $Value.Innertext = $nstat.StatusCode
   $livestat = $nstat.Status + ", Not Invokable"
   $online = "offline info"
}

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
[void]$Result.AppendChild($NotifyChanged)
[void]$Result.AppendChild($ValueLookup)
[void]$Result.AppendChild($Mode)
    
[void]$PRTG.AppendChild($Result)

# Wait time
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')    
$Unit = $xmldoc.CreateElement('Unit')
$Mode = $xmldoc.CreateElement('Mode')
$NotifyChanged = $xmldoc.CreateElement('NotifyChanged')
    
$Channel.InnerText = "Remote wait time (sec)"
$Unit.InnerText = "TimeSeconds"
$Mode.Innertext = "Absolute"
$Value.Innertext = $duration

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($NotifyChanged)
[void]$Result.AppendChild($Mode)
    
[void]$PRTG.AppendChild($Result)


foreach ($item in $Drivelist) {
   
    # Report FREE GB Absolute
    $Result = $xmldoc.CreateElement('Result')
    $Channel = $xmldoc.CreateElement('Channel')
    $Value = $xmldoc.CreateElement('Value')
    $Unit = $xmldoc.CreateElement('Unit')
    $CustomUnit = $xmldoc.CreateElement('CustomUnit')
    $Mode = $xmldoc.CreateElement('Mode')
    $Float = $xmldoc.CreateElement('Float')
    $NotifyChanged = $xmldoc.CreateElement('NotifyChanged')
    # $ValueLookup =  $xmldoc.CreateElement('ValueLookup')

    $cname = $item.Machine + " " + $item.Letter + " Free (GB)" 
    $Channel.InnerText = $cname
    $Unit.InnerText = "Custom"
    $CustomUnit.InnerText = "GB"
    $Mode.Innertext = "Absolute"
    $Float.Innertext = "1"
    # $ValueLookup.Innertext = 'xxx'
    $Value.Innertext = $item.FreeGB

    [void]$Result.AppendChild($Channel)
    [void]$Result.AppendChild($Value)
    [void]$Result.AppendChild($Unit)
    [void]$Result.AppendChild($CustomUnit)
    [void]$Result.AppendChild($NotifyChanged)
    [void]$Result.AppendChild($Float)
    [void]$Result.AppendChild($Mode)
    
    [void]$PRTG.AppendChild($Result)

    # Report FREE GB Relative
    $Result = $xmldoc.CreateElement('Result')
    $Channel = $xmldoc.CreateElement('Channel')
    $Value = $xmldoc.CreateElement('Value')
    $Unit = $xmldoc.CreateElement('Unit')
    $Mode = $xmldoc.CreateElement('Mode')
    $Float = $xmldoc.CreateElement('Float')
    $NotifyChanged = $xmldoc.CreateElement('NotifyChanged')
    # $ValueLookup =  $xmldoc.CreateElement('ValueLookup')

    $cname = $item.Machine + " " + $item.Letter + " Free (%)" 
    $Channel.InnerText = $cname
    $Unit.InnerText = "Percent"
    $Mode.Innertext = "Absolute"
    $Float.Innertext = "1"
    # $ValueLookup.Innertext = 'xxx'
    $freerel = [math]::Round( $item.FreeGB * 100 / $item.TotalGB, 2)
    $Value.Innertext = $freerel

    [void]$Result.AppendChild($Channel)
    [void]$Result.AppendChild($Value)
    [void]$Result.AppendChild($Unit)
    
    [void]$Result.AppendChild($NotifyChanged)
    [void]$Result.AppendChild($Float)
    [void]$Result.AppendChild($Mode)
    
    [void]$PRTG.AppendChild($Result)

    # Report status
    $Result = $xmldoc.CreateElement('Result')
    $Channel = $xmldoc.CreateElement('Channel')
    $Value = $xmldoc.CreateElement('Value')
    $Unit = $xmldoc.CreateElement('Unit')
    $CustomUnit = $xmldoc.CreateElement('CustomUnit')
    $Mode = $xmldoc.CreateElement('Mode')
    $NotifyChanged = $xmldoc.CreateElement('NotifyChanged')
    $ValueLookup =  $xmldoc.CreateElement('ValueLookup')

    $cname = $item.Machine + " " + $item.Letter + " (" + $item.Label.Trim() + ")"
    $Channel.InnerText = $cname
    $Unit.InnerText = "Custom"
    $Mode.Innertext = "Absolute"
    # $Float.Innertext = "1"
    $ValueLookup.Innertext = 'IndividualDriveStatus'
    $Value.Innertext = $item.OK

    [void]$Result.AppendChild($Channel)
    [void]$Result.AppendChild($Value)
    [void]$Result.AppendChild($Unit)
    [void]$Result.AppendChild($CustomUnit)
    [void]$Result.AppendChild($NotifyChanged)
    [void]$Result.AppendChild($ValueLookup)
    [void]$Result.AppendChild($Mode)
    
    [void]$PRTG.AppendChild($Result)
    

}

# Add error block

$ErrorValue = $xmldoc.CreateElement('Error')
$ErrorText = $xmldoc.CreateElement('Text')

if ($scripterror) {
    $Errorvalue.InnerText = "1"
    $ErrorText.InnerText = $scripterrormsg + " *** Scriptversion=$scriptversion *** "
}
else {
    $ErrorValue.InnerText = "0"
    
    $message = "Machine $myhost (now $livestat) *** Drives: $nrofdrives *** warning: $nrofwarning *** Critical: $nrofcritical *** Timestamp: $mytime ($online) *** Script $scriptversion"
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


