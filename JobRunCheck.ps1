param (
    [string]$LOGGING = "NO", 
    [string]$myHost  = "????" ,
    [int] $sensorid = 77 
)
# $LOGGING = 'YES'
# $myHost = "ADHC-2"

$myhost = $myhost.ToUpper()

$ScriptVersion = " -- Version: 4.0.2"

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
& "$LocalInitVar" "SILENT"

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

# END OF COMMON CODING


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


try {
    if ($log) {
        Add-Content $logfile "==> Get list of jobstatus files for computer $myHost"
    }
    $logdir = $ADHC_OutputDirectory + $ADHC_JobStatus
    $logList = Get-ChildItem $logdir -File | Select Name,FullName | Where-Object {$_.Name.ToUpper() -match $myHost}
       
}
catch {
    if ($log) {
        Add-Content $logfile "==> Reading directoy $dir failed"
    }
    $scripterror = $true
    $errortext = $error[0]
    $scripterrormsg = "Reading directoy $dir failed - $errortext"
    
}

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
        # Init boottime file if not existent
        $str = $ADHC_BootTime.Split("\")
        $dir = $ADHC_OutputDirectory + $str[0]
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
        $bootfile = $ADHC_OutputDirectory + $ADHC_BootTime.Replace($ADHC_Computer, $myHost)
        $lt = Test-Path $bootfile
        if (!$lt) {
            Set-Content $bootfile "$MyHost|01-01-2000 00:00:00|01-01-2000 00:00:00" -force
        }
        
        # Read bootfile
        $bootrec = Get-Content $bootfile
        if (!$bootrec) {
            $bootrec =  "$MyHost|01-01-2000 00:00:00|01-01-2000 00:00:00|0" 
        }
        $bootsplit = $bootrec.Split("|")
        $starttime = [datetime]::ParseExact($bootsplit[1],"dd-MM-yyyy HH:mm:ss",$null)
        $stoptime = [datetime]::ParseExact($bootsplit[2],"dd-MM-yyyy HH:mm:ss",$null)
                
        $diff = NEW-TIMESPAN –Start $starttime –End $stoptime
        # Only check job status if computer has been up for >2 hour
        if ($diff.TotalMinutes -ge 120) {
            $checkruns = $true
        }
        else {
            $checkruns = $false
        }
        if ($log) {
            $bt = $starttime.ToString()
            Add-Content $logfile "==> Last boottime = $bt"
        }
    }
    catch {
        if ($log) {
            Add-Content $logfile "==> Getting boottime failed for $myHost"
          
        }
        $scripterror = $true
        $errortext = $error[0]
        $scripterrormsg = "==> Getting boottime failed for $myHost - $errortext"
    }
}

# Process all jobstatus files

if (!$scripterror) {
    $missingfile = $false
    try { 
        if ($log) {
            Add-Content $logfile "==> Interprete each jobstatus file"
        }
        $resultlist = @()
        $total = 0
        $stat0 = 0
        $stat2 = 0
        $stat6 = 0
        $MaxCode = 0

        
        foreach ($logdataset in $loglist) {
            $readsuccess = $false
            $skippit = $false
            $trycount = 0
            do {
                try {
                    $trycount += 1
                    $a = Get-Content $logdataset.FullName
                    $readsuccess = $true
                }
                catch {
                    $f = $logdataset.FullName
                    $errortext = $error[0]
                    if ($trycount -le 5) {
                        if ($log) {                            
                            Add-Content $logfile "==> Attempt number $trycount - reading $f failed: $errortext"
                            Add-Content $logfile "==> Wait 5 seconds and retry"
                        }
                        Start-Sleep -Seconds 5
                    }
                    else {
                        if ($log) {                            
                            Add-Content $logfile "==> File $f failed skipped"
                        }
                        $skippit = $true
                        $readsuccess = $false  
                    }
                }
            } until ($readsuccess -or $skippit)

            if ($readsuccess) {
                $args = $a.Split("|")
                $machine = $args[0]
                $Job = $args[1]
                $Timestamp = [datetime]::ParseExact($args[4],"dd-MM-yyyy HH:mm:ss",$null)
                if ($timestamp -gt $boottime) {
                        $runstatus = 0 #ok
                        $stat0 +=1
                    }
                else {
                    if ($checkruns) {                
                        $runstatus = 6 #late
                        $stat6 +=1                
                    }
                    else {
                        $runstatus = 2 # boot is too recent
                        $stat2 +=1
                    }
                }
                $Maxcode = [math]::Max($Maxcode, $runstatus)
           
                $obj = [PSCustomObject] [ordered] @{Job = $Job;
                                                Machine = $machine; 
                                                Timestamp = $Timestamp;
                                                RunStatus = $runstatus}
                $resultlist += $obj 
                $Total = $Total + 1; 
                    
            }
            else {
                $missingfile = $true
            }
            # $resultlist | Out-gridview
        
        }
    }
    catch {
        if ($log) {
            Add-Content $logfile "==> Processing logfiles failed at $logdataset"
          
        }
        $scripterror = $true
        $errortext = $error[0]
        $scripterrormsg = "==> Processing logfiles failed at $logdataset - $errortext"
    
    }
}

if ($log) {
    Add-Content $logfile "==> Create XML"
}

[xml]$xmldoc = New-Object System.Xml.XmlDocument
$decl = $xmldoc.CreateXmlDeclaration('1.0','Windows-1252',$null)

[void]$xmldoc.AppendChild($decl)

$PRTG = $xmldoc.CreateElement('PRTG')

# Overall status (PRIMARY CHANNEL)

$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
$NotifyChanged = $xmldoc.CreateElement('NotifyChanged')
$ValueLookup =  $xmldoc.CreateElement('ValueLookup')

$Channel.InnerText = "Overall Job Run status"
$Unit.InnerText = "Custom"
$Mode.Innertext = "Absolute"
$ValueLookup.Innertext = 'OverallRUNStatus'

if ($scripterror) {
    $Value.Innertext = "12"
} 
else {
    if ($missingfile) {
        $Value.Innertext = "5"
    }
    else { 
        $Value.Innertext = $maxcode
    }
}

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
[void]$Result.AppendChild($NotifyChanged)
[void]$Result.AppendChild($ValueLookup)
[void]$Result.AppendChild($Mode)
    
[void]$PRTG.AppendChild($Result)

foreach ($item in $resultlist) {
   
    # Report each JOB as Channel
    $Result = $xmldoc.CreateElement('Result')
    $Channel = $xmldoc.CreateElement('Channel')
    $Value = $xmldoc.CreateElement('Value')
    $Unit = $xmldoc.CreateElement('Unit')
    $CustomUnit = $xmldoc.CreateElement('CustomUnit')
    $Mode = $xmldoc.CreateElement('Mode')
    $NotifyChanged = $xmldoc.CreateElement('NotifyChanged')
    $ValueLookup =  $xmldoc.CreateElement('ValueLookup')

    $cname = $item.Machine + "/" + $item.Job
    $Channel.InnerText = $cname
    $Unit.InnerText = "Custom"
    $Mode.Innertext = "Absolute"
    $ValueLookup.Innertext = 'IndividualRUNStatus'

    $Value.Innertext = $item.RunStatus

    [void]$Result.AppendChild($Channel)
    [void]$Result.AppendChild($Value)
    [void]$Result.AppendChild($Unit)
    [void]$Result.AppendChild($CustomUnit)
    [void]$Result.AppendChild($NotifyChanged)
    [void]$Result.AppendChild($ValueLookup)
    [void]$Result.AppendChild($Mode)
    
    [void]$PRTG.AppendChild($Result)
    

}

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

# Add error block

$ErrorValue = $xmldoc.CreateElement('Error')
$ErrorText = $xmldoc.CreateElement('Text')

if ($scripterror) {
    $Errorvalue.InnerText = "1"
    $ErrorText.InnerText = $scripterrormsg + " *** Scriptversion=$scriptversion *** "
}
else {
    $ErrorValue.InnerText = "0"
    $bt = $starttime.ToString()
    $message = "Machine $myhost (now $livestat) last booted $bt ($online) *** Total jobs: $Total *** Jobs Executed: $stat0 *** Jobs waiting to run: $Stat2 *** Jobs NOT run (error): $stat6 *** Script $scriptversion"
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


