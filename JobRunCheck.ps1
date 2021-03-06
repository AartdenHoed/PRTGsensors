﻿param (
    [string]$LOGGING = "NO", 
    [string]$myHost  = "????" ,
    [int] $sensorid = 77 
)
#$LOGGING = 'YES'
#$myHost = "adhc"

$myhost = $myhost.ToUpper()

$ScriptVersion = " -- Version: 3.1.1"

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
        # get boottime of machine
        if ($log) {
            Add-Content $logfile "==> Get boottime from machine $myHost"
        }
        # $boot = Invoke-Expression('systeminfo | find /i "Boot Time"')
        $invokable = $true
        if ($myHost -eq $ADHC_Computer.ToUpper()) {
            $begin = Get-Date
                      
            $bt = Get-CimInstance -Class Win32_OperatingSystem | Select-Object LastBootUpTime
            $boottime = $bt.LastBootUpTime

            $end = Get-Date
            $duration = ($end - $begin).seconds
        }
        else {
            try {
                $b = Get-Date
                $myjob = Invoke-Command -ComputerName $myhost `
                    -ScriptBlock { Get-CimInstance -Class Win32_OperatingSystem | Select-Object LastBootUpTime } -Credential $ADHC_Credentials `
                    -JobName BootJob  -AsJob
                # $bt = Invoke-Command -ComputerName $myhost -ScriptBlock { Get-CimInstance -Class Win32_OperatingSystem | Select-Object LastBootUpTime } -Credential $ADHC_Credentials
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
                    $boottime = (Receive-Job -Name BootJob).LastBootUpTime
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
            $bootrec =  "$MyHost|01-01-2000 00:00:00|01-01-2000 00:00:00" 
        }
        $bootsplit = $bootrec.Split("|")
        $starttime = [datetime]::ParseExact($bootsplit[1],"dd-MM-yyyy HH:mm:ss",$null)
        $stoptime = [datetime]::ParseExact($bootsplit[2],"dd-MM-yyyy HH:mm:ss",$null)
        # If node is NOT up, get last boottime from dataset, else update dataset
        # and Update stoptime if not already done so
        if (!$invokable) {           
            $boottime = $starttime            
            if ($stoptime -lt $starttime) { # update only first time after computer down
                $stoptime = Get-Date
                $bootrec = "$MyHost" + "|" + $boottime.ToString("dd-MM-yyyy HH:mm:ss") + "|" + $stoptime.ToString("dd-MM-yyyy HH:mm:ss")
                Set-Content $bootfile "$bootrec" 
            }
            $now = $stoptime
        }
        # if node is UP, update the bootfile with boottime
        else { # update file only first time after boot
            if ($boottime -gt $startime) {
                $bootrec = "$MyHost" + "|" + $boottime.ToString("dd-MM-yyyy HH:mm:ss") + "|" + $stoptime.ToString("dd-MM-yyyy HH:mm:ss")
                Set-Content $bootfile "$bootrec"
            }
            $now = Get-Date
        }          
        
        $diff = NEW-TIMESPAN –Start $boottime –End $now
        # Only check job status if computer has been up for >1,5 hour
        if ($diff.TotalMinutes -ge 90) {
            $checkruns = $true
        }
        else {
            $checkruns = $false
        }
        if ($log) {
            $bt = $boottime.ToString()
            Add-Content $logfile "==> Boottime = $bt, Node $MyHost INVOKABLE=$invokable"
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
            $a = Get-Content $logdataset.FullName
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
        # $resultlist | Out-gridview
        
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
   $Value.Innertext = $maxcode
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
    $bt = $boottime.ToString()
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


