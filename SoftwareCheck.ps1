param (
    [string]$LOGGING = "YES", 
    [string]$myHost  = "NONE",
    [int]$sensorid = 77 
)
# $LOGGING = 'YES'
# $myHost = "ADHC-2"

$myhost = $myhost.ToUpper()
$logging = $logging.ToUpper()

$ScriptVersion = " -- Version: 2.0"

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

# If LOGGING=YES create sensorlog if nog existent, and INIT it. 

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
    if ($log) {
        Add-Content $logfile $StringWriter.ToString()
    }
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

# Run WMIC to get software inventory if system appears to be invokable

if (!$scripterror) {
    $duration = 0
    try {
        # Software info of machine
        if ($log) {
            Add-Content $logfile "==> Get software info from machine $myHost"
        }
        $invokable = $true
        if ($myHost -eq $ADHC_Computer.ToUpper()) {
            $begin = Get-Date
                    
            $SoftwareInfo = WMIC  product get Name,Vendor,Version,InstallLocation,InstallDate

            $end = Get-Date
            $duration = ($end - $begin).seconds

        }
        else {
            try {
                $b = Get-Date
                $myjob = Invoke-Command -ComputerName $myhost `
                    -ScriptBlock { $a = WMIC  product get Name,Vendor,Version,InstallLocation,InstallDate ; return $a   } `
                    -Credential $ADHC_Credentials -JobName SoftwareJob  -AsJob
                
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
                    # write-host "YES"
                    $SoftwareInfo = (Receive-Job -Name SoftwareJob)
                    write-host "kom ik hier?"
                }
                else {
                    # write-host "NO"
                    $invokable = $false
                }
                
                $myjob | Stop-Job | Out-Null
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
            Add-Content $logfile "==> Getting software info from host failed for $myHost"
          
        }
        $scripterror = $true
        $errortext = $error[0]
        $scripterrormsg = "==> Getting software info from host failed for $myHost - $errortext"
    }

}

$nrofinstallationsWMIC = 0

if (!$scripterror) {
    if ($invokable) {
    # Node is UP, take real time WMIC info and write it tot dataset
        try {
            # 
             if ($log) {
                    Add-Content $logfile "==> Node is up, proces the WMIC output and write it to dataset"
                } 
            
            $first = $true
            $currentdate = Get-Date
            $prefixdate = $currentdate.ToString("yyyy-MM-dd HH:mm:ss").TrimEnd()
            $prefixcomp = $myhost.TrimEnd()
            $head1 = "Computer"
            $head2 = "Timestamp"
            $l = $head1.Length,$prefixcomp.Length | Measure-Object -maximum 
            $lenhead1 = $l.maximum + 2
            $blanks1 = 
            $l = $head2.Length,$prefixdate.Length | Measure-Object -maximum 
            $lenhead2= $l.Maximum + 2
            
            $tempfile = $ADHC_WmicTempDir + "TEMP_" + $myhost.TrimEnd() + ".txt"
                        
            foreach ($software in $SoftwareInfo){
                # Determine program name without arguments
                if ($first) {
                    $first = $false
                    $record = $head1.Padright($lenhead1, " ") + $head2.Padright($lenhead2, " ") + $software
                    Set-Content $TempFile $record -Encoding Unicode
                }
                else {
                    if ($software.TrimEnd() -eq "") {
                        continue
                    }
                    $record = $record = $prefixcomp.Padright($lenhead1, " ") + $prefixdate.Padright($lenhead2, " ") + $software
                    $nrofinstallationsWMIC = $nrofinstallationsWMIC + 1
                    Add-Content $TempFile $record -Encoding Unicode
                } 
            }
        }   
        catch {
            if ($log) {
                Add-Content $logfile "==> Processing WMIC output failed for $myHost"
          
            }
            $scripterror = $true
            $errortext = $error[0]
            $scripterrormsg = "==> Processing WMIC output failed for $myHost - $errortext"
            # exit
        }
    }
    else {
        $nrofinstallationsWMIC = -1
    }
}

# copy tempfile to definitive file if it has been created (host invokable)
if (!$scripterror) {
    if ($invokable) {
        
        $now = Get-Date
        $yyyymmdd = $now.ToString("yyyyMMdd").TrimEnd()      
        $ofile = $ADHC_WmicDirectory + "WMIC_" + $myhost.Replace("-","_") + "_" + $yyyymmdd + ".txt"

        if (Test-Path $ofile) {
            $action = 1
        }
        else {
            $action = 2
        }

        $cm1 = & $ADHC_CopyMoveScript $TempFile $ofile "MOVE" "REPLACE" "JSON" "WMIC,$process"  

        if ($log) {
            $cmlist = ConvertFrom-Json $cm1
            foreach ($m in $cmlist.MessageList) {
                $lvl = $m.Level
                $msg = $m.Message
                Add-Content $logfile "($lvl) - $msg"

            }
        }
        # copy definitive file to analyses file

        $anafile = $ADHC_WmicDirectory + "Analysis_Copy_" + $myhost.Replace("-","_") + ".txt"
        $cm2 = & $ADHC_CopyMoveScript $ofile $anafile "COPY" "REPLACE" "JSON" "WMIC,$process"  

        if ($log) {
            $cmlist = ConvertFrom-Json $cm2
            foreach ($m in $cmlist.MessageList) {
                $lvl = $m.Level
                $msg = $m.Message
                Add-Content $logfile "($lvl) - $msg"

            }
        }

    }
    else {
        $action = 0
    }
    $dsnmatch = "WMIC_" + $myhost.Replace("-","_").ToUpper() + "_" + "\d{8}" + ".txt"
    $nrofdsn = ((Get-ChildItem $ADHC_WmicDirectory -File | Select Name,FullName | Where-Object {$_.Name.ToUpper() -match $dsnmatch}) | Measure-Object).Count
    

}

# Get info from database and compare it with WMIC info

if (!$scripterror) {
                     
            
    if ($log) {
        Add-Content $logfile "==> Node is down, get info from SQL database"
    }

    # get computer ID
    $query = "Select ComputerID
                From dbo.Computer
                WHERE ComputerName = '" + $myhost + "'"  
    $DbResult = invoke-sqlcmd -ServerInstance '.\sqlexpress' -Database "Sympa" `
                    -Query "$query" `
                    -ErrorAction Stop 
    if (!$DbResult) {
        $scripterrormsg = "==> Host $myHost not found in database" 
        if ($log) {
            Add-Content $logfile $scripterrormsg          
        }
        $scripterror = $true
    }
    else {
        $computerid = $DbResult.ComputerID
    }            

    # get number of installations on that computer
    $query = "Select SUM([Count])
                From dbo.Installation
                WHERE ComputerID = " + $computerID + " AND EndDatetime is NULL"  
    $DBresult = invoke-sqlcmd -ServerInstance '.\sqlexpress' -Database "Sympa" `
                -Query "$query" `
                -ErrorAction Stop 
    $nrofinstallationsDB = $DBresult.Item(0)

    # get number of unauthorized installations on that computer
    $query = "Select count(*) From dbo.Installation I, dbo.Component C
              WHERE i.ComponentID = c.ComponentID AND ComputerID = " + $computerid + " AND EndDatetime is NULL AND c.Authorized = 'N'  "
    $DBresult = invoke-sqlcmd -ServerInstance '.\sqlexpress' -Database "Sympa" `
                -Query "$query" `
                -ErrorAction Stop 
    $nrofillegals = $DBresult.Item(0)
            
}

# if no WMIC data available, no delta can be calculated
if ($nrofinstallationsWMIC -lt 0) {
    $installationdelta = 0
    $nrofinstallationsWMIC = $nrofinstallationsDB
}
else {
    $installationdelta = $nrofinstallationsWMIC - $nrofinstallationsDB
}


if ($log) {
    Add-Content $logfile "==> Create XML"
}

[xml]$xmldoc = New-Object System.Xml.XmlDocument
$decl = $xmldoc.CreateXmlDeclaration('1.0','Windows-1252',$null)

[void]$xmldoc.AppendChild($decl)

$PRTG = $xmldoc.CreateElement('PRTG')

# Node status (PRIMARY Channel)
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

# Number of Software items before
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')    
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
    
$Channel.InnerText = "Software Count Before"
$Value.Innertext = $nrofinstallationsDB
$Unit.InnerText = "Custom"
$CustomUnit.Innertext = "Number of installations"
$Mode.Innertext = "Absolute"

$Value.Innertext = $nrofinstallationsDB

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($Customunit)
[void]$Result.AppendChild($Mode)
    
[void]$PRTG.AppendChild($Result)

# Number of Software items after
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')    
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
    
$Channel.InnerText = "Software Count Now"
$Value.Innertext = $nrofinstallationsWMIC
$Unit.InnerText = "Custom"
$CustomUnit.Innertext = "Number of installations"
$Mode.Innertext = "Absolute"

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($Customunit)
[void]$Result.AppendChild($Mode)
    
[void]$PRTG.AppendChild($Result)

# Number of Software items delta
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')    
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
$LimitMode = $xmldoc.CreateElement('LimitMode')
$LimitMinError = $xmldoc.CreateElement('LimitMinError')
$LimitMaxError = $xmldoc.CreateElement('LimitMaxError')
    
$Channel.InnerText = "Software Count Delta"
$Value.Innertext = $installationDelta
$Unit.InnerText = "Custom"
$CustomUnit.Innertext = "Number of installations"
$Mode.Innertext = "Absolute"
$LimitMode.InnerText = "1"
$LimitMinError.InnerText = "0"
$LimitMaxError.InnerText = "0"

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($Customunit)
[void]$Result.AppendChild($Mode)
[void]$Result.AppendChild($LimitMode) 
[void]$Result.AppendChild($LimitMinError) 
[void]$Result.AppendChild($LimitMaxError)     
    
[void]$PRTG.AppendChild($Result)

# Number of illegal software items
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')    
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
$LimitMode = $xmldoc.CreateElement('LimitMode')
$LimitMinError = $xmldoc.CreateElement('LimitMinError')
$LimitMaxError = $xmldoc.CreateElement('LimitMaxError')
    
$Channel.InnerText = "Illegal Installations"
$Value.Innertext = $nrofillegals
$Unit.InnerText = "Custom"
$CustomUnit.Innertext = "Number of installations"
$Mode.Innertext = "Absolute"
$LimitMode.InnerText = "1"
$LimitMinError.InnerText = "0"
$LimitMaxError.InnerText = "0"

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($Customunit)
[void]$Result.AppendChild($Mode)
[void]$Result.AppendChild($LimitMode) 
[void]$Result.AppendChild($LimitMinError) 
[void]$Result.AppendChild($LimitMaxError)     
    
[void]$PRTG.AppendChild($Result)

 # Report file action (replace or create)
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
$NotifyChanged = $xmldoc.CreateElement('NotifyChanged')
$ValueLookup =  $xmldoc.CreateElement('ValueLookup')

$Channel.InnerText = "Software File Action"
$Unit.InnerText = "Custom"
$Mode.Innertext = "Absolute"
$ValueLookup.Innertext = 'SoftwareFileAction'

$Value.Innertext = $action

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
[void]$Result.AppendChild($NotifyChanged)
[void]$Result.AppendChild($ValueLookup)
[void]$Result.AppendChild($Mode)
    
[void]$PRTG.AppendChild($Result)

# Number of software files
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')    
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
$NotifyChanged = $xmldoc.CreateElement('NotifyChanged')
$ValueLookup =  $xmldoc.CreateElement('ValueLookup')
    
$Channel.InnerText = "Number of software files"
$Unit.InnerText = "Custom"
$Mode.Innertext = "Absolute"
$ValueLookup.Innertext = 'SoftwareFiles'

$Value.Innertext = $nrofdsn

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
    $formattime = $d.ToString("dd-MM-yyyy HH:mm:ss")
    $message = "Machine $myhost (now $livestat) *** $nrofinstallationsWMIC Software instances found *** $nrofdsn datasets waiting to be processed *** Timestamp: $formattime ($online) *** Script $scriptversion"
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

