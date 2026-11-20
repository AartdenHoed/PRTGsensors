param (
    [string]$LOGGING = "YES", 
    [string]$myHost  = "NONE",
    [int]$sensorid = 77 
)
# $LOGGING = 'YES'
# $myHost = "hoesto"

$myhost = $myhost.ToUpper()
$logging = $logging.ToUpper()

$ScriptVersion = " -- Version: 1.2.2"

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
        if ($nstat.UpDown -eq "Up") {
            $Up = $true
            $uptext = "Up"
        }
        else {
            $up = $false
            $uptext = "Down"
        }

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

# Get info from file if present


if (!$scripterror) {
    try {
        # 
        if ($log) {
            Add-Content $logfile "==> Process Restore Point Info"
        }
        # Determine name of info file (filled bij powershell script, if existent)
        
        $RestorePointInfoFile = $ADHC_OutputDirectory + $ADHC_RestorePointInfo.Replace($ADHC_Computer, $myHost)

        $dr = Test-Path $RestorePointInfoFile
        $now = Get-Date

        if (!$dr) {
             $RP_agedays = -2 # Unknown
             $RP_time = Get-Date -Year 2020 -Month 01 -Day 01 -Hour 01 -Minute 01 -Second 01 
             $RP_timetext = "~~~ UNKNOWN ~~~"
             $RP_desc = "~~~ No Description ~~~"
             $RP_eventtype = -1
        }
        else {
            # Node is up, and info dataset exists, so go get it
                        
            $RP_line = Get-Content $RestorePointInfoFile
            
            $split = $RP_line.Split("|")
            $RP_time = [datetime]::ParseExact($split[0],"dd-MM-yyyy HH:mm:ss",$null)
            $RP_timetext = $split[0]
            $RP_desc = $split[1]
            $RP_eventtype = $split[2] 
            $RP_count = $split[3]           

            $diff = NEW-TIMESPAN –Start $RP_time –End $now
            $RP_agedays = $diff.Days

            if (!$up) {
                $RP_agedays = -1
            }
        }

    }   
    catch {
        if ($log) {
            Add-Content $logfile "==> Processing input info failed for $myHost"
          
        }
        $scripterror = $true
        $errortext = $error[0]
        $scripterrormsg = "==> Processing input info failed for $myHost - $errortext"
    }
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
$Unit = $xmldoc.CreateElement('Unit')
$Mode = $xmldoc.CreateElement('Mode')
$ValueLookup =  $xmldoc.CreateElement('ValueLookup')
$Value = $xmldoc.CreateElement('Value')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$NotifyChanged = $xmldoc.CreateElement('NotifyChanged')

$Channel.InnerText = "Node Status"
$Unit.InnerText = "Custom"
$Mode.Innertext = "Absolute"
$ValueLookup.Innertext = 'UpDown'
if ($up) {
    $Value.Innertext = 0
} 
else { 
    $Value.Innertext = 1
}

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
[void]$Result.AppendChild($NotifyChanged)
[void]$Result.AppendChild($ValueLookup)
[void]$Result.AppendChild($Mode)
    
[void]$PRTG.AppendChild($Result)

# Event type

$Result = $xmldoc.CreateElement('Result')

$Channel = $xmldoc.CreateElement('Channel')
$Unit = $xmldoc.CreateElement('Unit')
$Mode = $xmldoc.CreateElement('Mode')
$ValueLookup =  $xmldoc.CreateElement('ValueLookup')
$Value = $xmldoc.CreateElement('Value')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$NotifyChanged = $xmldoc.CreateElement('NotifyChanged')

$Channel.InnerText = "Event Type"
$Unit.InnerText = "Custom"
$Mode.Innertext = "Absolute"
$Value.Innertext = $RP_eventtype
$CustomUnit.InnerText = "Code"

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
[void]$Result.AppendChild($NotifyChanged)
[void]$Result.AppendChild($Mode)
    
[void]$PRTG.AppendChild($Result)

# Age

$Result = $xmldoc.CreateElement('Result')

$Channel = $xmldoc.CreateElement('Channel')
$Unit = $xmldoc.CreateElement('Unit')
$Mode = $xmldoc.CreateElement('Mode')
$Value = $xmldoc.CreateElement('Value')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$NotifyChanged = $xmldoc.CreateElement('NotifyChanged')

$Channel.InnerText = "Restore Point Age"
$Unit.InnerText = "Custom"
$CustomUnit.InnerText = "Days"
$Mode.Innertext = "Absolute"
$Value.Innertext = $RP_agedays

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
[void]$Result.AppendChild($NotifyChanged)
[void]$Result.AppendChild($ValueLookup)
[void]$Result.AppendChild($Mode)
    
[void]$PRTG.AppendChild($Result)

# Number of restore points

$Result = $xmldoc.CreateElement('Result')

$Channel = $xmldoc.CreateElement('Channel')
$Unit = $xmldoc.CreateElement('Unit')
$Mode = $xmldoc.CreateElement('Mode')
$Value = $xmldoc.CreateElement('Value')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$NotifyChanged = $xmldoc.CreateElement('NotifyChanged')
$LimitMode = $xmldoc.CreateElement('LimitMode')
$LimitMinWarning = $xmldoc.CreateElement('LimitMinWarning')
$LimitMinError = $xmldoc.CreateElement('LimitMinError')

$Channel.InnerText = "Number of Restore Points"
$Unit.InnerText = "Count"
# $CustomUnit.InnerText = 
$Mode.Innertext = "Absolute"
$Value.Innertext = $RP_count

$LimitMode.InnerText = "1"
$LimitMinWarning.InnerText = 2
$LimitMinError.Innertext = 1

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
[void]$Result.AppendChild($NotifyChanged)
[void]$Result.AppendChild($ValueLookup)
[void]$Result.AppendChild($Mode)

[void]$Result.AppendChild($LimitMode) 
[void]$Result.AppendChild($LimitMinWarning)
[void]$Result.AppendChild($LimitMinError)
    
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
    $formattime = $now.ToString("dd-MM-yyyy HH:mm:ss")
    $message = "Machine $myhost (now $uptext) *** Last Restore Point: $RP_timetext ($RP_desc) *** Script $scriptversion"
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