param (
    [string]$LOGGING = "YES", 
    [int]$sensorid = 77 
)
# $LOGGING = 'NO'

$ScriptVersion = " -- Version: 1.0.7"

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
$myhost = $ADHC_Computer

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
    if ($log) {
        Add-Content $logfile $StringWriter.ToString()
    }
}

# END OF COMMON CODING

try {
    # Read dataset that is being filled by the LuchtCLub Powershell batch script
    $timestamp = Get-Date
    $str = $ADHC_LuchtClubInfo.Split("\")
    $dir = $ADHC_OutputDirectory + $str[0] + "\" + $str[1]
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    $LuchtClubInfoFile = $ADHC_OutputDirectory + $ADHC_LuchtClubInfo.Replace($ADHC_Computer, $myHost)
    $dr = Test-Path $LuchtClubInfoFile
    if (!$dr) { 
        if ($log) {
            Add-Content $logfile "$LuchtClubInfoFile does not exist, presume all values to be zero"
        }
        $TotalsensorsCurrent = 0
        $ActivesensorsCurrent = 0
        $InactivesensorsCurrent = 0
        $ExcludedsensorsCurrent = 0
        $TotalsensorsDelta = 0
        $ActivesensorsDelta = 0
        $InactivesensorsDelta = 0
        $ExcludedsensorsDelta = 0
    }
    else {    
        $luchtlines = Get-Content $LuchtClubInfoFile
        $deltalist = "Delta's found for: "      
        foreach ($line in $luchtlines) {
            $split = $line.Split("|")
            $subject = $split[2]
            
            Switch ($subject) {
                "Total" {
                    $TotalsensorsCurrent = $split[0]                    
                    $TotalsensorsDelta = $split[0] - $split[1]  
                    if ($TotalsensorsDelta -ne 0) {
                        $scripterror = $true
                        $deltalist = $deltalist + "Total sensors - "
                    }                
                }  
                "Active" {
                    $ActivesensorsCurrent = $split[0]
                    $ActivesensorsDelta = $split[0] - $split[1]
                    if ($ActivesensorsDelta -ne 0) {
                        $scripterror = $true
                        $deltalist = $deltalist + "Active sensors - "
                    }     
                }
                "Inactive" {                 
                    $InactivesensorsCurrent = $split[0]
                    $InactivesensorsDelta = $split[0] - $split[1]
                    if ($InActivesensorsDelta -ne 0) {
                        $scripterror = $true
                        $deltalist = $deltalist + "InActive sensors - "
                    }     
                }
                "Excluded" {
                    $ExcludedsensorsCurrent = $split[0]
                    $ExcludedsensorsDelta = $split[0] - $split[1]
                    if ($ExcludedsensorsDelta -ne 0) {
                        $scripterror = $true
                        $deltalist = $deltalist + "Excluded sensors "
                    }       

                }
                Default {
                    $emsg = "Invalid tag $subject encountered, timestamp = " + $split[3]
                    if ($log) {
                        Add-Content $logfile $emsg        
                    }    
                    throw $emsg
                    
                }
            }
                      
        } 
        if ($scripterror) {
            $scripterrormsg = $deltalist
        }
    } 
}

catch {
    $scripterror = $true
    $scripterrormsg = $error[0]  
    if ($log) {
        Add-Content $logfile $scripterrormsg       
    }    

}


if ($log) {
    Add-Content $logfile "==> Create XML"
}

[xml]$xmldoc = New-Object System.Xml.XmlDocument
$decl = $xmldoc.CreateXmlDeclaration('1.0','Windows-1252',$null)

[void]$xmldoc.AppendChild($decl)

$PRTG = $xmldoc.CreateElement('PRTG')

# Active sensors
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('Custom')
$Value = $xmldoc.CreateElement('Value')    
$Mode = $xmldoc.CreateElement('Mode') 
$Float = $xmldoc.CreateElement('Float')  

$Channel.InnerText = "Active Sensors"
$Unit.InnerText = "Custom"
$Customunit.Innertext = "Number of sensors"
$Value.InnerText = $ActiveSensorsCurrent
$Mode.Innertext = "Absolute"
$Float.InnerText = "0"
    
[void]$Result.AppendChild($Channel)    
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Mode)
[void]$Result.AppendChild($Float)    

[void]$PRTG.AppendChild($Result)

# Active sensors delta
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('Custom')
$Value = $xmldoc.CreateElement('Value')    
$Mode = $xmldoc.CreateElement('Mode')   
$Float = $xmldoc.CreateElement('Float') 
$LimitMode = $xmldoc.CreateElement('LimitMode')
$LimitMinError = $xmldoc.CreateElement('LimitMinError')
$LimitMaxError = $xmldoc.CreateElement('LimitMaxError')

$Channel.InnerText = "Active Sensors Delta"
$Unit.InnerText = "Custom"
$Customunit.Innertext = "Number of sensors"
$Value.InnerText = $ActiveSensorsDelta
$Mode.Innertext = "Absolute"
$Float.InnerText = "0"
$LimitMode.InnerText = "1"
$LimitMinError.InnerText = "0"
$LimitMaxError.InnerText = "0"
    
[void]$Result.AppendChild($Channel)    
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Mode)
[void]$Result.AppendChild($Float) 
[void]$Result.AppendChild($LimitMode) 
[void]$Result.AppendChild($LimitMinError) 
[void]$Result.AppendChild($LimitMaxError)  

[void]$PRTG.AppendChild($Result)

# InActive sensors
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('Custom')
$Value = $xmldoc.CreateElement('Value')    
$Mode = $xmldoc.CreateElement('Mode') 
$Float = $xmldoc.CreateElement('Float')   

$Channel.InnerText = "InActive Sensors"
$Unit.InnerText = "Custom"
$Customunit.Innertext = "Number of sensors"
$Value.InnerText = $InActiveSensorsCurrent
$Mode.Innertext = "Absolute"
$Float.InnerText = "0"
    
[void]$Result.AppendChild($Channel)    
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Mode)
[void]$Result.AppendChild($Float)    

[void]$PRTG.AppendChild($Result)

# InActive sensors delta
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('Custom')
$Value = $xmldoc.CreateElement('Value')    
$Mode = $xmldoc.CreateElement('Mode')  
$Float = $xmldoc.CreateElement('Float') 
$LimitMode = $xmldoc.CreateElement('LimitMode')
$LimitMinError = $xmldoc.CreateElement('LimitMinError')
$LimitMaxError = $xmldoc.CreateElement('LimitMaxError') 

$Channel.InnerText = "InActive Sensors Delta"
$Unit.InnerText = "Custom"
$Customunit.Innertext = "Number of sensors"
$Value.InnerText = $InActiveSensorsDelta
$Mode.Innertext = "Absolute"
$Float.InnerText = "0"
$LimitMode.InnerText = "1"
$LimitMinError.InnerText = "0"
$LimitMaxError.InnerText = "0"
    
[void]$Result.AppendChild($Channel)    
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Mode)
[void]$Result.AppendChild($Float) 
[void]$Result.AppendChild($LimitMode) 
[void]$Result.AppendChild($LimitMinError) 
[void]$Result.AppendChild($LimitMaxError)     

[void]$PRTG.AppendChild($Result)

# Exluded
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('Custom')
$Value = $xmldoc.CreateElement('Value')    
$Mode = $xmldoc.CreateElement('Mode')  
$Float = $xmldoc.CreateElement('Float')  

$Channel.InnerText = "Excluded Sensors"
$Unit.InnerText = "Custom"
$Customunit.Innertext = "Number of sensors"
$Value.InnerText = $ExcludedSensorsCurrent
$Mode.Innertext = "Absolute"
$Float.InnerText = "0"
    
[void]$Result.AppendChild($Channel)    
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Mode)
[void]$Result.AppendChild($Float)    

[void]$PRTG.AppendChild($Result)

# Excluded delta
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('Custom')
$Value = $xmldoc.CreateElement('Value')    
$Mode = $xmldoc.CreateElement('Mode')   
$Float = $xmldoc.CreateElement('Float') 
$LimitMode = $xmldoc.CreateElement('LimitMode')
$LimitMinError = $xmldoc.CreateElement('LimitMinError')
$LimitMaxError = $xmldoc.CreateElement('LimitMaxError')

$Channel.InnerText = "Excluded Sensors Delta"
$Unit.InnerText = "Custom"
$Customunit.Innertext = "Number of sensors"
$Value.InnerText = $ExcludedSensorsDelta
$Mode.Innertext = "Absolute"
$Float.InnerText = "0"
$LimitMode.InnerText = "1"
$LimitMinError.InnerText = "0"
$LimitMaxError.InnerText = "0"
    
[void]$Result.AppendChild($Channel)    
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Mode)
[void]$Result.AppendChild($Float) 
[void]$Result.AppendChild($LimitMode) 
[void]$Result.AppendChild($LimitMinError) 
[void]$Result.AppendChild($LimitMaxError)     

[void]$PRTG.AppendChild($Result)

# Total sensors
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('Custom')
$Value = $xmldoc.CreateElement('Value')    
$Mode = $xmldoc.CreateElement('Mode')  
$Float = $xmldoc.CreateElement('Float')  

$Channel.InnerText = "Total Sensors"
$Unit.InnerText = "Custom"
$Customunit.Innertext = "Number of sensors"
$Value.InnerText = $TotalSensorsCurrent
$Mode.Innertext = "Absolute"
$Float.InnerText = "0"
    
[void]$Result.AppendChild($Channel)    
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Mode)
[void]$Result.AppendChild($Float)    

[void]$PRTG.AppendChild($Result)

# Total sensors delta
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('Custom')
$Value = $xmldoc.CreateElement('Value')    
$Mode = $xmldoc.CreateElement('Mode')   
$Float = $xmldoc.CreateElement('Float') 
$LimitMode = $xmldoc.CreateElement('LimitMode')
$LimitMinError = $xmldoc.CreateElement('LimitMinError')
$LimitMaxError = $xmldoc.CreateElement('LimitMaxError')

$Channel.InnerText = "Total Sensors Delta"
$Unit.InnerText = "Custom"
$Customunit.Innertext = "Number of sensors"
$Value.InnerText = $TotalSensorsDelta
$Mode.Innertext = "Absolute"
$Float.InnerText = "0"
$LimitMode.InnerText = "1"
$LimitMinError.InnerText = "0"
$LimitMaxError.InnerText = "0"
    
[void]$Result.AppendChild($Channel)    
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Mode)
[void]$Result.AppendChild($Float) 
[void]$Result.AppendChild($LimitMode) 
[void]$Result.AppendChild($LimitMinError) 
[void]$Result.AppendChild($LimitMaxError)     

[void]$PRTG.AppendChild($Result)
    

# Add error block

$ErrorValue = $xmldoc.CreateElement('Error')
$ErrorText = $xmldoc.CreateElement('Text')

if ($scripterror) {
    $Errorvalue.InnerText = "0"
    $ErrorText.InnerText = $scripterrormsg + " *** Scriptversion=$scriptversion *** "
}
else {
    $Errorvalue.InnerText = "0"    
    $message = "Luchtclub monitoring (ID=$sensorid) - no delta's encountered *** Script $scriptversion"
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
