param (
    [string]$LOGGING = "NO", 
    [int]$sensorid = 77 
)
# $LOGGING = 'YES'

$logging = $logging.ToUpper()

$ScriptVersion = " -- Version: 1.2.1"

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
    if ($log) {
        Add-Content $logfile $StringWriter.ToString()
    }
}

# END OF COMMON CODING

#--- Creëer HTTP header -------------------------------------------------
$baseurl = 'https://openapi.alphaess.com/api'

# Bereken UNIX timestamp (aantal seconden sinds 1-1-1970 om 00:00 uur tot huidige UTC tijd)
$StartDate = Get-Date -year 1970 -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0 -Millisecond 0
$EndDate = (Get-Date).ToUniversalTime()
$diff = NEW-TIMESPAN –Start $StartDate –End $EndDate
$unixTimestamp = [Math]::Floor([decimal]($diff.TotalSeconds))

# SIGN string conform api documentatie
$secureString = $ADHC_ESSappId + $ADHC_ESSsecretkey + $unixtimestamp
# encrypt SIGN string met $HA512
$bytes = [System.Text.Encoding]::UTF8.GetBytes($securestring)
$sha512 = [System.Security.Cryptography.SHA512]::Create()
$hashBytes = $sha512.ComputeHash($bytes)
# Omzetten naar hexadecimale string
$sign = [System.BitConverter]::ToString($hashBytes).Replace("-", "").ToLower()

# creëer header conform API documentatie
$headers = @{
    appId = $ADHC_ESSappId
    timeStamp = $unixtimestamp
    sign = $sign

}

if (!$scripterror) {
    try {
        # input parameter conform api documentatie
        $parm = "?" + "sysSn=" + "ALD081025090564"
        $verb = "/getSumDataForCustomer"
        $uri = $baseUrl + $verb + $parm
        $response = Invoke-RestMethod -Method Get `
            -Uri "$uri" `
            -Headers $headers `
            -ContentType 'application/json'

        if ($log) {
             Add-Content $logfile " "
             Add-Content $logfile $uri   
             Add-Content $logfile $response
             Add-Content $logfile $response.data  
        }

    }
    Catch {
        if ($log) {
            Add-Content $logfile "==> API call 1 failed"
        }
        $scripterror = $true
        $errortext = $error[0]
        $scripterrormsg = "==> API call 1 failed - $errortext"

    }
    finally {
        
    }
}

#$response
#$response.data
$VandaagOpgewekt      =  $response.data.epvtoday
$VandaagVerbruikt     =  $response.data.eload
$Vandaagteruggeleverd =  $response.data.eoutput
$VandaagNetafname     =  $response.data.einput
$VandaagOpgeladen     =  $response.data.echarge
$VandaagOntladen      =  $response.data.edischarge

$TotaalOpwekking           =  $response.data.epvtotal
$TotaalZelfverbruikPerc    =  $response.data.eselfConsumption
$TotaalZelfvoorzieningPerc = $response.data.eselfSufficiency

if (!$scripterror) {
    try {
        # input parameter conform api documentatie
        $parm = ""
        $verb = "/getEssList"
        $uri = $baseUrl + $verb + $parm
        $response = Invoke-RestMethod -Method Get `
            -Uri "$uri" `
            -Headers $headers `
            -ContentType 'application/json'

        if ($log) {
            Add-Content $logfile " "
            Add-Content $logfile $uri   
            Add-Content $logfile $response
            Add-Content $logfile $response.data    
        }
    }
    Catch {
        if ($log) {
            Add-Content $logfile "==> API call 2 failed"
        }
        $scripterror = $true
        $errortext = $error[0]
        $scripterrormsg = "==> API call 2 failed - $errortext"

    }
    finally {
       
    }
}

#$response
#$response.data
$systeemnaam = $response.data.sysSn
$systeemstatus = $response.data.emsstatus

if (!$scripterror) {
    try {
        # input parameter conform api documentatie
        $parm = "?" + "sysSn=" + "ALD081025090564"
        $verb = "/getLastPowerData"
        $uri = $baseUrl + $verb + $parm
        $response = Invoke-RestMethod -Method Get `
            -Uri "$uri" `
            -Headers $headers `
            -ContentType 'application/json'

        if ($log) {  
            Add-Content $logfile " "
            Add-Content $logfile $uri   
            Add-Content $logfile $response
            Add-Content $logfile $response.data    
        }
    }
    Catch {
        if ($log) {
            Add-Content $logfile "==> API call 3 failed"
        }
        $scripterror = $true
        $errortext = $error[0]
        $scripterrormsg = "==> API call 2 failed - $errortext"

    }
    finally {
       
    }
}

#$response
#$response.data
#write-host "Oplaadpercentage      " $response.data.soc
$oplaadpercentage = $response.data.soc
$zonnelevering = $response.data.ppv
$pbat = $response.data.pbat
if ($pbat -gt 0) {
    $batterijlevering = $pbat
    $batterijafname = 0
}
else {
    $batterijafname = $pbat * -1
    $batterijlevering = 0
}
$huisafname = $response.data.pload
$pgrid = $response.data.pgrid
if ($pgrid -gt 0) {
    $gridlevering = $pgrid
    $gridafname = 0
}
else {
    $gridafname = $pgrid *-1
    $gridlevering = 0
}

if ($log) {
    Add-Content $logfile "==> Write XML"
}

[xml]$xmldoc = New-Object System.Xml.XmlDocument
$decl = $xmldoc.CreateXmlDeclaration('1.0','Windows-1252',$null)

[void]$xmldoc.AppendChild($decl)

$PRTG = $xmldoc.CreateElement('PRTG')

# Batterij laadstatus (PRIMARY Channel)
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')
$Float = $xmldoc.CreateElement('Float')
$Unit = $xmldoc.CreateElement('Unit')
$Mode = $xmldoc.CreateElement('Mode')
$NotifyChanged = $xmldoc.CreateElement('NotifyChanged')

$Channel.InnerText = "Oplaadpercentage"
$Value.InnerText = $oplaadpercentage
$Float.InnerText = "1"
$Unit.InnerText = "Percent"
$Mode.Innertext = "Absolute"

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Float)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($NotifyChanged)
[void]$Result.AppendChild($Mode)
    
[void]$PRTG.AppendChild($Result)

# CURRENT
# Aanvoer vanuit het net 
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')
$Float = $xmldoc.CreateElement('Float')
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
$NotifyChanged = $xmldoc.CreateElement('NotifyChanged')

$Channel.InnerText = "Aanvoer vanuit het net"
$Value.InnerText = $gridlevering
$Float.InnerText = "1"
$Unit.InnerText = "Custom"
$CustomUnit.InnerText = "Watt"
$Mode.Innertext = "Absolute"

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Float)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
[void]$Result.AppendChild($NotifyChanged)
[void]$Result.AppendChild($Mode)
    
[void]$PRTG.AppendChild($Result)

# Aanvoer vanuit de zonnepanelen 
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')
$Float = $xmldoc.CreateElement('Float')
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
$NotifyChanged = $xmldoc.CreateElement('NotifyChanged')

$Channel.InnerText = "Aanvoer vanuit de zonnepanelen"
$Value.InnerText = $zonnelevering
$Float.InnerText = "1"
$Unit.InnerText = "Custom"
$CustomUnit.InnerText = "Watt"
$Mode.Innertext = "Absolute"

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Float)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
[void]$Result.AppendChild($NotifyChanged)
[void]$Result.AppendChild($Mode)
    
[void]$PRTG.AppendChild($Result)

# Aanvoer vanuit de battterij 
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')
$Float = $xmldoc.CreateElement('Float')
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
$NotifyChanged = $xmldoc.CreateElement('NotifyChanged')

$Channel.InnerText = "Aanvoer vanuit de batterij"
$Value.InnerText = $batterijlevering
$Float.InnerText = "1"
$Unit.InnerText = "Custom"
$CustomUnit.InnerText = "Watt"
$Mode.Innertext = "Absolute"

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Float)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
[void]$Result.AppendChild($NotifyChanged)
[void]$Result.AppendChild($Mode)
    
[void]$PRTG.AppendChild($Result)

# Levering het net 
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')
$Float = $xmldoc.CreateElement('Float')
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
$NotifyChanged = $xmldoc.CreateElement('NotifyChanged')

$Channel.InnerText = "Levering aan het net"
$Value.InnerText = $gridafname
$Float.InnerText = "1"
$Unit.InnerText = "Custom"
$CustomUnit.InnerText = "Watt"
$Mode.Innertext = "Absolute"

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Float)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
[void]$Result.AppendChild($NotifyChanged)
[void]$Result.AppendChild($Mode)
    
[void]$PRTG.AppendChild($Result)

# Levering aan het huis 
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')
$Float = $xmldoc.CreateElement('Float')
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
$NotifyChanged = $xmldoc.CreateElement('NotifyChanged')

$Channel.InnerText = "Levering aan het huis"
$Value.InnerText = $huisafname
$Float.InnerText = "1"
$Unit.InnerText = "Custom"
$CustomUnit.InnerText = "Watt"
$Mode.Innertext = "Absolute"

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Float)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
[void]$Result.AppendChild($NotifyChanged)
[void]$Result.AppendChild($Mode)
    
[void]$PRTG.AppendChild($Result)

# Levering de battterij 
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')
$Float = $xmldoc.CreateElement('Float')
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
$NotifyChanged = $xmldoc.CreateElement('NotifyChanged')

$Channel.InnerText = "Levering aan de batterij"
$Value.InnerText = $batterijafname
$Float.InnerText = "1"
$Unit.InnerText = "Custom"
$CustomUnit.InnerText = "Watt"
$Mode.Innertext = "Absolute"

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Float)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
[void]$Result.AppendChild($NotifyChanged)
[void]$Result.AppendChild($Mode)
    
[void]$PRTG.AppendChild($Result)


# Totaal VANDAAG

# Aanvoer vanuit het net 
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')
$Float = $xmldoc.CreateElement('Float')
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
$NotifyChanged = $xmldoc.CreateElement('NotifyChanged')

$Channel.InnerText = "Vandaag netafname"
$Value.InnerText = $VandaagNetafname 
$Float.InnerText = "1"
$Unit.InnerText = "Custom"
$CustomUnit.InnerText = "KwH"
$Mode.Innertext = "Absolute"

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Float)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
[void]$Result.AppendChild($NotifyChanged)
[void]$Result.AppendChild($Mode)
    
[void]$PRTG.AppendChild($Result)

# Aanvoer vanuit de zonnepanelen 
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')
$Float = $xmldoc.CreateElement('Float')
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
$NotifyChanged = $xmldoc.CreateElement('NotifyChanged')

$Channel.InnerText = "Vandaag opgewekt"
$Value.InnerText = $VandaagOpgewekt
$Float.InnerText = "1"
$Unit.InnerText = "Custom"
$CustomUnit.InnerText = "KwH"
$Mode.Innertext = "Absolute"

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Float)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
[void]$Result.AppendChild($NotifyChanged)
[void]$Result.AppendChild($Mode)
    
[void]$PRTG.AppendChild($Result)

# Aanvoer vanuit de battterij 
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')
$Float = $xmldoc.CreateElement('Float')
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
$NotifyChanged = $xmldoc.CreateElement('NotifyChanged')

$Channel.InnerText = "Vandaag batterij ontladen"
$Value.InnerText = $VandaagOntladen      
$Float.InnerText = "1"
$Unit.InnerText = "Custom"
$CustomUnit.InnerText = "KwH"
$Mode.Innertext = "Absolute"

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Float)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
[void]$Result.AppendChild($NotifyChanged)
[void]$Result.AppendChild($Mode)
    
[void]$PRTG.AppendChild($Result)

# Levering aan het net 
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')
$Float = $xmldoc.CreateElement('Float')
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
$NotifyChanged = $xmldoc.CreateElement('NotifyChanged')

$Channel.InnerText = "Vandaag teruglevering net"
$Value.InnerText = $Vandaagteruggeleverd
$Float.InnerText = "1"
$Unit.InnerText = "Custom"
$CustomUnit.InnerText = "KwH"
$Mode.Innertext = "Absolute"

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Float)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
[void]$Result.AppendChild($NotifyChanged)
[void]$Result.AppendChild($Mode)
    
[void]$PRTG.AppendChild($Result)

# Levering aan het huis 
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')
$Float = $xmldoc.CreateElement('Float')
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
$NotifyChanged = $xmldoc.CreateElement('NotifyChanged')

$Channel.InnerText = "Vandaag verbruikt"
$Value.InnerText = $VandaagVerbruikt
$Float.InnerText = "1"
$Unit.InnerText = "Custom"
$CustomUnit.InnerText = "KwH"
$Mode.Innertext = "Absolute"

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Float)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
[void]$Result.AppendChild($NotifyChanged)
[void]$Result.AppendChild($Mode)
    
[void]$PRTG.AppendChild($Result)

# Levering aan de battterij 
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')
$Float = $xmldoc.CreateElement('Float')
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
$NotifyChanged = $xmldoc.CreateElement('NotifyChanged')

$Channel.InnerText = "Vandaag batterij opgeladen"
$Value.InnerText = $VandaagOpgeladen
$Float.InnerText = "1"
$Unit.InnerText = "Custom"
$CustomUnit.InnerText = "KwH"
$Mode.Innertext = "Absolute"

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Float)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
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
    $Errorvalue.InnerText = "0"  
    $formattime = $d.ToString("dd-MM-yyyy HH:mm:ss")
    $ErrorText.InnerText = "$systeemnaam heeft status '$systeemstatus' *** Timestamp: $formattime *** Script$scriptversion"
} 
[void]$PRTG.AppendChild($ErrorValue)
[void]$PRTG.AppendChild($ErrorText)
    
[void]$xmldoc.Appendchild($PRTG)


writeXmlToScreen $xmldoc

if ($log) {
    $thisdate = Get-Date
    Add-Content $logfile "==> END $thisdate"
}



