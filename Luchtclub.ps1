param (
    [string]$LOGGING = "NO", 
    [int]$sensorid = 77 
)
# $LOGGING = 'NO'

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
}

# END OF COMMON CODING

try {
    # Start with link tot my sensor
    $uri = "https://api-samenmeten.rivm.nl/v1.0/Things(5361)"
    $cururi = $uri
    if ($log) {
        Add-Content $logfile "==> Get sensor object (Thing) <$cururi> "          
    }    
    $Sensorobj = Invoke-RestMethod  -Uri $uri 
    $sensorid = $Sensorobj.'@iot.id'
    $sensorname = $sensorobj.name

    # next links to follow 
    $locationlink = $Sensorobj.'Locations@iot.navigationLink'
    $datastreamlink = $Sensorobj.'Datastreams@iot.navigationLink'

    # get location of sensor
    $uri = $locationlink
    $cururi = $uri
    if ($log) {
        Add-Content $logfile "==> Get sensor location <$cururi> "          
    }    
    $Locationobj = Invoke-RestMethod  -Uri $uri 

    $Oosterlengte = $locationobj.value.location.coordinates[0]
    $Noorderbreedte = $locationobj.value.location.coordinates[1]

    # get data from sensor
    $uri = $datastreamlink
    $cururi = $uri
    if ($log) {
        Add-Content $logfile "==> Get data object <$cururi> "          
    }    
    $Datastreamobj = Invoke-RestMethod  -Uri $uri 
    
    $Metingen = @()
    $Currentdate = Get-Date
    $CurrentTime = $Currentdate.ToString()
    
    # loop through metingen
    for ($y=0; $y -le $Datastreamobj.value.count-1; $y++) {
        $measurement = $Datastreamobj.value[$y]
        $MetingObj = [PSCustomObject] [ordered] @{Unit = "?";
                                            Name = "?";
                                            Description = "?";
                                          Observationslink = "?";
                                          Propertylink = "?";
                                          LastMeasureTime = '';
                                          CurrentTime = '';
                                          MinutesAgo = 0;
                                          LastResult = 0
                                          }
        $MetingObj.Unit = $measurement.unitOfMeasurement.Symbol
        $MetingObj.Observationslink = $measurement.'Observations@iot.navigationLink'
        $MetingObj.Propertylink = $measurement.'ObservedProperty@iot.navigationLink'

        # get the most recent data of this meting
        $uri2 = $MetingObj.Observationslink
        $cururi = $uri2
        if ($log) {
            Add-Content $logfile "==> Get most recent measurement <$cururi> "          
        }
        
        $ObservationsObj = Invoke-RestMethod  -Uri $uri2 

        # the first meting in the list is the most recent one
        $MetingObj.LastMeasureTime = Get-Date -Date $ObservationsObj.value[0].phenomenonTime 
        $Hulp = Get-Date -Date $MetingObj.LastMeasureTime       
        $diff = NEW-TIMESPAN -Start $Hulp -End $CurrentDate
                
        $MetingObj.MinutesAgo = [math]::round($diff.TotalMinutes,0)
        $MetingObj.LastResult = $ObservationsObj.value[0].result
        $MetingObj.CurrentTime = $CurrentTime

        # Get the properties of this meting
        $uri3 = $Metingobj.PropertyLink
        $cururi = $uri3
        if ($log) {
            Add-Content $logfile "==> Get properties of this measurement <$cururi> "          
        }
       
        $PropertyObj = Invoke-RestMethod  -Uri $uri3 

        $MetingObj.Name = $PropertyObj.name
        $MetingObj.Description = $PropertyObj.description

        $Metingen += $MetingObj
     
    }

}

catch {
    $scripterror = $true
    $errortext = $error[0]  
    $scripterrormsg = "==> API Processing <$cururi> to RIVM site failed - $errortext" 
    if ($log) {
        Add-Content $logfile $scripterrormsg         
    }    

}

# $metingen

if ($log) {
    Add-Content $logfile "==> Create XML"
}

[xml]$xmldoc = New-Object System.Xml.XmlDocument
$decl = $xmldoc.CreateXmlDeclaration('1.0','Windows-1252',$null)

[void]$xmldoc.AppendChild($decl)

$PRTG = $xmldoc.CreateElement('PRTG')

# Luchtclub metingen
foreach ($item in $metingen) {
    $Result = $xmldoc.CreateElement('Result')
    $Channel = $xmldoc.CreateElement('Channel')
    $Unit = $xmldoc.CreateElement('Unit')
    $CustomUnit = $xmldoc.CreateElement('CustomUnit')
    $Value = $xmldoc.CreateElement('Value')    
    $Mode = $xmldoc.CreateElement('Mode')
    $FLoat = $xmldoc.CreateElement('Float')    

    $Channel.InnerText = $item.Description + " (" + $item.Name + ")"
    $Unit.InnerText = "Custom"
    $Customunit.Innertext = $item.unit
    $Value.InnerText = $item.Lastresult
    $Mode.Innertext = "Absolute"
    $Float.InnerText = "1"
    
    [void]$Result.AppendChild($Channel)    
    [void]$Result.AppendChild($Unit)
    [void]$Result.AppendChild($CustomUnit)
    [void]$Result.AppendChild($Value)
    [void]$Result.AppendChild($Mode)
    [void]$Result.AppendChild($Float)    

    [void]$PRTG.AppendChild($Result)

    $Result = $xmldoc.CreateElement('Result')
    $Channel = $xmldoc.CreateElement('Channel')
    $Unit = $xmldoc.CreateElement('Unit')
    $CustomUnit = $xmldoc.CreateElement('CustomUnit')
    $Value = $xmldoc.CreateElement('Value')    
    $Mode = $xmldoc.CreateElement('Mode')
    $FLoat = $xmldoc.CreateElement('Float')    

    $Channel.InnerText = "Observation age "+ $item.Name
    $Unit.InnerText = "Custom"
    $Customunit.Innertext = "Minutes"
    $Value.InnerText = $item.MinutesAgo
    $Mode.Innertext = "Absolute"
    $Float.InnerText = "0"
    
    [void]$Result.AppendChild($Channel)    
    [void]$Result.AppendChild($Unit)
    [void]$Result.AppendChild($CustomUnit)
    [void]$Result.AppendChild($Value)
    [void]$Result.AppendChild($Mode)
    [void]$Result.AppendChild($Float)    

    [void]$PRTG.AppendChild($Result)
    
}
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Value = $xmldoc.CreateElement('Value')    
$Mode = $xmldoc.CreateElement('Mode')
$FLoat = $xmldoc.CreateElement('Float')    

$Channel.InnerText = "Oosterlengte"
$Unit.InnerText = "Custom"
$Customunit.Innertext = "Graden"
$Value.InnerText = $Oosterlengte
$Mode.Innertext = "Absolute"
$Float.InnerText = "1"
    
[void]$Result.AppendChild($Channel)    
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Mode)
[void]$Result.AppendChild($Float)    

[void]$PRTG.AppendChild($Result)

$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Value = $xmldoc.CreateElement('Value')    
$Mode = $xmldoc.CreateElement('Mode')
$FLoat = $xmldoc.CreateElement('Float')    

$Channel.InnerText = "Noorderbreedte"
$Unit.InnerText = "Custom"
$Customunit.Innertext = "Graden"
$Value.InnerText = $Noorderbreedte
$Mode.Innertext = "Absolute"
$Float.InnerText = "1"
    
[void]$Result.AppendChild($Channel)    
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Mode)
[void]$Result.AppendChild($Float)    

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
    $message = "Luchtclub sensor $sensorname ($sensorid) *** Script $scriptversion"
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
