param (
    [string]$LOGGING = "YES", 
    [string]$Ctype = "NONE",
    [int]$sensorid = 77
)

$Ctype = $Ctype.ToUpper()

$ScriptVersion = " -- Version: 2.0.3"

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

$scripterror = $false

if (!$ADHC_InitSuccessfull) {
    # Write-Warning "YES"
    throw $ADHC_InitError
}

if ($LOGGING.ToUpper() -eq "YES") {$log = $true} else {$log = $false}

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
        Add-Content $logfile "==> Read IP address database"
    }
    $query = "SELECT  [Naam],[IPaddress],[Type] FROM [PRTG].[dbo].[IPadressen] WHERE [Type] = '" + $CTYPE + "'"
    $ctypelist = invoke-sqlcmd -ServerInstance ".\SQLEXPRESS" -Database "PRTG" `
                        -Query "$query" `
                        -ErrorAction Stop
    if (!$ctypelist) {
        if ($log) {
            Add-Content $logfile "==> No Node found with type $ctype"
            $scripterrormsg = "No Node found with type $ctype"
            $scripterror = $true
        }
    }         
}
catch {
    
    $scripterror = $true
    $errortext = $error[0]
     
    $scripterrormsg = "Reading IP address database failed --- $errortext"
    if ($log) {
        Add-Content $logfile "==> $scripterrormsg"
    }    
}

$maxcode = 0
$nrofping = 0
$nrofcached = 0
$nrofboth = 0
$nrofinactive = 0
$nroftotal = 0
$nroferror = 0


if (!$scripterror) {
    try {
        if ($log) {
            Add-Content $logfile "==> Check status of every IP with type $ctype"          
        }
        $ResultList = @()
        # Check al entries from database against 
        foreach ( $dbentry in $ctypelist) {
            
            $i = $dbentry.IPaddress 
            $nstat = & $ADHC_NodeInfoScript "$i" "$LOGGING"
            
            $ipobj = [PSCustomObject] [ordered] @{IPname = $dbentry.Naam;
                                            IPaddress = $nstat.IPaddress;
                                            IPcached = $nstat.IPcached; 
                                            IPtype = $ctype;
                                            IPping = $nstat.IPpingable;
                                            IPstatus = $nstat.Status;
                                            IPstatusCode = $nstat.StatusCode}
            $ResultList += $ipobj
            $nroftotal += 1
            if ($log) {
                foreach ($m in $nstat.MessageList) {
                    $lvl = $m.Level
                    $msg = $m.Message
                    Add-COntent $logfile "($lvl) - $msg"

                }
            }
            switch ($nstat.StatusCode) {
                0 { $nrofinactive += 1 } 
                3 { $nrofcached +=1 }
                6 { $nrofping += 1 } 
                9 { $nrofboth +=1} 
                12 { $nroferror +=1}
                default {
                    $x = $nstat.Statuscode 
                    $scripterrormsg = "==> $x is an unknown IP statuscode" 
                    if ($log) {
                        Add-Content $logfile $scripterrormsg                                 
                    }
                    $scripterror = $true

                }
            }
            $maxcode = [math]::max($maxcode, $nstat.Statuscode) 
                      
        }
    }
    catch {
        
        $scripterror = $true
        $errortext = $error[0]
        $scripterrormsg = "==> IP status check failed - $errortext"
        if ($log) {
            Add-Content $logfile "==> $scripterrormsg"
          
        }
    }
    finally {
       
    }
}


if ($log) {
    Add-Content $logfile "==> Create XML"
}

[xml]$xmldoc = New-Object System.Xml.XmlDocument
$decl = $xmldoc.CreateXmlDeclaration('1.0','Windows-1252',$null)

[void]$xmldoc.AppendChild($decl)

$PRTG = $xmldoc.CreateElement('PRTG')

# Overall status

$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
$NotifyChanged = $xmldoc.CreateElement('NotifyChanged')
$ValueLookup =  $xmldoc.CreateElement('ValueLookup')

$Channel.InnerText = "Overall status $Ctype nodes"
$Unit.InnerText = "Custom"
$Mode.Innertext = "Absolute"
$ValueLookup.Innertext = 'OverallCtypeStatus'

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

    $ipsplit = $item.IPaddress.Split(" .")
    $ipnr = (“{0:d3}” -f [int]$ipsplit[3].Trim()) 
    $cname =   "(IP" + $ipnr + ") " + $item.IPname.Trim()

    $Channel.InnerText = $cname
    $Unit.InnerText = "Custom"
    $Mode.Innertext = "Absolute"
    $ValueLookup.Innertext = 'IndividualCtypeStatus'
           
    $Value.Innertext = $item.IPstatusCode

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
   
    $message = "Nodes in group $ctype *** Total: $nroftotal *** Inactive: $nrofinactive *** Cached, not pingable: $nrofcached *** Not Cached, but Pingable: $nrofping  *** Cached+Pingable: $nrofboth *** In error: $nroferror *** Script $scriptversion"
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


