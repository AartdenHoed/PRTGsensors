param (
    [string]$LOGGING = "YES",
    [int]$sensorid = 77   
)

$ScriptVersion = " -- Version: 3.5"

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
        Add-Content $logfile "==> Get IP/Mac address from this computer"
    }
    $ComputerName = $env:computername
    $OrgSettings = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $ComputerName -EA Stop | ? { $_.DNSDomain -eq "fritz.box" }
    $myip = $OrgSettings.IPAddress[0]

    $ip = (([ipaddress] $myip).GetAddressBytes()[0..2] -join ".") + "."
    $ip = $ip.TrimEnd(".")
}
catch {
    
    $scripterror = $true
    $errortext = $error[0]
    $scripterrormsg = "Getting IP/MAC address failed - $errortext"
    if ($log) {
        Add-Content $logfile "==> $scripterrormsg"
    }
    
}

#Searches the ARP table for IPs that match the scheme and parses out the data into a db table

if (!$scripterror) {

    try { 
        if ($log) {
            Add-Content $logfile "==> Truncate ARP table"
        }
        $query = "TRUNCATE TABLE dbo.ARP" 
        invoke-sqlcmd -ServerInstance ".\SQLEXPRESS" -Database "PRTG" `
                        -Query "$query" `
                        -ErrorAction Stop
        }
    catch {
        if ($log) {
            Add-Content $logfile "==> Truncate ARP table failed"
        }
        $scripterror = $true
        $errortext = $error[0]
        $scripterrormsg = "Truncate ARP database failed - $errortext"
    
    }
}

if (!$scripterror) { 
    try {
        if ($log) {
            Add-Content $logfile "==> Give ARP -A command and fill ARP table in PRTG database"
        }
        $arpa = (arp -a) 
        foreach ($line in $arpa) {
            # Write-Warning "Line: $line"
            $words =  $line.TrimStart() -split '\s+'
            $thisIP = $words[0].Trim()
            if ($thisIP -match $ip) {
                $thisMac = $words[1] 
                # Write-Warning "ThisMac: $thisMac"
                if (!(($thisMac -eq "---") -or ($thisMac -eq "Address") -or ($thisMac -eq $null) -or ($thisMac -eq "ff-ff-ff-ff-ff-ff") -or ($thisMac -eq "static"))) {
                    $thisMac = $thisMac.Replace("-",":")
                    $query = "INSERT INTO [dbo].[ARP] ([IPaddress],[MACaddress]) VALUES('" + 
                                    $thisip + "','" +
                                    $thisMac + "')"
                    invoke-sqlcmd -ServerInstance ".\SQLEXPRESS" -Database "PRTG" `
                            -Query "$query" `
                            -ErrorAction Stop
                }
               
            }
        }
    }
    catch {
        if ($log) {
            Add-Content $logfile "==> Filling ARP table failed"
        }
        $scripterror = $true
        $errortext = $error[0]
        $scripterrormsg = "Fill ARP database failed - $errortext"
    }
    
}

# Check the IP's and MAC adresseses via LEFT JOIN

if (!$scripterror) {
    if ($log) {
        Add-Content $logfile "==> Run SQL query (left join) to determine discrepancies"
    }
    $query = "SELECT [Naam]
          ,db.[IPaddress] as dbIPaddress
          ,arp.[IPaddress] as arpIPaddress
          ,db.[MACaddress] as dbMACaddress
	      ,arp.[MACaddress] as arpMACaddress
          ,db.AltMAC as dbAltMAC
          ,db.Authorized as Authorized
      FROM [PRTG].[dbo].[IPadressen] db      
      full outer join [PRTG].[dbo].[ARP] arp on db.IPaddress = arp.IPaddress 
      WHERE db.Authorized = 'Y'or db.Pingable = 'Y'or arp.IPaddress <> ''
      order by db.IPaddress"
    $joinresult = invoke-sqlcmd -ServerInstance ".\SQLEXPRESS" -Database "PRTG" `
                    -Query "$query" `
                    -ErrorAction Stop
}

$resultlist = @()
$somethingrotten = $false
$warning = $false


if (!$scripterror) {
    if ($log) {
        Add-Content $logfile "==> Determine status per IP"
    }

    foreach ($entry in $joinresult) {
        
        if ([string]::IsNullOrEmpty($entry.dbIPaddress)) {
             $unknownIP = $true
             $somethingrotten = $true
        }
        else { 
            if (!($entry.dbIPaddress.Trim() -match "^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$")) {
                # NO ip address. Just check for numerics. No check on valid range (0-255) needed
                # don't add it to list
                continue
            }
            $unknownIP = $false

            if ($entry.Authorized -ne "Y") {
                $IPstatus = "** Illegal **, "
                $somethingrotten = $true
                $illegal = $true
            }
            else {
                $IPstatus = ""
                $illegal = $false
            }
          
            if  ([string]::IsNullOrEmpty($entry.arpMACaddress)) {
                $IPstatus = $IPstatus + "Not cached"
                $wrongMAC = $false
                $altmac = $false
            }
            else {
                $IPstatus = $IPstatus + "Cached"
            }
            try {
                $pingable = $true
                $ping = Test-Connection -COmputerName $entry.dbIPaddress.Trim() -Count 1
            }
            catch {
                $pingable = $false    
            }
            finally {
                if ($pingable) {
                    $IPstatus = $IPstatus +", Pingable"
                }
                else {
                    $IPstatus = $IPstatus +", Not pingable"
                }
            }

            if ($IPstatus -eq "Cached") {
                if ($entry.dbMACaddress.ToUpper() -eq $entry.arpMACaddress.ToUpper()) { 
                    $wrongMAC = $false
                    $altMAC = $false
                }
                else {
                    if (![string]::IsNullOrEmpty($entry.dbAltMAC)) {
                        
                        if ($entry.arpMACaddress.ToUpper() -eq  $entry.dbAltMAC.ToUpper()) {
                            $wrongMAC = $false
                            $altMAC = $true
                            $warning = $true
                        }
                        # write-host "Bingo"
                  
                    }
                    else {
                        $wrongMAC = $true
                        $altMAC = $false
                        $somethingrotten = $true
                    }
                }
            }
            
        }
   
        $obj = [PSCustomObject] [ordered] @{Naam = $entry.Naam;
                                            dbIP = $entry.dbIPaddress; 
                                            arpIP = $entry.arpIPaddress; 
                                            dbMAC = $entry.dbMACaddress; 
                                            arpMAC = $entry.arpMACaddress; 
                                            unknownIP = $unknownIP; 
                                            wrongMAC = $wrongMAC; 
                                            altMAC = $altMAC;
                                            IPstatus = $IPstatus;
                                            Illegal = $Illegal}
        $resultlist += $obj
    
    } 
}


#$resultlist | Out-GridView

if ($log) {
    Add-Content $logfile "==> Create XML"
}

$total = 0
$cacheping = 0
$cachenoping = 0
$nocacheping = 0
$nocachenoping = 0
$nrofunknown = 0
$nrofwrongmac = 0
$nrofalternates = 0
$nrofillegal = 0


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

$Channel.InnerText = "Overall status"
$Unit.InnerText = "Custom"
$Mode.Innertext = "Absolute"
$ValueLookup.Innertext = 'OverallIPStatus'

if ($scripterror) {
    $Value.Innertext = "3"
} 
else { 
    if ($somethingrotten) {
        $Value.InnerText = "2"
    } 
    else {
        if ($warning) {
            $Value.Innertext = "1"
        }
        else {
            $Value.Innertext = "0"
        }
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
    $useIP = $item.dbIP
    if ([string]::IsNullOrEmpty($item.dbIP)) {
        $useIP = $item.arpIP
    }
    
    $total = $total + 1
    # Report each IP as Channel
    $Result = $xmldoc.CreateElement('Result')
    $Channel = $xmldoc.CreateElement('Channel')
    $Value = $xmldoc.CreateElement('Value')
    $Unit = $xmldoc.CreateElement('Unit')
    $CustomUnit = $xmldoc.CreateElement('CustomUnit')
    $Mode = $xmldoc.CreateElement('Mode')
    $NotifyChanged = $xmldoc.CreateElement('NotifyChanged')
    $ValueLookup =  $xmldoc.CreateElement('ValueLookup')

    $ipsplit = $useIP.Split(" .")
    $ipnr = (“{0:d3}” -f [int]$ipsplit[3].Trim()) 
    $cname = "IP" + $ipnr + " (" + $item.Naam.Trim() + ")"
    $Channel.InnerText = $cname
    $Unit.InnerText = "Custom"
    $Mode.Innertext = "Absolute"
    $ValueLookup.Innertext = 'IndividualIPStatus'

    switch ($item.IPstatus) {
        "Cached, Pingable" { 
            $thisval = 0
            $cacheping +=1
        }
        "Not cached, Pingable" {
            $thisval = 1
            $nocacheping +=1
        }
        "Cached, Not pingable" { 
            $thisval = 2
            $cachenoping += 1
        }
        "Not cached, Not pingable" {
            $thisval = 3
            $nocachenoping += 1
        }  
        "** Illegal **, Cached, Pingable" { 
            $thisval = 0
            $cacheping +=1
        }
        "** Illegal **, Not cached, Pingable" {
            $thisval = 1
            $nocacheping +=1
        }
        "** Illegal **, Cached, Not pingable" { 
            $thisval = 2
            $cachenoping += 1
        }
        "** Illegal **, Not cached, Not pingable" {
            $thisval = 3
            $nocachenoping += 1
        }        
      
        default {
            $thisval = 8 
            if ($log) {
                Add-Content $logfile "==> IP-status $IPstatus unknown"
            }  
        }
    }
    if ($item.AltMAC) {
        $thisval = 4
        $nrofalternates = $nrofalternates + 1
    }
    if ($item.unknownIP) {
        $thisval = 5
        $nrofunknown = $nrofunknown + 1
    } 
    if ($item.wrongMAC) {
        $thisval = 6
        $nrofwrongmac = $nrofwrongmac + 1
    } 
    if ($item.Illegal) {
        $thisval = 7
        $nrofillegal = $nrofillegal + 1
    } 
    $Value.Innertext = $thisval

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
    $message = "Total IP's: $Total *** Illegal: $nrofillegal *** Cached, Not Pingable: $cachenoping *** Cached, Pingable: $cacheping *** Not Cached, Pingable: $nocacheping *** Inactive: $nocachenoping *** Alternate MACs: $nrofalternates *** Unknown IP's: $nrofunknown *** Wrong MAC adresses: $nrofwrongmac *** Script Version: $scriptversion"
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


