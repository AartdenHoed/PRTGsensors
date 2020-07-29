param (
    [string]$LOGGING = "NO"    
)
# $LOGGING = 'YES'
CLS
if ($LOGGING -eq "YES") {$log = $true} else {$log = $false}
$logfile = "D:\AartenHetty\OneDrive\ArpA\Sensor.log"
if ($log) {
    $thisdate = Get-Date
    &{Write-Warning "==> START $thisdate"}  6>&1 5>&1  4>&1 3>&1 2>&1 > $logfile
}

$scriptversion = "1.6"
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



if ($log) {   
    &{Write-Warning "==> Script Version: $scriptversion"}  6>&1 5>&1  4>&1 3>&1 2>&1 >> $logfile
} 

try {
    if ($log) {
        &{Write-Warning "==> Get IP/Mac address from this computer"}  6>&1 5>&1  4>&1 3>&1 2>&1 >> $logfile
    }
    $ComputerName = $env:computername
    $OrgSettings = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $ComputerName -EA Stop | ? { $_.DNSDomain -eq "fritz.box" }
    $myip = $OrgSettings.IPAddress[0]

    $ip = (([ipaddress] $myip).GetAddressBytes()[0..2] -join ".") + "."
    $ip = $ip.TrimEnd(".")
}
catch {
    if ($log) {
        &{Write-Warning "==> ARP-A command failed"}  6>&1 5>&1  4>&1 3>&1 2>&1 >> $logfile
    }
    $scripterror = $true
    $errortext = $error[0]
    $scripterrormsg = "Getting IP/MAC address failed - $errortext"
    
}

#Searches the ARP table for IPs that match the scheme and parses out the data into a db table

if (!$scripterror) {

    try { 
        if ($log) {
        &{Write-Warning "==> Truncate ARP table"}  6>&1 5>&1  4>&1 3>&1 2>&1 >> $logfile
        }
        $query = "TRUNCATE TABLE dbo.ARP" 
        invoke-sqlcmd -ServerInstance ".\SQLEXPRESS" -Database "PRTG" `
                        -Query "$query" `
                        -ErrorAction Stop
        # add this computer to list (will not be in ARP -A output)
        
        if ($log) {
            &{Write-Warning "==> Insert this computer in ARP table"}  6>&1 5>&1  4>&1 3>&1 2>&1 >> $logfile
        }
        $a = Get-NetAdapter | ? {$_.Name -eq "Wi-Fi"}
        $mymac = $a.MacAddress.Replace("-",":")
        $query = "INSERT INTO [dbo].[ARP] ([IPaddress],[MACaddress]) VALUES('" + 
                    $myip + "','"+
                    $mymac + "')"
        invoke-sqlcmd -ServerInstance ".\SQLEXPRESS" -Database "PRTG" `
                        -Query "$query" `
                        -ErrorAction Stop
        }
    catch {
        if ($log) {
            &{Write-Warning "==> Truncate + initial entry in ARP table failed"}  6>&1 5>&1  4>&1 3>&1 2>&1 >> $logfile
        }
        $scripterror = $true
        $errortext = $error[0]
        $scripterrormsg = "Truncate ARP database failed - $errortext"
    
    }
}

if (!$scripterror) { 
    try {
        if ($log) {
            &{Write-Warning "==> Give ARP -A command and fill ARP table in PRTG database"}  6>&1 5>&1  4>&1 3>&1 2>&1 >> $logfile
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
                                    $thisip + "','"+
                                    $thismac + "')"
                    invoke-sqlcmd -ServerInstance ".\SQLEXPRESS" -Database "PRTG" `
                            -Query "$query" `
                            -ErrorAction Stop
                }
               
            }
        }
    }
    catch {
        if ($log) {
            &{Write-Warning "==> Filling ARP table failed"}  6>&1 5>&1  4>&1 3>&1 2>&1 >> $logfile
        }
        $scripterror = $true
        $errortext = $error[0]
        $scripterrormsg = "Fill ARP database failed - $errortext"
    }
    
}

# Check the IP's and MAC adresseses via LEFT JOIN

if (!$scripterror) {
    if ($log) {
        &{Write-Warning "==> Run SQL query (left join) to determine discrepancies"}  6>&1 5>&1  4>&1 3>&1 2>&1 >> $logfile
    }
    $query = "SELECT [Naam]
          ,db.[IPaddress] as dbIPaddress
          ,arp.[IPaddress] as arpIPaddress
          ,db.[MACaddress] as dbMACaddress
	      ,arp.[MACaddress] as arpMACaddress
      FROM [PRTG].[dbo].[IPadressen] db
      full outer join [PRTG].[dbo].[ARP] arp on db.IPaddress = arp.IPaddress order by db.IPaddress"
    $joinresult = invoke-sqlcmd -ServerInstance ".\SQLEXPRESS" -Database "PRTG" `
                    -Query "$query" `
                    -ErrorAction Stop
}

$resultlist = @()
$somethingrotten = $false

if (!$scripterror) {
    if ($log) {
        &{Write-Warning "==> Determine status per IP"}  6>&1 5>&1  4>&1 3>&1 2>&1 >> $logfile
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
          
            if  ([string]::IsNullOrEmpty($entry.arpMACaddress)) {
                $IPstatus = "Inactive"
                $wrongMAC = $false
            }
            else {
                $IPstatus = "Active"
                if ($entry.dbMACaddress.ToUpper() -eq $entry.arpMACaddress.ToUpper()) { 
                    $wrongMAC = $false
                }
                else {
                    $wrongMAC = $true
                    $somethingrotten = $true
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
                                            IPstatus = $IPstatus}
        $resultlist += $obj
    
    } 
}


#$resultlist | Out-GridView

if ($log) {
    &{Write-Warning "==> Create XML"}  6>&1 5>&1  4>&1 3>&1 2>&1 >> $logfile
}

$total = 0
$nrofactive = 0
$nrofinactive = 0
$nrofunknown = 0
$nrofwrongmac = 0


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
        $Value.Innertext = "1"
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
    $cname = "IP" + $ipnr
    $Channel.InnerText = $cname
    $Unit.InnerText = "Custom"
    $Mode.Innertext = "Absolute"
    $ValueLookup.Innertext = 'IndividualIPStatus'

    if ($item.IPstatus -eq "Active") { 
        $thisval = 0
        $nrofactive = $nrofactive + 1
    }
    else {
        $thisval = 1
        $nrofinactive = $nrofinactive + 1
    }
    if ($item.unknownIP) {
        $thisval = 3
        $nrofunknown = $nrofunknown + 1
    } 
    if ($item.wrongMAC) {
        $thisval = 4
        $nrofwrongmac = $nrofwrongmac + 1
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
    $message = "Total IP's: $Total *** Active: $nrofactive *** Inactive: $nrofinactive *** Unknown IP's: $nrofunknown *** Wrong MAC adresses: $nrofwrongmac *** Script Version: $scriptversion"
    $ErrorText.InnerText = $message
} 
[void]$PRTG.AppendChild($ErrorValue)
[void]$PRTG.AppendChild($ErrorText)
    
[void]$xmldoc.Appendchild($PRTG)

if ($log) {
    &{Write-Warning "==> Write XML"}  6>&1 5>&1  4>&1 3>&1 2>&1 >> $logfile
}

WriteXmlToScreen $xmldoc

if ($log) {
    $thisdate = Get-Date
    &{Write-Warning "==> END $thisdate"}  6>&1 5>&1  4>&1 3>&1 2>&1 >> $logfile
}


