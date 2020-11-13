param (
    [string]$LOGGING = "NO", 
    [string]$Ctype = "Iets"
)
#$LOGGING = 'YES'
#$Ctype = "Guest"
$Ctype = $Ctype.ToUpper()

$ScriptVersion = " -- Version: 1.1"

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
& "$LocalInitVar"

$scripterror = $false

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
    $logfile = $dir + $process + ".log" 

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

# Determin IP range

try {
    if ($log) {
        Add-Content $logfile "==> Get IP address range"
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

if (!$scripterror) { 
    try {
        if ($log) {
            Add-Content $logfile "==> Give ARP -A command to get all active IP-Addresses"
        }
        $iplist = @()
        $arpa = (arp -a) 
        foreach ($line in $arpa) {
            # Write-Warning "Line: $line"
            $words =  $line.TrimStart() -split '\s+'
            $thisIP = $words[0].Trim()
            if ($thisIP -match $ip) {
                $thisMac = $words[1] 
                # Write-Warning "ThisMac: $thisMac"
                if (!(($thisMac -eq "---") -or ($thisMac -eq "Address") -or ($thisMac -eq $null) -or ($thisMac -eq "ff-ff-ff-ff-ff-ff") -or ($thisMac -eq "static"))) {
                    $iplist += $thisip   
                }               
            }
        }
    }
    catch {        
        $scripterror = $true
        $errortext = $error[0]
        $scripterrormsg = "Getting active IP addresses failed - $errortext"
        if ($log) {
            Add-Content $logfile "==> $scripterrormsg"
        }
    }
    
}
$maxcode = 0
$nrofping = 0
$nrofactive = 0
$nrofinactive = 0
$nroftotal = 0
$activeIPfound = $false
$pingsuccess = $false

if (!$scripterror) {
    try {
        if ($log) {
            Add-Content $logfile "==> Check status of every IP with type $ctype"          
        }
        $ResultList = @()
        # Check al entries from database against list of active IP-addresses
        foreach ( $dbentry in $ctypelist) {
            $nroftotal = $nroftotal + 1
            if ($iplist -contains $dbentry.IPaddress) {
                $ipactive = $true
                $activeIPfound = $true
                $maxcode = 3
                $ipping = $false
                try {
                    $ping = Test-Connection -COmputerName $dbentry.IPaddress -Count 1
                    $ipping = $true
                    $pingsuccess = $true
                    $maxcode = 6
                }
                catch {
                    $ipping = $false
                }
            }
            else {
                $ipactive = $false
                $ipping = $false
            }
            if ($ipping) { $nrofping = $nrofping + 1} 
            if ($ipactive) {
                $nrofactive = $nrofactive + 1 
            } 
            else {
                $nrofinactive = $nrofinactive + 1
            } 
            
            
            $ipobj = [PSCustomObject] [ordered] @{IPname = $dbentry.Naam;
                                            IPaddress = $dbentry.IPaddress;
                                            IPactive = $ipactive; 
                                            IPtype = $ctype;
                                            IPping = $ipping;
                                            }
            $ResultList += $ipobj                
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

    $ipcode = 0
    if ($item.IPactive) {$ipcode = "3"} 
    if ($item.IPping) { $ipcode = "6"}
    
    $Value.Innertext = $ipcode

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
   
    $message = "Nodes in group $ctype *** Total: $nroftotal *** Inactive: $nrofinactive *** Active: $nrofactive *** Pingable: $nrofping  *** Script $scriptversion"
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


