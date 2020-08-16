param (
    [string]$LOGGING = "NO"    
)
# $LOGGING = 'YES'

$ScriptVersion = " -- Version: 1.1.1"

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
        Add-Content $logfile "==> Get list of jobstatus files"
    }
    $logdir = $ADHC_OutputDirectory + $ADHC_JobStatus
    $logList = Get-ChildItem $logdir -File | Select Name,FullName
   
}
catch {
    if ($log) {
        Add-Content $logfile "==> Reading directoy $dir failed"
    }
    $scripterror = $true
    $errortext = $error[0]
    $scripterrormsg = "Reading directoy $dir failed - $errortext"
    
}



# Process all jobstatus files

if (!$scripterror) {

    try { 
        if ($log) {
            Add-Content $logfile "==> Interprete each jobstatus file"
        }
        $resultlist = @()
        $total = 0
        $Nrofnono = 0
        $NrofChangeNoact = 0
        $NrofAction = 0
        $NrofError = 0
        $MaxCode = 0

        foreach ($logdataset in $loglist) {
            $a = Get-Content $logdataset.FullName
            $args = $a.Split("|")
            $Machine = $args[0]
            $Job = $args[1]
            $Jobstatus = $args[2]
            $obj = [PSCustomObject] [ordered] @{Machine = $Machine;
                                            Job = $Job; 
                                            Jobstatus = $Jobstatus}
            $resultlist += $obj 
            $Total = $Total + 1;
            $Maxcode = [math]::Max($Maxcode, $Jobstatus)
            switch ($jobstatus) {
                "0" { $Nrofnono = $Nrofnono + 1}
                "3" { $NrofChangeNoact = $NrofChangeNoact + 1}
                "6" { $NrofAction = $NrofAction + 1}
                "9" { $NrofError = $NrofError + 1}
                default {
                    if ($log) {
                        Add-Content $logfile "==> Invalid jobstatus $jobstatus"
                    }
                    $scripterror = $true
                    $errortext = $error[0]
                    $scripterrormsg = "Invalid jobstatus $jobstatus"
                }
            }
            
            
        }
        $resultlist | Out-gridview

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

# Overall status
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
$NotifyChanged = $xmldoc.CreateElement('NotifyChanged')
$ValueLookup =  $xmldoc.CreateElement('ValueLookup')

$Channel.InnerText = "Overall JOB status"
$Unit.InnerText = "Custom"
$Mode.Innertext = "Absolute"
$ValueLookup.Innertext = 'OverallJOBStatus'

if ($scripterror) {
    $Value.Innertext = "12"
} 
else { 
   $Value.Innertext = $Maxcode
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
    $ValueLookup.Innertext = 'IndividualJOBStatus'

    $Value.Innertext = $item.Jobstatus

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
    $message = "Total jobs: $Total *** No Change: $Nrofnono *** Change, no action: $NrofChangeNoact *** Action required: $NrofAction *** Jobs in error: $Nroferror *** Script Version: $scriptversion"
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


