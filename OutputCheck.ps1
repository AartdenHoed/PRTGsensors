param (
    [string]$LOGGING = "NO", 
    [int]$sensorid = 77 
)
# $LOGGING = 'YES'

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

# If not exsitent, INIT outputcheck file. 

try {

    $OutputCheckFile = $ADHC_OutputDirectory + $ADHC_OutputCheck
    $dr = Test-Path $OutputCheckFile
    if (!$dr) {
        $str = $ADHC_OutputCheck.Split("\")
        $dir = $ADHC_OutputDirectory + $str[0] + "\" + $str[1]
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
    
        $initlist = Get-ChildItem -Path $ADHC_OutputDirectory -Recurse -File -Force  | Select FullName
        $first = $true 
        foreach ($ds in $initlist) {
            $rec = "U|" + $ds.Fullname + "|"
            if ($first) {
                Set-Content $OutputCheckFile $rec
                $first = $false
            } 
            else {
                Add-Content $OutputCheckFile $rec
            }    

        }
    }
    # Read the file into list
    $ifile = Get-Content $OutputCheckFile
    $ilist = @()
    foreach ($r in $ifile) {
        $ispl = $r.split("|")
        $state = $ispl[0]
        $dsname = $ispl[1]
        $obj0 = [PSCustomObject] [ordered] @{State = $state;
                                             FullName = $dsname}
        $ilist += $obj0

    }
}
catch {
     if ($log) {
            Add-Content $logfile "==> Creating/reading info file failed"          
        }
        $scripterror = $true
        $errortext = $error[0]
        $scripterrormsg = "==> Creating/reading info file failed"

}


# Get BOOT info
$bootlist = @()
if (!$scripterror) {
    if ($log) {
        Add-Content $logfile "==> Read all bootfiles and validate them"
    }
    try {
        $b = $ADHC_OutputDirectory + $ADHC_Bootdir
        $bootfiles = Get-ChildItem -Path $b -Recurse -File -Force | Select Name, FullName
        foreach ($bootfile in $bootfiles) {
            $n = $bootfile.Name
            $spl = $bootfile.Name -split "_"
            $hostname = $spl[0].ToUpper()

            if ($ADHC_Hoststring.ToUpper().Contains($hostname)) {
                # Write-Host "$Hostname found"
                $illegalbootfile = $false

                $bootline = Get-Content $Bootfile.FullName 
                $bootsplit = $bootline.Split("|")
                $Bootstart = [datetime]::ParseExact($bootsplit[1],"dd-MM-yyyy HH:mm:ss",$null)
                $Bootstop = [datetime]::ParseExact($bootsplit[2],"dd-MM-yyyy HH:mm:ss",$null)
                $diff = NEW-TIMESPAN –Start $Bootstart –End $Bootstop
                $Uptime = $diff.TotalMinutes
        
                $obj2 = [PSCustomObject] [ordered] @{BootFile = $n;
                                                    Host = $hostname;
                                                    Bootstart = $Bootstart
                                                    Bootstop = $Bootstop
                                                    Uptime = $Uptime}
                $bootlist += $obj2
            } 
            else {
                # Write-Host "$Hostname not found"
                $illegalbootfile = $true
            }
        
        } 
    }
    catch {
        if ($log) {
            Add-Content $logfile "==> Reading bootfiles failed"          
        }
        $scripterror = $true
        $errortext = $error[0]
        $scripterrormsg = "==> Reading bootfiles failed - $errortext"
    }
}

$dirobj = @()

# Get list of directories
if (!$scripterror) {
    if ($log) {
        Add-Content $logfile "==> Create a list of directories"
    }
    try {
        $dirlist = Get-ChildItem -Path $ADHC_OutputDirectory -Recurse -Directory -Force  | Select FullName
        foreach ($entry in $dirlist) {
        
            $f = $entry.FullName
            
            $obj1 = [PSCustomObject] [ordered] @{Path = $f}
            $dirobj += $obj1

        }  
    }
    catch {
        if ($log) {
            Add-Content $logfile "==> Creating directory list for $ADHC_OutputDirectory failed"          
        }
        $scripterror = $true
        $errortext = $error[0]
        $scripterrormsg = "==> Creating directory list for $ADHC_OutputDirectory failed - $errortext"
    }

}

# Process directories en check for OLD files
$xmlobj = @()
if (!$scripterror) {
    if ($log) {
        Add-Content $logfile "==> Check the files in each directory"
    }
    try {

        $totOLD = 0
        $totCurrent = 0
        $totNoCheck = 0
        $totaal = 0
        $olist = @()

        foreach ($obj1 in $dirobj) {
            $dslist = Get-ChildItem -Path $obj1.Path -File -Force | Select Name, FullName, DirectoryName, LastWriteTime

            $myOLD = 0
            $myCurrent = 0
            $myNoCheck = 0

            $d = $obj1.Path.Split("\")
            $i = $d.Count - 1
            $directory = $d[$i]
               
            foreach ($file in $dslist) {
                $totaal += 1
                $Lastupdate = $file.LastWriteTime
                           
                $fname = $file.FullName
                $dsname = $file.Name
                $j = $dsname.Split("_")
                $dshost = $j[0].ToUpper()
                $dsstate = "?"
                if (!($ADHC_Hoststring.ToUpper().Contains($dshost)) ) {

                    switch ($dshost.ToUpper()) {
                        "WMIC"      {
                            $dshost = "<nocheck>"
                        }
                        "#OVERALL" {
                            $dshost = "<nocheck>"
                        }
                        "Analysis" {
                            $dshost = "<nocheck>"
                        }
                        "CURRENT"  {
                            $dshost = "<nocheck>"
                        }
                        default     {
                            if ($directory.ToUpper() -eq "SECURITY") {
                                $dshost = "<nocheck>"
                            }
                            else {
                                $dshost = "ADHC-2"
                            }


                        }

                    }

                }

                if ($dshost -eq "<nocheck>") {
                   
                    $dsstate = "U"
                
                } 
                else {
                    $thishost = $bootlist | Where-Object Host -eq $dshost

                    if ($Lastupdate -ge $thishost.Bootstart) {                        
                        $dsstate = "C"
                    }
                    else {
                        
                        if ($thishost.Uptime -le 120) {
                            $thisds = $ilist | Where-Object Fullname -eq $fname   
                            $dsstate = $thisds.State                                                                         
                        }                        
                        else{                            
                            $dsstate = "O"
                        }
                    }

                }
                switch ($dsstate) {
                    "C" {
                        $myCurrent += 1
                        $totCurrent += 1
                    }
                    "O" {
                        $myOld += 1
                        $totOld += 1
                    }
                    "U" {
                         $myNocheck += 1
                         $totNoCheck += 1
                    }
                    default {
                        if ($log) {
                           Add-Content $logfile "==> Wrong dataset status encountered: $dsstate"          
                        }
                        $scripterror = $true
                        $scripterrormsg = "==> Wrong dataset status encountered: $dsstate for dataset $fname"
                    }

                }
                $obj0 = [PSCustomObject] [ordered] @{State = $dsstate;
                                             FullName = $fname}
                $olist += $obj0
                                       
            }

            # Write list to info dataset
            $first = $true 
            foreach ($o in $olist) {
                $rec = $o.State + "|" + $o.Fullname + "|"
                if ($first) {
                    Set-Content $OutputCheckFile $rec
                    $first = $false
                } 
                else {
                    Add-Content $OutputCheckFile $rec
                }    

            }
                   

            $obj3 = [PSCustomObject] [ordered] @{Directory = $directory;
                                                Currentfiles = $myCurrent;
                                                OldFiles = $myOld;
                                                NotCheckedFiles = $myNoCheck }
            $xmlobj += $obj3
             
        }       

    }
    catch {
        if ($log) {
            Add-Content $logfile "==> Checking file timestamps failed"          
        }
        $scripterror = $true
        $errortext = $error[0]
        $scripterrormsg = "==> Checking file timestamps failed - $errortext"

    }
 
}

if ($log) {
    Add-Content $logfile "==> Create XML"
}

[xml]$xmldoc = New-Object System.Xml.XmlDocument
$decl = $xmldoc.CreateXmlDeclaration('1.0','Windows-1252',$null)

[void]$xmldoc.AppendChild($decl)

$PRTG = $xmldoc.CreateElement('PRTG')

# Overall file status

$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
$NotifyChanged = $xmldoc.CreateElement('NotifyChanged')
$ValueLookup =  $xmldoc.CreateElement('ValueLookup')

$Channel.InnerText = "Overall File Status"
$Unit.InnerText = "Custom"
$Mode.Innertext = "Absolute"
$ValueLookup.Innertext = 'OverallFileStatus'

if ($scripterror) {
    $Value.Innertext = "12"
} 
else { 
    if ($illegalbootfile) {
        $Value.Innertext = "8"
    }
    else {
        if ($totOld -eq 0) {
            if ($totNoCheck -eq 0) {
            $Value.Innertext = "0"
            }
            else {
                $Value.Innertext = "1"
            }     
        }
        else {
            $Value.Innertext = "4"
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

# Node status
foreach ($xml in $xmlobj) {
    $Result = $xmldoc.CreateElement('Result')
    $Channel = $xmldoc.CreateElement('Channel')
    $Value = $xmldoc.CreateElement('Value')
    $Unit = $xmldoc.CreateElement('Unit')
    $CustomUnit = $xmldoc.CreateElement('CustomUnit')
    $Mode = $xmldoc.CreateElement('Mode')
    $NotifyChanged = $xmldoc.CreateElement('NotifyChanged')
    $ValueLookup =  $xmldoc.CreateElement('ValueLookup')

    $Channel.InnerText = $xml.Directory
    $Unit.InnerText = "Custom"
    $Mode.Innertext = "Absolute"
    $ValueLookup.Innertext = 'IndividualFileStatus'

    if ($xml.Oldfiles -eq 0) {
        if ($xml.NotCheckedFiles -eq 0) {
            $Value.Innertext = "0"
        }
        else {
            $Value.Innertext = "1"
        }     
    }
    else {
        $Value.Innertext = "4"
    }
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
    $O = $totOLD 
    $C = $totCurrent 
    $N = $totNoCheck 
    $T = $totaal 
    
    $message = "Datasets total = $T, current = $C, not checked = $N, old = $O *** Script $scriptversion"
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
