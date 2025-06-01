$Version = " -- Version: 1.5"

# COMMON coding
CLS

# init flags
$StatusOBJ = [PSCustomObject] [ordered] @{Scripterror = $false;
                                          ScriptChange = $false;
                                          ScriptAction = $false;
                                          RecordsLogged = $false
                                          }

$InformationPreference = "Continue"
$WarningPreference = "Continue"
$ErrorActionPreference = "Stop"

# ------------------ FUNCTIONS
function Report ([string]$level, [string]$line, [object]$Obj, [string]$file ) {
    switch ($level) {
        ("N") {$rptline = $line}
        ("H") {
            $rptline = "-------->".Padright(10," ") + $line
        }
        ("I") {
            $rptline = "Info    *".Padright(10," ") + $line
        }
        ("A") {
            $rptline = "Caution *".Padright(10," ") + $line
        }
        ("B") {
            $rptline = "        *".Padright(10," ") + $line
        }
        ("C") {
            $rptline = "Change  *".Padright(10," ") + $line
            $obj.scriptchange = $true
        }
        ("W") {
            $rptline = "Warning *".Padright(10," ") + $line
            $obj.scriptaction = $true
        }
        ("E") {
            $rptline = "Error   *".Padright(10," ") + $line
            $obj.scripterror = $true
        }
        ("G") {
            $rptline = "GIT:    *".Padright(10," ") + $line
        }
        default {
            $rptline = "Error   *".Padright(10," ") + "Messagelevel $level is not valid"
            $Obj.Scripterror = $true
        }
    }
    Add-Content $file $rptline

}


# ------------------------ END OF FUNCTIONS

# ------------------------ START OF MAIN CODE


$Node = " -- Node: " + $env:COMPUTERNAME

$myname = $MyInvocation.MyCommand.Name
$enqprocess = $myname.ToUpper().Replace(".PS1","")
$FullScriptName = $MyInvocation.MyCommand.Definition
$mypath = $FullScriptName.Replace($MyName, "")

$LocalInitVar = $mypath + "InitVar.PS1"
$InitObj = & "$LocalInitVar" "OBJECT"

if ($Initobj.AbEnd) {
    # Write-Warning "YES"
    throw "INIT script $LocalInitVar Failed"

}  
   
# END OF COMMON CODING

$gp = Get-Process -id $pid 
$ProcessID = $gp.Id
$ProcessName = $gp.Name      

# Init reporting file
$dir = $ADHC_TempDirectory + $ADHC_LocalCpuTemperature.Directory
New-Item -ItemType Directory -Force -Path $dir | Out-Null
$tempfile = $dir + $ADHC_LocalCpuTemperature.Name

# Init jobstatus file
$dir = $ADHC_OutputDirectory + $ADHC_Jobstatus
New-Item -ItemType Directory -Force -Path $dir | Out-Null
$p = $myname.Split(".")
$process = $p[0]
$jobstatus = $ADHC_OutputDirectory + $ADHC_Jobstatus + $ADHC_Computer + "_" + $Process + ".jst" 

$errorcount = 0
$loop = 0
$myname = $MyInvocation.MyCommand.Name
    
$FullScriptName = $MyInvocation.MyCommand.Definition
$mypath = $FullScriptName.Replace($MyName, "")

$sensorscript = $mypath + "CpuTemperature.ps1"

do {
    $d = Get-Date
    $Datum = " -- Date: " + $d.ToString("dd-MM-yyyy")
    $Tijd = " -- Time: " + $d.ToString("HH:mm:ss")
    $loop = $loop + 1
    $Scriptmsg = "*** STARTED *** " + $mypath + " -- PowerShell script " + $MyName + $Version + $Datum + $Tijd +$Node
    Write-Information $Scriptmsg 
    Set-Content $Tempfile $Scriptmsg -force
    if ($loop -le 1) {
        foreach ($entry in $InitObj.MessageList){
            Report $entry.Level $entry.Message $StatusObj $Tempfile
        }
    }

    Report "I" "Process name = $ProcessName, proces ID = $ProcessID" $StatusObj $Tempfile

    Report "I"  "Iteration number $loop ($errorcount errors until now)" $StatusObj $Tempfile
        
    try {

        Report "I" "Started... sensor script = $sensorscript" $StatusObj $Tempfile
        $Sensor = & "$Sensorscript" YES $ADHC_Computer 9000
        Report "I" "Ended... sensor script = $sensorscript" $StatusObj $Tempfile

    }
    catch {
        Report "E" "Error !!!" $StatusObj $Tempfile
        $errorcount += 1
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
        $Dump = $_.Exception.ToSTring()
        Report "E" "Message = $ErrorMessage" $StatusObj $Tempfile
        Report "E" "Failed Item = $Faileditem" $StatusObj $Tempfile
        Report "E" "Dump = $Dump" $StatusObj $Tempfile
    } 
    finally {
       
        if  ($StatusObj.scripterror) {
            $dt = Get-Date
            $jobline = $ADHC_Computer + "|" + $process + "|" + "9" + "|" + $version + "|" + $dt.ToString("dd-MM-yyyy HH:mm:ss")
            Set-Content $jobstatus $jobline
       
            Add-Content $jobstatus "Failed item = $FailedItem"
            Add-Content $jobstatus "Errormessage = $ErrorMessage"
            Add-Content $jobstatus "Dump info = $dump"

            Report "E" "Failed item = $FailedItem" $StatusObj $Tempfile
            Report "E" "Errormessage = $ErrorMessage" $StatusObj $Tempfile
            Report "E" "Dump info = $dump" $StatusObj $Tempfile
            }
        else {
            Report "I" ">>> Iteration ended normally $Datum $Tijd" $StatusObj $Tempfile
            Report "N" " " $StatusObj $Tempfile
   
            $dt = Get-Date
            $jobline = $ADHC_Computer + "|" + $process + "|" + "0" + "|" + $version + "|" + $dt.ToString("dd-MM-yyyy HH:mm:ss")
            Set-Content $jobstatus $jobline

        }
        Report "I" "Wait 300 seconds..." $StatusObj $Tempfile
        Report "N" " " $StatusObj $Tempfile
        try { #  copy temp file
        
            $deffile = $ADHC_OutputDirectory + $ADHC_LocalCpuTemperature.Directory + $ADHC_LocalCpuTemperature.Name 
            $CopMov = & $ADHC_CopyMoveScript $TempFile $deffile "COPY" "REPLACE" $TempFile  
            
        }
        Catch {
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            $Dump = $_.Exception.ToSTring()
            $dt = Get-Date
            $jobline = $ADHC_Computer + "|" + $process + "|" + "9" + "|" + $version + "|" + $dt.ToString("dd-MM-yyyy HH:mm:ss")
            Set-Content $jobstatus $jobline
            Add-Content $jobstatus "Failed item = $FailedItem"
            Add-Content $jobstatus "Errormessage = $ErrorMessage"
            Add-Content $jobstatus "Dump info = $Dump"
            $errorcount += 1    

        }    
        
        
        Start-Sleep -s 300
    }

} Until ($errorcount -gt 10)

Report "I" "Ended with error count = $errorcount" $StatusObj $Tempfile