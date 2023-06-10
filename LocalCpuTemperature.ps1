$Version = " -- Version: 1.3.2"

# COMMON coding
CLS

# init flags
$global:scripterror = $false
$global:scriptaction = $false
$global:scriptchange = $false

$global:recordslogged = $false

$InformationPreference = "Continue"
$WarningPreference = "Continue"
$ErrorActionPreference = "Stop"

# ------------------ FUNCTIONS
function Report ([string]$level, [string]$line) {
    switch ($level) {
        ("N") {$rptline = $line}
        ("I") {
            $rptline = "Info    *".Padright(10," ") + $line
        }
        ("H") {
            $rptline = "-------->".Padright(10," ") + $line
        }
        ("A") {
            $rptline = "Caution *".Padright(10," ") + $line
        }
        ("B") {
            $rptline = "        *".Padright(10," ") + $line
        }
        ("C") {
            $rptline = "Change  *".Padright(10," ") + $line
            $global:scriptchange = $true
        }
        ("W") {
            $rptline = "Warning *".Padright(10," ") + $line
            $global:scriptaction = $true
        }
        ("E") {
            $rptline = "Error   *".Padright(10," ") + $line
            $global:scripterror = $true
            
        }
        default {
            $rptline = "Error   *".Padright(10," ") + "Messagelevel $level is not valid"
            $global:scripterror = $true
        }
    }
    Add-Content $tempfile $rptline

}

# ------------------------ END OF FUNCTIONS

# ------------------------ START OF MAIN CODE


$Node = " -- Node: " + $env:COMPUTERNAME

$myname = $MyInvocation.MyCommand.Name
$enqprocess = $myname.ToUpper().Replace(".PS1","")
$FullScriptName = $MyInvocation.MyCommand.Definition
$mypath = $FullScriptName.Replace($MyName, "")

$LocalInitVar = $mypath + "InitVar.PS1"
& "$LocalInitVar"

if (!$ADHC_InitSuccessfull) {
    # Write-Warning "YES"
    throw $ADHC_InitError
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

Report "I" "Started... sensor script = $sensorscript"

do {
    $d = Get-Date
    $Datum = " -- Date: " + $d.ToString("dd-MM-yyyy")
    $Tijd = " -- Time: " + $d.ToString("HH:mm:ss")
    $loop = $loop + 1
    $Scriptmsg = "*** STARTED *** " + $mypath + " -- PowerShell script " + $MyName + $Version + $Datum + $Tijd +$Node
    Write-Information $Scriptmsg 
    Set-Content $Tempfile $Scriptmsg -force


    Report "I" "Process name = $ProcessName, proces ID = $ProcessID"

    Report "I"  "Iteration number $loop"
        
    try {
        
        & "$Sensorscript" YES $ADHC_Computer 9000

    }
    catch {
        Report "E" "Error !!!"
        $errorcount += 1
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
        $Dump = $_.Exception.ToSTring()
        Report "E" "Message = $ErrorMessage"
        Report "E" "Failed Item = $Faileditem"
        Report "E" "Dump = $Dump"
    } 
    finally {
       
        if  ($global:scripterror) {
             $dt = Get-Date
            $jobline = $ADHC_Computer + "|" + $process + "|" + "9" + "|" + $version + "|" + $dt.ToString("dd-MM-yyyy HH:mm:ss")
            Set-Content $jobstatus $jobline
       
            Add-Content $jobstatus "Failed item = $FailedItem"
            Add-Content $jobstatus "Errormessage = $ErrorMessage"
            Add-Content $jobstatus "Dump info = $dump"

            Report "E" "Failed item = $FailedItem"
            Report "E" "Errormessage = $ErrorMessage"
            Report "E" "Dump info = $dump"
            }
        else {
            Report "I" ">>> Script (iteration) ended normally $Datum $Tijd"
            Report "N" " "
   
            $dt = Get-Date
            $jobline = $ADHC_Computer + "|" + $process + "|" + "0" + "|" + $version + "|" + $dt.ToString("dd-MM-yyyy HH:mm:ss")
            Set-Content $jobstatus $jobline

        }
        Report "I" "Wait 300 seconds..."
        Report "N" " "
        try { #  copy temp file
        
            $deffile = $ADHC_OutputDirectory + $ADHC_LocalCpuTemperature.Directory + $ADHC_LocalCpuTemperature.Name 
            if ($loop -eq 1) {
                & $ADHC_CopyMoveScript $TempFile $deffile "MOVE" "REPLACE" $TempFile  
            }
            else {
                & $ADHC_CopyMoveScript $TempFile $deffile "MOVE" "APPEND" $TempFile 
            }
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

Report "I" "Ended with error count = $errorcount"