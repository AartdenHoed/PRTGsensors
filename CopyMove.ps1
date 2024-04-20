param (
     [string]$InputFile = "In"  ,
     [string]$OutputFile = "Out" ,
     [string]$Action = "Move/Copy" ,
     [string]$Mode = "Replace/Append" ,
     [string]$Messages = "MSG",
     [string]$Lock = ""
)
#$InputFile = "D:/ADHC_Home/ADHC_Temp/WmicFiles/WMIC_ADHC.log"
#$Action = "move"
#$Mode = "Append"
#$OutputFile = "D:/ADHC_Home/OneDrive/ADHC Output/WmicFiles/WMIC_ADHC.log"  
#$Messages =   JSON   - return JSON
#              MSG    - write-host messages
#              SILENT - suppress messages
#              OBJECT - return messages in object
#              <filename> - write messages to <filename> 

#$Lock ="Resource,Process"
#$messages = "OBJECT"

$Action = $Action.ToUpper()
$Mode = $Mode.ToUpper()
$Messages = $Messages.ToUpper()
$Actionlist = @("MOVE","COPY")
$Modelist = @("REPLACE","APPEND")

$ScriptVersion = " -- Version: 1.4.1"
$ReturnOBJ = [PSCustomObject] [ordered] @{AbEnd = $false;
                                                  MessageList = @()
                                                 }

function Report ([string]$level, [string]$line, [object] $Obj, [string] $target) {
    
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
            
        }
        ("W") {
            $rptline = "Warning *".Padright(10," ") + $line
            
        }
        ("E") {
            $rptline = "Error   *".Padright(10," ") + $line
            $Obj.AbEnd = $true
        }
        default {
            $rptline = "Error   *".Padright(10," ") + "Messagelevel $level is not valid"
            $Obj.AbEnd = $true
        }
    }
    
    
    switch ($Target) {
        "MSG" {
            Switch ($Level) {
                "E" {Write-Error $rptline}
                "W" {Write-Warning $rptline}
                "A" {Write-Warning $rptline}
                default {Write-Information $rptline}
            }
        }
       
        "JSON" {
            $msgentry = [PSCustomObject] [ordered] @{Message = $rptline;
                                                     Level = "I"}
            $Obj.MessageList += $msgentry    
        }
        "SILENT" { }
        "OBJECT" { 
            
            $msgentry = [PSCustomObject] [ordered] @{Message = $rptline;
                                                     Level = "I"}
            $Obj.MessageList += $msgentry
        }
        Default  { 
            if (!(Test-Path $Messages)) {
                Throw "Messages file $messages does not exist"
            }
            Add-Content $Messages $rptline
        }    
    }
}


# COMMON coding
Try {
    $InformationPreference = "Continue"
    $WarningPreference = "Continue"
    $ErrorActionPreference = "Stop"

    $Node = " -- Node: " + $env:COMPUTERNAME
    $d = Get-Date
    $Datum = " -- Date: " + $d.ToString("dd-MM-yyyy")
    $Tijd = " -- Time: " + $d.ToString("HH:mm:ss")

    $myname = $MyInvocation.MyCommand.Name
    $p = $myname.Split(".")
    $process = $p[0]
    $FullScriptName = $MyInvocation.MyCommand.Definition
    $mypath = $FullScriptName.Replace($MyName, "")

    $Scriptmsg = "*** STARTED *** " + $mypath + " -- PowerShell script " + $MyName + $ScriptVersion + $Datum + $Tijd +$Node
    Report "N" $Scriptmsg $ReturnOBJ $Messages

    $LocalInitVar = $mypath + "InitVar.PS1"
    $InitObj = & "$LocalInitVar" "OBJECT"

    if ($Initobj.AbEnd) {
        # Write-Warning "YES"
        throw "INIT script $LocalInitVar Failed"

    }
    foreach ($entry in $InitObj.MessageList){
        Report $entry.Level $entry.Message $ReturnOBJ $Messages
    }


}
Catch {
    
    $errortext = $error[0]
    $scripterrormsg = "Input validation failed - $errortext"
    Report "E" "$scripterrormsg" $ReturnOBJ $Messages

}

# Determine input parameters and verify input


if (!$ReturnObj.AbEnd) {
    try {    
        $mymsg = "Input Validation: $Action file $InputFile to $Outputfile (mode: $Mode)"
        Report "I"  $mymsg $ReturnOBJ $Messages
        if ($lock) {
            Report "I" "... With lock parameters $lock" $ReturnOBJ $Messages
        }
        switch ($Messages) {
           "JSON"   {$mymsg = "Messages will be returned via JSON"}
           "MSG"    {$mymsg = "Messages will be written to terminal"}
           "SILENT" {$mymsg = "Messages will be suppressed"}
           "OBJECT" {$mymsg = "Messages will be returned via OBJECT"}
           default  {$mymsg = "Messages will be written to file $Messages"}
        }
        Report "I"  $mymsg $ReturnOBJ $Messages

        if (!(Test-Path $InputFile)) {
            Throw "Input file $InputFile does not exist"
        }
        if (!($Actionlist -contains $action)) {
            Throw "Action $action is invalid"
        }
        if (!($Modelist -contains $Mode)) {
            Throw "Mode $mode is invalid"
        }
        if ($Lock) {
            $lockverbs = $lock.SPlit(",")
            if ($lockverbs.Count -ne 2) {
                Throw "Lock parameter $lock must consist of exactly two words"    
            } 
            $Resource = $lockverbs[0]
            $Process = $lockverbs[1]
            If (!($Resource -match "\w+")) {
                Throw "Lock resource $Resource is not valid"
            }
            If (!($Process -match "\w+")) {
                Throw "Lock process $process is not valid"
            }
        }

    
    }
    catch {    
        
        $errortext = $error[0]
        $scripterrormsg = "Input validation failed - $errortext"
        Report "E" "$scripterrormsg" $ReturnOBJ $Messages
    }
}

if (!$ReturnObj.AbEnd) {
    try {
        Report "I" "Perform requested actions" $ReturnOBJ $Messages
        if ($lock) {
            $m = & $ADHC_LockScript "Lock" "$Resource" "$Process" "10" "OBJECT"
            $ENQfailed = $false
            foreach ($msgentry in $m.MessageList) {
                $msglvl = $msgentry.level
                if ($msglvl -eq "E") {
                    $EQNfailed = $true
                }
                $msgtext = $msgentry.Message
                Report $msglvl $msgtext $ReturnOBJ $Messages
            } 
            if ($ENQfailed) {
                throw "ENQ Failed - Fatal error"
            }           
        } 
        $outdir = Split-Path $Outputfile
        if (!(Test-Path $outdir)) {
            New-Item  -Path "$outdir" -ItemType directory -force | Out-Null
            Report "I" "Directory $outdir did not exist but has been created now" $ReturnOBJ $Messages
        }
        if ($Action -eq "COPY") {
            if ($Mode -eq "REPLACE") {
                copy-item -path "$inputfile" -destination "$outputfile" -force
            }
            else {
                # MODE = APPEND
                $tussenfile = "$inputfile" + "T"
                copy-item -path "$inputfile" -destination "$tussenfile" -force
                $o = Get-Content $tussenfile                 
                Add-Content $outputfile $o
                Remove-Item $tussenfile
            }

           
        }
        else {
            # ACTION = MOVE
            if ($Mode -eq "REPLACE") {
                move-item -path "$inputfile" -destination "$outputfile" -force
            }
            else {
                # MODE = APPEND
                $tussenfile = "$inputfile" + "T"
                move-item -path "$inputfile" -destination "$tussenfile" -force
                $o = Get-Content $tussenfile                      
                Add-Content $outputfile $o
                Remove-Item $tussenfile
            }               
            
        }
        if ($Messages.ToUpper() -eq $Inputfile.ToUpper()) {
            $Messages = $OutputFile                           # from here messages are directed to output file
        }

        Report "I" "$ACTION with $MODE $inputfile to $outputfile completed" $ReturnOBJ $Messages
        if ($lock) {
            $m = & $ADHC_LockScript "Free" "$Resource" "$Process" "10" "OBJECT"
            foreach ($msgentry in $m.MessageList) {
                $msglvl = $msgentry.level
                $msgtext = $msgentry.Message
                Report $msglvl $msgtext $ReturnOBJ $Messages
            }            
        } 
              
    }
    catch {
        
        $errortext = $error[0]
        $scripterrormsg = "Requested actions failed - $errortext"
        Report "E" "$scripterrormsg" $ReturnOBJ $Messages
    }
}

$d = Get-Date
$Datum = " -- Date: " + $d.ToString("dd-MM-yyyy")
$Tijd = " -- Time: " + $d.ToString("HH:mm:ss")
$Scriptmsg = "*** ENDED ***** " + $mypath + " -- PowerShell script " + $MyName + $ScriptVersion + $Datum + $Tijd +$Node
Report "N" $Scriptmsg $ReturnOBJ $Messages

    
switch ($Messages) {        
    "JSON" {
        $ReturnJSON = ConvertTo-JSON $ReturnOBJ 
        return $ReturnJSON 
    } 
    Default { 
        return $Returnobj
    }
       
}

