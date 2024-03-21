# Lock/unlock a global resource
param (
    [string]$Action = "LOCK", 
    [string]$EnqName = "EnqName",
    [string]$Process = "Process", 
    [int] $Duration = 10,
    [string] $Mode = "xx" 
)

#TestValues####################################
#$Action = "FREE"
#$ENQNAME = "DEPLOY"
#$PROCESS = "Ikkuh"
#$waittime = 15
#$Mode = "SILENT"
#TestValues####################################
$Version = " -- Version: 3.4.1"
$Mode = $mode.ToUpper()
$Waittime = 150

$InformationPreference = "Continue"
$WarningPreference = "Continue"
$ErrorActionPreference = "Stop"

$global:MessageList = @()


class GlobalLockException : System.Exception  { 
    GlobalLockException( [string]$message) : base($message) {

    }

    GlobalLockException() {

    }
}

function CreateLockRecord ([string]$Status, [string]$Machine, [string]$Who, [string]$What, [datetime]$Start, [datetime]$Stop)  {
   
    $record = $Status.ToUpper() + "|" + $Machine + "|" + $Who + "|" + $What + "|" + $Start.ToSTring("dd-MM-yyyy HH:mm:ss") + "|" + $Stop.ToSTring("dd-MM-yyyy HH:mm:ss")
    return $record
}
function AddMessage ([string]$level, [string]$msg) {
    $msgentry = [PSCustomObject] [ordered] @{Level = $level;
                                             Message = $msg}
    $global:MessageList += $msgentry
    # Write-Host $msg
    
    return  
}

function Lock ([string]$InternalAction, [string]$Machine, [string]$Who, [string]$What) {

    switch ($InternalAction.ToUpper()) {
        "INIT" {
            $LockName = $Computer + "_" + $Process + "_"+ $EnqName
            $LockDsn = $LockName + ".glock"
            $LockFullName = $ADHC_LockDir + $LockDsn

            if (!(Test-Path $ADHC_LockDir)) {
                New-Item -ItemType Directory -Force -Path $ADHC_LockDir | Out-Null
            } 
            if (!(Test-Path $LockFullName)) {
                New-Item -Path $ADHC_LockDir -Name "$LockDsn" -ItemType File | Out-Null
                $i1 =[datetime]::ParseExact("01-01-2000 00:00:00","dd-MM-yyyy HH:mm:ss",$null) 
                $i2 = [datetime]::ParseExact("01-01-2000 00:10:00","dd-MM-yyyy HH:mm:ss",$null)
                $R = CreateLockRecord "INIT" $Computer $Process $EnqName $i1 $i2
                Set-Content $LockFullNAme $R -force
            }
            return "ok"    

        }
        "TEST" {
            $search = "^.+_.+_" + $what + "\.glock$"
            # Write-Host "Zoekargument $search"
            # Write-Host "Directory $ADHC_Lockdir"            
            
            $locklist = Get-ChildItem -Path $ADHC_LockDir | Select-Object FullName,Name | Where-Object {$_.Name -match "$search" }            
            $RetryCount = 0
            do { 
                
                $RetryCount += 1
                $ResourceFree = $True
                AddMessage "I" "Try number $RetryCOunt to lock resource $what for process $Process on computer $Computer"  
                        
                foreach ($entry in $locklist) {
                    
                    $E = $entry.FullName
                    # Write-Host "Dataset $E"
                    $lockrecord = Get-Content $entry.FullName
                    if ($lockrecord) {
                        $lockbits = $lockrecord.Split("|")
                    }
                    else {
                        $lockbits = @("NO","N0")
                    }
                    if ($lockbits.Count -lt 6) {                                   # file corrupted for some reason 
                        $i1 =[datetime]::ParseExact("01-01-2000 00:00:00","dd-MM-yyyy HH:mm:ss",$null) 
                        $i2 = [datetime]::ParseExact("01-01-2000 00:10:00","dd-MM-yyyy HH:mm:ss",$null)
                        $lockrecord = CreateLockRecord "FREE" "???" "????" "DEPLOY" $i1 $i2
                        Set-Content $E $lockrecord -force    
                        AddMessage "A" "Lockfile $E is corrupted, and has been bypassed" 
                        $lockbits = $lockrecord.Split("|") 
                    }
                    # $lockbits
                    $Lockstatus  = $lockbits[0]
                    $lockmachine = $lockbits[1]
                    $lockprocess = $lockbits[2]
                    $lockenqname = $lockbits[3]
                    $lockfrom    = [datetime]::ParseExact($lockbits[4],"dd-MM-yyyy HH:mm:ss",$null) 
                    $lockuntil   = [datetime]::ParseExact($lockbits[5],"dd-MM-yyyy HH:mm:ss",$null)
                    # Write-Host $lockuntil.ToSTring()
                    $Now = Get-Date
                    $u = $lockuntil.ToString("dd-MM-yyyy HH:mm:ss")
                    # write-host $lockstatus
                    if ($lockstatus -eq "LOCK") {
                        if ($lockuntil -gt $now) {   # Resource is locked by a process
                            $ResourceFree = $false
                            $msg = "Resource $lockenqname locked by process $lockprocess on computer $lockmachine Until $u"  
                            AddMessage "I" $msg
                            if (($Mode -ne "JSON") -and ($Mode -ne "SILENT"))  {
                                Write-Host $msg
                            } 
                        }
                                            
                    }
                    
                } 
                if (!$ResourceFree) { 
                    $msg = "Resource $what not available now. Wait for $waittime seconds" 
                    AddMessage "A" $msg
                    if (($Mode -ne "JSON") -and ($Mode -ne "SILENT")) {
                        Write-Host $msg
                    } 
                    Start-Sleep -s $waittime
                } 
                else {
                    AddMessage "I" "Resource $what not locked by any other process" 
                }
            } until ($ResourceFree)
            
            return "ok" 

        }
        "LOCK"{
            # Set lock
            $Now = Get-Date
            $End = $now.AddMinutes($duration) 
            $LockName = $Computer + "_" + $Process + "_"+ $EnqName
            $LockDsn = $LockName + ".glock"
            $LockFullName = $ADHC_LockDir + $LockDsn
            $R = CreateLockRecord "LOCK" $Computer $Process $EnqName $Now $End
            $u = $End.ToString("dd-MM-yyyy HH:mm:ss")
            Set-Content $LockFullNAme $R  
            AddMessage "I" "Process $Process on computer $Computer locked resource $what successfully now until $u"
            return "ok"                
 
        }
        "FREE"{
            $LockName = $Computer + "_" + $Process + "_"+ $EnqName
            $LockDsn = $LockName + ".glock"
            $LockFullName = $ADHC_LockDir + $LockDsn

            $lockrecord = Get-Content $lockFullName
            $lockbits = $lockrecord.Split("|")
            # $lockbits
            $Lockstatus  = $lockbits[0]
            $lockmachine = $lockbits[1]
            $lockprocess = $lockbits[2]
            $lockenqname = $lockbits[3]
            $lockfrom    = [datetime]::ParseExact($lockbits[4],"dd-MM-yyyy HH:mm:ss",$null) 
            $lockuntil   = [datetime]::ParseExact($lockbits[5],"dd-MM-yyyy HH:mm:ss",$null)
            $R = CreateLockRecord "FREE" $Computer $Process $EnqName $lockfrom $lockuntil
            Set-Content $LockFullNAme $R 
            AddMessage "I" "Resource $what freed by process $Process on computer $COmputer"
            return "ok" 
        }
        "VRFY"{
            $search = "^.+" + $what + "\.glock$"
            # Write-Host "Zoekargument $search"
            $locklist = Get-ChildItem -Path $ADHC_LockDir | Select-Object FullName,Name | Where-Object {$_.Name -match $search}            
            $ResourceFree = $True
              
            foreach ($entry in $locklist) {
                $E = $entry.FullNAme
                $lockrecord = Get-Content $entry.FullName
                $lockbits = $lockrecord.Split("|")
                # $lockbits
                $Lockstatus  = $lockbits[0]
                $lockmachine = $lockbits[1]
                $lockprocess = $lockbits[2]
                $lockenqname = $lockbits[3]
                $lockfrom    = [datetime]::ParseExact($lockbits[4],"dd-MM-yyyy HH:mm:ss",$null) 
                $lockuntil   = [datetime]::ParseExact($lockbits[5],"dd-MM-yyyy HH:mm:ss",$null)
                $Now = Get-Date
                if (($lockmachine -eq $Machine) -and ($lockprocess -eq $who)) {
                    # this is my self, skip
                    continue
                }
                if ($lockstatus -eq "LOCK") {
                    if ($lockuntil -gt $now) {   # Resource is locked by a process
                        $ResourceFree = $false
                    }
                    
                }
            } 
            
            if (!$ResourceFree) { 
                AddMessage "A" "Failed to verify resource $what for process $Process on computer $Computer" 
                return "nok"
            }
            else {
                AddMessage "I" "Resource $what succesfully verified for process $Process on computer $Computer" 
                return "ok"
            }            

      
        }
        Default {
            $MyError = [GlobalLockException]::new("Internal action $internalaction Unknown...")
            throw $MyError
        }

    }
    return
}

try {
    
    $Node = " -- Node: " + $env:COMPUTERNAME
    $d = Get-Date
    $Datum = " -- Date: " + $d.ToString("dd-MM-yyyy")
    $Tijd = " -- Time: " + $d.ToString("HH:mm:ss")
    $myname = $MyInvocation.MyCommand.Name
    $FullScriptName = $MyInvocation.MyCommand.Definition
    $mypath = $FullScriptName.Replace($MyName, "")
       
    $myname = $MyInvocation.MyCommand.Name
    $Scriptmsg = "*** STARTED *** " + $mypath + " -- PowerShell script " + $MyName + $Version + $Datum + $Tijd +$Node

    AddMessage "N" $Scriptmsg
    if (($Mode -ne "JSON") -and ($Mode -ne "SILENT")) {
        Write-Host $scriptmsg
    }

    $LocalInitVar = $mypath + "InitVar.PS1"  
    & "$LocalInitVar" -JSON Silent
    
    if (!$ADHC_InitSuccessfull) {
        # Write-Warning "YES"
        throw $ADHC_InitError
    }  

    $EnqSuccess = $true

    # ENQ Name
    $Computer = $ADHC_Computer

    $COmputer = $Computer.Replace("_", "-").ToUpper().Trim()
    $Process =  $Process.Replace("_", "-").ToUpper().Trim()
    $EnqName =  $EnqName.Replace("_", "-").ToUpper().Trim()

    
    Try {
        # Check availabiliy of ONEDRIVE. If NOT available, abort execution
        $OD = Test-Connection -COmputerName onedrive.live.com -Count 1 -ErrorAction Stop
    }
    Catch {
        #Write-Host "Catch"
        if ($action -eq "LOCK"){
            $MyError = [GlobalLockException]::new("OneDrive not available, $Action of resource $EnqName impossible")
            throw $MyError
        }
        else {
            AddMessage "A" "Onedrive no longer available... could not FREE resource $Enqname for process $process on computer $computer"
        }
       
    }
    

    
    Switch ($Action.ToUpper()) {
        "LOCK" {
            $lockset = $false
            $trycount = 0
            DO {
                $trycount += 1
                try {
                    $a = Lock "Init" $Computer $process $EnqName
                    # Check if lock is free
                    $b = Lock "Test" $Computer $process $EnqName
                    # write-host "test return"
                    # If Test returns, the lock is free, so get it!
                    $c = Lock "Lock" $Computer $process $EnqName 
                    # After this, verify that no other process crossed the lock 
                    $d = Lock "VRFY" $Computer $process $EnqName

                    if ($d -eq "ok") {
                        $lockset = $true
                    }
                    else {                    
                        AddMessage "A" "Resource $what could not be verified, retry..." 
                        $e = Lock "FREE" $Computer $process $EnqName
                    }
                }
                Catch {                   
                   $ErrorMessage = $_.Exception.Message
                   AddMessage "A" "Lock failed due to external reason: " 
                   AddMessage "A" $ErrorMessage
                   $lockset = $false
                   Start-Sleep -Seconds 5
                   if ($trycount -ge 5) {
                        AddMessage "A" "Persistent lock failure (5 attempts executed)"
                        $MyError = [GlobalLockException]::new("Persistent lock failure (5 attempts executed)")
                        throw $MyError
                   }
                }
    
            } until ($lockset)

            
        }

        "FREE" {
            try {
                $f = Lock "FREE" $Computer $process $EnqName
            }
            catch {
                $ErrorMessage = $_.Exception.Message
                AddMessage "A" "FREE action failed due to error: $Errormessage"
            }           

        }
        Default {
            $MyError = [GlobalLockException]::new("External action $action Unknown...")
            throw $MyError
        }
    }


}
Catch {
   $ErrorMessage = $_.Exception.Message
   AddMessage  "E" $ErrorMessage
   $EnqSuccess = $false
    
}
$d = Get-Date
$Datum = " -- Date: " + $d.ToString("dd-MM-yyyy")
$Tijd = " -- Time: " + $d.ToString("HH:mm:ss")
$Scriptmsg = "*** ENDED ***** " + $mypath + " -- PowerShell script " + $MyName + $Version + $Datum + $Tijd +$Node
AddMessage "N" $Scriptmsg

if (($Mode -ne "JSON") -and ($Mode -ne "SILENT")) {
   
    Write-Host $scriptmsg
}


if ($Mode.ToUpper() -eq  "JSON" ) {
       
    $ReturnJSON = ConvertTo-JSON $global:MessageList     
     
    return $ReturnJSON 
}
else {
    
    Return $global:MessageList   
}


