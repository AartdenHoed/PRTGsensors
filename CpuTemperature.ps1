param (
    [string]$LOGGING = "YES", 
    [string]$myHost  = "NONE" ,
    [int]$sensorid = 77 
)
# $LOGGING = 'YES'
# $myHost = "hoesto"

function Running-Elevated
{
    $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $p = New-Object System.Security.Principal.WindowsPrincipal($id)

    if ($p.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) { 
        $adm = $true 
    }      
    else { 
        $adm = $false 
    }
    $MyAuth = [PSCustomObject] [ordered] @{ID = $id;
                                           Principal = $p; 
                                           Administrator = $adm}
    return $MyAuth 
 } 

$myhost = $myhost.ToUpper()


$ScriptVersion = " -- Version: 1.9.1"

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
    if ($log) {
        Add-Content $logfile $StringWriter.ToString()
    }
}

# END OF COMMON CODING
# Get Node status

$el = Running-Elevated

if ($log) {
    $u = $el.id.Name
    $ia = $el.Principal.Identity.IsAuthenticated
    Add-Content $logfile "==> Current user is $u (IsAuthenticated = $ia)" 
    if (-not($el.Administrator)) {
        Add-Content $logfile "==> Script NOT running as administrator"
    }
    else {
       Add-Content $logfile "==> Script running as administrator"
    }
}

if (!$scripterror) {
    try {
        if ($log) {
            Add-Content $logfile "==> Get NODE status for $myhost"
        }
        $nstat = & $ADHC_NodeInfoScript "$myHost"  "$LOGGING" 

    }
    Catch {
        if ($log) {
            Add-Content $logfile "==> Getting NODE status failed for $myhost"
        }
        $scripterror = $true
        $errortext = $error[0]
        $scripterrormsg = "Getting NODE status failed for $myhost - $errortext"

    }
    finally {
        if ($log) {
            foreach ($m in $nstat.MessageList) {
                $lvl = $m.Level
                $msg = $m.Message
                Add-COntent $logfile "($lvl) - $msg"

            }
        }
    }

}

$CpuTempScript = $mypath + "CpuTemperature2.PS1"
$duration = 0

if (!$scripterror) {
    try {
        # CpuTemp info of machine
        if ($log) {
            Add-Content $logfile "==> Get CPU temperature info from machine $myHost"
        }
        $invokable = $true
        
        try {
                           
            if ($log) {
                Add-Content $logfile "==> Remote script $CpuTempScript"
            }
            $b = Get-Date
                
            $myjob = Invoke-Command -ComputerName $myhost -FilePath $CpuTempscript  -Credential $ADHC_Credentials -JobName CpuTempJob  -AsJob 
                
            # write-host "Wait"
            $myjob | Wait-Job -Timeout 150 | Out-Null
            $e = Get-Date  

            if ($myjob) { 
                $mystate = $myjob.state
                $begin = $myjob.PSBeginTime
                $end = $myjob.PSEndTime
                $duration = ($end - $begin).seconds
                if ($duration -lt 0 ) {
                    $duration = ($e - $b).seconds
                }
            } 
            else {
                $mystate = "Unknown"
                $duration = ($e - $b).seconds 
            }
            if ($log) {
                $mj = $myjob.Name
                Add-Content $logfile "==> Remote job $mj ended with status $mystate"
            }
                                
            # Write-host $mystate
            if ($mystate -eq "Completed") {
                # write-host "YES"
                $CpuObject = (Receive-Job -Name CpuTempJob)
                $CpuTempInfo = $CpuObject.CPUlist
                if ($log) {
                    $m = "==> Called script ended with status " + $CpuObject.MyStatus + " --- Message: " + $CpuObject.Message
                    Add-Content $logfile $m
          
                } 
                if ($CpuObject.MyStatus -ne "Ok") {
                    Throw $CpuObject.Message
                }
                   
            }
            else {
                #write-host "NO"
                $invokable = $false
            }
            # write-host "Stop"
            $myjob | Stop-Job | Out-Null
            # write-host "Remove"
            $myjob | Remove-Job | Out-null
        }
        catch {
            # write-host "Catch"
            $invokable = $false
        }
        finally {
                
            # Write-Host $nodeisup
        }

    }
    catch {
        if ($log) {
            Add-Content $logfile "==> Getting CPU temperature info failed for $myHost"
          
        }
        $scripterror = $true
        $errortext = $error[0]
        $scripterrormsg = "==> Getting CPU temperature info failed for $myHost - $errortext"
    }
}
if (!$scripterror) {
    try {
        # 
        if ($log) {
            Add-Content $logfile "==> Process CPU temperature info"
        }
        # Init CpuTemp info file if not existent
        $str = $ADHC_CpuTempInfo.Split("\")
        $dir = $ADHC_OutputDirectory + $str[0] + "\" + $str[1]
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
        $CpuTempInfoFile = $ADHC_OutputDirectory + $ADHC_CpuTempInfo.Replace($ADHC_Computer, $myHost)
        $dr = Test-Path $CpuTempInfoFile
        $timestamp = Get-Date
        if (!$dr) {
            $rec = "$MyHost|INIT|||||" + $timestamp.ToString("dd-MM-yyyy HH:mm:ss")
            Set-Content $CpuTempInfoFile  $rec -force
        }
        
        $cpulist = @()
        if (!$invokable) {
            # Node not invokable, get info from file
            
            
            if ($log) {
                Add-Content $logfile "==> Node is down, get info from dataset"
            }
            $CpuTemplines = Get-Content $CpuTempInfoFile
            
            
            foreach ($entry in $CpuTemplines) {
                $split = $entry.Split("|")
                $CPUtype = $split[1]
                if ($CPUtype -eq "INIT") { 
                    # Status INIT, and no realtime info available
                    $timestamp = [datetime]::ParseExact($split[6].Trim(),"dd-MM-yyyy HH:mm:ss",$null)

                    $obj = [PSCustomObject] [ordered] @{Machine = $myhost; 
                                                    CPUtype = $CPUtype;
                                                    CPUname = "n/a";
                                                    CpuTempCurrent = 0;
                                                    CpuTempMin = 0;
                                                    CpuTempMax = 0;
                                                    Timestamp = $timestamp}                
                }
                else {
                                
                    $CpuName = $split[2]
                    $CpuTempCurrent = $split[3]
                    $CpuTempMin = $split[4]
                    $CpuTempMax = $split[5]                              
                    $timestamp = [datetime]::ParseExact($split[6].Trim(),"dd-MM-yyyy HH:mm:ss",$null)
                    $obj = [PSCustomObject] [ordered] @{Machine = $myhost; 
                                                        CPUtype = $CPUtype;
                                                        CPUname = $CPUname;
                                                        CpuTempCurrent = $CpuTempCurrent;
                                                        CpuTempMin = $CpuTempMin;
                                                        CpuTempMax = $CpuTempMax;
                                                        Timestamp = $timestamp}
                }
                $cpulist += $obj
                $leeftijd = ($d - $obj.timestamp).TotalMinutes
           }
        }

        else {
            # Node is UP, take real time info and write it tot dataset
               
            if ($log) {
                Add-Content $logfile "==> Node is up, get realtime info and write it to dataset"
            }         
            
            $firstrec = $true
            $timestamp = Get-Date
            foreach ($cpu in $CpuTempInfo){
                
                $obj = [PSCustomObject] [ordered] @{Machine = $myhost; 
                                                        CPUtype = $CPU.Type;
                                                        CPUname = $CPU.Name;
                                                        CpuTempCurrent = $Cpu.TempCurrent;
                                                        CpuTempMin = $Cpu.TempMin;
                                                        CpuTempMax = $Cpu.TempMax;
                                                        Timestamp = $timestamp}
                $cpulist += $obj
                
                $record = $myhost + "|" + $obj.CPUtype + "|" + $obj.CPUname + "|" + $obj.CpuTempCurrent + "|" + 
                        $obj.CpuTempMin + "|" + $obj.CpuTempMax + "|" + $timestamp.ToString("dd-MM-yyyy HH:mm:ss")
            
                if ($firstrec) {
                        Set-Content $CpuTempInfofile $record
                        # Write-host "First"
                        $firstrec = $false
                }
                else { 
                    Add-Content $CpuTempInfofile $record
                }
            }
                        
        }

    }   
    catch {
        if ($log) {
            Add-Content $logfile "==> Processing CpuTemp info failed for $myHost"
          
        }
        $scripterror = $true
        $errortext = $error[0]
        $scripterrormsg = "==> Processing CpuTemp info failed for $myHost - $errortext"
    }
}

if ($log) {
    Add-Content $logfile "==> Create XML"
}

[xml]$xmldoc = New-Object System.Xml.XmlDocument
$decl = $xmldoc.CreateXmlDeclaration('1.0','Windows-1252',$null)

[void]$xmldoc.AppendChild($decl)

$PRTG = $xmldoc.CreateElement('PRTG')

# Node status (PRIMARY Channel)
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
$NotifyChanged = $xmldoc.CreateElement('NotifyChanged')
$ValueLookup =  $xmldoc.CreateElement('ValueLookup')

$Channel.InnerText = "Node status"
$Unit.InnerText = "Custom"
$Mode.Innertext = "Absolute"
$ValueLookup.Innertext = 'NodeStatus'

if ($invokable) {
    $Value.Innertext = $nstat.StatusCode + 1
    $livestat = $nstat.Status + ", Invokable"
    $online = "realtime info"
} 
else { 
   $Value.Innertext = $nstat.StatusCode
   $livestat = $nstat.Status + ", Not Invokable"
   if (-not($el.Administrator) -and ($myHost -eq $ADHC_Computer.ToUpper()) -and ($leeftijd -le 10)) {
       $online = "near realtime info"
   }
   else {
       $online = "offline info"
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

# Wait time
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')    
$Unit = $xmldoc.CreateElement('Unit')
$Mode = $xmldoc.CreateElement('Mode')
$NotifyChanged = $xmldoc.CreateElement('NotifyChanged')
    
$Channel.InnerText = "Remote wait time (sec)"
$Unit.InnerText = "TimeSeconds"
$Mode.Innertext = "Absolute"
$Value.Innertext = $duration

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($NotifyChanged)
[void]$Result.AppendChild($Mode)
    
[void]$PRTG.AppendChild($Result)

foreach ($item in $cpulist) {

    $CpuType = $item.CPUtype

    If ($CPUtype -eq "INIT") {
        $CPUtype = "Not known yet (INIT phase)" 
        break
    }
   
    # Report Current Temp
    $Result = $xmldoc.CreateElement('Result')
    $Channel = $xmldoc.CreateElement('Channel')
    $Value = $xmldoc.CreateElement('Value')    
    $Unit = $xmldoc.CreateElement('Unit')
    $Mode = $xmldoc.CreateElement('Mode')
    $NotifyChanged = $xmldoc.CreateElement('NotifyChanged')
    
    $cname = $item.Machine + " " + $item.CpuName + " Current (C)"
    $Channel.InnerText = $cname
    $Unit.InnerText = "Temperature"
    $Mode.Innertext = "Absolute"
    $Value.Innertext = $item.CpuTempCurrent

    [void]$Result.AppendChild($Channel)
    [void]$Result.AppendChild($Value)
    [void]$Result.AppendChild($Unit)
   
    [void]$Result.AppendChild($NotifyChanged)
    
    [void]$Result.AppendChild($Mode)
    
    [void]$PRTG.AppendChild($Result)

    # Report Minimum Temp
    $Result = $xmldoc.CreateElement('Result')
    $Channel = $xmldoc.CreateElement('Channel')
    $Value = $xmldoc.CreateElement('Value')    
    $Unit = $xmldoc.CreateElement('Unit')
    $Mode = $xmldoc.CreateElement('Mode')
    $NotifyChanged = $xmldoc.CreateElement('NotifyChanged')
    
    $cname = $item.Machine + " " + $item.CpuName + " Minimum (C)"
    $Channel.InnerText = $cname
    $Unit.InnerText = "Temperature"
    $Mode.Innertext = "Absolute"
    $Value.Innertext = $item.CpuTempMin

    [void]$Result.AppendChild($Channel)
    [void]$Result.AppendChild($Value)
    [void]$Result.AppendChild($Unit)
    
    [void]$Result.AppendChild($NotifyChanged)
    
    [void]$Result.AppendChild($Mode)
    
    [void]$PRTG.AppendChild($Result)

    # Report Maximum Temp
    $Result = $xmldoc.CreateElement('Result')
    $Channel = $xmldoc.CreateElement('Channel')
    $Value = $xmldoc.CreateElement('Value')    
    $Unit = $xmldoc.CreateElement('Unit')
    $Mode = $xmldoc.CreateElement('Mode')
    $NotifyChanged = $xmldoc.CreateElement('NotifyChanged')
    
    $cname = $item.Machine + " " + $item.CpuName + " Maximum (C)"
    $Channel.InnerText = $cname
    $Unit.InnerText = "Temperature"
    $Mode.Innertext = "Absolute"
    $Value.Innertext = $item.CpuTempMax

    [void]$Result.AppendChild($Channel)
    [void]$Result.AppendChild($Value)
    [void]$Result.AppendChild($Unit)
  
    [void]$Result.AppendChild($NotifyChanged)
   
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
    $formattime = $timestamp.ToString("dd-MM-yyyy HH:mm:ss")
    $message = "Machine $myhost (now $livestat) *** CPU type: $cputype *** Timestamp: $formattime ($online) *** Script $scriptversion"
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