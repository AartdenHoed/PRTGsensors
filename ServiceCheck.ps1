

param (
    [string]$LOGGING = "YES", 
    [string]$myHost  = "NONE" ,
    [int]$sensorid = 77 
)
# $LOGGING = 'YES'
# $myHost = "hoesto"

$myhost = $myhost.ToUpper()

$ScriptVersion = " -- Version: 1.0.4"

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


$duration = 0

if (!$scripterror) {
    try {
        # Service info of machine
        if ($log) {
            Add-Content $logfile "==> Get service info from machine $myHost"
        }
        $invokable = $true
        if ($myHost -eq $ADHC_Computer.ToUpper()) {
            $begin = Get-Date
                    
            $ServiceInfo = Get-WmiObject win32_service | select PSComputerName, SystemName, Name, Caption, Displayname, `
                                     PathName, ServiceType, StartMode, `
                                     Started, State, Status, ExitCode, Description   

            $end = Get-Date
            $duration = ($end - $begin).seconds

        }
        else {
            try {
                $b = Get-Date
                $myjob = Invoke-Command -ComputerName $myhost `
                    -ScriptBlock { Get-WmiObject win32_service | select PSComputerName, SystemName, Name, Caption, Displayname, `
                                     PathName, ServiceType, StartMode, `
                                     Started, State, Status, ExitCode, Description   } `
                    -Credential $ADHC_Credentials -JobName ServiceJob  -AsJob
                
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
                    #write-host "YES"
                    $ServiceInfo = (Receive-Job -Name ServiceJob)
                }
                else {
                    #write-host "NO"
                    $invokable = $false
                }
                
                $myjob | Stop-Job | Out-Null
                $myjob | Remove-Job | Out-null
            }
            catch {
                $invokable = $false
            }
            finally {
                # Write-Host $nodeisup
            }
        }
    }
    catch {
        if ($log) {
            Add-Content $logfile "==> Getting service info failed for $myHost"
          
        }
        $scripterror = $true
        $errortext = $error[0]
        $scripterrormsg = "==> Getting service info failed for $myHost - $errortext"
    }

}

if (!$scripterror) {
    try {
        # 
        if ($log) {
            Add-Content $logfile "==> Process service info"
        }
        # Init service info file if not existent
        $str = $ADHC_ServiceList.Split("\")
        $dir = $ADHC_OutputDirectory + $str[0] + "\" + $str[1]
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
        $ServiceListFile = $ADHC_OutputDirectory + $ADHC_ServiceList.Replace($ADHC_Computer, $myHost)
        $dr = Test-Path $ServiceListFile
        $timestamp = Get-Date
        if (!$dr) {
            $rec = "$MyHost|INIT||||||||||||||" + $timestamp.ToString("dd-MM-yyyy HH:mm:ss")
            Set-Content $ServiceListFile  $rec -force
        }
        
        $servicelist = @()
        if (!$invokable) {
            # Node not invokable, get info from file
            
            
            if ($log) {
                Add-Content $logfile "==> Node is down, get info from dataset"
            }
            $Servicelines = Get-Content $ServiceListFile
            
            
            foreach ($entry in $Servicelines) {
                $split = $entry.Split("|")
                $sysname = $split[1]
                if ($sysname -eq "INIT") { 
                    # Status INIT, and no realtime info available
                    $timestamp = [datetime]::ParseExact($split[15].Trim(),"dd-MM-yyyy HH:mm:ss",$null)

                    $obj = [PSCustomObject] [ordered] @{ComputerName = $myhost ;
                                                        SystemName     = $sysname ;
                                                        Name           = "Unknown";
                                                        Caption        = "Unknown";
                                                        Displayname    = "Unknown";
                                                        PathName       = "n/a";
                                                        ServiceType    = "n/a";
                                                        StartMode      = "n/a";
                                                        Started        = $false
                                                        State          = "n/a";
                                                        Status         = "n/a";
                                                        ExitCode       = 999
                                                        Description    = "n/a";
                                                        ProgramName    = "n/a";
                                                        Software       = "Unknown"
                                                        Timestamp      = $timestamp}                
                }
                else {
                                
                    $ComputerName     = $split[0]
                    $SystemName  = $split[1]                                                        
                    $Name        = $split[2]
                    $Caption     = $split[3]
                    $Displayname = $split[4]
                    $PathName    = $split[5]
                    $ServiceType = $split[6]
                    $StartMode   = $split[7]
                    $Started     = $split[8]
                    $State       = $split[9]
                    $Status      = $split[10]
                    $ExitCode    = $split[11]
                    $Description = $split[12]
                    $ProgramName = $split[13]
                    $Software    = $split[14]                               
                    $timestamp = [datetime]::ParseExact($split[15].Trim(),"dd-MM-yyyy HH:mm:ss",$null)
                    $obj = [PSCustomObject] [ordered] @{ComputerName = $ComputerName ;
                                                        SystemName     = $SystemName ;
                                                        Name           = $Name ;
                                                        Caption        = $Caption ;
                                                        Displayname    = $DisplayName;
                                                        PathName       = $PathName;
                                                        ServiceType    = $ServiceType;
                                                        StartMode      = $StartMode;
                                                        Started        = $Started
                                                        State          = $State;
                                                        Status         = $Status;
                                                        ExitCode       = $ExitCode;
                                                        Description    = $Description;
                                                        ProgramName    = $ProgramName;
                                                        Software       = $Software;
                                                        Timestamp      = $timestamp}      
                }
                $servicelist += $obj
                
           }
        }

        else {
            # Node is UP, take real time info and write it tot dataset
               
            if ($log) {
                Add-Content $logfile "==> Node is up, get realtime info and write it to dataset"
            }         
            
            $firstrec = $true
            $timestamp = Get-Date
            $timestring = $timestamp.ToString("dd-MM-yyyy HH:mm:ss")

            foreach ($service in $ServiceInfo){
                # Determine program name without arguments
                $thispath = $service.PathName
                $ProgramName = ''
                if ($thispath -match '"(.*?)"') {
                    $ProgramName = $matches[1]
                }
                else {
                    if ($thispath -match '(.*?)\s'){
                        $ProgramName = $matches[1]
                    }
                    else {
                        $ProgramName = $thispath    
                    }
                }
                if (-not $ProgramName) {
                    $ProgramName = "Unknown"
                }
                # Guess program name form directory name
                $software = " "
                $spl = $ProgramName.Split("\")
                if ($spl.count -eq 3) {
                    $software = $spl[1]
                }
                else {
                    $software = $spl[2]
                }
                if (-not $Software) {
                    $Software = "Unknown"
                }
                
                $obj = [PSCustomObject] [ordered] @{ComputerName = $service.PSComputerName ;
                                                        SystemName     = $service.SystemName ;
                                                        Name           = $service.Name ;
                                                        Caption        = $service.Caption ;
                                                        Displayname    = $service.DisplayName;
                                                        PathName       = $service.PathName;
                                                        ServiceType    = $service.ServiceType;
                                                        StartMode      = $service.StartMode;
                                                        Started        = $service.Started;
                                                        State          = $service.State;
                                                        Status         = $service.Status;
                                                        ExitCode       = $service.ExitCode;
                                                        Description    = $service.Description;
                                                        ProgramName    = $ProgramName;
                                                        Software       = $Software;
                                                        Timestamp      = $timestamp}      
                $servicelist += $obj
                
                $record = $obj.PSComputerName + "|" +
                          $obj.SystemName + "|" +
                          $obj.Name + "|" +
                          $obj.Caption + "|" +
                          $obj.DisplayName + "|" +
                          $obj.PathName  + "|" +
                          $obj.ServiceType  + "|" +
                          $obj.StartMode  + "|" +
                          $obj.Started + "|" +
                          $obj.State + "|" +
                          $obj.Status  + "|" +
                          $obj.ExitCode  + "|" +
                          $obj.Description  + "|" +
                          $obj.ProgramName  + "|" +
                          $obj.Software + "|" +
                          $timestring
            
                if ($firstrec) {
                    Set-Content $ServiceListfile $record
                    # Write-host "First"
                    $firstrec = $false
                }
                else { 
                    Add-Content $ServiceListfile $record
                }
            }
                        
        }

    }   
    catch {
        if ($log) {
            Add-Content $logfile "==> Processing Service info failed for $myHost"
          
        }
        $scripterror = $true
        $errortext = $error[0]
        $scripterrormsg = "==> Processing Service info failed for $myHost - $errortext"
    }
}


$nrofservices = $servicelist.Count

if ($nrofservices -eq 1) {
    if ($servicelist[0].SystemName -eq "INIT") {
        $nrofservices = 0
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
   $online = "offline info"
   
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

# Service status
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')    
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
$NotifyChanged = $xmldoc.CreateElement('NotifyChanged')
$ValueLookup =  $xmldoc.CreateElement('ValueLookup')
    
$Channel.InnerText = "Service Status"
$Unit.InnerText = "Custom"
$Mode.Innertext = "Absolute"
$ValueLookup.Innertext = 'ServiceStatus'

$Value.Innertext = 0    #====================== LET OP

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($NotifyChanged)
[void]$Result.AppendChild($Mode)
    
[void]$PRTG.AppendChild($Result)

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
    $message = "Machine $myhost (now $livestat) *** $nrofservices Services found *** Timestamp: $formattime ($online) *** Script $scriptversion"
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

