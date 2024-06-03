param (
    [string]$LOGGING = "YES", 
    [string]$myHost  = "NONE" ,
    [int]$sensorid = 77 
)
# $LOGGING = 'YES'
# $myHost = "ADHC-2"

$myhost = $myhost.ToUpper()

$ScriptVersion = " -- Version: 3.3"

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

$CheckDate = Get-Date
$timestring = $CheckDate.ToString("yyyy-MM-ddTHH:mm:ss")

# Get computerID to use
$query = "SELECT [ComputerID]      
                FROM [dbo].[Computer]
                WHERE ComputerName = '" + $myhost + "'"
$DbResult = invoke-sqlcmd -ServerInstance '.\sqlexpress' -Database "Sympa" `
                            -Query "$query" `
                            -ErrorAction Stop 
if (!$DbResult) {
    $scripterrormsg = "==> Host $myHost not found in database" 
    if ($log) {
        Add-Content $logfile $scripterrormsg          
    }
    $scripterror = $true
}
else {
    $computerid = $DbResult.ComputerID
}
# Get componentID to use
$query = "SELECT [ComponentID]      
                FROM [dbo].[Component]
                WHERE ComponentNameTemplate = '*** Unknown ***'"
$DbResult = invoke-sqlcmd -ServerInstance '.\sqlexpress' -Database "Sympa" `
                            -Query "$query" `
                            -ErrorAction Stop 
if (!$DbResult) {
    $scripterrormsg = "==> Component(Template) '*** Unknown ***' not found in database" 
    if ($log) {
        Add-Content $logfile $scripterrormsg          
    }
    $scripterror = $true
}
else {
    $componentid = $DbResult.ComponentID
}

if (!$scripterror) {
    try {
        # 
        if ($log) {
            Add-Content $logfile "==> Process service info"
        }
                
        if (!$invokable) {
            # Node not invokable, get info from file
            
            
            if ($log) {
                Add-Content $logfile "==> Node is down, get info from SQL database"
            }

            $query = "Select Count(*)
                        From dbo.Service
                        WHERE SystemName = '" + $myhost + "' AND ChangeState = 'Current'"  
            $DBresult = invoke-sqlcmd -ServerInstance '.\sqlexpress' -Database "Sympa" `
                        -Query "$query" `
                        -ErrorAction Stop 
            $nrofservices = $DBresult.Item(0)
            
        }

        else {
            # Node is UP, take real time info and write it tot dataset
               
            if ($log) {
                Add-Content $logfile "==> Node is up, get realtime info and write it to database"
            }         
            
            $nrofservices = $serviceinfo.Count
            
            foreach ($service in $ServiceInfo){
                # Determine program name without arguments
                if ($service.Description -eq $null) {
                    $service.Description = "n/a"
                }
                # Write-Host $service.Name
                $thispath = $service.PathName
                if ($service.Name -Match "\w+_[0-9a-fA-F]{5}") {
                    $s = $service.Name.Split("_")
                    $service.Name = $s[0]
                    $suffix = $s[1]    
                }
                else {
                    $suffix = ""
                }
                if ($thispath -ne $null) {
                
                    $thispath = $thispath.Replace('"', '') 

                    
                    if ($thispath.toupper() -match "\w+\.EXE\s*") {
                        $pname = $matches[0].Trim()
                        $epos = $thispath.ToUpper().IndexOf($pname)
                        $DirName = $thispath.substring(0, $epos-1)
     
                        $ProgramName = $thispath.substring($epos,$pname.Length)
                    }
                    $FullName = $Dirname + '\' + $ProgramName
                                       
                    if ($FullName -eq $thispath) {
                        $Parameter = ''
                    }
                    else {
                        $Parameter = $thispath.Replace($FullName + " ", '')
                    }

                    # Guess program name form directory name
                    $software = " "
                    $spl = $FullName.Split("\")
                    if ($spl.count -eq 3) {
                        $software = $spl[1]
                    }
                    else {
                        $software = $spl[2]
                    }
                    if (-not $Software) {
                        $Software = "Unknown"
                    }
                }
                else {
                    $Software = "Unknown"
                    $Parameter = ''
                    $Dirname = "Unknown"
                    $ProgramName = "Unknown"

                }
                
                $obj = [PSCustomObject] [ordered] @{ComputerName = $service.PSComputerName ;
                                                        SystemName     = $service.SystemName ;
                                                        Name           = $service.Name ;
                                                        Suffix         = $suffix;
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
                                                        DirName        = $DirName;
                                                        ProgramName    = $ProgramName;
                                                        Parameter      = $Parameter;
                                                        Software       = $Software;
                                                        CheckDate      = $CheckDate}      
                                
                 
                # Update database
                $query = "SELECT * FROM dbo.Service WHERE SystemName = '" + $obj.SystemName + 
                            "' AND Name = '" + $obj.Name + 
                            "' AND ChangeState = 'Current'" 
                $DBresult = invoke-sqlcmd -ServerInstance '.\sqlexpress' -Database "Sympa" `
                        -Query "$query" `
                        -ErrorAction Stop
                if ($DBresult -ne $null) {
                    if ($DBresult.Dirname.Trim() -eq $obj.DirName.Trim() -and $DBresult.ProgramName.Trim() -eq $obj.ProgramName.Trim() -and $DBresult.Parameter.Trim() -eq $obj.Parameter.Trim()) {
                        # service not changed, just update CheckDate
                        # Write-Host "Service not changed, just update checkdate"
                        $query = "UPDATE dbo.Service
                                   SET [PSComputerNAme] = '" + $obj.ComputerName + "'
                                      ,[Caption] = '"        + $obj.Caption      + "'
                                      ,[Suffix] = '"         + $obj.Suffix       + "'
                                      ,[DisplayName] = '"    + $obj.DisplayName  + "'  
                                      ,[PathName] = '"       + $obj.PathName     + "'
                                      ,[ServiceType] = '"    + $obj.ServiceType  + "'
                                      ,[StartMode] = '"      + $obj.StartMode    + "'
                                      ,[Started] = '"        + $obj.Started      + "'
                                      ,[State] = '"          + $obj.State        + "'
                                      ,[Status] = '"         + $obj.Status       + "'
                                      ,[ExitCode] = "        + $obj.ExitCode     + "
                                      ,[Description] = '"    + $obj.Description.Replace("'","''")  + "'
                                      ,[DirName] = '"        + $Obj.DirName      + "'
                                      ,[ProgramName] = '"    + $obj.ProgramName  + "'
                                      ,[Parameter] = '"      + $obj.Parameter    + "'
                                      ,[CheckDate] = '"      + $timestring       + 
                                      "' WHERE SystemName = '" + $obj.SystemName + 
                                         "' AND Name = '" + $obj.Name + 
                                         "' AND ChangeState = 'Current'"   
                        $DBresult = invoke-sqlcmd -ServerInstance '.\sqlexpress' -Database "Sympa" `
                        -Query "$query" `
                        -ErrorAction Stop  
                    }
                    else {
                        # Service Changed. Turn current record to old, and insert new current record with new startime
                        # Write-Host "Service changed, set chagestate to old and create new current record"
                        $query = "UPDATE dbo.Service
                                   SET [ChangeState] = 'Old'
                                      ,[CheckDate] = '"      + $timestring       + 
                                      "' WHERE SystemName = '" + $obj.SystemName + 
                                         "' AND Name = '" + $obj.Name + 
                                         "' AND ChangeState = 'Current'"   
                        $DBresult = invoke-sqlcmd -ServerInstance '.\sqlexpress' -Database "Sympa" `
                                    -Query "$query" `
                                    -ErrorAction Stop  
                        $query = "INSERT INTO dbo.Service
                               ([PSComputerNAme]
                               ,[SystemName]
                               ,[Name]
                               ,[Suffix]
                               ,[Caption]
                               ,[DisplayName]
                               ,[PathName]
                               ,[ServiceType]
                               ,[StartMode]
                               ,[Started]
                               ,[State]
                               ,[Status]
                               ,[ExitCode]
                               ,[Description]
                               ,[DirName]
                               ,[ProgramName]
                               ,[Parameter]
                               ,[ChangeState]
                               ,[StartDate]
                               ,[CheckDate]
                               ,[ComponentID]
                               ,[ComputerID)
                            VALUES
                                ('" + $obj.ComputerName + "','"+
                                        $obj.SystemName   + "','"+
                                        $obj.Name         + "','"+
                                        $obj.Suffix       + "','"+
                                        $obj.Caption      + "','"+
                                        $obj.DisplayName  + "','"+  
                                        $obj.PathName     + "','"+
                                        $obj.ServiceType  + "','"+
                                        $obj.StartMode    + "','"+
                                        $obj.Started      + "','"+
                                        $obj.State        + "','"+
                                        $obj.Status       + "',"+
                                        $obj.ExitCode     + ",'"+
                                        $obj.Description.Replace("'","''")  + "','"+
                                        $Obj.DirName      + "','"+
                                        $obj.ProgramName  + "','"+
                                        $obj.Parameter    + "','Current','"+
                                        $timestring       + "','"+
                                        $timestring       + "'," +
                                        $componentid + "," +
                                        $computerid + ")"    
                        $DBresult = invoke-sqlcmd -ServerInstance '.\sqlexpress' -Database "Sympa" `
                            -Query "$query" `
                            -ErrorAction Stop 
                    }

                }
                else {
                    # service not yet in database, so add it
                    # Write-Host "Service not yet in database, so add it"
                    # first determine computerid
                    

                    $query = "INSERT INTO dbo.Service
                               ([PSComputerNAme]
                               ,[SystemName]
                               ,[Name]
                               ,[Suffix]
                               ,[Caption]
                               ,[DisplayName]
                               ,[PathName]
                               ,[ServiceType]
                               ,[StartMode]
                               ,[Started]
                               ,[State]
                               ,[Status]
                               ,[ExitCode]
                               ,[Description]
                               ,[DirName]
                               ,[ProgramName]
                               ,[Parameter]
                               ,[ChangeState]
                               ,[StartDate]
                               ,[CheckDate]
                               ,[ComponentID]
                               ,[ComputerID])
                            VALUES
                                ('" + $obj.ComputerName + "','"+
                                        $obj.SystemName   + "','"+
                                        $obj.Name         + "','"+
                                        $obj.Suffix       + "','"+
                                        $obj.Caption      + "','"+
                                        $obj.DisplayName  + "','"+  
                                        $obj.PathName     + "','"+
                                        $obj.ServiceType  + "','"+
                                        $obj.StartMode    + "','"+
                                        $obj.Started      + "','"+
                                        $obj.State        + "','"+
                                        $obj.Status       + "',"+
                                        $obj.ExitCode     + ",'"+
                                        $obj.Description.Replace("'","''")  + "','"+
                                        $Obj.DirName      + "','"+
                                        $obj.ProgramName  + "','"+
                                        $obj.Parameter    + "','Current','"+
                                        $timestring       + "','"+
                                        $timestring       + "'," +
                                        $componentid + "," +
                                        $computerid + ")"    
                $DBresult = invoke-sqlcmd -ServerInstance '.\sqlexpress' -Database "Sympa" `
                        -Query "$query" `
                        -ErrorAction Stop 
                                       
                }
            }
            # Set all service that have not been found anymore to OLD
            $query = "UPDATE dbo.Service
	                SET ChangeState = 'Old', CheckDate = '" + $timestring + 
                     "' WHERE SystemName = '" +$myhost + "' AND ChangeState = 'Current' AND CheckDate < '" + $timestring + "'  SELECT @@ROWCOUNT"
                    $DBresult = invoke-sqlcmd -ServerInstance '.\sqlexpress' -Database "Sympa" `
                        -Query "$query" `
                        -ErrorAction Stop 
            $del = $DBresult.Item(0)
            # Write-Host "Deleted $del"                        
        }        

    }   
    catch {
        if ($log) {
            Add-Content $logfile "==> Processing Service info failed for $myHost"
          
        }
        $scripterror = $true
        $errortext = $error[0]
        $scripterrormsg = "==> Processing Service info failed for $myHost - $errortext"
        # exit
    }
}

if (!$scripterror) {

    $query = "Select Count(*)
                From dbo.Service
                WHERE SystemName = '" + $myhost + "' AND ((ChangeState = 'Current' AND StartDate > (DATEADD(hour,-8,GETDATE())))
                                                                OR
                                                           (ChangeState = 'Old' AND CheckDate > (DATEADD(hour,-8,GETDATE()))))"  
    $DBresult = invoke-sqlcmd -ServerInstance '.\sqlexpress' -Database "Sympa" `
                            -Query "$query" `
                            -ErrorAction Stop 
    $recent8hours = $DBresult.Item(0)

    $query = "Select Count(*)
                From dbo.Service
                WHERE SystemName = '" + $myhost + "' AND ((ChangeState = 'Current' AND StartDate > (DATEADD(day,-8,GETDATE())))
                                                                OR
                                                           (ChangeState = 'Old' AND CheckDate > (DATEADD(day,-8,GETDATE()))))"  
    $DBresult = invoke-sqlcmd -ServerInstance '.\sqlexpress' -Database "Sympa" `
                            -Query "$query" `
                            -ErrorAction Stop 
    $recent8days = $DBresult.Item(0)
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

# Service status (total number of services) 
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')    
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
$NotifyChanged = $xmldoc.CreateElement('NotifyChanged')
$ValueLookup =  $xmldoc.CreateElement('ValueLookup')
    
$Channel.InnerText = "Total number of services"
$Unit.InnerText = "Custom"
$Mode.Innertext = "Absolute"
$ValueLookup.Innertext = 'ServiceStatus'

$Value.Innertext = $nrofservices

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($NotifyChanged)
[void]$Result.AppendChild($Mode)
    
[void]$PRTG.AppendChild($Result)

# Service status (changes last 8 hours) 
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')    
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
$NotifyChanged = $xmldoc.CreateElement('NotifyChanged')
$ValueLookup =  $xmldoc.CreateElement('ValueLookup')
    
$Channel.InnerText = "Service changes last 8 hours"
$Unit.InnerText = "Custom"
$Mode.Innertext = "Absolute"
$ValueLookup.Innertext = 'ServiceStatus'

$Value.Innertext = $recent8hours

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($NotifyChanged)
[void]$Result.AppendChild($Mode)
    
[void]$PRTG.AppendChild($Result)

# Service status (changes last 8 days) 
$Result = $xmldoc.CreateElement('Result')
$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')    
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
$NotifyChanged = $xmldoc.CreateElement('NotifyChanged')
$ValueLookup =  $xmldoc.CreateElement('ValueLookup')
    
$Channel.InnerText = "Service changes last 8 days"
$Unit.InnerText = "Custom"
$Mode.Innertext = "Absolute"
$ValueLookup.Innertext = 'ServiceStatus'

$Value.Innertext = $recent8days

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
    $formattime = $CheckDate.ToString("dd-MM-yyyy HH:mm:ss")
    $message = "Machine $myhost (now $livestat) *** $nrofservices Services found *** $recent8hours | $recent8days Services changed last 8 hours | days *** Timestamp: $formattime ($online) *** Script $scriptversion"
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

