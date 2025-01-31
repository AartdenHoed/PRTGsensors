param (
    [string]$LOGGING = "YES", 
    [string]$myHost  = "NONE" ,
    [int]$sensorid = 77 
)
# $LOGGING = 'YES'
# $myHost = "ADHC-2"

$myhost = $myhost.ToUpper()

$ScriptVersion = " -- Version: 5.3.4"

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

$resultobj = [PSCustomObject] [ordered] @{deleted= 0;
                                          added = 0;
                                          readded = 0;
                                          updated = 0;
                                          changed = 0;
                                          current = 0;
                                          illegal = 0;
                                          totalactivedb = 0;
                                          totalactivewmic = 0}

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

# Get componentID of Unknown
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
    if ($invokable) {
        if ($log) {
            Add-Content $logfile "==> Node is up, put realtime service info in object"
        }
        $objlist = @()
        $resultobj.TotalActiveWmic = $serviceinfo.Count
            
        foreach ($service in $ServiceInfo) {
            # Determine program name without arguments
            if ($service.Description -eq $null) {
                $service.Description = "n/a"
            }
            # Write-Host $service.Name
            $thispath = $service.PathName
            if ($service.Name -Match "\w+_[0-9a-fA-F]{5,6}") {
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

                # Guess program name from directory name
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
                                                    CheckDate      = $timestring;
                                                    ChangeState    = "?"}  
            $objlist += $obj 
        }  
    } 
    else {
        if ($log) {
            Add-Content $logfile "==> Node is down, get info from SQL database"
        }       
    }

}

if ((!$scripterror) -and ($invokable)) {
    try {
        if ($log) {
            Add-Content $logfile "==> Update SQL database with realtime info"
        }   
        foreach ($item in $objlist) {                  
                 
            # Update record database if already present
            $query = "SELECT * FROM dbo.Service WHERE SystemName = '" + $item.SystemName + 
                        "' AND Name = '" + $item.Name + "'"
            $DBresult = invoke-sqlcmd -ServerInstance '.\sqlexpress' -Database "Sympa" `
                    -Query "$query" `
                    -ErrorAction Stop
            if ($DBresult -ne $null) {                            
                
                if ($DBresult.Parameter -is [DBNull]) {
                    $cparm = ''
                }
                else {
                    $cparm = $DBresult.Parameter
                }
                if ($DBresult.Dirname.Trim() -eq $item.DirName.Trim() -and 
                        $DBresult.ProgramName.Trim() -eq $item.ProgramName.Trim() -and 
                        $cparm.Trim() -eq $item.Parameter.Trim()) {
                    $changed = $false
                    $updated = $false
                }
                else {
                    if ($DBresult.DirectoryTemplate -is [DBNull]) {
                        $DBresult.DirectoryTemplate = ' '
                    }
                    if (($DBresult.DirectoryTemplate -ne $null) -and ($DBresult.DirectoryTemplate.Trim() -ne '')) {
                        $matchpattern = $DBresult.DirectoryTemplate.Trim() + "$"                                         
                        if ($item.DirName.Trim() -match $matchpattern) {
                            $changed = $false
                            $updated = $true
                        }
                        else {
                            $changed = $true
                            $updated = $false
                        }
                    }
                    else {
                        $changed = $true
                        $updated = $false
                    } 
                    
                }               
                if ($changed -or $updated) {
                    # Service Changed
                    $olddirname = $DBresult.DirName
                    $oldprogramname = $DBresult.ProgramName
                    $oldparameter = $DBresult.Parameter
                    if ($changed) {
                        $newstatus = "Changed"
                    }
                    else {
                        $newstatus = "Updated"
                    }
                    $oldstatus = $DBresult.ChangeState  
                    $startdate = $timestring                             
                }
                else {   
                    # if service NOT changed or updated
                    $olddirname = $DBresult.OldDirNAme
                    $oldprogramname = $DBresult.OldProgramName
                    $oldparameter = $DBresult.OldParameter
                    $startdate = $DBresult.StartDate 
                    
                    switch ($DBresult.ChangeState.Trim()) {
                        "Current" {
                            $newstatus = "Current"
                            $oldstatus = $DBresult.OldChangeState                             
                        }
                        "Deleted" {
                            $newstatus = "ReAdded"                            
                            $oldstatus = "Deleted" 
                        } 
                        "Changed" {    
                            $startdate = $DBresult.StartDate                        
                            if ($DBresult.StartDate -lt $CheckDate.AddHours(-32)) {
                                $newstatus = "Current" 
                                $oldstatus = "Changed"                                
                            }  
                            else {
                                $newstatus = "Changed" 
                                $oldstatus = $DBresult.OldChangeState
                            }                             
                       
                        }
                        "Updated" {    
                            $startdate = $DBresult.StartDate                        
                            if ($DBresult.StartDate -lt $CheckDate.AddHours(-32)) {
                                $newstatus = "Current" 
                                $oldstatus = "Updated"                                
                            }  
                            else {
                                $newstatus = "Updated" 
                                $oldstatus = $DBresult.OldChangeState
                            }                             
                       
                        }
                        "Added" {
                            $startdate = $DBresult.StartDate                        
                            if ($DBresult.StartDate -lt $CheckDate.AddHours(-32)) {
                                $newstatus = "Current" 
                                $oldstatus = "Added"                                
                            }  
                            else {
                                $newstatus = "Added" 
                                $oldstatus = $DBresult.OldChangeState
                            }           
                        }
                        "ReAdded" {
                            $startdate = $DBresult.StartDate                        
                            if ($DBresult.StartDate -lt $CheckDate.AddHours(-32)) {
                                $newstatus = "Current" 
                                $oldstatus = "ReAdded"                                
                            }  
                            else {
                                $newstatus = "ReAdded" 
                                $oldstatus = $DBresult.OldChangeState
                            }           
                        }
                        Default {
                            $scripterrormsg = "==> Database failure: " + $DBresult.ChangeState + " is an invalid CHANGESTATE"
                            if ($log) {
                                Add-Content $logfile $scripterrormsg
          
                            }
                            $scripterror = $true
                            throw $scripterrormsg
                        }
                    }
                }              
                
                     
                # service found, actualize it
                $query = "UPDATE dbo.Service
                        SET [PSComputerNAme] = '" + $item.ComputerName + "'
                            ,[Caption] = '"        + $item.Caption      + "'
                            ,[Suffix] = '"         + $item.Suffix       + "'
                            ,[DisplayName] = '"    + $item.DisplayName  + "'  
                            ,[PathName] = '"       + $item.PathName     + "'
                            ,[ServiceType] = '"    + $item.ServiceType  + "'
                            ,[StartMode] = '"      + $item.StartMode    + "'
                            ,[Started] = '"        + $item.Started      + "'
                            ,[State] = '"          + $item.State        + "'
                            ,[Status] = '"         + $item.Status       + "'
                            ,[ExitCode] = "        + $item.ExitCode     + "
                            ,[Description] = '"    + $item.Description.Replace("'","''")  + "'
                            ,[Software] = '"       + $item.SOftware     + "'
                            ,[DirName] = '"        + $item.DirName      + "'
                            ,[ProgramName] = '"    + $item.ProgramName  + "'
                            ,[Parameter] = '"      + $item.Parameter    + "'
                            ,[OldDirName] = '"     + $olddirname      + "'
                            ,[OldProgramName] = '" + $oldprogramName  + "'
                            ,[OldParameter] = '"   + $oldparameter    + "'
                            ,[OldCHangeState] = '" + $oldstatus       + "'
                            ,[CheckDate] = '"      + $timestring       + "'
                            ,[StartDate] = '"      + $startdate        + "'
                            ,[ChangeState] = '"    + $newstatus        + 
                            "' WHERE SystemName = '" + $item.SystemName + 
                                "' AND Name = '" + $item.Name + "'"   
                $DBresult = invoke-sqlcmd -ServerInstance '.\sqlexpress' -Database "Sympa" `
                    -Query "$query" `
                    -ErrorAction Stop  
                   
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
                            ,[Software]
                            ,[DirName]
                            ,[ProgramName]
                            ,[Parameter]
                            ,[ChangeState]
                            ,[StartDate]
                            ,[CheckDate]
                            ,[ComponentID]
                            ,[ComputerID])
                        VALUES
                            ('" + $item.ComputerName + "','"+
                                    $item.SystemName   + "','"+
                                    $item.Name         + "','"+
                                    $item.Suffix       + "','"+
                                    $item.Caption      + "','"+
                                    $item.DisplayName  + "','"+  
                                    $item.PathName     + "','"+
                                    $item.ServiceType  + "','"+
                                    $item.StartMode    + "','"+
                                    $item.Started      + "','"+
                                    $item.State        + "','"+
                                    $item.Status       + "',"+
                                    $item.ExitCode     + ",'"+
                                    $item.Description.Replace("'","''")  + "','"+
                                    $item.Software     + "','"+
                                    $item.DirName      + "','"+
                                    $item.ProgramName  + "','"+
                                    $item.Parameter    + "','Added','"+
                                    $timestring       + "','"+
                                    $timestring       + "'," +
                                    $componentid + "," +
                                    $computerid + ")"    
            $DBresult = invoke-sqlcmd -ServerInstance '.\sqlexpress' -Database "Sympa" `
                    -Query "$query" `
                    -ErrorAction Stop 
                                       
            }
        } # end of foreach
        # Set all service that have not been found anymore to OLD
        $query = "UPDATE dbo.Service
	            SET ChangeState = 'Deleted', CheckDate = '" + $timestring + 
                    "' WHERE SystemName = '" + $myhost + "' AND CheckDate < '" + $timestring + "'  SELECT @@ROWCOUNT"
                $DBresult = invoke-sqlcmd -ServerInstance '.\sqlexpress' -Database "Sympa" `
                    -Query "$query" `
                    -ErrorAction Stop 
        $resultobj.deleted = $DBresult.Item(0)
        # Write-Host "Deleted $del" 
    }
    catch {
        if ($log) {
            Add-Content $logfile "==> Database processing failed for $myHost"
          
        }
        $scripterror = $true
        $errortext = $error[0]
        $scripterrormsg = "==> Database processing failed for $myHost - $errortext"
        # exit
    }
}                                  
  
if (!$scripterror) {
    try {     

        # get totals from database
        $query = "Select Count(*)
                        From dbo.Service
                    WHERE SystemName = '" + $myhost + "' AND ComponentID = " + $componentid   
        $DBresult = invoke-sqlcmd -ServerInstance '.\sqlexpress' -Database "Sympa" `
                    -Query "$query" `
                    -ErrorAction Stop 
        $resultobj.Illegal = $DBresult.Item(0)

        $query = "Select Count(*)
                        From dbo.Service
                    WHERE SystemName = '" + $myhost + "' AND ChangeState = 'Changed'"  
        $DBresult = invoke-sqlcmd -ServerInstance '.\sqlexpress' -Database "Sympa" `
                    -Query "$query" `
                    -ErrorAction Stop 
        $resultobj.Changed = $DBresult.Item(0)

        $query = "Select Count(*)
                        From dbo.Service
                    WHERE SystemName = '" + $myhost + "' AND ChangeState = 'Updated'"  
        $DBresult = invoke-sqlcmd -ServerInstance '.\sqlexpress' -Database "Sympa" `
                    -Query "$query" `
                    -ErrorAction Stop 
        $resultobj.Updated = $DBresult.Item(0)

        $query = "Select Count(*)
                    From dbo.Service
                    WHERE SystemName = '" + $myhost + "' AND ChangeState = 'Added'"  
        $DBresult = invoke-sqlcmd -ServerInstance '.\sqlexpress' -Database "Sympa" `
                    -Query "$query" `
                    -ErrorAction Stop 
        $resultobj.Added = $DBresult.Item(0)

        $query = "Select Count(*)
                    From dbo.Service
                    WHERE SystemName = '" + $myhost + "' AND ChangeState = 'ReAdded'"  
        $DBresult = invoke-sqlcmd -ServerInstance '.\sqlexpress' -Database "Sympa" `
                    -Query "$query" `
                    -ErrorAction Stop 
        $resultobj.ReAdded = $DBresult.Item(0)   
        
        $query = "Select Count(*)
                    From dbo.Service
                    WHERE SystemName = '" + $myhost + "' AND ChangeState = 'Current'"  
        $DBresult = invoke-sqlcmd -ServerInstance '.\sqlexpress' -Database "Sympa" `
                    -Query "$query" `
                    -ErrorAction Stop 
        $resultobj.Current = $DBresult.Item(0)  
        
        $resultobj.totalactivedb =  $resultobj.Current + $resultobj.Added + $resultobj.ReAdded + $resultobj.Changed   
        
        if (($result.totalactivedb -ne $result.totalactivewmic) -and ($result.totalactivewmic -ne 0)) {
            throw "Totals of DB and WMIC do not match !"
        }    

    }  
    catch {
        if ($log) {
            Add-Content $logfile "==> Reading totals out of database failed for $myHost"          
        }
        $scripterror = $true
        $errortext = $error[0]
        $scripterrormsg = "==> Reading totals out of database failed for $myHost - $errortext"
        # exit    
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

# Service status (total number of services) 
$Result = $xmldoc.CreateElement('Result')

$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')    
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
    
$Channel.InnerText = "Total number of active services"
$Value.Innertext = $resultobj.totalactivedb
$Unit.InnerText = "Custom"
$CustomUnit.Innertext = "Services"
$Mode.Innertext = "Absolute"

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
[void]$Result.AppendChild($Mode)
    
[void]$PRTG.AppendChild($Result)

# Service status (total number of current services) 
$Result = $xmldoc.CreateElement('Result')

$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')    
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
    
$Channel.InnerText = "Total number of unchanged services"
$Value.Innertext = $resultobj.current
$Unit.InnerText = "Custom"
$CustomUnit.Innertext = "Services"
$Mode.Innertext = "Absolute"

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
[void]$Result.AppendChild($Mode)
    
[void]$PRTG.AppendChild($Result)

# Service status (total number of changed services) 
$Result = $xmldoc.CreateElement('Result')

$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')    
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
$LimitMode = $xmldoc.CreateElement('LimitMode')
$LimitMinError = $xmldoc.CreateElement('LimitMinError')
$LimitMaxError = $xmldoc.CreateElement('LimitMaxError')
    
$Channel.InnerText = "Total number of changed services"
$Value.Innertext = $resultobj.changed
$Unit.InnerText = "Custom"
$CustomUnit.Innertext = "Services"
$Mode.Innertext = "Absolute"
$LimitMode.InnerText = "1"
$LimitMinError.InnerText = "0"
$LimitMaxError.InnerText = "0"

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
[void]$Result.AppendChild($Mode)
[void]$Result.AppendChild($LimitMode) 
[void]$Result.AppendChild($LimitMinError) 
[void]$Result.AppendChild($LimitMaxError)     
    
[void]$PRTG.AppendChild($Result)

# Service status (total number of updated services) 
$Result = $xmldoc.CreateElement('Result')

$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')    
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
$LimitMode = $xmldoc.CreateElement('LimitMode')
$LimitMinError = $xmldoc.CreateElement('LimitMinError')
$LimitMaxError = $xmldoc.CreateElement('LimitMaxError')
    
$Channel.InnerText = "Total number of updated services"
$Value.Innertext = $resultobj.updated
$Unit.InnerText = "Custom"
$CustomUnit.Innertext = "Services"
$Mode.Innertext = "Absolute"

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
[void]$Result.AppendChild($Mode)
    
[void]$PRTG.AppendChild($Result)

# Service status (total number of Added services) 
$Result = $xmldoc.CreateElement('Result')

$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')    
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
$LimitMode = $xmldoc.CreateElement('LimitMode')
$LimitMinError = $xmldoc.CreateElement('LimitMinError')
$LimitMaxError = $xmldoc.CreateElement('LimitMaxError')
    
$Channel.InnerText = "Total number of added services"
$Value.Innertext = $resultobj.added
$Unit.InnerText = "Custom"
$CustomUnit.Innertext = "Services"
$Mode.Innertext = "Absolute"
$LimitMode.InnerText = "1"
$LimitMinError.InnerText = "0"
$LimitMaxError.InnerText = "0"

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
[void]$Result.AppendChild($Mode)
[void]$Result.AppendChild($LimitMode) 
[void]$Result.AppendChild($LimitMinError) 
[void]$Result.AppendChild($LimitMaxError)     
    
[void]$PRTG.AppendChild($Result)

# Service status (total number of re-added services) 
$Result = $xmldoc.CreateElement('Result')

$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')    
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
    
$Channel.InnerText = "Total number of re-added services"
$Value.Innertext = $resultobj.readded
$Unit.InnerText = "Custom"
$CustomUnit.Innertext = "Services"
$Mode.Innertext = "Absolute"

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
[void]$Result.AppendChild($Mode)
    
[void]$PRTG.AppendChild($Result)

# Service status (total number of deleted services) 
$Result = $xmldoc.CreateElement('Result')

$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')    
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
    
$Channel.InnerText = "Total number of deleted services"
$Value.Innertext = $resultobj.deleted
$Unit.InnerText = "Custom"
$CustomUnit.Innertext = "Services"
$Mode.Innertext = "Absolute"

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
[void]$Result.AppendChild($Mode)
    
[void]$PRTG.AppendChild($Result)

# Service status (total number of illegal services (linked to unknown component) 
$Result = $xmldoc.CreateElement('Result')

$Channel = $xmldoc.CreateElement('Channel')
$Value = $xmldoc.CreateElement('Value')    
$Unit = $xmldoc.CreateElement('Unit')
$CustomUnit = $xmldoc.CreateElement('CustomUnit')
$Mode = $xmldoc.CreateElement('Mode')
    
$Channel.InnerText = "Total number of illegal services"
$Value.Innertext = $resultobj.illegal
$Unit.InnerText = "Custom"
$CustomUnit.Innertext = "Services"
$Mode.Innertext = "Absolute"

[void]$Result.AppendChild($Channel)
[void]$Result.AppendChild($Value)
[void]$Result.AppendChild($Unit)
[void]$Result.AppendChild($CustomUnit)
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
    $nrofservices = $resultobj.TotalACtiveDB
    $message = "Machine $myhost (now $livestat) *** $nrofservices Active Services found *** Timestamp: $formattime ($online) *** Script $scriptversion"
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

