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
                                           Administrator = $adm;
                                           Version="1.0"}
    return $MyAuth 
 } 
$status = "Ok"
$CpuList = @()
$el = Running-Elevated

$u = $el.id.Name
$ia = $el.Principal.Identity.IsAuthenticated
$v = $el.Version

if (-not($el.Administrator)) {
    $msg =  "*** Script (version $v) NOT running as administrator (Current user is $u (IsAuthenticated = $ia))"
    $status = "Error"
}
else {
    $msg =  "*** Script (version $v) running as administrator (Current user is $u (IsAuthenticated = $ia))"
} 

if ($status -eq "Ok") {
 
    Add-Type -Path "D:\Software\OpenHardwareMonitor\OpenHardwareMonitorLib.dll"
 
    $Comp = New-Object -TypeName OpenHardwareMonitor.Hardware.Computer
 
    $Comp.Open()
 
    $Comp.CPUEnabled = $true
 
    $Comp.RAMEnabled = $true
 
    $Comp.MainboardEnabled = $true
 
    $Comp.FanControllerEnabled = $true
 
    $Comp.GPUEnabled = $true
 
    $Comp.HDDEnabled = $true


    $templist = @() 
    $packagefound = $false
    ForEach ($HW in $Comp.Hardware) {
 
        $HW.Update()    
 
        If ( $hw.HardwareType -eq "CPU"){
            $type = $hw.name.ToString()
        
            ForEach ($Sensor in $HW.Sensors) {
 
                If ($Sensor.SensorType -eq "Temperature"){
                    $name =  $Sensor.Name 
                    $tempcurrent = $Sensor.Value.ToString()
                    $tempmin =$Sensor.Min.ToString()
                    $tempmax = $Sensor.Max.ToString()
            
        
                    $obj = [PSCustomObject] [ordered] @{Type = $type;
                                                        Name = $name;
                                                        TempCurrent = $tempcurrent;
                                                        TempMin = $tempmin;
                                                        TempMax = $tempmax}

                    $templist += $obj
                }

                if ($name -eq "CPU Package") {
                    $ptype = $type
                    $pname = $name
                    $ptempcurrent = $tempcurrent
                    $ptempmin = $tempmin
                    $ptempmax = $tempmax
                    $packagefound = $true
                }
            } 
                                            

        }
    
    }
    $Comp.Close()

   
    if ($packagefound) {
        $obj = [PSCustomObject] [ordered] @{Type = $ptype;
                                                    Name = $pname;
                                                    TempCurrent = $ptempcurrent;
                                                    TempMin = $ptempmin;
                                                    TempMax = $ptempmax}
        $CpuList += $obj
    }

    foreach ($c in $templist) {
        if ($c.Name -ne "CPU Package") {
            $CpuList += $c
        
        }
    }
}
$Global:ReturnObject = [PSCustomObject] [ordered] @{Message = $msg;
                                            MyStatus = $status;
                                            CPUlist = $CpuList}

Return $Global:ReturnObject