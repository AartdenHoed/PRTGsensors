function Running-Elevated
{
   $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
   $p = New-Object System.Security.Principal.WindowsPrincipal($id)
   if ($p.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator))
   { Write-Output $true }      
    else
   { Write-Output $false }   
}
$status = "Ok"
$CpuList = @()
if (-not(Running-Elevated)) {
    $msg =  "*** Script NOT running as administrator"
    $status = "Error"
}
else {
    $msg =  "*** Script running as administrator"
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