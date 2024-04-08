# write-host "Entered VPN check2"

# write-host "Echt!"

$myURI = "https://www.expressvpn.com/nl/what-is-my-ip" 
# write-host "Net"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 
# write-host "Invoke"
try {
    $req = Invoke-Webrequest -URI $myURI -ErrorAction Stop 
}

catch {
    # write-host "Web request failed"
    $errortext = $error[0]
    throw  "$errortext"
}

# write-host "ip address"
$searchClass = "ip-address" 
$ipaddress = $req.ParsedHtml.getElementsByClassName($searchClass)
$x = $ipaddress | select innerHTML | foreach {$_.innerHTML} 
# # write-host $x

# write-host "info"
$searchClass = "info" 
$info = $req.ParsedHtml.getElementsByClassName($searchClass)
$y = $info | select innerHTML | foreach {$_.innerHTML} 
# # write-host $y

if ($x -like "*class=green*") {
    $status = 0
}
else {
    $status = 2
}
# write-host $status

$regExp = ">([^)]+)<"

if ($x -match $regexp) {
    $ip = $matches[1]
}
else {
    $ip = "Not FOund"
    $status = 3
}
# write-host $ip

if ($y -match $regexp) {
    $info = $matches[1]
}
else {
    $info = "Not FOund"
    $status = 3
} 
# write-host $info


$VpnObject = [PSCustomObject] [ordered] @{IpAddress = $ip;
                                            Info = $info;
                                            Status = $status}

# write-host "Exit"

return $VpnObject