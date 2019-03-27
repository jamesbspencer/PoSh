#Another script by James Spencer

$bases="OU=OU1,DC=domain,DC=local","OU=OU2,DC=domain,DC=local","OU=OU3,DC=domain,DC=local"

$groups="Group1","Group2","Group3" | Sort-Object

$srvdata = $null
$custdata = $null
$warns = 0
$crits = 0
$date = Get-Date

foreach ($group in $groups){
    Write-Host $group
    $filter="*" + $group + "*"
    foreach($base in $bases){
        $servers=Get-ADComputer -Filter "Description -like '$filter' -and OperatingSystem -like '*Windows*'" -SearchBase $base -Properties * | Sort-Object
        foreach($server in $servers){
            $name=$server.Name.ToUpper()
            $desc=$server.Description
            $os=$server.OperatingSystem #-replace "Microsoft Windows Server",""
                Try{
                    $icmp=Test-Connection -ComputerName $name -Count 1 -ErrorAction Stop
                    Try{
                        $net = Get-WmiObject win32_networkadapterconfiguration -ComputerName $name -Filter IPEnabled=True -ErrorAction Stop
                        $ip1=($net.IPAddress -split "\r\n")[0]
                        $ip2=($net.IPAddress -split "\r\n")[1]
                        $status="Good"
                        Write-Host $name $desc $os $ip1 $ip2 $status -ForegroundColor Green
                        }
                    Catch{
                        $ip1=$icmp.IPV4Address.IPAddressToString
                        $ip2=""
                        $status="WMI"
                        write-host $name $desc $os $ip1 $ip2 $status -ForegroundColor Yellow
                        }
                }
                Catch{
                    $ip1="0.0.0.0"
                    $ip2=""
                    $status="DNS"
                    Write-Host $name $desc $os $ip1 $status $icmp.StatusCode -ForegroundColor Red
                    }
            if($name -match "RADFGF*"){
                $rdp = 
@"
<img src="./extra/rdp.png" height="20px" width="20px" alt="RDP" title="Remote Desktop Connection" onclick="rdp2('$name');"/>
"@
                }
            else{
                $rdp =
@"
<img src="./extra/rdp.png" height="20px" width="20px" alt="RDP" title="Remote Desktop Connection" onclick="rdp1('$name');"/>
"@
                } #End server if/else
            #Create the MMC links
            $mmc =
@"
<img src="./extra/mmc.png" height="20px" width="20px" alt="CompMgmt" title="Computer Management" onclick="mmc('$name');"/>
"@
            #Create the admin share links
            $admin = 
@"
<img src="./extra/admin.png" height="20px" width="20px" alt="C$" title="C Share" onclick="admin('$name');"/>
"@
            #Create the server table
            if ($status -eq "1"){
                $crits++
                $srvdata +=
@"
                <tr style="background:red;" id="$name">
                    <td>$name <br/>$ip1 $ip2</td>
                    <td>$desc</td>
                    <td>$os</td>
                    <td>$status - ICMP</td>
                    <td>$rdp $mmc $admin</td>
                </tr>
            
"@
                }
            elseif ($status -eq "DNS"){
                $crits++
                $srvdata +=
@"
                <tr style="background:red;" id="$name">
                    <td>$name <br/>$ip1 $ip2</td>
                    <td>$desc</td>
                    <td>$os</td>
                    <td>DNS/Down</td>
                    <td>$rdp $mmc $admin</td>
                </tr>
            
"@
                }
            elseif ($status -eq "WMI"){
                $warns++
                $srvdata +=
@"
                <tr style="background:yellow;" id="$name">
                    <td>$name <br/>$ip1 $ip2</td>
                    <td>$desc</td>
                    <td>$os</td>
                    <td>WMI</td>
                    <td>$rdp $mmc $admin</td>
                </tr>
            
"@
                }
            else{
$srvdata +=
@"
                <tr id="$name">
                    <td>$name<br/>$ip1 $ip2</td>
                    <td>$desc</td>
                    <td>$os</td>
                    <td>UP</td>
                    <td>$rdp $mmc $admin</td>
                </tr>
            
"@
                } #End if server status      
            } #END foreach server
        } #End foreach base
    $custdata +=
@"
        <table style="width:100%" class="table1">
            <caption><b>$group</b></caption>
            <tr>
                <th>Server Name</th>
                <th>Description</th>
                <th>OS</th>
                <th>Status</th>
                <th>Links</th>
            </tr>
            $srvdata
        </table>
        <br>
"@
    $srvdata = $null
    
    } #End foreach group

#How long did it take to get this far
$elapsedTime =($(get-date) - $date).seconds
#Generate the HTML Page
$out =
@"
<html>
    <head>
        <title>ERP Windows Server Status v3.0</title>
        <meta http-equiv="x-ua-compatible" content="ie=9">
        <meta http-equiv="refresh" content="900">
		<link rel="stylesheet" type="text/css" href="./extra/style.css">
        <HTA:APPLICATION
			border="thin"
			borderStyle="normal"
			caption="yes"
			icon="./favicon.ico"
			maximizeButton="yes"
			minimizeButton="yes"
			showInTaskbar="yes"
			windowState="normal"
			innerBorder="yes"
			navigable="yes"
			scroll="auto"
			scrollFlat="yes" />
        <script type="text/javascript" language="javascript">
        function rdp1(a) {
			if (a === undefined){
				return;
				}
			var cmd = "c:/windows/system32/mstsc.exe /v:"+a+":49800";
			WshShell = new ActiveXObject("WScript.Shell");
			WshShell.Run(cmd, 1, false);
			}
        function rdp2(b) {
			if (b === undefined){
				return;
				}
			var cmd = "c:/windows/system32/mstsc.exe /v:"+b+":3389";
			WshShell = new ActiveXObject("WScript.Shell");
			WshShell.Run(cmd, 1, false);
			}
        function mmc(c) {
			if (c === undefined){
				return;
				}
			var cmd = "c:/windows/system32/compmgmt.msc -a /computer="+c;
			WshShell = new ActiveXObject("WScript.Shell");
			WshShell.Run(cmd, 1, false);
			}
        function admin(d) {
			if (d === undefined){
				return;
				}
			var cmd = "file://"+d+"/c$";
			WshShell = new ActiveXObject("WScript.Shell");
			WshShell.Run(cmd, 1, false);
			}
        </script>
    </head>
    <body>
        <table align="center" class="table1">
			<tr>
				<th>Warnings</th>
				<th>Critical</th>
			</tr>
			<tr>
				<td style="background:yellow;" align="center">$warns</td>
				<td style="background:red;" align="center">$crits</td>
			</tr>
		</table>
        <h3>ERP Windows Server Status v3.0</h3>
        <p>Page Generated on $date and took $elapsedTime seconds to create.</p>

        <br>
        $custdata
    </body>
</html>
"@
$out | Out-File -FilePath "$PSScriptRoot\index.hta" -Force
Write-Host "Finished creating main page."
