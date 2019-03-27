##Test Remote Management##
##Created by: James Spencer##
##To Run: Test-RemoteManagement.ps1 -server servername ##

#Command line parameter to test 1 computer
param($server)

#The AD OUs to be searched for computers
$ous = "OU=OU1,DC=domain,DC=local"

#The Server description to be searched for.
$filter = "Test"

#Modify the filter for our search.
$filter="*" + $filter + "*"

#Get computers from each OU if the command line paramater is blank
if ($server -eq $null){
    foreach($ou in $ous){
        #Get the servers that match our OU and search criteria. 
        $servers=Get-ADComputer -Filter "Description -like '$filter' -and OperatingSystem -like '*Windows*'" -SearchBase $ou -Properties * | Sort-Object
        }
    }
    else {
        $servers = Get-ADComputer -Identity $server -Properties *
        }

foreach($server in $servers){
    Write-Host $server.name $server.description
    $icmp = Test-Connection -ComputerName $server.name -Count 1 -Quiet
    Write-Host "ICMP Test: $icmp"
    $rdp = Test-NetConnection -ComputerName $server.name -CommonTCPPort RDP -ErrorAction SilentlyContinue -InformationLevel Quiet
    Write-Host "RDP Test: $rdp"
    $wrm = Test-NetConnection -ComputerName $server.Name -CommonTCPPort WINRM -ErrorAction SilentlyContinue -InformationLevel Quiet
    Write-Host "WinRM (TCP: 5985) Test: $wrm"
    $smb = Test-NetConnection -ComputerName $server.name -CommonTCPPort SMB -ErrorAction SilentlyContinue -InformationLevel Quiet
    Write-Host "SMB (TCP: 445) Test: $smb"
    $dcom = Test-NetConnection -ComputerName $server.name -Port 135 -ErrorAction SilentlyContinue -InformationLevel Quiet
    Write-host "DCOM (TCP: 135) Test: $dcom"
    Try{
        $wmi = Get-WmiObject -ComputerName RADFC2A13V12099 -Class win32_computersystem -ErrorAction Stop
        Write-Host "WMI Test: True"
        }
    Catch {
        Write-host "WMI Test: False"
        }
    Write-Host "-------------------------------------------------"
    }

#Cleanup
Clear-Variable -name server*
