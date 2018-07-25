$computername = @()
$computername = 'robot-x', 'fmmvoltron'
$Task = forEach ($Computer in $Computername)
{
	(New-Object System.Net.NetworkInformation.Ping).SendPingAsync($Computer)
}
[Threading.Tasks.Task]::WaitAll($Task)
$Task.Result

(New-Object System.Net.NetworkInformation.Ping).Send($computer)