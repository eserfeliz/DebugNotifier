[CmdletBinding()]
Param(
    [Parameter(Mandatory=$True,Position=1)][string]$email,
    $computername = @(hostname),
    $debuggers = @('windbg.exe','cdb.exe','werfault.exe')
)

$workstations = New-Object System.Collections.ArrayList
$servers = New-Object System.Collections.ArrayList

Function wait([int]$time)
{
	timeout -t $time;
}

Function pause()
{
    Write-Host "Press any key to continue ..."
    $null = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

Function Set-MachineTypes
{
    foreach ($machine in $computername)
    {
        $err = Check-WMIConnectivity($machine)
        $err
        if ($err -ne $null)
        {
            $workstations.Add($machine) | Out-Null
        }
        else
        {
            $servers.Add($machine) | Out-Null
        }
    }
    $info = New-Object psobject -Property @{
        servers = $servers
        workstations = $workstations
    }
    return $info
}

Function Check-WMIConnectivity($machine)
{
    gwmi Win32_Process -ComputerName $machine -ErrorVariable err -ErrorAction SilentlyContinue | Out-Null
    if ($err)
    {
        return $err
    }
}

Function Check-DebuggerInstance($machine)
{
    gwmi Win32_Process -ComputerName $machine -ErrorVariable err -ErrorAction SilentlyContinue | Where-Object {$debuggers -contains $_.Name}
}

Function Check-Servers
{
    $servers | %{Check-DebuggerInstance($_)}
}

Function Output-Info($debugcheck)
{
    $debugcheck.CommandLine -match " (\d+)" | Out-Null
    $trappedpid = $matches[1]
    Write-Warning "A trap has been detected on $($debugcheck.PSComputerName)"
    Write-Host "Windbg Process ID = $($debugcheck.ProcessId)"
    $MailMessage += "`nWindbg Process ID = $($debugcheck.ProcessId). `r`n"
    Write-Host "Debugger Command line = $($debugcheck.CommandLine)"
    $MailMessage += "Debugger Command line = $($debugcheck.CommandLine). `r`n"
    Write-Host "Trapped Process ID = $trappedpid"
    $MailMessage += "Trapped Process ID = $($trappedpid). `r`n"
    Write-Host "Trapped Process Name = $((Get-Process -ComputerName $debugcheck.PSComputerName | where-object {$_.Id -eq $trappedpid}).ProcessName)"
    $MailMessage += "Trapped Process Name = $((Get-Process -ComputerName $debugcheck.PSComputerName | where-object {$_.Id -eq $trappedpid}).ProcessName). `r`n"
    Send-Mail($MailMessage) | Out-Null
}

Function Send-Mail($MailMessage)
{
    $smtpServer = "mail.com"
    $serverName = hostname
    $time = Get-Date

    Write-Host "qwinsta = $qw"

    $msg = new-object Net.Mail.MailMessage
    $smtp = new-object Net.Mail.SmtpClient($smtpServer)
    $msg.From = "fromTest@stress.local"
    $msg.ReplyTo = "noreply@sys2.local"
    Write-Host "$email"
    $msg.To.Add("$email")
    $msg.subject = "Error has been detected at $($debugcheck.PSComputerName)";
    $msg.body = "A trap has been detected on server: $($debugcheck.PSComputerName) at $($time).`n`n`nprocess info: $MailMessage"

    if ($psexec = (Get-ChildItem -Path \*\*\* -Filter psexec.exe).FullName[0])
    {
        $remotemachine = "\\$($debugcheck.PSComputerName)"
        $command = "qwinsta"
        $switch = "/counter"
        $qw = &$psexec $remotemachine $command $switch
        $qw = $qw | Out-String
        Write-Host "qwinsta = `n$qw"
        $msg.body += "`n`n`n qwinsta information:"
        $msg.body += "`n $($qw)"
    }
    else
    {
        Write-Host "Cannot find psexec on this machine to run qwinsta."
        $msg.body += "`n`n`nCannot find psexec to gather qwinsta info."
    }
    

    Write-Host "Sending notification email"
    $smtp.Send($msg)
}

$machines = Set-MachineTypes
while ($servers -ne $null)
{
    $debugcheck = Check-Servers
    if ($debugcheck -ne $null -and $servers -ne $null)
    {
        Output-Info($debugcheck)
        $servers = $servers -ne $debugcheck.PSComputerName
        $debugcheck = $null
        wait 5
    }
    else
    {
        Write-Host -ForegroundColor Green "No traps found"
        wait 30
    }
}
