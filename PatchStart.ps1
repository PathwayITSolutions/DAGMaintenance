<#
    .SYNOPSIS
        Automated Exchange DAG Maintenance and Patching v4.2

    .DESCRIPTION
        The first half of the Automated Exchange DAG Maintenance and Patching script team,
        putting the server in Maintenance mode, Patching and if needed rebooting the server.        

        This script needs no inputs, as it will read the server name, the DAG details, and
        the required FQDN from the environment.
        It requires An Exchange window in Admin mode so that it can address the cluster.

    .EXAMPLE
        .\PatchStart.ps1

    .LINK
        https://www.pathwayit.net

    .NOTES
        Work in Progress, several items in the TODO
        TODO: Suspend-ClusterNode should read the output and act on it
        TODO: Message-Redirect should read the output and act on it
#>
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if ($currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host -ForegroundColor Green "Admin mode, continuing"
}
else {
    Write-Host -NoNewLine -ForegroundColor Red "Non-Admin mode, exiting, Press any key to continue..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit
}

if (Get-Module -ListAvailable -Name PSWindowsUpdate) {
    Write-Host "PSWindowsUpdate module installed, updating if possible"
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    Update-Module -Name PSWindowsUpdate
    Write-Host -ForegroundColor Green "PSWindowsUpdate module on latest version"
}
else {
    Write-Host -ForegroundColor Red "PSWindowsUpdate module not installed, exiting, Press any key to continue..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    Write-Host "You can install PSWindowsUpdate with 'Install-Module -Name PSWindowsUpdate'. You may need to 'Set-PSRepository -Name PSGallery -InstllationPolicy Trusted' as well."
    exit
}

Write-Host "Checking Environment..."
$ThisServer = $env:computername
$ThisEnvDAG = 1
$ThisEnvDAG = Get-DatabaseAvailabilityGroup
if ($ThisEnvDAG -eq 1) {
    Write-Host -ForegroundColor Red "Not a DAG Environment, exiting, Press any key to continue..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    Write-Host "This script is for DAG members only. No DAG was detected in this environment. If you feel this is incorrect, please contact support@pathwayit.net"
    exit
}
Write-Host "Looking for a Redirect Server..."
Foreach ($DAG in $ThisEnvDAG) {
    if ($DAG.Servers.Name -contains $ThisServer) { $DAGSet = $DAG.Servers }
}
$RedirectTarget = 1
While ($RedirectTarget -eq 1 ) {
    Foreach ($Server in $DAGSet.Name) {
        $RTHTState = Get-ServerComponentState $Server -Component HubTransport
        if ((Test-Connection $Server -Quiet) -and ($RTHTState.State -eq "Active") -and ($Server -ne $ThisServer)) {
            $RedirectTarget = ([System.Net.Dns]::GetHostByName("$Server")).HostName
            Write-Host -ForegroundColor Green "Redirect Server Found"
            break
        }
    }
    if ($RedirectTarget -eq 1) {
        $RedirectTarget = 2
    }
}
if ($RedirectTarget -eq 2) {
    Write-Host -NoNewLine -ForegroundColor Red "No suitable target for Redirect-Message, exiting, Press any key to continue..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit
}
Write-Host "Moving Databases..."
Move-ActiveMailboxDatabase -Server $ThisServer -ActivateOnServer $RedirectTarget -MoveAllDatabasesOrNone -SkipClientExperienceChecks -Confirm:$false -ErrorAction Stop
Write-Host "Databases Moved"

Write-Host "Set HubTransport Draining..."
Set-ServerComponentState -Identity $ThisServer -Component HubTransport -State Draining -Requester Maintenance
Start-Sleep -Seconds 5
$HTState = Get-ServerComponentState $ThisServer -Component HubTransport
$HTCounter = 0
While (($HTState.State -ne "Inactive") -and ($HTCounter -lt 5)) {
    Set-ServerComponentState -Identity $ThisServer -Component HubTransport -State Draining -Requester Maintenance
    Start-Sleep -Seconds 5
    $HTState = Get-ServerComponentState $ThisServer -Component HubTransport
    $HTCounter++
}
if ($HTState.State -eq "Inactive") {
    Write-Host -ForegroundColor Green "HubTransport Draining Set"
}
else {
    Write-Host -ForegroundColor Red "HubTransport Draining attempted but not confirmed"
}

Write-Host "Messages Redirecting..."
Redirect-Message -Server $ThisServer -Target "$RedirectTarget" -Confirm:$False | Out-Null

Write-Host "Suspending Cluster Node..."
Suspend-ClusterNode $ThisServer | Out-Null
Write-Host -ForegroundColor Green "Cluster Node Suspended"

Write-Host "Disabling and Blocking DatabaseCopy Activation..."
Set-MailboxServer $ThisServer -DatabaseCopyActivationDisabledAndMoveNow $true
Set-MailboxServer $ThisServer -DatabaseCopyAutoActivationPolicy Blocked
Start-Sleep -Seconds 2
$MBXServer = Get-MailboxServer $ThisServer
$MBXCounter = 0
While (($MBXCounter -lt 5) -and (($MBXServer.DatabaseCopyActivationDisabledAndMoveNow -ne $True) -or ($MBXServer.DatabaseCopyAutoActivationPolicy -ne "Blocked"))) {
    Set-MailboxServer $ThisServer -DatabaseCopyActivationDisabledAndMoveNow $true
    Set-MailboxServer $ThisServer -DatabaseCopyAutoActivationPolicy Blocked
    Start-Sleep -Seconds 2
    $MBXServer = Get-MailboxServer $ThisServer
    $MBXCounter++
}
if (($MBXServer.DatabaseCopyActivationDisabledAndMoveNow) -and ($MBXServer.DatabaseCopyAutoActivationPolicy -eq "Blocked")) {
    Write-Host -ForegroundColor Green "DatabaseCopy Activation Disabled and Blocked"
}
else {
    Write-Host -ForegroundColor Red "DatabaseCopy Activation Disabled and Blocked attempted but not confirmed"
}

Write-Host "ServerWideOffline..."
Set-ServerComponentState $ThisServer -Component ServerWideOffline -State Inactive -Requester Maintenance
Start-Sleep -Seconds 2
$SWOState = Get-ServerComponentState $ThisServer -Component ServerWideOffline
$SWOCounter = 0
While (($SWOState.State -ne "Inactive") -and ($SWOCounter -lt 5)) {
    Set-ServerComponentState $ThisServer -Component ServerWideOffline -State Inactive -Requester Maintenance
    Start-Sleep -Seconds 2
    $SWOState = Get-ServerComponentState $ThisServer -Component ServerWideOffline
    $SWOCounter++
}
if ($SWOState.State -eq "Inactive") {
    Write-Host -ForegroundColor Green "ServerWideOffline Set"
}
else {
    Write-Host -ForegroundColor Red "ServerWideOffline Set attempted but not confirmed"
}

Write-Host "Patching Starting..."
Get-WUInstall -AcceptAll -AutoReboot -Install