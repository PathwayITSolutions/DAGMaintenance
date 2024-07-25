<#
    .SYNOPSIS
        Automated Exchange DAG Maintenance and Patching v4.2

    .DESCRIPTION
        The second half of the Automated Exchange DAG Maintenance and Patching script team,
        Bringing a patched server back online. This reverses all the steps taken in PatchStart.ps1
        and then removes any inetpub files older than 60 days from the LogFiles section.

        This script needs no inputs, as it will read the server name from the environment.
        It requires An Exchange window in Admin mode so that it can address the cluster.

    .EXAMPLE
        .\PatchEnd.ps1

    .LINK
        https://www.pathwayit.net

    .NOTES
        Work in Progress, items in the TODO
        TODO: Resume-ClusterNode should read the output and act on it
        TODO: Script needs to read which databases belong on ThisServer as primary and move them Get-MailboxDatabase | Select Server, Name, DatabaseCopies,ActivationPreference
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
    $PSWUState = Get-WUInstall
    if ($null -ne $PSWUState) {
        Write-Host -ForegroundColor Red "Patching not completed, running PSWindowsUpdate again. Restart this script once complete..."
        Get-WUInstall -AcceptAll -AutoReboot -Install
        exit
    }
}
else {
    Write-Host -ForegroundColor Red "Unable to check patch status, check your PSWindowsUpdate installation"
}

$ThisServer = $env:computername

Write-Host "Setting ServerWideOffline to Active..."
Set-ServerComponentState $ThisServer -Component ServerWideOffline -State Active -Requester Maintenance
Start-Sleep -Seconds 2
$SWOState = Get-ServerComponentState $ThisServer -Component ServerWideOffline
$SWOCounter = 0
While (($SWOState.State -ne "Active") -and ($SWOCounter -lt 5)) {
    Set-ServerComponentState $ThisServer -Component ServerWideOffline -State Active -Requester Maintenance
    Start-Sleep -Seconds 2
    $SWOState = Get-ServerComponentState $ThisServer -Component ServerWideOffline
    $SWOCounter++
}
if ($SWOState.State -eq "Active") {
    Write-Host -ForegroundColor Green "ServerWideOffline Set to Active"
}
else {
    Write-Host -ForegroundColor Red "ServerWideOffline Set to Active attempted but not confirmed"
}

Write-Host "Resuming Cluster Node..."
Resume-ClusterNode -Name $ThisServer | Out-Null
Write-Host -ForegroundColor Green "Cluster Node Resumed"

Write-Host "Setting Unrestricted DatabaseCopy Auto Activation..."
Set-MailboxServer $ThisServer -DatabaseCopyAutoActivationPolicy Unrestricted
Set-MailboxServer $ThisServer -DatabaseCopyActivationDisabledAndMoveNow $False
Start-Sleep -Seconds 2
$MBXServer = Get-MailboxServer $ThisServer
$MBXCounter = 0
While (($MBXCounter -lt 5) -and (($MBXServer.DatabaseCopyActivationDisabledAndMoveNow) -or ($MBXServer.DatabaseCopyAutoActivationPolicy -ne "Unrestricted"))) {
    Set-MailboxServer $ThisServer -DatabaseCopyAutoActivationPolicy Unrestricted
    Set-MailboxServer $ThisServer -DatabaseCopyActivationDisabledAndMoveNow $False
    Start-Sleep -Seconds 2
    $MBXServer = Get-MailboxServer $ThisServer
    $MBXCounter++
}
if (($MBXServer.DatabaseCopyActivationDisabledAndMoveNow -eq $False) -and ($MBXServer.DatabaseCopyAutoActivationPolicy -eq "Unrestricted")) {
    Write-Host -ForegroundColor Green "DatabaseCopy Activation Unrestricted and Unblocked"
}
else {
    Write-Host -ForegroundColor Red "DatabaseCopy Activation Unrestricted and Unblocked attempted but not confirmed"
}

Write-Host "Setting HubTransport Active..."
Set-ServerComponentState $ThisServer -Component HubTransport -State Active -Requester Maintenance
Start-Sleep -Seconds 2
$HTState = Get-ServerComponentState $ThisServer -Component HubTransport
$HTCounter = 0
While (($HTState.State -ne "Active") -and ($HTCounter -lt 5)) {
    Set-ServerComponentState $ThisServer -Component HubTransport -State Active -Requester Maintenance
    Start-Sleep -Seconds 2
    $HTState = Get-ServerComponentState $ThisServer -Component HubTransport
    $HTCounter++
}
if ($HTState.State -eq "Active") {
    Write-Host -ForegroundColor Green "HubTransport Active Set"
}
else {
    Write-Host -ForegroundColor Red "HubTransport Active Set attempted but not confirmed"
}

Write-Host "Testing Service Health..."
$TSH = Test-ServiceHealth $ThisServer
$TSHCounter = 0
While ($TSHCounter -lt 3) {
    if ($TSH[$TSHCounter].RequiredServicesRunning) { 
        Write-Host -ForegroundColor Green "$($TSH[$TSHCounter].Role) OK"
        $TSHCounter++
    }
    else {
        $DSTotal = $($TSH[$TSHCounter].ServicesNotRunning).Count
        foreach ($DownService in $TSH[$TSHCounter].ServicesNotRunning) {
            $ServiceState = Get-Service $DownService
            $DSCounter = 0
            Write-Host "Service $($ServiceState.DisplayName) not Running..."
            While (($ServiceState.Status -ne "Running") -and ($DSCounter -lt 5)) {
                Start-Service $DownService
                Start-Sleep -Seconds 5
                $ServiceState = Get-Service $DownService
                $TSH = Test-ServiceHealth $ThisServer
                $DSCounter++
            }
            if ($ServiceState.Status -eq "Running") {
                Write-Host "$($DownService) started"
                $DSTotal--
            }
            else {
                Write-Host -ForegroundColor Red "$($DownService) failed to start, investigate"
                $DSTotal--
                if ($DSTotal -eq 0) {
                    Write-Host -ForegroundColor Red "$($TSH[$TSHCounter].Role) Down"
                    $TSHCounter++
                }
            }
        }
    }
}

$FileAge = (Get-Date).AddDays(-60)
$FilePath = "C:\inetpub\logs\LogFiles"
Write-Host "Deleting inetpub LogFiles older than 60 days..."
$FileList = Get-ChildItem -Path $FilePath -Recurse | Where-Object { !$_.PSIsContainer -and $_.CreationTime -lt $FileAge }
Write-Host -ForegroundColor Green "Deleting $($FileList.Count) files"
Get-ChildItem -Path $FilePath -Recurse -Force | Where-Object { !$_.PSIsContainer -and $_.CreationTime -lt $FileAge } | Remove-Item -Force