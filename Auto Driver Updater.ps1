#set Window Title
Set-TimeZone -Id "Romance Standard Time"
Start-Sleep -Seconds 1
$host.ui.RawUI.WindowTitle = "Auto Driver Updater v3.0 (19 jun 2023) for Foxway A/S by Johny Bartholdy Jensen " + (Get-Date).ToString("\[HH\:mm\]")
$monitor = Get-WmiObject -ns root/wmi -class wmiMonitorBrightNessMethods -EA SilentlyContinue
if ($monitor) {$monitor.WmiSetBrightness(1,100)}
Clear-Host

# predefined
$rebootrequired = 0
$looped = 0

# all functions
function text_time {Write-Host("Updating time ...`n") -Foregroundcolor Green}
function text_search {Write-Host("`nSearching for driver updates ...") -Foregroundcolor Green}
function text_download {Write-Host("Downloading drivers ...") -Foregroundcolor Green}
function text_install_start {Write-Host("`nInstalling drivers ...") -Foregroundcolor Green}
function text_install_done {Write-Host("`nInstallation complete!") -Foregroundcolor Green}
function text_no_update {Write-Host("`nNo updates available!") -Foregroundcolor Red}
function text_no_internet {Write-Host("`nNo internet connection found!") -Foregroundcolor Red}
function text_reboot {Write-Host("`nReboot required! Rebooting in 10 seconds!") -Foregroundcolor Red}
function text_research_max {Write-Host("`nMaximum number of searches has been reached! Rebooting in 10 seconds!") -Foregroundcolor Red}
function test_ifmodel {
    $systemfamily = Get-CimInstance win32_computersystem | select-object systemfamily
    if ($systemfamily -match "ThinkPad P16s Gen 1"){
        Add-Type -AssemblyName PresentationFramework
        $msgBoxInput = [System.Windows.MessageBox]::Show('The P16s needs to have run "TVSU" before running "Auto Driver Updater". Ask the team leader for more information.' + [System.Environment]::NewLine + [System.Environment]::NewLine + 'Did you run TVSU?','ThinkPad P16s dectected','YesNo','Error')
        switch ($msgBoxInput) {
            'Yes' {}
            'No' {exit}
        }
    }
}
function test_tcp_port {
    param(
        [string]$IP,
        [int[]]$Port,
        [int]$TCPTimeout = 100
    )
    $Result = [System.Collections.Generic.List[psobject]]::new()
    foreach ($Item in $Port) {
        $TCPClient = New-Object -TypeName System.Net.Sockets.TCPClient
        $AsyncResult = $TCPClient.BeginConnect($IP, $Item, $null, $null)
        $Wait = $AsyncResult.AsyncWaitHandle.WaitOne($TCPtimeout)
        if ($Wait) {
            try {
                $null = $TCPClient.EndConnect($AsyncResult)
            } catch {
                $Issue = $Error[0].Exception.InnerException.SocketErrorCode
            } finally {
                $Result.Add([pscustomobject]@{
                    IP = $IP
                    Port = $Item
                    IsOpen = $TCPClient.Connected
                    Notes = $Issue
                })
            }
        } Else {
            $Result.Add([pscustomobject]@{
            IP = $IP
            Port = $Item
            IsOpen = $TCPClient.Connected
            Notes = 'Timeout occurred connecting to port'
            })
        }
        $Issue = $Null
        $TCPClient.Dispose()
    }
    return $Result
}
function test_network {
    #$netcheckv1 = (Get-NetConnectionProfile).IPv4Connectivity -contains "Internet"
    #$netcheckv2 = Get-NetAdapter | where Status -eq "Up"
    #$netcheckv3 = if(Test-Connection 8.8.8.8 -Count 1 -ErrorAction SilentlyContinue){$true}else{$false}
    #$netcheckv4 = (New-Object Net.Sockets.TcpClient "8.8.8.8", 53).Connected
    $netcheckv5 = if((test_tcp_port 8.8.8.8 53 | select-Object isopen) -match "True") {$true}else{$false}
    if ((!$netcheckv5)) {        
        text_no_internet
        Pause
        exit
    }
}

# start running script
test_ifmodel
test_network
slmgr -ato

text_time
# Start the Windows Time service
try {
    $service = Get-Service -Name w32time -ErrorAction Stop
    if ($service.Status -ne "Running") {
        Start-Service -Name w32time
        Write-Host "Windows Time service started successfully."
    } else {
        Write-Host "Windows Time service is already running."
    }
} catch {
    Write-Host "Failed to start the Windows Time service."
}

# Trigger time synchronization
try {
    w32tm /resync /rediscover | Out-Null
    Write-Host "Windows Time synchronization triggered successfully."
} catch {
    Write-Host "Failed to trigger time synchronization."
}

while ($looped -lt 7) {
    $looped += 1
    
    # search and list all missing drivers
    test_network
    $Session = New-Object -ComObject Microsoft.Update.Session           
    $Searcher = $Session.CreateUpdateSearcher() 
    $Searcher.ServiceID = '7971f918-a847-4430-9279-4a52d1efe18d' # added the Microsoft Update Service GUID
    $Searcher.SearchScope =  1 # MachineOnly
    $Searcher.ServerSelection = 2 # Third Party
    test_network
    $Criteria = "IsInstalled=0 and Type='Driver' and ISHidden=0"
    text_search
    $SearchResult = $Searcher.Search($Criteria)          
    $Updates = $SearchResult.Updates
    
    # if no updates end loop
    if (-not ($Updates | Select-Object Driverclass, DriverModel)) {break}

    # show available drivers
    $Updates | Select-Object Driverclass, DriverModel | Format-Table -AutoSize -HideTableHeaders

    # check if downloads are available and trigger the downloading
    text_download
    $UpdatesToDownload = New-Object -ComObject Microsoft.Update.UpdateColl
    $updates | ForEach-Object {$UpdatesToDownload.Add($_) | out-null}
    $UpdateSession = New-Object -ComObject Microsoft.Update.Session
    $Downloader = $UpdateSession.CreateUpdateDownloader()
    $Downloader.Updates = $UpdatesToDownload
    Write-Progress -Activity "Downloading drivers" -Status "Starting downloading ..." -PercentComplete 0
    $TotalUpdates = $Downloader.Updates.Count
    $CompletedUpdates = 0
    $Downloader.Updates | ForEach-Object {
        $CompletedUpdates++
        $Update = $_
        $UpdateTitle = "[$CompletedUpdates/$TotalUpdates] " + $Update.DriverModel
        $PercentComplete = ($CompletedUpdates / $TotalUpdates) * 100
        Write-Progress -Activity "Downloading drivers" -Status "Downloading ..." -CurrentOperation $UpdateTitle -PercentComplete $PercentComplete
        $Downloader.Download() | Out-Null
    }
    Write-Progress -Activity "Downloading drivers" -Completed
        
    # check if the drivers are all downloaded and trigger the installation
    text_install_start
    $UpdatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl
    $Updates | ForEach-Object {$UpdatesToInstall.Add($_) | Out-Null}
    $Installer = $UpdateSession.CreateUpdateInstaller()
    $Installer.Updates = $UpdatesToInstall
    $TotalUpdates = $Installer.Updates.Count
    Write-Progress -Activity "Installing Drivers" -Status "Starting installation ..." -PercentComplete 0
    $CompletedInstalls = 0
    $Installer.Updates | ForEach-Object {
        $CompletedInstalls++
        $Update = $_    
        $UpdateTitle = "[$CompletedInstalls/$TotalUpdates] " + $Update.DriverModel
        $PercentComplete = ($CompletedInstalls / $TotalUpdates) * 100
        Write-Progress -Activity "Installing Drivers" -Status "Installing ..." -CurrentOperation $UpdateTitle -PercentComplete $PercentComplete
        $Installer.Install() | Out-Null
    }
    Write-Progress -Activity "Installing Drivers" -Completed
    $rebootrequired = 1
}

# clean up
$updateSvc.Services | Where-Object {$_.IsDefaultAUService -eq $false -and $_.ServiceID -eq "7971f918-a847-4430-9279-4a52d1efe18d"} | ForEach-Object {$UpdateSvc.RemoveService($_.ServiceID)}

# internet check
test_network
text_no_update

# reboot required check
if(($rebootrequired -gt 0) -or ($looped -gt 6)) {  
    text_reboot
    Start-Sleep -Seconds 10
    Restart-Computer -Force
}

# max count check
if(($rebootrequired -eq 0) -and ($looped -gt 6)) { 
    text_research_max
    if($rebootrequired -lt 1) {
        Start-Sleep -Seconds 10
        Restart-Computer -Force
    }
}

Write-Host(" ")
Pause
