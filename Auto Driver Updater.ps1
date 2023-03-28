#set Window Title
$host.ui.RawUI.WindowTitle = "Auto Driver Updater v2.7 (15 jun 2022) for DCC Support by Johny Bartholdy Jensen"
cls

# predefined
$rebootrequired = 0
$looped = 0

function Get-nointernet {
    Write-Host("`n*************************************************") -Fore red -Back black
    Write-Host("******                                     ******") -Fore red -Back black
    Write-Host("******    No internet connection found!    ******") -Fore red -Back black
    Write-Host("******                                     ******") -Fore red -Back black
    Write-Host("*************************************************`n") -Fore red -Back black
    Pause
    exit
}

function Get-dlfailed {
    Write-Host("`n*************************************************") -Fore red -Back black
    Write-Host("******                                     ******") -Fore red -Back black
    Write-Host("******  Download failed! Please try again! ******") -Fore red -Back black
    Write-Host("******                                     ******") -Fore red -Back black
    Write-Host("*************************************************`n") -Fore red -Back black
    Pause
    exit
}
function Test-TCPPort {
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

function Get-netcheck {
    #$netcheckv1 = (Get-NetConnectionProfile).IPv4Connectivity -contains "Internet"
    #$netcheckv2 = Get-NetAdapter | where Status -eq "Up"
    #$netcheckv3 = if(Test-Connection 8.8.8.8 -Count 1 -ErrorAction SilentlyContinue){$true}else{$false}
    #$netcheckv4 = (New-Object Net.Sockets.TcpClient "8.8.8.8", 53).Connected
    $netcheckv5 = if((Test-TCPPort 8.8.8.8 53 | select-Object isopen) -match "True") {$true}else{$false}
    if ((!$netcheckv5)) {        
        Get-nointernet
        Pause
    }
}
Get-netcheck

# update time
Write-Host("Updating time ...") -Fore Green
net start w32time
w32tm /resync /rediscover

# search and list all missing drivers
Get-netcheck
$Session = New-Object -ComObject Microsoft.Update.Session           
$Searcher = $Session.CreateUpdateSearcher() 
$Searcher.ServiceID = '7971f918-a847-4430-9279-4a52d1efe18d' # added the Microsoft Update Service GUID
$Searcher.SearchScope =  1 # MachineOnly
$Searcher.ServerSelection = 2 # Third Party
Get-netcheck
$Criteria = "IsInstalled=0 and Type='Driver' and ISHidden=0"
Write-Host("`nSearching for Driver Update ...") -Fore Green  
$SearchResult = $Searcher.Search($Criteria)          
$Updates = $SearchResult.Updates

# loop
while (($looped -lt 7) -and ($Updates | select Title, DriverModel, DriverVerDate, Driverclass, DriverManufacturer | fl)) {
    $looped += 1
    
    # show available drivers
    $Updates | select Title, DriverModel, DriverVerDate, Driverclass, DriverManufacturer | fl

    # download the drivers from Microsoft
    $UpdatesToDownload = New-Object -Com Microsoft.Update.UpdateColl
    $updates | % { $UpdatesToDownload.Add($_) | out-null }
    Write-Host("Downloading Drivers ...")  -Fore Green  
    $UpdateSession = New-Object -Com Microsoft.Update.Session
    $Downloader = $UpdateSession.CreateUpdateDownloader()
    $Downloader.Updates = $UpdatesToDownload
    if ($Downloader.Download()) {
        $Downloader.Download()
    } else {
        Get-dlfailed
    }

    # check if the drivers are all downloaded and trigger the installation
    $UpdatesToInstall = New-Object -Com Microsoft.Update.UpdateColl
    $updates | % { if($_.IsDownloaded) { $UpdatesToInstall.Add($_) | out-null } else {
        Get-dlfailed
    }}

    Write-Host("Installing Drivers ...")  -Fore Green
    $Installer = $UpdateSession.CreateUpdateInstaller()
    $Installer.Updates = $UpdatesToInstall
    $InstallationResult = $Installer.Install()
    Write-Host("`nInstallation complete!") -Fore Green
    if($InstallationResult.RebootRequired) {
        $rebootrequired = 1
    }

    # search and list all missing Drivers
    $Session = New-Object -ComObject Microsoft.Update.Session           
    $Searcher = $Session.CreateUpdateSearcher() 
    $Searcher.ServiceID = '7971f918-a847-4430-9279-4a52d1efe18d' # added the Microsoft Update Service GUID
    $Searcher.SearchScope =  1 # MachineOnly
    $Searcher.ServerSelection = 2 # Third Party
    $Criteria = "IsInstalled=0 and Type='Driver' and ISHidden=0"
    Write-Host("`nResearching for Driver Update ...") -Fore Green  
    $SearchResult = $Searcher.Search($Criteria)          
    $Updates = $SearchResult.Updates
}

# internet check
Get-netcheck
Write-Host("`nNo updates available!") -Fore red

# reboot required check
if(($rebootrequired -gt 0) -or ($looped -gt 6)) {  
    Write-Host("`nReboot required! Rebooting in 10 seconds!") -Fore Red
    shutdown -r -t 10
}

# max count check
if(($rebootrequired = 0) -and ($looped -gt 6)) { 
    Write-Host("`nMaximum number of searches has been reached! Rebooting in 10 seconds!") -Fore Red
    if($rebootrequired -lt 1) {
        shutdown -r -t 10
    }
}
Write-Host(" ")
Pause