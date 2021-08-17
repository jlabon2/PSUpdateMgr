
Import-Module -Name `
    (Join-Path -Path $PSScriptRoot -ChildPath  '\mod\internal.psm1'),`
    (Join-Path -Path $PSScriptRoot -ChildPath  '\mod\PoshRSJob\PoshRSJob.psm1')

Set-WindowVisibility 
$configPath = Join-Path -Path $PSScriptRoot -ChildPath  'config.json'
$xamlPath = Join-Path -Path $PSScriptRoot -ChildPath  'MainWindow.xaml'
$rdcMan = Join-Path -Path $PSScriptRoot -ChildPath  'RDCMan.exe'

# loaded required DLLs
foreach ($dll in ((Get-ChildItem -Path (Join-Path -Path $PSScriptRoot -ChildPath lib) -Filter *.dll).FullName)) { $null = [System.Reflection.Assembly]::LoadFrom($dll) }

New-HashTables

$varhash.jsonPath = Join-Path -Path $PSScriptRoot -ChildPath  'prev.json'
$varHash.tableRefresh = [System.Diagnostics.Stopwatch]::StartNew()
$varHash.Finish = $false
$varHash.listUpdated = $false
$varHash.selectedFilter = "All"

$varHash.filterChoice = @{
    'All items'     = 'All'
    'Needs Updates' = 'UpdateNeeded'
    'Needs Reboot'  = 'RebootNeeded'
}

$varHash.psModPath = @{
    psmodBase8  = (Get-Module -Name PSwindowsUpdate -ListAvailable | Where-Object { $_.Version -eq "2.2.0.2" }).ModuleBase # 2.2 seems to work w 2k8?
    psmodBase12 = (Get-Module -Name PSwindowsUpdate -ListAvailable | Where-Object { $_.Version -eq "2.2.0.2" }).ModuleBase
}

Set-WPFControls -TargetHash $guiHash -XAMLPath $xamlPath -WindowName "UM2.MainWindow"

$valueHash.ObjectModel = New-Object System.Collections.ObjectModel.ObservableCollection[Object]

if (!(Test-Path $configPath)) {
    $guiHash.initialUpdatePanel.Visibility = 'Collapsed'
    $guiHash.actionConfigWindow.IsOpen = $true
    $guiHash.actionConfigWindow.IsModal = $true
    $guiHash.actionConfigWindow.CloseByEscape = $false
    $guihash.actionConfigWindow.ShowCloseButton = $false
    $guiHash.actionConfigWindow.Title  = "Settings"
    $guiHash.toolConfigPanel.Visibility = "Visible"
    $guiHash.toolCancel.IsEnabled = $false 
}

else {

    $configImport = Get-Content -Path $configPath | ConvertFrom-Json 

    $varHash.logDir     = $configImport.LogPath
    $script:LogDir      = $varHash.logDir
    $varHash.WSUSServer = $configImport.Server
    $varHash.WSUSport   = $configImport.Port
    $varHash.Groups     = $configImport.Groups 

    $varHash.LogDirCurrent = (Join-Path -Path $varHash.logDir -ChildPath ((Get-Date -format "MMM yyyy").ToString()))
    $varhash.logListing = [System.Collections.ArrayList]@()

    if (!(Test-Path $varHash.LogDirCurrent)) {
        New-Item -ItemType Directory -Path $varHash.LogDirCurrent
    }

    Get-ChildItem $varHash.LogDir | Where-Object { $_.LastWriteTime -gt (Get-Date).AddMonths(-3) } | ForEach-Object {
        $dirDate = $_.Name
        Get-ChildItem $_.FullName | ForEach-Object { $varHash.LogListing.Add((
                    [PSCustomObject]@{
                        Fullname = $_.FullName
                        Name     = $_.Name -replace '.log'
                        Date     = $dirDate
                    })) | Out-Null
        }
    } 

    if (Test-Path $varHash.jsonPath) {
        $jsonImport = Get-Content $varhash.jsonPath | ConvertFrom-JSON
        foreach ($item in ($jsonImport.PSObject.Properties.Name)) {
            if ($jsonImport.$item.lastSyncTime -gt (Get-Date)) {
                $jsonImport.$item.lastSyncTime = (Get-Date $jsonImport.$item.LastSyncTime).AddHours(-7)
            }
            $valueHash.syncedHash.$item = $jsonImport.$item
        }
        $valueHash.syncedHash.Values | ForEach-Object { $valueHash.ObjectModel.Add($_) }
        $varHash.firstRunDone = $true
        $varHash.listPopulated = $true
        $varHash.initialUpdate = $true
        $guiHash.updateSwitch.isEnabled = $true
    } 

    else {
        $guiHash.initialUpdate.Content = "No history found. Updating..."
        $guiHash.missingText.Text = "Pulling information... This may take a moment."
        $varHash.firstRunDone = $false
    }

    $valuehash.ViewModel = [System.Windows.Data.CollectionViewSource]::GetDefaultView($valueHash.objectModel)
    $guiHash.resultsGrid.ItemsSource = $valueHash.ViewModel
    $valueHash.ViewModel.IsLiveFiltering = $true
    $valueHash.viewModel.IsLiveSorting = $true

    @("UpdatesNeeded", "RebootNeeded", "Connectivity", "Updater", "ComputerName", "LastSyncTime", "OS") | ForEach-Object {
        $valueHash.viewModel.LiveFilteringProperties.Add($_)
        $valueHash.viewModel.LiveSortingProperties.Add($_)
    }

    Start-RSJob -Name objectModelUpdate -Argument $valueHash, $varHash, $guiHash -ScriptBlock {
        param ($valueHash, $varHash, $guiHash)
        $ErrorActionPreference = 'Continue'
        do {
            do { Start-Sleep -seconds 5 } until ($varHash.newRefresh -eq $true -or $varHash.Finish -eq $true)

            if ($varHash.newRefresh) {
                if (($valueHash.objectModel).Count -eq 0) {
                    $valueHash.syncedHash.Values | ForEach-Object { $valueHash.ObjectModel.Add($_) }    
                }
                foreach ($compName in $valueHash.syncedHash.Keys) {
                    if ($compName -notin $valueHash.objectModel.ComputerName) {
                        $valueHash.objectModel.Add($valueHash.syncedHash.$compName)
                    }
                    else {
                        foreach ($prop in ($valueHash.syncedHash.$compName).PSObject.Properties.Name) {
                            if ($prop -notmatch "(connectivity|updater)") {
                                if ($valueHash.objectModel.Where( { $_.ComputerName -eq $compName }).$prop -ne $valueHash.syncedHash.$compName.$prop) {
                                    ($valueHash.objectModel | Where-Object { $_.ComputerName -eq $compName }).$prop = $valueHash.syncedHash.$compName.$prop
                                }
                            }
                        }
                        foreach ($prop in ($valueHash.syncedStateHash.$compName).PSObject.Properties.Name) {
                            if ($valueHash.objectModel.Where( { $_.ComputerName -eq $compName }).$prop -ne $valueHash.syncedStateHash.$compName.$prop) {
                                ($valueHash.objectModel | Where-Object { $_.ComputerName -eq $compName }).$prop = $valueHash.syncedStateHash.$compName.$prop
                            }
                        }        
                    }
                }
                $valueHash.syncedHash | ConvertTo-Json | Out-File $varHash.jsonPath -Force
                $varHash.newRefresh = $false

                if ($varHash.firstRunDone -ne $true) {
                    $guiHash.Window.Dispatcher.Invoke([action] {
                            $guiHash.resultsGrid.ItemsSource = $valueHash.ViewModel
                            $guiHash.missingText.Text = "No results based on filter criteria"
                            $varHash.firstRunDone = $true
                    })
                }

                $guiHash.Window.Dispatcher.Invoke([action]{$guiHash.updateSwitch.IsEnabled = $false})

            }
        } until ($varHash.Finish -eq $true)
        exit
    }

    Start-RSJob -Name WSUSPull -ArgumentList $valueHash, $varHash, $guiHash -ModulesToImport "UpdateServices" -FunctionsToImport Get-TimeDifference, Get-TimeSpanStringValue {
        param ($valueHash, $varHash, $guiHash)

        $ErrorActionPreference = 'Continue'
        [void][Reflection.Assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration")

        do {
            $tempHash = @{}
            $WSUS = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer($varHash.WSUSServer, $false, $varHash.WSUSPort)
            $computerScope = New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope

            $WSUS.GetComputerTargetGroups() | Where-Object { $_.Name -in $varHash.Groups } | ForEach-Object {
                $computerScope.ComputerTargetGroups.Add($_) | Out-Null
            }

            $updateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope -Property @{
                UpdateApprovalActions      = [Microsoft.UpdateServices.Administration.UpdateApprovalActions]::Install
                IncludedInstallationStates = @("Failed", "NotInstalled", "Downloaded")
            }
        
            $rebootScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope -Property @{
                IncludedInstallationStates = "InstalledPendingReboot"
            }

            $systems = $WSUS.GetComputerTargets($computerScope)

            foreach ($system in $systems) {
                $tempHash.($system.FullDomainName) = [PSCustomObject]@{
                    ComputerName  = $system.FullDomainName
                    OS            = $system.OSDescription -replace 'Edition' -replace "\(full installation\)"
                    Connectivity  = $null
                    LastSyncTime  = [System.TimeZoneInfo]::ConvertTimeFromUTC($system.LastReportedStatusTime, [System.TimeZoneInfo]::FindSystemTimeZoneByID((Get-WmiObject Win32_TimeZone).StandardName))
                    UpdatesNeeded = ($system.GetUpdateInstallationInfoPerUpdate($updateScope) | Measure-Object).Count
                    RebootNeeded  = if ($system.GetUpdateInstallationInfoPerUpdate($rebootScope)) { $true.toString() }
                    else { $false.toString() }
                    Updater       = "Waiting"
                }
            }

            $tempHash.Keys | ForEach-Object { $tempHash.$_ | Add-Member -Force -MemberType NoteProperty -Name 'LastSyncDiff' -Value (Get-TimeDifference -StartTime $tempHash.$_.LastSyncTime) }        
            $tempHash.Keys | ForEach-Object { $tempHash.$_ | Add-Member -Force -MemberType NoteProperty -Name 'DiffTag' -Value (Get-TimeSpanStringValue -TimeDiffString $tempHash.$_.LastSyncDiff) }

            if ($varHash.listUpdated -ne $true ) { 
                 if ($valueHash.syncedHash.Keys.Count -ne $tempHash.Keys.Count) {
                    $valueHash.syncedHash = [hashtable]::Synchronized(@{})
                }

                foreach ($key in $tempHash.Keys) { $valueHash.syncedHash.$key = $tempHash.$key }
               
                    
                $varHash.List = $valueHash.syncedHash.Keys
                $varHash.listPopulated = $true

            }
            else {
                foreach ($tempKey in $tempHash.Keys) {
                    foreach ($prop in ($tempHash.$tempKey).PSObject.Properties.Name) {
                        if ($varHash.listPopulated -ne $true -or $prop -notmatch "(connectivity|updater)") {
                            if ($valueHash.syncedHash.$tempKey.$prop -ne $tempHash.$tempKey.$prop) {
                                $valueHash.syncedHash.$tempKey.$prop = $tempHash.$tempKey.$prop
                            }
                        }
                    }
                }
            }
            $varHash.listUpdated = $true
            $varHash.newRefresh = $true
            $varHash.firstRunDone = $true
            Start-Sleep -Seconds 15

        } until ($varHash.Finish -eq $true)

    }
}
$guiHash.filterTextBox.Add_KeyDown( 
    {
        $keyPress = $args[1].Key
        if ($keyPress -eq 'Return') {
            $varHash.SearchText = $guiHash.FilterTextBox.Text
            $global:SearchItem = $varHash.SearchText
            Set-Filter -SearchItem $searchItem -SearchList $valueHash.ObjectModel -SelectedFilter $varHash.filterChoice.($guiHash.FilterSelection.SelectedValue.Content)
        }
    }
)

$guiHash.FilterSelection.Add_SelectionChanged({  Set-Filter -SearchItem $guiHash.FilterTextBox.Text -SearchList $valueHash.ViewModel -SelectedFilter $varHash.filterChoice.($guiHash.FilterSelection.SelectedValue.Content)})

$guiHash.resultsGrid.Add_SelectionChanged( 
    {
        if (($guiHash.resultsGrid.SelectedItems | Measure-Object).Count -ge 1) {
            $guiHash.FlyOut.isOpen = $true
            if (($guiHash.resultsGrid.Selecteditems | Measure-Object).Count -eq 1) {
                $guiHash.selectionLabel.Content = $guiHash.resultsGrid.SelectedItems.ComputerName
                $guiHash.installWinHeader.Content = $guiHash.SelectionLabel.Content
                $guiHash.abortWinHeader.Content = $guiHash.SelectionLabel.Content
                $guiHash.rebootWinHeader.Content = $guiHash.SelectionLabel.Content
                $guihash.mstscButton.Tag = $null
            
                $varHash.CurrentLogs = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
            
                $varHash.LogListing | Where-Object { $_.Name -eq "$($(($guiHash.resultsGrid.SelectedItems.ComputerName) -split '\.')[0])" } | 
                    Select-Object FullName, Name, @{Label = 'Date'; Expression = { Get-Date($_.Date) } } |
                    Sort-Object -Unique FullName | Sort-Object -Property Date -Descending | 
                    Select-Object FullName, Name, @{Label = 'Date'; Expression = { (Get-Date($_.Date) -format "MMM yyyy") } } |
                    ForEach-Object { $varHash.CurrentLogs.Add($_) }
        
                if ($currentLogs.Date -notcontains (Get-Date -Format "MMM yyyy").toString()) {
                    if (Test-Path ((Join-Path -Path $varHash.LogDirCurrent -ChildPath "$($(($guiHash.resultsGrid.SelectedItems.ComputerName) -split '\.')[0]).log"))) {
                        $varHash.currentLogs.Add((Get-ChildItem -Path $varHash.LogDirCurrent -Filter ("$($(($guiHash.resultsGrid.SelectedItem.ComputerName) -split '\.')[0]).log") |
                                Select-Object FullName, @{Label = 'Name'; Expression = { $_.Name -replace '.log' } }, @{Label = 'Date'; Expression = { (Get-Date -Format "MMM yyyy") } }))
                    }
                }

                if ($varHash.currentLogs) {
                    $guiHash.logButton.ItemsSource = $varHash.currentLogs
                    $guiHash.logButton.SelectedIndex = 0
                    $guiHash.logButton.Visibility = "Visible"
                }
                else { $guiHash.logButton.Visibility = "Hidden" }
        
            }
      

            else {
                $guiHash.mstscButton.Tag = 'Multi'
                $guiHash.SelectionLabel.content = "$(($GuiHash.resultsGrid.SelectedItems | Measure-Object).Count) systems selected"
                $guiHash.installWinHeader.Content = $guiHash.selectionLabel.Content
                $guiHash.abortWinHeader.Content = $guiHash.selectionLabel.Content
                $guiHash.rebootWinHeader.Content = $guiHash.selectionLabel.Content
                $guiHash.logButton.Visibility = "Hidden"
            }      
        }

        else {$guiHash.FlyOut.isOpen = $false}

        $varhash.SelectedCount = ($guiHash.resultsGrid.SelectedItems | Measure-Object).Count
        $varHash.SelectedIndex = $guiHash.resultsGrid.SelectedIndex
    }
)

$guiHash.buttonSelectAll.Add_Click( { $guiHash.resultsGrid.SelectAll() } )

$guiHash.buttonUnselectAll.Add_Click( { $guiHash.resultsGrid.UnselectAll() } )

$guiHash.logButton.Add_Click( {
        $lineNum = 001
        $guiHash.LogText.Text = ""
        $guiHash.logWindow.Title = "$($guiHash.resultsGrid.SelectedItems.ComputerName) Logs"
    
        foreach ($line in (Get-Content $guiHash.logBUtton.SelectedValue.FullName)) {
            $guiHash.logText.AppendText("$($lineNum.toString("00"))] $($line) `r`n")
            $lineNum++
        }

        $guiHash.logWindow.IsOpen = $true

    })

$guiHash.logButton.Add_DropDownClosed( {
        $lineNum = 001
        $guiHash.LogText.Text = ""
        $guiHash.logWindow.Title = "$($guiHash.resultsGrid.SelectedItems.ComputerName) Logs"
    
        foreach ($line in (Get-Content $guiHash.logBUtton.SelectedValue.FullName)) {
            $guiHash.logText.AppendText("$($lineNum.toString("00"))] $($line) `r`n")
            $lineNum++
        }

        $guiHash.logWindow.IsOpen = $true

    })

$guiHash.resultsGrid.Add_IsVisibleChanged( {
        if ($varHash.FirstRunDone -eq $false) {  $valueHAsh.viewModel.Filter = $null }

    })

 Start-RSJob -Name "StatusCheckParent" -ArgumentList $varHash, $valueHash, $guiHash -ModulesToImport "ScheduledTasks" {
                param ($varHash, $valueHash, $guiHash)
                
                do {Start-Sleep -Seconds 5} until ($guiHash.resultsGrid.Visibility -eq 'Visible' -and $varHash.listUpdated -eq $true)
                $varHash.statusRefresh = [System.Diagnostics.Stopwatch]::StartNew()

                do {
                    Get-RSJob | Remove-RSJob

                    Start-RSJob -Name "StatusCheck" -Batch "StatusCheck" -Throttle 6 -InputObject $valueHash.syncedHash.Keys -ArgumentList $varHash, $valueHash -ModulesToImport "ScheduledTasks" -ScriptBlock {
                        param ($varHash, $valueHash)
                        $connectStatus = Test-Connection -ComputerName $_ -Quiet -Count 2

                        if ($valueHash.syncedHash.$_.UpdatesNeeded -gt 0) {
                       
                            if ($connectStatus) {                           
                                if ($valueHash.syncedHash.$_.OS -notlike "*2008*") {                                          
                                    $rs = Start-RSJob -Name "GetTask" -ArgumentList $_ -ScriptBlock {
                                    try { Get-ScheduledTask -CimSession $args[0] | Where-Object { $_.TaskName -eq "PSwindowsUpdate" } }
                                        catch [System.Management.Automation.SetValueInvocationException] { $task[0].State }
                                        catch [System.Management.Automation.RuntimeException] { "Not started" }
                                        catch { "UseLegacy" }                                   
                                    }
                                    Wait-RSJob -Name "GetTask" -Timeout 30
                                    $psTask = $rs | Receive-RSJob

                                    if ($psTask -eq 'UseLegacy') { $useLegacy = $true }
                                    if ($null -ne $psTask -and $psTask[0].State.ToString().Length -le 3) { $useLegacy = $true }                               
                                }

                                if ($valueHash.syncedHash.$_.OS -like "*2008*" -or $useLegacy -eq $true) {
                                    $rs = Start-RSJob -Name "GetTask" -ArgumentList $_ -ScriptBlock { schtasks /query /s $args[0] /FO CSV | ConvertFrom-CSV | Where-Object { $_.TaskName -like "*PSWINDOWSUPDATE*" } }
                                    Wait-RSJob -Name "GetTask" -Timeout 30
                                    $psTask = $rs | Receive-RSJob
                                }

                                if ($null -eq $psTask) { $psTask = "Not started" }
                                elseif ($psTask[0] -is [CimInstance]) {
                                    if ((Get-Date).AddSeconds(120) -lt ($psTask | Get-ScheduledTaskInfo).NextRunTime) {
                                        $schedTime = Get-Date ($psTask | Get-ScheduledTaskInfo).NextRunTime -Format 'MM/dd/yyyy hh:mm'
                                        $psTask = "Scheduled: $schedTime"
                                    }
                                    elseif (($psTask[0].State -ne 'Running' -and ($psTask | Get-ScheduledTaskInfo).NextRunTime -and (($psTask | Get-ScheduledTaskInfo).NextRunTime -lt (Get-Date).AddYears(-1)))) {
                                        #$psTask | Unregister-ScheduledTask -Confirm:$false
                                        $psTask = "Not started (stale)"
                                        $rerun = Get-ScheduledTask -CimSession $_ -ErrorAction SilentlyContinue | Where-Object { $_.TaskName -eq 'PSWU-Rerun' }
                                       # if ($rerun) { $rerun[0] | Unregister-ScheduledTask -Confirm:$false }
                                    }
                                    else { $psTask = $psTask[0].State }                              
                                }                          
                                elseif ($psTask[0] -is [PSCustomObject]) {
                                    if ((Get-Date).AddSeconds(120) -lt (Get-Date(($psTask).'Next Run Time')) -and (($psTask.Status -ne "Running"))) {
                                        $schedTime = Get-Date(($psTask).'Next Run Time') -Format 'MM/dd/yyyy hh:mm'
                                        $psTask = "Scheduled: $schedTime"
                                    }                             
                                    else { $psTask = $psTask[0].State }     
                                }
                                else { $psTask = "Not started" }
                            }
                            else { $psTask = "N/A" }

                        }
                        elseif ($valueHash.syncedHash.$_.UpdatesNeeded -eq 0 -and $connectStatus) { $psTask = "Complete" }
                        else { $psTask = "N/A" }


                         if ($valueHash.syncedStateHash.Keys -notcontains $_) {
                            $valueHash.syncedStateHash.Add($_,[PSCustomObject]@{
                                Connectivity = $null
                                Updater      = 'N/A'
                            })
                        }

                        $valueHash.syncedStateHash.$_.Connectivity = $connectStatus.toString()
                        $valueHash.syncedStateHash.$_.Updater      = if (![string]::IsNullOrWhiteSpace($psTask)) { $psTask.ToString() }
                                                                     else { "N/A" }

                        
                        
                    }
                    
                    Wait-RSJob -Batch "StatusCheck" -Timeout 60

                    $varHash.newRefresh = $true
                    do {} while ($varHash.statusRefresh.Elapsed.TotalSeconds -lt 60)
                    $varHash.statusRefresh.Restart()
                    

                } until ($varHash.Finish -eq $true)
            }


$guiHash.checkInButton.Add_Click( {
        if ($guiHash.resultsGrid.SelectedItems) {
            $checkinList = @($guiHash.resultsGrid.SelectedItems)
            $varHash.completeCheckinCount = $checkinList.Count
            $varHash.TotalCheckInCount = $checkinList.Count
    
            Start-RSJob -InputObject $checkinList -ArgumentList $varHash -Name 'CheckIn' -Batch 'CheckInBatch' -Throttle 4 -FunctionsToImport 'Invoke-WSUSReport' {
                param ($varHash)
                $varHash.completeCheckinCount--
                Start-RSJob -ArgumentList $_.ComputerName -Name 'CheckInComp' -FunctionsToImport 'Invoke-WSUSReport' -ScriptBlock { Invoke-WSUSReport -ComputerName $args[0] }
                Wait-RSJob -Name 'CheckInComp' -Timeout 25
                Exit
            }
    
            Start-RSJob -ArgumentList $guiHash, $varHash -Name 'CheckInMonitor' -ModulesToImport 'PoshRSJob' -ScriptBlock {
                param ($guiHash, $varhash)

                $checkCount = $varHash.completeCheckinCount

                do {
                    $toolTipString = "Forcing report in on client $($varHash.TotalCheckInCount - $varHash.completeCheckinCount) of $($varHash.TotalCheckInCount)"
                    $guiHash.Window.Dispatcher.Invoke([action] { $guiHash.checkInButton.ToolTip = $toolTipString })
                    Start-Sleep -Seconds 1
                } while ($varHash.completeCheckInCount -gt 0)

                $guiHash.Window.Dispatcher.Invoke([action]{ $guiHash.CheckinButton.ToolTip = "Force check in on selected clients" })
            }
        }

    })

$guiHash.clearButton.Add_Click( {
        $varHash.SearchText = $null
        $global:SearchItem = $null
        $guiHash.FilterTextBox.Clear()
        $valueHash.viewModel.Filter = $null
        Set-Filter -SearchItem $searchItem -SearchList $valueHash.viewModel -SelectedFilter $varHash.FilterChoice.($GuiHash.FilterSelection.SelectedValue.Content)
    })

$guiHash.Window.Add_Activated( {
        $guiHash.FlyOut.Background = ($guiHash.Window.GlowBrush.Color).toString()
    })

$guiHash.Window.Add_Deactivated( {
        $guiHash.FlyOut.Background = ($guiHash.Window.NonActiveGlowBrush.Color).toString()
    })

$guiHash.resultsGrid.Add_Sorting( {
        if ($null -like $guiHash.FilterTextBox.Text -and $guiHash.filterSelection.SelectedValue.Content -match "All|All items") {
            $valueHash.viewModel.Filter = $null
        }
        else {
            $varHash.searchText = $guiHash.filterTextBox.Text
            $global:SearchItem = $varHash.SearchText
            Set-Filter -SearchItem $searchItem -SearchList $valueHash.viewModel -SelectedFilter $varHash.FilterChoice.($GuiHash.FilterSelection.SelectedValue.Content)
        }
    })

$guiHash.actionConfigWindow.Add_ClosingFinished( {
        $guiHash.updateConfigPanel.Visibility = "Collapsed"
        $guiHash.abortConfigPanel.Visibility = "Collapsed"
        $guiHash.rebootConfigPanel.Visibility = "Collapsed"
        $guiHash.toolConfigPanel.Visibility = "Collapsed"
    })

$guiHash.installButton.Add_click( {
        $guiHash.updateConfigPanel.Visibility = "Visible"
        $guiHash.actionConfigWindow.Title = "Configure Update Install"
        $guiHash.actionConfigWindow.IsOpen = $true
    })

$guiHash.updateCancel.Add_Click( {
        $guiHash.actionConfigWindow.IsOpen = $false
    })

$guiHash.updateStart.Add_Click( {
        if ($guiHash.resultsGrid.SelectedItems) {
            $updateList = @($guiHash.resultsGrid.SelectedItems)
            $varHash.completeInstallCount = $updateList.Count
            $varHash.totalInstallCount = $updateList.Count

            $varhash.installArgs = @{
                forceInstall = $guiHash.updateForceStart.IsOn
                reboot       = $guiHash.updateRebootToggle.IsOn
                rerun        = $guiHash.updaterReRun.IsOn
                date         = if ($guiHash.updateDelayToggle.IsOn) { Get-Date($guiHash.installScheduleDate.SelectedDate.toShortDateString(), $(Get-Date($guiHash.installScheduleTime.Text)).toShortTimeString() -join " ") }
                              
            }

            Start-RSJob -InputObject $updateList -ArgumentList $varHash, $logDir, $guiHash, $valueHash -Name Install -Batch InstallBatch -Throttle 2 -ModulesToImport PSWindowsUpdate -ScriptBlock {
                param ($varHash, $logDir, $guiHash, $valueHash)
                $varHash.completeInstallCount--

                if ($_.UpdatesNeeded -eq 0 -and $varHash.installArgs.ForceInstall -ne $true) { exit }

                if ($varHash.installArgs.Date) { $action = 'Scheduling task' }
                else { $action = 'Starting task' }

                $invokeSettings = @{
                    ComputerName = $_.ComputerName
                    Confirm      = $false
                }

                 if ($varHash.installArgs.Date) { $invokeSettings.TriggerDate = $varhash.installArgs.Date }
                 else { $invokeSettings.RunNow = $true }

                $logName = ($_.ComputerName -split '\.')[0] + '.log'

                if ($_.OS -like "*2008*") {
                    $psModBase = $varHash.psModPath.psmodBase8

                    if (!(Test-Path "\\$($_.ComputerName)\c$\Windows\System32\WindowsPowerShell\v1.0\Modules\PSWindowsUpdate")) {
                        Copy-Item -Path $psModBase -Destination "\\$($_.ComputerName)\c$\Windows\System32\WindowsPowerShell\v1.0\Modules\PSWindowsUpdate" -Recurse
                    }

                    $patchStatus = schtasks /s $_.ComputerName /query /FO CSV | ConvertFrom-CSV | Where-Object { $_.TaskName -like "*PSWINDOWSUPDATE*" }

                    if ($patchStatus) {
                        if ($patchStatus.Status -eq 'Running' -and $varHash.forceInstall -ne $true) { exit }
                        else { schtasks /delete /s $_.ComputerName /tn "\PSWindowsUpdate" /f }
                    }

                    if ($varHash.installArgs.Reboot -eq $true) {
                        Invoke-WUJob @invokeSettings -Script "Import-Module PSWindowsUpdate; Start-Transcript -Path `'$logDir\$(Get-Date -format 'MMM yyyy')\$logName`' -Append; Get-WUInstall -AcceptAll -AutoReboot -Verbose; Stop-Transcript | Out-Null; wuauclt /reportnow /detectnow"
                    }
                    else {
                        Invoke-WUJob @invokeSettings -Script "Import-Module PSWindowsUpdate; Start-Transcript -Path `'$logDir\$(Get-Date -format 'MMM yyyy')\$logName`' -Append; Get-WUInstall -AcceptAll -IgnoreReboot -Verbose; Stop-Transcript | Out-Null; wuauclt /reportnow /detectnow"
                    }
                
                    if ($varhash.installArgs.rerun -eq $true) {
                        Invoke-WUJob -TriggerAtStart -Confirm:$false -ComputerName $_.ComputerName -TaskName 'PSWU-Rerun' -Script "Import-Module PSWindowsUpdate; Start-Transcript -Path `'$logDir\$(Get-Date -format 'MMM yyyy')\$logName`' -Append; Get-WUInstall -AcceptAll -IgnoreReboot -Verbose; Stop-Transcript | Out-Null; wuauclt /reportnow /detectnow"
                    }

                }

                else {
                    $psModBase = $varHash.psModPath.psModBase12

                    if (!(Test-Path "\\$($_.ComputerName)\c$\Windows\System32\WindowsPowerShell\v1.0\Modules\PSWindowsUpdate")) {
                        Copy-Item -Path $psModBase -Destination "\\$($_.ComputerName)\c$\Windows\System32\WindowsPowerShell\v1.0\Modules\PSWindowsUpdate" -Recurse
                    }

                    if ((Get-Service -ComputerName $_.ComputerName -Name WinRM).Status -ne 'Running') { Get-Service -ComputerName $_.ComputerName -Name WinRM | Start-Service }

                    $patchStatus = Get-ScheduledTask -CimSession $_.ComputerName -TaskName PSWindowsUpdate -ErrorAction SilentlyContinue

                    if ($patchStatus) {
                        if ($patchStatus.Status -eq 'Running' -and $varhash.forceInstall -ne $true) { exit }
                        else { $patchStatus | Unregister-ScheduledTask -Confirm:$false }
                    }

                    if ($varHash.installArgs.Reboot -eq $true) {
                        Invoke-WUJob @invokeSettings -Script "Import-Module PSWindowsUpdate; Install-WindowsUpdate -AcceptAll -AutoReboot -Verbose *>> `'$logDir\$(Get-Date -Format 'MMM yyyy')\$logName`'; wuauclt /reportnow /detectnow"
                    }
                    else {
                        Invoke-WUJob @invokeSettings -Script "Import-Module PSWindowsUpdate; Install-WindowsUpdate -AcceptAll -IgnoreReboot -Verbose *>> `'$logDir\$(Get-Date -Format 'MMM yyyy')\$logName`'; wuauclt /reportnow /detectnow"
                    }

                    if ($varhash.installArgs.rerun -eq $true) {
                        Invoke-WUJob -TriggerAtStart -Confirm:$false -ComputerName $_.ComputerName -TaskName 'PSWU-Rerun' -Script "Import-Module PSWindowsUpdate; Install-WindowsUpdate -AcceptAll -IgnoreReboot -Verbose *>> `'$logDir\$(Get-Date -Format 'MMM yyyy')\$logName`'; wuauclt /reportnow /detectnow"
                    }

                }

                $comp = $_.ComputerName
                $valueHash.syncedStateHash.$comp.Updater = $action
                #$valueHash.ObjectModel[[Array]::IndexOf($valueHash.objectModel.ComputerName, $comp)].Updater = $action
                $varHash.newRefresh = $true
                    
            }

            Start-RSJob -ArgumentList $guiHash, $varHash -Name 'InstallMonitor' -ScriptBlock {
                param ($guiHash, $varhash)
              
                do {
                    $toolTipString = "Starting install on client $($varHash.totalInstallCount - $varHash.completeInstallCount) of $($varHash.totalInstallCount)"
                    $guiHash.Window.Dispatcher.Invoke([action] { $guiHash.installButton.ToolTip = $toolTipString })
                    Start-Sleep -Seconds 1
                } while ($varHash.completeInstallCount -gt 0)

                $guiHash.Window.Dispatcher.Invoke([action]{$guiHash.installButton.ToolTip = "Configure install on selected clients" })
            }
        }

    $guiHash.actionConfigWindow.IsOpen = $false

    }
)

$guiHash.updateSwitch.Add_IsEnabledChanged( {
        if ($guiHash.updateSwitch.isEnabled -eq $false) {
            $valueHash.viewModel.Refresh()
            $guiHash.updateSwitch.IsEnabled = $true
            $global:SearchItem = $varHash.SearchText

            if ($searchItem) {
                Set-Filter -SearchItem $SearchItem -SearchList $valueHash.ViewModel -SelectedFilter $varHash.FilterChoice.($guiHash.FilterSelection.SelectedValue.Content)
            }
        }          

        if ($guiHash.initialUpdatePanel.Visibility -ne 'Collapsed') { 
            $guiHash.initialUpdatePanel.Visibility = "Collapsed"
            $varHash.initialUpdate = $true 
        }
        
        Get-RSJob -State Completed | Remove-RSJob

    }
)

$guihash.cancelInstallButton.Add_Click( {
        $guiHash.abortConfigPanel.Visibility = "Visible"
        $guiHash.actionConfigWindow.Title = "Abort Update Install"
        $guihash.actionConfigWindow.IsOpen = $true
    }
)

$guiHash.abortStart.Add_Click( {
        if ($guiHash.resultsGrid.SelectedItems) {
            $abortList = @($guiHash.resultsGrid.SelectedItems) 
            $varhash.completeAbortCount = $abortList.Count
            $varhash.totalAbortCount = $abortList.Count
            $varHash.includeScheduled = $guiHash.abortScheduledToggle.IsOn
        
            Start-RsJob -InputObject $abortList -ArgumentList $varHash -Name 'Abort' -Batch 'AbortBatch' -Throttle 3 -ScriptBlock {
            param ($varHash)
            $varHash.completeAbortCount--

                if ($_.OS -like "*2008*") {
                    $patchStatus = schtasks /s $_.ComputerName /query /FO CSV | ConvertFrom-CSV | Where-Object { $_.TaskName -like "*PSWINDOWSUPDATE*" }
                    $rerunStatus = schtasks /s $_.ComputerName /query /FO CSV | ConvertFrom-CSV | Where-Object { $_.TaskName -like "*PSWU-Rerun*" }
                    
                    if ($patchStatus) {
                        if (($patchStatus.State -eq 'Ready' -and ((Get-Date($PatchStatus.'Next Run Time')) -gt (Get-Date))) -and $varHash.includeScheduled -ne $true) { exit }
                        else {
                            schtasks /delete /s $_.ComputerName /tn "\PSwindowsUpdate" /f 
                            if ($rerunStatus) { schtasks /delete /s $_.ComputerName /tn "\PSWU-Rerun" /f }
                        }
                    }
                }

                else {
                    $patchStatus = Get-ScheduledTask -CimSession $_.ComputerName -TaskName PSWindowsUpdate -ErrorAction SilentlyContinue
                    $rerunStatus = Get-ScheduledTask -CimSession $_.ComputerName -TaskName PSWU-Rerun -ErrorAction SilentlyContinue

                    if ($patchStatus) {
                        if ($patchStatus.State -eq 'Ready' -and ((($patchStatus | Get-ScheduledTaskInfo).NextRunTime) -gt (Get-Date)) -and $varHash.includeScheduled -ne $true) { exit }
                        else {
                            $patchStatus | Unregister-ScheduledTask -Confirm:$false
                            if ($rerunStatus) { $rerunStatus | Unregister-ScheduledTask -Confirm:$false }
                        }
                    }
                }
            }

            Start-RSJob -ArgumentList $guiHash, $varHash -Name "AbortMonitor" -ScriptBlock {
                param ($guiHash, $varHash)

                do {
                    $toolTipString = "Aborting updater on client $($varHash.totalAbortCount - $varHash.completeAbortCount) of $($varHash.totalAbortCount)"
                    $guiHash.Window.Dispatcher.Invoke([action] { $guiHash.cancelInstallButton.ToolTip = $toolTipString })
                    Start-Sleep -Seconds 1
                } while ($varHash.completeAbortCount -gt 0)

                $guiHash.Window.Dispatcher.Invoke([action]{$guiHash.cancelInstallButton.ToolTip = "Abort install on selected clients"})            
            }
        }

        $guiHAsh.actionConfigWindow.IsOpen = $false

    }
)

$guiHash.abortCancel.Add_Click( {
        $guihash.actionConfigWindow.IsOpen = $false
    })

$guiHash.rebootButton.Add_Click( {
        $guiHash.rebootConfigPanel.Visibility = "Visible"
        $guiHash.actionConfigWindow.Title = "Post-Installation Reboot"
        $guiHash.actionConfigWindow.IsOpen = $true
    })

$guiHash.rebootCancel.Add_Click( {
        $guiHash.actionConfigWindow.IsOpen = $false
    })




$guiHash.rebootStart.Add_Click( {
        if ($guiHash.resultsGrid.SelectedItems) {
            $rebootList = @($guiHash.resultsGrid.SelectedItems) 
            $varhash.completeRebootCount = $rebootList.Count
            $varhash.totalRebootCount = $rebootList.Count
            $varHash.forceReboot = $guiHash.rebootForceToggle.IsOn
    
            Start-RsJob -InputObject $rebootList -ArgumentList $varHash, $logDir -Name 'Reboot' -Batch 'RebootBatch' -Throttle 2 -ModulesToImport PSwindowsUpdate -ScriptBlock {
                param ($varHash, $logDir)
                $varHash.completeRebootCount--

                if ($_.OS -like "*2008*") { $psModBase = $varHash.psModPath.psModBase8 }
                else { $psModBase = $varHash.psModPath.psModBase12 }

                if (!(Test-Path "\\$($_.ComputerName)\c$\Windows\System32\WindowsPowerShell\v1.0\Modules\PSWindowsUpdate")) {
                    Copy-Item -Path $psModBase -Destination "\\$($_.ComputerName)\c$\Windows\System32\WindowsPowerShell\v1.0\Modules\PSWindowsUpdate" -Recurse
                }

                if ($varHash.forceReboot -eq $true) {  
                    shutdown /r /t 0 /m \\$($_.ComputerName) 
                    "$(Get-Date -format t) - Remote reboot attempted" | Out-File -Append "$logDir\$(Get-Date -Format 'MMM yyyy')\$($($_.ComputerName -split '\.')[0]).log"
                }
                else {
                    try { Get-WURebootStatus -ComputerName $_.ComputerName -AutoReboot -ErrorAction Stop -Verbose *>> "$logDir\$(Get-Date -Format 'MMM yyyy')\$($($_.ComputerName -split '\.')[0]).log" }
                    catch { 
                        Get-WURebootStatus -ComputerName $_.ComputerName -AutoReboot
                        "$(Get-Date -format t) - Remote reboot attempted" | Out-File -Append "$logDir\$(Get-Date -Format 'MMM yyyy')\$($($_.ComputerName -split '\.')[0]).log"
                    }
                }
            }
            Start-RSJob -ArgumentList $guiHash, $varHash -Name "RebootMonitor" -ScriptBlock {
                param ($guiHash, $varHash)

                do {
                    $toolTipString = "Initiating reboot on client $($varHash.totalRebootCount - $varHash.completeRebootCount) of $($varHash.totalRebootCount)"
                    $guiHash.Window.Dispatcher.Invoke([action] { $guiHash.rebootButton.ToolTip = $toolTipString })
                    Start-Sleep -Seconds 1
                } while ($varHash.completeRebootCount -gt 0)

                $guiHash.Window.Dispatcher.Invoke([action]{$guiHash.rebootButton.ToolTip = "Reboot selected clients"})            
            }

            $guiHash.actionConfigWindow.IsOpen = $false
        }
    }
)

$guiHash.mstscButton.Add_Click({
    if ($guiHash.resultsGrid.SelectedItems.Count -eq 1) { mstsc /admin /v $guiHash.resultsGrid.SelectedItem.ComputerName}
    else { $guiHash.resultsGrid.SelectedItems.ComputerName | Connect-RDCMan -RDCManPath $rdcMan }
})

#config validation events
$guiHash.checkServer.Add_Click({
    if (![string]::IsNullOrEmpty($guiHash.settingWSUS.Text)) {
        if (($guiHash.settingWSUS.Text.Split(':')).Count -eq 2 -and
            ($guiHash.settingWSUS.Text.Split(':')[1] -as [int]) -and
            (Test-Connection -Quiet -Count 1 -Computer ($guiHash.settingWSUS.Text.Split(':')[0]))) {

                # Entry was valid - add to hashtable and check for WSUS specific validity
                $varHash.newConfig.Server = $guiHash.settingWSUS.Text.Split(':')[0]
                $varHash.newConfig.Port = $guiHash.settingWSUS.Text.Split(':')[1]
        
            try {
                [void][reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration")            
                $WSUS = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer($varHash.newConfig.Server ,$False,$varHash.newConfig.Port)
                $varHash.newConfig.Groups = [array]($WSUS.GetComputerTargetGroups()).Name
                $guiHash.toolComputerListTypes.ItemsSource = $varHash.newConfig.Groups
                $guiHash.toolServerStatus.Tag = 'True'
            }
        
            catch { $guiHash.toolServerStatus.Tag = 'False' }
        }
        
        else { $guiHash.toolServerStatus.Tag = 'False' }
    }
})

$guiHash.settingWsus.Add_TextChanged({ $guiHash.toolServerStatus.Tag = $null})

$guiHash.checkLogging.Add_Click({
    if (![string]::IsNullOrEmpty($guiHash.settingLogPath.Text)) {
        if ([bool]([uri]$guiHash.settingLogPath.Text).IsUnc -and
            (Test-Path $guiHash.settingLogPath.Text)) {
            $varHash.newConfig.LogPath = $guiHash.settingLogPath.Text
            $guiHash.toolLoggingStatus.Tag = 'True'
        }

        else { $guiHash.toolLoggingStatus.Tag = 'False' }
    }
})

$guiHash.settingLogPath.Add_TextChanged({ $guiHash.toolLoggingStatus.Tag = $null})

$guiHash.headerConfigUpdate.Add_Click({
        $guiHash.toolConfigPanel.Visibility = "Visible"
        $guiHash.actionConfigWindow.Title = "Settings"      

        $guiHash.settingWSUS.Text = $varHash.WSUSServer + ':' + $varHash.WSUSport
        $guiHash.settingLogPath.Text = $varHash.logDir

        $guiHash.actionConfigWindow.IsOpen = $true

})

$guiHash.toolCancel.Add_Click( {
        $guiHash.actionConfigWindow.IsOpen = $false
        $varHash.newConfig = @{}
    })

$guiHash.toolSave.Add_Click({
    $varHash.newConfig.Groups = [array]$guiHash.toolComputerListTypes.SelectedItems
    $varHash.newConfig | ConvertTo-Json | Out-File (Join-Path -Path $PSScriptRoot -ChildPath 'config.json')
    $guiHash.Window.Close()
    $varHash.Finish = $true
    Start-Process -WindowStyle Minimized -FilePath "$PSHOME\powershell.exe" -ArgumentList " -ExecutionPolicy Bypass -NonInteractive -File $(Join-Path $PSScriptRoot -ChildPath 'PSUpdate.ps1')"
    exit
})

$guiHash.Window.ShowDialog() | Out-Null
$varHash.Finish = $true 