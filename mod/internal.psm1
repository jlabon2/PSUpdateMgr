function Invoke-WSUSReport {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string[]]$ComputerName = $env:COMPUTERNAME
    )
    
    process {
        if ($computerName -ne $env:COMPUTERNAME -and !(Test-Connection -Count 1 -Quiet -ComputerName $ComputerName)) {
            Write-Error -Message "$ComputerName could not be reached"
        }

        elseif (Invoke-Command -ComputerName $ComputerName -ScriptBlock { ((Get-Item "HKLM:Software\Policies\Microsoft\Windows\WindowsUpdate" | Get-ItemProperty).WuServer -like "*SCCM*") }) {
            Write-Error -Message "Client is configured to use SCCM; must be using WSUS directly to continue."
        }

        else {
            Invoke-Command -ComputerName $ComputerName -ScriptBlock {
                wuauclt /reportnow /detectnow
                (New-Object -ComObject "Microsoft.Update.Session").CreateUpdateSearcher().Search($null).Updates
                Start-Sleep -Seconds 1
                wuauclt /reportnow /detectnow
            } | Out-Null
        }
    }
}

function Get-TimeDifference {
[CmdletBinding()]
    param (
         [Parameter(Mandatory = $true)][datetime]$StartTime,
         [datetime]$EndTime
    )

    if (!$EndTime) { $EndTime = Get-Date }

    $diffString = switch (New-TimeSpan -Start $StartTime -End $EndTime) {
        {$_.Days -ge 30}   { "$([math]::Floor($_.Days/30)) month"; break}
        {$_.Days -ge 1}    { "$($_.Days) day"; break}
        {$_.Hours -ge 1}   { "$($_.Hours) hour"; break}
        {$_.Minutes -ge 1} { "$($_.Minutes) minute"; break}
        {$_.Seconds -ge 1} { "$($_.Seconds
        ) second"; break}
    }


    if ($diffString -notlike "1 *") {$diffString = $diffString + 's'}
    
    $diffString + ' ' + 'ago'
   
}

function Get-TimeSpanStringValue {
    param ([string]$timeDiffString) 

    $tagValue = switch -Wildcard ($timeDiffString) {
        '*years*'  {8; break}
        '*month*' {7; break}
        '*days*'   {6; break}
        '*day*'    {5; break}
        '*hours*'  {4; break}
        '*hour*'   {3; break}
        '*minutes*'{2; break}
        default    {1}

    }

    [string]$tagValue
}

function Set-Filter {
    [CmdletBinding()]
    param (
        [string]$SearchItem,
        [System.Windows.Data.CollectionView]$SearchList,
        [string]$SelectedFilter
    )

    $valueHash.viewModel.Filter = $null
    
    if (($searchItem -ne $null -or $guiHash.filterTextBox.Text -ne $null) -and ($selectedFilter -match "All|AllItems")) {
        $valueHash.viewModel.Filter = { param ($searchList) $searchList.ComputerName | Where-Object { $_ -match $searchItem } }
    }

    elseif ($SelectedFilter -eq "RebootNeeded") {
        $valueHash.viewModel.Filter = { param ($searchList) $searchList | Where-Object { $_.RebootNeeded -eq $true -and $_.ComputerName -match $searchItem } }
    }

    elseif ($SelectedFilter -eq "UpdateNeeded") {
        $valueHash.viewModel.Filter = { param ($searchList) $searchList | Where-Object { $_.UpdatesNeeded -ge 1 -and $_.ComputerName -match $searchItem } }
    }
}

function Connect-RDCMan {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName, Position = 0)][string[]]$ComputerName,
        [Parameter(Mandatory = $true)][string]$RDCManPath
    )

    begin {
        $xmlString = [String]::Empty

        $stringHeader = @'
<?xml version="1.0" encoding="utf-8"?>
<RDCMan programVersion="2.82" schemaVersion="3">
  <file>
    <credentialsProfiles />
    <properties>
      <expanded>True</expanded>
        <name>Computer List</name>
    </properties>
'@   

        $xmlString = $xmlString + $stringHeader
    
    }
    process {
        $stringItem = @"
      <server>
        <properties>
          <name>$ComputerName</name>
        </properties>
      </server>
"@

        $xmlString = $xmlString + $stringItem

        if ($connectList) { [string]$connectList = $connectList + ',' + $computerName }
        else { [string]$connectList = $computerName }
    }

    end {
        $xmlFooter = @'
        </file>
        <connected />
        <favorites />
        <recentlyUsed />
      </RDCMan>
'@

        $xmlString = $xmlString + $xmlFooter
        $xmlString | Out-File -Encoding utf8 -FilePath (Join-Path $env:TEMP -ChildPath "temp.rdg")

        Start-Process $rdcManPath -ArgumentList "$($(Join-Path $env:TEMP -ChildPath "temp.rdg")) /c $connectList"
    }
}

function Set-WindowVisibility {
    if ($host.name -eq 'ConsoleHost') {
        $SW_HIDE, $SW_SHOW = 0, 5
        $TypeDef = '[DllImport("User32.dll")]public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);'
        Add-Type -MemberDefinition $TypeDef -Namespace Win32 -Name Functions
        $hWnd = (Get-Process -Id $PID).MainWindowHandle
        [Win32.Functions]::ShowWindow($hWnd, $SW_HIDE) | Out-Null
    }
}

function Set-WPFControls {
    param (
        [Parameter(Mandatory)]$XAMLPath,
        [Parameter(Mandatory)]$WindowName,
        [Parameter(Mandatory)][Hashtable]$TargetHash
    ) 

    $inputXML = Get-Content -Path $XAMLPath
    
    $inputXML = $inputXML -replace 'mc:Ignorable="d"', '' -replace "x:Class=`"$($WindowName)`"" -replace 'x:N', 'N' -replace '^<Win.*', '<Window'
    [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
    [xml]$XAML = $inputXML

    $xmlReader = (New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $XAML)
    
    try { $TargetHash.Window = [Windows.Markup.XamlReader]::Load($xmlReader) }
    catch { Write-Warning -Message "Unable to parse XML, with error: $($Error[0])" }

    ## Load each named control into PS hashtable
    foreach ($controlName in ($XAML.SelectNodes('//*[@Name]').Name)) { $TargetHash.$controlName = $TargetHash.Window.FindName($controlName) }

}

function New-HashTables {
    # Stores values log0ging missing or errored items during init
    $global:valueHash = [hashtable]::Synchronized(@{ })

    # Stores config values imported JSON, during config, or both
    $global:valueHash.syncedHash = [hashtable]::Synchronized(@{ })

    $global:valueHash.syncedStateHash = [hashtable]::Synchronized(@{ })

    # Stores WPF controls
    $global:guiHash = [hashtable]::Synchronized(@{ })

    # Stores config'd vars
    $global:varHash = [hashtable]::Synchronized(@{ })

    # Changed config values
    $varHash.newConfig = @{}
    
}