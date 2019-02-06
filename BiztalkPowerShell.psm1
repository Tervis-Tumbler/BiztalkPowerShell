function Set-BiztalSQLComputerName {
    param (
        [String]$BiztalSQLComputerName
    )
    $Script:BiztalSQLComputerName = $BiztalSQLComputerName
}

function Get-BiztalkApplication {
    param (
        $ApplicationName = "Tervis.Integration"
    )
    [void] [System.reflection.Assembly]::LoadWithPartialName("Microsoft.BizTalk.ExplorerOM")
    $Catalog = New-Object Microsoft.BizTalk.ExplorerOM.BtsCatalogExplorer
    $Catalog.ConnectionString = "SERVER=$Script:BiztalSQLComputerName;DATABASE=BizTalkMgmtDb;Integrated Security=SSPI"
    $Catalog.Applications[$ApplicationName]
}

function Export-BiztalkState {
    param (
        $BiztalkStatePath = "C:\Biztalk State", 
        [parameter(Mandatory = $true)][string]$Environment
    )

    $TervisIntegration =  Get-BiztalkApplication -ApplicationName "Tervis.Integration"

    $TervisIntegration.SendPorts | 
    Select-Object @{Name="HostName";Expression={$_.PrimaryTransport.SendHandler.name}}, name, status, @{Name="Address";Expression={$_.PrimaryTransport.Address}} | 
    sort Hostname |
    Export-Csv -Path $BiztalkStatePath\SendPorts.csv -NoTypeInformation

    $TervisIntegration.Orchestrations | 
    select @{Name="HostName";Expression={$_.host.name}}, @{Name="Name";Expression={$_.FullName}}, status |
    sort Hostname |
    Export-Csv -Path $BiztalkStatePath\Orchestrations.csv -NoTypeInformation

    $TervisIntegration.ReceivePorts | 
    select -ExpandProperty Receivelocations | 
    select @{Name="HostName";Expression={$_.ReceiveHandler.name}}, name, enable, address   | 
    sort Hostname |
    Export-Csv -Path $BiztalkStatePath\ReceiveLocations.csv -NoTypeInformation
}

function Compare-BiztalkState () {
    param (
        $BiztalkStatePath = "C:\Biztalk State", 
        [parameter(Mandatory = $true)][string]$Environment,
        [switch]$IncludeEqual
    )

    $TervisIntegration =  Get-BiztalkApplication -ApplicationName "Tervis.Integration"

    #Check Send Ports
    $Canonical = import-csv $BiztalkStatePath\SendPorts.csv
    $Current = $TervisIntegration.SendPorts | select @{Name="HostName";Expression={$_.PrimaryTransport.SendHandler.name}}, name, status, @{Name="Address";Expression={$_.PrimaryTransport.Address}} | sort Hostname
    Compare-Object -ReferenceObject $Canonical -DifferenceObject $Current -Property HostName,Name,Status,Address -IncludeEqual:$IncludeEqual | FT

    #Check Orchestrations    
    $Canonical = import-csv $BiztalkStatePath\Orchestrations.csv
    $Current = $TervisIntegration.Orchestrations | select @{Name="HostName";Expression={$_.host.name}}, @{Name="Name";Expression={$_.FullName}}, status | sort Hostname
    Compare-Object -ReferenceObject $Canonical -DifferenceObject $Current -Property HostName,Name,Status -IncludeEqual:$IncludeEqual | FT

    #Check Receive Locations
    $Canonical = import-csv $BiztalkStatePath\ReceiveLocations.csv
    $Current = $TervisIntegration.ReceivePorts | select -ExpandProperty Receivelocations | select @{Name="HostName";Expression={$_.ReceiveHandler.name}}, name, enable, address | sort Hostname
    Compare-Object -ReferenceObject $Canonical -DifferenceObject $Current -Property HostName,Name,Enable,Address -IncludeEqual:$IncludeEqual | FT
}


function Import-BiztalkState {
    param (
        $BiztalkStatePath = "C:\Biztalk State", 
        [parameter(Mandatory = $true)][string]$Environment
    )
    $TervisIntegration =  Get-BiztalkApplication -ApplicationName "Tervis.Integration"    

    #Set Send Ports state to match file
    $Canonical = import-csv $BiztalkStatePath\SendPorts.csv
    $Current = $TervisIntegration.SendPorts | select @{Name="HostName";Expression={$_.PrimaryTransport.SendHandler.name}}, name, status | sort HostName
    $IntersectionOfCurrentAndCanonical = $Canonical | where { $_.name -in $Current.name }
    $ItemsOutOfSynch = Compare-Object -ReferenceObject $IntersectionOfCurrentAndCanonical -DifferenceObject $Current -Property HostName,Name,Status
    $ItemsToChange = $ItemsOutOfSynch | where {$_.SideIndicator -eq "<="}
    $TervisIntegration.SendPorts | where {$_.name -in $ItemsToChange.Name} | 
        % { $SendPort = $_; $SendPort.Status = [Microsoft.BizTalk.ExplorerOM.PortStatus] ($ItemsToChange | where {$_.name -eq $SendPort.Name} | select -expandproperty Status)}
    $Catalog.SaveChanges();

    #Set Send Ports that were not in the state file to be bound
    #$TervisIntegration.SendPorts | where {$_.name -notin $ItemsToChange.Name} | 
    #    %{ $_.Status = [Microsoft.BizTalk.ExplorerOM.PortStatus] "Bound" }
    #$Catalog.SaveChanges();


    #Set Orchestrations state to match file
    $Canonical = import-csv $BiztalkStatePath\Orchestrations.csv
    $Current = $TervisIntegration.Orchestrations | select @{Name="HostName";Expression={$_.host.name}}, @{Name="Name";Expression={$_.FullName}}, Status | sort Hostname
    $IntersectionOfCurrentAndCanonical = $Canonical | where { $_.name -in $Current.name }
    $ItemsOutOfSynch = Compare-Object -ReferenceObject $IntersectionOfCurrentAndCanonical -DifferenceObject $Current -Property HostName,Name,Status
    $ItemsToChange = $ItemsOutOfSynch | where {$_.SideIndicator -eq "<="}
    $TervisIntegration.Orchestrations | where {$_.FullName -in $ItemsToChange.Name} | 
        % { $Orchestration = $_; $Orchestration.Status = [Microsoft.BizTalk.ExplorerOM.OrchestrationStatus] ($ItemsToChange | where {$_.name -eq $Orchestration.FullName} | select -expandproperty Status)}
    $Catalog.SaveChanges();
    
    #Set orchestrations not in the state file to the Unenlisted state
    #$TervisIntegration.Orchestrations | where {$_.FullName -notin $ItemsToChange.Name} | 
    #    %{ $_.Status = [Microsoft.BizTalk.ExplorerOM.OrchestrationStatus] "Unenlisted" }
    #$Catalog.SaveChanges();


    #Set Receive Locations state to match file
    $Canonical = import-csv $BiztalkStatePath\ReceiveLocations.csv | Select HostName, Name, @{Name="Enable";Expression={[boolean]::Parse($_.Enable)}}
    $Current = $TervisIntegration.ReceivePorts | select -ExpandProperty ReceiveLocations | select @{Name="HostName";Expression={$_.ReceiveHandler.name}}, Name, Enable | sort Hostname
    $IntersectionOfCurrentAndCanonical = $Canonical | where { $_.name -in $Current.name }
    $ItemsOutOfSynch = Compare-Object -ReferenceObject $IntersectionOfCurrentAndCanonical -DifferenceObject $Current -Property HostName,Name,Enable
    $ItemsToChange = $ItemsOutOfSynch | where {$_.SideIndicator -eq "<="}
    $TervisIntegration.ReceivePorts | select -ExpandProperty ReceiveLocations | where {$_.Name -in $ItemsToChange.Name} | 
        % { $ReceiveLocation = $_; $ReceiveLocation.Enable = [boolean]::Parse(($ItemsToChange | where {$_.name -eq $ReceiveLocation.Name} | select -expandproperty Enable))}
    $Catalog.SaveChanges();
}

function Test-BiztalkState () {
    param (
        $CorrectBiztalkStatePath = "C:\Correct Biztalk State",
        [parameter(Mandatory = $true)][string]$Environment,
        $From,
        $To,
        $SMTPServer
    )

    $ComparisonResults = Compare-BiztalkState -Environment $Environment -BiztalkStatePath $CorrectBiztalkStatePath
    if ($ComparisonResults) { 
        Send-TervisMailMessage -From $From -to $To -subject "Incorrect Biztalk State Detected" -Body ($ComparisonResults | FT -autosize | out-string -Width 200) 
    }
}
