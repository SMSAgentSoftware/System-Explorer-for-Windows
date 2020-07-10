###requires -Version 2 -Modules CimCmdlets

#region Functions

# Function to read registry values
function Read-Registry 
{
    if (Test-Path -Path 'HKLM:\SOFTWARE\Wow6432Node\SMSAgent\System Explorer')
    {
        $regsql = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\SMSAgent\System Explorer' -Name SQLServer | Select-Object -ExpandProperty SQLServer
        $regdb = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\SMSAgent\System Explorer' -Name Database | Select-Object -ExpandProperty Database
        $enabled = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\SMSAgent\System Explorer' -Name ConfigMgrEnabled | Select-Object -ExpandProperty ConfigMgrEnabled
    }

    if (Test-Path -Path 'HKLM:\SOFTWARE\SMSAgent\System Explorer')
    {
        $regsql = Get-ItemProperty -Path 'HKLM:\SOFTWARE\SMSAgent\System Explorer' -Name SQLServer | Select-Object -ExpandProperty SQLServer
        $regdb = Get-ItemProperty -Path 'HKLM:\SOFTWARE\SMSAgent\System Explorer' -Name Database | Select-Object -ExpandProperty Database
        $enabled = Get-ItemProperty -Path 'HKLM:\SOFTWARE\SMSAgent\System Explorer' -Name ConfigMgrEnabled | Select-Object -ExpandProperty ConfigMgrEnabled
    }

    $Window.SQLServer = $regsql
    $Window.Database = $regdb
    $Window.ConfigMgrEnabled = $enabled
}

# Function to create a seperate runspace to run the main code
function Start-RSJob
{
    param(
        [parameter(Mandatory = $True,Position = 0)]
        [ScriptBlock]$Code,
        [parameter()]
        $Arguments,
        [parameter()]
        $Functions,
        $Runspaces
    )
    $hash = @{}
    $InitialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $Scope = [System.Management.Automation.ScopedItemOptions]::AllScope
    foreach ($Function in $Functions)
    {
        $name = $Function.Split('\')[1]
        $Def = Get-Content $Function -ErrorAction Stop
        $SessionStateFunction = New-Object -TypeName System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList $name, $Def, $Scope, $null
        $InitialSessionState.Commands.Add($SessionStateFunction)
    }
    
    $Runspace = [runspacefactory]::CreateRunspace($InitialSessionState)
    $PowerShell = [powershell]::Create()
    $PowerShell.runspace = $Runspace
    $Runspace.Open()
    [void]$PowerShell.AddScript($Code)
    foreach ($Argument in $Arguments)
    {
        [void]$PowerShell.AddArgument($Argument)
    }
    $hash.PowerShell = $PowerShell
    $job = $PowerShell.BeginInvoke()
    $Hash.Job = $job
    $Runspaces.RS = $Hash
}

# Function to create a custom array where there are multiple instances in a class.  This allows each instance to be displayed horizontally instead of vertically, where the instances are difficult to distinguish
function Prepare-CustomArray 
{
    Param ($ListItem,$Hardware)
    # Put the hardware info into a new variable
    $Source = $Hardware.GetEnumerator() | Where-Object -FilterScript {
        $_.Name -match $ListItem
    }

    # Find unique property names
    $Names = $Source.Value.Name | Select-Object -Unique
    
    # Find Count of items
    $Count = $Source.Value.count   
    
    # Create an array container
    $RowItems = @()

    # Create a counter object
    $i = 0

    # Create a new object for each instance by adding the value for each item, then add the object to the array
    do 
    {
        $Source | ForEach-Object -Process {
            $Item = $_
            $obj = New-Object -TypeName psobject
            $Names | ForEach-Object -Process {
                $name = $_
                # Check that the named property exists the instance...
                if ($Item.Value.Name -contains $_)
                {
                    $Value = ($Item.Value | Where-Object -FilterScript {
                            $_.Name -eq $name
                    }).Value
                    $i ++
                }
                # If not, assign a null value
                Else 
                {
                    $Value = $null
                }
                Add-Member -InputObject $obj -MemberType NoteProperty -Name $_ -Value $Value
            }
            $RowItems += $obj
        }
    }
    Until ($i -eq $Count)
    
    # Return the array
    $RowItems = $RowItems | Sort-Object -Property Caption
    Return $RowItems
}

# Function to convert data into a datatable, which displays better in a WPF datagrid object   
function ConvertTo-DataTable 
{
    [CmdletBinding()] 
    param([Parameter(Position = 0, Mandatory = $True, ValueFromPipeline = $True)] [PSObject[]]$InputObject) 
 
    Begin 
    { 
        $dt = New-Object -TypeName Data.datatable   
        $First = $True  
    } 
    Process 
    { 
        foreach ($object in $InputObject) 
        { 
            $DR = $dt.NewRow()   
            foreach($property in $object.PsObject.get_properties()) 
            {   
                if ($First) 
                {   
                    $Col = New-Object -TypeName Data.DataColumn   
                    $Col.ColumnName = $property.Name.ToString()   
                    $dt.Columns.Add($Col) 
                }   
                if ($property.Gettype().IsArray) 
                {
                    $DR.Item($property.Name) = $property.value | ConvertTo-Xml -As String -NoTypeInformation -Depth 1
                }   
                else 
                { 
                    try
                    {
                        $DR.Item($property.Name) = $property.value
                    }
                    Catch 
                    {

                    }
                } 
            }   
            $dt.Rows.Add($DR)   
            $First = $false 
        } 
    }  
      
    End 
    { 
        Write-Output @(,($dt)) 
    }  
}

# Function to Switch displayed item names with their corresponding WMI class
function Switch-ClassNames 
{
    param ($ListItem)
    # Set WMI Class name based on the list item name
    switch ($ListItem)
    {
        'Operating System' 
        {
            $Script:Class = 'OperatingSystem'
        }
        'Computer System' 
        {
            $Script:Class = 'ComputerSystem'
        }
        'Disk Drive' 
        {
            $Script:Class = 'DiskDrive'
        }
        'Disk Partition' 
        {
            $Script:Class = 'DiskPartition'
        }
        'Logical Disk' 
        {
            $Script:Class = 'LogicalDisk'
        }
        'Network Adapter' 
        {
            $Script:Class = 'NetworkAdapter'
        }
        'Network Adapter Configuration' 
        {
            $Script:Class = 'NetworkAdapterConfiguration'
        }
        'Physical Memory' 
        {
            $Script:Class = 'PhysicalMemory'
        }
        'Sound Device' 
        {
            $Script:Class = 'SoundDevice'
        }
        'USB Device' 
        {
            $Script:Class = 'USBDevice'
        }
        'Video Controller' 
        {
            $Script:Class = 'VideoController'
        }
        'User Profiles' 
        {
            $Script:Class = 'UserProfile'
        }
        'Shares' 
        {
            $Script:Class = 'Share'
        }
        'Printers' 
        {
            $Script:Class = 'Printer'
        }
        'Installed Windows Features' 
        {
            $Script:Class = 'OptionalFeature'
        }
        'Network Login Profiles' 
        {
            $Script:Class = 'NetworkLoginProfile'
        }
        'Mapped Drives' 
        {
            $Script:Class = 'NetworkConnection'
        }
        default 
        {
            $Script:Class = $ListItem
        }
    }
}

# Function to load the data from WMI
function Load-Data 
{
    param ($Credentials,$ListItem,$ComputerName,$Class,$Window)
    # Load the data for the selected class
    try
    {
        # Test if credentials have been passed, if so create a Cim session to use them...
        if ($Credentials -ne 'none')
        { 
            $option = New-CimSessionOption -Protocol Wsman
            $Script:Session = New-CimSession -SessionOption $option -ComputerName $ComputerName -Credential $Credentials

            if ($ListItem -eq 'Installed Windows Features')
            {
                $Script:TempVar = Get-CimInstance -CimSession $Session -ClassName Win32_$Class -Property * -Filter 'InstallState = 1' -ErrorAction Stop
            }
            Else 
            {
                $Script:TempVar = Get-CimInstance -CimSession $Session -ClassName Win32_$Class -Property * -ErrorAction Stop
            }
        }
        # ...otherwise use the computer name directly
        Else 
        { 
            if ($ListItem -eq 'Installed Windows Features')
            {
                $Script:TempVar = Get-CimInstance -ComputerName $ComputerName -ClassName Win32_$Class -Property * -Filter 'InstallState = 1' -ErrorAction Stop
            }
            Else 
            {
                $Script:TempVar = Get-CimInstance -ComputerName $ComputerName -ClassName Win32_$Class -Property * -ErrorAction Stop
            }
        }
    }
    catch
    {
        $myerror = $_
        $obj = New-Object -TypeName psobject
        Add-Member -InputObject $obj -MemberType NoteProperty -Name Error -Value $myerror
        $Window.GUI.Dispatcher.Invoke(
            [action]{
                $Window.dataGrid.ItemsSource = [array]$obj
        })
        Return
    }
}

# Function to add the list of populated properties for each class and instance to a hash table, appending a unique number to each instance 
function Populate-HardwareHashTable 
{
    param($ListItem,$Hardware,$TempVar)

    $i = -1
    $TempVar | ForEach-Object -Process {
        $i ++
        $x = $ListItem
        $x = $x + $i
        # Define the properties to get for specific classes...
        If ($ListItem -eq 'User Profiles') 
        {
            $Hardware.$x = $_.CimInstanceProperties |
            Select-Object -Property Name, Value |
            Where-Object -FilterScript {
                $_.Value -ne $null -and $_.Value -ne '' -and $_.Name -in ('HealthStatus', 'LastUseTime', 'LocalPath', 'SID', 'Loaded', 'Special', 'RefCount')
            }
        }
        # ...Else define unneeded classes and values to filter out
        Else 
        {
            $Hardware.$x = $_.CimInstanceProperties |
            Select-Object -Property Name, Value |
            Where-Object -FilterScript {
                $_.Value -ne $null -and $_.Value -ne '' -and $_.Name -notin ('CreationClassName', 'CSCreationClassName', 'CSName', 'SystemName', 'SystemCreationClassName')
            }
        }
    }
}

# Function to reformat certain values and dates and expand property arrays into strings
function Reformat-Values 
{
    param($Hardware,$ListItems)

    # Operating System | Rounding values
    $OSRoundingList = @(
        'FreePhysicalMemory'
        'FreeSpaceInPagingFiles'
        'FreeVirtualMemory'
        'MaxProcessMemorySize'
        'SizeStoredInPagingFiles'
        'TotalVirtualMemorySize'
        'TotalVisibleMemorySize'
    )
    $OSRoundingList | ForEach-Object -Process {
        $Item = $_
        $HashEntryName = $Hardware.GetEnumerator() |
        Where-Object -FilterScript {
            $_.Name -match 'Operating System' -or $_.Name -match 'OperatingSystem'
        } |
        Select-Object -ExpandProperty Name
        ($Hardware.$HashEntryName | Where-Object -FilterScript {
                $_.Name -eq $Item
        }).Value = "$([math]::Round((($Hardware.$HashEntryName | Where-Object -FilterScript {$_.Name -eq $Item
        }).Value / 1MB),1)) GB"
    }

    # Operating System | reformat dates
    $OSDateList = @(
        'InstallDate'
        'LastBootUpTime'
        'LocalDateTime'
    )
    $OSDateList | ForEach-Object -Process {
        $Item = $_
        $HashEntryName = $Hardware.GetEnumerator() |
        Where-Object -FilterScript {
            $_.Name -match 'Operating System' -or $_.Name -match 'OperatingSystem'
        } |
        Select-Object -ExpandProperty Name
        ($Hardware.$HashEntryName | Where-Object -FilterScript {
                $_.Name -eq $Item
        }).Value = (($Hardware.$HashEntryName | Where-Object -FilterScript {
                    $_.Name -eq $Item
        }).Value).toString('dd/MMM/yyyy hh:mm')
    }

    # Computer System | Rounding values
    $HashEntryName = $Hardware.GetEnumerator() |
    Where-Object -FilterScript {
        $_.Name -match 'Computer System' -or $_.Name -match 'ComputerSystem'
    } |
    Select-Object -ExpandProperty Name
    $HashEntryName | ForEach-Object -Process {
        $name = $_
        ($Hardware.$name  | Where-Object -FilterScript {
                $_.Name -eq 'TotalPhysicalMemory'
        }).Value = "$([math]::Round((($Hardware.$name  | Where-Object -FilterScript {$_.Name -eq 'TotalPhysicalMemory'
        }).Value / 1GB))) GB"
    }

    # Disk Drive | Rounding values
    $HashEntryName = $Hardware.GetEnumerator() |
    Where-Object -FilterScript {
        $_.Name -match 'Disk Drive' -or $_.Name -match 'DiskDrive'
    } |
    Select-Object -ExpandProperty Name
    $HashEntryName | ForEach-Object -Process {
        $name = $_
        ($Hardware.$name | Where-Object -FilterScript {
                $_.Name -eq 'Size'
        }).Value = "$([math]::Round((($Hardware.$name | Where-Object -FilterScript {$_.Name -eq 'Size'
        }).Value / 1GB))) GB"
    }

    # Disk Partition | Rounding values
    $HashEntryName = $Hardware.GetEnumerator() |
    Where-Object -FilterScript {
        $_.Name -match 'Disk Partition' -or $_.Name -match 'DiskPartition'
    } |
    Select-Object -ExpandProperty Name
    $HashEntryName | ForEach-Object -Process {
        $name = $_
        ($Hardware.$name | Where-Object -FilterScript {
                $_.Name -eq 'Size'
        }).Value | ForEach-Object -Process {
            $Value = $_
            if ($Value -lt 1000000000)
            {
                ($Hardware.$name | Where-Object -FilterScript {
                        $_.Name -eq 'Size' -and $_.Value -eq $Value
                }).Value = "$([math]::Round(($Value / 1MB))) MB"
            }
            Else 
            {
                ($Hardware.$name | Where-Object -FilterScript {
                        $_.Name -eq 'Size' -and $_.Value -eq $Value
                }).Value = "$([math]::Round(($Value / 1GB))) GB"
            }
        }
    }

    # Logical Disk | Rounding values
    $LogicalDiskRoundingList = @(
        'FreeSpace'
        'Size'
    )
    $LogicalDiskRoundingList | ForEach-Object -Process {
        $Item = $_
        $HashEntryName = $Hardware.GetEnumerator() |
        Where-Object -FilterScript {
            $_.Name -match 'Logical Disk' -or $_.Name -match 'LogicalDisk'
        } |
        Select-Object -ExpandProperty Name
        $HashEntryName | ForEach-Object -Process {
            $name = $_
            $Hardware.$name |
            Where-Object -FilterScript {
                $_.Name -eq $Item
            } |
            ForEach-Object -Process {
                $Value = $_.Value
                ($Hardware.$name  |
                    Where-Object -FilterScript {
                        $_.Name -eq $Item -and $_.Value -eq $Value
                }).Value = "$([math]::Round(($Value / 1GB),1)) GB"
            }
        }
    }

    # Physical Memory | Rounding values
    $HashEntryName = $Hardware.GetEnumerator() |
    Where-Object -FilterScript {
        $_.Name -match 'Physical Memory' -or $_.Name -match 'PhysicalMemory'
    } |
    Select-Object -ExpandProperty Name
    $HashEntryName | ForEach-Object -Process {
        $name = $_
        ($Hardware.$name  | Where-Object -FilterScript {
                $_.Name -eq 'Capacity'
        }).Value = "$([math]::Round((($Hardware.$name  | Where-Object -FilterScript {$_.Name -eq 'Capacity'
        }).Value / 1GB))) GB"
    }

    # Volume | Rounding values
    $VolumeRoundingList = @(
        'FreeSpace'
        'Capacity'
    )
    $VolumeRoundingList | ForEach-Object -Process {
        $Item = $_
        $HashEntryName = $Hardware.GetEnumerator() |
        Where-Object -FilterScript {
            $_.Name -match 'Volume'
        } |
        Select-Object -ExpandProperty Name
        $HashEntryName | ForEach-Object -Process {
            $name = $_
            $Hardware.$name |
            Where-Object -FilterScript {
                $_.Name -eq $Item
            } |
            ForEach-Object -Process {
                $Value = $_.Value
                ($Hardware.$name |
                    Where-Object -FilterScript {
                        $_.Name -eq $Item -and $_.Value -eq $Value
                }).Value = "$([math]::Round(($Value / 1GB),1)) GB"
            }
        }
    }

    # Video Controller | Rounding values
    'AdapterRAM' | ForEach-Object -Process {
        $Item = $_
        $HashEntryName = $Hardware.GetEnumerator() |
        Where-Object -FilterScript {
            $_.Name -match 'Video Controller' -or $_.Name -match 'VideoController'
        } |
        Select-Object -ExpandProperty Name
        $HashEntryName | ForEach-Object -Process {
            $name = $_
            $Hardware.$name |
            Where-Object -FilterScript {
                $_.Name -eq $Item
            } |
            ForEach-Object -Process {
                $Value = $_.Value
                ($Hardware.$name |
                    Where-Object -FilterScript {
                        $_.Name -eq $Item -and $_.Value -eq $Value
                }).Value = "$([math]::Round(($Value / 1GB),1)) GB"
            }
        }
    }

    # Convert properties with array elements to a string
    $ListItems | ForEach-Object -Process {
        $name = $_
        $HashEntryName = $Hardware.GetEnumerator() |
        Where-Object -FilterScript {
            $_.Name -match $name
        } |
        Select-Object -ExpandProperty Name
        $HashEntryName | ForEach-Object -Process {
            $Hardware.$_ | ForEach-Object -Process {
                try
                {
                    If (($_.Value.GetType()).IsArray)
                    {
                        $_.Value = "$($_.Value)"
                    }
                }
                Catch 
                {

                }
            }
        }
    } 
}

# Function to process the data into a datatable and display the results in the GUI
function Display-DataInGUI 
{
    param($Source,$Instances,$ListItem,$Hardware,$Window,$Description)
    if ($Source -eq 'Client WMI' -or $Source -eq 'Client WMI (Full)')
    {
        # If more than one instance exists, display horizontally...
        if ($Instances.$ListItem -gt 0)
        {
            $Array = Prepare-CustomArray -ListItem $ListItem -Hardware $Hardware
            # If array is blank, notify of no results found
            if ($Array -eq $null)
            {
                $obj = New-Object -TypeName psobject
                Add-Member -InputObject $obj -MemberType NoteProperty -Name Info -Value 'No results'
                $Window.GUI.Dispatcher.Invoke(
                    [action]{
                        $Window.dataGrid.ItemsSource = [array]$obj
                        $Window.ClassDescription.Text = $Description
                })
            }
            Else 
            {
                $Hardware.Table = $Array | ConvertTo-DataTable
                $Window.GUI.Dispatcher.Invoke(
                    [action]{
                        $Window.dataGrid.ItemsSource = $Hardware.Table.DefaultView
                        $Window.ClassDescription.Text = $Description
                })
            }
        }
        # Else display vertically 
        Else 
        {
            $HashEntryName = $Hardware.GetEnumerator() |
            Where-Object -FilterScript {
                $_.Name -match $ListItem
            } |
            Select-Object -ExpandProperty Name
            # If no name found in hash table, notify of no results found
            if ($HashEntryName -eq $null)
            {
                $obj = New-Object -TypeName psobject
                Add-Member -InputObject $obj -MemberType NoteProperty -Name Info -Value 'No results'
                $Window.GUI.Dispatcher.Invoke(
                    [action]{
                        $Window.dataGrid.ItemsSource = [array]$obj
                        $Window.ClassDescription.Text = $Description
                })
            }
            Else 
            {
                $Hardware.Table = $Hardware.$HashEntryName |
                Sort-Object -Property Name |
                ConvertTo-DataTable
                $Window.GUI.Dispatcher.Invoke(
                    [action]{
                        $Window.dataGrid.ItemsSource = $Hardware.Table.DefaultView
                        $Window.ClassDescription.Text = $Description
                })
            }
        }
    }
}

# Function to check if the target system is online and if alternate credentials are required to access it
function Check-SystemAndCredentials
{
    Param($ComputerName,$Window)

        # Check that the system is online
        If (!(Test-Connection -ComputerName $ComputerName -Count 1 -Quiet))
        {
            $Window.Online = $false
            return
        }

        # Test whether credentials are required
        try
        {
            Get-CimInstance -ComputerName $ComputerName -ClassName win32_operatingsystem -ErrorAction Stop
            $Window.Online = $True
        }
        Catch
        {
            if ($Error[0] | Select-String -Pattern 'Access is denied')
            {
                # Prompt for the right credentials to use
                try
                {
                    $Credentials = $Window.host.ui.PromptForCredential('Credentials required', "Access is denied to $($ComputerName).  Enter credentials to connect.", '', '')
                    $Window.Credentials = $Credentials
                    $Window.Online = $True
                    
                }
                Catch {}
            }
            Else
            {
                # Some other error?  Return the error in the datagrid and exit the code
                $myerror = $($Error[0].Exception.Message)
                $obj = New-Object -TypeName psobject
                Add-Member -InputObject $obj -MemberType NoteProperty -Name Error -Value $myerror
                $Window.GUI.Dispatcher.Invoke(
                    [action]{
                        $Window.dataGrid.ItemsSource = [array]$obj
                        $Window.dataGrid.Foreground = 'Red'
                })
                $Window.Online = $false
                return
            }
        }
}

# Function holding the commands that run when the Connect button is clicked
function Thunderbirds-AreGo 
{
    param ($Runspaces,$Window)
    # Remove any previously created runspaces
    $Runspaces.GetEnumerator() | foreach {
    $_.Value.PowerShell.Stop()
    $_.Value.PowerShell.Runspace.Dispose()
    }
    $Runspaces.Clear()

    # Create new hashtables to store the hardware and instance data
    $Global:Hardware = [hashtable]::Synchronized(@{})
    $Global:Instances = [hashtable]::Synchronized(@{})
    
    # Remove credentials if they are present from a previous run
    if ($Window.Credentials)
    {
        $Window.Remove('Credentials')
    }
    
    # Remove online status if present from a previous run
    if ($Window.Online -eq $True -or $Window.Online -eq $false)
    {
        $Window.Remove('Online')
    }

    # Set the computer name as a variable
    $Script:ComputerName = $Window.ComputerName.Text

    # Read registry values
    Read-Registry

    # Check that the system is online and the right credentials are being used
    #if ($Window.Source.SelectedItem -match 'Client')
    #{
        Check-SystemAndCredentials -ComputerName $ComputerName -Window $Window
        do 
        {

        }
        until ($Window.Online -eq $True -or $Window.Online -eq $false)
    #}

    # Get resource ID if the system exists in the ConfigMgr database
    if ($ResourceID)
    {
        Remove-Variable -Name ResourceID -Force -Scope Global
    }
    Get-ResourceID
    
    # If system is online, enable the listview, advise user to select category
    if ($Window.Online -and $Window.ConfigMgrEnabled -eq 'True')
    { 
        $items = @()
        $Window.Source.Items | foreach {
            $items += $_
            }
        $items | foreach {
            [void]$Window.Source.Items.Remove($_)
        }
        $obj = New-Object -TypeName psobject
        Add-Member -InputObject $obj -MemberType NoteProperty -Name Info -Value 'Select a category'
        $Window.dataGrid.ItemsSource = [array]$obj
        $Window.dataGrid.Foreground = 'Black'
        $Window.listView.IsEnabled = $True
        $Window.Source.IsEnabled = $True
        'Client WMI', 'ConfigMgr DB', 'Client WMI (Full)' | ForEach-Object {
        [void]$Window.Source.Items.Add("$_")
        }
        $Window.Source.SelectedIndex = 0
    }

    If ($Window.Online -and ($Window.ConfigMgrEnabled -eq 'False' -or !$Window.ConfigMgrEnabled))
    {
        $items = @()
        $Window.Source.Items | foreach {
            $items += $_
            }
        $items | foreach {
            [void]$Window.Source.Items.Remove($_)
        }
        $obj = New-Object -TypeName psobject
        Add-Member -InputObject $obj -MemberType NoteProperty -Name Info -Value 'Select a category'
        $Window.dataGrid.ItemsSource = [array]$obj
        $Window.dataGrid.Foreground = 'Black'
        $Window.listView.IsEnabled = $True
        $Window.Source.IsEnabled = $True
        'Client WMI', 'Client WMI (Full)' | ForEach-Object {
        [void]$Window.Source.Items.Add("$_")
        }
        $Window.Source.SelectedIndex = 0
    }

    If (!$Window.Online -and $Window.ConfigMgrEnabled -eq 'True' -and $ResourceID)
    {
        $items = @()
        $Window.Source.Items | foreach {
            $items += $_
            }
        $items | foreach {
            [void]$Window.Source.Items.Remove($_)
        }
        $obj = New-Object -TypeName psobject
        Add-Member -InputObject $obj -MemberType NoteProperty -Name Info -Value 'The target system is currently offline or inaccessible.  You can view data from the ConfigMgr database only. Select a category'
        $Window.dataGrid.ItemsSource = [array]$obj
        $Window.dataGrid.Foreground = 'Black'
        $Window.listView.IsEnabled = $True
        $Window.Source.IsEnabled = $True
        'ConfigMgr DB' | ForEach-Object {
        [void]$Window.Source.Items.Add("$_")
        }
        $Window.Source.SelectedIndex = 0
        $Window.ClassDescription.Text = "Using Configuration Manager Database"
        $Window.ClassDescription.IsEnabled = $false
    }

    If (!$Window.Online -and $Window.ConfigMgrEnabled -eq 'True' -and !$ResourceID)
    {
        $obj = New-Object -TypeName psobject
        Add-Member -InputObject $obj -MemberType NoteProperty -Name Warning -Value 'The target system is currently offline or inaccessible and the system cannot be found in the ConfigMgr database.'
        $Window.dataGrid.ItemsSource = [array]$obj
        $Window.dataGrid.Foreground = 'Red'
        $Window.listView.IsEnabled = $False
        $Window.Source.IsEnabled = $False
        $Window.ClassDescription.Text = "Using Configuration Manager Database"
        $Window.ClassDescription.IsEnabled = $false
    }

    If (!$Window.Online -and ($Window.ConfigMgrEnabled -eq 'False' -or !$Window.ConfigMgrEnabled))
    {
        $obj = New-Object -TypeName psobject
        Add-Member -InputObject $obj -MemberType NoteProperty -Name Warning -Value 'The target system is offline and ConfigMgr database use is not enabled.  No data can be viewed for this system.'
        $Window.dataGrid.ItemsSource = [array]$obj
        $Window.dataGrid.Foreground = 'Orange'
        $Window.listView.IsEnabled = $False
        $Window.Source.IsEnabled = $False
        $Window.ClassDescription.Text = ""
        $Window.ClassDescription.IsEnabled = $false
    }
}

# Function that defines what happens when an item is selected from the listview
function ListViewItem-Selected 
{
    param($Runspaces,$Window,$DontUpdate,$Code,$Hardware,$Instances,$ComputerName)
    
    # Dispose existing runspaces
    $Runspaces.GetEnumerator() | where {$_.Name -ne "Number"} | foreach {
    $_.Value.PowerShell.Stop()
    $_.Value.PowerShell.Runspace.Dispose()
    }
    $Runspaces.Clear()

    # Set some variables
    $ListItem = $This.SelectedItem.Content
    $Source = $Window.Source.SelectedItem

    # Clear some variables
    $Window.ClassDescription.Text = ''
    $Window.Enum.Text = ''
    if ($Hardware.Table)
    {
        $Hardware.Remove('Table')
    }

    # Notify user that we are loading data
    $obj = New-Object -TypeName psobject
    Add-Member -InputObject $obj -MemberType NoteProperty -Name Info -Value 'Loading...'
    $Window.dataGrid.ItemsSource = [array]$obj

    # Trigger the main code in a separate runspace
    # Prevent the code from running if the source has been changed, which also triggers a change of the listitem
    if ($DontUpdate)
    {

    }
    Else
    {
        # Set blank credential if not populated
        if (!($Window.Credentials))
        {
            $Window.Credentials = 'none'
        }
        $Credentials = $Window.Credentials

        # Run the code that gets and formats the data.  This is 'offloaded' to a seperate runspace to keep the GUI responsive.
        $Functions = @(
            'Function:\Start-RSJob'
            'Function:\Prepare-CustomArray'
            'Function:\ConvertTo-DataTable'
            'Function:\Switch-ClassNames'
            'Function:\Load-Data'
            'Function:\Populate-HardwareHashTable'
            'Function:\Reformat-Values'
            'Function:\Display-DataInGUI'
        )
        
        $Arguments = @(
            $ListItem
            $Hardware
            $Instances
            $Window
            $ComputerName
            $Credentials
            $Source   
        )

        $Params = @{
            Code      = $Code
            Arguments = $Arguments
            Runspaces = $Runspaces
            Functions = $Functions
        }

        Start-RSJob @Params
    }
}

# Function that defines what happens when an item is selected from the listview
function ListViewItemSQL-Selected 
{
    param($Runspaces,$Window,$DontUpdate,$SQLCode,$Hardware,$ComputerName)
    
    # Dispose existing runspaces
    $Runspaces.GetEnumerator() | where {$_.Name -ne "Number"} | foreach {
    $_.Value.PowerShell.Stop()
    $_.Value.PowerShell.Runspace.Dispose()
    }
    $Runspaces.Clear()

    # Set some variables
    $ListItem = $Window.ListView.SelectedItem.Content

    # Clear some variables
    $Window.ClassDescription.Text = 'Using Configuration Manager Database'
    $Window.ClassDescription.IsEnabled = $false
    $Window.Enum.Text = ''
    if ($Hardware.Table)
    {
        $Hardware.Remove('Table')
    }

    # Notify user that we are loading data
    $obj = New-Object -TypeName psobject
    Add-Member -InputObject $obj -MemberType NoteProperty -Name Info -Value 'Loading...'
    $Window.dataGrid.ItemsSource = [array]$obj

    # Trigger the main code in a separate runspace
    # Prevent the code from running if the source has been changed, which also triggers a change of the listitem
    if ($DontUpdate)
    {

    }
    Else
    {
        $Functions = @(
            'Function:\Start-RSJob'
            'Function:\Load-SQLData'
        )
        
        $Arguments = @(
            $ListItem
            $Hardware
            $Window
            $ComputerName
            $ResourceID
        )

        $Params = @{
            Code      = $SQLCode
            Arguments = $Arguments
            Runspaces = $Runspaces
            Functions = $Functions
        }

        Start-RSJob @Params
    }
}

# Function to update the items in the listView based on WMI Win32_ Classes in the cimv2 namespace
function Update-ListViewItems 
{
    # Source: Client WMI (Full)
    if ($Window.Source.SelectedItem -eq 'Client WMI (Full)')
    {
        # Define excluded classes (take too long to enumerate or do not report useful data)
        $Exclusions = @(
            'Win32_Account'
            'Win32_AccountSID'
            'Win32_ActionCheck'
            'Win32_CheckCheck'
            'Win32_Directory'
            'Win32_DirectorySpecification'
            'Win32_FileSpecification'
            'Win32_Group'
            'Win32_GroupInDomain'
            'Win32_GroupUser'
            'Win32_InstalledSoftwareElement'
            'Win32_MSIResource'
            'Win32_NTLogEvent'
            'Win32_NTLogEventComputer'
            'Win32_NTLogEventUser'
            'Win32_ODBCAttribute'
            'Win32_Perf'
            'Win32_PnPSignedDriverCIMDataFile'
            'Win32_Product'
            'Win32_ProductResource'
            'Win32_PublishComponentAction'
            'Win32_RegistryAction'
            'Win32_SelfRegModuleAction'
            'Win32_SettingCheck'
            'Win32_ShortcutFile'
            'Win32_SoftwareElement'
            'Win32_SoftwareElementAction'
            'Win32_SoftwareElementCheck'
            'Win32_SoftwareElementCondition'
            'Win32_SoftwareFeature'
            'Win32_SoftwareFeatureAction'
            'Win32_SoftwareFeatureCheck'
            'Win32_SoftwareFeatureParent'
            'Win32_SoftwareFeatureSoftwareElements'
            'Win32_SubDirectory'
            'Win32_UserAccount'
            'Win32_UserDesktop'
            'Win32_UserInDomain'
            'Win32_AllocatedResource'
            'Win32_ApplicationCommandLine'
            'Win32_AssociatedProcessorMemory'
            'Win32_CIMLogicalDeviceCIMDataFile'
            'Win32_ClassicCOMApplicationClasses'
            'Win32_ClassicCOMClassSettings'
            'Win32_ClientApplicationSetting'
            'Win32_ClusterShare'
            'Win32_CollectionStatistics'
            'Win32_COMApplicationClasses'
            'Win32_COMApplicationSettings'
            'Win32_COMClassAutoEmulator'
            'Win32_COMClassEmulator'
            'Win32_ControllerHasHub'
            'Win32_DCOMApplicationAccessAllowedSetting'
            'Win32_DependentService'
            'Win32_DeviceBus'
            'Win32_DeviceSettings'
            'Win32_DfsNode'
            'Win32_DfsNodeTarget'
            'Win32_DfsTarget'
            'Win32_DiskDrivePhysicalMedia'
            'Win32_DiskDriveToDiskPartition'
            'Win32_DriverForDevice'
            'Win32_IDEControllerDevice'
            'Win32_ImplementedCategory'
            'Win32_InstalledProgramFramework'
            'Win32_LoadOrderGroupServiceDependencies'
            'Win32_LoadOrderGroupServiceMembers'
            'Win32_LoggedOnUser'
            'Win32_LogicalDiskRootDirectory'
            'Win32_LogicalDiskToPartition'
            'Win32_LogicalFileAccess'
            'Win32_LogicalFileAuditing'
            'Win32_LogicalFileGroup'
            'Win32_LogicalFileOwner'
            'Win32_LogicalFileSecuritySetting'
            'Win32_LogicalProgramGroupDirectory'
            'Win32_LogicalProgramGroupItemDataFile'
            'Win32_LogonSessionMappedDisk'
            'Win32_ManagedSystemElementResource'
            'Win32_MemoryArrayLocation'
            'Win32_MemoryDeviceArray'
            'Win32_MemoryDeviceLocation'
            'Win32_MountPoint'
            'Win32_NetworkAdapterSetting'
            'Win32_NTLogEventLog'
            'Win32_OfflineFilesMachineConfiguration'
            'Win32_OperatingSystemAutochkSetting'
            'Win32_OperatingSystemQFE'
            'Win32_PhysicalMemoryLocation'
            'Win32_PingStatus'
            'Win32_PNPAllocatedResource'
            'Win32_PNPDevice'
            'Win32_PatchFile'
            'Win32_PrinterDriverDll'
            'Win32_PrinterSetting'
            'Win32_ProductCheck'
            'Win32_ProductSoftwareFeatures'
            'Win32_ProgramGroupContents'
            'Win32_ProtocolBinding'
            'Win32_Property'
            'Win32_RoamingProfileMachineConfiguration'
            'Win32_SecuritySettingOfLogicalFile'
            'Win32_SecuritySettingOfObject'
            'Win32_SerialPortSetting'
            'Win32_SessionProcess'
            'Win32_SessionResource'
            'Win32_ShareToDirectory'
            'Win32_ServiceSpecificationService'
            'Win32_SID'
            'Win32_SIDandAttributes'
            'Win32_SystemBIOS'
            'Win32_SystemBootConfiguration'
            'Win32_SystemDesktop'
            'Win32_SystemDevices'
            'Win32_ShortcutSAP'
            'Win32_SystemDriverPNPEntity'
            'Win32_SystemLoadOrderGroups'
            'Win32_SoftwareElementResource'
            'Win32_SystemNetworkConnections'
            'Win32_SystemOperatingSystem'
            'Win32_SystemPartitions'
            'Win32_SystemProcesses'
            'Win32_SystemProgramGroups'
            'Win32_SystemResources'
            'Win32_SystemServices'
            'Win32_SystemSetting'
            'Win32_SystemSystemDriver'
            'Win32_SystemTimeZone'
            'Win32_SystemUsers'
            'Win32_USBControllerDevice'
            'Win32_UserProfile'
            'Win32_VolumeQuota'
            'Win32_WMIElementSetting'
            'Win32_ApplicationCommandLine'
            'Win32_ApplicationService'
            'Win32_Binary'
            'Win32_BindImageAction'
            'Win32_ClassInfoAction'
            'Win32_CommandLineAccess'
            'Win32_Condition'
            'Win32_CreateFolderAction'
            'Win32_DuplicateFileAction'
            'Win32_EnvironmentSpecification'
            'Win32_ExtensionInfoAction'
            'Win32_FontInfoAction'
            'Win32_IniFileSpecification'
            'Win32_LaunchCondition'
            'Win32_ManagedSystemElementResource'
            'Win32_MIMEInfoAction'
            'Win32_MoveFileAction'
            'Win32_NTDomain'
            'Win32_NTLogEventLog'
            'Win32_ODBCDataSourceAttribute'
            'Win32_ODBCDataSourceSpecification'
            'Win32_ODBCDriverAttribute'
            'Win32_ODBCDriverSoftwareElement'
            'Win32_ODBCDriverSpecification'
            'Win32_ODBCSourceAttribute'
            'Win32_ODBCTranslatorSpecification'
            'Win32_Patch'
            'Win32_PatchFile'
            'Win32_PatchPackage'
            'Win32_ProductCheck'
            'Win32_ProgIDSpecification'
            'Win32_Property'
            'Win32_RemoveFileAction'
            'Win32_RemoveIniAction'
            'Win32_ReserveCost'
            'Win32_ServiceControl'
            'Win32_ServiceSpecification'
            'Win32_ServiceSpecificationService'
            'Win32_ShortcutAction'
            'Win32_ShortcutSAP'
            'Win32_SoftwareElementResource'
            'Win32_TypeLibraryAction'
            'Win32_ClassicCOMClass'
            'Win32_ClassicCOMClassSetting'
            'Win32_ClassInfoAction'
            'Win32_COMClass'
            'Win32_ComSetting'
            'Win32_LogicalProgramGroup'
            'Win32_Reliability'
            'Win32_ReliabilityRecords'
        )

        # Create an exclusion array for "Perf" classes
        $perf = Get-CimClass -Namespace ROOT\cimv2 |
        Where-Object -FilterScript {
            $_.CimClassName -match 'Win32_' -and ($_.CimClassName -match 'PerfFormattedData' -or $_.CimClassName -match 'PerfRawData')
        } |
        Sort-Object -Property CimClassName |
        Select-Object -ExpandProperty CimClassName
        
        # Create an exclusion array for "PNP" classes
        $pnp = Get-CimClass -Namespace ROOT\cimv2 |
        Where-Object -FilterScript {
            $_.CimClassName -match 'Win32_PNPDeviceProperty'
        } |
        Sort-Object -Property CimClassName |
        Select-Object -ExpandProperty CimClassName
        
        # Create an exclusion array for "NamedJob" classes
        $nj = Get-CimClass -Namespace ROOT\cimv2 |
        Where-Object -FilterScript {
            $_.CimClassName -match 'Win32_NamedJobObject'
        } |
        Sort-Object -Property CimClassName |
        Select-Object -ExpandProperty CimClassName
        
        # Get the Names of the Win32_ Cim Classes in cimv2 namespace and add to an array
        $CimInstanceNames = @()
        $CimNames = Get-CimClass -Namespace ROOT\cimv2 |
        Where-Object -FilterScript {
            $_.CimClassName -match 'Win32_' -and $_.CimClassName -notin $perf -and $_.CimClassName -notin $pnp -and $_.CimClassName -notin $nj -and $_.CimClassName -notin $Exclusions
        } |
        Sort-Object -Property CimClassName |
        Select-Object -ExpandProperty CimClassName |
        ForEach-Object -Process {
            $CimInstanceNames += $_.Split('_')[1]
        }

        # Add the class names to the listview
        $Window.ListItems = $CimInstanceNames
        $Window.ListItems | ForEach-Object -Process {
            $Item = New-Object -TypeName Windows.Controls.ListViewItem #[windows.controls.listviewitem]::new()
            $Item.Content = $_
            $Window.listView.AddChild($Item)
        }
    }

    # Source: Client WMI (Simple)
    if ($Window.Source.SelectedItem -eq 'Client WMI')
    {
        # Define List of Hardware Categories (corresponding to WMI classes)
        $Window.ListItems = @(
            'Operating System'
            'BIOS'
            'Computer System'
            'Disk Drive'
            'Disk Partition'
            'Keyboard'
            'Logical Disk'
            'Network Adapter'
            'Network Adapter Configuration'
            'Physical Memory'
            'Processor'
            'Sound Device'
            'Timezone'
            'USB Device'
            'Video Controller'
            'Volume'
            'User Profiles'
            'Shares'
            'Printers'
            'Installed Windows Features'
            'Network Login Profiles'
            'Mapped Drives'
        )

        # Populate ListView Items with Hardware categories
        $Window.ListItems | ForEach-Object -Process {
            $Item = New-Object -TypeName Windows.Controls.ListViewItem #[windows.controls.listviewitem]::new()
            $Item.Content = $_
            $Window.listView.AddChild($Item)
        }
    }

    # Source: ConfigMgrDB (Simple)
    if ($Window.Source.SelectedItem -eq 'ConfigMgr DB')
    {
        Get-ResourceID
        
        $Views = @()

        [string]$Query = "Select Name from sys.Views 
            where Name like 'v_GS%'
        Order by Name"
        [int]$ConnectionTimeout = 15
        [int]$CommandTimeout = 600
    
        # Define connection string
        $connectionString = "Server=$($Window.SQLServer);Database=$($Window.Database);Integrated Security=SSPI;Connection Timeout=$ConnectionTimeout"
    
        # Open the connection
        $connection = New-Object -TypeName System.Data.SqlClient.SqlConnection
        $connection.ConnectionString = $connectionString
        Try
        {
            $connection.Open()
        }
        Catch
        {
            $MyError = $_
            $obj = New-Object -TypeName psobject
            Add-Member -InputObject $obj -MemberType NoteProperty -Name Error -Value "Could not connect to $($Window.Database) on $($Window.SQLServer). $_"
            $Window.GUI.Dispatcher.Invoke(
                [action]{
                    $Window.dataGrid.ItemsSource = [array]$obj
                    $Window.dataGrid.Foreground = 'Red'
            })
            Return
        }
        
        # Execute the query 
        $command = $connection.CreateCommand()
        $command.CommandText = $Query
        $command.CommandTimeout = $CommandTimeout
        $reader = $command.ExecuteReader()
        
        # Load results to a data table
        $table = New-Object -TypeName 'System.Data.DataTable'
        $table.Load($reader)
        $SQLViews = $table | Select-Object -ExpandProperty Name
         
        # Close the connection
        $connection.Close()

        $SQLViews | ForEach-Object {
            $Views += $_.Substring(5)
        }

        # Populate ListView Items with Hardware categories
        $Window.ListItems = $Views
        $Window.ListItems | ForEach-Object -Process {
            $Item = New-Object -TypeName Windows.Controls.ListViewItem #[windows.controls.listviewitem]::new()
            $Item.Content = $_
            $Window.listView.AddChild($Item)
        }
    }

    # If system is online, enable the listview, advise user to select category
    if ($Window.Online)
    { 
        $obj = New-Object -TypeName psobject
        Add-Member -InputObject $obj -MemberType NoteProperty -Name Info -Value 'Select a category'
        $Window.dataGrid.ItemsSource = [array]$obj
        $Window.dataGrid.Foreground = 'Black'
        $Window.listView.IsEnabled = $True
        $Window.ClassDescription.Text = '[WMI Class Description]'
    }
}

# Function to display the WMI description in a separate window when clicked
function Open-WMIDescriptionWindow 
{
    # Create a Window
    $NewWindow = New-Object Windows.Window
    $NewWindow.SizeToContent = 'WidthAndHeight'
    $NewWindow.WindowStyle = 'none'
    $NewWindow.WindowStartupLocation = 'CenterScreen'
    $NewWindow.Background = 'White'

        
    # Add a label control
    $NewTextBox = New-Object Windows.Controls.TextBox
    $NewTextBox.Height = 300
    $NewTextBox.Width = 500
    $NewTextBox.Text = $Hardware.Description
    $NewTextBox.FontSize = 16
    $NewTextBox.FontWeight = 'Normal'
    $NewTextBox.Foreground = 'Blue'
    $NewTextBox.VerticalContentAlignment = 'Top'
    $NewTextBox.HorizontalContentAlignment = 'Left'
    $NewTextBox.TextWrapping = 'Wrap'
    $NewTextBox.VerticalScrollBarVisibility = 'Auto'
    $NewTextBox.IsReadOnly = $True
    $NewTextBox.BorderThickness = 0
    $NewTextBox.Padding = 5

    # Add an event to close the window when clicked
    $NewWindow.Add_PreviewMouseDown({
            $This.Close()
    })

    # Add an event to change the cursor to a hand when the mouse goes over the window
    $NewTextBox.Add_MouseEnter({
            $This.Cursor = 'Hand'
    })

    # Or display the window until clicked
    $NewWindow.Content = $NewTextBox
    [void]$NewWindow.ShowDialog()
}


# Function to display the WMI description in a separate window when clicked
function Open-AboutWindow 
{
    # Create a Window
    $NewWindow = New-Object Windows.Window
    $NewWindow.SizeToContent = 'WidthAndHeight'
    #$NewWindow.AllowsTransparency = $true
    $NewWindow.WindowStyle = 'none'
    $NewWindow.WindowStartupLocation = 'CenterScreen'
    $NewWindow.ResizeMode = 'NoResize'

        
    # Add a label control
    $Text = "SYSTEM EXPLORER FOR WINDOWS allows deep exploration of system and hardware information for a local or remote computer as exposed through WMI Win32 classes (if the system is online), as well as from the System Center Configuration Manager Database (whether the system is online or offline).`nIt is a WPF application coded in PowerShell.`
        `nVersion: 1.0 `
        `nAuthor: Trevor Jones `
    `nWebsite: http://smsagent.wordpress.com"
    $NewTextBox = New-Object Windows.Controls.TextBox
    $NewTextBox.Height = 300
    $NewTextBox.Width = 400
    $NewTextBox.Text = $Text
    $NewTextBox.FontSize = 16
    $NewTextBox.FontWeight = 'Normal'
    $NewTextBox.Foreground = 'Blue'
    $NewTextBox.VerticalContentAlignment = 'Top'
    $NewTextBox.HorizontalContentAlignment = 'Left'
    $NewTextBox.TextWrapping = 'Wrap'
    $NewTextBox.VerticalScrollBarVisibility = 'Auto'
    $NewTextBox.IsReadOnly = $True
    $NewTextBox.BorderThickness = 0
    $NewTextBox.Padding = 5
    $NewTextBox.Background = 'AliceBlue'

    # Add an event to close the window when clicked
    $NewWindow.Add_PreviewMouseDown({
            $This.Close()
    })

    # Add an event to change the cursor to a hand when the mouse goes over the window
    $NewTextBox.Add_MouseEnter({
            $This.Cursor = 'Hand'
    })

    # Or display the window until clicked
    $NewWindow.Content = $NewTextBox
    [void]$NewWindow.ShowDialog()
}

# Function to export the current datagrid values to csv
function Export-ToCSV 
{
    if ($Hardware.Table)
    {
        $file = "$env:Temp\$(Get-Random).csv"
        $CSVFiles.Add($file)
        $Hardware.Table | Export-Csv -Path $file -NoTypeInformation
        Invoke-Item $file
    }
    Else
    {
        Open-NoCSVWindow
    }
}

# Function to display the WMI description in a separate window when clicked
function Open-NoCSVWindow 
{
    # Create a Window
    $NewWindow = New-Object Windows.Window
    $NewWindow.SizeToContent = 'WidthAndHeight'
    $NewWindow.WindowStyle = 'none'
    $NewWindow.WindowStartupLocation = 'CenterScreen'
    $NewWindow.ResizeMode = 'NoResize'

        
    # Add a label control
    $Text = 'There is no data to output!'
    $NewTextBox = New-Object Windows.Controls.TextBox
    $NewTextBox.Height = 100
    $NewTextBox.Width = 300
    $NewTextBox.Text = $Text
    $NewTextBox.FontSize = 20
    $NewTextBox.FontWeight = 'Bold'
    $NewTextBox.Foreground = 'Red'
    $NewTextBox.VerticalContentAlignment = 'Center'
    $NewTextBox.HorizontalContentAlignment = 'Center'
    $NewTextBox.TextWrapping = 'Wrap'
    $NewTextBox.IsReadOnly = $True
    $NewTextBox.BorderThickness = 0
    $NewTextBox.Padding = 5
    $NewTextBox.Background = 'White'

    # Add an event to close the window when clicked
    $NewWindow.Add_PreviewMouseDown({
            $This.Close()
    })

    # Add an event to change the cursor to a hand when the mouse goes over the window
    $NewTextBox.Add_MouseEnter({
            $This.Cursor = 'Hand'
    })

    # Or display the window until clicked
    $NewWindow.Content = $NewTextBox
    [void]$NewWindow.ShowDialog()
}

# Function to get the resourceID of the computer
function Get-ResourceID 
{
    if ($Window.ConfigMgrEnabled -eq 'True')
    {
        [string]$Query = "Select ResourceID from v_R_System where Name0='$ComputerName'"
        [int]$ConnectionTimeout = 15
        [int]$CommandTimeout = 600

        # Define connection string
        $connectionString = "Server=$($Window.SQLServer);Database=$($Window.Database);Integrated Security=SSPI;Connection Timeout=$ConnectionTimeout"

        # Open the connection
        $connection = New-Object -TypeName System.Data.SqlClient.SqlConnection
        $connection.ConnectionString = $connectionString
        try
        {
            $connection.Open()
        }
        Catch
        {
            $MyError = $_
            $obj = New-Object -TypeName psobject
            Add-Member -InputObject $obj -MemberType NoteProperty -Name Error -Value "Could not connect to $($Window.Database) on $($Window.SQLServer). $_"
            $Window.GUI.Dispatcher.Invoke(
                [action]{
                    $Window.dataGrid.ItemsSource = [array]$obj
                    $Window.dataGrid.Foreground = 'Red'
            })
            Return
        }

    
        # Execute the query 
        $command = $connection.CreateCommand()
        $command.CommandText = $Query
        $command.CommandTimeout = $CommandTimeout
        $reader = $command.ExecuteReader()
    
        # Load results to a data table
        $table = New-Object -TypeName 'System.Data.DataTable'
        $table.Load($reader)
        $Global:ResourceID = $table | Select-Object -ExpandProperty ResourceID
     
        # Close the connection
        $connection.Close()
    }
}

# Function to load the SQL data
function Load-SQLData 
{
    param($Window,$ListItem,$Hardware,$ComputerName,$ResourceID)
    If (!$ResourceID)
    {
        $obj = New-Object -TypeName psobject
        Add-Member -InputObject $obj -MemberType NoteProperty -Name Error -Value "No Resource ID could be found in the ConfigMgr Database for $ComputerName"
        $Window.GUI.Dispatcher.Invoke(
            [action]{
                $Window.dataGrid.ItemsSource = [array]$obj
        })
        Continue
    }

    [string]$Query = "Select * from v_GS_$ListItem where ResourceID='$ResourceID'"
    [int]$ConnectionTimeout = 30
    [int]$CommandTimeout = 600

    # Define connection string
    $connectionString = "Server=$($Window.SQLServer);Database=$($Window.Database);Integrated Security=SSPI;Connection Timeout=$ConnectionTimeout"

    # Open the connection
    $connection = New-Object -TypeName System.Data.SqlClient.SqlConnection
    $connection.ConnectionString = $connectionString
    $connection.Open()
    
    # Execute the query 
    $command = $connection.CreateCommand()
    $command.CommandText = $Query
    $command.CommandTimeout = $CommandTimeout
    $reader = $command.ExecuteReader()
    
    # Load results to a data table
    $table = New-Object -TypeName 'System.Data.DataTable'
    $table.Load($reader)
    $Hardware.Table = $table
     
    # Close the connection
    $connection.Close()

    if ($table.Rows.Count -eq 0)
    {
        $obj = New-Object -TypeName psobject
        Add-Member -InputObject $obj -MemberType NoteProperty -Name Info -Value 'No results.'
        $Window.GUI.Dispatcher.Invoke(
            [action]{
                $Window.dataGrid.ItemsSource = [array]$obj
        })
    }
    Else
    {
        $Window.GUI.Dispatcher.Invoke(
            [action]{
                $Window.dataGrid.ItemsSource = $table.DefaultView
        })
    }
}

# Function to load Settings Window
function Show-SettingsWindow 
{
    function Read-Registry 
    {
        if (Test-Path -Path 'HKLM:\SOFTWARE\Wow6432Node\SMSAgent\System Explorer')
        {
            $regsql = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\SMSAgent\System Explorer' -Name SQLServer | Select-Object -ExpandProperty SQLServer
            $regdb = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\SMSAgent\System Explorer' -Name Database | Select-Object -ExpandProperty Database
            $enabled = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\SMSAgent\System Explorer' -Name ConfigMgrEnabled | Select-Object -ExpandProperty ConfigMgrEnabled
        }

        if (Test-Path -Path 'HKLM:\SOFTWARE\SMSAgent\System Explorer')
        {
            $regsql = Get-ItemProperty -Path 'HKLM:\SOFTWARE\SMSAgent\System Explorer' -Name SQLServer | Select-Object -ExpandProperty SQLServer
            $regdb = Get-ItemProperty -Path 'HKLM:\SOFTWARE\SMSAgent\System Explorer' -Name Database | Select-Object -ExpandProperty Database
            $enabled = Get-ItemProperty -Path 'HKLM:\SOFTWARE\SMSAgent\System Explorer' -Name ConfigMgrEnabled | Select-Object -ExpandProperty ConfigMgrEnabled
        }

        if ($regsql -ne $null -and $regsql -ne '')
        {
            $GUI.SQLServer.Text = $regsql
            $Window.SQLServer = $regsql
        }
        if ($regdb -ne $null -and $regdb -ne '')
        {
            $GUI.Database.Text = $regdb
            $Window.Database = $regdb
        }
        if ($enabled -ne $null -and $enabled -ne '' -and $enabled -ne 'False')
        {
            $GUI.checkBox.IsChecked = $True
        }
        if ($enabled -ne $null -and $enabled -ne '' -and $enabled -ne 'True')
        {
            $GUI.checkBox.IsChecked = $false
        }
        if ($GUI.checkBox.IsChecked -eq $false)
        {
            $GUI.SQLServer.IsEnabled = $false  
            $GUI.Database.IsEnabled = $false 
            $GUI.SQLLabel.IsEnabled = $false 
            $GUI.DatabaseLabel.IsEnabled = $false 
            $GUI.Button_Save.IsEnabled = $false
        }
    }


    function Write-Registry 
    {
        $SQLServer = $GUI.SQLServer.Text
        $Window.SQLServer = $GUI.SQLServer.Text
        $Database = $GUI.Database.Text
        $Window.Database = $GUI.Database.Text
        $ConfigMgrEnabled = $GUI.checkBox.IsChecked
        $Window.ConfigMgrEnabled = $GUI.checkBox.IsChecked

        try 
        {
            if (Test-Path -Path 'HKLM:\SOFTWARE\Wow6432Node\SMSAgent\System Explorer')
            {
                Start-Process -FilePath powershell.exe -ArgumentList "-Command ""Set-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\SMSAgent\System Explorer' -Name SQLServer -Value $SQLServer; Set-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\SMSAgent\System Explorer' -Name Database -Value $Database; Set-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\SMSAgent\System Explorer' -Name ConfigMgrEnabled -Value $ConfigMgrEnabled""" -Verb runas -WindowStyle Hidden -Wait -ErrorAction Stop
            }

            if (Test-Path -Path 'HKLM:\SOFTWARE\SMSAgent\System Explorer')
            {
                Start-Process -FilePath powershell.exe -ArgumentList "-Command ""Set-ItemProperty -Path 'HKLM:\SOFTWARE\SMSAgent\System Explorer' -Name SQLServer -Value $SQLServer; Set-ItemProperty -Path 'HKLM:\SOFTWARE\SMSAgent\System Explorer' -Name Database -Value $Database; Set-ItemProperty -Path 'HKLM:\SOFTWARE\SMSAgent\System Explorer' -Name ConfigMgrEnabled -Value $ConfigMgrEnabled""" -Verb runas -WindowStyle Hidden -Wait -ErrorAction Stop
            }

            $GUI.Saved.Visibility = 'Visible'
        }
        Catch 
        {
            $GUI.Saved.Content = 'Failed!'
            $GUI.Saved.Foreground = 'Red'
            $GUI.Saved.Visibility = 'Visible'
        }
    }


    # Define XAML
    [XML]$XAML = @'
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="MainWindow" Height="176.773" Width="300" WindowStartupLocation="CenterScreen" AllowsTransparency="True" WindowStyle="None" Background="AliceBlue">
    <Grid>
        <Button x:Name="Button_Save" Content="Save" Background="Aquamarine" BorderThickness="0" HorizontalAlignment="Left" Margin="129,132,0,0" VerticalAlignment="Top" Width="79" Height="35"/>
        <Button x:Name="Button_Close" Content="Close" Background="Aquamarine" BorderThickness="0" HorizontalAlignment="Left" Margin="213,132,0,0" VerticalAlignment="Top" Width="79" Height="35"/>
        <Label x:Name="label" Content="Use ConfigMgr Database?" FontSize="20" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top"/>
        <CheckBox x:Name="checkBox" HorizontalAlignment="Left" Margin="254,14,0,0" VerticalAlignment="Top">
            <CheckBox.LayoutTransform>
                <ScaleTransform ScaleX="2" ScaleY="2" />
            </CheckBox.LayoutTransform>
        </CheckBox>
        <Label x:Name="SQLLabel" Content="SQL Server" BorderThickness="0" FontSize="20" HorizontalAlignment="Left" Margin="10,47,0,0" VerticalAlignment="Top"/>
        <Label x:Name="DatabaseLabel" Content="Database" BorderThickness="0" FontSize="20" HorizontalAlignment="Left" Margin="10,84,0,0" VerticalAlignment="Top"/>
        <TextBox x:Name="SQLServer" CharacterCasing="Upper" FontSize="13" BorderThickness="0" VerticalContentAlignment="Center" HorizontalAlignment="Left" Height="32" Margin="121,52,0,0" TextWrapping="Wrap" Text="&lt;SQLServer&gt;\&lt;Instance&gt;" VerticalAlignment="Top" Width="169"/>
        <TextBox x:Name="Database" CharacterCasing="Upper" FontSize="13" BorderThickness="0" VerticalContentAlignment="Center" HorizontalAlignment="Left" Height="32" Margin="121,89,0,0" TextWrapping="Wrap" Text="CM_XXX" VerticalAlignment="Top" Width="169"/>
        <Label x:Name="Saved" Content="Saved!" Visibility="Hidden" BorderThickness="0" FontSize="20" Foreground="Aquamarine" FontWeight="Bold" HorizontalAlignment="Left" Margin="30,130,0,0" VerticalAlignment="Top"/>
    </Grid>
</Window>
'@

    # Load the XAML, add it to the hash table, and also add the GUI elements
    $reader = (New-Object -TypeName System.XML.XMLNodeReader -ArgumentList $XAML)
    $script:GUI = @{}
    $GUI.Window = [Windows.Markup.XAMLReader]::Load($reader)
    $XAML.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object -Process {
        $GUI.$($_.Name) = $GUI.Window.FindName($_.Name)
    }

    # Add an event on the Close button
    $GUI.Button_Close.Add_Click({
            $GUI.Window.Close()
    })

    # Add an event on the Save button
    $GUI.Button_Save.Add_Click({
            Write-Registry
    })

    # Add an event when clicking into the SQL Server text box
    $GUI.SQLServer.Add_PreviewMouseDown({
            $This.Text = ' '
    })

    # Add an event when clicking into the SQL Server text box
    $GUI.SQLServer.Add_GotKeyboardFocus({
            $This.Text = ' '
    })

    # Add an event when clicking into the Database text box
    $GUI.Database.Add_PreviewMouseDown({
            $This.Text = ' '
    })

    # Add an event when clicking into the Database text box
    $GUI.Database.Add_GotKeyboardFocus({
            $This.Text = ' '
    })

    # Add an event when the window loads
    $GUI.Window.Add_Loaded({
            Read-Registry
    })

    # Add an event when the checkbox is checked
    $GUI.checkBox.Add_Checked({
            $GUI.SQLServer.IsEnabled = $True  
            $GUI.Database.IsEnabled = $True 
            $GUI.SQLLabel.IsEnabled = $True 
            $GUI.DatabaseLabel.IsEnabled = $True 
            $GUI.Button_Save.IsEnabled = $True
    })

    # Add an event when the checkbox is unchecked
    $GUI.checkBox.Add_Unchecked({
            $GUI.SQLServer.IsEnabled = $false  
            $GUI.Database.IsEnabled = $false 
            $GUI.SQLLabel.IsEnabled = $false 
            $GUI.DatabaseLabel.IsEnabled = $false 
    })

    $GUI.Window.Add_Closed({
            If ($GUI.checkBox.IsChecked -eq $True)
            {
                if ($Window.Source.Items -contains 'ConfigMgr DB')
                {

                }
                Else
                {
                    If ($Window.Source.Items.Count -ge 1)
                    {
                        $Window.source.Items.Add('ConfigMgr DB')
                        $Window.Source.SelectedIndex = 0
                    }
                }
            }
            If ($GUI.checkBox.IsChecked -eq $false)
            {
                If ($Window.Source.Items.Count -ge 1)
                {
                    $Window.source.Items.Remove('ConfigMgr DB')
                    $Window.Source.SelectedIndex = 0
                }
            }
    })

    # Show the window

    [void]$GUI.window.ShowDialog()
}

#endregion




#region PrepareGUI

# Define XAML
Add-Type -AssemblyName PresentationFramework
[XML]$XAML = @'
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="System Explorer for Windows" Height="507.875" Width="865.186" WindowStartupLocation="CenterScreen">
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="310"/>
            <ColumnDefinition Width="5"/>
            <ColumnDefinition Width="343*"/>
            <ColumnDefinition Width="199"/>
        </Grid.ColumnDefinitions>
        <GridSplitter Grid.Column="1" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Width="auto" Height="auto" Background="white" ToolTip="Resize" />
        <DataGrid x:Name="dataGrid" Grid.Column="2" Background="White" FontSize="14" Height="auto" Width="auto" VerticalAlignment="Stretch" HorizontalAlignment="Stretch" GridLinesVisibility="None" IsReadOnly="True" CanUserSortColumns="True" Margin="1,72,0,0" Grid.ColumnSpan="2" />
        <ListView x:Name="listView" FontSize="16" Background="White" Padding="5" Foreground="Blue" BorderBrush="White" FontWeight="Bold" Margin="0,120,0,0"/>
        <Label x:Name="label" Content="Computer Name:" HorizontalAlignment="Left" VerticalAlignment="Top" Height="39" Width="141" FontWeight="Bold" FontSize="16" VerticalContentAlignment="Center" HorizontalContentAlignment="Left"/>
        <TextBox x:Name="ComputerName" Height="39" Margin="141,0,0,0" TextWrapping="Wrap" CharacterCasing="Upper" BorderBrush="White" Text="" VerticalAlignment="Top" FontWeight="Bold" FontSize="16" VerticalContentAlignment="Center" HorizontalContentAlignment="Center"/>
        <Button x:Name="button" Content="Connect" Margin="0,72,0,0" VerticalAlignment="Top" Height="48" BorderThickness="0" Background="AliceBlue" FontWeight="Bold" FontSize="20"/>
        <Label x:Name="label_Copy" Content="Source:" HorizontalAlignment="Left" VerticalAlignment="Top" Height="34" Width="141" FontWeight="Bold" FontSize="16" VerticalContentAlignment="Center" HorizontalContentAlignment="Left" Margin="0,38,0,0"/>
        <ComboBox x:Name="Source" Margin="141,39,0,0" VerticalAlignment="Top" Height="33" Background="White" BorderBrush="White" FontSize="16" VerticalContentAlignment="Center" HorizontalContentAlignment="Left"/>
        <TextBox x:Name="ClassDescription" Grid.Column="2" Padding="5" HorizontalAlignment="Stretch" Height="72" TextWrapping="Wrap" Text="[WMI Class Description]" VerticalAlignment="Top" MinWidth="343" BorderThickness="0" Foreground="Blue" IsReadOnly="True" FontSize="14" VerticalScrollBarVisibility="Auto"/>
        <Button x:Name="Settings" Content="Settings" Grid.Column="3" BorderThickness="0" Background="AliceBlue" HorizontalAlignment="Right" VerticalAlignment="Top" Width="100" MinWidth="100" Height="36" Margin="0,0,-1,0"/>
        <Button x:Name="About" Content="About" Grid.Column="3" BorderThickness="0" Background="AliceBlue" HorizontalAlignment="Right" VerticalAlignment="Top" Width="100" MinWidth="100" Height="36" Margin="0,36,-1,0"/>
        <Button x:Name="Export" Content="Export to CSV" Grid.Column="3" BorderThickness="0" Background="AliceBlue" HorizontalAlignment="Right" VerticalAlignment="Top" Width="100" MinWidth="100" Height="36" Margin="0,0,99,0"/>
        <Label x:Name="label1" Content="Enumeration Time (s)" FontSize="9" Grid.Column="3" VerticalContentAlignment="Top" HorizontalAlignment="Right" Margin="0,36,99,0" VerticalAlignment="Top" Height="21" Width="100"/>
        <TextBox x:Name="Enum" Grid.Column="3" FontSize="9" HorizontalContentAlignment="Center" BorderThickness="0" HorizontalAlignment="Right" Height="15" Margin="0,57,99,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="100" IsReadOnly="True"/>
    </Grid>
</Window>
'@

# Create a synchronized hashtable available between runspaces
$Global:Window = [hashtable]::Synchronized(@{})
$Runspaces = [hashtable]::Synchronized(@{}) #[System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
$Global:CSVFiles = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))

# Load the XAML, add it to the hash table, and also add the GUI elements
$reader = (New-Object -TypeName System.XML.XMLNodeReader -ArgumentList $XAML)
$Window.GUI = [Windows.Markup.XAMLReader]::Load($reader)
$XAML.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object -Process {
    $Window.$($_.Name) = $Window.GUI.FindName($_.Name)
}

# Add the PS host so that runspace can use it for credential prompting
$Window.Host = $host

# Set the default computername to the current system
$Window.ComputerName.Text = $env:COMPUTERNAME

# Load values from registry
Read-Registry

# Set the window size to 65% of screen
$Window.GUI.Height = ([System.Windows.SystemParameters]::PrimaryScreenHeight * 0.65)
$Window.GUI.Width = ([System.Windows.SystemParameters]::PrimaryScreenWidth * 0.65)

# Set the App icons
if (Test-Path -Path "$env:ProgramFiles\SMSAgent\System Explorer for Windows\Map.ico")
{
    $Window.GUI.Icon = "$env:ProgramFiles\SMSAgent\System Explorer for Windows\Map.ico"
}
if (Test-Path -Path "${env:ProgramFiles(x86)}\SMSAgent\System Explorer for Windows\Map.ico")
{
    $Window.GUI.Icon = "${env:ProgramFiles(x86)}\SMSAgent\System Explorer for Windows\Map.ico"
}




# Define List of Hardware Categories (corresponding to WMI classes)
$Window.ListItems = @(
    'Operating System'
    'BIOS'
    'Computer System'
    'Disk Drive'
    'Disk Partition'
    'Keyboard'
    'Logical Disk'
    'Network Adapter'
    'Network Adapter Configuration'
    'Physical Memory'
    'Processor'
    'Sound Device'
    'Timezone'
    'USB Device'
    'Video Controller'
    'Volume'
    'User Profiles'
    'Shares'
    'Printers'
    'Installed Windows Features'
    'Network Login Profiles'
    'Mapped Drives'
)

# Populate ListView Items with Hardware categories
$Window.ListItems | ForEach-Object {
    $Item = New-Object -TypeName Windows.Controls.ListViewItem #[windows.controls.listviewitem]::new()
    $Item.Content = $_
    $Window.listView.AddChild($Item)
}

# Disable the listView until data is loaded
$Window.Listview.IsEnabled = $false

# Disable the Source combo until data is loaded
$Window.Source.IsEnabled = $false
#endregion



#region NestedCode

# Define the WMI code to run in the runspace
$Code = {
    param($ListItem,$Hardware,$Instances,$Window,$ComputerName,$Credentials,$Source)

    # The code that will run in a seperate runspace.  This will add populated WMI class values to a hash table
    $NestedCode = {
        param($ListItem,$Hardware,$Instances,$Window,$ComputerName,$Credentials,$Source)
        
        $StartTime = Get-Date
        $ListItems = $Window.ListItems

        # Switch displayed item names with their corresponding WMI class
        Switch-ClassNames -ListItem $ListItem

        # Load the data for the selected class
        Load-Data -Credentials $Credentials -ListItem $ListItem -ComputerName $ComputerName -Class $Class -Window $Window

        # Populate the number of instances in this class, which will be used to determine whether to display vertically (single instance) or horizontally (multiple instances)
        if ($TempVar.Count)
        {
            $Instances.$ListItem = $TempVar.Count
        }
        Else 
        {
            $Instances.$ListItem = 0
        }

        # Add the list of populated properties for each class and instance to a hash table, appending a unique number to each instance 
        if ($TempVar)
        {
            Populate-HardwareHashTable -ListItem $ListItem -Hardware $Hardware -TempVar $TempVar
        }

        # Reformat certain values and dates and expand property arrays into strings
        Reformat-Values -Hardware $Hardware -ListItems $ListItems

        # Variable cleanup
        Remove-Variable -Name TempVar

        # Close Cim-session
        Remove-CimSession $Session

        # Get WMI Class Description
        $o = New-Object System.Management.ObjectGetOptions
        $o.UseAmendedQualifiers = $True
        #$c = [System.Management.ManagementClass]::new("Win32_$Class",$o)
        $c = New-Object System.Management.ManagementClass -ArgumentList "Win32_$Class",$o
        $Hardware.Description = $c.GetQualifierValue('Description')

        # Process the data into a datatable and display the results in the GUI
        $Functions = @(
            'Function:\Prepare-CustomArray'
            'Function:\ConvertTo-DataTable'
        )
        $Params = @{
            Source      = $Source
            Instances   = $Instances
            ListItem    = $ListItem
            Hardware    = $Hardware
            Window      = $Window
            Description = $Hardware.Description
        }

        Display-DataInGUI $Functions @Params 

        # Display enum time in window
        $EndTime = Get-Date
        $EnumTime = [math]::Round(($EndTime - $StartTime).TotalSeconds,2)
        $Window.GUI.Dispatcher.Invoke(
            [action]{
                $Window.Enum.Text = $EnumTime
        })
    }
    
    # Invoke the code in a seperate runspace
    
    $Functions = @(
        'Function:\Prepare-CustomArray'
        'Function:\ConvertTo-DataTable'
        'Function:\Switch-ClassNames'
        'Function:\Load-Data'
        'Function:\Populate-HardwareHashTable'
        'Function:\Reformat-Values'
        'Function:\Display-DataInGUI'
    )

    $Arguments = @(
        $ListItem
        $Hardware
        $Instances
        $Window
        $ComputerName
        $Credentials
        $Source
    )

    $Params = @{
        Code      = $NestedCode
        Arguments = $Arguments
        Runspaces = $Runspaces
        Functions = $Functions
    }

    Start-RSJob @Params
}

# Define the SQL code to run in the runspace
$SQLCode = {
    param($ListItem,$Hardware,$Window,$ComputerName,$ResourceID)

    $StartTime = Get-Date

    Load-SQLData -Window $Window -ListItem $ListItem -Hardware $Hardware -ComputerName $ComputerName -ResourceID $ResourceID

    # Display enum time in window
    $EndTime = Get-Date
    $EnumTime = [math]::Round(($EndTime - $StartTime).TotalSeconds,2)
    $Window.GUI.Dispatcher.Invoke(
        [action]{
            $Window.Enum.Text = $EnumTime
    })
}

#endregion


#region Events

# GUI event which gets the data for the entered computer name when the button is clicked
$Window.Button.add_Click({
        Thunderbirds-AreGo Function:\Check-SystemAndCredentials, Function:\Start-RSJob, Function:\Get-ResourceID -Runspaces $Runspaces -Window $Window
})


# GUI event which gets the data for the entered computer name when the enter button is hit
$Window.ComputerName.Add_KeyDown({
        if ($_.Key -eq 'Return')
        {
            Thunderbirds-AreGo Function:\Check-SystemAndCredentials, Function:\Start-RSJob, Function:\Get-ResourceID -Runspaces $Runspaces -Window $Window
        }
})


# GUI event which disables the listView when the computername gets changed
$Window.ComputerName.Add_TextChanged({
        $Window.listView.IsEnabled = $false
        $Window.Source.IsEnabled = $false
})


# GUI event to populate the datagrid when a list item is selected
$Window.Listview.add_SelectionChanged({
        If ($Window.Source.SelectedItem -match 'Client')
        {
            $Functions = @(
                'Function:\Start-RSJob'
                'Function:\Prepare-CustomArray'
                'Function:\ConvertTo-DataTable'
                'Function:\Switch-ClassNames'
                'Function:\Load-Data'
                'Function:\Populate-HardwareHashTable'
                'Function:\Reformat-Values'
                'Function:\Display-DataInGUI'
            )
            $Params = @{
                Runspaces    = $Runspaces
                Window       = $Window
                DontUpdate   = $DontUpdate
                Code         = $Code
                Hardware     = $Hardware
                Instances    = $Instances
                ComputerName = $ComputerName
            }
            ListViewItem-Selected $Functions @Params
        }

        If ($Window.Source.SelectedItem -match 'ConfigMgr')
        {
            $Functions = @(
                'Function:\Start-RSJob'
                'Function:\Load-SQLData'
            )
            $Params = @{
                Runspaces    = $Runspaces
                Window       = $Window
                DontUpdate   = $DontUpdate
                SQLCode      = $SQLCode
                Hardware     = $Hardware
                ComputerName = $ComputerName
            }
            ListViewItemSQL-Selected $Functions @Params
        }
})

# GUI event to update listView items when source changed
$Window.Source.add_SelectionChanged({
        # Set a variable which prevents the "listViewItem-Selected" code from running if the source has been changed, which also triggers a change of the listitem
        $DontUpdate = $True

        # Clear item list in listview
        $Window.ListView.Items.Clear()

        # Enable the Description box
        $Window.ClassDescription.IsEnabled = $True
    
        # Update the ListView Items
        Update-ListViewItems

        # Release the variable
        $DontUpdate = $false
})

# Change the cursor to a hand when mouse enter the WMI class descsription box
$Window.ClassDescription.add_MouseEnter({
        if ($This.Text -eq '[WMI Class Description]')
        {

        }
        Else
        {
            $This.Cursor = 'Hand'
        }
})

# Open a separate Window when WMI class descsription box is clicked
$Window.ClassDescription.add_PreviewMouseDown({
        if ($This.Text -eq '[WMI Class Description]')
        {

        }
        Else
        {
            Open-WMIDescriptionWindow
        }
})

# Change the cursor to a hand when mouse enter the WMI class descsription box
$Window.About.add_Click({
        Open-AboutWindow
})

# Load the Settings window
$Window.Settings.add_Click({
        Show-SettingsWindow
})

# Change the cursor to a hand when mouse enter the WMI class descsription box
$Window.Export.add_Click({
        Export-ToCSV
})

# On window close
$Window.GUI.Add_Closed({
        # Remove any remaining runspaces
    $Runspaces.GetEnumerator() | where {$_.Name -ne "Number"} | foreach {
    $_.Value.PowerShell.Stop()
    $_.Value.PowerShell.Runspace.Dispose()
    }

    $Runspaces.Clear()
    
        # Remove any CSV files created in the temp directory 
        $CSVFiles | ForEach-Object {
            Remove-Item $_ -Force
        }
})

#endregion




#region DisplayGUI

# If code is running in ISE, use ShowDialog()...
if ($psISE)
{ 
    $null = $Window.GUI.ShowDialog()
}
# ...otherwise run application in it's own app context
Else
{
    # Make PowerShell Disappear #comment our for development
    #$windowcode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);' 
    #$asyncwindow = Add-Type -MemberDefinition $windowcode -name Win32ShowWindowAsync -namespace Win32Functions -PassThru 
    #$null = $asyncwindow::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0)
    
    #$app = [Windows.Application]::new()
    $app = New-Object Windows.Application
    $app.Run($Window.GUI)
    
    #$Window.GUI.ShowDialog()
}
#endregion