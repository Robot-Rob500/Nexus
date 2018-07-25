#--------------------------------------------
# Declare Global Variables and Functions here
#--------------------------------------------

#region Registry Self-Healing

$registryHealPaths = New-Object PSObject -Property @{
	
	DropdownButtons = "HKCU:\Software\Nexus\DropdownButtons"
	DropdownButtonsPanel1 = "HKCU:\Software\Nexus\DropdownButtons\Panel1"
	DropdownButtonsPanel2 = "HKCU:\Software\Nexus\DropdownButtons\Panel2"
	DropdownButtonsPanel3 = "HKCU:\Software\Nexus\DropdownButtons\Panel3"
	DropdownButtonsPanel4 = "HKCU:\Software\Nexus\DropdownButtons\Panel4"
	Folders = "HKCU:\Software\Nexus\Folders"
	Multiping = "HKCU:\Software\Nexus\Multi-Ping"
	Paths = "HKCU:\Software\Nexus\Paths"
	WebBrowsers = "HKCU:\Software\Nexus\WebBrowsers"
	WebsiteGroups = "HKCU:\Software\Nexus\WebsiteGroups"
	Websites = "HKCU:\Software\Nexus\Websites"
	
}

$array = @()
$array = $registryHealPaths.psobject.Properties | Select Name, Value | Sort-Object Name


foreach ($object in $array)
{
	$selfHealTest = Test-Path $object.Value
	if ($selfHealTest -eq $false)
	{
		$baseHealPathLast = ($object.Value).Split("\")[-1]
		$baseHealPath = ($object.Value).Replace("\$baseHealPathLast", "")
		New-Item -Path $baseHealPath -Name $baseHealPathLast -Force
	}
}

#endregion

#region Folder Self-Healing

$folderHealProperties = New-Object PSObject -Property @{
	
	BrowserIconsPath = "C:\ProgramData\Nexus\$env:username\Browser\Icons"
	IconsPath = "C:\ProgramData\Nexus\$env:username\Data\Icons"
	ImagesPath = "C:\ProgramData\Nexus\$env:username\Data\Images"
	InstallLocation = "C:\Program Files\Nexus"
	SCCMNameSpace = "root\sms\site_ENT"
	SCCMServer = "SYS01PCM12BRLW.FMOL-HS.LOCAL"
	ToolsPath = "C:\Program Files\Nexus\Tools"
	
}

$array = @()
$array = $folderHealProperties.psobject.Properties | Select Name, Value | Sort-Object Name

foreach ($object in $array)
{
	New-ItemProperty -Path "HKCU:\Software\Nexus\Paths" -Name $object.Name -Value $object.Value -Force
}

#endregion

#region Global Variables
$global:toolsPath = (Get-ItemProperty -Path "HKCU:\Software\Nexus\Paths" -Name ToolsPath).ToolsPath
$psexec = "$toolsPath\psexec.exe"
$global:browserIconPath = (Get-ItemProperty -Path "HKCU:\Software\Nexus\Paths" -Name "BrowserIconsPath").BrowserIconsPath
$global:programInstallPath = (Get-ItemProperty -Path "HKCU:\Software\Nexus\Paths" -Name "InstallLocation").InstallLocation
$global:programIconPath = (Get-ItemProperty -Path "HKCU:\Software\Nexus\Paths" -Name "IconsPath").IconsPath
$global:programImagePath = (Get-ItemProperty -Path "HKCU:\Software\Nexus\Paths" -Name "ImagesPath").ImagesPath
$global:SCCMServer = (Get-ItemProperty -Path "HKCU:\Software\Nexus\Paths" -Name SCCMServer).SCCMServer
$global:SCCMNameSpace = (Get-ItemProperty -Path "HKCU:\Software\Nexus\Paths" -Name SCCMNameSpace).SCCMNameSpace
#endregion

#region Get-IP

function Get-IP
{
	[Cmdletbinding()]
	Param (
		[alias('dnsHostName')]
		[Parameter(ValueFromPipelineByPropertyName = $true, ValueFromPipeline = $true)]
		[string]$Computername = $Computername
	)
	Process
	{
		$NICs = Get-WmiObject Win32_NetworkAdapterConfiguration -Filter "IPEnabled='$True'" -ComputerName $ComputerName
		foreach ($Nic in $NICs)
		{
			$myobj = @{
				Name = $Nic.Description
				'MAC Address' = $Nic.MACAddress
				IPV4 = $Nic.IPAddress | where{ $_ -match "\d+\.\d+\.\d+\.\d+" }
				IPV6 = $Nic.IPAddress | where{ $_ -match "\:\:" }
				'IPV4 Subnet' = $Nic.IPSubnet | where{ $_ -match "\d+\.\d+\.\d+\.\d+" }
				'Default Gateway' = $Nic.DefaultIPGateway | Select -First 1
				'DNS Server' = $Nic.DNSServerSearchOrder
				'WINS Primary' = $Nic.WINSPrimaryServer
				'WINS Secondary' = $Nic.WINSSecondaryServer
			}
			$obj = New-Object PSObject -Property $myobj
			$obj.PSTypeNames.Clear()
			$obj.PSTypeNames.Add('BSonPosh.IPInfo')
			$obj
		}
	}
}

#endregion

#region Open-RemoteRegistry

function Open-RemoteRegistry ($computerName = "127.0.0.1")
{
	Start-Job  {
		Add-Type -AssemblyName Microsoft.VisualBasic
		Add-Type -AssemblyName System.Windows.Forms
		
		regedit
		
		Start-Sleep -Seconds 1
		
		[Microsoft.VisualBasic.Interaction]::AppActivate("Regedit")
		[System.Windows.Forms.SendKeys]::SendWait("%FC")
		
		Start-Sleep -Seconds 1
		
		[System.Windows.Forms.SendKeys]::SendWait("$using:computerName{ENTER}")
	}
}

#endregion

#region Converto-Datatable

function ConvertTo-DataTable
{
	[OutputType([System.Data.DataTable])]
	param (
		[ValidateNotNull()]
		$InputObject,
		[ValidateNotNull()]
		[System.Data.DataTable]$Table,
		[switch]$RetainColumns,
		[switch]$FilterWMIProperties)
	
	if ($Table -eq $null)
	{
		$Table = New-Object System.Data.DataTable
	}
	
	if ($InputObject -is [System.Data.DataTable])
	{
		$Table = $InputObject
	}
	else
	{
		if (-not $RetainColumns -or $Table.Columns.Count -eq 0)
		{
			#Clear out the Table Contents
			$Table.Clear()
			
			if ($InputObject -eq $null) { return } #Empty Data
			
			$object = $null
			#find the first non null value
			foreach ($item in $InputObject)
			{
				if ($item -ne $null)
				{
					$object = $item
					break
				}
			}
			
			if ($object -eq $null) { return } #All null then empty
			
			#Get all the properties in order to create the columns
			foreach ($prop in $object.PSObject.Get_Properties())
			{
				if (-not $FilterWMIProperties -or -not $prop.Name.StartsWith('__'))#filter out WMI properties
				{
					#Get the type from the Definition string
					$type = $null
					
					if ($prop.Value -ne $null)
					{
						try { $type = $prop.Value.GetType() }
						catch { }
					}
					
					if ($type -ne $null) # -and [System.Type]::GetTypeCode($type) -ne 'Object')
					{
						[void]$table.Columns.Add($prop.Name, $type)
					}
					else #Type info not found
					{
						[void]$table.Columns.Add($prop.Name)
					}
				}
			}
			
			if ($object -is [System.Data.DataRow])
			{
				foreach ($item in $InputObject)
				{
					$Table.Rows.Add($item)
				}
				return @(, $Table)
			}
		}
		else
		{
			$Table.Rows.Clear()
		}
		
		foreach ($item in $InputObject)
		{
			$row = $table.NewRow()
			
			if ($item)
			{
				foreach ($prop in $item.PSObject.Get_Properties())
				{
					if ($table.Columns.Contains($prop.Name))
					{
						$row.Item($prop.Name) = $prop.Value
					}
				}
			}
			[void]$table.Rows.Add($row)
		}
	}
	
	return @(, $Table)
}

#endregion

#region Reopen Dropdown
function Reopen-Dropdown
{
	$formDropDown.Dispose()
	$formDropdown.Close()
	$MainForm.Activate()
	Call-dropdown_psf
}
#endregion

#region Reopen Sites

function Reopen-Sites
{
	$formSites.Dispose()
	$formSites.Close()
	$MainForm.Activate()
	Call-sites_psf
}

#endregion

#region Reopen Site-Add

function Reopen-Site-Add
{
	$global:formSiteAddRelaunched = New-Object PSObject -Property @{
		
		Relaunched = $true
		X = $formSiteAdd.Location.X
		Y = $formSiteAdd.Location.Y
		
	}
	Call-browser-add_psf
	$formSiteAdd.Dispose()
	$formSiteAdd.Close()
	Call-site-add_psf
}

#endregion

#region Reopen Folders

function Reopen-Folders
{
	$formFolders.Dispose()
	$formFolders.Close()
	$MainForm.Activate()
	Call-folders_psf
}

#endregion

#region Sort-ListViewColumn
function Sort-ListViewColumn
{
	param (
		[ValidateNotNull()]
		[Parameter(Mandatory = $true)]
		[System.Windows.Forms.ListView]$ListView,
		[Parameter(Mandatory = $true)]
		[int]$ColumnIndex,
		[System.Windows.Forms.SortOrder]$SortOrder = 'None')
	
	if (($ListView.Items.Count -eq 0) -or ($ColumnIndex -lt 0) -or ($ColumnIndex -ge $ListView.Columns.Count))
	{
		return;
	}
	
	#region Define ListViewItemComparer
	try
	{
		$local:type = [ListViewItemComparer]
	}
	catch
	{
		Add-Type -ReferencedAssemblies ('System.Windows.Forms') -TypeDefinition  @" 
	using System;
	using System.Windows.Forms;
	using System.Collections;
	public class ListViewItemComparer : IComparer
	{
	    public int column;
	    public SortOrder sortOrder;
	    public ListViewItemComparer()
	    {
	        column = 0;
			sortOrder = SortOrder.Ascending;
	    }
	    public ListViewItemComparer(int column, SortOrder sort)
	    {
	        this.column = column;
			sortOrder = sort;
	    }
	    public int Compare(object x, object y)
	    {
			if(column >= ((ListViewItem)x).SubItems.Count)
				return  sortOrder == SortOrder.Ascending ? -1 : 1;
		
			if(column >= ((ListViewItem)y).SubItems.Count)
				return sortOrder == SortOrder.Ascending ? 1 : -1;
		
			if(sortOrder == SortOrder.Ascending)
	        	return String.Compare(((ListViewItem)x).SubItems[column].Text, ((ListViewItem)y).SubItems[column].Text);
			else
				return String.Compare(((ListViewItem)y).SubItems[column].Text, ((ListViewItem)x).SubItems[column].Text);
	    }
	}
"@ | Out-Null
	}
	#endregion
	
	if ($ListView.Tag -is [ListViewItemComparer])
	{
		#Toggle the Sort Order
		if ($SortOrder -eq [System.Windows.Forms.SortOrder]::None)
		{
			if ($ListView.Tag.column -eq $ColumnIndex -and $ListView.Tag.sortOrder -eq 'Ascending')
			{
				$ListView.Tag.sortOrder = 'Descending'
			}
			else
			{
				$ListView.Tag.sortOrder = 'Ascending'
			}
		}
		else
		{
			$ListView.Tag.sortOrder = $SortOrder
		}
		
		$ListView.Tag.column = $ColumnIndex
		$ListView.Sort()#Sort the items
	}
	else
	{
		if ($SortOrder -eq [System.Windows.Forms.SortOrder]::None)
		{
			$SortOrder = [System.Windows.Forms.SortOrder]::Ascending
		}
		
		#Set to Tag because for some reason in PowerShell ListViewItemSorter prop returns null
		$ListView.Tag = New-Object ListViewItemComparer ($ColumnIndex, $SortOrder)
		$ListView.ListViewItemSorter = $ListView.Tag #Automatically sorts
	}
}
#endregion

#region Load-Combobox
function Load-ComboBox
{
	Param (
		[ValidateNotNull()]
		[Parameter(Mandatory = $true)]
		[System.Windows.Forms.ComboBox]$ComboBox,
		[ValidateNotNull()]
		[Parameter(Mandatory = $true)]
		$Items,
		[Parameter(Mandatory = $false)]
		[string]$DisplayMember,
		[switch]$Append
	)
	
	if (-not $Append)
	{
		$ComboBox.Items.Clear()
	}
	
	if ($Items -is [Object[]])
	{
		$ComboBox.Items.AddRange($Items)
	}
	elseif ($Items -is [Array])
	{
		$ComboBox.BeginUpdate()
		foreach ($obj in $Items)
		{
			$ComboBox.Items.Add($obj)
		}
		$ComboBox.EndUpdate()
	}
	else
	{
		$ComboBox.Items.Add($Items)
	}
	
	$ComboBox.DisplayMember = $DisplayMember
}
#endregion

#region Add-ListViewItem

function Add-ListViewItem
{
	Param (
		[ValidateNotNull()]
		[Parameter(Mandatory = $true)]
		[System.Windows.Forms.ListView]$ListView,
		[ValidateNotNull()]
		[Parameter(Mandatory = $true)]
		$Items,
		[int]$ImageIndex = -1,
		[string[]]$SubItems,
		$Group,
		[switch]$Clear)
	
	if ($Clear)
	{
		$ListView.Items.Clear();
	}
	
	$lvGroup = $null
	if ($Group -is [System.Windows.Forms.ListViewGroup])
	{
		$lvGroup = $Group
	}
	elseif ($Group -is [string])
	{
		#$lvGroup = $ListView.Group[$Group] # Case sensitive
		foreach ($groupItem in $ListView.Groups)
		{
			if ($groupItem.Name -eq $Group)
			{
				$lvGroup = $groupItem
				break
			}
		}
		
		if ($lvGroup -eq $null)
		{
			$lvGroup = $ListView.Groups.Add($Group, $Group)
		}
	}
	
	if ($Items -is [Array])
	{
		$ListView.BeginUpdate()
		foreach ($item in $Items)
		{
			$listitem = $ListView.Items.Add($item.ToString(), $ImageIndex)
			#Store the object in the Tag
			$listitem.Tag = $item
			
			if ($SubItems -ne $null)
			{
				$listitem.SubItems.AddRange($SubItems)
			}
			
			if ($lvGroup -ne $null)
			{
				$listitem.Group = $lvGroup
			}
		}
		$ListView.EndUpdate()
	}
	else
	{
		#Add a new item to the ListView
		$listitem = $ListView.Items.Add($Items.ToString(), $ImageIndex)
		#Store the object in the Tag
		$listitem.Tag = $Items
		
		if ($SubItems -ne $null)
		{
			$listitem.SubItems.AddRange($SubItems)
		}
		
		if ($lvGroup -ne $null)
		{
			$listitem.Group = $lvGroup
		}
	}
}

#endregion

#region Load-DataGridView

function Load-DataGridView
{
	Param (
		[ValidateNotNull()]
		[Parameter(Mandatory = $true)]
		[System.Windows.Forms.DataGridView]$DataGridView,
		[ValidateNotNull()]
		[Parameter(Mandatory = $true)]
		$Item,
		[Parameter(Mandatory = $false)]
		[string]$DataMember
	)
	$DataGridView.SuspendLayout()
	$DataGridView.DataMember = $DataMember
	
	if ($Item -is [System.ComponentModel.IListSource]`
	-or $Item -is [System.ComponentModel.IBindingList] -or $Item -is [System.ComponentModel.IBindingListView])
	{
		$DataGridView.DataSource = $Item
	}
	else
	{
		$array = New-Object System.Collections.ArrayList
		
		if ($Item -is [System.Collections.IList])
		{
			$array.AddRange($Item)
		}
		else
		{
			$array.Add($Item)
		}
		$DataGridView.DataSource = $array
	}
	
	$DataGridView.ResumeLayout()
}
#endregion

#region Get-Printers

function Get-Printers($ComputerName)
{
	$output = @()
	if (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet)
	{
		$Hive = [long]$HIVE_HKU = 2147483651
		$sessions = Get-WmiObject -ComputerName $ComputerName -Class win32_process | ?{ $_.name -eq "explorer.exe" }
		if ($sessions)
		{
			foreach ($explorer in $sessions)
			{
				$sid = ($explorer.GetOwnerSid()).sid
				$owner = $explorer.GetOwner()
				$RegProv = get-WmiObject -List -Namespace "root\default" -ComputerName $ComputerName | Where-Object { $_.Name -eq "StdRegProv" }
				$PrinterList = $RegProv.EnumKey($Hive, "$($sid)\Printers\Connections")
				if ($PrinterList.sNames.count -gt 0)
				{
					foreach ($printer in $printerList.sNames)
					{
						"$($printer)`t$(($RegProv.GetStringValue($Hive, "$($sid)\Printers\Connections\$($printer)", "RemotePath")).sValue)"
					}
				}
				else { write-debug "No mapped printers on $($ComputerName)" }
			}
		}
		else { write-debug "explorer.exe not running on $($ComputerName)" }
	}
	else { write-debug "Can't connect to $($ComputerName)" }
	return $output
}

#endregion

#region Update-Hostname-List
function Update-Hostname-List
{
	if (($computername_field.Text -notlike "") -and ($computername_field.Text -notlike $null))
	{
		<#
		$path = @()
		[int]$number = $null
		$path = Get-ItemProperty -Path "HKCU:\Software\Microsoft\Terminal Server Client\Default" -Name *
		[int]$number = ($path | GM | ? { $_.Name -like "MR*" }).Count
		$newEntry = "MRU$number"
		Set-ItemProperty -Path "HKCU:\Software\Microsoft\Terminal Server Client\Default" -Name "$newEntry" -Value "$($computername_field.Text)"
		if ($computername_field.AutoCompleteCustomSource -notcontains "$newEntry")
		{
			$computername_field.AutoCompleteCustomSource.AddRange($($computername_field.Text))
		}#>
		
		$path = @()
		[int]$number = $null
		$path = Get-ItemProperty -Path "HKCU:\Software\Microsoft\Terminal Server Client\Default" -Name *
		$hostNames = @(); $hostNames = $path.psobject.properties | ? { $_.Name -like "MRU*" } | Select Name, Value
		[int]$number = $hostNames.Count
		$newEntry = "MRU$number"
		if ($($hostNames.Value) -notcontains $($computername_field.Text))
		{
			Set-ItemProperty -Path "HKCU:\Software\Microsoft\Terminal Server Client\Default" -Name "$newEntry" -Value "$($computername_field.Text)"
			$computername_field.AutoCompleteCustomSource.AddRange($($computername_field.Text))
		}
	}
}
#endregion

#region Import-AD-Module
function Import-AD-Module
{
	$getADModule = Get-Module
	if ($getADModule.Name -notcontains "ActiveDirectory")
	{
		Import-Module -Name ActiveDirectory
	}
}
#endregion

#region Get-Admin-Crendentials
function Get-Admin-Credentials
{
	$global:adminCredAdd = $true
	$global:standardCredAdd = $false
	Call-credentials_psf
}
#endregion

#region Get-Standard-Crendentials
function Get-Standard-Credentials
{
	$global:adminCredAdd = $false
	$global:standardCredAdd = $true
	Call-credentials_psf
}
#endregion

#region Get-Device-Info

function Get-Device-Info
{
	$global:getDeviceInfo = $true
	$computername_field.Text = $global:deviceInfoName
	Call-system-info_psf
}

#endregion

#region Close-System-Info-Options

function System-Info
{
	$formsysinfooptions.Opacity = 0
	$formsysinfooptions.Close()
	Call-system-info_psf
}

#endregion

#region Mainform-Foreground
function Mainform-Foreground
{
	#region Window Foreground
	$script:nexusForeground = (Get-ItemProperty -Path "HKCU:\Software\Nexus\" -Name "NexusForeground").NexusForeground
	if ($script:nexusForeground -like 'Checked')
	{
		$Mainform.TopMost = $true
		$this.TopMost = $true
	}
	else
	{
		$MainForm.TopMost = $false
		$this.TopMost = $false
	}
	#endregion
}

#endregion