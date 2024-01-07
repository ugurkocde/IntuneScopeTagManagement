Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# Define WPF XAML
[xml]$xaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Intune Scope Tag Management"  Height="520" Width="780" ResizeMode="NoResize">

    <Grid Margin="5,10,0,2">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="246*"/>
            <ColumnDefinition Width="529*"/>
            <ColumnDefinition Width="0*"/>
        </Grid.ColumnDefinitions>
        <ListBox x:Name="scopetags_listbox" Margin="381,149,0,0" VerticalAlignment="Top" Height="243" RenderTransformOrigin="1.393,0.499" Grid.Column="1" HorizontalAlignment="Left" Width="130" BorderBrush="Black" FontSize="14">
            <ListBox.ItemTemplate>
                <DataTemplate>
                    <CheckBox Content="{Binding Name}" IsChecked="{Binding IsSelected, Mode=TwoWay}" />
                </DataTemplate>
            </ListBox.ItemTemplate>
        </ListBox>
        <Button x:Name="refresh_button" Content="&#x51;" Margin="0,121,165,0" VerticalAlignment="Top" Background="#FF96FF98" HorizontalAlignment="Right" Width="23" Height="23" FontSize="16" FontFamily="Wingdings 3" Grid.Column="1"/>

        <Button x:Name="connect_graph_button" Content="Connect to Graph" Margin="0,37,317,0" VerticalAlignment="Top" Background="#FFFFE793" Grid.Column="1" HorizontalAlignment="Right" Width="140" Height="28" FontSize="16"/>
        <Button x:Name="assign_scope_tags_button" Content="Assign Scope Tag" Margin="381,397,0,0" VerticalAlignment="Top" Background="#FFCCFF91" Grid.Column="1" HorizontalAlignment="Left" Width="130" Height="31" FontSize="16"/>
        <TextBox x:Name="searchBox" HorizontalAlignment="Left" Margin="142,439,0,0" VerticalAlignment="Top" Width="272" Height="29" Grid.ColumnSpan="2" FontSize="16"/>

        <ListView x:Name="device_list" Margin="27,149,165,0" Grid.ColumnSpan="2" Height="279" VerticalAlignment="Top" BorderBrush="Black" FontSize="14">
            <ListView.View>
                <GridView>
                    <GridViewColumn/>
                </GridView>
            </ListView.View>
        </ListView>
        <TextBlock HorizontalAlignment="Left" Margin="27,439,0,0" TextWrapping="Wrap" VerticalAlignment="Top" FontSize="16" Height="21" Width="115"><Run Text="Search"/><Run Text=" Device"/><Run Text=":"/></TextBlock>
        <TextBlock HorizontalAlignment="Center" TextWrapping="Wrap" VerticalAlignment="Top" FontSize="24" Width="358" Grid.ColumnSpan="2" Height="32" FontWeight="Bold"><Run Text="Intune Scope"/><Run Language="de-de" Text=" "/><Run Text="Tag Management"/></TextBlock>
        <TextBlock Margin="27,123,89,0" TextWrapping="Wrap" FontSize="16" FontWeight="Bold" Height="21" VerticalAlignment="Top"><Run Language="de-de" Text="Devices"/></TextBlock>
        <TextBlock Margin="381,123,60,0" TextWrapping="Wrap" VerticalAlignment="Top" FontSize="16" FontWeight="Bold" Grid.Column="1" Height="21"><Run Language="de-de" Text="Scope Tags"/></TextBlock>
        <Button x:Name="export_button" Content="Export Devices and Scope Tags" Margin="281,439,0,0" VerticalAlignment="Top" Background="#FF83B2FF" Grid.Column="1" HorizontalAlignment="Left" Width="230" Height="30" FontSize="16"/>
        <TextBlock HorizontalAlignment="Left" Margin="27,94,0,0" TextWrapping="Wrap" VerticalAlignment="Top" FontSize="16" FontWeight="Bold" Height="21" Width="76"><Run Text="Tenant"/><Run Language="de-de" Text=" ID:"/></TextBlock>
        <TextBlock HorizontalAlignment="Left" Margin="27,68,0,0" TextWrapping="Wrap" VerticalAlignment="Top" FontSize="16" FontWeight="Bold" Height="21" Width="66"><Run Language="de-de" Text="Account:"/></TextBlock>
        <TextBlock x:Name="tenant_id_text" HorizontalAlignment="Left" Margin="117,94,0,0" TextWrapping="Wrap" VerticalAlignment="Top" FontSize="16" Width="377" Text="" Grid.ColumnSpan="2" Height="21"/>
        <TextBlock x:Name="account_text" HorizontalAlignment="Left" Margin="117,68,0,0" TextWrapping="Wrap" VerticalAlignment="Top" FontSize="16" Width="377" Text="" Grid.ColumnSpan="2" Height="21"/>
        <TextBlock HorizontalAlignment="Left" Margin="419,30,0,0" TextWrapping="Wrap" VerticalAlignment="Top" FontSize="10" Height="32" Width="92" Grid.Column="1" FontStyle="Italic"><Run Text="Version 1.0"/><LineBreak/><Run Text="Author: "/><Run Language="de-de" Text="@ugurkocde"/></TextBlock>
        <Button x:Name="twitter_button" Content="Twitter" Margin="419,10,0,0" VerticalAlignment="Top" Background="White" Grid.Column="1" HorizontalAlignment="Left" Width="40" Height="18" FontSize="10"/>
        <Button x:Name="github_button" Content="GitHub" Margin="465,10,0,0" VerticalAlignment="Top" Background="White" Grid.Column="1" HorizontalAlignment="Left" Width="40" Height="18" FontSize="10"/>


    </Grid>

    </Window>
"@

# Parse XAML
$reader = (New-Object System.Xml.XmlNodeReader $xaml) 
$Window = [Windows.Markup.XamlReader]::Load( $reader )


# Define GridViewColumns for device name, ID and Scope Tags
$deviceListView = $Window.FindName('device_list')
$gridView = New-Object System.Windows.Controls.GridView
$deviceListView.View = $gridView

$deviceNameColumn = New-Object System.Windows.Controls.GridViewColumn
$deviceNameColumn.Header = "Device Name"
$deviceNameColumn.DisplayMemberBinding = New-Object System.Windows.Data.Binding("DeviceName")
$gridView.Columns.Add($deviceNameColumn)

$deviceIdColumn = New-Object System.Windows.Controls.GridViewColumn
$deviceIdColumn.Header = "Device ID"
$deviceIdColumn.DisplayMemberBinding = New-Object System.Windows.Data.Binding("DeviceId")
$gridView.Columns.Add($deviceIdColumn)

$scopeTagsColumn = New-Object System.Windows.Controls.GridViewColumn
$scopeTagsColumn.Header = "Scope Tag"
$scopeTagsColumn.DisplayMemberBinding = New-Object System.Windows.Data.Binding("ScopeTags")
$gridView.Columns.Add($scopeTagsColumn)

# Define a DataTemplate for the Checkbox
$dataTemplate = New-Object System.Windows.DataTemplate
$frameworkElementFactory = New-Object System.Windows.FrameworkElementFactory([System.Windows.Controls.CheckBox])

# Corrected Binding
$binding = New-Object System.Windows.Data.Binding("Selected")
$frameworkElementFactory.SetValue([System.Windows.Controls.CheckBox]::IsCheckedProperty, $binding)

$dataTemplate.VisualTree = $frameworkElementFactory

# Add a new GridViewColumn for the Checkbox
$checkBoxColumn = New-Object System.Windows.Controls.GridViewColumn
$checkBoxColumn.CellTemplate = $dataTemplate
$checkBoxColumn.Header = "Select"
$gridView.Columns.Insert(0, $checkBoxColumn) 

# Global variable to store all devices
$global:allDevices = @()

# Function to Get Managed Devices
function Get-ManagedDevices {
    param($Uri)
    $devices = @()
    try {
        $response = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/managedDevices" -Method GET

        foreach ($device in $response.value) {
            # Retrieve scope tags for the device
            $deviceDetailsResponse = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/managedDevices('$($device.id)')" -Method GET
            $scopeTagIds = $deviceDetailsResponse.roleScopeTagIds

            # Fetch the display names of each scope tag
            $scopeTagNames = @()
            foreach ($tagId in $scopeTagIds) {
                $tagDetailsResponse = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/roleScopeTags/$tagId" -Method GET
                $scopeTagNames += $tagDetailsResponse.displayName
            }

            $devices += New-Object PSObject -Property @{
                DeviceName = $device.deviceName  
                DeviceId = $device.id             
                ScopeTags = ($scopeTagNames -join ", ")
                Selected = $false
            }
        }

        # Recursively call for next link if available
        if ($response.'@odata.nextLink') {
            $devices += Get-ManagedDevices -Uri $response.'@odata.nextLink'
        }
    } catch {
        Write-Host "Error fetching devices: $($_.Exception.Message)"
    }

    $global:allDevices += $devices
    return $devices
}


# Set the list view's initial items source
$deviceListView.ItemsSource = $global:allDevices

# Filter the device list
function Filter-DeviceList {
    param($SearchText)
    $filteredDevices = @($global:allDevices | Where-Object { $_.DeviceName -like "*$SearchText*" -or $_.DeviceId -like "*$SearchText*" })
    $deviceListView.ItemsSource = $filteredDevices
}

$searchBox = $Window.FindName('searchBox')
$searchBox.Add_TextChanged({ Filter-DeviceList $searchBox.Text })

# Function to Get Scope Tags
function Get-ScopeTags {
    param($Uri)
    $scopeTags = @()
    try {
        $response = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/roleScopeTags" -Method GET

        foreach ($scopeTag in $response.value) {
            $scopeTags += New-Object PSObject -Property @{
                Name = $scopeTag.displayName
                Id = $scopeTag.id
                IsSelected = $false  # Add IsSelected property for checkbox binding
            }
        }

        # Recursively call for next link if available
        if ($response.'@odata.nextLink') {
            $scopeTags += Get-ScopeTags -Uri $response.'@odata.nextLink'
        }
    } catch {
        Write-Host "Error fetching scope tags: $($_.Exception.Message)"
    }

    # Sort the scope tags alphabetically by their Name
    $sortedScopeTags = $scopeTags | Sort-Object -Property Name

    return $sortedScopeTags
}

# Set the list box's initial items source
$scopetags_listbox = $Window.FindName('scopetags_listbox')

# Assign Scope Tag to Devices
function Assign-ScopeTag {
    $selectedDevices = $global:allDevices | Where-Object { $_.Selected -eq $true }
    $selectedScopeTagItems = $scopetags_listbox.Items | Where-Object { $_.IsSelected -eq $true }

    foreach ($device in $selectedDevices) {
        if ($selectedScopeTagItems.Count -eq 0) {
            Write-Host "No scope tags selected for device $($device.DeviceId)"
            continue
        }

        $scopeTagIds = $selectedScopeTagItems | ForEach-Object { $_.Id }

        $uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices('$($device.DeviceId)')"
        $body = @{
            roleScopeTagIds = @($scopeTagIds)
        } | ConvertTo-Json

        # Print URI and body to the console for debugging
        Write-Host "URI: $uri"
        Write-Host "Body: $body"

        Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $body -ContentType "application/json"
        Write-Host "Attempted to set scope tags for device $($device.DeviceId)"
    }
        # Refresh data after scope tag assignment
        Get-Data
}

$assignScopeTagsButton = $Window.FindName('assign_scope_tags_button')
$assignScopeTagsButton.Add_Click({ Assign-ScopeTag })

function Get-Data {
    Write-Host "Fetching Data..."

    # Fetch Managed Devices
    $global:allDevices = Get-ManagedDevices -Uri $devicesUri
    $deviceListView.Dispatcher.Invoke([Action]{
        $deviceListView.ItemsSource = $global:allDevices
    })

    # Fetch Scope Tags
    $sortedScopeTags = Get-ScopeTags -Uri $scopetagsUri
    $scopetags_listbox.Dispatcher.Invoke([Action]{
        $scopetags_listbox.ItemsSource = $sortedScopeTags
    })
}

# Function to handle the Connect button click
function Connect-GraphButton {
    try {
        if (-not (Get-MgContext)) {
            Connect-MgGraph -Scopes "DeviceManagementManagedDevices.ReadWrite.All", "DeviceManagementRBAC.Read.All"

            # Check if the connection has been established
            $mgContext = Get-MgContext
            if ($mgContext -and $mgContext.Account -and $mgContext.TenantId) {
                $connectButton.IsEnabled = $false

                UpdateLoginInfo  # Update tenant and account info

                Get-Data  # Fetch devices and scope tags
            } else {
           }
        } else {
            Get-Data  # Fetch devices and scope tags
        }
    } catch {
        [System.Windows.MessageBox]::Show("Failed to connect to Microsoft Graph. Error: $($_.Exception.Message)", "Connection Error")
    }
}

# Function to update tenant and account information
function UpdateLoginInfo {
    $mgContext = Get-MgContext
    $tenantId = $mgContext.TenantId
    $accountId = $mgContext.Account

    $Window.Dispatcher.Invoke([Action]{
        $Window.FindName('tenant_id_text').Text = $tenantId
        $Window.FindName('account_text').Text = $accountId
    })
}

$devicesUri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices"
$scopetagsUri = "https://graph.microsoft.com/beta/deviceManagement/roleScopeTags"

# Find the Connect button and attach the event handler
$connectButton = $Window.FindName('connect_graph_button')
$connectButton.Add_Click({ Connect-GraphButton })

# Function to Export Data to CSV
function Export-Data {
    # Create Save File Dialog
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV files (*.csv)|*.csv"
    $saveFileDialog.Title = "Export Devices and ScopeTags in a CSV File"
    $saveFileDialog.FileName = "DevicesAndScopeTags.csv"

    # Show the Save File Dialog
    $result = $saveFileDialog.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $filePath = $saveFileDialog.FileName
        if (-not [string]::IsNullOrWhiteSpace($filePath)) {
            # Prepare data for export
            $exportData = $global:allDevices | Select-Object DeviceName, ScopeTags

            # Export to CSV
            $exportData | Export-Csv -Path $filePath -NoTypeInformation

            # Optionally, you can open the folder where the file is saved
            $folderPath = Split-Path -Parent $filePath
            Start-Process "explorer.exe" $folderPath
        }
    }
}

# Refresh data when the Refresh button is clicked
$refreshButton = $Window.FindName('refresh_button')
$refreshButton.Add_Click({ Get-Data })

# Find the Export button and attach the event handler
$exportButton = $Window.FindName('export_button')
$exportButton.Add_Click({ Export-Data })

# Check if already connected to Microsoft Graph and fetch data if so
if (Get-MgContext) {
    UpdateLoginInfo  # Call to update tenant and account info
    Get-Data
    $connectButton.IsEnabled = $false
}

# Find the Twitter button and attach the event handler
$twitterButton = $Window.FindName('twitter_button')
$twitterButton.Add_Click({
    Start-Process "https://twitter.com/UgurKocDe"
})

# Find the GitHub button and attach the event handler
$githubButton = $Window.FindName('github_button')
$githubButton.Add_Click({
    Start-Process "https://github.com/ugurkocde/IntuneScopeTagManagement"
})

# Show Window
$Window.ShowDialog() | Out-Null