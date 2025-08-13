Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

# XAML layout
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="CM Instant Deploy" Height="400" Width="500" ResizeMode="CanResize" WindowStartupLocation="CenterScreen"
        Background="#FFEDF0F5" FontFamily="Segoe UI">
    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="40"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="120"/>
        </Grid.ColumnDefinitions>

        <!-- Title -->
        <TextBlock Grid.Row="0" Grid.Column="0" Grid.ColumnSpan="3" Text="CM Instant Deploy" 
                   FontSize="24" FontWeight="Bold" Foreground="#FF2D2D30" HorizontalAlignment="Center"/>
        <Separator Grid.Row="1" Grid.ColumnSpan="3" Margin="0,10"/>

        <!-- Computer Name -->
        <Label Grid.Row="2" Grid.Column="0" Content="Computer Name:" VerticalAlignment="Center" FontSize="14"/>
        <TextBox Grid.Row="2" Grid.Column="1" Name="Input_ComputerName" Height="28" FontSize="14" Padding="5"/>
        <Button Grid.Row="2" Grid.Column="2" Name="Button_TestConnection" Content="Test Connection" 
                Height="28" Width="110" Margin="10,0,0,0" Background="#FF007ACC" Foreground="White" BorderThickness="0"
                Cursor="Hand" FontWeight="SemiBold"/>

        <!-- Status Text -->
        <TextBlock Grid.Row="3" Grid.Column="0" Grid.ColumnSpan="3" Name="TextBlock_Status" 
                   Text="" FontSize="13" Foreground="Gray" Margin="0,15,0,0" TextWrapping="Wrap"/>

        <!-- CM Server Name -->
        <Label Grid.Row="4" Grid.Column="0" Content="CM Server Name:" VerticalAlignment="Center" FontSize="14" Visibility="Collapsed" Name="Label_CMServer"/>
        <TextBox Grid.Row="4" Grid.Column="1" Name="Input_CMServer" Height="28" FontSize="14" Padding="5" Visibility="Collapsed"/>

        <!-- CM Site Code -->
        <Label Grid.Row="5" Grid.Column="0" Content="CM Site Code:" VerticalAlignment="Center" FontSize="14" Visibility="Collapsed" Name="Label_CMSite"/>
        <TextBox Grid.Row="5" Grid.Column="1" Name="Input_CMSite" Height="28" FontSize="14" Padding="5" Visibility="Collapsed"/>

        <!-- App Picker -->
        <Button Grid.Row="6" Grid.Column="1" Name="Button_SelectApp" Content="Select Application" 
                Height="28" Width="140" Margin="0,15,0,0" Background="#FF4CAF50" Foreground="White" 
                BorderThickness="0" Cursor="Hand" FontWeight="SemiBold" Visibility="Collapsed" HorizontalAlignment="Left"/>

        <!-- New Status TextBlock below SelectApp button -->
        <TextBlock Grid.Row="7" Grid.Column="1" Name="TextBlock_Status2" 
                   Text="" FontSize="13" Foreground="Gray" Margin="0,10,0,0" TextWrapping="Wrap"/>

        <!-- Deploy Button (hidden by default), width removed for dynamic sizing -->
        <Button Grid.Row="8" Grid.Column="1" Name="Button_Deploy" Content="Deploy Application" 
                Height="28" MinWidth="140" Margin="0,10,0,0" Background="#FF2196F3" Foreground="White" 
                BorderThickness="0" Cursor="Hand" FontWeight="SemiBold" Visibility="Collapsed" HorizontalAlignment="Left"/>
    </Grid>
</Window>
"@

# Load XAML
$reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Bind elements
$Input_ComputerName = $window.FindName("Input_ComputerName")
$Button_TestConnection = $window.FindName("Button_TestConnection")
$TextBlock_Status = $window.FindName("TextBlock_Status")
$Label_CMServer = $window.FindName("Label_CMServer")
$Input_CMServer = $window.FindName("Input_CMServer")
$Label_CMSite = $window.FindName("Label_CMSite")
$Input_CMSite = $window.FindName("Input_CMSite")
$Button_SelectApp = $window.FindName("Button_SelectApp")
$Button_Deploy = $window.FindName("Button_Deploy")
$TextBlock_Status2 = $window.FindName("TextBlock_Status2")

$global:ManagementPoint = (Get-CimInstance -ClassName SMS_Authority -Namespace 'root\ccm').CurrentManagementPoint  
$siteCodePath = (Get-CimInstance -ClassName SMS_Authority -Namespace 'root\ccm').Name.Replace(":","_")
$global:siteCode = $siteCodePath.Replace("SMS_","")


# Expand form fields after successful ping
function Show-AdvancedFields {
    $Label_CMServer.Visibility = "Visible"
    $Input_CMServer.Visibility = "Visible"
    $Label_CMSite.Visibility = "Visible"
    $Input_CMSite.Visibility = "Visible"
    $Button_SelectApp.Visibility = "Visible"
    
}

# Event: Test Connection
$Button_TestConnection.Add_Click({
    $computer = $Input_ComputerName.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($computer)) {
        $TextBlock_Status.Text = "⚠️ Please enter a computer name."
        return
    }
    $TextBlock_Status.Text = "🔄 Testing connection to $computer ..."
    $ping = Test-Connection -ComputerName $computer -Count 2 -Quiet -ErrorAction SilentlyContinue
    if ($ping) {
        $TextBlock_Status.Text = "✅ Connection to $computer succeeded."
        Show-AdvancedFields
        $Input_CMSite.text = $global:siteCode
        $Input_CMServer.Text = $global:ManagementPoint
        $
    } else {
        $TextBlock_Status.Text = "❌ Connection to $computer failed."
    }
})

# Event: Select Application
$Button_SelectApp.Add_Click({

[CmdletBinding()]
param (
    [string]$SiteServer = $global:ManagementPoint,  # Change to your SCCM site server
    [string]$SiteCode   = $global:siteCode         # Change to your SCCM site code
)
$Button_SelectApp.Content = "Please wait..."
Add-Type -AssemblyName PresentationFramework


function Get-SCCMApplicationTree {
    $namespace = "root\sms\site_$SiteCode"

    Write-Host "Querying folders..."
    $folders = Get-WmiObject -Namespace "root\sms\site_$global:siteCode" -Class SMS_ObjectContainerNode -ComputerName $global:ManagementPoint |
        Where-Object { $_.ObjectType -eq 6000 }

    Write-Host "Querying folder items..."
    $folderItems = Get-WmiObject -Namespace "root\sms\site_$global:siteCode" -Class SMS_ObjectContainerItem -ComputerName $global:ManagementPoint |
        Where-Object { $_.ObjectType -eq 6000 }

    Write-Host "Querying applications via Get-CMApplication..."
    Function Connect-SCCM-Server{

param(
$SiteCode,
$servername
)

#
# Press 'F5' to run this script. Running this script will load the ConfigurationManager
# module for Windows PowerShell and will connect to the site.
#
# This script was auto-generated at '11/2/2023 3:58:08 PM'.

# Site configuration
$SiteCode = $SiteCode # Site code 
$ProviderMachineName = $servername # SMS Provider machine name

# Customizations
$initParams = @{}
#$initParams.Add("Verbose", $true) # Uncomment this line to enable verbose logging
#$initParams.Add("ErrorAction", "Stop") # Uncomment this line to stop the script on any errors

# Do not change anything below this line

# Import the ConfigurationManager.psd1 module 
if((Get-Module ConfigurationManager) -eq $null) {
    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
}

# Connect to the site's drive if it is not already present
if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
}

# Set the current location to be the site code.
Set-Location "$($SiteCode):\" @initParams
}
Connect-SCCM-Server -SiteCode $global:siteCode -servername $global:ManagementPoint
    $applications = Get-CMApplication -Fast

    # Build folder lookup: key = ContainerNodeID
    $folderLookup = @{ }
    $folderLookup[0] = @{
        Name = "Applications"
        ParentID = $null
        Items = @()
    }
    foreach ($folder in $folders | Sort-Object Name) {
        $folderLookup[$folder.ContainerNodeID] = @{
            Name = $folder.Name
            ParentID = $folder.ParentContainerNodeID
            Items = @()
        }
    }
    Write-Host "Folder lookup entries: $($folderLookup.Count)"

    # Build app lookup by extracting GUID from CI_UniqueID
    $appLookup = @{}
    foreach ($app in $applications) {
        if ($app.CI_UniqueID -match "Application_([0-9a-f\-]+)") {
            $guid = $matches[1]
            $displayName = if ($app.LocalizedDisplayName) { $app.LocalizedDisplayName } else { $app.Name }
            $appLookup[$guid] = $displayName
        }
    }
    Write-Host "App lookup entries: $($appLookup.Count)"

    # Map folder items to apps by matching folder InstanceKey with app GUID
    $mappedApps = 0
    foreach ($item in $folderItems) {
        $containerID = $item.ContainerNodeID
        $instanceKey = $item.InstanceKey

        if (-not $folderLookup.ContainsKey($containerID)) {
            Write-Verbose "Skipping orphan container: $containerID"
            continue
        }

        # InstanceKey looks like "ScopeId_xxx/Application_GUID", extract GUID part
        if ($instanceKey -match "Application_([0-9a-f\-]+)") {
            $appGuid = $matches[1]
            if ($appLookup.ContainsKey($appGuid)) {
                $folderLookup[$containerID].Items += $appLookup[$appGuid]
                $mappedApps++
            }
        }
    }
    Write-Host "Mapped applications to folders: $mappedApps"

    # Add apps with no folder mapping to root folder
    $mappedAppGuids = $folderItems | ForEach-Object {
        if ($_.InstanceKey -match "Application_([0-9a-f\-]+)") { $matches[1] }
    }
    foreach ($guid in $appLookup.Keys) {
        if (-not ($mappedAppGuids -contains $guid)) {
            $folderLookup[0].Items += $appLookup[$guid]
        }
    }
    Write-Host "Orphan applications added to root: $($folderLookup[0].Items.Count)"

    # Sort folder items alphabetically
    foreach ($key in $folderLookup.Keys) {
        $folderLookup[$key].Items = $folderLookup[$key].Items | Sort-Object
    }

    return $folderLookup
}

function Build-TreeNodes {
    param (
        $lookup,
        $parentID = 0
    )
    $nodes = @()
    foreach ($key in ($lookup.Keys | Sort-Object { $lookup[$_].Name })) {
        if ($lookup[$key].ParentID -eq $parentID) {
            $node = New-Object System.Windows.Controls.TreeViewItem
            $node.Header = $lookup[$key].Name
            $node.FontWeight = "Bold"

            foreach ($appName in $lookup[$key].Items) {
                $appNode = New-Object System.Windows.Controls.TreeViewItem
                $appNode.Header = $appName
                $appNode.Tag = "Application"
                $node.Items.Add($appNode) | Out-Null
            }

            $childNodes = Build-TreeNodes -lookup $lookup -parentID $key
            foreach ($childNode in $childNodes) {
                $node.Items.Add($childNode) | Out-Null
            }

            $nodes += $node
        }
    }
    return $nodes
}

Write-Host "Starting to build SCCM Applications Tree..."

$lookup = Get-SCCMApplicationTree
$treeNodes = Build-TreeNodes -lookup $lookup

Write-Host "Launching GUI window..."

$global:selectedApp = $null  # Variable to hold the selected application name

$window = New-Object System.Windows.Window
$window.Title = "SCCM Applications"
$window.Width = 600
$window.Height = 800
$window.WindowStartupLocation = "CenterScreen"

$treeView = New-Object System.Windows.Controls.TreeView
foreach ($node in $treeNodes) {
    $treeView.Items.Add($node) | Out-Null
}

$treeView.Add_SelectedItemChanged({
    if ($treeView.SelectedItem -is [System.Windows.Controls.TreeViewItem]) {
        $selected = $treeView.SelectedItem
        if ($selected.Tag -eq "Application") {
            $global:selectedApp = $selected.Header
            $window.Close()
        }
    }
})

$window.Content = $treeView
$window.ShowDialog() | Out-Null

# After window closes, you can use $selectedApp to get the chosen application
Write-Host "Selected application: $global:selectedApp"
#$Button_SelectApp.content = $global:selectedApp
$TextBlock_Status2.Visibility = "visible"
$TextBlock_Status2.text = "✅ $global:SelectedApp selected."
$Button_SelectApp.Content = "Change application..."
$computer = $Input_ComputerName.text
$Button_Deploy.Content = "Deploy $global:selectedApp to $computer"
$Button_Deploy.Visibility = "Visible"
})

$Button_Deploy.Add_Click({
    $date = (get-date).ToShortDateString()
    $computer = $Input_ComputerName.Text
    $ResourceID = (get-cmdevice -Name $computer).resourceID
    $popupWindow = New-Object System.Windows.Window
    $popupWindow.Title = "Deployment Status"
    $popupWindow.Width = 500
    $popupWindow.Height = 350
    $popupWindow.WindowStartupLocation = "CenterScreen"

    $textBox = New-Object System.Windows.Controls.TextBox
    $textBox.AcceptsReturn = $true
    $textBox.TextWrapping = "Wrap"
    $textBox.VerticalScrollBarVisibility = "Auto"
    $textBox.IsReadOnly = $true
    $textBox.FontSize = 14
    $textBox.Margin = '10'
    $textBox.Height = 300
    $popupWindow.Content = $textBox

    $popupWindow.Show()

    # Append each message with an icon, update UI immediately after each append

    $textBox.AppendText("🛠️  Creating temporary device collection.`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{})
    New-CMDeviceCollection -Name "Temp - $computer - $global:selectedApp - $date" -LimitingCollectionName "All Workstations" -ErrorAction Stop 


    $textBox.AppendText("🖥️  Adding $computer to device collection.`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{})
    Add-CMDeviceCollectionDirectMembershipRule -CollectionName "Temp - $computer - $global:selectedApp - $date" -ResourceId $ResourceID -ErrorAction stop
   
    $textBox.AppendText("📤  Creating deployment.`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{})
    New-CMApplicationDeployment -Name $global:selectedApp -DeployAction Install -DeployPurpose Available -ErrorAction stop -CollectionName "Temp - $computer - $global:selectedApp - $date" 
    
     
    $textBox.AppendText("😴  Sleeping.`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{})
    Start-Sleep -Seconds 15
   
    $textBox.AppendText("🔄  Forcing $computer to check in.`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{})
    Invoke-WmiMethod -ComputerName $computer -Namespace root\ccm -Class sms_client -Name TriggerSchedule "{00000000-0000-0000-0000-000000000021}" 
    Invoke-WmiMethod -ComputerName $computer -Namespace root\ccm -Class sms_client -Name TriggerSchedule "{00000000-0000-0000-0000-000000000121}" 
    Invoke-WmiMethod -ComputerName $computer -Namespace root\ccm -Class sms_client -Name TriggerSchedule "{00000000-0000-0000-0000-000000000002}" 
    Invoke-WmiMethod -ComputerName $computer -Namespace root\ccm -Class sms_client -Name TriggerSchedule "{00000000-0000-0000-0000-000000000113}" 
    Invoke-WmiMethod -ComputerName $computer -Namespace root\ccm -Class sms_client -Name TriggerSchedule "{00000000-0000-0000-0000-000000000022}" 
    Invoke-WmiMethod -ComputerName $computer -Namespace root\ccm -Class sms_client -Name TriggerSchedule "{00000000-0000-0000-0000-000000000001}" 
    

     $textBox.AppendText("😴  Sleeping again (20 seconds).`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{})
    Start-Sleep -Seconds 20

    
# Max attempts: 5 minutes / 30 seconds = 10 checks
$maxAttempts = 10
$attempt = 1

while ($attempt -le $maxAttempts) {
    $textBox.AppendText("🔍  Checking to see if application is available for install.`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{}
    )

   # Get the SCCM application unique ID
       Function Connect-SCCM-Server{

param(
$SiteCode,
$servername
)

#
# Press 'F5' to run this script. Running this script will load the ConfigurationManager
# module for Windows PowerShell and will connect to the site.
#
# This script was auto-generated at '11/2/2023 3:58:08 PM'.

# Site configuration
$SiteCode = $SiteCode # Site code 
$ProviderMachineName = $servername # SMS Provider machine name

# Customizations
$initParams = @{}
#$initParams.Add("Verbose", $true) # Uncomment this line to enable verbose logging
#$initParams.Add("ErrorAction", "Stop") # Uncomment this line to stop the script on any errors

# Do not change anything below this line

# Import the ConfigurationManager.psd1 module 
if((Get-Module ConfigurationManager) -eq $null) {
    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
}

# Connect to the site's drive if it is not already present
if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
}

# Set the current location to be the site code.
Set-Location "$($SiteCode):\" @initParams
}
Connect-SCCM-Server -SiteCode $global:siteCode -servername $global:ManagementPoint
$CMID = Get-CMApplication -Fast |
    Where-Object LocalizedDisplayName -eq $global:selectedApp |
    Select-Object -ExpandProperty CI_UniqueID

# Remove the trailing "/<number>" if present
$CMIDTrimmed = $CMID -replace "/\d+$",""

# Get available applications on the client and check for match
$AvailableApplications = Get-WmiObject -ComputerName $computer `
    -Query "SELECT * FROM CCM_Application" `
    -Namespace "ROOT\ccm\ClientSDK" |
Where-Object { $_.ID -eq $CMIDTrimmed }

    if ($AvailableApplications) {
        $textBox.AppendText("🚦  Application is available, starting install.`r`n")
        $textBox.ScrollToEnd()
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
            [System.Windows.Threading.DispatcherPriority]::Background,
            [action]{}
        )

       $appID = $AvailableApplications.ID
       $appRevision = $AvailableApplications.Revision
       $appMachineTarget = $AvailableApplications.IsMachineTarget

       try{
            Invoke-Command -ComputerName $Computer -ScriptBlock {
             ([WmiClass]'Root\CCM\ClientSDK:CCM_Application').Install($using:appID, $using:appRevision, $using:appMachineTarget, 0, "Normal", $false)
                }

                  $textBox.AppendText("🚀  Application install started!`r`n")
        $textBox.ScrollToEnd()
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
            [System.Windows.Threading.DispatcherPriority]::Background,
            [action]{}
        )

                }

                catch{

                  $textBox.AppendText("❌  Application install failed to start.`r`n")
        $textBox.ScrollToEnd()
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
            [System.Windows.Threading.DispatcherPriority]::Background,
            [action]{}
        )
        break
                }


                 do {


                
                    $WMI =  Get-WmiObject -ComputerName $computer -Query "SELECT * FROM CCM_Application" -Namespace "ROOT\ccm\ClientSDK" | Where-Object { $_.ID -eq $CMIDTrimmed}
                    $WMIEval = $WMI.EvaluationState
                    $WMIInstallState = $WMI.InstallState
                    $WMIError = $WMI.ErrorCode
                    $InProgress = $WMI.InProgressActions
                    switch ($WMIEval) {
                       0  { 
    $Content = "No state information is available." 
    $textBox.AppendText("➤  Status: $Content`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{}
    )
}
1  { 
    $Content = "Application is enforced to desired/resolved state." 
    $textBox.AppendText("➤  Status: $Content`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{}
    )
}
2  { 
    $Content = "Application isn't required on the client." 
    $textBox.AppendText("➤  Status: $Content`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{}
    )
}
3  { 
    $Content = "Application is available for enforcement (install or uninstall based on resolved state). Content may/may not have been downloaded." 
    $textBox.AppendText("➤  Status: $Content`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{}
    )
}
4  { 
    $Content = "Application last failed to enforce (install/uninstall)." 
    $textBox.AppendText("➤  Status: $Content`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{}
    )
}
5  { 
    $Content = "Application is currently waiting for content download to complete." 
    $textBox.AppendText("➤  Status: $Content`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{}
    )
}
6  { 
    $Content = "Application is currently waiting for content download to complete." 
    $textBox.AppendText("➤  Status: $Content`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{}
    )
}
7  { 
    $Content = "Application is currently waiting for its dependencies to download." 
    $textBox.AppendText("➤  Status: $Content`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{}
    )
}
8  { 
    $Content = "Application is currently waiting for a service (maintenance) window." 
    $textBox.AppendText("➤  Status: $Content`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{}
    )
}
9  { 
    $Content = "Application is currently waiting for a previously pending reboot." 
    $textBox.AppendText("➤  Status: $Content`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{}
    )
}
10 { 
    $Content = "Application is currently waiting for serialized enforcement." 
    $textBox.AppendText("➤  Status: $Content`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{}
    )
}
11 { 
    $Content = "Application is currently enforcing dependencies." 
    $textBox.AppendText("➤  Status: $Content`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{}
    )
}
12 { 
    $Content = "Application is currently installing." 
    $textBox.AppendText("➤  Status: $Content`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{}
    )
}
13 { 
    $Content = "Application install/uninstall enforced and soft reboot is pending." 
    $textBox.AppendText("➤  Status: $Content`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{}
    )
}
14 { 
    $Content = "Application installed/uninstalled and hard reboot is pending." 
    $textBox.AppendText("➤  Status: $Content`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{}
    )
}
15 { 
    $Content = "Update is available but pending installation." 
    $textBox.AppendText("➤  Status: $Content`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{}
    )
}
16 { 
    $Content = "Application failed to evaluate." 
    $textBox.AppendText("➤  Status: $Content`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{}
    )
}
17 { 
    $Content = "Application is currently waiting for an active user session to enforce." 
    $textBox.AppendText("➤  Status: $Content`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{}
    )
}
18 { 
    $Content = "Application is currently waiting for all users to sign out." 
    $textBox.AppendText("➤  Status: $Content`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{}
    )
}
19 { 
    $Content = "Application is currently waiting for a user sign in." 
    $textBox.AppendText("➤  Status: $Content`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{}
    )
}
20 { 
    $Content = "Application in progress, waiting for retry." 
    $textBox.AppendText("➤  Status: $Content`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{}
    )
}
21 { 
    $Content = "Application is waiting for presentation mode to be switched off." 
    $textBox.AppendText("➤  Status: $Content`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{}
    )
}
22 { 
    $Content = "Application is pre-downloading content (downloading outside of install job)." 
    $textBox.AppendText("➤  Status: $Content`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{}
    )
}
23 { 
    $Content = "Application is pre-downloading dependent content (downloading outside of install job)." 
    $textBox.AppendText("➤  Status: $Content`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{}
    )
}
24 { 
    $Content = "Application download failed (downloading during install job)." 
    $textBox.AppendText("➤  Status: $Content`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{}
    )
}
25 { 
    $Content = "Application pre-downloading failed (downloading outside of install job)." 
    $textBox.AppendText("➤  Status: $Content`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{}
    )
}
26 { 
    $Content = "Download success (downloading during install job)." 
    $textBox.AppendText("➤  Status: $Content`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{}
    )
}
27 { 
    $Content = "Post-enforce evaluation." 
    $textBox.AppendText("➤  Status: $Content`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{}
    )
}
28 { 
    $Content = "Waiting for network connectivity." 
    $textBox.AppendText("➤  Status: $Content`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{}
    )
}


                      
                    }


                    Start-Sleep -Seconds 15

          
                } until (!($InProgress))

$InstallState = (Get-WmiObject -ComputerName $computer -Query "SELECT * FROM CCM_Application"  -Namespace "ROOT\ccm\ClientSDK" |Where-Object { $_.ID -eq $CMIDTrimmed}).installstate
if ($InstallState -eq "Installed"){
$Content = "SUCCESS: Successfull installed $global:selectedApp on $computer ! " 
    $textBox.AppendText("🎉  $Content`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{}
    )
}
else{
$ErrorCode = (Get-WmiObject -ComputerName $computer -Query "SELECT * FROM CCM_Application"  -Namespace "ROOT\ccm\ClientSDK" |Where-Object { $_.ID -eq $CMIDTrimmed}).ErrorCode
$Content = "FAILURE: Failed to install $global:selectedApp on $computer ! ERROR CODE: $ErrorCode " 
    $textBox.AppendText("👎  $Content`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{}
    )
}


$Content = "Cleaning up deplyoment & device collection " 
    $textBox.AppendText("🗑️  $Content`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{}
        )

        # Call your install script/function here
        # Example: Install-MyApp -ComputerName $computer -AppName $global:selectedApp



        
        Remove-CMDeviceCollection -Name "Temp - $computer - $global:selectedApp - $date" -Force -Confirm:$false


        $Content = "Complete! " 
    $textBox.AppendText("🏁  $Content`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{}
        )

        # Call your install script/function here
        # Example: Install-MyApp -ComputerName $computer -AppName $global:selectedApp


        break
    }
    else {
        Invoke-WmiMethod -ComputerName $computer -Namespace root\ccm -Class sms_client -Name TriggerSchedule "{00000000-0000-0000-0000-000000000021}" 
    Invoke-WmiMethod -ComputerName $computer -Namespace root\ccm -Class sms_client -Name TriggerSchedule "{00000000-0000-0000-0000-000000000121}" 
    Invoke-WmiMethod -ComputerName $computer -Namespace root\ccm -Class sms_client -Name TriggerSchedule "{00000000-0000-0000-0000-000000000002}" 
    Invoke-WmiMethod -ComputerName $computer -Namespace root\ccm -Class sms_client -Name TriggerSchedule "{00000000-0000-0000-0000-000000000113}" 
    Invoke-WmiMethod -ComputerName $computer -Namespace root\ccm -Class sms_client -Name TriggerSchedule "{00000000-0000-0000-0000-000000000022}" 
    Invoke-WmiMethod -ComputerName $computer -Namespace root\ccm -Class sms_client -Name TriggerSchedule "{00000000-0000-0000-0000-000000000001}" 
        $textBox.AppendText("❌  Application not available, sleeping... (Attempt $attempt of $maxAttempts)`r`n")
        $textBox.ScrollToEnd()
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
            [System.Windows.Threading.DispatcherPriority]::Background,
            [action]{}
        )

        Start-Sleep -Seconds 30
    }

    $attempt++
}

if ($attempt -gt $maxAttempts) {
    $textBox.AppendText("⏱  Application did not become available within 5 minutes.`r`n")
    $textBox.ScrollToEnd()
}
    <#
    $textBox.AppendText("🚦  Application is available, starting install.`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{})
    Start-Sleep -Seconds 2

    $textBox.AppendText("😴  Sleeping, will check status in 20 seconds.`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{})
    Start-Sleep -Seconds 2

    $textBox.AppendText("🎉  Application successfully installed.`r`n")
    $textBox.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action]{})#>
})




# Show main window
$window.ShowDialog() | Out-Null
