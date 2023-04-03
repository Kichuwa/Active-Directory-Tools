# Baseline tool to work with
 
Get-ADUser -Identity Arons -Properties proxyaddresses | Select-Object Name, @{L = "ProxyAddresses"; E = { ($_.ProxyAddresses -like 'smtp:*') -join ";"}}

# ============================================================================

# Last Major Addition 4/3/23 - DS

Clear-Host

Import-Module ActiveDirectory | Out-Null

$AssemblyDictionary = (
"Accessibility",
"Anonymously Hosted DynamicMethods Assembly",
"MetadataViewProxies_b32ae82b-c6f6-49e4-84e0-e60a2410395f",
"Microsoft.GeneratedCode",
"Microsoft.Management.Infrastructure",
"Microsoft.PowerShell.Commands.Utility",
"Microsoft.PowerShell.Editor",
"Microsoft.PowerShell.GPowerShell",
"Microsoft.PowerShell.GraphicalHost",
"Microsoft.PowerShell.ISECommon",
"Microsoft.PowerShell.Security",
"mscorlib",
"powershell_ise",
"PresentationCore",
"PresentationFramework",
"PresentationFramework.Aero2",
"PresentationFramework-SystemCore",
"PresentationFramework-SystemData",
"PresentationFramework-SystemXml",
"SMDiagnostics",
"System",
"System.ComponentModel.Composition",
"System.Configuration",
"System.Configuration.Install",
"System.Core",
"System.Data",
"System.Diagnostics.Tracing",
"System.DirectoryServices",
"System.Drawing",
"System.Management",
"System.Management.Automation",
"System.Numerics",
"System.Runtime.Serialization",
"System.ServiceModel.Internals",
"System.Transactions",
"System.Windows.Forms",
"System.Xaml",
"System.Xml",
"UIAutomationProvider",
"UIAutomationTypes",
"WindowsBase"
)

# Adds all dependencies to dictionary
Foreach($Assembly in $AssemblyDictionary){

    $tsttemp = $Assembly.ToString()

    [reflection.assembly]::loadwithpartialname("$tsttemp") | Out-Null
}

# Logging Timestamp

$DateFormat = "yyyy-MM-dd HH:mm:ss"
$Date = [DateTime]::Now.ToString($dateFormat)

$TimestampFormat = "HH:mm:ss"
$Timestamp = [DateTime]::Now.ToString($TimestampFormat)

$LogDateFormat = "dd-MM-yyyy"
$LogDate = [DateTime]::Today.ToString($LogDateFormat)

$RootPath = "C:\Temp\"
$AppPath = "C:\Temp\ProxyAddress-Tool\"
$LogFile = "log-$LogDate.txt"




$frmInitialScreen = New-Object system.Windows.Forms.Form
$frmInitialScreen.Text = 'Proxy Address Editor'
$frmInitialScreen.Width = 750
$frmInitialScreen.Height = 550
$frmInitialScreen.Location = New-Object System.Drawing.Size(450,295)
$frmInitialScreen.StartPosition = "CenterScreen"
$frmInitialScreen.TopMost = $true
$frmInitialScreen.MaximumSize = New-Object system.drawing.size(650, 295)
$frmInitialScreen.MinimumSize = New-Object system.drawing.size(650, 295)
$frmInitialScreen.Opacity = 0.95

# Search By user Input area.

$lblSearchItems = New-Object system.windows.Forms.Label
$lblSearchItems.Text = 'Search:'
$lblSearchItems.AutoSize = $true
$lblSearchItems.Width = 25
$lblSearchItems.Height = 10
$lblSearchItems.location = New-Object system.drawing.size(5,10)
$lblSearchItems.Font = "Microsoft Sans Serif,10"
$lblSearchItems.Anchor = "Left, Top"
$frmInitialScreen.controls.Add($lblSearchItems)

$txtGetUsernameInfo = New-Object system.windows.Forms.TextBox
$txtGetUsernameInfo.Width = 180
$txtGetUsernameInfo.Height = 30
$txtGetUsernameInfo.location = New-Object system.drawing.size(5,30)
$txtGetUsernameInfo.Font = "Microsoft Sans Serif,10"
$txtGetUsernameInfo.TabIndex = 2
$txtGetUsernameInfo.KeyUp
$frmInitialScreen.controls.Add($txtGetUsernameInfo)
$txtGetUsernameInfo_KeyDown = {}
$txtGetUsernameInfo_KeyDown = [System.Windows.Forms.KeyEventHandler]{}
$txtGetUsernameInfo_KeyDown = [System.Windows.Forms.KeyEventHandler]{
    if ($_.KeyCode -eq 'Enter')
    {
        $btnSearchByUsername.PerformClick()
    }
}
$txtGetUsernameInfo.Anchor = "Left, Top"
$txtGetUsernameInfo.Add_KeyDown($txtGetUsernameInfo_KeyDown)

$btnSearchByUsername = New-Object system.windows.Forms.Button
$btnSearchByUsername.Text = 'Search By Name'
$btnSearchByUsername.Width = 180
$btnSearchByUsername.Height = 30
$btnSearchByUsername.location = New-Object system.drawing.size(5,55)
$btnSearchByUsername.Font = "Microsoft Sans Serif,10"
$btnSearchByUsername.TabStop = $false
$btnSearchByUsername.Anchor = "Left, Top"
$frmInitialScreen.controls.Add($btnSearchByUsername)

# Add and Delete Buttons

$btnAddProxy = New-Object system.windows.Forms.Button
$btnAddProxy.Text = 'Add Proxy Address'
$btnAddProxy.Width = 180
$btnAddProxy.Height = 45
$btnAddProxy.location = New-Object system.drawing.size(5,105)
$btnAddProxy.Font = "Microsoft Sans Serif,10"
$btnAddProxy.TabStop = $false
$btnAddProxy.Anchor = "Left, Top"
$frmInitialScreen.controls.Add($btnAddProxy)

$btnDeleteProxy = New-Object system.windows.Forms.Button
$btnDeleteProxy.Text = 'Remove Proxy Address'
$btnDeleteProxy.Width = 180
$btnDeleteProxy.Height = 45
$btnDeleteProxy.location = New-Object system.drawing.size(5,150)
$btnDeleteProxy.Font = "Microsoft Sans Serif,10"
$btnDeleteProxy.TabStop = $false
$btnDeleteProxy.Anchor = "Left, Top"
$frmInitialScreen.controls.Add($btnDeleteProxy)

# Exit Button

$btnExit = New-Object system.windows.Forms.Button
$btnExit.Text = 'Exit'
$btnExit.Width = 180
$btnExit.Height = 30
$btnExit.location = New-Object system.drawing.size(5,215)
$btnExit.Font = "Microsoft Sans Serif,10"
$btnExit.TabStop = $false
$btnExit.Anchor = "Left, Top"
$frmInitialScreen.controls.Add($btnExit)

# User Listings

$lblUsers = New-Object system.windows.Forms.Label
$lblUsers.Text = 'Users:'
$lblUsers.AutoSize = $true
$lblUsers.Width = 25
$lblUsers.Height = 10
$lblUsers.location = New-Object system.drawing.size(195,10)
$lblUsers.Font = "Microsoft Sans Serif,10"
$lblUsers.Anchor = "Left, Top"
$frmInitialScreen.controls.Add($lblUsers)

$lstShowUser = New-Object system.windows.Forms.ListBox
$lstShowUser.Text = "listView"
$lstShowUser.Width = 200
$lstShowUser.Height = 225
$lstShowUser.location = New-Object system.drawing.point(195,30)
$lstShowUser.TabIndex = 4
$lstShowUser.SelectionMode = "One"
$lstShowUser.Anchor = "Left, Right, Top"
$frmInitialScreen.controls.Add($lstShowUser)


# Show Proxies
$lblProxies = New-Object system.windows.Forms.Label
$lblProxies.Text = 'Proxies:'
$lblProxies.AutoSize = $true
$lblProxies.Width = 25
$lblProxies.Height = 10
$lblProxies.location = New-Object system.drawing.size(410,10)
$lblProxies.Font = "Microsoft Sans Serif,10"
$lblProxies.Anchor = "Left, Top"
$frmInitialScreen.controls.Add($lblProxies)

$lstShowProxy = New-Object system.windows.Forms.ListBox
$lstShowProxy.Text = "listView"
$lstShowProxy.Width = 200
$lstShowProxy.Height = 225
$lstShowProxy.location = New-Object system.drawing.point(410,30)
$lstShowProxy.TabIndex = 4
$lstShowProxy.SelectionMode = "One"
$lstShowProxy.Anchor = "Left, Right, Top"
$frmInitialScreen.controls.Add($lstShowProxy)

# ================================



# ================================
[void] $frmInitialScreen.ShowDialog()