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
$frmInitialScreen.Location = New-Object System.Drawing.Size(450,300)
$frmInitialScreen.StartPosition = "CenterScreen"
$frmInitialScreen.TopMost = $true
$frmInitialScreen.MaximumSize = New-Object system.drawing.size(650, 400)
$frmInitialScreen.MinimumSize = New-Object system.drawing.size(550, 300)

# Search By user Input area.

$lblSearchItems = New-Object system.windows.Forms.Label
$lblSearchItems.Text = 'Search:'
$lblSearchItems.AutoSize = $true
$lblSearchItems.Width = 25
$lblSearchItems.Height = 10
$lblSearchItems.location = New-Object system.drawing.size(5,20)
$lblSearchItems.Font = "Microsoft Sans Serif,10"
$lblSearchItems.Anchor = "Left, Top"
$frmInitialScreen.controls.Add($lblSearchItems)

$txtGetUsernameInfo = New-Object system.windows.Forms.TextBox
$txtGetUsernameInfo.Width = 180
$txtGetUsernameInfo.Height = 30
$txtGetUsernameInfo.location = New-Object system.drawing.size(5,40)
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
$btnSearchByUsername.location = New-Object system.drawing.size(5,65)
$btnSearchByUsername.Font = "Microsoft Sans Serif,10"
$btnSearchByUsername.TabStop = $false
$btnSearchByUsername.Anchor = "Left, Top"
$frmInitialScreen.controls.Add($btnSearchByUsername)





# ================================
[void] $frmInitialScreen.ShowDialog()