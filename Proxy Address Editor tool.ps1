# Last Major Addition 4/4/23 - DS

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

$LogDateFormat = "MM-dd-yyyy"
$LogDate = [DateTime]::Today.ToString($LogDateFormat)

$RootPath = "C:\Temp\"
$AppPath = "C:\Temp\ProxyAddress-Tool\"
$LogFile = "log-$LogDate.txt"


# Initial Form

$frmInitialScreen = New-Object system.Windows.Forms.Form
$frmInitialScreen.Text = 'Proxy Address Editor'
$frmInitialScreen.Width = 850
$frmInitialScreen.Height = 550
$frmInitialScreen.Location = New-Object System.Drawing.Size(550,295)
$frmInitialScreen.StartPosition = "CenterScreen"
$frmInitialScreen.TopMost = $true
$frmInitialScreen.MaximumSize = New-Object system.drawing.size(950, 295)
$frmInitialScreen.MinimumSize = New-Object system.drawing.size(750, 295)
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
$lstShowUser.Anchor = "Left, Top"
$frmInitialScreen.controls.Add($lstShowUser)


# Show Proxies

$lblProxies = New-Object system.windows.Forms.Label
$lblProxies.Text = 'Proxy Addresses:'
$lblProxies.AutoSize = $true
$lblProxies.Width = 25
$lblProxies.Height = 10
$lblProxies.location = New-Object system.drawing.size(410,10)
$lblProxies.Font = "Microsoft Sans Serif,10"
$lblProxies.Anchor = "Left, Right, Top"
$frmInitialScreen.controls.Add($lblProxies)

$lstShowProxy = New-Object system.windows.Forms.ListBox
$lstShowProxy.Text = "listView"
$lstShowProxy.Width = 400
$lstShowProxy.Height = 225
$lstShowProxy.location = New-Object system.drawing.point(410,30)
$lstShowProxy.TabIndex = 4
$lstShowProxy.SelectionMode = "One"
$lstShowProxy.Anchor = "Left, Right, Top"
$frmInitialScreen.controls.Add($lstShowProxy)


# End of Current GUI Stage =======================================

#Connection test

#Initial test to check if the user has access to AD by trying to search the name of the current user.

$ErrorActionPreference = 'continue'
$CurrentUser = [Environment]::UserName

$RootCheck = Test-Path $RootPath
$AppCheck = Test-Path $AppPath

if($RootCheck -eq $true){
    if($AppCheck -eq $true){

        ("") | Out-File $AppPath$LogFile -Append
        ("=====================") | Out-File $AppPath$LogFile -Append
        ("Launched Proxy Editor at $Date") | Out-File $AppPath$LogFile -Append
        ("=====================") | Out-File $AppPath$LogFile -Append
    }
    else{

        mkdir $AppPath
        ("") | Out-File $AppPath$LogFile -Append
        ("=====================") | Out-File $AppPath$LogFile -Append
        ("Launched Proxy Editor at $Date") | Out-File $AppPath$LogFile -Append
        ("=====================") | Out-File $AppPath$LogFile -Append
    }
}
else{

    mkdir $RootPath
    mkdir $AppPath
    ("") | Out-File $AppPath$LogFile -Append
    ("=====================") | Out-File $AppPath$LogFile -Append
    ("Launched Proxy Editor at $Date") | Out-File $AppPath$LogFile -Append
    ("=====================") | Out-File $AppPath$LogFile -Append
}

Try{

    Get-ADUser -Filter "Name -like '*$CurrentUser*'" | Select-Object name, samaccountname
    $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
    ("$Timestamp - Connected to Active Directory as < $CurrentUser >")| Out-File $AppPath$LogFile -Append
    ("-~-") | Out-File $AppPath$LogFile -Append

}Catch{

    $BlankLine = ("------")
    $ConnectionFailed = ("Unable to connect to Active Directory.")
    $ConnectionFailed2 = ("Please ensure you have AD Services installed including the Powershell Module.")

    $lstVerboseOutput.Items.Add($BlankLine)
    $lstVerboseOutput.Items.Add($ConnectionFailed)
    $lstVerboseOutput.Items.Add($BlankLine)
    $lstVerboseOutput.Items.Add($ConnectionFailed2)
    $lstVerboseOutput.Items.Add($BlankLine)

    $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
    ("$Timestamp - $ConnectionFailed")| Out-File $AppPath$LogFile -Append
    ("$Timestamp - $ConnectionFailed2")| Out-File $AppPath$LogFile -Append
    ("-~-") | Out-File $AppPath$LogFile -Append
}



# Search Active Directory for users that match a string to some degree which then pulls the closest match
# ================================
# Button for search
# ================================
$btnSearchByUsername.Add_Click({

    BeginningTimestamp("Search by Username.")

    $lstShowUser.SelectionMode = "MultiExtended"
    if($txtGetUsernameInfo.Text -eq ""){
        
        Alert "No Query" "You must enter a user to search for."
    }
    else{
        # Clear Both Boxes on Search
        Try{
            foreach($Item in $lstShowUser){
                $lstShowUser.Items.Clear()
                $lstShowProxy.Items.Clear()
            }
        }
        Catch{
            continue
        }
        # Attempt to see if users match a string of characters
        Try{
            $UsernameSearch = $txtGetUsernameInfo.Text
            $UsernameSearch = $UsernameSearch -replace ".{0}$"
            $UserProfiles = Get-ADUser -Filter "Name -like '*$UsernameSearch*'" | Select-Object name, samaccountname
            
            # If no Match disable ability to select from list
            if($UserProfiles.Count -eq 0){
                $lstShowUser.Enabled = $false
                $lstShowUser.Items.Add("No Results for $UsernameSearch")
                $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                ("$Timestamp - User  search for < $UsernameSearch > returned no results.")| Out-File $AppPath$LogFile -Append
            }
            # If match allow Enum
            else {
                $lstShowUser.Enabled = $true
                foreach($User in $UserProfiles){

                    $Username = $User.name
                    if($Username -notmatch "tmp"){
                        if($Username -notmatch "tmpl"){
    
                            $lstShowUser.Items.Add($Username)
                            $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                            ("$Timestamp - User  search for < $UsernameSearch > returned - $Username.")| Out-File $AppPath$LogFile -Append
                        }
                    }
                }
            }
        }Catch{
            # Throw if extraneous error occurs. I.E. no AD connectivity.
            Alert "Error" "Nothing to show."
            $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
            ("$Timestamp - Nothing to show.")| Out-File $AppPath$LogFile -Append
        }
        ("-~-") | Out-File $AppPath$LogFile -Append
    }
})

# When a user highlights a user in the results box from user search output is shown in second box
# ====================================
# Update Prox List on User Highlight
# ====================================

$lstShowUser.Add_SelectedIndexChanged({

    $SelectedUser = $lstShowUser.SelectedItem

    BeginningTimestamp("Selected user $SelectedUser")

    # Selects all proxy Adderesses that match a SMTP or smtp object.
    $Proxies = Get-ADUser -Filter "Name -eq '$SelectedUser'" -Properties proxyaddresses | Select-Object -ExpandProperty proxyaddresses | Where-Object {$_ -like 'smtp:*'} 

    # Clear List on new Selection
    $lstShowProxy.Items.Clear()

    # Disable the Proxy output if no valid return
    if($Proxies.Count -eq 0){
        $lstShowProxy.Enabled = $false
        $lstShowProxy.Items.Add("No Proxy Address found for $SelectedUser")
        $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
        ("$Timestamp - No Proxy Address found for $SelectedUser")| Out-File $AppPath$LogFile -Append
    }
    # Enable and allow use of proxy listing is returns are valid
    else {
        $lstShowProxy.Enabled = $true
        foreach($Proxy in $Proxies){
            $lstShowProxy.Items.Add($Proxy)
            $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
            ("$Timestamp - Proxy Search for < $SelectedUser > returned - $Proxy.")| Out-File $AppPath$LogFile -Append
        }
    }
    ("-~-") | Out-File $AppPath$LogFile -Append
})

# Creates a pop-up for the 
# ================================
# Add Proxy Button
# ================================

# TODO
$btnAddProxy.Add_Click({

})


# Deletes the proxy from the selected users returned proxies.
# ================================
# Delete Proxy Button
# ================================

# TODO
$btnDeleteProxy.Add_Click({
    
})



# Enforces the use of the exit button such that any other method of closing will not work including forced close.
# ================================
# Exit Button Function
# ================================
$btnExit.Add_Click({

    BeginningTimestamp("Disconnected from Active Directory as < $CurrentUser > - Program Closed.")
    
    $frmInitialScreen.Add_FormClosing({
        $_.Cancel=$false
    })

    $frmInitialScreen.Close()
})

# Basic Timestamp Function for the tracking of method using 
# ================================
# Timestamp Function
# ================================
Function BeginningTimestamp([String] $functionLabel){

    $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
    ("") | Out-File $AppPath$LogFile -Append
    ("") | Out-File $AppPath$LogFile -Append
    ("$Timestamp - FUNCTION - $functionLabel") | Out-File $AppPath$LogFile -Append
}


# Alert Box for Users
# Overloaded Method to Ensure this cannot be pushed behind initial application and ignored
# ================================
# MessageBox Function
# ================================
function Show-MessageBox {
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true, Position = 0)]
        [string]$Message,

        [parameter(Mandatory = $false)]
        [string]$Title = 'MessageBox in PowerShell',

        [ValidateSet("OKOnly", "OKCancel", "AbortRetryIgnore", "YesNoCancel", "YesNo", "RetryCancel")]
        [string]$Buttons = "OKCancel",

        [ValidateSet("Critical", "Question", "Exclamation", "Information")]
        [string]$Icon = "Information"
    )
    Add-Type -AssemblyName Microsoft.VisualBasic

    [Microsoft.VisualBasic.Interaction]::MsgBox($Message, "$Buttons,SystemModal,$Icon", $Title)
}

# Shortened Call function for the alert
# Note: This works as a POWERSHELL Method, not DOTNET therefore do not use parenthesis when using the function
# ================================
# Alert Function
# ================================
function Alert([String] $Title, [String] $Message){
    Show-MessageBox -Title "$Title" -Message "$Message" -Icon Information -Buttons OKOnly
}


# ========================================================
# Show Screen and Ensure the X button cannot be used.
# ========================================================
$frmInitialScreen.add_FormClosing({
    $_.Cancel = $true
})
[void] $frmInitialScreen.ShowDialog()