# Last Major Addition 4/25/23 - DS
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

$DateFormat = "yyyy-MM-dd HH:mm:ss"
$Date = [DateTime]::Now.ToString($dateFormat)

$TimestampFormat = "HH:mm:ss"
$Timestamp = [DateTime]::Now.ToString($TimestampFormat)

$LogDateFormat = "dd-MM-yyyy"
$LogDate = [DateTime]::Today.ToString($LogDateFormat)

$RootPath = "C:\Temp\"
$AppPath = "C:\Temp\AD-Tool\"
$LogFile = "log-$LogDate.txt"


# ========================================================
# Start of GUI layout
# ========================================================

#Form structure

$frmInitialScreen = New-Object system.Windows.Forms.Form
$frmInitialScreen.Text = 'Active Directory Tool'
$frmInitialScreen.Width = 640
$frmInitialScreen.Height = 710
$frmInitialScreen.Location = New-Object System.Drawing.Size(200,300)
$frmInitialScreen.StartPosition = "CenterScreen"
$frmInitialScreen.TopMost = $true
$frmInitialScreen.MaximumSize = New-Object system.drawing.size(640, 900)
$frmInitialScreen.MinimumSize = New-Object system.drawing.size(640, 710)
$frmInitialScreen.MaximizeBox = $false
$frmInitialScreen.Minimizebox = $false
$frmInitialScreen.ControlBox = $false


$lblSearchItems = New-Object system.windows.Forms.Label
$lblSearchItems.Text = 'Search:'
$lblSearchItems.AutoSize = $true
$lblSearchItems.Width = 25
$lblSearchItems.Height = 10
$lblSearchItems.location = New-Object system.drawing.size(5,20)
$lblSearchItems.Font = "Microsoft Sans Serif,10"
$lblSearchItems.Anchor = "Left, Top"
$frmInitialScreen.controls.Add($lblSearchItems)

$lblLeftSideBox = New-Object system.windows.Forms.Label
$lblLeftSideBox.Text = 'Source:'
$lblLeftSideBox.AutoSize = $true
$lblLeftSideBox.Width = 25
$lblLeftSideBox.Height = 10
$lblLeftSideBox.location = New-Object system.drawing.size(195,20)
$lblLeftSideBox.Font = "Microsoft Sans Serif,10"
$lblLeftSideBox.Anchor = "Left, Top"
$frmInitialScreen.controls.Add($lblLeftSideBox)

$lblRightSideBox = New-Object system.windows.Forms.Label
$lblRightSideBox.Text = 'Destination:'
$lblRightSideBox.AutoSize = $true
$lblRightSideBox.Width = 25
$lblRightSideBox.Height = 10
$lblRightSideBox.location = New-Object system.drawing.size(418,20)
$lblRightSideBox.Font = "Microsoft Sans Serif,10"
$lblRightSideBox.Anchor = "Left, Top"
$frmInitialScreen.controls.Add($lblRightSideBox)

$lblVerboseOutput = New-Object system.windows.Forms.Label
$lblVerboseOutput.Text = 'Verbose:'
$lblVerboseOutput.AutoSize = $true
$lblVerboseOutput.Width = 25
$lblVerboseOutput.Height = 10
$lblVerboseOutput.location = New-Object system.drawing.size(4,362)
$lblVerboseOutput.Font = "Microsoft Sans Serif,10"
$frmInitialScreen.controls.Add($lblVerboseOutput)

$lblFootnote = New-Object system.windows.Forms.Label
$lblFootnote.Text = "-This will require domain admin to copy users, unlock, or locate lockouts-" 
$lblFootnote.AutoSize = $true
$lblFootnote.Width = 25
$lblFootnote.Height = 10
$lblFootnote.location = New-Object system.drawing.size(4,656) # was 4 450
$lblFootnote.Font = "Microsoft Sans Serif,8"
$lblFootnote.Anchor = "Left, Bottom"
$frmInitialScreen.controls.Add($lblFootnote)

# Search Text Box

$txtGetFirstGroupInfo = New-Object system.windows.Forms.TextBox
$txtGetFirstGroupInfo.Width = 178
$txtGetFirstGroupInfo.Height = 20
$txtGetFirstGroupInfo.location = New-Object system.drawing.size(5,40)
$txtGetFirstGroupInfo.Font = "Microsoft Sans Serif,10"
$txtGetFirstGroupInfo.TabIndex = 0
$frmInitialScreen.controls.Add($txtGetFirstGroupInfo)
$txtGetFirstGroupInfo.Anchor = "Left, Top"

# Buttons below are vertical (Searches and Clear)

$btnShowFirstGroups = New-Object system.windows.Forms.Button
$btnShowFirstGroups.Text = 'Search Source Group'
$btnShowFirstGroups.Width = 180
$btnShowFirstGroups.Height = 45
$btnShowFirstGroups.location = New-Object system.drawing.size(4,70)
$btnShowFirstGroups.Font = "Microsoft Sans Serif,10"
$btnShowFirstGroups.TabStop = $false
$btnShowFirstGroups.Anchor = "Left, Top"
$frmInitialScreen.controls.Add($btnShowFirstGroups)

$btnShowSecondGroups = New-Object system.windows.Forms.Button
$btnShowSecondGroups.Text = 'Search Destination Group'
$btnShowSecondGroups.Width = 180
$btnShowSecondGroups.Height = 45
$btnShowSecondGroups.location = New-Object system.drawing.size(4,120)
$btnShowSecondGroups.Font = "Microsoft Sans Serif,10"
$btnShowSecondGroups.TabStop = $false
$btnShowSecondGroups.Anchor = "Left, Top"
$frmInitialScreen.controls.Add($btnShowSecondGroups)

$btnSearchByUsername = New-Object system.windows.Forms.Button
$btnSearchByUsername.Text = 'Search By Name'
$btnSearchByUsername.Width = 180
$btnSearchByUsername.Height = 45
$btnSearchByUsername.location = New-Object system.drawing.size(4,170)
$btnSearchByUsername.Font = "Microsoft Sans Serif,10"
$btnSearchByUsername.TabStop = $false
$btnSearchByUsername.Anchor = "Left, Top"
$frmInitialScreen.controls.Add($btnSearchByUsername)

$btnClearListBoxes = New-Object system.windows.Forms.Button
$btnClearListBoxes.Text = 'Clear'
$btnClearListBoxes.Width = 180
$btnClearListBoxes.Height = 45
$btnClearListBoxes.location = New-Object system.drawing.size(4,220)
$btnClearListBoxes.Font = "Microsoft Sans Serif,10"
$btnClearListBoxes.TabIndex = 3
$btnClearListBoxes.Anchor = "Left, Top"
$frmInitialScreen.controls.Add($btnClearListBoxes)


#Buttons below are horizontal

$btnUserInfo = New-Object system.windows.Forms.Button
$btnUserInfo.Text = 'User Info'
$btnUserInfo.Width = 150
$btnUserInfo.Height = 40
$btnUserInfo.location = New-Object system.drawing.size(4,270)
$btnUserInfo.Font = "Microsoft Sans Serif,10"
$btnUserInfo.TabIndex = 6
$frmInitialScreen.controls.Add($btnUserInfo)

$btnShowHelp = New-Object system.windows.Forms.Button
$btnShowHelp.Text = 'Show Help'
$btnShowHelp.Width = 150
$btnShowHelp.Height = 40
$btnShowHelp.location = New-Object system.drawing.size(4,310)
$btnShowHelp.Font = "Microsoft Sans Serif,10"
$btnShowHelp.TabIndex = 13
$frmInitialScreen.controls.Add($btnShowHelp)


$btnShowFirstUsers = New-Object system.windows.Forms.Button
$btnShowFirstUsers.Text = 'Show Group Members'
$btnShowFirstUsers.Width = 150
$btnShowFirstUsers.Height = 40
$btnShowFirstUsers.location = New-Object system.drawing.size(160,270)
$btnShowFirstUsers.Font = "Microsoft Sans Serif,10"
$btnShowFirstUsers.TabIndex = 8
$frmInitialScreen.controls.Add($btnShowFirstUsers)

$btnShowMemberships = New-Object system.windows.Forms.Button
$btnShowMemberships.Text = 'Show Users Groups'
$btnShowMemberships.Width = 150
$btnShowMemberships.Height = 40
$btnShowMemberships.location = New-Object system.drawing.size(160,310)
$btnShowMemberships.Font = "Microsoft Sans Serif,10"
$btnShowMemberships.TabIndex = 9
$frmInitialScreen.controls.Add($btnShowMemberships)




$btnCopyUsersOver = New-Object system.windows.Forms.Button
$btnCopyUsersOver.Text = 'Copy All Users'
$btnCopyUsersOver.Width = 150
$btnCopyUsersOver.Height = 40
$btnCopyUsersOver.location = New-Object system.drawing.size(315,270)
$btnCopyUsersOver.Font = "Microsoft Sans Serif,10"
$btnCopyUsersOver.TabIndex = 10
$frmInitialScreen.controls.Add($btnCopyUsersOver)

$btnCopySelectedUsers = New-Object system.windows.Forms.Button
$btnCopySelectedUsers.Text = 'Copy Selected Users'
$btnCopySelectedUsers.Width = 150
$btnCopySelectedUsers.Height = 40
$btnCopySelectedUsers.location = New-Object system.drawing.size(315,310)
$btnCopySelectedUsers.Font = "Microsoft Sans Serif,10"
$btnCopySelectedUsers.TabIndex = 7
$frmInitialScreen.controls.Add($btnCopySelectedUsers)




$btnGetLockoutBlame = New-Object system.windows.Forms.Button
$btnGetLockoutBlame.Text = 'Lockout Location'
$btnGetLockoutBlame.Width = 150
$btnGetLockoutBlame.Height = 40
$btnGetLockoutBlame.location = New-Object system.drawing.size(470,270)
$btnGetLockoutBlame.Font = "Microsoft Sans Serif,10"
$btnCopyUsersOver.TabIndex = 11
$frmInitialScreen.controls.Add($btnGetLockoutBlame)

$btnUnlockAccount = New-Object system.windows.Forms.Button
$btnUnlockAccount.Text = 'Unlock User'
$btnUnlockAccount.Width = 150
$btnUnlockAccount.Height = 40
$btnUnlockAccount.location = New-Object system.drawing.size(470,310)
$btnUnlockAccount.Font = "Microsoft Sans Serif,10"
$btnCopyUsersOver.TabIndex = 11
$frmInitialScreen.controls.Add($btnUnlockAccount)


$btnExit = New-Object system.windows.Forms.Button
$btnExit.Text = 'Exit'
$btnExit.Width = 75
$btnExit.Height = 40
$btnExit.location = New-Object system.drawing.size(545,610)
$btnExit.Font = "Microsoft Sans Serif,10"
$btnExit.TabIndex = 14
$btnExit.Anchor = "Bottom, Right"
$frmInitialScreen.controls.Add($btnExit)

#List boxes below

$lstShowFirstGroups = New-Object system.windows.Forms.ListBox
$lstShowFirstGroups.Text = "listView"
$lstShowFirstGroups.Width = 200
$lstShowFirstGroups.Height = 230
$lstShowFirstGroups.location = New-Object system.drawing.point(195,40)
$lstShowFirstGroups.TabIndex = 4
$lstShowFirstGroups.SelectionMode = "One"
$lstShowFirstGroups.Anchor = "Left, Right, Top"
$frmInitialScreen.controls.Add($lstShowFirstGroups)

$lstShowSecondGroups = New-Object system.windows.Forms.ListBox
$lstShowSecondGroups.Text = "listView"
$lstShowSecondGroups.Width = 200
$lstShowSecondGroups.Height = 230
$lstShowSecondGroups.location = New-Object system.drawing.point(419,40)
$lstShowSecondGroups.Anchor = "Right, Left, Top"
$lstShowSecondGroups.TabIndex = 5

$frmInitialScreen.controls.Add($lstShowSecondGroups)

$lstVerboseOutput = New-Object system.windows.Forms.ListBox
$lstVerboseOutput.Text = "listView"
$lstVerboseOutput.Width = 615
$lstVerboseOutput.Height = 270 # was 120
$lstVerboseOutput.location = New-Object system.drawing.point(4,383)
$lstVerboseOutput.TabStop = $false
$lstVerboseOutput.Anchor = "Top, Bottom, Left, Right"
$frmInitialScreen.controls.Add($lstVerboseOutput)

# ========================================================
# End of GUI Layout
# ========================================================


# ========================================================
# Start of Active Directory Verification
# ========================================================

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
        ("Launched AD-Tool at $Date") | Out-File $AppPath$LogFile -Append
        ("=====================") | Out-File $AppPath$LogFile -Append
    }
    else{

        mkdir $AppPath
        ("") | Out-File $AppPath$LogFile -Append
        ("=====================") | Out-File $AppPath$LogFile -Append
        ("Launched AD-Tool at $Date") | Out-File $AppPath$LogFile -Append
        ("=====================") | Out-File $AppPath$LogFile -Append
    }
}
else{

    mkdir $RootPath
    mkdir $AppPath
    ("") | Out-File $AppPath$LogFile -Append
    ("=====================") | Out-File $AppPath$LogFile -Append
    ("Launched AD-Tool at $Date") | Out-File $AppPath$LogFile -Append
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

# ========================================================
# End of Active Directory Verification
# ========================================================


# ========================================================
# Start of Code
# ========================================================


# ========================================================
# Show Source Group
# ========================================================

$btnShowFirstGroups.Add_Click({

    BeginningTimestamp("Search First Group.")

    $lstShowFirstGroups.SelectionMode = "One"
    if($txtGetFirstGroupInfo.Text -eq ""){

        ClearVerboseOutput
        $VerboseMessage = ("You can't leave the source box blank.")
        $lstVerboseOutput.Items.Add($VerboseMessage)
    }
    else{
        Try{
            foreach($Item in $lstShowFirstGroups){
            $lstShowFirstGroups.Items.Clear()
            }
        }
        Catch{
            continue
        }
        Try{
            $GroupnameSearch = $txtGetFirstGroupInfo.Text
            $GroupList = Get-ADGroup -Filter "name -like '*$GroupnameSearch*'" | Select-Object SamAccountName

            ForEach($Group in $GroupList){
        
                $Groupname = $Group.SamAccountName
                $lstShowFirstGroups.Items.Add($GroupName)
                
                $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                ("$Timestamp - Group search for < $GroupnameSearch > returned - $Groupname.")| Out-File $AppPath$LogFile -Append
            }
        }Catch{

            ClearVerboseOutput
            $VerboseMessage = ("Couldn't import module for Active Directory.")
            $lstVerboseOutput.Items.Add($VerboseMessage)

            $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
            ("$Timestamp - $VerboseMessage.")| Out-File $AppPath$LogFile -Append
        }
        ("-~-") | Out-File $AppPath$LogFile -Append    
    }
})


# ========================================================
# Show Target Group
# ========================================================

$btnShowSecondGroups.Add_Click({

    BeginningTimestamp("Search Second Group.")

    if($txtGetFirstGroupInfo.Text -eq ""){

        ClearVerboseOutput
        $VerboseMessage = ("You can't leave the detination box blank.")
        $lstVerboseOutput.Items.Add($VerboseMessage)
    }
    else{
        Try{
            foreach($Item in $lstShowSecondGroups){

            $lstShowSecondGroups.Items.Clear()
            }
        }
        Catch{
            continue
        }
        Try{

            $GroupnameSearch2 = $txtGetFirstGroupInfo.Text
            $GroupList2 = Get-ADGroup -Filter "name -like '*$GroupnameSearch2*'" | Select-Object SamAccountName

            ForEach($Group2 in $GroupList2){
        
                $Groupname2 = $Group2.SamAccountName
                $lstShowSecondGroups.Items.Add($GroupName2)
                $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                ("$Timestamp - Group search for < $GroupnameSearch2 > returned - $Groupname2.")| Out-File $AppPath$LogFile -Append
            }
        }Catch{

            ClearVerboseOutput
            $VerboseMessage = ("Couldn't import module for Active Directory.")
            $lstVerboseOutput.Items.Add($VerboseMessage)

            $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
            ("$Timestamp - $VerboseMessage.")| Out-File $AppPath$LogFile -Append
        }
        ("-~-") | Out-File $AppPath$LogFile -Append
    }
})

# ========================================================
# Show Individuals By Name
# ========================================================

$btnSearchByUsername.Add_Click({

    BeginningTimestamp("Search by Username.")

    $lstShowFirstGroups.SelectionMode = "MultiExtended"
    if($txtGetFirstGroupInfo.Text -eq ""){

        ClearVerboseOutput
        $VerboseMessage = ("You can't leave the username box blank.")
        $lstVerboseOutput.Items.Add($VerboseMessage)
    }
    else{
        Try{
            foreach($Item in $lstShowFirstGroups){

            $lstShowFirstGroups.Items.Clear()
            }
        }
        Catch{
            continue
        }
        # Eventually implement a better regex for search results
        Try{
            $UsernameSearch = $txtGetFirstGroupInfo.Text
            $UsernameSearch = $UsernameSearch -replace ".{0}$"
            $UserProfiles = Get-ADUser -Filter "Name -like '*$UsernameSearch*'" | Select-Object name, samaccountname
            
            foreach($User in $UserProfiles){

                $Username = $User.name
                if($Username -notmatch "tmp"){
                    if($Username -notmatch "tmpl"){

                        $lstShowFirstGroups.Items.Add($Username)
                        $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                        ("$Timestamp - User  search for < $UsernameSearch > returned - $Username.")| Out-File $AppPath$LogFile -Append
                    }
                }
            }
        }Catch{

            ClearVerboseOutput
            $VerboseMessage = ("Couldn't import module for Active Directory.")
            $lstVerboseOutput.Items.Add($VerboseMessage)

            $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
            ("$Timestamp - $VerboseMessage.")| Out-File $AppPath$LogFile -Append
        }
        ("-~-") | Out-File $AppPath$LogFile -Append
    }
})

# ========================================================
# Clear boxes
# ========================================================

$btnClearListBoxes.Add_Click({
    
    foreach($Item in $lstShowFirstGroups){

        $lstShowFirstGroups.Items.Clear()
        $lstShowSecondGroups.Items.Clear()
    }

    $txtGetFirstGroupInfo.Text = ""
    $txtGetFirstGroupInfo.Text = ""
    $txtGetFirstGroupInfo.Text = ""
})


# ========================================================
# Find domain Controller where user is locked out
# ========================================================

$btnGetLockoutBlame.Add_Click({

    BeginningTimestamp("Search for Lockout.")

    ClearVerboseOutput
    $VerboseMessage = ("This requires using an admin account.")
    $lstVerboseOutput.Items.Add($VerboseMessage)
    
    Try{
        $DomainController = Get-ADDomainController -Discover -Service PrimaryDC
        $ComputerName = $DomainController.HostName

        # ======================================
        # Not in Use Yet =======================
        #$DC = $ComputerName.ToString()

        Try{
            if($lstShowFirstGroups.SelectedItem -eq $null){
                ClearVerboseOutput
                $VerboseMessage = ("You need to highlight a user to find a lockout. $ComputerName ")
                $lstVerboseOutput.Items.Add($VerboseMessage)    
            }
            else{
                $SelectedUser = $lstShowFirstGroups.SelectedItem
                $UserProfile = Get-ADUser -Filter "Name -eq '$SelectedUser'" | Select-Object name, samaccountname
                $UserProfileSAM = $UserProfile.samaccountname
            
                Try{
                    $Filter = "*[System[EventID=4740] and EventData[Data[@Name='TargetUserName']='$UserProfileSAM']]"
                    $Events = Get-WinEvent -ComputerName $ComputerName -Logname Security -FilterXPath $Filter
                    $LockoutOutput = $Events | Select-Object TimeCreated,
                                                            @{Name='User Name';Expression={$_.Properties[0].Value}},
                                                            @{Name='Source Host';Expression={$_.Properties[1].Value}}
                    ClearVerboseOutput
                    $lstVerboseOutput.Items.Add($LockoutOutput)
                    $Success = ("Task completed.")
                    $lstVerboseOutput.Items.Add($Success)

                    $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                    ("$Timestamp - $LockoutOutput.")| Out-File $AppPath$LogFile -Append
                    ("$Timestamp - $Success.")| Out-File $AppPath$LogFile -Append


                }Catch{
                    ClearVerboseOutput
                    $VerboseMessage = ("Failed to obtain lockout. Please ensure you are running as admin to complete this search.")
                    $lstVerboseOutput.Items.Add($VerboseMessage)
                    $lstVerboseOutput.Items.Add("")
                    $lstVerboseOutput.Items.Add("Error can be found in log file - C:\Temp\AD-Tool\$LogFile")

                    $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                    ("$Timestamp - $VerboseMessage.")| Out-File $AppPath$LogFile -Append
                    $_ | Out-File $AppPath$LogFile -Append
                    ("") | Out-File $AppPath$LogFile -Append
                }
            }
        }Catch{
            ClearVerboseOutput
            $VerboseMessage = ("Could not find user: < $SelectedUser >, please ensure you have highlighted a user and not a group.")
            $lstVerboseOutput.Items.Add($VerboseMessage)

            $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
            ("$Timestamp - $VerboseMessage.")| Out-File $AppPath$LogFile -Append
        }
    }Catch{
        ClearVerboseOutput
        $VerboseMessage = ("Couldn't connect to a domain controller.")
        $lstVerboseOutput.Items.Add($VerboseMessage)

        $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
        ("$Timestamp - $VerboseMessage.")| Out-File $AppPath$LogFile -Append
    }
    ("-~-") | Out-File $AppPath$LogFile -Append
})


# ========================================================
# Unlocking Account function
# ========================================================

$btnUnlockAccount.Add_Click({

    BeginningTimestamp("Unlock a user account.")

    ClearVerboseOutput
    $VerboseMessage = ("This requires using an admin account.")
    $lstVerboseOutput.Items.Add($VerboseMessage)
    

    Try{
        if($lstShowFirstGroups.SelectedItem -eq $null){
            ClearVerboseOutput
            $VerboseMessage = ("You need to highlight a user to unlock their account.")
            $lstVerboseOutput.Items.Add($VerboseMessage)    
        }
        else{

            $SelectedUser = $lstShowFirstGroups.SelectedItem
            $UserProfile = Get-ADUser -Filter "Name -eq '$SelectedUser'" -Properties  name, samaccountname, LockedOut
            $UserProfileSAM = $UserProfile.samaccountname
            $UserProfileLocked = $UserProfile.LockedOut
            
            Try{

                If($UserProfileLocked -eq $false){
                
                    ClearVerboseOutput
                    $VerboseMessage = ("Cannot unlock the selected account because is not currently locked out.")
                    $lstVerboseOutput.Items.Add($VerboseMessage)

                    $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                    ("$Timestamp - The account for < $SelectedUser > was already unlocked.")| Out-File $AppPath$LogFile -Append

                }
                elseif($UserProfileLocked -eq $true){

                    Unlock-ADAccount -Identity $UserProfileSAM

                    $UserProfile = Get-ADUser -Filter "Name -eq '$SelectedUser'" -Properties  name, samaccountname, LockedOut
                    $UserProfileSAM = $UserProfile.samaccountname
                    $UserProfileLocked = $UserProfile.LockedOut

                    If($UserProfileLocked -eq $false){

                        ClearVerboseOutput
                        $VerboseMessage = ("Account has been unlocked.")
                        $lstVerboseOutput.Items.Add($VerboseMessage)

                        $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                        ("$Timestamp - The account for < $SelectedUser > has been unlocked.")| Out-File $AppPath$LogFile -Append
                    }
                    elseif($UserProfileLocked -eq $true){

                        ClearVerboseOutput
                        $VerboseMessage = ("Checking account . .")
                        $lstVerboseOutput.Items.Add($VerboseMessage)
                        $VerboseMessage = ("")
                        $lstVerboseOutput.Items.Add($VerboseMessage)
                        $VerboseMessage = ("Account is locked, either unlocking failed or account has been locked again instantly.")
                        $lstVerboseOutput.Items.Add($VerboseMessage)

                        $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                        ("$Timestamp - The account for < $SelectedUser > failed to be unlocked.")| Out-File $AppPath$LogFile -Append
                    }
                }
            }Catch{
                ClearVerboseOutput
                $VerboseMessage = ("Failed to unlock. Please ensure you are running as admin to perform this function.")
                $lstVerboseOutput.Items.Add($VerboseMessage)
                $lstVerboseOutput.Items.Add("")
                $lstVerboseOutput.Items.Add("Error can be found in log file - C:\Temp\AD-Tool\$LogFile")

                $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                ("$Timestamp - $VerboseMessage.")| Out-File $AppPath$LogFile -Append
                $_ | Out-File $AppPath$LogFile -Append
                ("") | Out-File $AppPath$LogFile -Append
            }
        }
    }Catch{
        ClearVerboseOutput
        $VerboseMessage = ("Could not find user: < $SelectedUser >, please ensure you have highlighted a user and not a group.")
        $lstVerboseOutput.Items.Add($VerboseMessage)

        $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
        ("$Timestamp - $VerboseMessage.")| Out-File $AppPath$LogFile -Append
    }

    ("-~-") | Out-File $AppPath$LogFile -Append
})


# ========================================================
# Show members of Group
# ========================================================

$btnShowFirstUsers.Add_click({

    BeginningTimestamp("Show Users in Group.")

    $lstShowFirstGroups.SelectionMode = "MultiExtended"
    if($lstShowFirstGroups.SelectedItem -eq $null){

        ClearVerboseOutput
        $VerboseMessage = ("You need to search for a group, or highlight a group from the first box.")
        $lstVerboseOutput.Items.Add($VerboseMessage)
    }
    else{
        Try{
            $SelectedGroup = $lstShowFirstGroups.SelectedItem

            foreach($Item in $lstShowFirstGroups){

                $lstShowFirstGroups.Items.Clear()
            }
        }Catch{
            continue
        }   
        Try{
            $Counter = 0
            $UserProfiles = Get-ADGroupMember -Identity $SelectedGroup -Recursive | Get-ADUser -Properties displayname, objectclass, name
            ForEach ($UserProfile in $UserProfiles){
 
                $Username = $UserProfile.name
                
                #=============================================
                # Not in Use yet =============================
                # $ObjectClass = $UserProfile.objectclass
                # $DisplayName = $UserProfile.displayname
                
                foreach($User in $UserProfile){

                    if($Username -notmatch "tmpl"){
                        if($Username -notmatch "tmp"){

                            $lstShowFirstGroups.Items.Add($Username)
                            $counter += 1
                        
                            $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                            ("$Timestamp - Group < $SelectedGroup > contains user - $Username.")| Out-File $AppPath$LogFile -Append
                        }
                    }
                }
        
                ClearVerboseOutput
                $VerboseMessage = ("$Counter users found.")
                $lstVerboseOutput.Items.Add($VerboseMessage)

            }
        }Catch{

            ClearVerboseOutput
            $VerboseMessage = ("Ensure you haven't highlighted a user already.")
            $lstVerboseOutput.Items.Add($VerboseMessage)
        }
    }
    ("-~-") | Out-File $AppPath$LogFile -Append       
})


# ========================================================
# Copies users over to new source group
# ========================================================

$btnCopyUsersOver.Add_Click({

    BeginningTimestamp("Copy All Users in First Group to Second Group.")

    Try{   
        $FirstSelectedGroup = $lstShowFirstGroups.SelectedItem
        $SecondSelectedGroup = $lstShowSecondGroups.SelectedItem

        if($FirstSelectedGroup -eq $null -and $SecondSelectedGroup -eq $null){

            ClearVerboseOutput
            $VerboseMessage = ("You need to search and select a group from each box.")
            $lstVerboseOutput.Items.Add($VerboseMessage)
        }
        elseif($FirstSelectedGroup -eq $null){
            
            ClearVerboseOutput
            $VerboseMessage = ("You haven't selected the first group.")
            $lstVerboseOutput.Items.Add($VerboseMessage)
        }
        elseif($SecondSelectedGroup -eq $null){
            
            ClearVerboseOutput
            $VerboseMessage = ("You haven't selected the second group")
            $lstVerboseOutput.Items.Add($VerboseMessage)
        }
        else{
            Try{

                ClearVerboseOutput
                $Group = Get-ADGroupMember -Identity $FirstSelectedGroup -Recursive | Get-ADUser -Properties displayname, objectclass, name, samaccountname
                $Count = 0

                ForEach ($GroupMember in $Group){

                    $Username = $GroupMember.samaccountname
                    $LifeName = $GroupMember.name

                    Try{
                        if($Lifename -notmatch "tmp"){
                            if($Lifename -notmatch "tmpl"){

                                $VerboseMessage = ("Would be copying $Lifename from $FirstSelectedGroup to $SecondSelectedGroup")
                                $lstVerboseOutput.Items.Add($VerboseMessage)
                                $Count += 1
                            }
                        }
                    }Catch{
                                    
                        $VerboseMessage = ("Could not copy $GroupMember to $SecondSelectedGroup")
                        $lstVerboseOutput.Items.Add($VerboseMessage)
                    }
                }
                # Confirm choices before copying any users.
                if([System.Windows.Forms.MessageBox]::Show("Copy $Count users?", "Confirm",[System.Windows.Forms.MessageBoxButtons]::OKCancel) -eq "OK"){

                    ClearVerboseOutput
                    $Count = 0
                                
                    foreach($GroupMember in $Group){

                        $Username = $GroupMember.samaccountname
                        $LifeName = $GroupMember.name

                        Try{
                            if($Lifename -notmatch "tmp"){
                                if($Lifename -notmatch "tmpl"){

                                    $VerboseMessage = ("Copying $Lifename to $SecondSelectedGroup")
                                    $lstVerboseOutput.Items.Add($VerboseMessage)

                                    Add-ADGroupmember -Identity $SecondSelectedGroup -Members $Username
                                    $Count += 1

                                    $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                                    ("$Timestamp - Copied < $LifeName > from < $FirstSelectedGroup > to - $SecondSelectedGroup.")| Out-File $AppPath$LogFile -Append
                                }
                            }
                        }Catch{   
                                     
                            $VerboseMessage = ("Could not copy $SelectedUser to $SecondSelectedGroup")
                            $lstVerboseOutput.Items.Add($VerboseMessage)

                            $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                            ("$Timestamp - FAILED to copy < $LifeName > from < $FirstSelectedGroup > to - $SecondSelectedGroup. Error:")| Out-File $AppPath$LogFile -Append
                            $_ | Out-File $AppPath$LogFile -Append
                            ("") | Out-File $AppPath$LogFile -Append
                        }
                        $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                        ("$Timestamp - Copied $Count users.")| Out-File $AppPath$LogFile -Append
                    }
                }
                # If the user ends up hitting Cancel
                else{

                    ClearVerboseOutput
                    $VerboseMessage = ("Cancelled. No users copied.")
                    $lstVerboseOutput.Items.Add($VerboseMessage)
                }
            }Catch{

                ClearVerboseOutput
                $VerboseMessage = ("Could not transfer users between groups.")
                $lstVerboseOutput.Items.Add($VerboseMessage)

                $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                ("$Timestamp - FAILED to copy users between groups.")| Out-File $AppPath$LogFile -Append
            }
        }
    }Catch{

        ClearVerboseOutput
        $VerboseMessage = ("Ensure you have selected a group in each box.")
        $lstVerboseOutput.Items.Add($VerboseMessage)
    }
    ("-~-") | Out-File $AppPath$LogFile -Append
})


# ===============================================================
# Show members of group 
# Note: Will throw error if selecting a non-OU very error prone
# ===============================================================

$btnShowMemberships.Add_Click({
    
    BeginningTimestamp("Show Which Groups the User Belongs to.")

    $SelectedUser = $lstShowFirstGroups.SelectedItem

    if($SelectedUser -eq $null){

        ClearVerboseOutput
        $VerboseMessage = ("You need to highlight a user to view their memberships.")
        $lstVerboseOutput.Items.Add($VerboseMessage)
    }
    else{
        Try{

            foreach($Item in $lstShowSecondGroups){

                $lstShowSecondGroups.Items.Clear()
            }
        }Catch{
            continue
        }
        Try{

            $UserProfile = Get-ADUser -Filter "Name -eq '$SelectedUser'" | Select-Object name, samaccountname, userprincipalname
            $UserProfileSAM = $UserProfile.SamAccountName
            $Memberships = Get-ADPrincipalGroupMembership -Identity $UserProfileSAM | Select-Object name | Sort-Object

            Foreach ($Group in $Memberships){

                $lstShowSecondGroups.Items.Add($Group)
                $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                ("$Timestamp - < $UserProfileSAM > belongs to - $Group.")| Out-File $AppPath$LogFile -Append

            }
        }Catch{ 
            # ClearVerboseOutput
            $VerboseMessage = ("Could not display memberships. Ensure you have selected a user and not a group.")
            $lstVerboseOutput.Items.Add($VerboseMessage)

            $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
            ("$Timestamp - Failed to diplayed memberships, please ensure this a username is selected.")| Out-File $AppPath$LogFile -Append
        }
    }
    ("-~-") | Out-File $AppPath$LogFile -Append
})


# ========================================================
# Show User's Information
# ========================================================

$btnUserInfo.Add_Click({

    BeginningTimestamp("Get details about selected user.")
    
    $SelectedUser = $lstShowFirstGroups.SelectedItem
    $UserSAMAccount = Get-ADUser -Filter "Name -eq '$SelectedUser'" | Select-Object -ExpandProperty SamAccountName
    
    # Get Path to User
    $DN = (Get-ADUser -Identity $UserSAMAccount -Properties DistinguishedName).DistinguishedName
    # $OUPath splits DN into 2 parts, grabs second part and replaces OU, DC, and CN references to mimic a folder structure
    $OUPath = (($DN -split ',', 2)[1] -replace '^OU=','') -replace ',OU=', '\'
    $OUPath = $OUPath -replace ',DC=[^,]+', '' -replace '^CN=', ''
    # $UserName Splits indefinitely, grabs first instance and replaces references to "CN=""
    $UserName = ($DN -split ',')[0] -replace '^CN=',''
     # Reassemble the DC units as the FQDN for human readability
    $DN = $DN -replace '^.*?,DC=',''
    $DN = $DN -replace ',DC=', '.'
    $DN = "\\$DN\$OUPath\$UserName"
    

    if($SelectedUser -eq $null){

        ClearVerboseOutput
        $VerboseMessage = ("You need to highlight a user to view their details.")
        $lstVerboseOutput.Items.Add($VerboseMessage)
    }
    else{
        Try{
            ClearVerboseOutput
            $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
            ("$Timestamp - Results for - $SelectedUser")| Out-File $AppPath$LogFile -Append

            $ComputerProperties = Get-ADComputer -Filter "Description -like '*$SelectedUser*'" -Properties *
            $CompProperties = ("DNSHostName", "Description", "IPv4Address")

            $AllProperties = Get-ADUser -Filter "Name -eq '$SelectedUser'" -Properties *
            $Properties = ("DisplayName", "SamAccountName", "Title", "Department", "Description", "LockedOut", "logonCount", "LastLogonDate", "OfficePhone", "whenCreated", "whenChanged")



            Foreach ($Property in $CompProperties){

                Try{

                    $PropertyValue = $ComputerProperties.$Property

                    $PropertyObj = New-Object psobject -Property @{$Property = $PropertyValue}

                    $VerboseInfo = $PropertyObj | Format-List | Out-String

                    $lstVerboseOutput.Items.Add("$VerboseInfo")
                    $lstVerboseOutput.Items.Add("")
                    $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                    ("$Timestamp - $Property - $PropertyValue")| Out-File $AppPath$LogFile -Append
                
                }Catch{

                    $VerboseInfo = ("Couldn't obtain $Property for $SelectedUser")

                    $lstVerboseOutput.Items.Add("$VerboseInfo")
                    $lstVerboseOutput.Items.Add("")
                    $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                    ("$Timestamp - Couldn't obtain < $Property > for - $SelectedUser")| Out-File $AppPath$LogFile -Append
                }
            }

            # Path to User, heavily customized and cannot be moved to enhanced for loop
            $lstVerboseOutput.Items.Add("Active Directory Path: $DN")
            $lstVerboseOutput.Items.Add("")

            Foreach ($Property in $Properties){

                $PropertyValue = $AllProperties.$Property


                $PropertyObj = New-Object psobject -Property @{$Property = $PropertyValue}

                $VerboseInfo = $PropertyObj | Format-List | Out-String

                $lstVerboseOutput.Items.Add("$VerboseInfo")
                $lstVerboseOutput.Items.Add("")
                $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                ("$Timestamp - $Property - $PropertyValue")| Out-File $AppPath$LogFile -Append
            }

            
        }Catch{

            ClearVerboseOutput
            $VerboseMessage = ("Could not display details. Ensure you have selected a user and not a group.")
            $lstVerboseOutput.Items.Add($VerboseMessage)

            $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
            ("$Timestamp - Failed to diplayed details, please ensure this a username is selected.")| Out-File $AppPath$LogFile -Append
        }
    }
    ("-~-") | Out-File $AppPath$LogFile -Append
})


# ========================================================
# Show Others what each button does
# ========================================================

$btnShowHelp.Add_Click({

    ClearVerboseOutput

    $BlankLine = ("------")

    $lstVerboseOutput.Items.Add($BlankLine)
    $lstVerboseOutput.Items.Add("Permissions within this app are based on your user account.")
    $lstVerboseOutput.Items.Add($BlankLine)
    $lstVerboseOutput.Items.Add("Search for the source and destination groups using the search boxes.")
    $lstVerboseOutput.Items.Add($BlankLine)
    $lstVerboseOutput.Items.Add("After highlighting a group from each box, select 'Copy All Users' ")
    $lstVerboseOutput.Items.Add("to copy all the users from the left to the right group.")
    $lstVerboseOutput.Items.Add($BlankLine)
    $lstVerboseOutput.Items.Add("You can also search for users by username or name which will be ")
    $lstVerboseOutput.Items.Add("displayed in the first box.")
    $lstVerboseOutput.Items.Add($BlankLine)
    $lstVerboseOutput.Items.Add("Highlight users from the first box and choose 'Copy Selected Users'")
    $lstVerboseOutput.Items.Add("to copy only these users to the destination group.")
    $lstVerboseOutput.Items.Add($BlankLine)
    $lstVerboseOutput.Items.Add("Highlight a user and select 'Show Users Groups' to fill the ")
    $lstVerboseOutput.Items.Add("second box with all the groups that the user belongs to.")
    $lstVerboseOutput.Items.Add($BlankLine)
    $lstVerboseOutput.Items.Add("Highlight a user and select 'Lockout Location' to ")
    $lstVerboseOutput.Items.Add("find the machine that locked out their account.")
    $lstVerboseOutput.Items.Add($BlankLine)
    $lstVerboseOutput.Items.Add("Highlight a group in the source box and select 'Show Group Members'")
    $lstVerboseOutput.Items.Add("to display the users that are part of that group.")
    $lstVerboseOutput.Items.Add($BlankLine)
    $lstVerboseOutput.Items.Add("Highlight a user and select 'Get User Info' to see")
    $lstVerboseOutput.Items.Add("their machine name, AD info, IP Address etc.")

})

# ========================================================
# Copy Over Selected Users to Another Group
# ========================================================

$btnCopySelectedUsers.Add_Click({

    BeginningTimestamp("Copy Selected Users to Second Group.")

    ClearVerboseOutput
    
    $SelectedUsers = $lstShowFirstGroups.SelectedItems
    $SecondSelectedGroup = $lstShowSecondGroups.SelectedItem

    $Count = 0

    if($SelectedUsers -eq $null){

        ClearVerboseOutput
        $VerboseMessage = ("You need to highlight at least one user to copy them to a group.")
        $lstVerboseOutput.Items.Add($VerboseMessage)
    }
    elseif($SecondSelectedGroup -eq $null){
            
        ClearVerboseOutput
        $VerboseMessage = ("You haven't selected a group to copy the users to.")
        $lstVerboseOutput.Items.Add($VerboseMessage)
    }
    else{
        Try{

            $SelectedUser = $SelectedUsers[0]
            $UserProfile = Get-ADUser -Filter "Name -eq '$SelectedUser'" | Select-Object name, samaccountname
            $username = $UserProfile.samaccountname

            if($Username -eq "" -or $username -eq $null){

                ClearVerboseOutput
                $VerboseMessage = ("Ensure you have selected a user and not a group.")
                $lstVerboseOutput.Items.Add($VerboseMessage)
            }
            else{
                foreach($SelectedUser in $SelectedUsers){

                    $UserProfile = Get-ADUser -Filter "Name -eq '$SelectedUser'" | Select-Object name, samaccountname
                    $Username = $UserProfile.samaccountname
                    $LifeName = $UserProfile.name

                    Try{

                        if($Lifename -notmatch "tmp"){
                            if($Lifename -notmatch "tmpl"){

                                $VerboseMessage = ("Would be copying $Lifename to $SecondSelectedGroup")
                                $lstVerboseOutput.Items.Add($VerboseMessage)
                                $Count += 1
                            }
                        }
                    }Catch{                
                        $VerboseMessage = ("Could not copy $SelectedUser to $SecondSelectedGroup")
                        $lstVerboseOutput.Items.Add($VerboseMessage)
                    }
                }
                #Ask the user to confirm before copying any users
                if([System.Windows.Forms.MessageBox]::Show("Copy $Count users?", "Confirm",[System.Windows.Forms.MessageBoxButtons]::OKCancel) -eq "OK"){

                    ClearVerboseOutput
                    $Count = 0
                                
                    foreach($SelectedUser in $SelectedUsers){

                        $UserProfile = Get-ADUser -Filter "Name -eq '$SelectedUser'" | Select-Object name, samaccountname
                        $Username = $UserProfile.samaccountname
                        $LifeName = $UserProfile.name

                        Try{
                            if($Lifename -notmatch "tmp"){
                                if($Lifename -notmatch "tmpl"){

                                    $VerboseMessage = ("Copying $Lifename to $SecondSelectedGroup")
                                    $lstVerboseOutput.Items.Add($VerboseMessage)
                                    Add-ADGroupmember -Identity $SecondSelectedGroup -Members $Username
                                    $Count += 1

                                    $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                                    ("$Timestamp - Copied < $LifeName > to - $SecondSelectedGroup.")| Out-File $AppPath$LogFile -Append
                                }
                            }
                        }Catch{
                            
                            $VerboseMessage = ("Could not copy $SelectedUser to $SecondSelectedGroup")
                            $lstVerboseOutput.Items.Add($VerboseMessage)

                            $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
                            ("$Timestamp - FAILED to copy < $LifeName > to - $SecondSelectedGroup. Error:")| Out-File $AppPath$LogFile -Append
                            $_ | Out-File $AppPath$LogFile -Append
                            ("") | Out-File $AppPath$LogFile -Append
                        }
                    }
                }
                else{

                    ClearVerboseOutput
                    $VerboseMessage = ("Cancelled.")
                    $lstVerboseOutput.Items.Add($VerboseMessage)
                }
            }
        }Catch{

            ClearVerboseOutput
            $VerboseMessage = ("Ensure you have selected a user and not a group.")
            $lstVerboseOutput.Items.Add($VerboseMessage)
        }
    }
    ("-~-") | Out-File $AppPath$LogFile -Append   
})

# ========================================================
# Close Program Correctly
# ========================================================

$btnExit.Add_Click({

    ("-~-") | Out-File $AppPath$LogFile -Append
    BeginningTimestamp("Disconnected from Active Directory as < $CurrentUser > - Program Closed.")
    
    $frmInitialScreen.Add_FormClosing({
        $_.Cancel=$false
    })

    $frmInitialScreen.Close()
})

# ========================================================
# Text Cleaning Function
# ========================================================

Function ClearVerboseOutput(){

    Try{
        foreach($Item in $lstShowFirstGroups){

            $lstVerboseOutput.Items.Clear()
        }
    }Catch{
        continue
    }
}

# ========================================================
# Beginning Timestamp Function
# ========================================================

Function BeginningTimestamp([String] $functionLabel){

    $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
    ("") | Out-File $AppPath$LogFile -Append
    ("") | Out-File $AppPath$LogFile -Append
    ("$Timestamp - FUNCTION - $functionLabel") | Out-File $AppPath$LogFile -Append
    ("/ End of Session /") | Out-File $AppPath$LogFile -Append
}

# ========================================================
# Code End
# ========================================================

# Disable use of the X/ Close button  
$frmInitialScreen.add_FormClosing({
    $_.Cancel = $true
})
#Launch form
[void] $frmInitialScreen.ShowDialog()