# ============================================================================================================
# Logging Format. Ensures path exists or it will make it exist.
# ============================================================================================================

$TimestampFormat = "HH:mm:ss"
$Timestamp = [DateTime]::Now.ToString($TimestampFormat)

$LogDateFormat = "dd-MM-yyyy"
$LogDate = [DateTime]::Today.ToString($LogDateFormat)

$RootPath = "C:\Users\administrator.REDCLIFF\Desktop\"
$AppPath = "$RootPath\UAC-Log Files\"
$LogFile = "log-$LogDate.txt"

$RootCheck = Test-Path $RootPath
$AppCheck = Test-Path $AppPath

if($RootCheck -eq $true){
    if($AppCheck -eq $true){
        # doesn't need to do anything
    }
    else{
        mkdir $AppPath
    }
}
else{
    mkdir $RootPath
    mkdir $AppPath
}

# ==============================================================================================================
# Check status
# ==============================================================================================================
$UACStatus = [bool](Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System).EnableLUA


# Checks if value is equal to true 
if ($UACStatus -eq $true){
    leaveNote
    Set-ItemProperty -Path REGISTRY::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System -Name ConsentPromptBehaviorAdmin -Value 0
}

# Timestamp and output
Function LeaveNote(){
    $Timestamp = [DateTime]::Now.ToString($TimestampFormat)
    ("UAC Turned back on at $Timestamp") | Out-File $AppPath$LogFile -Append
}