#*************************************************************** Constant variables *****************************************************************************
$burntToastUNCPath = "Accessible UNC path for burntoast"            # UNC path to the BurntToast module on the NAS                                        
$burnToastLocalPath = "$env:LOCALAPPDATA\BurntToast"                # Local path for the BurntToast module in %localappdata%
$NmbrJours = 10                                                     # Number of days before notifying the user
$HeroLogo = "The herologo path you want to use"                     # Path to the Semat hero image
$PauseSecondes = 1                                                  # Wait time before sending the notification  


$ErrorActionPreference = 'SilentlyContinue' # Force execution without prompt and avoid visible errors

#*************************************************************** Import Module ************************************************************************************
if (-not (Test-Path $burnToastLocalPath)) { # If the module is not already present locally
    if (Test-Path $burntToastUNCPath) {
        New-Item -ItemType Directory -Path $burnToastLocalPath -Force | Out-Null # Create the local folder if needed       
        Copy-Item -Path "$burntToastUNCPath\*" -Destination $burnToastLocalPath -Recurse -Force # Copy the contents of the network folder to the local folder
    }
}

if (Test-Path $burnToastLocalPath) { # Import the module from the local folder
    Import-Module $burnToastLocalPath -Force -ErrorAction SilentlyContinue
}

#*************************************************************** PC uptime verification ***************************************************************
$Days = $NmbrJours #Set the number of days before notifying the user

# Check how long the computer has not been shut down.
$cim = Get-CimInstance win32_operatingsystem
$uptime = (Get-Date) - ($cim.LastBootUpTime)
$uptimeDays = $Uptime.Days


# Return Exit code 0 if this computer has not been online for too long, ending the script.
if ($uptimeDays -LT $Days) { # For testing I am using "-GT" greater than to get the notification if it has been on for less than $days days / For production use "-LT" less than
Exit 0
}

#***************************************************************** Get current user ***************************************************************
$CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$SamAccountName = $CurrentUser.Split('\')[1]

#***************************************************************** LDAP searcher to get user name *************************************************
$Searcher = New-Object DirectoryServices.DirectorySearcher
$Searcher.Filter = "(&(objectCategory=person)(objectClass=user)(sAMAccountName=$SamAccountName))"
$Searcher.PropertiesToLoad.Add("givenName") | Out-Null

$Result = $Searcher.FindOne()
$GivenName = $Result.Properties["givenname"]

Start-Sleep -Seconds $PauseSecondes

#*************************************************************** Notifications *****************************************************************

#********************************************************** Notification construction ******************************************************
$Text1 = New-BTText -Content "Hello $GivenName, your computer has been up for $uptimeDays days"
$Text2 = New-BTText -Content "Can you please restart it as soon as possible ?" 

#****************************************************** Buttons and actions ***********************************************************************
$Button = New-BTButton -Content "Report" -snooze -id 'SnoozeTime'                                   # Button to snooze the notification
$Button2 = New-BTButton -Content "Reboot" -Arguments "ToastReboot:" -ActivationType Protocol        # Button to restart the computer
$5Min = New-BTSelectionBoxItem -Id 5 -Content '5 minutes'                                           # Snooze for 5 minutes
$10Min = New-BTSelectionBoxItem -Id 10 -Content '10 minutes'                                        # Snooze for 10 minutes
$1Hour = New-BTSelectionBoxItem -Id 60 -Content '1 heure'                                           # Snooze for 1 hour
$4Hour = New-BTSelectionBoxItem -Id 240 -Content '4 heures'                                         # Snooze for 4 hours
$1Day = New-BTSelectionBoxItem -Id 1440 -Content '1 jour'                                           # Snooze for 1 day
$Items = $5Min, $10Min, $1Hour, $4Hour, $1Day
$SelectionBox = New-BTInput -Id 'SnoozeTime' -DefaultSelectionBoxItemId 10 -Items $Items            # Default snooze value set to 10 minutes
$heroimage = New-BTImage -Source $HeroLogo -HeroImage                                               # Create a hero image for the notification

$action = New-BTAction -Buttons $Button, $Button2 -inputs $SelectionBox                             # Notification actions, turns "SelectionBoxItem" into buttons

$Binding = New-BTBinding -Children $Text1, $Text2 -HeroImage $heroimage                             # Bind texts and icon to the notification

$Visual = New-BTVisual -BindingGeneric $Binding                                                     # Create the notification visual with texts and icon

$Audio = New-BTAudio -Silent                                                                        # Set audio to silent to allow choosing the notification duration. Otherwise, the default sound is played, so notification duration is short "Short".

$Content = New-BTContent -Visual $Visual -Actions $action -Audio $Audio -Duration Long              # Create the notification content with visual, actions, and audio, and set the notification duration to "Long" (about 25 seconds). Can be replaced by "Short" for a short duration (about 5 seconds).
    
#***************************************************** Send the notification ***********************************************************************

Submit-BTNotification -Content $Content
