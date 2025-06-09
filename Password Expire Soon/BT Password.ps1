#*************************************************************** Constant variables *************************************************************************
$burntToastUNCPath = "Accessible UNC path for burntoast"        # UNC path to the BurntToast module on the NAS
$burnToastLocalPath = "$env:LOCALAPPDATA\BurntToast"            # Local path for the BurntToast module in %localappdata%
#$LogoImage = "logo path you want to use"                       # Path to the image on the local client / Uncomment this line if you want to use a logo image
$NbreJoursExpiration = 400                                      # Password expiration duration in days; if the password expires in less than x days, the notification will be displayed.
$WaitingSec = 1                                                 # Waiting time before sending the notification.
$HeroLogo = "The herologo path you want to use"                # Path to the hero image for the notification


#*************************************************************** Script start *******************************************************************************

$ErrorActionPreference = 'SilentlyContinue' # Force execution without prompt and avoid visible errors


if (-not (Test-Path $burnToastLocalPath)) { # If the module is not already present locally
    if (Test-Path $burntToastUNCPath) {
        New-Item -ItemType Directory -Path $burnToastLocalPath -Force | Out-Null # Create the local folder if needed
        Copy-Item -Path "$burntToastUNCPath\*" -Destination $burnToastLocalPath -Recurse -Force # Copy the content from the network folder to the local folder
    }
}
# Import the module from the local folder
if (Test-Path $burnToastLocalPath) {
    Import-Module $burnToastLocalPath -Force -ErrorAction SilentlyContinue
}

#***************************************************************** Get current user ***************************************************************
$CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$SamAccountName = $CurrentUser.Split('\')[1]

#***************************************************************** LDAP searcher to get information *************************************************
$Searcher = New-Object DirectoryServices.DirectorySearcher
$Searcher.Filter = "(&(objectCategory=person)(objectClass=user)(sAMAccountName=$SamAccountName))"
$Searcher.PropertiesToLoad.Add("msDS-UserPasswordExpiryTimeComputed") | Out-Null
$Searcher.PropertiesToLoad.Add("givenName") | Out-Null

$Result = $Searcher.FindOne()

Start-Sleep -Seconds $WaitingSec

#****************************************************************** Notification start ********************************************************************
if ($Result) {
    $PasswordExpirationTime = $Result.Properties["msds-userpasswordexpirytimecomputed"]
    $GivenName = $Result.Properties["givenname"]

    #************************************************************** Check if the password expiration date is near **********************************
    if ($PasswordExpirationTime.Count -gt 0) {
        $ExpirationDate = [datetime]::FromFileTime($PasswordExpirationTime[0])
        $DateToday = Get-Date
        $DateThreshold = $NbreJoursExpiration # Password expiration duration in days; if the password expires in less than x days, the notification will be displayed.
        $DateWithThreshold = $DateToday.AddDays($DateThreshold)

        #********************************************************** Notification construction *************************************************************
        if (($ExpirationDate -lt $DateWithThreshold) -and ($ExpirationDate -gt $DateToday)) {
            # Text
            $Text10 = New-BTText -Content "Password Expiration"                                                                       # Notification title
            $Text20 = New-BTText -Content "Hello $GivenName, your password expires on $($ExpirationDate.ToString('dd/MM/yyyy'))."     # Notification content

            #****************************************************** Buttons and actions **************************************************************************
            $Button10 = New-BTButton -Content "Snooze" -snooze -id 'SnoozeTime'   # Button to snooze the notification
            $Button20 = New-BTButton -Content "Dismiss" -Dismiss                  # Button to dismiss the notification
            $5Min = New-BTSelectionBoxItem -Id 5 -Content '5 minutes'             # Snooze for 5 minutes
            $10Min = New-BTSelectionBoxItem -Id 10 -Content '10 minutes'          # Snooze for 10 minutes
            $1Hour = New-BTSelectionBoxItem -Id 60 -Content '1 hour'              # Snooze for 1 hour
            $4Hour = New-BTSelectionBoxItem -Id 240 -Content '4 hours'            # Snooze for 4 hours
            $1Day = New-BTSelectionBoxItem -Id 1440 -Content '1 day'              # Snooze for 1 day
            $Items = $5Min, $10Min, $1Hour, $4Hour, $1Day
            $SelectionBox = New-BTInput -Id 'SnoozeTime' -DefaultSelectionBoxItemId 10 -Items $Items    # Default snooze value set to 10 minutes
            $action = New-BTAction -Buttons $Button10, $Button20 -inputs $SelectionBox                  # Notification actions, turns "SelectionBoxItem" into buttons

            #****************************************************** AppLogo: create the object with crop "circle" **************************************************
            $heroimage = New-BTImage -Source $HeroLogo -HeroImage                          # Create a hero image for the notification
            # Uncomment the next line if you want to use a logo image
            <#$AppLogo71 = [Microsoft.Toolkit.Uwp.Notifications.ToastGenericAppLogo]@{      # Create an AppLogo object for the notification
                Source = $LogoImage                                                         # Icon path
                HintCrop = 'circle'                                                         # Crop the icon as a circle
            }#>

            $Binding61 = New-BTBinding -Children $Text10, $Text20 <#-AppLogoOverride $AppLogo71#> -HeroImage $heroimage # Bind texts and icon to the notification

            #***************************************************** Create the visual *******************************************************************************
            $Visual = New-BTVisual -BindingGeneric $Binding61 # Create the visual for the notification with texts and icon

            #***************************************************** Notification audio ******************************************************************************
            $Audio1 = New-BTAudio -Silent # Set audio to silent to allow choosing the notification duration. Otherwise, the default sound is played, so notification duration is short "Short".

            #***************************************************** Create the content and submit the notification **************************************************
            $Content = New-BTContent -Visual $Visual -Actions $action -Audio $Audio1 -Duration Long # Create the notification content with visual, actions, and audio, and set the notification duration to "Long" (about 25 seconds). Can be replaced by "Short" for a short duration (about 5 seconds).
            
            #***************************************************** Send the notification **************************************************************************
            Submit-BTNotification -Content $Content 
        }
    }
}