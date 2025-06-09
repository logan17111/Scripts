#*************************************************************** Modifiable variables *************************************************************************
$burntToastUNCPath = "Accessible UNC path for burntoast"        # UNC path to the BurntToast module on the NAS
$burnToastLocalPath = "$env:LOCALAPPDATA\BurntToast"            # Local path for the BurntToast module in %localappdata%
$WaitingSec = 200                                               # Waiting time before sending the notification.
$HeroLogo = "The herologo path you want to use"                 # Path to the hero image for the notification

#*************************************************************** Start of script *******************************************************************************

$ErrorActionPreference = 'SilentlyContinue' # Force execution without prompt and avoid visible errors

if (-not (Test-Path $burnToastLocalPath)) { # If the module is not already present locally
    if (Test-Path $burntToastUNCPath) {
        New-Item -ItemType Directory -Path $burnToastLocalPath -Force | Out-Null # Create the local folder if needed
        Copy-Item -Path "$burntToastUNCPath\*" -Destination $burnToastLocalPath -Recurse -Force # Copy the contents of the network folder to the local folder
    }
}
# Import the module from the local folder
if (Test-Path $burnToastLocalPath) {
    Import-Module $burnToastLocalPath -Force -ErrorAction SilentlyContinue
}

#***************************************************************** Get current user ****************************************************************************
$CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$SamAccountName = $CurrentUser.Split('\')[1]

#***************************************************************** LDAP searcher to get information ************************************************************
$Searcher = New-Object DirectoryServices.DirectorySearcher
$Searcher.Filter = "(&(objectCategory=person)(objectClass=user)(sAMAccountName=$SamAccountName))"
$Searcher.PropertiesToLoad.Add("msDS-UserPasswordExpiryTimeComputed") | Out-Null
$Searcher.PropertiesToLoad.Add("givenName") | Out-Null

$Result = $Searcher.FindOne()
$GivenName = $Result.Properties["givenname"]

Start-Sleep -Seconds $WaitingSec

#****************************************************************** Start of notification **********************************************************************

        #********************************************************** Notification construction *******************************************************************
            # Text
            $Text10 = New-BTText -Content "Hello $GivenName"                              # Notification title
            $Text20 = New-BTText -Content "blablabla"                                     # Notification content

            #****************************************************** Buttons and actions *************************************************************************
            $Button10 = New-BTButton -Content "Access to ..." -Arguments "https://ahahahah.com" -ActivationType Protocol    # Button to access the URL
            $Button20 = New-BTButton -Content "Ignore" -Dismiss                                                             # Button to ignore the notification
            $action = New-BTAction -Buttons $Button10, $Button20                    

            #****************************************************** AppLogo: create the object with crop "circle" ***********************************************
            $heroimage = New-BTImage -Source $HeroLogo -HeroImage                  # Create a hero image for the notification

            $Binding61 = New-BTBinding -Children $Text10, $Text20 -HeroImage $heroimage # Bind the texts and icon to the notification

            #***************************************************** Create the visual ****************************************************************************
            $Visual = New-BTVisual -BindingGeneric $Binding61 # Create the visual for the notification with texts and icon

            #***************************************************** Notification audio ***************************************************************************
            $Audio1 = New-BTAudio -Silent # Set the audio to silent to be able to choose the notification duration. Otherwise, the default sound is played so notification duration is short "Short".

            #***************************************************** Create the content and submit the notification ***********************************************
            $Content = New-BTContent -Visual $Visual -Actions $action -Audio $Audio1 -Duration Short # Create the notification content with visual, actions, and audio, and set the notification duration to "Long" (about 25 seconds). Replaceable by "Short" for a short duration (about 5 seconds).
            
            #***************************************************** Send the notification ************************************************************************
            Submit-BTNotification -Content $Content 
