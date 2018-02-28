#!/bin/sh
################################################################################
# Enable Automatic Updates & Register Apps in MAU
# Will only set the defaults in the plist file. The user can overwrite those settings
# Will Register the Apps in MAU (otherwise the Apps won't be updated until the are startet once.
#
# Discussion on JAMF Nation: https://www.jamf.com/jamf-nation/discussions/22853/best-practices-for-automating-office-2016-updates
#
# 15.02.2017 / STA
#
################################################################################

loggedInUser=$(/usr/bin/python -c 'from SystemConfiguration import SCDynamicStoreCopyConsoleUser; import sys; username = (SCDynamicStoreCopyConsoleUser(None, None, None) or [None])[0]; username = [username,""][username in [u"loginwindow", None, u""]]; sys.stdout.write(username + "\n");')


## Delete existing prefs in Home folder
#if [ -f /Users/$loggedInUser/Library/Preferences/com.microsoft.autoupdate2.plist ];then
#rm /Users/$loggedInUser/Library/Preferences/com.microsoft.autoupdate2.plist
#else echo "no plist in ~/Library/Preferences/"
#fi
#sleep 2
## Delete existing prefs in / f
#if [ -f /Library/Preferences/com.microsoft.autoupdate2.plist ];then
#rm /Library/Preferences/com.microsoft.autoupdate2.plist
#else echo "no plist in /Library/Preferences/"
#fi

# enable AutomaticDownload 
enableUpdates() {
/usr/bin/defaults write /Library/Preferences/com.microsoft.autoupdate2 HowToCheck AutomaticDownload
/usr/bin/defaults write /Users/$loggedInUser/Library/Preferences/com.microsoft.autoupdate2 HowToCheck AutomaticDownload

## Shorter checkin interval (2min) for testing. Default is 12h
#/usr/bin/defaults write /Library/Preferences/com.microsoft.autoupdate2 UpdateCheckFrequency -int 2
#/usr/bin/defaults write /Users/$loggedInUser/Library/Preferences/com.microsoft.autoupdate2 UpdateCheckFrequency -int 2
}

# Register Office Apps in MAU
# Add office apps to /Library/Preferences/com.microsoft.autoupdate2 plist
trustMAU() {
applications="
Word.app
Excel.app
PowerPoint.app
OneNote.app
Outlook.app
Silverlight.app"

for application in $applications
do
     domain="/Library/Preferences/com.microsoft.autoupdate2"
     defaults_cmd="/usr/bin/defaults"
     application_info_plist="/Applications/Microsoft $application/Contents/Info.plist"
     lcid="1033"

     if /bin/test -f "$application_info_plist"
     then
          application_bundle_signature=$($defaults_cmd read "$application_info_plist" CFBundleSignature)
          application_bundle_version=$($defaults_cmd read "$application_info_plist" CFBundleVersion)
          application_id=$(printf "%s%02s" "$application_bundle_signature" "${application_bundle_version%%.*}")
          $defaults_cmd write $domain Applications -dict-add "/Applications/Microsoft $application" "{ 'Application ID' = $application_id; LCID = $lcid ; }"
     fi
done
}

# Register Office Apps in MAU
# Add office apps to ~/Library/Preferences/com.microsoft.autoupdate2 plist (User!)
trustMAU_user() {
applications="
Word.app
Excel.app
PowerPoint.app
OneNote.app
Outlook.app
Silverlight.app"

for application in $applications
do
     domain="/Users/$loggedInUser/Library/Preferences/com.microsoft.autoupdate2"
     defaults_cmd="/usr/bin/defaults"
     application_info_plist="/Applications/Microsoft $application/Contents/Info.plist"
     lcid="1033"

     if /bin/test -f "$application_info_plist"
     then
          application_bundle_signature=$($defaults_cmd read "$application_info_plist" CFBundleSignature)
          application_bundle_version=$($defaults_cmd read "$application_info_plist" CFBundleVersion)
          application_id=$(printf "%s%02s" "$application_bundle_signature" "${application_bundle_version%%.*}")
          $defaults_cmd write $domain Applications -dict-add "/Applications/Microsoft $application" "{ 'Application ID' = $application_id; LCID = $lcid ; }"
     fi
done
}

#add Skype to /Library/Preferences/com.microsoft.autoupdate2 plist as well
trustMAU_Skype() {
domain="/Library/Preferences/com.microsoft.autoupdate2"
defaults_cmd="/usr/bin/defaults"
application_info_plist="/Applications/Skype for Business.app/Contents/Info.plist"
lcid="1033"

if /bin/test -f "$application_info_plist"
then
     application_bundle_signature=$($defaults_cmd read "$application_info_plist" CFBundleSignature)
     application_bundle_version=$($defaults_cmd read "$application_info_plist" CFBundleVersion)
     application_id=$(printf "%s%02s" "$application_bundle_signature" "${application_bundle_version%%.*}")
     $defaults_cmd write $domain Applications -dict-add "/Applications/Skype for Business.app" "{ 'Application ID' = $application_id; LCID = $lcid ; }"
fi
}

#add Skype to ~/Library/Preferences/com.microsoft.autoupdate2 plist as well (User!)
trustMAU_Skype_user() {
domain="/Users/$loggedInUser/Library/Preferences/com.microsoft.autoupdate2"
defaults_cmd="/usr/bin/defaults"
application_info_plist="/Applications/Skype for Business.app/Contents/Info.plist"
lcid="1033"

if /bin/test -f "$application_info_plist"
then
     application_bundle_signature=$($defaults_cmd read "$application_info_plist" CFBundleSignature)
     application_bundle_version=$($defaults_cmd read "$application_info_plist" CFBundleVersion)
     application_id=$(printf "%s%02s" "$application_bundle_signature" "${application_bundle_version%%.*}")
     $defaults_cmd write $domain Applications -dict-add "/Applications/Skype for Business.app" "{ 'Application ID' = $application_id; LCID = $lcid ; }"
fi
}

# Create a LaunchAgent 
createLaunchAgent() {
cat <<EOF > /Library/LaunchAgents/com.microsoft.update.agent.plist
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
	<key>Disabled</key>
	<false/>
	<key>Label</key>
	<string>com.microsoft.update.agent</string>
	<key>ProgramArguments</key>
	<array>
		<string>/Library/Application Support/Microsoft/MAU2.0/Microsoft AutoUpdate.app/Contents/MacOS/Microsoft AU Daemon.app/Contents/MacOS/Microsoft AU Daemon</string>
		<string>-checkForUpdates</string>
	</array>
	<key>RunAtLoad</key>
	<true/>
	<key>StartInterval</key>
	<integer>43200</integer>
</dict>
</plist>
EOF
}

# Execute the Functions
trustMAU
trustMAU_user
trustMAU_Skype
trustMAU_Skype_user
enableUpdates
createLaunchAgent

# Set rights
chmod +r /Users/$loggedInUser/Library/Preferences/com.microsoft.autoupdate2.plist

#kill cfprefsd to load new plist (will restart automatically) 
killall cfprefsd

exit 0
