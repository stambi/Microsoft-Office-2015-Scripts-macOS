#!/bin/sh
################################################################################
# Trust MAU - Needs to run (re register) after every MAU Update! Run periodically 
#
# Discussion on JAMF Nation: https://www.jamf.com/jamf-nation/discussions/22853/best-practices-for-automating-office-2016-updates
#
# 15.02.2017 / STA
#
################################################################################

 loggedInUser=$(/usr/bin/python -c 'from SystemConfiguration import SCDynamicStoreCopyConsoleUser; import sys; username = (SCDynamicStoreCopyConsoleUser(None, None, None) or [None])[0]; username = [username,""][username in [u"loginwindow", None, u""]]; sys.stdout.write(username + "\n");')

su - "$loggedInUser" -c '/System/Library/Frameworks/CoreServices.framework/Frameworks/LaunchServices.framework/Support/lsregister -R -f -trusted "/Library/Application Support/Microsoft/MAU2.0/Microsoft AutoUpdate.app"'
su - "$loggedInUser" -c '/System/Library/Frameworks/CoreServices.framework/Frameworks/LaunchServices.framework/Support/lsregister -R -f -trusted "/Library/Application Support/Microsoft/MAU2.0/Microsoft AutoUpdate.app/Contents/MacOS/Microsoft AU Daemon.app"'
su - "$loggedInUser" -c 'defaults delete com.microsoft.autoupdate2 LastUpdate'

exit 0
