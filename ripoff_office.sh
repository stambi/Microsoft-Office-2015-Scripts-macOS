#!/bin/bash
######
#08.05.2017 / STA
#RipOff Microsoft Office 356 - leave Outlook & Outlook db's installed
######

#Set Variables
loggedinuser=$(/usr/bin/python -c 'from SystemConfiguration import SCDynamicStoreCopyConsoleUser; import sys; username = (SCDynamicStoreCopyConsoleUser(None, None, None) or [None])[0]; username = [username,""][username in [u"loginwindow", None, u""]]; sys.stdout.write(username + "\n");')
logfile=/tmp/officeripoff.log


#Close Applications
# Friendly close Office applications if open
osascript -e 'try' -e 'quit app "Microsoft Word"' -e 'end try'
osascript -e 'try' -e 'quit app "Microsoft Excel"' -e 'end try'
osascript -e 'try' -e 'quit app "Microsoft PowerPoint"' -e 'end try'
osascript -e 'try' -e 'quit app "Microsoft OneNote"' -e 'end try'

sleep 10

# Force quit if there are still running Office applications.
killall "Microsoft Word"
killall "Microsoft Excel"
killall "Microsoft PowerPoint"
killall "Microsoft OneNote"

sleep 5


#Delete in Appliactions Folder
echo "###### Delete Apps ######" >> $logfile
for app in "Excel" "Word" "Powerpoint" "OneNote"
do
  if [ -d "/Applications/Microsoft $app.app" ]; then
    rm -rf "/Applications/Microsoft $app.app"
    echo "$(date): Delete App: Microsoft $app" >> $logfile
  fi
done


#Delete User Containers
echo "###### Delete User Containers ######" >> $logfile
for container in "Excel" "Word" "Powerpoint" "oneneote.mac" "errorreporting" "netlib.shipassertprocess" "Office365ServiceV2" "RMS-XPCService"
do
  if [ -d "/Users/$loggedinuser/Library/Containers/com.microsoft.$container" ]; then
    rm -rf "/Users/$loggedinuser/Library/Containers/com.microsoft.$container"
    echo "$(date): Delete User Containers: /Users/$loggedinuser/Library/Containers/com.microsoft.$container" >> $logfile
  fi
done


#Delete Preferences
echo "###### Delete Preferences ######" >> $logfile
for preference in "Excel" "onenote.mac" "Powerpoint" "Word"
do
  if [ -e "/Library/Preferences/com.microsoft.$preference.plist" ]; then
    rm -f "/Library/Preferences/com.microsoft.$preference.plist"
    echo "$(date): Delete Preferences: /Library/Preferences/com.microsoft.$preference.plist" >> $logfile
  fi
done


#Forget Package Receipts
echo "###### Forget Package Receipts ######" >> $logfile
for pkg in "Excel" "OneNote" "PowerPoint" "Word"
do
  pkgname=$(pkgutil --pkgs | grep $pkg)
  if [ -n "$pkgname" ]; then
    pkgutil --forget com.microsoft.package.Microsoft_$pkg.app
    echo "$(date): Forget Package Receipts: com.microsoft.package.Microsoft_$pkg.app" >> $logfile
  fi
done


#Remove items from dock
echo "###### Forget Package Receipts ######" >> $logfile
for dockitem in "Microsoft%20Word.app" "Microsoft%20OneNote.app" "Microsoft%20PowerPoint.app" "Microsoft%20Excel.app"
do
dloc=$(defaults read "/Users/$loggedinuser/Library/Preferences/com.apple.dock.plist" persistent-apps | grep _CFURLString\" | awk -v x="$dockitem" '$0~x {print NR-1}')
  if [ -n "$dloc" ]; then
    /usr/libexec/PlistBuddy -c "Delete persistent-apps:$dloc" "/Users/$loggedinuser/Library/Preferences/com.apple.dock.plist"
    echo "$(date): $dloc removed from dock" >> $logfile
  fi
done

killall Dock
echo "$(date): Dock: Items removed from Dock" >> $logfile

echo "###### Office 2016 removed ######" >> $logfile
exit 0
