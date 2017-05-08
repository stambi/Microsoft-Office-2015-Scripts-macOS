#!/bin/bash
######
#08.05.2017 / STA
#RipOff Microsoft Office 356
######

#Get loggedin Users
loggedinuser=$(ls -l /dev/console | awk '{ print $3 }')
logfile=/tmp/officeripoff.log

#Delete in Appliactions Folder
for app in "Excel" "Word" "Powerpoint" "OneNote"
do
  if [ -d "/Applications/Microsoft $app.app" ]; then
#    rm -rf "/Applications/Microsoft $app.app"
    echo "$(date): Delete App: Microsoft $app" >> $logfile
  fi
done

#Delete User Containers
for container in "Excel" "Word" "Powerpoint" "oneneote.mac" "errorreporting" "netlib.shipassertprocess" "Office365ServiceV2" "RMS-XPCService"
do
  if [ -d "/Users/$loggedinuser/Library/Containers/com.microsoft.$container" ]; then
#    rm -rf "/Users/$loggedinuser/Library/Container/microsoft.$container"
    echo "$(date): Delete User Containers: /Users/$loggedinuser/Library/Containers/com.microsoft.$container" >> $logfile
  fi
done
