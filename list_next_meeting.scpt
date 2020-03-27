#!/usr/bin/osascript
 
-- This script will output a terse string
-- indicating when your next meeting is occuring, and 
-- also where. This could be useful for many reasons.
-- I use this inside a tmux info pane so that I can focus
-- on a terminal window without missing a meeting alert.

on replace(A, B, theText)
     set {TID, AppleScript'stext item delimiters} to {AppleScript'stext item delimiters, {A}}
     set {theTextItems, AppleScript'stext item delimiters} to {text items of theText, {B}}
     set {theText, AppleScript'stext item delimiters} to {theTextItems as text, TID}
     return theText
end replace

set Cals2Check to "Calendar"
set curdate to current date
set outsidedate to (curdate + 43200) --The number at the end determines how many seconds to look into the future for a meeting

set delims to AppleScript's text item delimiters
if Cals2Check contains ", " then
 set AppleScript's text item delimiters to {", "}
else
 set AppleScript's text item delimiters to {","}
end if
set caltitles to every text item of Cals2Check
set AppleScript's text item delimiters to delims

tell application "Microsoft Outlook"

--We need to get the ID of each calendar, as the names are not always unique (this may be an issue with mounted shared calendars)
 set calIDs to {}
 repeat with i from 1 to number of items in caltitles
  set caltitle to item i of caltitles
  set calIDs to calIDs & (id of every calendar whose name is caltitle)
 end repeat


--Now we get a list of events from each of the calendar that match our time criteria
 set calEvents to {}
 repeat with i from 1 to number of items in calIDs
  set CalID to item i of calIDs

  tell (calendar id CalID)
   set calEvents to calEvents & (every calendar event whose (start time > (curdate - 300)) and (start time < (outsidedate)))
  end tell

 end repeat


 --we grab the "next" calendar event
 set nextEventTitle to {}
 repeat with i from 1 to number of items in calEvents
  if nextEventTitle is {} then
   set nextEventTitle to item i of calEvents
  else
   if start time of item i of calEvents is less than start time of item 1 of nextEventTitle then
    set nextEventTitle to item i of calEvents
   end if
  end if

 end repeat


 if nextEventTitle is not {} then

  set MeetingLocation to location of item 1 of nextEventTitle
  if MeetingLocation is missing value then
   set MeetingLocation to "?"
  end if
  set MeetingTitle to subject of item 1 of nextEventTitle

  set MeetingStartDate to start time of item 1 of nextEventTitle
  set MeetingStartTime to time string of MeetingStartDate
  set revisedTime to MeetingStartTime as string
  set myAMPM to text -1 thru -3 of revisedTime
  set revisedTime to text 1 thru -7 of revisedTime
  set cleanLocation to my replace("ConfRm_","",MeetingLocation)
  set cleanLocation to my replace(" (Rangeway)","",cleanLocation)

  return MeetingTitle & "->" & cleanLocation & "@" & revisedTime
 else
  return "No meeting in the time frame specified"

 end if

end tell
