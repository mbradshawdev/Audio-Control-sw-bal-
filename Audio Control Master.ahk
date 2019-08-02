/*==================================================================\
|              Audio Control (Sw Bal) - Version 1.44.1              |
\===================================================================/

DESCRIPTION:
 Script to automate common interactions with Win7/10 "Sound" (mmsys.cpl) interface
 Primarily to switch audio devices (set default) and change balance values on 5.1 devices (hence 'sw' & 'bal')
 Limited support for toggling some 'Enhancements' (Fill, Loudness..) exists but no longer a focus of this script

HOW TO USE:
 This script was built around my own needs but can be adjusted to your own needs in 'CUSTOMISATION' section below

TO-DO:
 Finish customisation options (move all references to names/nicknames to 'CUSTOMISATION' section
 Add arrays for nickname strings
 Add support for up to 10 devices
 Experiment with making changes with Sound window hidden
 Extra help for closing any possible sub-windows of Sound main window (eg 'HDMI Audio Device Properties')

CUSTOMISATION:
 Replace these with your device* and preferred names:
*/
   dev1Name = Speakers
   dev1Nick = Desk

   dev2Name = Headphones
   dev2Nick = Bed

   dev3Name = Bluetooth Audio Renderer
   dev3Nick = Blu

   dev4Name = AMD HDMI Output
   dev4Nick = TV

   dev5Name = Your Device Name Here
   dev5Nick = Your device nickname here

/*Device names are displayed and can be renamed in the classic 'Sound' interface (Start->run->mmsys.cpl)
 To rename: Open properties and the device name should be highlighted and editable

CHANGELOG:
 1.44.0 02-08-2019  Big tidy up, added simple customisation support
 1.43.3 27/07/2019  TV Speakers added
 1.43.2       2018  Minor Text Fixes lol
 1.43.1       2018  Increased bluetooth wait time 
 1.43         2017  More 'Bal51' options and standard set (no space/dot)
 1.42         2017  Loudness Equalisation Update (loudMode,loudToggle)
 1.41         2017  Set Update
 1.40         2017  Set Update
 1.32         2017  Manual shift update 27.11.2017
 1.30.0 01/11/2017  Log starts Nov 2017
*/
;--------------------------------------------------------------------------------------------------------------
;=========SETTINGS=============================================================================================

#NoEnv
;#Warn 
#SingleInstance force
SetTitleMatchMode, RegEx
SetWorkingDir %A_ScriptDir%

;========MAIN PROGRAM==========================================================================================

Loop
 IfWinNotActive Launchy		;Ensure Launchy is hidden
   break			;Usually not a problem

WinGet, restoreActiveWindow, ID, A
restoreActiveWindow = ahk_id %restoreActiveWindow%	;Ensure previously top-most window is active again after

GoSub, GetOpsInfo					;Get info about which operation+arguments from filename

;Loop { ;RetryLoop					;On second and later loops values come fron user input

 if shiftMode {
  GoSub, GetShiftValues
  GoSub, FindDevice					;When setting bal fill hsould always be off, this ensures it
  GoSub, PropertiesLevelsBalance				;Properties should already be open but skips just that button press and jumps to  second (tab1) tab
  GoSub, DoShifts
  GoSub, CloseAll
  Sleep 3000
  ExitApp
  }  

 if setMode {
  GoSub, GetSetValues
  GoSub, FindDevice
  GoSub, PropertiesLevelsBalance			;Properties should already be open but skips just that button press and jumps to second (tab1) tab
  GoSub, GetCurrentVals
  GoSub, SetBalance
  GoSub, CloseAll
  Sleep 3000
  ExitApp
  }

 if balanceMode {					;From filename, Balance mode IS selected
  if !defaultMode && !manualMode
   GoSub, BalValsFilename				;Get desired balance values from filename
  if !balValsRaw && !defaultMode && !manualMode		;filename values not found
   GoSub, BalValsUserInput				;Request balance values from user
  if !balValsRaw and !desBal1 and !desBal2 and !desBal3 and !desBal4 and !desBal5 and !desBal6		 
   GoSub, DefaultInput					;still no values, use defaults set in this script
  GoSub, ParseValues					;Finally, parse the values
  GoSub, FindDevice					;When setting bal fill hsould always be off, this ensures it
;  GoSub, PropertiesEnhancements				;When adjusting balance..
;  GoSub, FillToggle					;..fill should always be off. SubRoutines BackToSound and BackToProperties were for this but unnecessary
  GoSub, PropertiesLevelsBalance				;Properties should already be open but skips just that button press and jumps to second (tab1) tab
  GoSub, GetCurrentVals
  GoSub, SetBalance
  GoSub, CloseAll
  Sleep 3000
  ExitApp
  }

 else if fillMode {					;Fill mode first to allow 'sp fill mode' in filename
  GoSub, FindDevice
  GoSub, PropertiesEnhancements
  GoSub, FillToggle
  GoSub, CloseAll
  Sleep 3000 
  }

 else if loudMode {					;Fill mode first to allow 'sp fill mode' in filename
  GoSub, FindDevice
  GoSub, PropertiesEnhancements
  GoSub, LoudToggle
  GoSub, CloseAll
  Sleep 3000 
  }

 else if switchMode {				;From filename, SWITCH mode is selected
  ToolTip Setting %desDev% as default device
  GoSub, FindDevice
  GoSub, SetDefaultDevice
  GoSub, CloseAll
  Sleep 3000
  ExitApp
  }

 else
  MsgBox Fail idk
;} ;RetryLoop

Return

;--------------------------------------------------------------------------------------------------------------
;=========SUB-ROUTINES=========================================================================================

GetOpsInfo:

opsInfoRaw := A_ScriptName				;On first run script checks filename for options
	
Loop { ;GetOpsLoop					;Loop to try again: First from filename info, subsequent from user inputs

 ;Operations
  if (InStr(opsInfoRaw, "Shift")) or (InStr(opsInfoRaw, "Up")) or (InStr(opsInfoRaw, "Increase")) or (InStr(opsInfoRaw, "Down")) or (InStr(opsInfoRaw, "decrease"))
   shiftMode = 1					;Shift mode on (SHIFT, UP, INCREASE, DOWN, DECREASE)
  else if InStr(opsInfoRaw, "Set")			;Set mode on (SET)
   setMode = 1
  else if InStr(opsInfoRaw, "Bal")			;Balance mode on (BAL)
   balanceMode := 1					;, toggleOff = 1 (disabling this as fill off causes difficulties
  if (InStr(opsInfoRaw, "Sw"))				;not (InStr(opsInfoRaw, "Sp"))
   switchMode = 1					;Device switch mode only on SW

 ;Legacy commands, leave for now
  if InStr(opsInfoRaw, "Fill")				;Speaker Fill mode
   fillMode := 1					;, toggleOff := ""
  if InStr(opsInfoRaw, "Loud")				;Speaker Fill mode
   loudMode := 1					;, toggleOff := ""

;----------------------------------
 ;Devices

;To-Do:
; if matches official name up to 10
;  use that
; else check nickname arrays up to 10
;  use associated official name 'device#OfficialName' I guess
; else
;  NoDeviceFlag = 1 and continue

Loop, 10
 if InStr(opsInfoRaw, dev%A_Index%Name) or InStr(opsInfoRaw, dev%A_Index%Nick)
  {
   desDev := dev%A_Index%Name
   break
  }
 
;----------------------------------
 ;Special Options
  if (InStr(opsInfoRaw, "Manual")) or (InStr(opsInfoRaw, "Menu"))
   manualMode = 1					;Manual: just open to desired menu
  if (InStr(opsInfoRaw, "Default"))
   defaultMode = 1					;Default skips to using default bal vals
  if InStr(opsInfoRaw, "Off")				;Toggle Speaker Fill, Loudness Equalisation (possibly more later) OFF
    toggleOff = 1					; Default Fill, LoudEq etc mode is on, turning it off requires this additional 'off'
  ;else if (InStr(opsInfoRaw, "On"))			;Disabled for now because ON is default
   ;onMode = 1

;----------------------------------
;Presets						;Passes on preset values ultimately to BalValsFilename SubRoutine
 if (InStr(opsInfoRaw, "bal51 30")) or (InStr(opsInfoRaw, "bal 51 30")) {
  opsInfoRaw = bal 30 30 90 10 100 100			
  userOpsMode = 1
  }

 else if (InStr(opsInfoRaw, "bal51 50")) {
  opsInfoRaw = bal 50 50 90 10 100 100			
  userOpsMode = 1
  }

 else if (InStr(opsInfoRaw, "bal51")) {
  opsInfoRaw = bal 70 70 90 01 100 100			
  userOpsMode = 1
  }

 else if (InStr(opsInfoRaw, "Stereo")) {
  opsInfoRaw = bal 100 100 0 10 0 0			;Passes on preset values ultimately to BalValsFilename SubRoutine
  userOpsMode = 1
  balanceMode = 1					;This line removes necessary to add 'bal'
  }
 else if (InStr(opsInfoRaw, "Annoy")) {
  opsInfoRaw = bal 0 0 0 0 0 100			;Back Right only (used to fight back against very inconsiderate neighbours with a loud TV
  userOpsMode = 1
  }

;----------------------------------
;Special Rules
 if !desDev						;If no desired device is set..
  if balanceMode or shiftMode or setMode or fillMode or loudMode		;..currently for all modes..
   desDev = Speakers					;..assume Desk Speakers
  ;desDev = Default Device				; NOT ..assume current (Default) Device
 ;if (InStr(opsInfoRaw, "Stereo"))	;idk was to do with fill but now meh
 ;if fillMode && !onMode && !toggleOff			;If fill mode selected but not on or off
 ; onMode = 1	
					;..assume on. UNNECESSARY due "if not 'off' turn on"
;----------------------------------			;test MsgBox opsInfoRaw %opsInfoRaw%`nfillMode: %fillMode%`ntoggleOff: "%toggleOff%"

;Continue Or Try Again With User Input (UserOps)
 if userOpsMode or shiftMode or setMode or balanceMode or switchMode or fillMode or loudMode	;At least one operation recognised
  break							;break GetOpsLoop
 else {
  userOpsMode = 1
  inputMode = User Input (UserOps)
  InputBox, opsInfoRaw, Operation Select, Enter desired Operation: Switch`, Balance`, Shift`, Fill`nDevice: Desk`, Bed`, Bluetooth`nAnd values/arguments (optional):`nManual`, Default`, Off`, 80 80 90 20 100 100 (bal)`, 20 (shift) etc`n%errorText%
  errorText = Error try again...
  }
 } ;GetOpsLoop						
inputErrorCount = 
errorText = 
Return

;--------------------------------------------------------------------------------------------------------------
GetShiftValues:

 if userOpsMode						;User has input UserOps instructions 
  shiftValRaw := opsInfoRaw				;Initially grab potential raw shift values from OpsInfo user input
 else 
  shiftValRaw := A_ScriptName				;Initially grab potential raw shift values from filename

 inputMode = Filename					;Mode is filename unless changed to userinput
 Loop { ;MainShiftLoop						;Test raw shift values (initially from filename then from user input)
  if InStr(shiftValRaw, "front")				;CHANNELS 
   shiftChan1 := 1, shiftChan2 = 2			;If front is found the channels to shift are 1 and 2
  else if InStr(shiftValRaw, "cen")			;If cen: just channel 3 etc...
   shiftChan1 := 3
  else if InStr(shiftValRaw, "sub")
   shiftChan1 := 4
  else if InStr(shiftValRaw, "back left")
   shiftChan1 := 5
  else if InStr(shiftValRaw, "back right")
   shiftChan1 := 6
  else if InStr(shiftValRaw, "rear")			;
   shiftChan1 := 5, shiftChan2 = 6			;
  else { ;NoChannel					;None of 'front' 'cen' 'sub' or 'rear' were found
   chanError := "Channel"
   } ;NoChannel

  if InStr(shiftValRaw, "down")				;SHIFT DIRECTION (look for 'up' or 'down' in raw values)
   shiftDir := "Left"
  else if InStr(shiftValRaw, "up")	
   shiftDir := "Right"
  else { ;NoDir						;Neither 'up' nor 'down' was found
   if chanError
    dirError := ", Direction"
   else
    dirError := "Direction"
   } ;NoDir

 RegExMatch(shiftValRaw, "\d+", shiftMag)		;SHIFT MAGNITUDE (Find a 1 or more digit number)
 if shiftMag not between 1 and 100			;Check 1 - 100 inclusive
  {
   if chanError or dirError
    magError1 := ", Magnitude"
   else
    magError1 := "Magnitude"    			;magError2 :=  % "# Magnitude not valid (" shiftMag ") #`n"

   shiftMag =
  }

 if shiftChan1 && shiftDir && shiftMag			;Everything was successful
  break							;break out of the loop

 else {	;UserInputMode					;One or more element was not successful (channel direction magnitude)
  inputmode = User Input (Shift)				;Just for reporting outcome
  if !shiftFail						;No previous fails do first input box
   InputBox, shiftValRaw, Input desired SHIFT, %chanError%%dirError%%magError1% not found/valid in filename%magError2%`nEnter desired shift:,,,150
  else							;Previous fails, do second input box with error message
   InputBox, shiftValRaw, Input desired SHIFT (try again), %chanError%%dirError%%magError1% still not found/valid%magError2%`nEnter desired shift and include:`nFRONT`, CEN`, SUB or REAR and..`nUP or DOWN and..`nA value from 1 to 100
  if ErrorLevel {
   ToolTip Giving up`, huh?`nQuitting...
   Sleep 1000
   WinActivate, %restoreActiveWindow%
   ExitApp
   }
  shiftFail++
  } ;UserInputMode

 chanError = 
dirError = 
magError1 =
magError2 =
 } ;MainShiftLoop
  shiftFail = 
Return

;--------------------------------------------------------------------------------------------------------------
GetSetValues:

 if userOpsMode						;User has input UserOps instructions 
  setValRaw := opsInfoRaw				;Initially grab potential raw set values from OpsInfo user input
 else 
  setValRaw := A_ScriptName				;Initially grab potential raw set values from filename
 inputMode = Filename					;Mode is filename unless changed to userinput

 Loop { ;MainSetLoop					;Test raw set values (initially from filename then from user input)
  if InStr(setValRaw, "front left")
   setChan1 := 1
  else if InStr(setValRaw, "front right")
   setChan1 := 2
  else if InStr(setValRaw, "front")			;CHANNELS 
   setChan1 := 1, setChan2 = 2				;If front/rear setChan1 and 2 are used to hold 1/2 and 5/6
  else if InStr(setValRaw, "cen")			;If cen: just channel 3 etc...
   setChan1 := 3
  else if InStr(setValRaw, "sub")
   setChan1 := 4
  else if InStr(setValRaw, "back left")
   setChan1 := 5
  else if InStr(setValRaw, "back right")
   setChan1 := 6
  else if InStr(setValRaw, "rear")
   setChan1 := 5, setChan2 = 6
  else { ;NoChannel					;None of 'front' 'cen' 'sub' or 'rear' were found
   chanError := "Channel"
   } ;NoChannel

 RegExMatch(setValRaw, "\d+", setVal)		;SHIFT MAGNITUDE (Find a 1 or more digit number)
 if setVal not between -1 and 100			;Check 0 - 100 inclusive
  {
   if chanError
    magError1 := ", Magnitude"
   else
    magError1 := "Magnitude"    			;

   shiftMag =
  }

 if setChan1 && (setVal > -1)				;Everything was successful
  break							;break out of the loop

 else {	;UserInputMode					;One or more element was not successful (channel or value)
  inputmode = User Input (Set)				;Just for reporting outcome
  if !setFail						;No previous fails do first input box
   InputBox, setValRaw, Input desired SET value, %chanError%%magError1% not found/valid in filename`nEnter desired SET value:,,,150
  else							;Previous fails, do second input box with error message
   InputBox, setValRaw, Input desired SET value (try again), %chanError%%magError1% still not found/valid`nEnter desired SET value and include:`nFRONT`, CEN`, SUB or REAR and..`nA value from 0 to 100
  if ErrorLevel {
   ToolTip Giving up`, huh?`nQuitting...
   Sleep 1000
   WinActivate, %restoreActiveWindow%
   ExitApp
   }
  setFail++
  } ;UserInputMode

 chanError = 
 magError1 =
 } ;MainSetLoop
 setFail = 
 balanceMode = 1 


 desBal%setChan1% := setVal
 if setChan2
  desBal%setChan2% := setVal

Return

;--------------------------------------------------------------------------------------------------------------
BalValsFilename:

 if userOpsMode						;User has input UserOps instructions 
  RegExMatch(opsInfoRaw, "(\d+[\s,-\.]+){5,5}\d+", balValsRaw)
  ;OLD  balValRaw := opsInfoRaw				;Initially grab potential raw balance values from this OpsInfo user input
 else {							;Original filename mode
  ;Get values from filename (one or more digit followed by space or ,-.)x5 then one more digit (more robust)
  RegExMatch(A_ScriptName, "(\d+[\s,-\.]+){5,5}\d+", balValsRaw)
  inputMode = Filename 					;Over-written if other balance value input mode used instead
  }
Return

;--------------------------------------------------------------------------------------------------------------
BalValsUserInput:
 InputBox, balValsRaw, Enter Balance Values (# # # # # #), Filename values missing or invalid`nEnter desired balance values for these speakers:`nLeft`, Right`, Centre`, SubWoofer`, Rear-Left`, Rear-Right`neg: 80 80 70 10 100 100
  inputMode = User Input (Balance Grouped)
 Loop {
  if !balValsRaw
   break
  if (balValsRaw = "x") {
   balValsRaw =
   GoSub, GetLevelsIndivid	;Legacy option: request each value individually (this only completes with valid values)
   break
   }
  valuesOk := RegExMatch(balValsRaw, "(\d+[\s,-\.]+){5,5}\d+", balValsRaw)
  if valuesOk
   break
  InputBox, balValsRaw, Enter Balance Values (# # # # # #), Filename values missing or invalid`nEnter desired balance values for these speakers:`nLeft`, Right`, Centre`, SubWoofer`, Rear-Left`, Rear-Right`neg: 80 80 70 10 100 100`nError`, try again
  }
Return

;--------------------------------------------------------------------------------------------------------------
DefaultInput:	;Default values
  inputMode = Default (Balance)
  desBal1 = 80	;Left
  desBal2 = 80	;Right	;Manually edit these values as desired!
  desBal3 = 70	;Centr
  desBal4 = 10	;Sub
  desBal5 = 100	;RearL
  desBal6 = 100	;RearR
Return

;--------------------------------------------------------------------------------------------------------------
SoundCloseOpen:
;Close Existing Sound window
 Loop 
  if WinExist("^Sound$") {				;Check if Sound is open already
   WinClose ^Balance$
   WinClose Properties$
   WinClose ^Sound$
   if (A_Index > 10)
    run TASKKILL /F /IM "rundll32.exe",, Hide		;Slow, only if necessary
   }
  else							;If it isn't open
   break						;Continue

;Open Sound options
 Loop {							
  Run, mmsys.cpl					;Open Sound menu (to Playback tab)
  WinWait, ^Sound$,, 2					;Wait for it for 2 seconds
  if ErrorLevel {					;Sound menu failed to load
   Loop 3						;
    run TASKKILL /F /IM "rundll32.exe",, Hide		;Task Kill x3 to close any bugged instance
   Sleep 200						;
   }
  else							;Sound loaded successfully 
   break						;continue on to next block
  ToolTip Fails: %A_Index% 				;Display fail count
  if (A_Index > 4) {					
   MsgBox, Part 1 failed`, try again? idk		;Something unexpected happened
   ExitApp
   }
  }
Return

;--------------------------------------------------------------------------------------------------------------
GetLevelsIndivid:	;Request inputs for each speaker individually
 inputMode = User Input (Balance Individual)
 ch1Name := "LEFT", ch2Name = "RIGHT", ch3Name = "CENTER", ch4Name = "SUB", ch5Name = "REAR-LEFT", ch6Name = "REAR-RIGHT"		;Set name for each channel
 Loop 6 {																;Six channels
  j := A_Index																;Counter for sub-loop
  InputBox, desBal%j%, % "desBalance Levels: " ch%j%Name, % "Enter desired " ch%j%Name " speaker balance (1-100)"				;First input
  Loop {																;loop to escape only on valid entry
   if desBal%j% =																;blank entry is valid
    break																;break sub-loop to enter next channel
   if desBal%j% is number															;must be a number AND..
    if (desBal%j% > -1) and (desBal%j% < 101)													;Between 0 and 100 inclusive
     break																;as 3 above re break
   InputBox, desBal%j%, % "Balance Levels: " ch%j%Name, % "Enter desired " ch%j%Name " speaker balance (1-100)`nError`,try again"		;Second input includes error message
   } ;sub-loop to ensure correct entry
  } ;per channel loop

Return

;--------------------------------------------------------------------------------------------------------------
ParseValues:
StringReplace, balVals, balValsRaw, `,, %A_SPACE%, All					;Change commas, to spaces
StringReplace, balVals, balVals, ., %A_SPACE%, All					;Change periods. to spaces
StringReplace, balVals, balVals, -, %A_SPACE%, All					;Change hyphens- to spaces

Loop {											;Reduce all multi spaces to single spaces
 StringReplace, balVals, balVals, %A_SPACE%%A_SPACE%, %A_SPACE%, UseErrorLevel
 if ErrorLevel = 0  ; No more replacements needed.
  break
 }

Loop, parse, balVals, %A_Space%`,							;Split eg 80 80 70 10 100 100 into desBal1 desBal2 etc
{
 desBal%A_Index% := A_LoopField
}
Return

;--------------------------------------------------------------------------------------------------------------
FindDevice:
Loop {	;RetryLoop						                ;Retries whole Switch process once or more
 GoSub, SoundCloseOpen
 
 ;manual + switchmode so just end script now that sound dialogue should already be open ready to switch devices or whatever manually
 if (switchMode and manualMode) or !desDev {              ;In either case just quit but different tooltips
  if (switchMode and manualMode)	                     ;Manual and Switch aka sw man
   ToolTip Sound dialogue opened for manual use (or error maybe idk)`nQuitting...
  else             						                 ;No device set, so just 
   ToolTip Sound options opened`, no device specified`nQuitting...
  Sleep 1000                                             ;Display one of two ToolTips for 1sec
  ExitApp                                                ;Just exit at this point
  }
 
 ToolTip Searching for "%desDev%"
 Loop {	;ListRestartLoop    	                        ;Restarts search from top of list
  ControlSend, SysListView321, {HOME}, ^Sound$						    ;Focus on first item in list
 ;WinActivate Sound
 ;Send {Home}
  ControlGet, i, List, Count, SysListView321, Sound 	;Get number of devices
  Loop { ;DeviceWaitLoop		                        ;Allows script to wait for "Not Plugged In" devices
   ControlGet, x, List, Selected, SysListView321, Sound	;Get info (x = Device: Name, Description, Status)
   if InStr(x, desDev)					                ;If device info (x) contains desired device (desDev) name
    if InStr(x, "Not Plugged In") {			            ;Device is disconnected (not plugged in) as with waiting on Bluetooth headset
     i = 100						                    ;count becomes timer to search for approx 0.5sec per 1
     notPlugged = 1					                    ;Desired device found but not connected (usually Bluetooth)
     ToolTip (%A_Index%/%i%) Waiting for "%DesDev%"`nPress [ESC] to cancel  ;Display allowed time to conenct (a few sec or i/2)
     Sleep 500
     }
    else ;if InStr(x, "Not Plug..	                    ;Device is not disconnected (probably connected)
     break 3						                    ;success, break all 3 loops (DeviceWait, ListRestart, and Retry loops) to continue (next hits the return at end of this SR)
   else { ;if InStr(x, desDev)	                        ;Not correct device
    if notPlugged					                    ;Correct device was selected but has disappeared (expected baheviour briefly before device connects)
     break						                        ;break per device loop to restart list (to re-find now-connected device)
    ;Send {Down}						                    ;Normal ops: Move down one to check next list item
    ControlSend, SysListView321, {Down}, ^Sound$		                    ;Normal ops: Not correct device so move down one to check next list item
    }
   if (A_Index >= i)					                ;Checked all list items OR reached end of i timer
    break 2                                             ;Break out of per device loop AND list restart loop (hits "retries whole thing once loop")
   } ;DeviceWaitLoop
  notPlugged = 						                    ;notPlug is cleared for restarting list to (hopefully) find newly connected device
  Sleep 2000						                    ;allowance for BT headset to report available
  } ;/ListRestartLoop
 failedAttempts++ 
if (A_Index >= 5) {
  MsgBox Failed attempts: %failedAttempts%`nDevice not found`nEnsure "show disconnected devices" is checked`nCheck name of device `/ filename`nQuitting...
  ExitApp
  }
 } ;/RetryLoop
Return

;--------------------------------------------------------------------------------------------------------------
PropertiesEnhancements:									;0 (no tab) is General, 1 Levels, 2 Enhancements
 if !WinExist("Properties$")
  ControlClick, &Properties, Sound,,,na							;[Properties] button if not already open
 Sleep 50
 SendMessage, 0x1330, 2,, SysTabControl321, Properties$					;1 Fancy-ass way to tab 
 Sleep 0  ; This line and the next are necessary only for certain tab controls.		;2 "
 SendMessage, 0x130C, 2,, SysTabControl321, Properties$					;3 "
 if fillMode && manualMode {								;Manual mode, just opens to Balance so quit at this point
  MsgBox Enhancements window opened for manual control press [OK] when done
  Sleep 500
  Exitapp
  }
Return

;--------------------------------------------------------------------------------------------------------------
PropertiesLevelsBalance:
 sleep 50
 if !WinExist("Properties$")
  ControlClick, &Properties, Sound,,,na							;[Properties] button if not already open
 Sleep 50
 WinWait Properties$
 SendMessage, 0x1330, 1,, SysTabControl321, Properties$				;1 Fancy-ass way to tab 
 Sleep 0  ; This line and the next are necessary only for certain tab controls.		;2 "
 SendMessage, 0x130C, 1,, SysTabControl321, Properties$				;3 "
 Loop
  if !WinExist("Balance") {								;If it's not open, open it
   ControlClick,Button3, Properties$,,,na						;Ctrl Click for Balance loop for efficiency+reliability
   sleep 100
   }
  else											;If it IS open stop trying to open it -> break -> return
   break
 if manualMode {									;Manual mode, just opens to Balance so quit at this point
  MsgBox Balance window opened for manual control press [OK] when done
  Sleep 500
  Exitapp
  }
Return

;--------------------------------------------------------------------------------------------------------------
FillToggle:
  Sleep 50
 ControlClick, &Restore Defaults, Properties$
  Sleep 50
 ControlFocus, SysListView321, Properties$
 if !toggleOff
  Send {Up 3}{Down}{Space}
  Sleep 50
 ControlClick, OK, Properties$
  Sleep 50
Return     

;--------------------------------------------------------------------------------------------------------------
LoudToggle:
  Sleep 50
 ControlClick, &Restore Defaults, Properties$
  Sleep 50
 ControlFocus, SysListView321, Properties$
 if !toggleOff
  Send {Up 3}{Down 3}{Space}
  Sleep 50
 ControlClick, OK, Properties$
  Sleep 50
Return   

/*
EnhancementsControl:
  Sleep 50
 ControlClick, &Restore Defaults, Properties$
  Sleep 50
 ControlFocus, SysListView321, Properties$
 if bassManagement {
  Send {Space}
  }
 Send {Down}
 if speakerFill {
  Send {Space}
  Sleep 50
  }
 Send {Down}
 if roomCorrect {
  Send {Space}
  Sleep 50
  }
 Send {Down}
 if loudnessEq {
  Send {Space}
  }
  Sleep 50
 ControlClick, OK, Properties$
  Sleep 50
*/

;--------------------------------------------------------------------------------------------------------------
GetCurrentVals:
 ;Get balance values: Find literal L,R,Sub,RL etc
  WinGetText, winStuff, Balance
  RegExMatch(winStuff, "L\s*\K\d+", curVal1)	;L Match literal L (or R, Sub, RL etc)
  RegExMatch(winStuff, "R\s*\K\d+", curVal2)	;\s* none or more whitespace
  RegExMatch(winStuff, "C\s*\K\d+", curVal3)	;\K but don't return this***
  RegExMatch(winStuff, "Sub\s*\K\d+", curVal4)	;\d+ then return 1 or more digits
  RegExMatch(winStuff, "RL\s*\K\d+", curVal5)	;RegExMatch auto terminates on \r or \n idk
  RegExMatch(winStuff, "RR\s*\K\d+", curVal6)	;***Previously used extra line: RegExMatch(balRR, "\d+", balRR)
  ;MsgBox %winStuff%`n`nL: %balL%`nR: %balR%`nC: %balC%`nSub: %balSub%`nRL: %balRL%`nRR: %balRR%
Return     

;--------------------------------------------------------------------------------------------------------------
SetBalance:
 Loop 6 {
  dir%A_Index% = Right							;Initialise each movement direction as Right
  move%A_Index% := desBal%A_Index% - curVal%A_Index%			;Get movement value (difference between 
  if move%A_Index% > 50							;If movement value is over 50..
   ToolTip VOLUME WARNING!						;display a volume warning tooltip
  if (move%A_Index% < 0)				 		;if movement value is negative.. (if not keep dir:Right and Move val is already positive)
   dir%A_Index% := "Left", move%A_Index% := Abs(move%A_Index%)		;set "Left" movement and take absoute value (-2 becomes 2)
  if move%A_Index% && !(desBal%A_Index% = "") 				;Skip if movement value is 0
   ;MsgBox % "Yes`, move" A_Index " & desBal: '" desBal%A_Index% else MsgBox % "No`, move" A_Index " & desBal: '" desBal%A_Index%
   ;Info: trackbar321 322 323 etc (luckily end # is 1-6), dir# = L or R for 1-6 similar move# is movement value for 1-6, sent to Balance window
   ControlSend, msctls_trackbar32%A_Index%, % "{" dir%A_Index% " " move%A_Index% "}", ^Balance$
;  MsgBox % "move" A_Index ": "move%A_Index% " = desBal" A_Index ": " desBal%A_Index% " - curVal" A_Index ": " curVal%A_Index% "dir" A_Index ": " dir%A_Index%
  }
    sleep 50								;Delay here prevents final keystroke from being lost
   ToolTip Speaker balance set`n%inputMode% values used`n_ %desBal1%_ %desBal2%_ %desBal3%_ %desBal4%_ %desBal5%_ %desBal6%_`nQuitting...
;Sometimes values shift up or down one due to rounding I guess
Return

;--------------------------------------------------------------------------------------------------------------
DoShifts:
; ControlSend, msctls_trackbar32%shiftChan1%, % "{" dir%A_Index% " " move%A_Index% "}", ^Balance$
 ControlSend, msctls_trackbar32%shiftChan1%, {%shiftDir% %shiftMag%}, ^Balance$
 if shiftChan2 {											;if rear or front
  ControlSend, msctls_trackbar32%shiftChan2%, {%shiftDir% %shiftMag%}, ^Balance$			;do second channel/speaker as well
  ToolTip %A_ScriptName%`nChannels %shiftChan1% & %shiftChan2% shifted %shiftDir% by %shiftMag%`nQuitting...
  }
 else	;Do single channel tooltip
  ToolTip %A_ScriptName%`nChannel %shiftChan1% shifted %shiftDir% by %shiftMag%`nQuitting...
 Sleep 50												;sleep here prevents final notch being lost
Return

;--------------------------------------------------------------------------------------------------------------
SetDefaultDevice:
 ControlClick,&Set Default,^Sound$,,,,na		;Control Click Set Active button (no mouse move)
  Sleep 50
 ControlClick,&Apply,^Sound$,,,,na			;As above for Apply
  Sleep 50
   ToolTip %desDev% set as default device`nQuitting...
Return

;--------------------------------------------------------------------------------------------------------------
CloseAll:
 ;Done and OK to close ..until closed lol
 Loop	
  if WinExist("^Sound$") {
   ControlClick,&OK,^Balance$,,,,na
   ControlClick,OK,Properties$,,,,na
   ControlClick,OK,^Sound$,,,,na
   }
  else
   break
 WinActivate, %restoreActiveWindow%
Return     

;--------------------------------------------------------------------------------------------------------------
BackToSound:
 ;Done and OK to close only back to sound window
 Loop	
  if WinExist("^Balance$") or WinExist("Properties$") {
   ControlClick,&OK,^Balance$,,,,na
   ControlClick,OK,Properties$,,,,na
   }
  else
   break
Return    

;--------------------------------------------------------------------------------------------------------------
BackToProperties:	;Probs won't use this, currently might as well just be single line: ControlClick,&OK,^Balance$,,,,na
 ;Done and OK to close only back to properties
 Loop	
  if WinExist("^Balance$") {
   ControlClick,&OK,^Balance$,,,,na
   }
  else
   break
Return    

;--------------------------------------------------------------------------------------------------------------
;=============HOTKEYS==========================================================================================

ESC::
 ToolTip Quitting
 Sleep 400
 ExitApp
Return

;--------------------------------------------------------------------------------------------------------------

;==============TIMERS==============
;No timers used for now
;--------------------------------------------------------------------------------------------------------------

