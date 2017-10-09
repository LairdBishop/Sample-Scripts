'==========================================================================
'
' NAME: 	Set Computer Name
'
' AUTHOR: 	Dell Computer Corporation - Laird Bishop
' DATE: 	01/07/2016
'
' COMMENT: 	This script is provided as an example which can be modified
'		to change a computer name.
'			
'		This code is provided without warrenty and its use does
'		not imply support.  Any use of this script i based on the assumption that:
'			You are testing this for suitability and
'			that you accept any risk resulting from its use
'			That any maintenance necessary is your responsibility
'
'==========================================================================
'##########################################################################
'#######################       SETTINGS       #############################
'##########################################################################
' Modify items in this section only to control script behavior.  
' Read through comments for instructions.

' How to build a custom prefix based on default gateway
' Assign a prefx for each of your default gateways
'This will only work in Post-Delivery configuration, therefore USE_WMI must be true.
Select Case GetDefaultGateway
	Case "10.208.54.1"
		strPrefix="AUS"
	Case "10.208.128.1"
		strPrefix="DAL"
	Case "10.208.196.1"
		strPrefix="HOU"
	Case Else
		strPrefix="UNK" 'Unknown
End Select

'@@@@@@@@@@@@@@@@@@@  BUILD A CUTOM COMPUTER NAME HERE  @@@@@@@@@@@@@@@@@@@
'	Uncomment only one of the following examples, or build your own based on these examples
	StrName = GetSerial '										1ZTBQN1
'	StrName = strPrefix & "-" & GetSerial '						AUS-1ZTBQN1
'	StrName = GetAsset '										405314
'	StrName = ChassisType & "-" & GetSerial ' 					D-1ZTBQN1
'	strName = strPrefix & "-" & ChassisType & "-" & GetAsset '	AUS-D-405314
'	strName = strPrefix & ChassisType & GetSerial '				AUSD1ZTBQN1
'	strName = strPrefix & "-" & ChassisType & "-" & GetSerial '	AUS-D-1ZTBQN1
'****************************  TEST MODE  **********************************
' TEST_MODE allows the script to be tested without actually changing the computer name.
' If TEST_MODE=True then only a message box containing the computer name will appear.
' If TEST_MODE=False then the computer name will actually be changed without any prompting.  No message box will appear.
' Change to False for production !!WARNING!!  IF FALSE WILL CHANGE THE NAME OF THE COMPUTER ON WHICH IT IS RUN WITHOUT FURTHER PROMPTING!!
Const TEST_MODE=True

'****************************   USE WMI   **********************************
' USE_WMI determines how the computer name is changed.
' If USE_WMI=True then the computer name will be changed using the Rename method of the Win32_ComputerSystem WMI class.
'	- If WMI is being used the script should be called from the Post-Delivery Configuration section of the task sequence
'	- The call should occur prior to any Join Domain or Workgroup step
' 	- The script should be followed by a reboot to the currently installed OS prior to joining the domain
' If USE_WMI=False then the script will set the OSDComputerName task sequence variable.
'	- This does not diectly change the computer name.  The variable will be consumed by the task sequence engine at the appropriate time.
'	- NOTE: If the call to this script follows the point in the task sequence when the OSDComputerName variable has already been consumed
'			then setting the variable will have no effect.  If that is the case follow the instructions above for USE_WMI=True.
'			Using this mode, the script should be placed just prior to the Apply Operating System step.
'	- Since Microsoft.SMS.TSEnvironment can only be instantiated during the execution of a task sequence, TEST_MODE cannot be used in this mode.
'			If the script is executed in this mode outside of a running task sequence an error will result.
Const USE_WMI=False

'**************************  USE INPUT BOX   ********************************
' NOTE: This option is interactive, and therefore must occur in Post-Delivery Configuration
' NOTE: This option can only be used in the Dell factory if USE_WMI=True.
' If USE_INPUT_BOX=True then the user will be prompted for a computer name.
Const USE_INPUT_BOX=False
	strPrompt = "Please enter a computername for this computer" 'Text which appears in the body of the dialog
	strTitle = "NAME THIS COMPUTER" 'Text which appears in the title bar
	strDefaultValue = strName 'Text will appear in the input box, already selected so that it can be typed over
'	intXPosition = 200 'The number of pixels from the left edge of the screen to display the input box (Without this it is hidden behind the task sequence progress dialog). COMMENT TO CENTER.
	intYPosition = 300 'The number of pixels from the top of the screen to display the input box (Without this it is hidden behind the task sequence progress dialog).

'############################# SCRIPT BODY #################################

' Prepare the computer name
If USE_INPUT_BOX then
	strName=AskUser(strPrompt,strTitle,strDefaultValue,intXPosition,intYPosition)
End If 'USE_INPUT_BOX

' Take action
If TEST_MODE then 
	msgBox StrName
Else
	If USE_WMI then
		Wscript.Quit(SetComputerName(strName))
	Else
		Wscript.Quit(CInt(SetOSDComputerName(strName)))
	End If
End If ' TEST_MODE

'#########################  FUNCTIONS AND SUBS  ###########################

Function GetSerial
'	Input: None
'	Returns: A string containing the Dell Service Tag (Serial number) as returned from BIOS
	Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")

	Set colservicetag=objWMIService.ExecQuery("Select * from Win32_Bios")
	For Each objservicetag in colservicetag
		strSerial = objservicetag.serialnumber
	Next
	GetSerial=strSerial
End Function  'GetSerial

Function GetAsset
'	Input: None
'	Returns: A string containing the Asset Tag as returned from BIOS
'		*NOTE*: The asset tag field must be populated for this to work.
'			Talk to your CFI PM to have asset tags set at the factory

	Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
	Set colAssetTag=objWMIService.ExecQuery("Select * from Win32_SystemEnclosure")
	For Each objAssetTag in colAssetTag
		strAsset = objAssetTag.SMBIOSAssetTag
	Next 'objAssetTag in colAssetTag
	'Return "BLANK" if the asset tag field is not populated
	If instr(strAsset," ") then strAsset="BLANK"
	GetAsset=strAsset
End Function  'GetAsset

Function ChassisType
'	Input: None
'	Returns: A string indicating the chassis type of the current computer
'		VM - Virtual Machine
'		L  - Laptop
'		D  - Desktop
'		U  - Unknown
'	For a complete list of possible ChassisTypes values refer to https://msdn.microsoft.com/en-us/library/aa394474(v=vs.85).aspx

	Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
	Set colChassis = objWMIService.ExecQuery("Select * from Win32_SystemEnclosure")
	For Each objChassis in colChassis
	    For  Each strChassisType in objChassis.ChassisTypes
			Select Case strChassisType
				Case "1"
					strType="VM"
				Case "8","9","10","14"
					strType = "L"
				Case "3","4","6","7","15"
					strType = "D"
				Case Else
					strType = "U"
			End Select
	    Next 'strChassisType in objChassis.ChassisTypes
	Next 'objChassis in colChassis
	ChassisType=strType
End Function 'ChassisType

Function SetComputerName(strComputerName)
'	Input:  A string containing a valid computer name
'	Returns: Return value of Win32_ComputerSystem.Rename method
	
	Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
	Set colComputers = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
	For Each objComputer in colComputers
	    SetComputerName = objComputer.Rename(strComputerName)
	Next 'objComputer in colComputers
End Function 'SetComputerName

Function SetOSDComputerName(strComputerName)
'	Input:  A string containing a valid computer name
'	Returns: Return value of 0 for success or any other value for failure

	' !! IF A TASK SEQUENCE IS NOT RUNNING THEN THIS WILL PODUCE AN ERROR !!
	SET env = CreateObject("Microsoft.SMS.TSEnvironment")

	env("OSDCOMPUTERNAME") = strComputerName
	If env("OSDCOMPUTERNAME") = strComputerName then
		SetOSDComputerName = 0 'Success
	Else
		SetOSDComputerName = 1 ' Failed to set variable
	End If 'env("OSDCOMPUTERNAME") = strComputerName
End Function 'SetOSDComputerName

Function AskUser(sPrompt,sTitle,sDefVal,iXPos,iYPos)
	If iXPos < 1 AND iYPos < 1 Then 
		AskUser=InputBox(sPrompt,sTitle,sDefVal)
		Exit Function
	ElseIf iXPos < 1 Then
		AskUser=InputBox(sPrompt,sTitle,sDefVal,,iYPos)
		Exit Function
	ElseIf iYPos < 1 Then
		AskUser=InputBox(sPrompt,sTitle,sDefVal,iXPos)
		Exit Function
	Else
		AskUser=InputBox(sPrompt,sTitle,sDefVal,iXPos,iYPos)
	End If
	
End Function 'AskUser

Function GetDefaultGateway
'	Input: None
'	Returns: A string containing IP address of the default gateway for the last NIC with IP enabled.
	Set objWMIService = GetObject("winmgmts:")
	Set colNicConfig = objWMIService.ExecQuery("SELECT * FROM " & _
	  "Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
	For Each objNicConfig In colNicConfig
	  Result = objNicConfig.DefaultIPGateway
	  If IsArray(Result) then strDefaultIPGateway = Join(objNicConfig.DefaultIPGateway, ",")
	Next
	GetDefaultGateway = CStr(strDefaultIPGateway)
End Function 'GetDefaultGateway