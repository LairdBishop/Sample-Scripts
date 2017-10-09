'==========================================================================
'

' NAME: 		Set Computer Name to Input Sample Script
'

' AUTHOR: 		Dell Computer Corporation

' DATE: 		02/27/2013
'

' COMMENT: 		This script is provided as an example which can be modified

'				to change a computer name based on a prompted value.

'			

'				This code is provided without warrenty and its use does

'				not imply support.
'

'==========================================================================

Const TEST_MODE = True 'Change to False for production !!WARNING!!  IF FALSE WILL CHANGE THE NAME OF THE COMPUTER ON WHICH IT IS RUN WITHOUT FURTHER PROMPTING!!


If TEST_MODE Then

	MsgBox(InputBox("Please enter a computername for this computer","NAME THIS COMPUTER","",30))

	WScript.Quit(0)

Else

	Return = SetComputerName(InputBox("Please enter a computername for this computer","NAME THIS COMPUTER","",30))

	WScript.Quit(Return)

End If



Function SetComputerName(strComputerName)

'	Input:  A string containing a valid computer name

'	Returns: Return value of Win32_ComputerSystem.Rename method

	Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")

	
Set colComputers = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")

	
For Each objComputer in colComputers

	    SetComputerName = objComputer.Rename(strComputerName)

	Next

End Function 'SetComputerName


