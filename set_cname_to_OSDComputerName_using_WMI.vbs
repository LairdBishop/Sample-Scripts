'==========================================================================
'
' NAME: 	Set Computer Name to OSDComputerName using WMI
'
' AUTHOR: 	Dell Computer Corporation - Laird Bishop
' DATE: 	05/12/2016
'
' COMMENT: 	This script is provided as an example which can be modified
'			to change a computer name.  Please test for suitability in your 
'			environment before placing in production.
'			
'			This code is provided without warrenty and its use does
'			not imply support.
'
'==========================================================================


'############################# SCRIPT BODY #################################

RetVal = SetComputerName(GetOSDComputerName)
Wscript.Quit(RetVal)

'#########################  FUNCTIONS AND SUBS  ###########################





Function SetComputerName(strComputerName)
'	Input:  A string containing a valid computer name
'	Returns: Return value of Win32_ComputerSystem.Rename method
	
	Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
	Set colComputers = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
	For Each objComputer in colComputers
	    SetComputerName = objComputer.Rename(strComputerName)
	Next 'objComputer in colComputers
End Function 'SetComputerName

Function GetOSDComputerName()
'	Input:  None
'	Returns: The value of the OSDComputerName task sequence variable

	' !! IF A TASK SEQUENCE IS NOT RUNNING THEN THIS WILL PODUCE AN ERROR !!
	SET env = CreateObject("Microsoft.SMS.TSEnvironment")

	GetOSDComputerName = env("OSDCOMPUTERNAME")

End Function 'GetOSDComputerName



