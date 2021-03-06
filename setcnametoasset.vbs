'==========================================================================
'
' NAME: 		Set Computer Name to Asset Tag Sample
'
' AUTHOR: 		Dell Computer Corporation
' DATE: 		6/28/2012
'
' COMMENT: 		This script is provided as an example which can be modified
'				to change a computer name based on the Asset tag field In
'				the BIOS.  
'			
'				This code is provided without warrenty and its use does
'				not imply support.
'
'  ASSUMPTIONS:	The Asset Tag field of the BIOS is populated with a valid
'				computer name prior to the execution of this script.
'
'==========================================================================

Return = SetComputerName(GetAssetTag)
WScript.Quit(Return)

Function SetComputerName(strComputerName)
'	Input:  A string containing a valid computer name
'	Returns: Return value of Win32_ComputerSystem.Rename method
	Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
	
	Set colComputers = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
	
	For Each objComputer in colComputers
	    SetComputerName = objComputer.Rename(strComputerName)
	Next
End Function 'SetComputerName


Function GetAssetTag()
'	Input: None
'	Returns: The contents of the Asset tag field of the BIOS
	Set objWMIService=GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
	Set colSMBIOSAssetTag=objWMIService.ExecQuery("Select SMBIOSAssetTag from Win32_SystemEnclosure")
	For Each objAssetTag in colSMBIOSAssetTag
		strName = objAssetTag.SMBIOSAssetTag
	Next
	If Len(strName) < 1 Then strName = "BLANK"
	GetAssetTag = strName
End Function 'GetAssetTag
