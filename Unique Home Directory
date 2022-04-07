ON ERROR RESUME NEXT

'--------------------------------------------------------------------------------------------------
' DECLARATIONS AND VARIABLES
'--------------------------------------------------------------------------------------------------

Dim WSHShell, WSHNetwork, objShell, objDomain, DomainString, UserString, UserObj, Path, Psvr, fso, HomeFolder

Set WSHShell = CreateObject("WScript.Shell")
Set WSHNetwork = CreateObject("WScript.Network")
Set objShell = CreateObject("WScript.Shell")
Set fso = CreateObject ("Scripting.FileSystemObject")

UserString = UCase(WSHNetwork.UserName)
strComputer = UCase(WSHNetwork.ComputerName)
WinDir = WshShell.ExpandEnvironmentStrings("%WinDir%")
ProgFileDir = WshShell.ExpandEnvironmentStrings("%PROGRAMFILES%")

DomainString = Wshnetwork.UserDomain

Set UserObj = GetObject("WinNT://" & DomainString & "/" & UserString)
Set objRootDSE = GetObject("LDAP://RootDSE")
UserObj.GetInfo

strDomain = objRootDSE.Get("DefaultNamingContext")
strDomain = replace(strDomain, "DC=", "")
ADDom = replace(strDomain, ",", ".")

strHomeDirectory = UserObj.Get("homeDirectory")

'--------------------------------------------------------------------------------------------------
' Start
'--------------------------------------------------------------------------------------------------
' *** EDIT THIS LINE ONLY *** 
BaseFolder = "\\{server name or path}\" & strComputer & "\"

Call MapHomeFolder(BaseFolder)
Wscript.Quit
'--------------------------------------------------------------------------------------------------
' End
'--------------------------------------------------------------------------------------------------



'--------------------------------------------------------------------------------------------------
' FUNCTIONS & SUBS
'--------------------------------------------------------------------------------------------------

Sub MapHomeFolder(BaseFolder)

	Set FSO = CreateObject("Scripting.FileSystemObject")

	If Not FSO.FolderExists(BaseFolder) Then 
		FSO.CreateFolder(BaseFolder)
	End If

	UDate = Now()

	HomeFolder = BaseFolder & Day(UDate) & "-" & Month(UDate) & "-" & Year(UDate) & "-" & Hour(UDate) & "-" & Minute(UDate) & "-" & Second(UDate)

	If Not FSO.FolderExists(HomeFolder) Then 
		FSO.CreateFolder(HomeFolder)
	End If

	Call MapNetworkDrive ("H",HomeFolder,"My Documents")

End Sub



'--------------------------------------------------------------------------------------------------
' DRIVE MAPPING FUNCTIONS
' DRIVE TYPES: 0 = Unknown, 1 = Removable, 2 = Fixed, 3 = Network, 4 = CD-ROM, 5 = RAM Disk
'--------------------------------------------------------------------------------------------------

Function MapNetworkDrive(DrvLetter, MapPath, NameSpace)

	RemoveNetworkDrive(DrvLetter)

	If fso.DriveExists(DrvLetter & ":") Then
	
			If fso.GetDrive(DrvLetter).DriveType = 1 or fso.GetDrive(DrvLetter).DriveType = 2 or fso.GetDrive(DrvLetter).DriveType = 4 Then

				'do nothing

			Else

				Set objShell = CreateObject("Shell.Application")
				WSHNetwork.MapNetworkDrive DrvLetter & ":", MapPath, True
				objShell.NameSpace(DrvLetter & ":").Self.Name = NameSpace

			End If
	Else

		Set objShell = CreateObject("Shell.Application")
		WSHNetwork.MapNetworkDrive DrvLetter & ":", MapPath, True
		objShell.NameSpace(DrvLetter & ":").Self.Name = NameSpace
		
	End If

End Function


Function RemoveNetworkDrive(DrvLetter)

	If fso.DriveExists(DrvLetter & ":") Then  
	
		If fso.GetDrive(DrvLetter).DriveType = 3 Then
			WSHNetwork.RemoveNetworkDrive DrvLetter & ":", True, True
		End If
		
	End If

End Function
