'// *************************************************************************** //
'// This script will copy ALL source files and folders to a Destination folder  //
'// As long as the source folder is newer by a minimum of 1 second or higher    //
'// *************************************************************************** //

On Error Resume Next

Dim SrcFolder, DstFolder, SrcDate, DstDate
Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")

' ***********************************************
' *** Edit these two lines only *****************
' ***********************************************
SrcFolder = "{enter source folder here}"
DstFolder = "{enter destination folder here}"
' ***********************************************

Call BuildPathIfNotExists(DstFolder)

If objFSO.FolderExists(SrcFolder) Then

	If objFSO.FolderExists(DstFolder) Then
	
		SrcDate = GetModifiedDate(SrcFolder)
		DstDate = GetModifiedDate(DstFolder)

		If DateDiff("s", DstDate, SrcDate) > 1 Then
		
			Set objWshShell = WScript.CreateObject("WScript.Shell")
			Call objWshShell.Run("XCOPY """ & SrcFolder & """ """ & DstFolder & """ /R /I /C /H /K /E /Y", 1, True)
			
		End If

	End If
	
End If


Sub BuildPathIfNotExists(FullPath)

  If Not objFSO.FolderExists(FullPath) Then

    Call BuildPathIfNotExists(objFSO.GetParentFolderName(FullPath))
    objFSO.CreateFolder FullPath
	
  End If
  
End Sub

Function GetModifiedDate(filespec)

	Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim f : Set f = objFSO.GetFolder(filespec)
	
	GetModifiedDate = f.DateLastModified
   
End Function
