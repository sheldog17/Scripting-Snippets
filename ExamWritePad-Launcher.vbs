On Error Resume Next

Dim srcFolder, trgFolder

srcFolder = "\\server name\ExamWritePad"
trgFolder = "C:\windows\Temp\ExamWritePad"

'Copy and then Execute EWP
CopyFilesAndFolders srcFolder, trgFolder
Set objShell = WScript.CreateObject("WScript.Shell")
objShell.Run trgFolder & "\ExamWritePad.exe", 1, False
WScript.Quit



Sub CopyFilesAndFolders (ByVal strSource, ByVal strDestination)

    Dim ObjFSO, ObjFolder, ObjSubFolder, ObjFile, files
    Dim TargetPath
	
    Set ObjFSO = CreateObject("scripting.filesystemobject")
	
    'connecting to the folder where is going to be searched
    Set ObjFolder = ObjFSO.GetFolder(strSource)
    TargetPath = Replace (objFolder.path & "\", strSource, strDestination,1,-1,vbTextCompare)
    
	If Not ObjFSO.FolderExists (TargetPath) Then ObjFSO.CreateFolder (TargetPath)
    
	Err.clear
	
    On Error Resume Next
	
    'Check all files in a folder
    For Each objFile In ObjFolder.files
        
		If Err.Number <> 0 Then Exit For 'If no permission or no files in folder
        
		On Error goto 0
        If CheckToCopyFile (objFile.path, TargetPath & "\" & objFile.name) Then 
            objFSO.copyfile objFile.path, TargetPath & "\" & objFile.name, True
        End If
		
    Next
	
    'Recurse through all of the subfolders
    On Error Resume Next
	
    Err.clear
	
    For Each objSubFolder In ObjFolder.subFolders
        If Err.Number <> 0 Then Exit For 'If no permission or no subfolder in folder
        On Error goto 0
        'For each found subfolder there will be searched for files
        CopyFilesAndFolders ObjSubFolder.Path & "\", TargetPath & ObjSubFolder.name & "\"
    Next
	
    Set ObjFile = Nothing
    Set ObjSubFolder = Nothing
    Set ObjFolder = Nothing
    Set ObjFSO = Nothing
	
End Sub

Function CheckToCopyFile (ByVal strSourceFilePath, ByVal strDestFilePath)
    Dim oFSO, oFile, SourceFileModTime, DestFileModTime
    
	CheckToCopyFile = True
    
	Set oFSO = CreateObject("scripting.filesystemobject")
	If Not oFSO.FileExists (strDestFilePath) Then Exit Function
    Set oFile = oFSO.GetFile (strSourceFilePath)
    SourceFileModTime = oFile.DateLastModified
    Set oFile = Nothing
    Set oFile = oFSO.GetFile (strDestFilePath)
    DestFileModTime = oFile.DateLastModified
    Set oFile = Nothing
    If SourceFileModTime =< DestFileModTime Then CheckToCopyFile = False
    Set oFSO = Nothing
	
End Function
