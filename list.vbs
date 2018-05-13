Option explicit
On Error Resume Next

dim objFSO, objStartFolder, aFiles, oShell, objLog
dim strFilePaths, n

Set objFSO = CreateObject("Scripting.FileSystemObject")
set oShell=CreateObject("WScript.Shell")

strFilePaths = array("")

objStartFolder = Replace(WScript.ScriptFullName, WScript.ScriptName,vbNullString )

If (objFSO.FolderExists(objStartFolder)) = False Then
                WScript.Echo "The path entered is either invalid or cannot be found. Please check and try again."
                WScript.Quit
End If

Set objLog = objFSO.OpenTextFile(objStartFolder & "\" & CreateObject("WScript.Network").UserName & Right(Day(Date) + 100, 2) & Right(Month(Date) + 100, 2) & Year(Date) & Right(Hour(Time) + 100, 2) & Right(Minute(Time) + 100, 2) & Right(Second(Time) + 100, 2) & ".csv" ,2, True) ' Output File

objLog.WriteLine chr(34) &  "FileName" & chr(34) & "," & chr(34) & "FullPath" & chr(34) & "," & chr(34) & "FileSize(bytes)" & chr(34) & "," & chr(34) & "DateCreated" & chr(34) & "," & chr(34) & "DateLastModified" & chr(34) & "," & chr(34) & "DateLastAccessed" & chr(34)
 

for n = lbound(strFilePaths) to ubound(strFilePaths)
                sGetFileDetails strFilePaths(n), objLog
Next
 
objLog.close

set objFSO = Nothing
set oShell = Nothing
set objLog = Nothing

msgbox "Done"
 

Sub sGetFileDetails(sFolder,objLog )
Dim oFileSys, oFolder, aFiles, aSubFolders, file, folder

Set oFileSys = WScript.CreateObject("Scripting.FileSystemObject")
Set oFolder = oFileSys.GetFolder(sFolder)
Set aFiles = oFolder.Files
Set aSubFolders = oFolder.SubFolders

For Each file in aFiles
                objLog.WriteLine chr(34) &  file.Name & chr(34) & "," & chr(34) & file.Path & chr(34) & "," & chr(34) & file.Size & chr(34) & "," & chr(34) & file.DateCreated & chr(34) & "," & chr(34) & file.DateLastModified & chr(34) & "," & chr(34) & file.DateLastAccessed & chr(34)
Next

For Each folder in aSubFolders
                sGetFileDetails folder.Path,objLog
Next

set aSubFolders = Nothing
set aFiles = Nothing
set oFolder = Nothing
set oFileSys = Nothing
               

End Sub