' InspectorFileRetention.vbs
' This script to if inspectors' uploaded files were older than 183 days when the md3200i storage size is reached to 95%
' This will configure in windows schedule and run everyday
' added by Jimmy @8Aug2018

' variable declaration
Const strPath = "G:\fileserver-107\inspector\savehere"

Dim objFSO
Dim filecount
Dim SizePercentage 
Dim CutOffDay
Dim strName
Dim flag
Dim strToday
Dim logFilename

strToday = Replace(Date(), "/" ,"_" )
logFilename = "E:\Log\InspectorFileRetention-" & strToday & ".log"
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set f = objFSO.OpenTextFile(logFilename, 8, true)
flag = true

' Call Search(strPath)

' Capture the program starting datetime
f.WriteLine "start ===== " & Now

' get the percentage of free space in md3200i
Call getFreeSpace() 

 ' Capture the program starting datetime
 f.WriteLine "Storage available percentage: " & SizePercentage

' if it is over 98 percent, remove files oder than 90 days
' if it is over 80 percent, remove files oder than 183 days
SizePercentage = 6
If SizePercentage < 3 Then
   CutOffDay = 90
   call Search(strPath)
ElseIf SizePercentage < 5 Then
   CutOffDay = 150
   call Search(strPath)
ElseIf SizePercentage < 13 Then
   CutOffDay = 183
   call Search(strPath)
End If

' WScript.Echo"Done."


Set fso = CreateObject("Scripting.FileSystemObject") 


' a function to delete files in folder and subfolder based on the cutoffday 
Sub Search(str)
    Dim objFolder, objSubFolder, objFile
    Set objFolder = objFSO.GetFolder(str)
    For Each objFile In objFolder.Files
        ' f.WriteLine "CutOffDay: " & CutOffDay  
        If objFile.DateCreated < (Now() - CutOffDay) Then
		    
	        if 	flag <> false Then 
		      f.WriteLine "folder: " & objFolder.Name & " , " & "fileCount: " & objFolder.Files.Count 
			  flag = false
			end if
			  objFile.Delete(True)
        End If
 
    Next
    For Each objSubFolder In objFolder.SubFolders
	   f.WriteLine "Subfolder: " & objSubFolder.Name & " , " & "fileCount: " & objSubFolder.Files.Count
       Search(objSubFolder.Path)
        ' Files have been deleted, now see if the folder is empty.
        'If (objSubFolder.Files.Count = 0) Then
        '   objSubFolder.Delete True
        'End If 
    Next
End Sub

' a function to get the available space of percentage
sub getFreeSpace()
 
 Set objWMIService = GetObject("winmgmts:")
 Set objLogicalDisk = objWMIService.Get("Win32_LogicalDisk.DeviceID='G:'")
 SizePercentage = CInt((objLogicalDisk.FreeSpace / objLogicalDisk.Size) * 100)
 
End Sub

f.WriteLine "END ===== " & Now

' exit program
set objFSO = Nothing
set objLogicalDisk = Nothing
set objWMIService = Nothing