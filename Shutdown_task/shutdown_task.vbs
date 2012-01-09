Dim oShell
Dim nHour
Set WshShell = CreateObject("WScript.Shell")
Set oShell = WScript.CreateObject("WScript.Shell")
Dim objFSO
Dim objFSOText
Dim objFolder
Dim objFile
Dim strDirectory
Dim strFile
Dim strText
Dim strText1
Dim strText2
Dim strText3
Dim strText4
Dim strText5
Dim dateStamp

dateStamp = Date()
nHour = Hour(Now())



strDirectory = "c:\Temp"
strFile = "\shutdown_log.txt"
strText = "Script initiated at  " & dateStamp & "  " & nHour
strText1 = "  Hour was between 12am and 7am  "
strText2 = "  No User input - shutdown initiated at  " & dateStamp & "  " & nHour
strText3 = "  User chose to shutdown at  " & dateStamp & "  " & nHour
strText4 = "  User canceled shutdown at  " & dateStamp & "  " & nHour
strText5 = "  Hour was not between 12am and 7am  "




Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(strDirectory)
Const ForAppending = 8
Set objFile = objFSO.OpenTextFile(strDirectory & strFile, ForAppending, True)


objFile.WriteLine(strText)

'start of program: IF the time is between 12am and 7am THEN continue program ELSE do nothing

if (nHour >= 0 And nHour < 7) then
objFile.WriteLine(strText1)

	intButton = WshShell.Popup ("To conserve energy Windows will automatically shutdown in 5 minutes.", 300, "ATUS Labs Alert!", 1 + 48)

	select case intButton
 		 case -1
   			strMessage = "Windows is Shutting Down Now."
				WScript.Sleep 300000
				oShell.Run "C:\Windows\system32\shutdown.exe /p /f"
				Set oShell = Nothing
				objFile.WriteLine(strText2)
 		 case 1
  			strMessage = "Windows is Shutting Down Now."
				WScript.Sleep 300
				oShell.Run "C:\Windows\system32\shutdown.exe /p /f"
				Set oShell = Nothing
				objFile.WriteLine(strText3)
 		 case 2
  			strMessage = "Shutdown has been canceled."
			objFile.WriteLine(strText4)
	end select

	WshShell.Popup strMessage, , , 64 

else
objFile.WriteLine(strText5)

end if


objFile.Close

Set objFSO = Nothing
Set objFolder = Nothing
Set objFile = Nothing



WScript.Quit
