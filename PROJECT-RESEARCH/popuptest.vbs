Dim oShell
Dim nHour
Set WshShell = CreateObject("WScript.Shell")
Set oShell = WScript.CreateObject("WScript.Shell")
nHour = Hour(Now())


if (nHour >= 0 or nHour < 7) then

intButton = WshShell.Popup ("Click a button to proceed.", 5, , 1 + 48)

select case intButton
  case -1
    strMessage = "Computer Will Shutdown Now."
	WScript.Sleep 100
	oShell.Run "shutdown -i"
	Set oShell = Nothing
  case 1
    strMessage = "You clicked the OK button."
	WScript.Sleep 100
	oShell.Run "shutdown -i"
	Set oShell = Nothing
  case 2
    strMessage = "You clicked the Cancel button."
end select

WshShell.Popup strMessage, , , 64 

end if