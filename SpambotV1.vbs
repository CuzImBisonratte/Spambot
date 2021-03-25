Set WshShell = CreateObject("WScript.Shell")
WScript.Sleep 1000
For i = 1 To 5
  WshShell.SendKeys "MESSAGE"
  WScript.Sleep 15
  WshShell.SendKeys "{ENTER}"
  WScript.Sleep 15
Next