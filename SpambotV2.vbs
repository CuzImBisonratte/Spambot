MyMessage = "Ich bin ne Nachricht" 'Nachricht zum senden
MyLoopValue = "3" 'Wie oft die Nachricht
MyWait = "15" 'Wie lang zwischen den nachrichten
Set WshShell = CreateObject("WScript.Shell")
WScript.Sleep 1000
For i = 1 To MyLoopValue
  WshShell.SendKeys MyMessage
  WScript.Sleep MyWait
  WshShell.SendKeys "{ENTER}"
  WScript.Sleep MyWait
Next