Dim Message, Title, Default, MyValue, MyMessage, MyWait
Message = "Wie oft soll die nachricht gesendet werden?"
Title = "Spambot V3"
Default = "1"
MyValue = InputBox(Message, Title, Default)

Message = "Welche Nachricht soll "+MyValue+" mal gesendet werden?"
Title = "Spambot V3"
Default = "Ich bin eine Nachricht!"
MyMessage = InputBox(Message, Title, Default)

Message = "Wie viele Millisekunden sollen zwischen den Nachrichten gewartet werden? (1 Sek = 1000 Millisek)"
Title = "Spambot V3"
Default = "15"
MyWait = InputBox(Message, Title, Default)

Message = "Wie viele Sekunden soll vor der ersten Nachricht gewartet werden?"
Title = "Spambot V3"
Default = "15"
MyWaitStart = InputBox(Message, Title, Default)

MsgBox "", vbYesNo, "Überprüfen"

Set WshShell = CreateObject("WScript.Shell")
WScript.Sleep MyWaitStart*1000
For i = 1 To MyValue
  WshShell.SendKeys MyMessage
  WScript.Sleep MyWait
  WshShell.SendKeys "{ENTER}"
  WScript.Sleep MyWait
Next