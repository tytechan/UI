WinWait("[CLASS:#32770]", "", 10)
ControlFocus("��", "", "Edit1")

#ControlSetText("��" ,"", "Edit1", $CmdLine[1])
ControlSetText("��" ,"", "Edit1", "123")

For $i=1 To $CmdLine[0]
   MsgBox(1,"����Ĳ�����",$CmdLine[$i])
Next


MsgBox(2,"$CmdLineRaw ",$CmdLineRaw)

Sleep(2000)
ControlClick("��", "","Button1");
