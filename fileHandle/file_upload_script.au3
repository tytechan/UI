WinWait("[CLASS:#32770]", "", 10)
ControlFocus("��", "", "Edit1")

ControlSetText("��" ,"", "Edit1", $CmdLine[1])
Sleep(2000)
ControlClick("��", "","Button1");
