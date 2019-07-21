WinWait("[CLASS:#32770]", "", 10)
ControlFocus("打开", "", "Edit1")

#ControlSetText("打开" ,"", "Edit1", $CmdLine[1])
ControlSetText("打开" ,"", "Edit1", "123")

For $i=1 To $CmdLine[0]
   MsgBox(1,"传入的参数是",$CmdLine[$i])
Next


MsgBox(2,"$CmdLineRaw ",$CmdLineRaw)

Sleep(2000)
ControlClick("打开", "","Button1");
