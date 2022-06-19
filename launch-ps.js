shell = WScript.createObject("WScript.Shell");
ret = shell.Run("powershell.exe -File " + WScript.Arguments.Item(0), 0, false);
WScript.Quit(ret);