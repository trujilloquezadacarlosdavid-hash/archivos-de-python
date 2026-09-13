' Lanza REFRESCAR_HEINZ.bat con la ventana OCULTA (0) para que no aparezca consola cada hora.
Set sh = CreateObject("WScript.Shell")
d = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))
sh.Run """" & d & "REFRESCAR_HEINZ.bat""", 0, False
