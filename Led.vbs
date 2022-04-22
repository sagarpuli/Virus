Set wshShell =wscript.CreateObject("WScript.Shell")
do
wscript.sleep 10
wshshell.sendkeys "{CAPSLOCK}"
wshshell.sendkeys "{NUMLOCK}"
wshshell.sendkeys "{SCROLLLOCK}"
loop