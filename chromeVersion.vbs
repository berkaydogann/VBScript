Set objFSO = CreateObject("Scripting.FileSystemObject")
file = "C:\Program Files\Google\Chrome\Application\chrome.exe"
set x = CreateObject("wscript.shell")
x.run "C:\Windows\system32\notepad.exe"
wscript.sleep 2000
x.sendkeys(objFSO.GetFileVersion(file))
