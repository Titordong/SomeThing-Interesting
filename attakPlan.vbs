Dim WshShell,goal,Str 

Set WshShell=WScript.CreateObject("WScript.Shell")

WshShell.Run"D:\Tencent\Bin\QQScLauncher.exe /uin:2486972462 /quicklunch:0D69D76A148C96AD3BD3F1342ECD460AB9BD04E223B626D8E3619539AB22AE296747E393DC293F33"

WScript.Sleep 2000

goal=Inputbox("输入轰炸次数",,20)

Str=InputBox("输入轰炸内容",,"")

WshShell.run "mshta vbscript:clipboardData.SetData("+""""+"text"+""""+","+""""&str&""""+")(close)",0,true

For i=1 To goal

WshShell.SendKeys "^v"

WshShell.SendKeys "{ENTER}"

WScript.Sleep 100

Next

WScript.Sleep 3000

WshShell.SendKeys"%{F4}"