Set ws=CreateObject("wscript.shell")

Dim a,cut,b,Str,flag
flag=false
cut=1
a=Inputbox("你相信这世界上存在圣诞老人吗？",,"相信")
MsgBox "..."
MsgBox "其实我就是圣诞老人"
MsgBox "虽然我没有礼物给你"
MsgBox "呜呜，你是不是生气啦"
MsgBox "那我给你变个魔术好不好"
a=Inputbox("要不要变魔术?",,"要")
If a<>"要" Then 
	MsgBox "呜呜，真的不给机会吗？"
	Do 
	b=InputBox(Str+"试一下吧",,"不")
	Str=Str&"请"
	If b<>"不"Then 
		flag=true
		Exit Do 
	Else cut=cut+1
	End If 
	If cut=5 Then 
		MsgBox "那好吧，再见！"
		Exit Do 
	End If
	Loop
Else flag=true
End If 
If flag=true Then 
Set WS=WScript.CreateObject("WScript.Shell")
WS.run("notepad"),3
WScript.Sleep 500
WS.AppActivate("notepad")
WS.SendKeys "+"
arr=Array("听","说","你","被","拐","卖","真","的","把","我","吓","坏","你","虽","从","小","痴","呆","却","对","社","会","无","害","谁","人","如","此","大","胆","竟","敢","拿","你","去","卖","我","真","替","他","担","心","卖","得","掉","才","怪","")
'f(arr)
'WScript.Sleep 3000
'For i=0 To 90
'	WScript.Sleep 30
'	WS.SendKeys "{BS}"
'Next 
arr=Array("好","啦","好","啦","不","逗","你","了","")
Clipboard="MsHta vbscript:ClipBoardData.setData(""Text"","""&str1&""")(Window.Close)"
WS.Run(Clipboard)
	For i=0 To 8
		'WScript.Sleep 200
		'Clipboard="MsHta vbscript:ClipBoardData.setData(""Text"","""&arr(i)&""")(Window.Close)"
		'WS.Run(Clipboard)
		'WS.SendKeys"^v"
	Next
'WScript.Sleep 3000
'For i=0 To 90
	'WScript.Sleep 30
	'WS.SendKeys "{BS}"
'Next 
AutoTime=75
say(AutoTime)

end If
Function link()
Set Seven = WScript.CreateObject("WScript.Shell")
strDesktop = Seven.SpecialFolders("Desktop")
set oShellLink = Seven.CreateShortcut(strDesktop & "\Titordong.url")
oShellLink.TargetPath = "https://www.cnblogs.com/Titordong/"
oShellLink.Save
Set oShellLink=Nothing
strDesktop = Seven.SpecialFolders(4)
mypath=strDesktop&"\Titordong.url"
Seven.run mypath
End Function

Function f(a)
	Dim str1
	Clipboard="MsHta vbscript:ClipBoardData.setData(""Text"","""&str1&""")(Window.Close)"
	WS.Run(Clipboard)
	For i=0 To 47
		WScript.Sleep 200
		Clipboard="MsHta vbscript:ClipBoardData.setData(""Text"","""&a(i)&""")(Window.Close)"
		WS.Run(Clipboard)
		WS.SendKeys"^v"
		If i>0 Then
		If(i Mod 6=0) Then
			WScript.Sleep 400
			WS.SendKeys "{ENTER}"
		End If
		End IF
	Next
End Function

Function say(AutoTime)
WScript.Sleep AutoTime
WS.SendKeys "D"
WScript.Sleep AutoTime
WS.SendKeys "e"
WScript.Sleep AutoTime
WS.SendKeys "a"
WScript.Sleep AutoTime
WS.SendKeys "r"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "f"
WScript.Sleep AutoTime
WS.SendKeys "r"
WScript.Sleep AutoTime
WS.SendKeys "i"
WScript.Sleep AutoTime
WS.SendKeys "e"
WScript.Sleep AutoTime
WS.SendKeys "n"
WScript.Sleep AutoTime
WS.SendKeys "d"
WScript.Sleep AutoTime
WS.SendKeys ":"
WScript.Sleep 1000
WS.SendKeys "{ENTER}"


WS.SendKeys "{ENTER}"


WScript.Sleep AutoTime
WS.SendKeys "H"
WScript.Sleep AutoTime
WS.SendKeys "o"
WScript.Sleep AutoTime
WS.SendKeys "w"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "a"
WScript.Sleep AutoTime
WS.SendKeys "r"
WScript.Sleep AutoTime
WS.SendKeys "e"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "y"
WScript.Sleep AutoTime
WS.SendKeys "o"
WScript.Sleep AutoTime
WS.SendKeys "u"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "g"
WScript.Sleep AutoTime
WS.SendKeys "o"
WScript.Sleep AutoTime
WS.SendKeys "i"
WScript.Sleep AutoTime
WS.SendKeys "n"
WScript.Sleep AutoTime
WS.SendKeys "g"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "r"
WScript.Sleep AutoTime
WS.SendKeys "e"
WScript.Sleep AutoTime
WS.SendKeys "n"
WScript.Sleep AutoTime
WS.SendKeys "c"
WScript.Sleep AutoTime
WS.SendKeys "e"
WScript.Sleep AutoTime
WS.SendKeys "n"
WScript.Sleep AutoTime
WS.SendKeys "t"
WScript.Sleep AutoTime
WS.SendKeys "l"
WScript.Sleep AutoTime
WS.SendKeys "y"
WScript.Sleep AutoTime
WS.SendKeys "?"
WScript.Sleep 1000
WS.SendKeys "{ENTER}"


WS.SendKeys "{ENTER}"


WScript.Sleep AutoTime
WS.SendKeys "I"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "h"
WScript.Sleep AutoTime
WS.SendKeys "e"
WScript.Sleep AutoTime
WS.SendKeys "a"
WScript.Sleep AutoTime
WS.SendKeys "r"
WScript.Sleep AutoTime
WS.SendKeys "d"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "t"
WScript.Sleep AutoTime
WS.SendKeys "h"
WScript.Sleep AutoTime
WS.SendKeys "a"
WScript.Sleep AutoTime
WS.SendKeys "t"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "y"
WScript.Sleep AutoTime
WS.SendKeys "o"
WScript.Sleep AutoTime
WS.SendKeys "u"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "h"
WScript.Sleep AutoTime
WS.SendKeys "a"
WScript.Sleep AutoTime
WS.SendKeys "v"
WScript.Sleep AutoTime
WS.SendKeys "a"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "m"
WScript.Sleep AutoTime
WS.SendKeys "a"
WScript.Sleep AutoTime
WS.SendKeys "d"
WScript.Sleep AutoTime
WS.SendKeys "e"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "a"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "g"
WScript.Sleep AutoTime
WS.SendKeys "i"
WScript.Sleep AutoTime
WS.SendKeys "r"
WScript.Sleep AutoTime
WS.SendKeys "l"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "f"
WScript.Sleep AutoTime
WS.SendKeys "r"
WScript.Sleep AutoTime
WS.SendKeys "i"
WScript.Sleep AutoTime
WS.SendKeys "e"
WScript.Sleep AutoTime
WS.SendKeys "n"
WScript.Sleep AutoTime
WS.SendKeys "d"
WScript.Sleep AutoTime
WS.SendKeys ","
WScript.Sleep 1000
WS.SendKeys "{ENTER}"


WScript.Sleep AutoTime
WS.SendKeys "a"
WScript.Sleep AutoTime
WS.SendKeys "n"
WScript.Sleep AutoTime
WS.SendKeys "d"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "s"
WScript.Sleep AutoTime
WS.SendKeys "h"
WScript.Sleep AutoTime
WS.SendKeys "e"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "i"
WScript.Sleep AutoTime
WS.SendKeys "s"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "b"
WScript.Sleep AutoTime
WS.SendKeys "e"
WScript.Sleep AutoTime
WS.SendKeys "a"
WScript.Sleep AutoTime
WS.SendKeys "u"
WScript.Sleep AutoTime
WS.SendKeys "t"
WScript.Sleep AutoTime
WS.SendKeys "i"
WScript.Sleep AutoTime
WS.SendKeys "f"
WScript.Sleep AutoTime
WS.SendKeys "u"
WScript.Sleep AutoTime
WS.SendKeys "l"
WScript.Sleep AutoTime
WS.SendKeys "."
WScript.Sleep 1000
WS.SendKeys "{ENTER}"


WScript.Sleep AutoTime
WS.SendKeys "I"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "a"
WScript.Sleep AutoTime
WS.SendKeys "m"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "v"
WScript.Sleep AutoTime
WS.SendKeys "e"
WScript.Sleep AutoTime
WS.SendKeys "r"
WScript.Sleep AutoTime
WS.SendKeys "y"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "s"
WScript.Sleep AutoTime
WS.SendKeys "a"
WScript.Sleep AutoTime
WS.SendKeys "d"
WScript.Sleep AutoTime
WS.SendKeys ","
WScript.Sleep AutoTime
WS.SendKeys "y"
WScript.Sleep AutoTime
WS.SendKeys "o"
WScript.Sleep AutoTime
WS.SendKeys "u"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "h"
WScript.Sleep AutoTime
WS.SendKeys "a"
WScript.Sleep AutoTime
WS.SendKeys "v"
WScript.Sleep AutoTime
WS.SendKeys "e"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "s"
WScript.Sleep AutoTime
WS.SendKeys "a"
WScript.Sleep AutoTime
WS.SendKeys "i"
WScript.Sleep AutoTime
WS.SendKeys "d"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "y"
WScript.Sleep AutoTime
WS.SendKeys "o"
WScript.Sleep AutoTime
WS.SendKeys "u"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "w"
WScript.Sleep AutoTime
WS.SendKeys "i"
WScript.Sleep AutoTime
WS.SendKeys "l"
WScript.Sleep AutoTime
WS.SendKeys "l"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "n"
WScript.Sleep AutoTime
WS.SendKeys "e"
WScript.Sleep AutoTime
WS.SendKeys "v"
WScript.Sleep AutoTime
WS.SendKeys "e"
WScript.Sleep AutoTime
WS.SendKeys "r"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "l"
WScript.Sleep AutoTime
WS.SendKeys "o"
WScript.Sleep AutoTime
WS.SendKeys "v"
WScript.Sleep AutoTime
WS.SendKeys "e"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "a"
WScript.Sleep AutoTime
WS.SendKeys "n"
WScript.Sleep AutoTime
WS.SendKeys "y"
WScript.Sleep AutoTime
WS.SendKeys "o"
WScript.Sleep AutoTime
WS.SendKeys "n"
WScript.Sleep AutoTime
WS.SendKeys "e"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "a"
WScript.Sleep AutoTime
WS.SendKeys "f"
WScript.Sleep AutoTime
WS.SendKeys "t"
WScript.Sleep AutoTime
WS.SendKeys "e"
WScript.Sleep AutoTime
WS.SendKeys "r"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "m"
WScript.Sleep AutoTime
WS.SendKeys "e"
WScript.Sleep AutoTime
WS.SendKeys "."
WScript.Sleep 1000
WS.SendKeys "{ENTER}"


WScript.Sleep AutoTime
WS.SendKeys "I"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "c"
WScript.Sleep AutoTime
WS.SendKeys "a"
WScript.Sleep AutoTime
WS.SendKeys "n"
WScript.Sleep AutoTime
WS.SendKeys "'"
WScript.Sleep AutoTime
WS.SendKeys "t"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "i"
WScript.Sleep AutoTime
WS.SendKeys "m"
WScript.Sleep AutoTime
WS.SendKeys "a"
WScript.Sleep AutoTime
WS.SendKeys "g"
WScript.Sleep AutoTime
WS.SendKeys "i"
WScript.Sleep AutoTime
WS.SendKeys "n"
WScript.Sleep AutoTime
WS.SendKeys "e"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "y"
WScript.Sleep AutoTime
WS.SendKeys "o"
WScript.Sleep AutoTime
WS.SendKeys "u"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "b"
WScript.Sleep AutoTime
WS.SendKeys "e"
WScript.Sleep AutoTime
WS.SendKeys "t"
WScript.Sleep AutoTime
WS.SendKeys "r"
WScript.Sleep AutoTime
WS.SendKeys "a"
WScript.Sleep AutoTime
WS.SendKeys "y"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "m"
WScript.Sleep AutoTime
WS.SendKeys "e"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "s"
WScript.Sleep AutoTime
WS.SendKeys "o"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "e"
WScript.Sleep AutoTime
WS.SendKeys "a"
WScript.Sleep AutoTime
WS.SendKeys "s"
WScript.Sleep AutoTime
WS.SendKeys "i"
WScript.Sleep AutoTime
WS.SendKeys "l"
WScript.Sleep AutoTime
WS.SendKeys "y"
WScript.Sleep AutoTime
WS.SendKeys "!"
WScript.Sleep 1000
WS.SendKeys "{ENTER}"


WS.SendKeys "{ENTER}"


WScript.Sleep AutoTime
WS.SendKeys "I"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "h"
WScript.Sleep AutoTime
WS.SendKeys "a"
WScript.Sleep AutoTime
WS.SendKeys "v"
WScript.Sleep AutoTime
WS.SendKeys "e"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "n"
WScript.Sleep AutoTime
WS.SendKeys "o"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "w"
WScript.Sleep AutoTime
WS.SendKeys "o"
WScript.Sleep AutoTime
WS.SendKeys "r"
WScript.Sleep AutoTime
WS.SendKeys "d"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "t"
WScript.Sleep AutoTime
WS.SendKeys "o"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "s"
WScript.Sleep AutoTime
WS.SendKeys "a"
WScript.Sleep AutoTime
WS.SendKeys "y"
WScript.Sleep AutoTime
WS.SendKeys ","
WScript.Sleep 1000
WS.SendKeys "{ENTER}"


WScript.Sleep AutoTime
WS.SendKeys "j"
WScript.Sleep AutoTime
WS.SendKeys "u"
WScript.Sleep AutoTime
WS.SendKeys "s"
WScript.Sleep AutoTime
WS.SendKeys "t"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "w"
WScript.Sleep AutoTime
WS.SendKeys "i"
WScript.Sleep AutoTime
WS.SendKeys "s"
WScript.Sleep AutoTime
WS.SendKeys "h"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "y"
WScript.Sleep AutoTime
WS.SendKeys "o"
WScript.Sleep AutoTime
WS.SendKeys "u"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "c"
WScript.Sleep AutoTime
WS.SendKeys "a"
WScript.Sleep AutoTime
WS.SendKeys "n"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "g"
WScript.Sleep AutoTime
WS.SendKeys "a"
WScript.Sleep AutoTime
WS.SendKeys "i"
WScript.Sleep AutoTime
WS.SendKeys "n"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "h"
WScript.Sleep AutoTime
WS.SendKeys "a"
WScript.Sleep AutoTime
WS.SendKeys "p"
WScript.Sleep AutoTime
WS.SendKeys "p"
WScript.Sleep AutoTime
WS.SendKeys "i"
WScript.Sleep AutoTime
WS.SendKeys "n"
WScript.Sleep AutoTime
WS.SendKeys "e"
WScript.Sleep AutoTime
WS.SendKeys "s"
WScript.Sleep AutoTime
WS.SendKeys "s"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "i"
WScript.Sleep AutoTime
WS.SendKeys "n"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "t"
WScript.Sleep AutoTime
WS.SendKeys "h"
WScript.Sleep AutoTime
WS.SendKeys "e"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "f"
WScript.Sleep AutoTime
WS.SendKeys "u"
WScript.Sleep AutoTime
WS.SendKeys "t"
WScript.Sleep AutoTime
WS.SendKeys "u"
WScript.Sleep AutoTime
WS.SendKeys "r"
WScript.Sleep AutoTime
WS.SendKeys "e"
WScript.Sleep AutoTime
WS.SendKeys "."
WScript.Sleep 1000
WS.SendKeys "{ENTER}"


WScript.Sleep AutoTime
WS.SendKeys "B"
WScript.Sleep AutoTime
WS.SendKeys "y"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "t"
WScript.Sleep AutoTime
WS.SendKeys "h"
WScript.Sleep AutoTime
WS.SendKeys "e"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "w"
WScript.Sleep AutoTime
WS.SendKeys "a"
WScript.Sleep AutoTime
WS.SendKeys "y"
WScript.Sleep AutoTime
WS.SendKeys ","
WScript.Sleep 1000
WS.SendKeys "{ENTER}"


WScript.Sleep AutoTime
WS.SendKeys "I"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "k"
WScript.Sleep AutoTime
WS.SendKeys "n"
WScript.Sleep AutoTime
WS.SendKeys "o"
WScript.Sleep AutoTime
WS.SendKeys "w"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "I"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "a"
WScript.Sleep AutoTime
WS.SendKeys "m"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "y"
WScript.Sleep AutoTime
WS.SendKeys "o"
WScript.Sleep AutoTime
WS.SendKeys "u"
WScript.Sleep AutoTime
WS.SendKeys "r"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "b"
WScript.Sleep AutoTime
WS.SendKeys "o"
WScript.Sleep AutoTime
WS.SendKeys "y"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "f"
WScript.Sleep AutoTime
WS.SendKeys "i"
WScript.Sleep AutoTime
WS.SendKeys "r"
WScript.Sleep AutoTime
WS.SendKeys "e"
WScript.Sleep AutoTime
WS.SendKeys "n"
WScript.Sleep AutoTime
WS.SendKeys "d"
WScript.Sleep AutoTime
WS.SendKeys ","
WScript.Sleep 1000
WS.SendKeys "{ENTER}"


WScript.Sleep AutoTime
WS.SendKeys "b"
WScript.Sleep AutoTime
WS.SendKeys "u"
WScript.Sleep AutoTime
WS.SendKeys "t"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "I"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "w"
WScript.Sleep AutoTime
WS.SendKeys "a"
WScript.Sleep AutoTime
WS.SendKeys "n"
WScript.Sleep AutoTime
WS.SendKeys "t"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "o"
WScript.Sleep AutoTime
WS.SendKeys "n"
WScript.Sleep AutoTime
WS.SendKeys "l"
WScript.Sleep AutoTime
WS.SendKeys "y"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "I"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "h"
WScript.Sleep AutoTime
WS.SendKeys "a"
WScript.Sleep AutoTime
WS.SendKeys "v"
WScript.Sleep AutoTime
WS.SendKeys "e"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "y"
WScript.Sleep AutoTime
WS.SendKeys "o"
WScript.Sleep AutoTime
WS.SendKeys "u"
WScript.Sleep AutoTime
WS.SendKeys "."
WScript.Sleep 1000
WS.SendKeys "{ENTER}"


WS.SendKeys "{ENTER}"


WScript.Sleep AutoTime
WS.SendKeys "S"
WScript.Sleep AutoTime
WS.SendKeys "o"
WScript.Sleep AutoTime
WS.SendKeys ","
WScript.Sleep AutoTime
WS.SendKeys "b"
WScript.Sleep AutoTime
WS.SendKeys "r"
WScript.Sleep AutoTime
WS.SendKeys "e"
WScript.Sleep AutoTime
WS.SendKeys "a"
WScript.Sleep AutoTime
WS.SendKeys "k"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "u"
WScript.Sleep AutoTime
WS.SendKeys "p"
WScript.Sleep AutoTime
WS.SendKeys "."
WScript.Sleep 1000
WS.SendKeys "{ENTER}"


WS.SendKeys "{ENTER}"


WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "-"
WScript.Sleep AutoTime
WS.SendKeys "-"
WScript.Sleep AutoTime
WS.SendKeys "-"
WScript.Sleep AutoTime
WS.SendKeys "-"
WScript.Sleep AutoTime
WS.SendKeys "Y"
WScript.Sleep AutoTime
WS.SendKeys "o"
WScript.Sleep AutoTime
WS.SendKeys "u"
WScript.Sleep AutoTime
WS.SendKeys "r"
WScript.Sleep AutoTime
WS.SendKeys "s"
WScript.Sleep AutoTime
WS.SendKeys " "
WScript.Sleep AutoTime
WS.SendKeys "H"
WScript.Sleep AutoTime
WS.SendKeys "o"
WScript.Sleep AutoTime
WS.SendKeys "n"
WScript.Sleep AutoTime
WS.SendKeys "e"
WScript.Sleep AutoTime
WS.SendKeys "y"
WScript.Sleep 1000
WS.SendKeys "{ENTER}"






End Function