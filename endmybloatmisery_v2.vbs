' EndMyBloat_HTA.vbs
Option Explicit

' Reminder popup
MsgBox "Reminder: this doesn't delete any files, to continue press OK", vbInformation, "End My Bloat Misery"

' Welcome popup with OK to start, Cancel/X closes
Dim answer
answer = MsgBox("Welcome to the End My Bloat Misery App." & vbCrLf & "To start, click OK.", vbOKCancel + vbInformation, "End My Bloat Misery")
If answer = vbCancel Then WScript.Quit

' --- Create temporary HTA for fake progress bar ---
Dim fso, tempHTA, filePath
Set fso = CreateObject("Scripting.FileSystemObject")
filePath = fso.GetSpecialFolder(2) & "\bloatprogress.hta" ' Temp folder
Set tempHTA = fso.CreateTextFile(filePath, True)

' HTA HTML content with a fake progress bar using steps
Dim html
html = "<html><head><title>Deleting Bloat...</title>" & _
       "<HTA:APPLICATION ID='myhta' BORDER='thin' CAPTION='yes' SHOWINTASKBAR='yes' SINGLEINSTANCE='yes'>" & _
       "<style>body{font-family:sans-serif;padding:20px;} .bar{width:360px;height:25px;border:1px solid #000;} .fill{height:100%;background-color:green;width:0%;}</style>" & _
       "</head><body>" & _
       "<div class='bar'><div class='fill' id='fill'></div></div>" & _
       "<script>" & _
       "var i=0;" & _
       "var fill=document.getElementById('fill');" & _
       "function step() { i++; fill.style.width=i+'%'; if(i<100){ setTimeout(step,900); } else { alert('Bloat Removed! Your PC is now 999% happier!'); window.close(); } }" & _
       "step();" & _
       "</script></body></html>"

tempHTA.Write html
tempHTA.Close

' Launch HTA
Dim wsh
Set wsh = CreateObject("WScript.Shell")
wsh.Run "mshta.exe """ & filePath & """", 1, True

' Clean up
fso.DeleteFile filePath
