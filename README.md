<div align="center">

## Personal StartUp\! As simple as it can get\.\.\.


</div>

### Description

Do you use different user logins in Windows?

Tired of having the same StartUp folder just because the users share the same Start Menu?

Then this solves your problem!
 
### More Info
 
How to use personal startups.

1) Add the source code below to a new project.

2) Make an exe file.

3) Create a shortcut in your StartUp folder that points to the exe.

4) Log Out and Log In again as the person you want to give personal startup.

5) Check in the directory where you put the exe.

6) Open the file UserID.txt with notepad (Replace UserID with yor UserID)

7) Add the path of the files you want to start when windows starts, separate the paths of the different files by hitting enter.

8) Done!

Starts different programs on startup depending on the who has logged in.

None that i know of.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Hyperswede](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/hyperswede.md)
**Level**          |Unknown
**User Rating**    |5.9 (606 globes from 102 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/hyperswede-personal-startup-as-simple-as-it-can-get__1-2047/archive/master.zip)

### API Declarations

```
Declare Function WNetGetUser Lib "mpr" _
  Alias "WNetGetUserA" (ByVal lpName As _
  String, ByVal lpUserName As String, _
  lpnLength As Long) As Long
Public Function UserID() As String
  Dim sUserNameBuffer As String * 255
  sUserNameBuffer = Space(255)
  Call WNetGetUser(vbNullString, _
    sUserNameBuffer, 255&)
    UserID = Left$(sUserNameBuffer, _
      InStr(sUserNameBuffer, _
      vbNullChar) - 1)
If UserID = "" Then UserID = "default"
End Function
```


### Source Code

```
Private Sub Form_Load()
Dim OpenWhat
'MsgBox UserID
On Error GoTo bwell
Open App.Path & "\" & UserID & ".txt" For Input As #1
On Error Resume Next
Do Until EOF(1)
Line Input #1, OpenWhat
Shell "Start " & OpenWhat
Loop
Close #1
End
bwell:
Open App.Path & "\" & UserID & ".txt" For Output As #2: Close #2
Resume
End Sub
```

