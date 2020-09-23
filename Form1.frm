VERSION 5.00
Object = "{33D148A5-290B-4B3B-8984-753CDF4D7199}#24.0#0"; "MapViewer.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "MazeRunner"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd 
      Left            =   3360
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "*.sav|*.sav"
   End
   Begin MapViewer.MapView Map 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   3836
   End
   Begin VB.Menu save 
      Caption         =   "&Save"
   End
   Begin VB.Menu load 
      Caption         =   "&Load"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim CurrentMap As String
Dim FoundKey As Boolean, DoorOpened As Boolean, SecretLocation As Boolean
Dim StartW, StartH, MyX As Long, MyY As Long, DirectionX As String
Private Sub Form_Load()
Map.ApplicationPath = App.Path
Map.LoadMap App.Path & "\Maps\level01.txt"
Form2.Show , Me
End Sub

Private Sub Form_Resize()
Me.Height = StartH
Me.Width = StartW
End Sub

Private Sub Map_DirectionChange(Direction As MapViewer.DirectionsX)
Select Case Direction
Case 3
Map.SetCharacter LoadPicture(App.Path & "\textures\character\down.gif")
Case 1
Map.SetCharacter LoadPicture(App.Path & "\textures\character\up.gif")
Case 4
Map.SetCharacter LoadPicture(App.Path & "\textures\character\left.gif")
Case 2
Map.SetCharacter LoadPicture(App.Path & "\textures\character\right.gif")
End Select
DirectionX = Direction
End Sub

Private Sub Map_LoadError(Description As String)
PutLog "Load Error: " & Description
MsgBox "Unable to start Maze Runner: " & Description, vbCritical
End
End Sub

Public Function Prepare()
Map.SetCharacter LoadPicture(App.Path & "\textures\character\right.gif")
If CurrentMap = "level01" Then
 Map.ChangeTile 5, 4, App.Path & "\Textures\Switches\SwitchOn.bmp", 0
 Map.ChangeTile 7, 2, App.Path & "\Textures\warp2.bmp", 0
End If
If CurrentMap = "level02" Then
 FoundKey = False
 DoorOpened = False
 SecretLocation = False
End If
End Function

Private Sub Map_MapLoad(Map As String, FileName As String)
PutLog "Load: " & Map & "-" & VBA.Left(FileName, InStr(FileName, ".") - 1)
CurrentMap = VBA.Left(FileName, InStr(FileName, ".") - 1)
Prepare
SetCaption "You are on " & Replace(VBA.Left(FileName, InStr(FileName, ".") - 1), "level", "Level ") & "!"
End Sub

Public Function SetCaption(str As String)
Form2.Label1.Caption = str
End Function

Public Function PutLog(str As String)
Close #1
Open App.Path & "\debug.txt" For Append As #1
Print #1, ""
Print #1, "{ " & Now & " } " & str
Print #1, ""
Close #1
End Function

Private Sub Map_Move(X As Long, Y As Long, TileImage As stdole.Picture, tiletype As String)
MyX = X
MyY = Y
If CurrentMap = "level01" Then
Select Case X & "," & Y
Case "9,2"
 If FoundKey = False Then
 FoundKey = True
 Map.ChangeTile X, Y, App.Path & "\Textures\floor.bmp", 0
 SetCaption "You found a key!"
 End If
Case "6,6"
 If FoundKey = False And DoorOpened = fals Then
 SetCaption "You see a door. Something is needed."
 ElseIf DoorOpened = False And FoundKey = True Then
 SetCaption "The key fits the door."
 Map.ChangeTile 5, 6, App.Path & "\textures\floor.bmp", 0
 End If
Case "4,6"
 DoEvents
 Sleep 300
 DoEvents
 SetCaption "You were magically teleported to another room. No exits here."
 Map.MovePlayer 5, 2
 Map.SetCharacter LoadPicture(App.Path & "\textures\character\down.gif")
 SecretLocation = True
 FoundKey = False
Case "5,3"
 If SecretLocation = True And FoundKey = False Then
 SetCaption "The forcefield needs to be disabled."
 End If
Case "5,4"
 If SecretLocation = True And FoundKey = False Then
 FoundKey = True
 Map.ChangeTile 5, 4, App.Path & "\Textures\Switches\SwitchOff.bmp", 0
 Map.ChangeTile 6, 3, App.Path & "\Textures\Floor.bmp", 0
 SetCaption "Forcefield Disabled"
 Else
 FoundKey = False
 Map.ChangeTile 5, 4, App.Path & "\Textures\Switches\SwitchOn.bmp", 0
 Map.ChangeTile 6, 3, App.Path & "\Textures\Warp.bmp", 1
 SetCaption "Forcefield Enabled"
 End If
Case "7,2"
 Map.LoadMap App.Path & "\maps\level02.txt"
End Select
End If
If CurrentMap = "level02" Then
Select Case X & "," & Y
Case "9,2"
 FoundKey = True
 Map.ChangeTile X, Y, App.Path & "\Textures\floor.bmp", 0
 SetCaption "You found a key!"
Case "7,3"
 If FoundKey = False Then
 SetCaption "You need a key!"
 Else
 Map.ChangeTile X, Y + 1, App.Path & "\textures\floor.bmp", 0
 FoundKey = False
 SetCaption "The door opened!"
 End If
Case "9,5"
 Map.ChangeTile 8, 6, App.Path & "\textures\floor.bmp", 0
 SetCaption "You hit the door-open button!"
Case "5,7"
 Map.SetCharacter LoadPicture(App.Path & "\textures\character\up.gif")
 Sleep 100
 DoEvents
 Map.MovePlayer 5, 5
End Select
End If
End Sub

Private Sub Map_WindowDimensions(Width As Long, Height As Long)
Form1.Width = Width + 100
Form1.Height = Height + 450
StartW = Width + 100
StartH = Height + 850
End Sub

Private Sub save_Click()
cd.InitDir = App.Path & "\Saves\"
On Error GoTo err
cd.FileName = ""
cd.ShowSave
Dim SaveStr As String
Open cd.FileName For Output As #1
Print #1, MyX
Print #1, MyY
Print #1, Form2.Label1.Caption
Print #1, FoundKey
Print #1, DoorOpened
Print #1, SecretLocation
Print #1, CurrentMap
Print #1, DirectionX
Map.OutPutInfo SaveStr
Print #1, SaveStr
Close #1
err:
End Sub

Private Sub load_Click()
On Error GoTo err
cd.InitDir = App.Path & "\Saves\"
cd.FileName = ""
cd.ShowOpen
Map.PrepeareMap
Open cd.FileName For Input As #1
Dim str As String, tag As String, texture As String, xp As String, yp As String
Input #1, str
str = Replace(str, " ", "")
xp = str
Input #1, str
str = Replace(str, " ", "")
yp = str
Map.MovePlayer CLng(Val(xp)), CLng(Val(yp))
Line Input #1, str
SetCaption "Loaded - " & str
Line Input #1, str
FoundKey = CBool(str)
Line Input #1, str
DoorOpened = CBool(str)
Line Input #1, str
SecretLocation = CBool(str)
Line Input #1, str
CurrentMap = str
Line Input #1, str
DirectionX = Val(str)
Call Map_DirectionChange(CStr(DirectionX))
Line Input #1, str
Line Input #1, str
Line Input #1, str
Line Input #1, str
While Not EOF(1)
Line Input #1, tag
Line Input #1, texture
Line Input #1, xp
Line Input #1, yp
xp = Val(xp)
yp = Val(yp)
tag = Val(tag)
Map.ChangeTile CLng(xp), CLng(yp), Replace(texture, ".\", App.Path & "\"), tag
Wend
Close #1
err:
End Sub
