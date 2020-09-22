VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00DBE3E2&
   BorderStyle     =   0  'None
   Caption         =   "PipeMan"
   ClientHeight    =   6210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8265
   Icon            =   "MainForm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   8265
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0CFD0&
      Height          =   315
      Left            =   2760
      ScaleHeight     =   255
      ScaleWidth      =   5265
      TabIndex        =   25
      Top             =   360
      Width           =   5325
      Begin VB.Label lblGameInfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00212D2E&
         Height          =   255
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   5265
      End
   End
   Begin VB.CommandButton cmdScore 
      BackColor       =   &H00C0CFD0&
      Caption         =   "&Hiscore"
      Height          =   495
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton cmdJump 
      BackColor       =   &H00C0CFD0&
      Caption         =   "&Jump to level"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4680
      Width           =   975
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2280
      Top             =   3240
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00DBE3E2&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Enabled         =   0   'False
      Height          =   1815
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1935
      Begin VB.TextBox txtPipesLeft 
         BackColor       =   &H00EEF1F2&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "0"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtElapsed 
         BackColor       =   &H00EEF1F2&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "0"
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtMultiply 
         BackColor       =   &H00EEF1F2&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "0"
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtTrigTime 
         BackColor       =   &H00EEF1F2&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "0"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtScore 
         BackColor       =   &H00EEF1F2&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "0"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtBonus 
         BackColor       =   &H00EEF1F2&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "0"
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Pipes to win"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   20
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Elapsed"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Multiply"
         Height          =   255
         Left            =   1080
         TabIndex        =   16
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Trigger time"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Score"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Bonus"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.PictureBox picAppBar 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   30
      ScaleHeight     =   300
      ScaleWidth      =   8205
      TabIndex        =   7
      Top             =   30
      Width           =   8205
      Begin VB.Image picControl 
         Height          =   270
         Index           =   0
         Left            =   7635
         Tag             =   "Minimize"
         Top             =   15
         Width           =   270
      End
      Begin VB.Image picControl 
         Height          =   270
         Index           =   1
         Left            =   7920
         Tag             =   "Exit"
         Top             =   15
         Width           =   270
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SAVELOGO"
      Height          =   495
      Left            =   2160
      TabIndex        =   6
      Top             =   4080
      Width           =   615
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSpeed 
      BackColor       =   &H00C0CFD0&
      Caption         =   "Speed &up"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdPause 
      BackColor       =   &H00C0CFD0&
      Caption         =   "&Pause"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4080
      Width           =   975
   End
   Begin VB.PictureBox picBoxes 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0CFD0&
      Height          =   2460
      Left            =   2160
      ScaleHeight     =   2400
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   720
      Width           =   540
   End
   Begin VB.PictureBox picField 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0CFD0&
      Height          =   5340
      Left            =   2760
      ScaleHeight     =   5280
      ScaleWidth      =   5280
      TabIndex        =   0
      Top             =   720
      Width           =   5340
      Begin VB.PictureBox picPause 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0CFD0&
         BorderStyle     =   0  'None
         Height          =   5280
         Left            =   0
         ScaleHeight     =   5280
         ScaleWidth      =   5280
         TabIndex        =   27
         Top             =   0
         Width           =   5280
         Visible         =   0   'False
      End
   End
   Begin VB.CommandButton cmdStop 
      BackColor       =   &H00C0CFD0&
      Caption         =   "S&top"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00C0CFD0&
      Caption         =   "&Start"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   975
   End
   Begin VB.OptionButton optSpeed 
      BackColor       =   &H00DBE3E2&
      Caption         =   "Sp&eed pipe"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1080
      TabIndex        =   24
      Top             =   5280
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.OptionButton optCheck 
      BackColor       =   &H00DBE3E2&
      Caption         =   "&Check pipe"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   23
      Top             =   5280
      Width           =   975
   End
   Begin VB.Image LogoLine 
      Height          =   90
      Left            =   30
      Stretch         =   -1  'True
      Top             =   1305
      Width           =   2130
   End
   Begin VB.Image LogoImg 
      Height          =   660
      Left            =   240
      Top             =   480
      Width           =   1800
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdPause_Click()
  PauseGame
End Sub

Private Sub cmdScore_Click()
  ShowHiscore False
End Sub

Private Sub cmdSpeed_Click()
  If TriggerTime > 0 Then
    Bonus = Bonus + (Fix(TriggerTime / 1.25) * 5) * BonusMulti
    TriggerTime = 0
    UpdateBoard
  Else
    Speed = Speed / 2
    
    If SpeedMulti = False Then
      BonusMulti = BonusMulti + 1
      UpdateBoard
      
      SpeedMulti = True
      End If
    End If
End Sub

Private Sub cmdStart_Click()
  Dim Continue As Boolean
  
  cmdStart.Enabled = False
  cmdStop.Enabled = True
  cmdPause.Enabled = True
  cmdSpeed.Enabled = True
  cmdScore.Enabled = False
  optSpeed.Enabled = False
  
  ElapsedTime = MSTimer
  Timer.Enabled = True
  Force = False
  
  Score = 0
  Bonus = 0
  CurrentMap = 0
  NextMap
  
  Do
    TimeLoop
    WaterLoop
    
    Score = Score + Int(Bonus)
    UpdateBoard
    EmptyBoxes
    
    Continue = False
    
    If Force = False Then
      If WinPipe = 0 Then
        PauseLoop 2000
        ShowPicture NextMapPic
        PauseLoop 2000
        NextMap
        Continue = True
        End If
      End If
    
  Loop Until Continue = False Or Force
  
  StopGame
  
  Timer.Enabled = False
  
  PauseLoop 2000
  ShowPicture GameOverPic
  UpdateBoard
  
  Dim Position As Integer
  
  Position = NewScore(Score)
  
  If Position > 0 Then
    PauseLoop 1000
    HighScore(Position).Score = Score
    
    ShowHiscore True
    
    EnterName = ""
    EnterX = 0
    EnterY = Position
    DrawBox EnterX * PipeSize, EnterY * PipeSize, 100
    EnterHiscore
    EnterLoop
    
    HighScore(Position).Player = EnterName
    
    SaveHiscore "hiscore.fdi"
    
    ShowHiscore False
  Else
    cmdScore.Enabled = True
    End If
  
  optSpeed.Enabled = True
  cmdStart.Enabled = True
End Sub

Private Sub cmdStop_Click()
  StopGame
End Sub

Private Sub Command2_Click()
  Dim asd As New PipeMan.Write
  asd.SaveFile "pipelogo.fdi", False
  
  asd.WriteString "PMPL", 4
  
  A = "0000000000007474714730020222222000241122213006522222630000000000000742273734002222201150022221326400656563202000000000000"
  For X = 0 To 10
  For Y = 0 To 10
    asd.WriteValue Mid(A, Y * 11 + X + 1, 1), 4
  Next
  Next
  A = "0000000000074273202313261206150202021371402020263202020000000000000714734734002222021350022213120000222202200000000000000"
  For X = 0 To 10
  For Y = 0 To 10
    asd.WriteValue Mid(A, Y * 11 + X + 1, 1), 4
  Next
  Next
  A = "00000000000" & _
      "07342734730" & _
      "01352135200" & _
      "02002200130" & _
      "02002200630" & _
      "00000000000" & _
      "07147347420" & _
      "02222022610" & _
      "02221312020" & _
      "02222022020" & _
      "00000000000"
  For X = 0 To 10
  For Y = 0 To 10
    asd.WriteValue Mid(A, Y * 11 + X + 1, 1), 4
  Next
  Next
  
  A = "00000000000" & _
      "00000000000" & _
      "00000000000" & _
      "73474227373" & _
      "13522226413" & _
      "20011220220" & _
      "20022653563" & _
      "00000000000" & _
      "00000000000" & _
      "00000000000" & _
      "00000000000"
  For X = 0 To 10
  For Y = 0 To 10
    asd.WriteValue Mid(A, Y * 11 + X + 1, 1), 4
  Next
  Next
  asd.CloseFile False
End Sub

Private Sub Command3_Click()
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If (Shift Mod 2 ^ (3 - 1) <> Shift Mod 2 ^ 3) And KeyCode = 115 Then
    End
    End If
  
  If GameOver = True And KeyCode = 27 And Shift = 0 Then
    PauseGame True
    frmMain.WindowState = 1
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If EnterMode = True Then
    Dim Value As Integer
    Dim PipeType As Integer
    
    Value = Asc(UCase(Chr(KeyAscii)))

    If Value = 32 Then
      PipeType = 101
    ElseIf Value >= 65 And Value <= 90 Then
      PipeType = Value + 235
    ElseIf Value = 8 And EnterX > 0 Then
      EnterX = EnterX - 1
      EnterName = Left(EnterName, Len(EnterName) - 1)
      PipeType = 100
    ElseIf Value = 13 Then
      EnterMode = False
      End If
    
    If PipeType > 0 And Len(EnterName) < 6 Then
      DrawBox EnterX * PipeSize, EnterY * PipeSize, PipeType - (PipeType = 100)
      
      If Len(EnterName) + (PipeType = 100) < 5 Then
        DrawBox (EnterX + 1 + (PipeType = 100)) * PipeSize, EnterY * PipeSize, 100
        End If
      If PipeType = 100 And Len(EnterName) < 5 Then
        DrawBox (EnterX + 1) * PipeSize, EnterY * PipeSize, 101
        End If
        
      If PipeType <> 100 Then
        EnterX = EnterX + 1
        EnterName = EnterName & Chr(Value)
        End If
      End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  End
End Sub

Private Sub picAppBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    MX = X
    MY = Y
    End If
End Sub

Private Sub picAppBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    frmMain.Move frmMain.Left + X - MX, frmMain.Top + Y - MY
    End If
End Sub

Private Sub picControl_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    picControl(Index).Picture = LoadResPicture("DOWN" & picControl(Index).Tag, 0)
    End If
End Sub

Private Sub picControl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    If X >= 0 And Y >= 0 And X <= 270 And Y <= 270 Then
      picControl(Index).Picture = LoadResPicture("DOWN" & picControl(Index).Tag, 0)
    Else
      picControl(Index).Picture = LoadResPicture("UP" & picControl(Index).Tag, 0)
      End If
    End If
End Sub

Private Sub picControl_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 And X >= 0 And Y >= 0 And X <= 270 And Y <= 270 Then
      Select Case Index
        Case 0
          PauseGame True
          frmMain.WindowState = 1
        Case 1
          End
      End Select
    End If
  
  picControl(Index).Picture = LoadResPicture("UP" & picControl(Index).Tag, 0)
End Sub

Private Sub picField_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 And GameOver = False And Pause = False Then
      PlaceBox CInt(X), CInt(Y)
      End If
End Sub

Private Sub Timer_Timer()
  frmMain.txtElapsed = Int((MSTimer - ElapsedTime) / 1000)
End Sub
