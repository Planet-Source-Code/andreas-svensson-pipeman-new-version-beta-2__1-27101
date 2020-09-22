Attribute VB_Name = "modOther"
Sub DrawBoxes()
  Dim Y As Integer
  
  For Y = 1 To 5
    frmMain.picBoxes.PaintPicture LoadResPicture(Boxes(6 - Y), 0), 0, (Y - 1) * 480
  Next
  
  'frmMain.picCurBox.PaintPicture LoadResPicture(Boxes(0), 0), 0, 0
End Sub

Sub AdvBoxes()
  Dim P As Integer
  
  For P = 1 To 4
    Boxes(P) = Boxes(P + 1)
  Next
  
  Boxes(5) = Int(Rnd * 7) + 1
  
  DrawBoxes
End Sub

Sub PlaceBox(X As Integer, Y As Integer, Optional BoxType As Integer = -1)
  Dim XX As Integer
  Dim YY As Integer
  
  XX = Int(X / 480)
  YY = Int(Y / 480)
  
  If BoxType = -1 Then
    BoxType = Boxes(1)
    If Field(XX, YY).PipeType <> 100 Then
      PauseLoop 200
      End If
    End If
  
  If Field(XX, YY).Watered = False And (Field(XX, YY).PipeType = 100 Or Field(XX, YY).PipeType < 8) Then
    Field(XX, YY).PipeType = BoxType
    DrawBox Int(X / PipeSize) * PipeSize, Int(Y / PipeSize) * PipeSize, BoxType
    
    AdvBoxes
    DrawBoxes
    End If
End Sub

Sub DrawBox(X As Integer, Y As Integer, PipeType As Integer)
  frmMain.picField.PaintPicture LoadResPicture(PipeType, 0), X, Y
End Sub

Sub GenBoxes()
  For P = 1 To 5
    Boxes(P) = Int(Rnd * 7) + 1
  Next
  
  DrawBoxes
End Sub

Sub EmptyBoxes()
  For P = 1 To 5
    Boxes(P) = 101
  Next
  
  DrawBoxes
End Sub

Sub PlaceStart(X As Integer, Y As Integer, PipeType As Integer)
  Dim OrientX As Integer
  Dim OrientY As Integer
  Dim XPlus As Integer
  Dim YPlus As Integer
  
  Field(X, Y).PipeType = PipeType
  Field(X, Y).Watered = True
  
  DrawBox X * PipeSize, Y * PipeSize, PipeType
  
  Select Case PipeType
    Case 102
      OrientY = -1
    Case 103
      OrientX = 1
      XPlus = 1
    Case 104
      OrientY = 1
      YPlus = 1
    Case 105
      OrientX = -1
  End Select
  
  WaterPos.X = 16 + OrientX * 6 + XPlus
  WaterPos.Y = 16 + OrientY * 6 + YPlus
  WaterPipe = 1
  WaterField.X = X
  WaterField.Y = Y
  WaterOrient.X = OrientX
  WaterOrient.Y = OrientY
End Sub

Sub ShowPicture(List() As Byte, Optional Quick As Boolean)
  ResetField
  
  Dim P As Integer
  
  For Y = 0 To 10
    For X = 0 To 10
      If Quick = False Then
        DoEvents
        PauseLoop 10
        End If
      DrawBox X * PipeSize, Y * PipeSize, CInt(List(X, Y))
      Field(X, Y).PipeType = List(X, Y)
    Next
  Next
  
  DrawField
End Sub

Sub ReadPics(List() As Byte, Filename As String, Index As Integer)
  Dim FPRead As New Read
  Dim P As Integer
  Dim X As Integer
  Dim Y As Integer
  
  If PLID = "" Then
    FPRead.LoadFile Filename
    PLID = FPRead.ReadString(4)
    FPRead.CloseFile
    End If
  
  If PLID = "PMPL" Then
    FPRead.LoadFile Filename
    FPRead.SeekFile (Index * 121) * 4 + 32
    
    For X = 0 To 10
      For Y = 0 To 10
        P = FPRead.ReadValue(4)
        
        If P = 0 Then
          P = 100
          End If
        List(X, Y) = P
      Next
    Next
    
    FPRead.CloseFile
  End If
End Sub

Sub UpdateBoard()
  frmMain.txtTrigTime = TriggerTime
  frmMain.txtScore = Score
  frmMain.txtPipesLeft = WinPipe
  frmMain.txtBonus = Int(Bonus)
  frmMain.txtMultiply = BonusMulti
End Sub

Function NewScore(Score As Long) As Integer
  Dim Pos1 As Integer
  Dim Pos2 As Integer
  Dim Position As Integer
  
  For Pos1 = 1 To UBound(HighScore)
    If Score > HighScore(Pos1).Score Then
    
      For Pos2 = UBound(HighScore) - 1 To Pos1 Step -1
        HighScore(Pos2 + 1).Score = HighScore(Pos2).Score
        HighScore(Pos2 + 1).Player = HighScore(Pos2).Player
      Next
      
      HighScore(Pos1).Score = 0
      HighScore(Pos1).Player = ""
      Position = Pos1
      Exit For
      End If
  Next
  
  NewScore = Position
End Function

Sub PauseGame(Optional State As Integer = 1)
  If State = 1 Then
    Pause = Not Pause
  Else
    Pause = State
  End If
  
  frmMain.picPause.Visible = Pause
End Sub

Sub StopGame()
  PauseGame False
  GameOver = True
  Force = True
  
  frmMain.cmdStart.Enabled = False
  frmMain.cmdStop.Enabled = False
  frmMain.cmdPause.Enabled = False
  frmMain.cmdSpeed.Enabled = False
  frmMain.cmdScore.Enabled = False
End Sub

Sub SaveHiscore(Filename As String)
  Dim FPWrite As New PipeMan.Write
  Dim P As Integer
  
  FPWrite.SaveFile Filename, False
  
  FPWrite.WriteString "PMHI", 4
  
  For P = 1 To 10
    FPWrite.WriteTable HighScore(P).Player, Tables(), 6
    FPWrite.WriteValue CCur(HighScore(P).Score), 17
  Next
  
  FPWrite.CloseFile
End Sub

Sub LoadHiscore(Filename As String)
  Dim FPRead As New PipeMan.Read
  Dim Inp As String
  Dim P As Integer
  Dim Pos As Integer
  Dim Score As Long
  Dim Player As String
  
  FPRead.LoadFile Filename
  Inp = FPRead.ReadString(4)
  
  If Inp = "PMHI" Then
    For P = 1 To 10
      Player = FPRead.ReadTable(Tables(), 6)
      Score = FPRead.ReadValue(17)
      
      Pos = NewScore(Score)
      
      If Pos > 0 Then
        HighScore(Pos).Score = Score
        HighScore(Pos).Player = Player
        End If
    Next
    
    FPRead.CloseFile
    End If
End Sub

Sub CreateTable(Data As String, Table() As Byte)
  Dim Pos As Integer
  
  ReDim Table(0 To Len(Data) - 1)
  
  For Pos = 1 To Len(Data)
    Table(Pos - 1) = Asc(Mid(Data, Pos, 1))
  Next
End Sub

Sub EnterHiscore()
  frmMain.cmdJump.Enabled = False
  frmMain.cmdPause.Enabled = False
  frmMain.cmdScore.Enabled = False
  frmMain.cmdSpeed.Enabled = False
  frmMain.cmdStart.Enabled = False
  frmMain.cmdStop.Enabled = False
  frmMain.optSpeed.Enabled = False
  
  EnterMode = True
End Sub

Function PChar(Char As String) As Integer
  PChar = Asc(Char) + 235
End Function

Sub DrawText(Text As String, PX As Integer, PY As Integer)
  For Y = 1 To Len(Text)
    Letter = PChar(Mid(Text, Y, 1))
    
    If Letter = 235 Or Letter = 267 Then
      Letter = 101
      End If
    
    Field(PY - 1 + Y, PX).PipeType = Letter
  Next
End Sub

Sub DrawNumber(Value As Long, Length As Integer, PX As Integer, PY As Integer)
  For Y = 1 To Length
      Field(PY + Y - 1, PX).PipeType = Asc(Mid(Format(Value, String(Length, "0")), Y, 1)) + 302
    Next
End Sub

Sub ShowHiscore(Optional TypeName As Boolean)
  ResetField
  
  frmMain.cmdScore.Enabled = False
  If TypeName Then
    DrawText "YOUR NAME", 0, 1
  Else
    DrawText "HISCORE", 0, 2
    End If
  
  Dim Letter As Integer
  
  For X = 1 To 10
    DrawText HighScore(X).Player, CInt(X), 0
    DrawNumber HighScore(X).Score, 5, CInt(X), 6
  Next
  DrawField
End Sub
