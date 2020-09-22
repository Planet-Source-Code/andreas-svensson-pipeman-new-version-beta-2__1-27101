Attribute VB_Name = "modWater"
Sub AdvWater()
  Select Case WaterPipe
    Case 1
      If WaterWater And WaterPos.X = 16 - WaterOrient.X And WaterPos.Y = 16 - WaterOrient.Y Then
        Bonus = Bonus + 10 * BonusMulti
        UpdateBoard
        End If
      
      ContWater
      DrawWater
      
    Case 2, 3
      ContWater
      DrawWater
      
    Case 8, 9, 10, 11
      If WaterWater And WaterPos.X = 16 - WaterOrient.X And WaterPos.Y = 16 - WaterOrient.Y Then
        Bonus = Bonus + 5 * BonusMulti
        UpdateBoard
        End If
      
      ContWater
      DrawWater
    
    Case 4, 5, 6, 7
      Dim X As Integer
      Dim Y As Integer
      
      If WaterOrient.X > 0 Then
        X = 3
      ElseIf WaterOrient.X < 0 Then
        X = -2
      ElseIf WaterOrient.Y > 0 Then
        Y = 3
      ElseIf WaterOrient.Y < 0 Then
        Y = -2
        End If
      
      If WaterPos.X = 16 + X And WaterPos.Y = 16 + Y Then
        
        Select Case WaterPipe
          Case 4
            OrientWater -1, 1
          Case 5
            OrientWater -1, -1
          Case 6
            OrientWater 1, -1
          Case 7
            OrientWater 1, 1
        End Select
        End If
        
      ContWater
      DrawWater
      
    Case 13
      If WaterPos.X > 10 - WaterOrient.X And WaterPos.X < 23 - WaterOrient.X And _
         WaterPos.Y > 10 - WaterOrient.Y And WaterPos.Y < 23 - WaterOrient.Y Then
        
        WaterAnim = -1
        If WaterPos.X = 16 - WaterOrient.X And WaterPos.Y = 16 - WaterOrient.Y Then
          Bonus = Bonus + 15 * BonusMulti
          UpdateBoard
          End If
        End If
      
      ContWater
      DrawWater
    
    Case 12
      If WaterPos.X > 11 - WaterOrient.X And WaterPos.X < 22 - WaterOrient.X And _
         WaterPos.Y > 11 - WaterOrient.Y And WaterPos.Y < 22 - WaterOrient.Y Then
        
        WaterAnim = -1
        If WaterPos.X = 16 - WaterOrient.X And WaterPos.Y = 16 - WaterOrient.Y Then
          For P = 1 To 2
            If Teleport(P).X = WaterField.X And Teleport(P).Y = WaterField.Y Then
              WaterField.X = Teleport(3 - P).X
              WaterField.Y = Teleport(3 - P).Y
              Exit For
              End If
          Next
          Bonus = Bonus + 5 * BonusMulti
          
          If WaterWater Then
            Bonus = Bonus + 20 * BonusMulti
            End If
          
          UpdateBoard
          End If
        End If
      
      ContWater
      DrawWater
  End Select
End Sub

Sub OrientWater(X As Integer, Y As Integer)
  If WaterOrient.X = -X Then
    WaterOrient.X = 0
    WaterOrient.Y = Y
    WaterPos.X = 16
    WaterPos.Y = 16 + Y * 2
  
  ElseIf WaterOrient.Y = -Y Then
    WaterOrient.Y = 0
    WaterOrient.X = X
    WaterPos.X = 16 + X * 2
    WaterPos.Y = 16
    End If
End Sub

Function CheckWater(X As Integer, Y As Integer) As Boolean
  If WaterOrient.X = -X Then
    CheckWater = True
  ElseIf WaterOrient.Y = -Y Then
    CheckWater = True
    End If
End Function

Function CheckOrient(X As Integer, Y As Integer) As Boolean
  If WaterOrient.X = X And WaterOrient.Y = Y Then
    CheckOrient = True
    End If
End Function

Sub CalcWater()
  Dim BAD As Boolean
  Dim OK As Boolean
  
  If WaterField.X < 0 Or WaterField.X > 10 Or _
     WaterField.Y < 0 Or WaterField.Y > 10 Then
     BAD = True
    End If
  
  If BAD = False Then
    If WaterWater Then
      CrossPoint = True
      End If
    
    WaterWater = Field(WaterField.X, WaterField.Y).Watered
    WaterPipe = Field(WaterField.X, WaterField.Y).PipeType
    Field(WaterField.X, WaterField.Y).Watered = True
    
    Select Case WaterPipe
      Case 1, 12, 13
        OK = True
        
      Case 2
        If WaterOrient.X = 0 Then
          OK = True
          End If
        
      Case 3
        If WaterOrient.Y = 0 Then
          OK = True
          End If
        
      Case 4, 5, 6, 7
        Select Case WaterPipe
          Case 4
            OK = CheckWater(-1, 1)
          Case 5
            OK = CheckWater(-1, -1)
          Case 6
            OK = CheckWater(1, -1)
          Case 7
            OK = CheckWater(1, 1)
        End Select
      
      Case 8
        OK = CheckOrient(-1, 0)
      
      Case 9
        OK = CheckOrient(0, -1)
      
      Case 10
        OK = CheckOrient(1, 0)
      
      Case 11
        OK = CheckOrient(0, 1)
      
    End Select
    End If
  
  If OK Then
    Score = Score + 10
    
    If WinPipe > 0 Then
      WinPipe = WinPipe - 1
      End If
    
    UpdateBoard
  Else
    WaterAnim = -1
    GameOver = True
    End If
End Sub

Sub DrawWater()
  Dim WaterPic As Integer
  Dim X As Integer
  Dim Y As Integer
  
  If WaterOrient.X <> 0 Then
    WaterPic = 200
    Y = Y - 2
  ElseIf WaterOrient.Y <> 0 Then
    WaterPic = 201
    X = X - 2
    End If
  
  X = (X + WaterPos.X + WaterField.X * 32 - 1) * Screen.TwipsPerPixelX
  Y = (Y + WaterPos.Y + WaterField.Y * 32 - 1) * Screen.TwipsPerPixelY
  
  If WaterAnim = 0 Then
    frmMain.picField.PaintPicture LoadResPicture(WaterPic, 0), X, Y
  Else
    Select Case WaterPipe
      Case 1, 4, 5, 6, 7
        Dim XX As Integer
        Dim YY As Integer
        Dim XW As Integer
        Dim YH As Integer
        
        If WaterOrient.X <> 0 Then
          XX = (WaterPos.X - 14) * Screen.TwipsPerPixelX
          YH = 6 * Screen.TwipsPerPixelY
          XW = Screen.TwipsPerPixelX
        ElseIf WaterOrient.Y <> 0 Then
          YY = (WaterPos.Y - 14) * Screen.TwipsPerPixelY
          XW = 6 * Screen.TwipsPerPixelX
          YH = Screen.TwipsPerPixelY
          End If
        
        If WaterAnim <> -1 Then
          frmMain.picField.PaintPicture LoadResPicture(WaterAnim, 0), X, Y, , , XX, YY, XW, YH
          End If
    End Select
    End If
  
  WaterAnim = 0
End Sub

Sub ContWater()
  WaterPos.X = WaterPos.X + WaterOrient.X
  WaterPos.Y = WaterPos.Y + WaterOrient.Y
  
  If WaterPos.X < 1 Then
    WaterPos.X = 32
    WaterField.X = WaterField.X - 1
    CalcWater
  ElseIf WaterPos.X > 32 Then
    WaterPos.X = 1
    WaterField.X = WaterField.X + 1
    CalcWater
  ElseIf WaterPos.Y < 1 Then
    WaterPos.Y = 32
    WaterField.Y = WaterField.Y - 1
    CalcWater
  ElseIf WaterPos.Y > 32 Then
    WaterPos.Y = 1
    WaterField.Y = WaterField.Y + 1
    CalcWater
    End If
  
  Select Case WaterPipe
    Case 1
    If WaterPos.X > 13 And WaterPos.X < 20 And _
       WaterPos.Y > 13 And WaterPos.Y < 20 Then
      
      If WaterWater Then
        WaterAnim = 202
        End If
      End If
    
    Case 4, 5, 6, 7
      If WaterPos.X > 13 And WaterPos.X < 20 And _
         WaterPos.Y > 13 And WaterPos.Y < 20 Then
        
        WaterAnim = 199 + WaterPipe
        End If
  End Select
End Sub
