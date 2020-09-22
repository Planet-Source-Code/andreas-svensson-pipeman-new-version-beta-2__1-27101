Attribute VB_Name = "modMap"
Sub NextMap()
  CurrentMap = CurrentMap + 1
  
  GameOver = False
  FireTime = False
  SpeedMulti = False
  
  Bonus = 0
  BonusMulti = 1
  UpdateBoard
  
  ResetField
  GenBoxes
  
  Select Case CurrentMap
    Case 1
      PlaceStart 4, 4, 103
      TriggerTime = 31
      WinPipe = 10
      Speed = 50
      frmMain.lblGameInfo.Caption = "Beta map"
      Field(3, 3).PipeType = 13
      Field(6, 6).PipeType = 8
      Field(7, 9).PipeType = 9
      Field(6, 4).PipeType = 10
      Field(8, 4).PipeType = 11
      Field(7, 7).PipeType = 101
      
      Field(5, 5).PipeType = 12
      Teleport(1).X = 5
      Teleport(1).Y = 5
      Field(8, 8).PipeType = 12
      Teleport(2).X = 8
      Teleport(2).Y = 8
    Case 2
      frmMain.lblGameInfo.Caption = "Old and bad map"
      PlaceStart Int(Rnd * 9) + 1, Int(Rnd * 9) + 1, 102 + Int(Rnd * 4)
      TriggerTime = 26
      WinPipe = 20
      Speed = 50
    Case 3
      frmMain.lblGameInfo.Caption = "Old and bad map"
      PlaceStart Int(Rnd * 9) + 1, Int(Rnd * 9) + 1, 102 + Int(Rnd * 4)
      TriggerTime = 21
      WinPipe = 25
      Speed = 40
    Case 4
      frmMain.lblGameInfo.Caption = "Old and bad map"
      PlaceStart Int(Rnd * 9) + 1, Int(Rnd * 9) + 1, 102 + Int(Rnd * 4)
      TriggerTime = 11
      WinPipe = 30
      Speed = 40
    Case 5
      frmMain.lblGameInfo.Caption = "Old and bad map"
      PlaceStart Int(Rnd * 9) + 1, Int(Rnd * 9) + 1, 102 + Int(Rnd * 4)
      TriggerTime = 6
      WinPipe = 10
      Speed = 50
    Case 6
      frmMain.lblGameInfo.Caption = "Old and bad map"
      For X = 1 To 10 Step 4
        For Y = 1 To 10 Step 4
          Field(X, Y).PipeType = 101
          Field(X, Y).Watered = True
        Next
      Next
      
      PlaceStart 2, 2, 102 + Int(Rnd * 4)
      TriggerTime = 31
      WinPipe = 35
      Speed = 35
    Case 7
      frmMain.lblGameInfo.Caption = "Old and bad map"
      PlaceStart Int(Rnd * 9) + 1, Int(Rnd * 9) + 1, 102 + Int(Rnd * 4)
      TriggerTime = 31
      WinPipe = 35
      Speed = 35
    Case 8
      frmMain.lblGameInfo.Caption = "Old and bad map"
      PlaceStart Int(Rnd * 9) + 1, Int(Rnd * 9) + 1, 102 + Int(Rnd * 4)
      TriggerTime = 41
      WinPipe = 50
      Speed = 20
    Case 9
      frmMain.lblGameInfo.Caption = "Old and bad map"
      PlaceStart Int(Rnd * 9) + 1, Int(Rnd * 9) + 1, 102 + Int(Rnd * 4)
      TriggerTime = 61
      WinPipe = 50
      Speed = 10
    Case 10
      frmMain.lblGameInfo.Caption = "Old and bad map"
      For X = 1 To 10 Step 2
        For Y = 1 To 10 Step 2
          Field(X, Y).PipeType = 101
          Field(X, Y).Watered = True
        Next
      Next
      
      PlaceStart 5, 5, 102 + Int(Rnd * 4)
      
      TriggerTime = 41
      WinPipe = 60
      Speed = 25
    Case 11
      frmMain.lblGameInfo.Caption = "You have made it, incredible"
      Force = True
      GameOver = True
  End Select
  
  DrawField
End Sub
