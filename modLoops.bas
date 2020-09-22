Attribute VB_Name = "modLoops"
Sub WaterLoop()
  Dim NextFrame As Long
  Dim Temp As Long
  
  Do Until GameOver
    NextFrame = MSTimer + Speed - Temp
    AdvWater
    
    Do
      Do
        DoEvents
        
        If Pause Then
          NextFrame = MSTimer
          End If
      Loop While Pause
    Loop Until MSTimer >= NextFrame
      
    Temp = (MSTimer - NextFrame)
  Loop
End Sub

Sub TimeLoop()
  Dim NextFrame As Long
  
  Do Until GameOver Or TriggerTime = 0
    NextFrame = MSTimer + 1000
    TriggerTime = TriggerTime - 1
    
    UpdateBoard
    
    Do
      Do
        DoEvents
      Loop While Pause
    Loop Until MSTimer >= NextFrame Or TriggerTime = 0
  Loop
End Sub

Sub PauseLoop(Time As Long)
  Dim NextFrame As Long
  
  NextFrame = MSTimer + Time
  
  Do
    DoEvents
  Loop Until MSTimer >= NextFrame
End Sub

Sub EnterLoop()
  Do
    DoEvents
  Loop While EnterMode
End Sub
