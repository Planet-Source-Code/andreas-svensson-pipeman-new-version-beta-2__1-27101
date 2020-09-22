Attribute VB_Name = "modField"
Sub DrawField()
  Dim X As Integer
  Dim Y As Integer
  Dim P As Integer
  
  For X = 0 To 10
    For Y = 0 To 10
      P = Field(X, Y).PipeType
      frmMain.picField.PaintPicture LoadResPicture(P, 0), X * PipeSize, Y * PipeSize
    Next
  Next
End Sub

Sub ResetField()
  Dim X As Integer
  Dim Y As Integer
  
  For X = 0 To 10
    For Y = 0 To 10
      Field(X, Y).Watered = False
      Field(X, Y).PipeType = 100
    Next
  Next
End Sub
