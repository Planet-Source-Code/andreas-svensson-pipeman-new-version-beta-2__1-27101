VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Write"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private CMN As New Common
Private BytePos As Byte
Private PosMulti As Byte
Private FileByte As Byte
Private Filenames As String
Public Filenumber As Integer

Property Get Position() As Currency
  If Filenumber >= 0 Then
    Position = (Seek(Filenumber) - 1) * 8 + BytePos
  Else
    Position = -1
    End If
End Property

Sub SaveFile(Filename As String, Optional Append As Boolean)
  If Filenumber >= 0 Then
    Err.Raise 3, , LoadResString(3): Exit Sub: End If
  
  Filenumber = FreeFile
  Filenames = Filename
  
  If Append = False And CMN.FileExists(Filename) Then
    Kill Filename
    End If
  
  Open Filename For Binary Access Write Lock Write As #Filenumber
  
  FileByte = 0
  PosMulti = 1
  BytePos = 0
End Sub

Sub CloseFile()
  If Filenumber <= -1 Then
    Err.Raise 2, , LoadResString(2): Exit Sub: End If
  
  If BytePos > 0 Then
    Put #Filenumber, , FileByte
    End If
  
  Close #Filenumber
  
  Filenumber = -1
  PosMulti = 0
  BytePos = 0
  FileByte = 0
  Filenames = ""
End Sub

Sub SeekFile(Position As Currency, Append As Boolean)
  If Filenumber <= -1 Then
    Err.Raise 2, , LoadResString(2): Exit Sub: End If
  
  Position = Int(Position)
  
  Seek #Filenumber, Int(Position / 8) + 1
  
  BytePos = Position Mod 8
  FileByte = 0
  PosMulti = 2 ^ (Position Mod 8)
  
  If Append Then
    Dim FPRead As New Read
    Dim Value As Integer
    
    FPRead.LoadFile Filenames
    FPRead.SeekFile Position - Position Mod 8
    Value = Position Mod 8
    If Value > 0 Then
      FileByte = FPRead.ReadValue(Value)
    Else
      FileByte = 0
      End If
    FPRead.CloseFile
    End If
End Sub

Sub WriteData(Bool As Boolean)
  If Filenumber <= -1 Then
    Err.Raise 2, , LoadResString(2): Exit Sub: End If
  
  BytePos = BytePos + 1
  
  If Bool Then
    FileByte = FileByte + PosMulti
    End If
  
  If BytePos = 8 Then
    Put #Filenumber, , FileByte
    FileByte = 0
    PosMulti = 1
    BytePos = 0
  Else
    PosMulti = PosMulti * 2
    End If
End Sub

Sub WriteHeader(Data As Currency, Length As Integer)
  If Filenumber <= -1 Then
    Err.Raise 2, , LoadResString(2): Exit Sub
  ElseIf Length <= 0 Then Err.Raise 4, , LoadResString(4): Exit Sub: End If
  
  Data = Int(Data)
  
  WriteValue Data, Length
End Sub

Sub WriteSpace(Length As Currency, Optional Header As Integer)
  If Filenumber <= -1 Then
    Err.Raise 2, , LoadResString(2): Exit Sub
  ElseIf Length < 0 Then Err.Raise 4, , LoadResString(4): Exit Sub
  ElseIf Header < 0 Then Err.Raise 7, , LoadResString(7): Exit Sub
  ElseIf Length = 0 And Header = 0 Then Err.Raise 8, , LoadResString(8): Exit Sub: End If
  
  Length = Int(Length)
  
  If Header > 0 Then
    WriteHeader Length, Header
    End If
  
  For Pos = 1 To Length
    WriteData False
  Next
End Sub

Sub WriteBooleans(Data() As Boolean, Optional Header As Integer)
  If Filenumber <= -1 Then
    Err.Raise 2, , LoadResString(2): Exit Sub
  ElseIf Header < 0 Then Err.Raise 7, , LoadResString(7): Exit Sub
  ElseIf CMN.ValidBArray(Data()) = False Then Err.Raise 9, , LoadResString(9): Exit Sub: End If
  
  Dim Pos As Long
  Dim Pos2 As Long
  
  If Header > 0 Then
    For Pos2 = UBound(Data()) To LBound(Data()) Step -1
      If Data(Pos2) Then Exit For
    Next
  Else
    Pos2 = UBound(Data()) - LBound(Data()) + 1
    End If
  
  Pos2 = Pos2 + LBound(Data()) - 1
  
  If Header > 0 Then
    WriteHeader CCur(Pos2), Header
    End If
  
  For Pos = LBound(Data()) To Pos2
    WriteData Data(Pos)
  Next
End Sub

Sub WriteValue(Data As Currency, Optional Length As Integer, Optional Header As Integer)
  If Filenumber <= -1 Then
    Err.Raise 2, , LoadResString(2): Exit Sub
  ElseIf Length < 0 Or Length > 49 Then Err.Raise 6, , LoadResString(6): Exit Sub
  ElseIf Header < 0 Or Header < CMN.GetBits(CMN.GetBits(Data)) And Header > 0 Then Err.Raise 7, , LoadResString(7): Exit Sub
  ElseIf Length = 0 And Header = 0 Then Err.Raise 8, , LoadResString(8): Exit Sub
  ElseIf Data < 0 Then Err.Raise 10, , LoadResString(10): Exit Sub: End If
  
  Data = Int(Data)
  
  Dim Pos As Integer                'The current position within the value
  Dim Multi As Currency             'Value for the mod operation
  Dim Match1 As Currency            'New bit match value
  Dim Match2 As Currency            'Old bit match value
  Dim Bool() As Boolean
  
  Multi = 1                         'Set MUTLI to multiplyable value
  
  If Header > 0 Then
    Length = CMN.GetBits(Data)
    End If
  
  If Length = 0 Then
    WriteHeader 0, Header
    Exit Sub
    End If
  
  ReDim Bool(1 To Length)
  
  Do Until Match1 = Data Or Length = Pos
    Pos = Pos + 1                   'Increase the position by one
    Multi = Multi * 2               'Double the mod value (2 ^ pos)
    Match1 = CMN.MODulus(Data, Multi)   'Modulus the value with the bit value
    
    If Match1 <> Match2 Then        'Check if values aren't the same
      Match2 = Match1               'Set the old value to the newer
      Bool(Pos) = True
      End If
  Loop
  
  If Header > 0 Then
    WriteHeader UBound(Bool), Header
    End If
  
  For Pos = LBound(Bool()) To UBound(Bool())
    WriteData Bool(Pos)
  Next
End Sub

Sub WriteTable(Data As String, Table() As Byte, Optional Length As Long, Optional Header As Integer)
  If Filenumber <= -1 Then
    Err.Raise 2, , LoadResString(2): Exit Sub
  ElseIf Length < 0 Or Length < Len(Data) And Header = 0 Then Err.Raise 6, , LoadResString(6): Exit Sub
  ElseIf Header < 0 Or Header < CMN.GetBits(CMN.GetBits(Len(Data))) And Header > 0 Then Err.Raise 7, , LoadResString(7): Exit Sub
  ElseIf Length = 0 And Header = 0 Then Err.Raise 8, , LoadResString(8): Exit Sub
  ElseIf CMN.ValidYArray(Table()) = False Then Err.Raise 9, , LoadResString(9): Exit Sub: End If
  
  Dim StrTable(0 To 255) As Byte
  Dim Pos As Currency
  Dim Size As Integer
  
  If Header > 0 Then
    WriteHeader Len(Data), Header
  ElseIf Length > 0 And Len(Data) < Length Then
    Data = Data & String(Length - Len(Data), " ")
    End If
  
  Size = CMN.GetBits(UBound(Table))
  
  For Pos = 1 To UBound(Table)
    StrTable(Table(Pos)) = Pos
  Next
  
  For Pos = 1 To Len(Data)
    WriteValue CCur(StrTable(Asc(Mid(Data, Pos, 1)))), CCur(Size)
  Next
End Sub

Sub WriteString(Data As String, Optional Length As Long, Optional Header As Integer)
  If Filenumber <= -1 Then
    Err.Raise 2, , LoadResString(2): Exit Sub
  ElseIf Length < 0 Or Length < Len(Data) And Header = 0 Then Err.Raise 4, , LoadResString(4): Exit Sub
  ElseIf Header < 0 Or Header < CMN.GetBits(CMN.GetBits(Len(Data))) And Header > 0 Then Err.Raise 5, , LoadResString(5): Exit Sub
  ElseIf Length = 0 And Header = 0 Then Err.Raise 8, , LoadResString(8): Exit Sub: End If
  
  Dim Pos As Currency
  
  If Header > 0 Then
    WriteHeader Len(Data), Header
  ElseIf Length > 0 And Len(Data) < Length Then
    Data = Data & String(Length - Len(Data), " ")
    End If
  
  For Pos = 1 To Len(Data)
    WriteValue CCur(Asc(Mid(Data, Pos, 1))), 8
  Next
End Sub

Private Sub Class_Initialize()
  Filenumber = -1
End Sub

Private Sub Class_Terminate()
  If Filenumber >= 0 Then
    CloseFile
    End If
End Sub

