VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Read"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private CMN As New Common
Private BytePos As Byte
Private PosMulti As Byte
Private FileByte As Byte
Public Filenumber As Integer

Property Get Position() As Currency
  If Filenumber >= 0 Then
    Position = (Seek(Filenumber) - 2) * 8 + BytePos
  Else
    Position = -1
    End If
End Property

Sub LoadFile(Filename As String)
  If Filenumber >= 0 Then
    Err.Raise 3, , LoadResString(3): Exit Sub
  ElseIf CMN.FileExists(Filename) = False Then Err.Raise 1, , LoadResString(1): Exit Sub: End If
  
  Filenumber = FreeFile
  
  Open Filename For Binary Access Read As #Filenumber
  
  Get #Filenumber, , FileByte
  PosMulti = 1
  BytePos = 0
End Sub

Sub CloseFile()
  If Filenumber <= -1 Then
    Err.Raise 2, , LoadResString(2): Exit Sub: End If
  
  Close #Filenumber
  
  Filenumber = -1
  PosMulti = 0
  BytePos = 0
  FileByte = 0
End Sub

Sub SeekFile(Position As Currency)
  If Filenumber <= -1 Then
    Err.Raise 2, , LoadResString(2): Exit Sub: End If
  
  Seek #Filenumber, Int((Position) / 8) + 1
  
  Get #Filenumber, , FileByte
  BytePos = Position Mod 8
  PosMulti = 2 ^ (Position Mod 8)
End Sub

Function ReadData() As Boolean
  If Filenumber <= -1 Then
    Err.Raise 2, , LoadResString(2): Exit Function: End If
  
  BytePos = BytePos + 1
  
  ReadData = (FileByte Mod 2 ^ (BytePos - 1) <> FileByte Mod 2 ^ BytePos) 'Check if it isn't the same as the higher bit then it is true
  
  If BytePos = 8 Then
    Get #Filenumber, , FileByte
    PosMulti = 1
    BytePos = 0
  Else
    PosMulti = PosMulti * 2
    End If
End Function

Function ReadHeader(Length As Integer) As Long
  If Filenumber <= -1 Then
    Err.Raise 2, , LoadResString(2): Exit Function
  ElseIf Length <= 0 Then Err.Raise 4, , LoadResString(4): Exit Function: End If
  
  ReadHeader = ReadValue(Length)
End Function

Sub ReadSpace(Optional Length As Integer, Optional Header As Integer)
  If Filenumber <= -1 Then
    Err.Raise 2, , LoadResString(2): Exit Sub
  ElseIf Length < 0 Then Err.Raise 6, , LoadResString(6): Exit Sub
  ElseIf Header < 0 Then Err.Raise 7, , LoadResString(7): Exit Sub
  ElseIf Length = 0 And Header = 0 Then Err.Raise 8, , LoadResString(8): Exit Sub: End If
  
  If Header > 0 Then
    Length = ReadHeader(Header)
    End If
  
  For Pos = 1 To Length
    ReadData
  Next
End Sub

Sub ReadBooleans(Data() As Boolean, Optional Header As Integer)
  If Filenumber <= -1 Then
    Err.Raise 2, , LoadResString(2): Exit Sub
  ElseIf CMN.ValidBArray(Data()) = False And Header = 0 Then Err.Raise 6, , LoadResString(6): Exit Sub
  ElseIf Header < 0 Then Err.Raise 7, , LoadResString(7): Exit Sub: End If
  
  Dim Pos As Long
  Dim Size As Integer
  
  If Header > 0 Then
    Size = ReadHeader(Header)
    ReDim Data(LBound(Data()) To Size - LBound(Data()) + 1)
    End If
  
  For Pos = LBound(Data()) To UBound(Data())
    Data(Pos) = ReadData
  Next
End Sub

Function ReadValue(Optional Length As Integer, Optional Header As Integer) As Currency
  If Filenumber <= -1 Then
    Err.Raise 2, , LoadResString(2): Exit Function
  ElseIf Length < 0 Or Length > 49 Then Err.Raise 6, , LoadResString(6): Exit Function
  ElseIf Header < 0 Then Err.Raise 7, , LoadResString(7): Exit Function
  ElseIf Length = 0 And Header = 0 Then Err.Raise 8, , LoadResString(8): Exit Function: End If
  
  Dim Pos As Integer
  Dim Multi As Currency
  
  If Header > 0 Then
    Length = ReadHeader(Header)
    End If
  
  Multi = 0.5
  
  For Pos = 1 To Length
    Multi = Multi * 2
    
    If ReadData Then
      ReadValue = ReadValue + Multi
      End If
  Next
End Function

Function ReadTable(Table() As Byte, Optional Length As Long, Optional Header As Integer) As String
  If Filenumber <= -1 Then
    Err.Raise 2, , LoadResString(2): Exit Function
  ElseIf Length < 0 Then Err.Raise 6, , LoadResString(6): Exit Function
  ElseIf Header < 0 Then Err.Raise 7, , LoadResString(7): Exit Function
  ElseIf Length = 0 And Header = 0 Then Err.Raise 8, , LoadResString(8): Exit Function
  ElseIf CMN.ValidYArray(Table()) = False Then Err.Raise 9, , LoadResString(9): Exit Function: End If
  
  Dim Pos As Currency
  
  If Header > 0 Then
    Length = ReadHeader(Header)
    End If
  
  Size = CMN.GetBits(UBound(Table))
  
  For Pos = 1 To Length
    A = ReadValue(CInt(Size))
    ReadTable = ReadTable & Chr(Table(A))
  Next
End Function

Function ReadString(Optional Length As Long, Optional Header As Integer) As String
  If Filenumber <= -1 Then
    Err.Raise 2, , LoadResString(2): Exit Function
  ElseIf Length < 0 Then Err.Raise 6, , LoadResString(6): Exit Function
  ElseIf Header < 0 Then Err.Raise 7, , LoadResString(7): Exit Function
  ElseIf Length = 0 And Header = 0 Then Err.Raise 8, , LoadResString(8): Exit Function: End If
  
  Dim Pos As Currency
  
  If Header > 0 Then
    Length = ReadHeader(Header)
    End If
  
  For Pos = 1 To Length
    ReadString = ReadString & Chr(ReadValue(8))
  Next
End Function

Private Sub Class_Initialize()
  Filenumber = -1
End Sub

Private Sub Class_Terminate()
  If Filenumber >= 0 Then
    CloseFile
    End If
End Sub
