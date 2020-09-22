VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Common"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function GetBits(Value As Currency) As Integer
  Dim Multi As Currency
  
  Multi = 1
  Do Until Multi > Value
    Multi = Multi * 2
    GetBits = GetBits + 1
  Loop
End Function

Function MODulus(ByVal Value1 As Currency, ByVal Value2 As Currency) As Currency
  Value1 = Int(Value1)                                  'Round value
  Value2 = Int(Value2)                                  'Round value
  MODulus = Int(Value1 - Int(Value1 / Value2) * Value2) 'Modulus formula
End Function

Function ValidBArray(Data() As Boolean) As Boolean
  On Error Resume Next
  
  Dim Temp As Long
  Temp = UBound(Data())
  
  ValidBArray = (Err.Number = 0)
End Function

Function ValidYArray(Data() As Byte) As Boolean
  On Error Resume Next
  
  Dim Temp As Long
  Temp = UBound(Data())
  
  ValidYArray = (Err.Number = 0)
End Function

Function FileExists(File As String) As Boolean
  On Error Resume Next              'If error occurs continue with next line
  
  Dim Filenumber As Integer         'The open filenumber
  Filenumber = FreeFile             'Retrieve a free filenumber
  
  Open File For Binary Access Read As Filenumber 'Open the file
    FileExists = (Err.Number = 0)   'If no error occured file exists
  Close Filenumber                  'Close the file
  
  Err.Clear                         'Clear if any errors occured
End Function
