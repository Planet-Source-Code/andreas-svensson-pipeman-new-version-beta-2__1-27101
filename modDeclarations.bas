Attribute VB_Name = "modDeclarations"
Public Declare Function MSTimer Lib "winmm.dll" Alias "timeGetTime" () As Long

Public WaterField As PointInfo
Public WaterPos As PointInfo
Public WaterPipe As Integer
Public WaterOrient As PointInfo
Public WaterAnim As Integer
Public WaterWater As Boolean

Public PLID As String
Public Field(0 To 10, 0 To 10) As BoxInfo
Public Boxes(1 To 5) As Integer
Public PipeSize As Integer

Public EnterMode As Boolean
Public EnterX As Integer
Public EnterY As Integer
Public EnterName As String

Public Score As Long
Public Bonus As Double
Public GameOver As Boolean
Public Pause As Boolean
Public TriggerTime As Integer
Public GameOverPic(0 To 10, 0 To 10) As Byte
Public NextMapPic(0 To 10, 0 To 10) As Byte
Public PipeManPic(0 To 10, 0 To 10) As Byte
Public PausePic(0 To 10, 0 To 10) As Byte
Public PausedField(0 To 10, 0 To 10) As Byte
Public Force As Boolean
Public CurrentMap As Integer
Public Speed As Integer
Public Teleport(1 To 2) As PointInfo
Public BonusMulti As Single
Public SpeedMulti As Boolean
Public ElapsedTime As Long
Public PauseTime As Long
Public Tables() As Byte

Public MX As Integer
Public MY As Integer

Public WinPipe As Integer

Public Type BoxInfo
  PipeType As Integer
  Watered As Boolean
End Type

Public Type PointInfo
  X As Integer
  Y As Integer
End Type

Public HighScore(1 To 10) As ScoreType

Private Type ScoreType
  Score As Long
  Player As String * 6
End Type
