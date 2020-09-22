Attribute VB_Name = "modMain"
Sub Main()
  frmStartUp.Picture = LoadResPicture("STARTLOGO", 0)
  frmStartUp.Show
  DoEvents
  ElapsedTime = MSTimer
  Randomize MSTimer
  
  PipeSize = 32 * Screen.TwipsPerPixelX
  
  ChDir App.Path
  
  EmptyBoxes
  frmMain.LogoLine.Picture = LoadResPicture("MENULINE", 0)
  frmMain.LogoImg.Picture = LoadResPicture("MENULOGO", 0)
  ReadPics GameOverPic(), "pipelogo.fdi", 0
  ReadPics NextMapPic(), "pipelogo.fdi", 1
  ReadPics PipeManPic(), "pipelogo.fdi", 2
  ReadPics PausePic(), "pipelogo.fdi", 3
  DrawBoxes
  ResetField
  
  CreateTable "ABCDEFGHIJKLMNOPQRSTUVWXYZ ", Tables()
  
  Dim P As Integer
  
  For P = 1 To 10
    HighScore(P).Player = Space(6)
  Next
  
  LoadHiscore "hiscore.fdi"
  
  Dim X As Integer
  Dim Y As Integer
  
  
  For X = 0 To 10
    For Y = 0 To 10
      P = PausePic(X, Y)
      frmMain.picPause.PaintPicture LoadResPicture(P, 0), X * PipeSize, Y * PipeSize
    Next
  Next
  
  Erase PausePic
  
  ShowPicture PipeManPic(), True
  
  frmMain.PSet (0, 0), RGB(172, 172, 162)
  frmMain.Line (0, 1)-(0, frmMain.Height - 2 * Screen.TwipsPerPixelY), RGB(157, 156, 149)
  frmMain.Line (1, 0)-(frmMain.Width - 2 * Screen.TwipsPerPixelY, 0), RGB(188, 187, 175)
  frmMain.Line (frmMain.Width - 2 * Screen.TwipsPerPixelY, 0)-(frmMain.Width - 2 * Screen.TwipsPerPixelY, frmMain.Height - 2 * Screen.TwipsPerPixelY), RGB(188, 187, 175)
  frmMain.Line (frmMain.Width - Screen.TwipsPerPixelY, 0)-(frmMain.Width - Screen.TwipsPerPixelY, frmMain.Height), RGB(119, 118, 106)
  frmMain.Line (0, frmMain.Height - 2 * Screen.TwipsPerPixelY)-(frmMain.Width - Screen.TwipsPerPixelY, frmMain.Height - 2 * Screen.TwipsPerPixelY), RGB(157, 156, 149)
  frmMain.Line (0, frmMain.Height - Screen.TwipsPerPixelY)-(frmMain.Width - Screen.TwipsPerPixelY, frmMain.Height - Screen.TwipsPerPixelY), RGB(57, 57, 55)
  
  frmMain.picAppBar.PaintPicture LoadResPicture("LOGOLEFT", 0), 0, 0
  frmMain.picAppBar.PaintPicture LoadResPicture("LOGOMIDDLE", 0), 3810, 0, frmMain.picAppBar.Width - 3810
  
  frmMain.picControl(1).Picture = LoadResPicture("UPEXIT", 0)
  frmMain.picControl(0).Picture = LoadResPicture("UPMINIMIZE", 0)
  Load frmMain
  
  GameOver = True
  
  frmMain.lblGameInfo.Caption = "PipeMan - Andreas Svensson"
  
  Do While ElapsedTime + 1000 > MSTimer
    DoEvents
  Loop
  
  Unload frmStartUp
  frmMain.Show
End Sub
