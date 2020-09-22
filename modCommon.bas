Attribute VB_Name = "modCommon"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal Operation As String, ByVal Filename As String, ByVal Parameters As String, ByVal Directory As String, ByVal ShowCommand As Long) As Long

Public Const SEShowNormal = 1      'Open with the default state
Public Const SEShowMinimized = 2   'Minimizes the window and activates another window
Public Const SEShowMaximized = 3   'Maximizes the window and retrieves focus
Public Const SEShowNoActivate = 4  'Does not activate the window
Public Const SEShow = 5            'Activates the window
Public Const SEMinimize = 6        'Minimizes the window
Public Const SEShowMinNoActive = 7 'Display window as minimized without activation
Public Const SEShowNA = 8          'Do not modify the windows state
Public Const SERestore = 9         'Restores the window and retrieves fouces
Public Const SEShowDefault = 10    'Complex command, look up in MSDN

'Navigate to a specific URL
Sub Navigate(URL As String, Optional Flags As Long = 1)
  ShellExecute 0&, "", URL, "", "", Flags 'Execute the command
End Sub
