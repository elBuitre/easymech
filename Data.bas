Attribute VB_Name = "Data"
Public Const GWL_EXSTYLE = (-20)
Public Const LWA_ALPHA = &H2
Public Const WS_EX_LAYERED = &H80000
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const factor = 100
Public Const pi = 3.14159265358979

Public activeBackground As Integer
Public PathToShell As String

Public hRgn As Long, X As Long
Public formWidth As Single, formHeight As Single
Public borderWidth As Single, titleHeight As Single

Public message As New MsgWin
Public GoOn As Integer
Public SavedProjectName As String
Public algorithm As Integer

Public xOffset, yOffset, counting As Integer

Public parmtr1, parmtr2, parmtr3, parmtr4 As Double
Public timeCounter As Double

' Initial conditions array
Public init_x(100) As Double
Public init_y(100) As Double

' Current point in the integration
Public xx, yy, r_x, r_y As Double
' For advanced Eulero and Runge-Kutta
Public x_A, y_A As Double

