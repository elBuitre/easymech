Attribute VB_Name = "Functions"
Option Explicit

Declare Function CreateRoundRectRgn Lib _
    "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, _
    ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Declare Function SetWindowRgn Lib _
    "user32" (ByVal hWnd As Long, ByVal hRgn As Long, _
    ByVal bRedraw As Boolean) As Long
Declare Function DeleteObject Lib _
    "gdi32" (ByVal hObject As Long) As Long
Declare Function PaintDesktop Lib _
    "user32" (ByVal hdc As Long) As Long
Declare Function SetWindowPos Lib "user32" _
(ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function SetLayeredWindowAttributes Lib "user32" ( _
ByVal hWnd As Long, _
ByVal crKey As Long, _
ByVal bAlpha As Byte, _
ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
ByVal hWnd As Long, _
ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
ByVal hWnd As Long, _
ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long

Dim step As Double

Public Sub FormFade(Fading As Integer, ByRef frmForm As Form, blnHide As Boolean)
    
    Dim MSG As Long
    Dim i As Long
    
    If blnHide = True Then
    
        For i = Fading To 0 Step -5
            'Set window style to layered
            MSG = GetWindowLong(frmForm.hWnd, GWL_EXSTYLE)
            MSG = MSG Or WS_EX_LAYERED
            SetWindowLong frmForm.hWnd, GWL_EXSTYLE, MSG
            'Set the opacity of the layer according the the parameters
            SetLayeredWindowAttributes frmForm.hWnd, 0, i, LWA_ALPHA
            frmForm.Refresh
        Next
        
    Else
    
        For i = 0 To Fading Step 5
            'Set window style to layered
            MSG = GetWindowLong(frmForm.hWnd, GWL_EXSTYLE)
            MSG = MSG Or WS_EX_LAYERED
            SetWindowLong frmForm.hWnd, GWL_EXSTYLE, MSG
            'Set the opacity of the layer according the the parameters
            SetLayeredWindowAttributes frmForm.hWnd, 0, i, LWA_ALPHA
            frmForm.Refresh
        Next
        
    End If
    
End Sub

Public Sub AlwaysOnTop(X As Form, Y As Boolean, xPos As Integer, yPos As Integer)

    Select Case Y
        Case Is = True
            SetWindowPos X.hWnd, HWND_TOPMOST, xPos, yPos, 0, 0, 1
        Case Is = False
            SetWindowPos X.hWnd, HWND_NOTOPMOST, 0, 0, 0, _
            0, SWP_NOACTIVATE Or SWP_SHOWWINDOW
    End Select

End Sub

Public Sub RoundedCorner(frm As Form)

    ' Calculate the form area
    borderWidth = (frm.Width - frm.ScaleWidth) / 2
    titleHeight = frm.Height - frm.ScaleHeight - borderWidth
    
    ' Convert to Pixels
    borderWidth = frm.ScaleX(borderWidth, vbTwips, vbPixels)
    titleHeight = frm.ScaleY(titleHeight, vbTwips, vbPixels)
    formWidth = frm.ScaleX(frm.ScaleWidth + borderWidth, vbTwips, vbPixels)
    formHeight = frm.ScaleY(frm.ScaleHeight + titleHeight, vbTwips, vbPixels)
    
    ' Create a round rectangle region around the graphics area of the form
    hRgn = CreateRoundRectRgn(borderWidth, titleHeight, formWidth + borderWidth, _
                              formHeight + titleHeight, 30, 30)
    
    ' Set the clipping area of the window using the resulting region
    SetWindowRgn frm.hWnd, hRgn, True
    
    ' Tidy up
    X = DeleteObject(hRgn)
    DoEvents
    PaintDesktop frm.hdc

End Sub

Public Sub BlinkLabel(lbl As Label, maxX, minX, maxY, minY As Integer, _
                      cursorX, cursorY As Single, activeColor, disabledColor As Integer)
    
    If cursorX <= maxX And cursorX >= minX And cursorY <= maxY And cursorY >= minY Then
        lbl.ForeColor = QBColor(activeColor)
    Else
        lbl.ForeColor = QBColor(disabledColor)
    End If

End Sub

Public Sub OpenSavedProject()

    Dim Notepad As Integer
    Dim NameLength As Integer
    
    If Dir(SavedProjectName) = "" Or SavedProjectName = OpenFrm.Drive1.Drive & "\" Then
        enableForm OpenFrm, False, 180
        message.ShowMessage "Il nome specificato non è valido!", "OkOnly"
    Else
        Unload OpenFrm
        Monitor.Show
        Monitor.Picture = LoadPicture(SavedProjectName)
    
        NameLength = Len(SavedProjectName)
        If Dir(Left(SavedProjectName, NameLength - 3) & "txt") <> "" Then
            Notepad = Shell(PathToShell & " " & Left(SavedProjectName, NameLength - 3) & "txt", vbNormalFocus)
        Else
            message.ShowMessage "File con i dati iniziali inesistente!", "OkOnly"
        End If
    End If
    
End Sub

Public Sub enableForm(frm As Form, enbl As Boolean, color As Integer)

    Dim ctl As Control
    
    If Not enbl Then
        For Each ctl In frm.Controls
            If TypeOf ctl Is Label Then ctl.ForeColor = RGB(180, 180, 180)
        Next ctl
    Else
        For Each ctl In frm.Controls
            If TypeOf ctl Is Label Then ctl.ForeColor = RGB(color, color, color)
        Next ctl
    End If
    
End Sub

Public Sub integratorLoop()

    timeCounter = 0
    Select Case algorithm
        Case 0
            eulero
        Case 1
            advancedEulero
        Case 2
            RungeKutta
    End Select
    
End Sub

Public Sub eulero()

    Do
        If factor * xx + xOffset > 0 And yOffset - (factor * yy) > 0 And factor * xx + xOffset < 900 And yOffset - (factor * yy) < 600 Then
            Monitor.PSet (factor * xx + xOffset, yOffset - (factor * yy))
        End If
        deriv
        xx = xx + Val(NewFrm.step) * r_x
        yy = yy + Val(NewFrm.step) * r_y
        timeCounter = timeCounter + Val(NewFrm.step)
    Loop While (timeCounter <= Val(NewFrm.final_time))

End Sub

Public Sub advancedEulero()

    Dim halfStep As Double
    Dim innerCounter As Integer

    halfStep = Val(NewFrm.step) / 2
    Do
        If factor * xx + xOffset > 0 And yOffset - (factor * yy) > 0 And factor * xx + xOffset < 900 And yOffset - (factor * yy) < 600 Then
            Monitor.PSet (factor * xx + xOffset, yOffset - (factor * yy))
        End If
        
        x_A = xx + halfStep * r_x
        y_A = yy + halfStep * r_y
        For innerCounter = 1 To 3
            xx = x_A + halfStep * r_x
            yy = y_A + halfStep * r_y
            deriv
        Next innerCounter
        timeCounter = timeCounter + Val(NewFrm.step)
    Loop While (timeCounter <= Val(NewFrm.final_time))

End Sub

Public Sub deriv()

    Dim num1, num2, den As Double
    
    Select Case NewFrm.problem
        Case "Oscillatore armonico"
            r_x = yy
            r_y = -xx
            
        Case "Repulsore armonico"
            r_x = yy
            r_y = xx
            
        Case "Pendolo"
            r_x = yy
            r_y = -Sin(xx)
            
        Case "Oscillatore armonico smorzato"
            r_x = yy
            r_y = -2 * parmtr1 * yy - xx
            
        Case "Sistema di Lotka-Volterra"
            r_x = parmtr1 * xx - parmtr2 * xx * yy
            r_y = -parmtr3 * yy + parmtr4 * xx * yy
            
        Case "Ciclo limite - Equazione di Van Der Pol"
            If parmtr1 > 1 Then
                r_x = parmtr1 * (yy - (xx * xx * xx / 3 - xx))
                r_y = -xx / parmtr1
            Else
                r_x = yy
                r_y = -parmtr1 * ((xx * xx) - 1) * yy - xx
            End If
            
        Case "Pendolo doppio"
            If xx > -3.74 And xx < 3.74 Then
                num1 = -r_y * r_y * Sin(xx - yy) - 2 * Sin(xx) + Cos(xx - yy) * (-r_x * r_x * Sin(xx - yy) + Sin(yy))
                num2 = 2 * r_x * r_x * Sin(xx - yy) - 2 * Sin(yy) + Cos(xx - yy) * (r_y * r_y * Sin(xx - yy) + 2 * Sin(xx))
                den = 2 - ((Cos(xx - yy) * Cos(xx - yy)))
                r_x = r_x + Val(NewFrm.step) * (num1 / den)
                r_y = r_y + Val(NewFrm.step) * (num2 / den)
            End If
    End Select

End Sub

Public Sub RungeKutta()
    
    Dim aa(4) As Double
    Dim bb(4) As Double
    Dim cc(4) As Double
    Dim k As Integer
    Dim j As Integer
    Dim i As Integer
    Dim forcingT As Double
    Dim forcingF As Double
    Dim forcingA As Double
    Dim k1x, k1v, k2x, k2v, k3x, k3v, k4x, k4v As Double
    
    Screen.MousePointer = vbHourglass
    Monitor.Line1.Visible = False
    Monitor.Line2.Visible = False
    Monitor.Label1.Visible = False
    Monitor.Label2.Visible = False
    Monitor.Label3.Visible = False
    Monitor.Label4.Visible = False
    
    If NewFrm.problem <> "Sezioni di Poincarè del pendolo forzato" Then
        step = Val(NewFrm.final_time) / (Val(NewFrm.step) * Val(NewFrm.points))
        aa(0) = 0.5
        aa(1) = 1 - Sqr(0.5)
        aa(2) = 1 + Sqr(0.5)
        aa(3) = 1 / 6
        bb(0) = 2
        bb(1) = 1
        bb(2) = 1
        bb(3) = 2
        cc(0) = 0.5
        cc(1) = 1 - Sqr(0.5)
        cc(2) = 1 + Sqr(0.5)
        cc(3) = 0.5
        deriv
        For k = 1 To Val(NewFrm.points)
            For j = 1 To Val(NewFrm.step)
                For i = 0 To 3
                    Call increment(aa(i), bb(i), cc(i))
                    deriv
                Next i
                timeCounter = timeCounter + step
            Next j
            If factor * xx + xOffset > 0 And yOffset - (factor * yy) > 0 And factor * xx + xOffset < 900 And yOffset - (factor * yy) < 600 Then
                Monitor.PSet (factor * xx + xOffset, yOffset - (factor * yy))
                Monitor.Refresh
            End If
        Next k
    Else
        forcingF = Val(NewFrm.parameter4)
        forcingT = 2 * 3.14159265358979 / forcingF
        forcingA = Val(NewFrm.parameter2)
        step = forcingT / Val(NewFrm.step)
    
        For i = 1 To Val(NewFrm.final_time)
            For j = 1 To Val(NewFrm.step)
                k1v = (forcingA * Cos(forcingF * timeCounter) - Sin(xx)) * step
                k1x = yy * step
                timeCounter = timeCounter + 0.5 * step
                k2v = (forcingA * Cos(forcingF * timeCounter) - Sin(xx + 0.5 * k1x)) * step
                k2x = (yy + 0.5 * k1v) * step
                k3v = (forcingA * Cos(forcingF * timeCounter) - Sin(xx + 0.5 * k2x)) * step
                k3x = (yy + 0.5 * k2v) * step
                timeCounter = timeCounter + 0.5 * step
                k4v = (forcingA * Cos(forcingF * timeCounter) - Sin(xx + k3x)) * step
                k4x = (yy + k3v) * step
                yy = yy + (k1v + 2 * k2v + 2 * k3v + k4v) / 6
                xx = xx + (k1x + 2 * k2x + 2 * k3x + k4x) / 6
                If xx > 3.14159265358979 Then xx = xx - 2 * 3.14159265358979
                If xx < -3.14159265358979 Then xx = xx + 2 * 3.14159265358979
            Next j
            If (xx * 140 + xOffset > 0) And (xx * 140 + xOffset < 900) And (yOffset - yy * 140 > 0) And (yOffset - yy * 140 < 600) Then
                Monitor.PSet (xx * 140 + xOffset, yOffset - yy * 140)
                Monitor.Refresh
            End If
        Next i
        
    End If
    
    Monitor.Line1.Visible = True
    Monitor.Line2.Visible = True
    Monitor.Label1.Visible = True
    Monitor.Label2.Visible = True
    Monitor.Label3.Visible = True
    Monitor.Label4.Visible = True
    Screen.MousePointer = vbNormal
    
End Sub

Sub increment(ap As Double, bp As Double, cp As Double)

    Dim aum, rr  As Double
    
    rr = r_x * step
    aum = ap * (rr - bp * x_A)
    xx = xx + aum
    x_A = x_A + 3 * aum - cp * rr
    
    rr = r_y * step
    aum = ap * (rr - bp * y_A)
    yy = yy + aum
    y_A = y_A + 3 * aum - cp * rr
    
End Sub

