VERSION 5.00
Begin VB.Form Monitor 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13500
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   ScaleHeight     =   650
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox yo 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   495
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   420
      Width           =   1050
   End
   Begin VB.TextBox xo 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   480
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   90
      Width           =   1050
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "d"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   11610
      TabIndex        =   16
      Top             =   975
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "c"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   11610
      TabIndex        =   15
      Top             =   675
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "b"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   11610
      TabIndex        =   14
      Top             =   375
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   11325
      TabIndex        =   13
      Top             =   75
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   12225
      TabIndex        =   12
      Tag             =   "--"
      Top             =   975
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   12225
      TabIndex        =   11
      Tag             =   "--"
      Top             =   675
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   12225
      TabIndex        =   10
      Top             =   375
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   12225
      TabIndex        =   9
      Top             =   75
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "yo"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   0
      TabIndex        =   6
      Top             =   390
      Width           =   375
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "xo"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Top             =   75
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Equilibrium"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   6270
      TabIndex        =   4
      Top             =   8745
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H008080FF&
      Height          =   120
      Left            =   6030
      Shape           =   3  'Circle
      Top             =   8820
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   480
      X2              =   985
      Y1              =   612
      Y2              =   612
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   -88
      X2              =   417
      Y1              =   612
      Y2              =   612
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   12810
      MouseIcon       =   "Monitor.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   9405
      Width           =   435
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   7185
      MouseIcon       =   "Monitor.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   9405
      Width           =   525
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Clean"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5760
      MouseIcon       =   "Monitor.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   9405
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   180
      MouseIcon       =   "Monitor.frx":091E
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   9405
      Visible         =   0   'False
      Width           =   435
   End
End
Attribute VB_Name = "Monitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim initial_ray, final_ray, error As Double 'For the error in the problem "Oscillatore armonico"

Private Sub Form_Load()

    counting = 0
    Monitor.Cls
    If NewFrm.problem <> "Sistema di Lotka-Volterra" Then
        Monitor.Line (0, 305)-(900, 305), &HFFFFFF
        Monitor.Line (450, 0)-(450, 610), &HFFFFFF
        xOffset = 450
        yOffset = 305
    Else
        Monitor.Line (0, 550)-(900, 550), &HFFFFFF
        Monitor.Line (40, 70)-(40, 610), &HFFFFFF
        xOffset = 40
        yOffset = 550
        Monitor.Circle (factor * (parmtr3 / parmtr4) + xOffset, yOffset - (factor * (parmtr1 / parmtr2))), 3, &HFF
        Shape1.Visible = True
        Label5.Visible = True
    End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Y < 580 Then 'Check if the cursor is inside the correct zone
    
        'Save the initial conditions in the array
        If NewFrm.problem <> "Sezioni di Poincarè del pendolo forzato" Then
            init_x(counting) = (X - xOffset) / factor
            init_y(counting) = (yOffset - Y) / factor
        Else
            init_x(counting) = (X - xOffset) / 140
            init_y(counting) = (yOffset - Y) / 140
        End If
        
        'Setting the initial conditions
        xx = init_x(counting)
        yy = init_y(counting)
        r_x = 0
        r_y = 0
        
        'Showing label
        If NewFrm.problem = "Oscillatore armonico smorzato" Then
            Label12.Font = "Symbol"
            Label12.Caption = "m"
            Label12.Visible = True
            Label8.Caption = FormatNumber(parmtr1, 2)
            Label8.Visible = True
        End If
        If NewFrm.problem = "Sistema di Lotka-Volterra" Then
            Label12.Font = "Symbol"
            Label13.Font = "Symbol"
            Label14.Font = "Symbol"
            Label15.Font = "Symbol"
            Label12.Caption = "a"
            Label13.Caption = "b"
            Label14.Caption = "c"
            Label15.Caption = "d"
            Label12.Visible = True
            Label13.Visible = True
            Label14.Visible = True
            Label15.Visible = True
            Label8.Caption = FormatNumber(parmtr1, 2)
            Label9.Caption = FormatNumber(parmtr2, 2)
            Label10.Caption = FormatNumber(parmtr3, 2)
            Label11.Caption = FormatNumber(parmtr4, 2)
            Label8.Visible = True
            Label9.Visible = True
            Label10.Visible = True
            Label11.Visible = True
        End If
        
        integratorLoop
        
        'Showing error
        If NewFrm.problem = "Oscillatore armonico" Then
            initial_ray = Sqr((init_x(counting) * init_x(counting)) + init_y(counting) * init_y(counting))
            final_ray = Sqr((xx * xx) + (yy * yy))
            Label12.Caption = "Error"
            Label12.Visible = True
            Label8.Caption = Format$(final_ray - initial_ray, "##e-00")
            Label8.Visible = True
        End If
       
        counting = counting + 1
    End If
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim xPos, yPos As Double

    If NewFrm.problem = "Pendolo doppio" Or NewFrm.problem = "Pendolo" Then
       xPos = (180 / 3.14159265358979) * (X - xOffset) / factor
       yPos = (180 / 3.14159265358979) * (-Y + yOffset) / factor
    ElseIf NewFrm.problem = "Sezioni di Poincarè del pendolo forzato" Then
       xPos = (180 / 3.14159265358979) * (X - xOffset) / 140
       yPos = (180 / 3.14159265358979) * (-Y + yOffset) / 140
    Else
        xPos = (X - xOffset) / factor
        yPos = (yOffset - Y) / factor
    End If
    xo.text = FormatNumber(xPos, 4)
    yo.text = FormatNumber(yPos, 4)
    
    BlinkLabel Label1, 41, 12, 644, 625, X, Y, 12, 15
    BlinkLabel Label2, 410, 441, 368, 354, X, Y, 12, 15
    BlinkLabel Label3, 24, 55, 368, 354, X, Y, 12, 15
    BlinkLabel Label4, 410, 441, 368, 354, X, Y, 12, 15

End Sub

Private Sub Label1_Click()

    enableForm Monitor, False, 180
    SaveFrm.Show 1
    
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    BlinkLabel Label1, 435, 0, 285, 0, X, Y, 12, 15
    
End Sub

Private Sub Label2_Click()

    For counting = 0 To 99
        init_x(counting) = 0
        init_y(counting) = 0
    Next counting
    
    counting = 0
    Monitor.Cls
    If NewFrm.problem <> "Sistema di Lotka-Volterra" Then
        Monitor.Line (0, 305)-(900, 305), &HFFFFFF
        Monitor.Line (450, 0)-(450, 610), &HFFFFFF
        xOffset = 450
        yOffset = 305
    Else
        Monitor.Line (0, 550)-(900, 550), &HFFFFFF
        Monitor.Line (40, 70)-(40, 610), &HFFFFFF
        xOffset = 40
        yOffset = 600
        Monitor.Circle (factor * (parmtr3 / parmtr4) + xOffset, yOffset - (factor * (parmtr1 / parmtr2))), 3, &HFF
        Shape1.Visible = True
        Label5.Visible = True
    End If
    
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    BlinkLabel Label2, 525, 0, 285, 0, X, Y, 12, 15

End Sub

Private Sub Label3_Click()
    
    Printer.Orientation = vbPRORLandscape
    Monitor.PrintForm
    Printer.Orientation = vbPRORPortrait

End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    BlinkLabel Label3, 525, 0, 285, 0, X, Y, 12, 15

End Sub

Private Sub Label4_Click()

    Unload Me
    enableForm MainFrm, True, 255

End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    BlinkLabel Label4, 435, 0, 285, 0, X, Y, 12, 15

End Sub
