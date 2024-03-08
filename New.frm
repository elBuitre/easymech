VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form NewFrm 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   6300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7095
   LinkTopic       =   "Form2"
   Picture         =   "New.frx":0000
   ScaleHeight     =   6300
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox parameter4 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4275
      TabIndex        =   21
      Top             =   5340
      Visible         =   0   'False
      Width           =   2070
   End
   Begin VB.TextBox parameter3 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4275
      TabIndex        =   20
      Top             =   4920
      Visible         =   0   'False
      Width           =   2070
   End
   Begin VB.TextBox parameter2 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   990
      TabIndex        =   17
      Top             =   5340
      Visible         =   0   'False
      Width           =   2070
   End
   Begin VB.TextBox parameter1 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   990
      TabIndex        =   16
      Top             =   4920
      Visible         =   0   'False
      Width           =   2070
   End
   Begin VB.TextBox points 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4995
      TabIndex        =   12
      Top             =   3465
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox step 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4995
      TabIndex        =   10
      Top             =   2925
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox final_time 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5010
      TabIndex        =   7
      Top             =   2400
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   90
      Top             =   165
   End
   Begin MSForms.OptionButton runge_kutta 
      Height          =   270
      Left            =   720
      TabIndex        =   29
      Top             =   3540
      Width           =   285
      VariousPropertyBits=   746588179
      BackColor       =   14737632
      ForeColor       =   -2147483630
      DisplayStyle    =   5
      Size            =   "503;476"
      Value           =   "0"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label17 
      Height          =   225
      Left            =   1155
      TabIndex        =   28
      Top             =   3555
      Width           =   1575
      BackColor       =   14737632
      VariousPropertyBits=   8388627
      Caption         =   "Runge Kutta"
      Size            =   "2778;397"
      FontName        =   "Microsoft Sans Serif"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label16 
      Height          =   225
      Left            =   1155
      TabIndex        =   27
      Top             =   2955
      Width           =   1575
      BackColor       =   14737632
      VariousPropertyBits=   8388627
      Caption         =   "Advanced Eulero"
      Size            =   "2778;397"
      FontName        =   "Microsoft Sans Serif"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.OptionButton advanced_eulero 
      Height          =   270
      Left            =   720
      TabIndex        =   26
      Top             =   2955
      Width           =   285
      VariousPropertyBits=   746588179
      BackColor       =   14737632
      ForeColor       =   -2147483630
      DisplayStyle    =   5
      Size            =   "503;476"
      Value           =   "0"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label15 
      Height          =   225
      Left            =   1155
      TabIndex        =   25
      Top             =   2400
      Width           =   675
      BackColor       =   14737632
      VariousPropertyBits=   8388627
      Caption         =   "Eulero"
      Size            =   "1191;397"
      FontName        =   "Microsoft Sans Serif"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.OptionButton eulero 
      Height          =   300
      Left            =   720
      TabIndex        =   24
      Top             =   2385
      Width           =   270
      VariousPropertyBits=   746588179
      BackColor       =   14737632
      ForeColor       =   -2147483630
      DisplayStyle    =   5
      Size            =   "476;529"
      Value           =   "0"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox problem 
      Height          =   315
      Left            =   330
      TabIndex        =   23
      Top             =   1230
      Width           =   3465
      VariousPropertyBits=   746604563
      DisplayStyle    =   3
      Size            =   "6112;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Microsoft Sans Serif"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   285
      Left            =   450
      TabIndex        =   22
      Top             =   180
      Width           =   6195
      VariousPropertyBits=   746604563
      Size            =   "10927;503"
      SpecialEffect   =   0
      FontName        =   "Microsoft Sans Serif"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "p4"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3915
      TabIndex        =   19
      Top             =   5370
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "p3"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3930
      TabIndex        =   18
      Top             =   4980
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "p2"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   660
      TabIndex        =   15
      Top             =   5370
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "p1"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   630
      TabIndex        =   14
      Top             =   4965
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Problem parameters"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   315
      TabIndex        =   13
      Top             =   4515
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00808080&
      Height          =   990
      Left            =   315
      Shape           =   4  'Rounded Rectangle
      Top             =   4770
      Visible         =   0   'False
      Width           =   6465
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Points"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3975
      TabIndex        =   11
      Top             =   3510
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Step"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3975
      TabIndex        =   9
      Top             =   3000
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Final time"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3975
      TabIndex        =   8
      Top             =   2445
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   2115
      Left            =   3765
      Shape           =   4  'Rounded Rectangle
      Top             =   2070
      Width           =   3015
   End
   Begin VB.Label help 
      BackStyle       =   0  'Transparent
      Caption         =   "Select an algorithm..."
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   4380
      MouseIcon       =   "New.frx":427B1
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   3000
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   2115
      Left            =   315
      Shape           =   4  'Rounded Rectangle
      Top             =   2070
      Width           =   3015
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Parameters"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5790
      TabIndex        =   4
      Top             =   1815
      Width           =   1005
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Algorithm"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   315
      TabIndex        =   3
      Top             =   1815
      Width           =   900
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Problem"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   315
      TabIndex        =   2
      Top             =   975
      Width           =   750
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      X1              =   4095
      X2              =   9165
      Y1              =   660
      Y2              =   660
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      X1              =   -2205
      X2              =   2865
      Y1              =   660
      Y2              =   660
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6180
      MouseIcon       =   "New.frx":42ABB
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   5970
      Width           =   585
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   315
      MouseIcon       =   "New.frx":42DC5
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   5970
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Height          =   2115
      Left            =   3750
      MouseIcon       =   "New.frx":430CF
      TabIndex        =   6
      Top             =   2070
      Width           =   3030
   End
End
Attribute VB_Name = "NewFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim direction As Boolean
Dim Label As String

Private Sub advanced_eulero_Click()

    algorithm = 1
    Label7.Visible = True
    final_time.Visible = True
    Label8.Caption = "Step"
    Label8.Visible = True
    step.Visible = True
    Label9.Visible = False
    points.Visible = False
    
End Sub

Private Sub eulero_Click()
    
    algorithm = 0
    Label7.Visible = True
    final_time.Visible = True
    Label8.Caption = "Step"
    Label8.Visible = True
    step.Visible = True
    Label9.Visible = False
    points.Visible = False
    
End Sub

Private Sub final_time_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then
        KeyAscii = 7
    End If
    
End Sub

Private Sub Form_Load()

    algorithm = -1
    problem.AddItem "Oscillatore armonico"
    problem.AddItem "Repulsore armonico"
    problem.AddItem "Pendolo"
    problem.AddItem "Oscillatore armonico smorzato"
    problem.AddItem "Sistema di Lotka-Volterra"
    problem.AddItem "Ciclo limite - Equazione di Van Der Pol"
    problem.AddItem "Pendolo doppio"
    problem.AddItem "Sezioni di Poincarè del pendolo forzato"
    Label11.Font.Bold = True
    Label12.Font.Bold = True
    Label13.Font.Bold = True
    Label14.Font.Bold = True
    RoundedCorner Me
    FormFade 220, NewFrm, False
    direction = True
    Label = String(100, " ") + "Creating a new project..."
    Timer1.Enabled = True

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    BlinkLabel Label1, 24, 41, 412, 398, X, Y, 12, 0
    BlinkLabel Label2, 410, 441, 412, 398, X, Y, 12, 0

End Sub

Private Sub Label1_Click()

    If problem.text = "" Or algorithm = -1 Or final_time.text = "" Or step.text = "" Or _
       (algorithm = 2 And points.Visible = True And points.text = "") Or _
       (parameter1.Visible = True And parameter1.text = "") Or _
       (parameter2.Visible = True And parameter2.text = "") Or _
       (parameter3.Visible = True And parameter3.text = "") Or _
       (parameter4.Visible = True And parameter4.text = "") Then
        message.ShowMessage "I campi non sono completi!", "OkOnly"
    Else
        If parameter1.Visible Then parmtr1 = Val(parameter1.text)
        If parameter2.Visible Then parmtr2 = Val(parameter2.text)
        If parameter3.Visible Then parmtr3 = Val(parameter3.text)
        If parameter4.Visible Then parmtr4 = Val(parameter4.text)
        Load Monitor
        Forms(2).Label1.Visible = True
        Forms(2).Label2.Visible = True
        Monitor.Show 1
    End If
    
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Forms.Count <= 2 Then BlinkLabel Label1, 255, 0, 210, 0, X, Y, 12, 0

End Sub

Private Sub Label2_Click()
    
    FormFade 200, Me, True
    Unload NewFrm
    enableForm MainFrm, True, 255

End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    BlinkLabel Label2, 585, 0, 210, 0, X, Y, 12, 0

End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If algorithm = -1 Then help.Visible = True
    
End Sub

Private Sub Label6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If algorithm = -1 Then help.Visible = False
    
End Sub

Private Sub parameter1_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then
        KeyAscii = 7
    End If

End Sub

Private Sub parameter2_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then
        KeyAscii = 7
    End If

End Sub

Private Sub parameter3_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then
        KeyAscii = 7
    End If

End Sub

Private Sub parameter4_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then
        KeyAscii = 7
    End If

End Sub

Private Sub points_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 7
    End If

End Sub

Private Sub problem_Click()

    eulero.Value = False
    advanced_eulero.Value = False
    runge_kutta.Value = False
    Label7.Visible = False
    Label8.Visible = False
    Label9.Visible = False
    final_time.Visible = False
    step.Visible = False
    points.Visible = False
    
    'Problems with one parameter
    If problem.text = "Oscillatore armonico smorzato" Or _
       problem.text = "Ciclo limite - Equazione di Van Der Pol" Then
            eulero.Enabled = True
            advanced_eulero.Enabled = True
            Label10.Visible = True
            Shape3.Visible = True
            Label11.Font = "Symbol"
            Label11.Caption = "m"
            Label11.Visible = True
            parameter1.Visible = True
            Label12.Visible = False
            parameter2.Visible = False
            Label13.Visible = False
            parameter3.Visible = False
            Label14.Visible = False
            parameter4.Visible = False
    'Problem with four parameters
    ElseIf problem.text = "Sistema di Lotka-Volterra" Then
            eulero.Enabled = True
            advanced_eulero.Enabled = True
            Label10.Visible = True
            Shape3.Visible = True
            Label11.Font = "Symbol"
            Label11.Caption = "a"
            Label11.Visible = True
            parameter1.Visible = True
            Label12.Font = "Symbol"
            Label12.Caption = "b"
            Label12.Visible = True
            parameter2.Visible = True
            Label13.Font = "Symbol"
            Label13.Caption = "c"
            Label13.Visible = True
            parameter3.Visible = True
            Label14.Font = "Symbol"
            Label14.Caption = "d"
            Label14.Visible = True
            parameter4.Visible = True
    'Problem with two parameters
    ElseIf problem.text = "Sezioni di Poincarè del pendolo forzato" Then
            eulero.Enabled = False
            advanced_eulero.Enabled = False
            runge_kutta.Value = True
            runge_kutta_Click
            parameter1.Visible = False
            parameter3.Visible = False
            Label10.Visible = True
            Shape3.Visible = True
            Label11.Font = "Microsoft Sans Serif"
            Label11.Caption = "Forcing amplitude"
            Label11.Width = 2100
            Label11.Visible = True
            Label12.Font = "Microsoft Sans Serif"
            Label12.Caption = "A"
            Label12.Visible = True
            parameter2.Visible = True
            Label13.Font = "Microsoft Sans Serif"
            Label13.Caption = "Forcing frequency"
            Label13.Width = 2100
            Label13.Visible = True
            Label14.Font = "Symbol"
            Label14.Caption = "W"
            Label14.Visible = True
            parameter4.Visible = True
    'Problems without parameters
    Else
            eulero.Enabled = True
            advanced_eulero.Enabled = True
            Label10.Visible = False
            Shape3.Visible = False
            Label11.Visible = False
            parameter1.Visible = False
            Label12.Visible = False
            parameter2.Visible = False
            Label13.Visible = False
            parameter3.Visible = False
            Label14.Visible = False
            parameter4.Visible = False
    End If

End Sub

Private Sub runge_kutta_Click()

    algorithm = 2
    Label7.Visible = True
    Label8.Visible = True
    step.Visible = True
    final_time.Visible = True
    final_time.text = ""
    step.text = ""
    If problem <> "Sezioni di Poincarè del pendolo forzato" Then
        Label8.Caption = "Steps"
        points.Visible = True
        Label9.Visible = True
    Else
        Label7.Caption = "Points"
        Label8.Caption = "Cyc/point"
        Label9.Visible = False
        points.Visible = False
    End If
    
End Sub

Private Sub step_KeyPress(KeyAscii As Integer)

    'For Eulero and advanced Eulero step can be real
    If algorithm = 0 Or algorithm = 1 Then
        If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then
            KeyAscii = 7
        End If
    'For Runge-Kutta step means number of steps and must be integer
    Else
        If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
            KeyAscii = 7
        End If
    End If
End Sub

Private Sub Timer1_Timer()
    
    If direction Then
        Label = Mid(Label, 2) & Left(Label, 1)
        TextBox1.text = Label
        If Left(Label, 1) <> " " Then direction = False
    Else
        Label = Right(Label, 1) & Left(Label, 87)
        TextBox1.text = Label
        If Right(Label, 1) <> " " Then direction = True
    End If

End Sub
