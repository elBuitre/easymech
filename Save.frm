VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form SaveFrm 
   Appearance      =   0  'Flat
   BackColor       =   &H0080FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7095
   LinkTopic       =   "Form2"
   Picture         =   "Save.frx":0000
   ScaleHeight     =   4260
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1665
      Left            =   4110
      TabIndex        =   3
      Top             =   1455
      Width           =   2445
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   375
      TabIndex        =   2
      Top             =   1455
      Width           =   2445
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   90
      Top             =   165
   End
   Begin MSForms.TextBox TextBox2 
      Height          =   330
      Left            =   390
      TabIndex        =   7
      Top             =   3480
      Width           =   6330
      VariousPropertyBits=   746604563
      Size            =   "11165;582"
      FontName        =   "Microsoft Sans Serif"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label4 
      Height          =   210
      Left            =   4125
      TabIndex        =   6
      Top             =   1200
      Width           =   2445
      BackColor       =   8454143
      VariousPropertyBits=   8388627
      Caption         =   "Directory"
      Size            =   "4313;370"
      FontName        =   "Microsoft Sans Serif"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label3 
      Height          =   240
      Left            =   390
      TabIndex        =   5
      Top             =   1200
      Width           =   2445
      BackColor       =   8454143
      VariousPropertyBits=   8388627
      Caption         =   "Drive"
      Size            =   "4313;423"
      FontName        =   "Microsoft Sans Serif"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   285
      Left            =   480
      TabIndex        =   4
      Top             =   210
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
      Left            =   6150
      MouseIcon       =   "Save.frx":427B1
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3900
      Width           =   585
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
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
      Left            =   360
      MouseIcon       =   "Save.frx":42ABB
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   3900
      Width           =   465
   End
End
Attribute VB_Name = "SaveFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim direction As Boolean
Dim Label As String

Private Sub Dir1_Change()

    TextBox2.text = Dir1.Path & "\"
    
End Sub

Private Sub Drive1_Change()

On Error GoTo ErrMng

    Dir1.Path = Drive1.Drive
    TextBox2.text = Left(Drive1.Drive, 2) + "\"
    
ErrMng:
    If Err = 68 Then
        Unload Me
        message.ShowMessage "Unità non presente!", "OkOnly"
    End If
    
End Sub

Private Sub Form_Load()

    RoundedCorner Me
    FormFade 220, SaveFrm, False
    direction = True
    Label = String(100, " ") + "Saving project..."
    Timer1.Enabled = True
    TextBox2.text = Left(Drive1.Drive, 2) + "\"

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    BlinkLabel Label1, 24, 55, 368, 354, X, Y, 12, 0
    BlinkLabel Label2, 410, 441, 368, 354, X, Y, 12, 0

End Sub

Private Sub Label1_Click()

    Dim file, ctr As Integer

    If Right(TextBox2.text, 1) <> "\" Then
        If Right(TextBox2.text, 4) <> ".bmp" Then
            If Right(TextBox2.text, 4) = ".txt" Then
                TextBox2.text = Left(TextBox2.text, Len(TextBox2.text) - 4) & ".bmp"
            Else
                TextBox2.text = TextBox2.text & ".bmp"
            End If
        End If
        SavePicture Monitor.Image, TextBox2.text
        file = FreeFile
        TextBox2.text = Left(TextBox2.text, Len(TextBox2.text) - 4) & ".txt"
        Open TextBox2.text For Output As file
        Print #file, "Computed problem: " & NewFrm.problem.text
        Select Case algorithm
            Case 0
                Print #file, "Integration system: Eulero"
            Case 1
                Print #file, "Integration system: Advanced Eulero"
            Case 2
                Print #file, "Integration system: Runge-Kutta"
        End Select
        If NewFrm.problem = "Sezioni di Poincarè del pendolo forzato" Then
            Print #file, "Points: "; Str(NewFrm.step.text)
            Print #file, "Cycles per point: "; Str(NewFrm.points.text)
            Print #file, "Forcing amplitude: "; Str(parmtr2)
            Print #file, "Forcing frequency: "; Str(parmtr4)
        Else
            Print #file, "Final time: "; NewFrm.final_time.text
            If NewFrm.points.Visible Then
                Print #file, "Steps: "; NewFrm.step.text
                Print #file, "Points: "; NewFrm.points.text
            Else
                Print #file, "Step: "; NewFrm.step.text
            End If
        End If
        If NewFrm.problem = "Oscillatore armonico smorzato" Then
            Print #file, "Coefficiente di smorzamento: "; Str(parmtr1)
        End If
        If NewFrm.problem = "Sistema di Lotka-Volterra" Then
            Print #file, "alfa:  "; Str(parmtr1)
            Print #file, "beta:  "; Str(parmtr2)
            Print #file, "gamma: "; Str(parmtr3)
            Print #file, "delta: "; Str(parmtr4)
        End If
        If NewFrm.problem = "Ciclo limite - Equazione di Van Der Pol" Then
            Print #file, "Coefficiente beta: "; Str(parmtr1)
        End If
        
        Print #file,
        Print #file, Spc(6); "Xo"; Spc(6); "|"; Spc(10); "Yo"
        Print #file, "--------------|--------------------"
        
        'Write initial conditions to file
        For ctr = 0 To counting - 1
            Print #file, Str(FormatNumber(init_x(ctr), 4, 0)); Tab; "|"; Tab; Str(FormatNumber(init_y(ctr), 4, 0))
        Next ctr
        
        Close file
        FormFade 200, Me, True
        Unload SaveFrm
        enableForm Monitor, True, 255

    Else
        message.ShowMessage "Type the name of file!", "OkOnly"
    End If
    
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    BlinkLabel Label1, 465, 0, 210, 0, X, Y, 12, 0

End Sub

Private Sub Label2_Click()
    
    FormFade 200, Me, True
    Unload SaveFrm
    enableForm Monitor, True, 255

End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    BlinkLabel Label2, 585, 0, 210, 0, X, Y, 12, 0

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
