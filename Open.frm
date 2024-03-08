VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form OpenFrm 
   Appearance      =   0  'Flat
   BackColor       =   &H0080FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   5865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7095
   LinkTopic       =   "Form2"
   Picture         =   "Open.frx":0000
   ScaleHeight     =   5865
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
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
      Height          =   2625
      Left            =   4140
      MultiSelect     =   1  'Simple
      Pattern         =   "*.bmp; *.txt"
      TabIndex        =   4
      Top             =   1440
      Width           =   2580
   End
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
      Left            =   375
      TabIndex        =   3
      Top             =   2400
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
      TabIndex        =   9
      Top             =   4305
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
   Begin MSForms.Label Label5 
      Height          =   195
      Left            =   375
      TabIndex        =   8
      Top             =   2130
      Width           =   2235
      BackColor       =   8454143
      VariousPropertyBits=   8388627
      Caption         =   "Directory"
      Size            =   "3942;344"
      FontName        =   "Microsoft Sans Serif"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label4 
      Height          =   225
      Left            =   4140
      TabIndex        =   7
      Top             =   1155
      Width           =   2235
      BackColor       =   8454143
      VariousPropertyBits=   8388627
      Caption         =   "Saved projects"
      Size            =   "3942;397"
      FontName        =   "Microsoft Sans Serif"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label3 
      Height          =   165
      Left            =   375
      TabIndex        =   6
      Top             =   1155
      Width           =   2235
      BackColor       =   8454143
      VariousPropertyBits=   8388627
      Caption         =   "Drive"
      Size            =   "3942;291"
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
      TabIndex        =   5
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
      MouseIcon       =   "Open.frx":427B1
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   5310
      Width           =   585
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Open"
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
      MouseIcon       =   "Open.frx":42ABB
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   5310
      Width           =   465
   End
End
Attribute VB_Name = "OpenFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim direction As Boolean
Dim Label As String

Private Sub Dir1_Change()

    File1.Path = Dir1.Path
    TextBox2.text = Dir1.Path
    
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

Private Sub File1_Click()

    TextBox2.text = Dir1.Path & "\" & File1.FileName
    
End Sub

Private Sub File1_DblClick()
    
    SavedProjectName = TextBox2.text
    OpenSavedProject
    
End Sub

Private Sub Form_Load()

    RoundedCorner Me
    FormFade 220, OpenFrm, False
    direction = True
    Label = String(100, " ") + "Opening saved project..."
    Timer1.Enabled = True
    TextBox2.text = Left(Drive1.Drive, 2) + "\"

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    BlinkLabel Label1, 24, 55, 368, 354, X, Y, 12, 0
    BlinkLabel Label2, 410, 441, 368, 354, X, Y, 12, 0

End Sub

Private Sub Label1_Click()
        
    SavedProjectName = TextBox2.text
    OpenSavedProject
    
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Forms.Count <= 2 Then BlinkLabel Label1, 465, 0, 210, 0, X, Y, 12, 0

End Sub

Private Sub Label2_Click()
    
    FormFade 200, Me, True
    Unload OpenFrm
    enableForm MainFrm, True, 255

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
