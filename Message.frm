VERSION 5.00
Begin VB.Form MessageFrm 
   Appearance      =   0  'Flat
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2145
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2145
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Phrase 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   240
      Left            =   210
      TabIndex        =   0
      Top             =   495
      Width           =   4215
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   3840
      MouseIcon       =   "Message.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1740
      Width           =   630
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   210
      MouseIcon       =   "Message.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1740
      Width           =   285
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
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   2205
      MouseIcon       =   "Message.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1740
      Width           =   270
   End
End
Attribute VB_Name = "MessageFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    
    RoundedCorner Me
    FormFade 190, MessageFrm, False

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    BlinkLabel Label1, 503, 522, 429, 443, X, Y, 15, 0
    BlinkLabel Label2, 370, 389, 429, 443, X, Y, 15, 0
    BlinkLabel Label3, 612, 631, 429, 443, X, Y, 15, 0

End Sub

Private Sub Label1_Click()
    
    GoOn = 0
    FormFade 180, Me, True
    Unload MessageFrm

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    BlinkLabel Label1, 270, 0, 180, 0, X, Y, 15, 0

End Sub

Private Sub Label2_Click()
    
    GoOn = 1
    FormFade 180, Me, True
    Unload MessageFrm

End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    BlinkLabel Label2, 285, 0, 210, 0, X, Y, 15, 0
    
End Sub

Private Sub Label3_Click()
    
    GoOn = 0
    FormFade 140, Me, True
    Unload MessageFrm

End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    BlinkLabel Label3, 630, 0, 210, 0, X, Y, 15, 0

End Sub
