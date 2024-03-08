VERSION 5.00
Begin VB.Form InfoFrm 
   BackColor       =   &H0000FF00&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   2100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10140
   LinkTopic       =   "Form4"
   ScaleHeight     =   2100
   ScaleWidth      =   10140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   465
      Top             =   1365
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Top             =   300
      Width           =   9180
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ver. 1.0"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   495
      TabIndex        =   4
      Top             =   750
      Visible         =   0   'False
      Width           =   9180
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   495
      TabIndex        =   3
      Top             =   1395
      Visible         =   0   'False
      Width           =   9180
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Technical Support: stefano_carniel@hotmail.com"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   495
      TabIndex        =   2
      Top             =   1095
      Visible         =   0   'False
      Width           =   9180
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4800
      TabIndex        =   0
      Top             =   1725
      Width           =   540
   End
End
Attribute VB_Name = "InfoFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Information As String

Private Sub Form_Load()
  
    RoundedCorner Me
    FormFade 180, InfoFrm, False
    
    Information = String(150, " ") + Chr(171) + " Dynatech " + Chr(174) + " the art of mechanical " + Chr(187)
    Timer1.Enabled = True

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    BlinkLabel Label3, 360, 323, 133, 113, X, Y, 12, 0

End Sub

Private Sub Label3_Click()
    
    FormFade 140, Me, True
    Unload InfoFrm
    enableForm MainFrm, True, 255

End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    BlinkLabel Label3, 525, 0, 270, 0, X, Y, 12, 0

End Sub

Private Sub Timer1_Timer()

    Information = Mid(Information, 2) & Left(Information, 1)
    Text1.text = Information
    If Information = String(46, " ") + Chr(171) + " Dynatech " + Chr(174) + " the art of mechanical " + Chr(187) + String(104, " ") Then
        Timer1.Enabled = False
        Label1.Visible = True
        Label2.Caption = Chr(171) + " Copyright " + Chr(169) + " Stefano Carniel " + Chr(187)
        Label2.Visible = True
        Label4.Visible = True
    End If
End Sub
