VERSION 5.00
Begin VB.Form MainFrm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Integrator"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Main.frx":08CA
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2520
      Top             =   8535
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   462
      X2              =   800
      Y1              =   41
      Y2              =   41
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   -2
      X2              =   336
      Y1              =   41
      Y2              =   41
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Open"
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
      Left            =   5745
      MouseIcon       =   "Main.frx":438FD
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   150
      Width           =   510
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "New"
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
      Left            =   240
      MouseIcon       =   "Main.frx":43C07
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   150
      Width           =   465
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   270
      TabIndex        =   3
      Top             =   8460
      Width           =   2070
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "About"
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
      Left            =   5715
      MouseIcon       =   "Main.frx":43F11
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   8445
      Width           =   570
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
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
      Left            =   11400
      MouseIcon       =   "Main.frx":4421B
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   150
      Width           =   360
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Change background"
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
      Left            =   9840
      MouseIcon       =   "Main.frx":44525
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   8460
      Width           =   1920
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    GoOn = -1
    If App.PrevInstance Then
        message.ShowMessage "Programma già in esecuzione!", "OkOnly"
        Do
            If GoOn = 0 Then End
        Loop
    End If
    Timer1.Enabled = True
    activeBackground = 0
    PathToShell = App.Path & "\shell\notepad.exe"
    Label4.Caption = Format(Now, "hh:mm:ss")
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Forms.Count = 1 Then
        BlinkLabel Label1, 790, 647, 572, 564, X, Y, 12, 15
        BlinkLabel Label2, 790, 750, 30, 10, X, Y, 12, 15
        BlinkLabel Label3, 421, 380, 572, 564, X, Y, 12, 15
        BlinkLabel Label5, 48, 16, 30, 10, X, Y, 12, 15
        BlinkLabel Label6, 417, 382, 30, 10, X, Y, 12, 15
    End If
    
End Sub

Private Sub Label1_Click()
   
    Dim backGroundName As String
       
    activeBackground = (activeBackground + 1) Mod 8
    backGroundName = App.Path & "\Backgrounds\background" & CStr(activeBackground) & ".jpg"
    MainFrm.Picture = LoadPicture(backGroundName)

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Forms.Count = 1 Then
        If X <= 1995 And X >= 0 And Y <= 270 And Y >= 0 Then
            Label1.ForeColor = RGB(255, 0, 0)
        Else
            Label1.ForeColor = RGB(255, 255, 255)
        End If
    End If

End Sub

Private Sub Label2_Click()
    
    FormFade 255, Me, True
    Unload Me
    End
    
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Forms.Count = 1 Then
        If X <= 600 And X >= 0 And Y <= 270 And Y >= 0 Then
            Label2.ForeColor = RGB(255, 0, 0)
        Else
            Label2.ForeColor = RGB(255, 255, 255)
        End If
    End If

End Sub

Private Sub Label3_Click()
   
    enableForm MainFrm, False, 180
    InfoFrm.Show 1
   
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Forms.Count = 1 Then
        If X <= 1995 And X >= 0 And Y <= 270 And Y >= 0 Then
            Label3.ForeColor = RGB(255, 0, 0)
        Else
            Label3.ForeColor = RGB(255, 255, 255)
        End If
    End If

End Sub

Private Sub Label5_Click()

    enableForm MainFrm, False, 180
    NewFrm.Show 1
    
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Forms.Count = 1 Then
        If X <= 600 And X >= 0 And Y <= 270 And Y >= 0 Then
            Label5.ForeColor = RGB(255, 0, 0)
        Else
            Label5.ForeColor = RGB(255, 255, 255)
        End If
    End If

End Sub

Private Sub Label6_Click()
    
    enableForm MainFrm, False, 180
    OpenFrm.Show 1

End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Forms.Count = 1 Then
        If X <= 600 And X >= 0 And Y <= 270 And Y >= 0 Then
            Label6.ForeColor = RGB(255, 0, 0)
        Else
            Label6.ForeColor = RGB(255, 255, 255)
        End If
    End If

End Sub

Private Sub Timer1_Timer()

    Label4.Caption = Format(Now, "hh:mm:ss")
    
End Sub
