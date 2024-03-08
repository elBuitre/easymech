VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MsgWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function ShowMessage(text As String, options As String)
  
    Select Case options
        Case "OkOnly"
            MessageFrm.Label3.Visible = False
            MessageFrm.Label2.Visible = False
            MessageFrm.Label1.Visible = True
        Case "OkCancel"
            MessageFrm.Label1.Visible = False
            MessageFrm.Label2.Visible = True
            MessageFrm.Label3.Visible = True
    End Select
    
    MessageFrm.Phrase.text = text
    MessageFrm.Show 1
    
End Function

