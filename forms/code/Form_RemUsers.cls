Attribute VB_Name = "Form_RemUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

'activate / deactivate user button
Private Sub ActUser_Click()
    'check the activated status of the selected user
    If Me.Active.Value = 1 Then
        'if the user is activated, deactivate the user
        Me.Active.Value = 0
        'set the button caption to activate user
        Me.ActUser.Caption = "Activate User"
    Else
        'if the user is deactivated, activate the user
        Me.Active.Value = 1
        'set the button caption to deactivated the user
        Me.ActUser.Caption = "De-Activate User"
    End If
    
End Sub

Private Sub Detail_Paint()
    'chexk the activated status of the selected user
    If Me.Active.Value = 1 Then
        'if the user is activated, set the button caption to deactivate user
        Me.ActUser.Caption = "De-Activate User"
    Else
        'if the user is deactivated, set the button caption to activate the user
        Me.ActUser.Caption = "Activate User"
    End If
End Sub
