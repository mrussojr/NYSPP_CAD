Attribute VB_Name = "Form_AuthUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub AuthUser_Click()
    'authorize user to use the cad system
    Me.Authorized.Value = 1
    'refresh form
    Me.Requery
    'close form if there are no more users to authorize
    If DCount("[Username]", "usersQuery", "[Authorized] = 0") = 0 Then
        DoCmd.Close acForm, "AuthUsers"
    End If
End Sub

Private Sub Form_Load()
    'check how many users need approval, and if none advise user and close form
    If DCount("[Username]", "usersQuery", "[Authorized] = 0") = 0 Then
        MsgBox "There are no users that need to be Authorized."
        DoCmd.Close acForm, "AuthUsers"
    End If
End Sub
