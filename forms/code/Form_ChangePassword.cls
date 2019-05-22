Attribute VB_Name = "Form_ChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Cancel_Click()
    If Me.OpenArgs = "CAD" Then
    
    Else
        'if the user got here from the cad by click on change password, then return to cad, else return to log in screen
        DoCmd.OpenForm "LogIn"
    End If
    
    'close the change password form
    DoCmd.Close acForm, "ChangePassword"
End Sub

Private Sub Form_Load()
    'get the users cad log in username from the current windows profile
    Me.Username.Value = GetUserName()
End Sub

Private Sub LogIn_Click()
    Dim uName As String, sql As String
    
    'get the cad username using the windows profile currently logged in
    uName = GetUserName()
    
    If Not Me.Username.Value = uName Then
        'if the username loaded into the form is not the same as the username that belongs with the currently logged in profile, advise user and cancel change password
        MsgBox "You must be currently logged in as " & uName & " to change your password."
        Exit Sub
    End If
    
    Dim x As Integer
    
    'count the number of users with entered user name and password, should return 1 if username and password match
    x = DCount("[ID]", "users", "[Username] = '" & uName & "' AND [Password] = '" & SHA1HASH(Me.OldPass.Value) & "'")

    If x = 1 Then
        If Len(Me.NewPass.Value) >= 6 Then 'check to make sure the new password is at least 6 characters in length
            If Not SHA1HASH(Me.NewPass.Value) = DLookup("[Password]", "users", "[Username] = '" & uName & "'") Then 'check to make sure the new password is not the same as the old password
                'set sql injection warnings off
                DoCmd.SetWarnings False
                'generate sql to update the users password from the old password to the new password
                sql = "UPDATE users SET users.Password = '" & SHA1HASH(Me.NewPass.Value) & "' WHERE (users.Username) = '" & uName & "'"
                'execute sql to update users password
                DoCmd.RunSQL sql
                'reenable sql injection warnings
                DoCmd.SetWarnings True
                'let the user know that their password has in fact been changed
                MsgBox "You password has successfully been changed."
                If Not Me.OpenArgs = "CAD" Then 'check to see if user got here from the cad or the login screen
                    'if not the cad, then send the user back to the log in screen
                    DoCmd.OpenForm "LogIn"
                End If
                'close the change password form
                DoCmd.Close acForm, "ChangePassword"
            Else
                'advise the user that their new password must be different from their old password
                MsgBox "Your new password must be different than your old password!"
                'set focus to the new password input so they can try again
                Me.NewPass.SetFocus
                'exit without saving any information
                Exit Sub
            End If
        Else
            'advise the user that their password must be at least 6 characters in length
            MsgBox "Your password must be at least 6 characters long!"
            'set the focus to the new password input so they may try again
            Me.NewPass.SetFocus
            'exit without saving any information
            Exit Sub
        End If
    Else
        'if the password and username do not match advise user to try again or contact cad administrator
        MsgBox "You have not provided the correct password on file for your username, please try again, or contact the CAD Administrator for assistance."
        'set the old password input to have the focus so the user can retry
        Me.OldPass.SetFocus
        'exit without saving any information
        Exit Sub
    End If
End Sub
