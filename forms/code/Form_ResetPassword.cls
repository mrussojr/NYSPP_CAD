Attribute VB_Name = "Form_ResetPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'cancel button
Private Sub Cancel_Click()
    'open the log in form
    DoCmd.OpenForm "LogIn"
    'close the reset password form
    DoCmd.Close acForm, "ResetPassword"
End Sub

Private Sub Form_Load()
    'look up the username associated with the windows profile currently logged in and set the username field
    Me.Username.Value = GetUserName()
End Sub

'reset password button
Private Sub LogIn_Click()
    'check that the value of the username field is the same that is retreived from the windows profile currently logged,
    'this prevents unauthorized users from changing anothers password
    If Not Me.Username.Value = GetUserName() Then
        'if the usernames do not match, advise the user of such and exit the procedure without resetiing the password
        MsgBox "You must be currently logged in as " & GetUserName() & " to reset your password."
        Exit Sub
    End If
    
    Dim x As Integer
    Dim NewPass As String
    
    'count to make sure that the username and email match
    x = DCount("[ID]", "users", "[Username] = '" & GetUserName() & "' AND [Email] = '" & Me.emailAddress.Value & "'")

    If x = 1 Then
        'if the username and email match
        'create a new password of 6 characters in length randomly generated
        NewPass = RandomPassword(6)
        'set sql injection warnings to false
        DoCmd.SetWarnings False
        'run sql to update the users password to the newly generated random password
        DoCmd.RunSQL "UPDATE users SET users.Password = '" & SHA1HASH(NewPass) & "' WHERE (users.Username) = '" & GetUserName() & "'"
        'reenable the sql injection warnings
        DoCmd.SetWarnings True
        'send the user a new pasword in an email
        SendNewPassword Me.emailAddress.Value, NewPass
        'advise the user that the password has been reset, and that an email has been sent with the password
        MsgBox "Your password has successfully been reset, please check your inbox for your new password."
        'open the log in form
        DoCmd.OpenForm "LogIn"
        'close the reset password form
        DoCmd.Close acForm, "ResetPassword"
    Else
        'if the username and email do not match, advise the user
        MsgBox "You have not provided the correct email address on file for your username, please try again, or contact the CAD Administrator for assistance."
        'set the focus to the email address field
        Me.emailAddress.SetFocus
        'exit the procedure for the user to correct the email address
        Exit Sub
    End If
End Sub
