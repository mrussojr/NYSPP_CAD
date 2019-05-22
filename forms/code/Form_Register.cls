Attribute VB_Name = "Form_Register"
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
    'close the register form
    DoCmd.Close acForm, "Register"
End Sub

Private Sub Form_Load()
    'when the form loads, get the username using the windows profile and fill in the username field
    Me.Username.Value = GetUserName()
End Sub

'registration button
Private Sub LogIn_Click()
    Dim uName As String, pass As String, eMailAddy As String, EJustice As String, insertSql As String
    Dim dId As Integer
    
    'check the value in the username field
    If IsNull(Me.Username.Value) = False Then
        'if the username is filled in, set the username variable to the username field
        uName = Me.Username.Value
    Else
        'if the username field is not filled in, advise the user to enter their username
        MsgBox "Please enter your Windows username."
        'set the focus to the username field
        Me.Username.SetFocus
        'exit the procedure for the user to enter their user name
        Exit Sub
    End If
    
    'check that the password has been filled in
    If IsNull(Me.Password.Value) = False Or Me.Password.Value <> "" Then
        'if there is a password entered, check the length of the password in the password field
        If Len(Me.Password.Value) >= 6 Then
            'if the password length is 6 characters or more, set the password variable to the password entered in the form
            pass = SHA1HASH(Me.Password.Value)
        Else
            'if the password is not at least 6 characters in length, advise the user that the minimum password length is 6
            MsgBox "Your password must be a minimum of 6 characters in length."
            'set the focus to the password field
            Me.Password.SetFocus
            'exit the procedure for the user to enter a new password
            Exit Sub
        End If
    Else
        'if the password field is blank, advise the user to create a password
        MsgBox "Please create a password to log in to the system."
        'exit the procedure so that the user can create a password
        Exit Sub
    End If
    
    'check that the email address field is filled in
    If IsNull(Me.emailAddress.Value) = False Then
        'if an email address has been entered, check that it is a valid email
        If VerifyEmail(Me.emailAddress.Value) = True Then
            'if the email address is valid, set the email address variable to the email address entered on the form
            eMailAddy = Me.emailAddress.Value
        Else
            'if the email address is not valid, advise the user to enter a valid email address
            MsgBox "Please enter a valid email address."
            'set the email address field to have the focus
            Me.emailAddress.SetFocus
            'exit the procedure so that the user can enter a valid email address
            Exit Sub
        End If
    Else
        'if no email has been entered, advise the user to enter their work email address
        MsgBox "Please enter your work email address."
        'set the email address field to have the focus
        Me.emailAddress.SetFocus
        'exit the procedure so the user can enter their email address
        Exit Sub
    End If
    
    'check the ejustice username has been entered
    If IsNull(Me.EJusticeID.Value) = False Then
        'if the ejustice username has been filled in, set the ejustice variable to be the username entered on the form
        EJustice = Me.EJusticeID.Value
    Else
        'if the ejustice username has not been filled in, advise the user that the their ejustice username must be entered
        MsgBox "Please enter your E-Justice Portal username."
        'set the ejustice username field to have the focus
        Me.EJusticeID.SetFocus
        'exit the procedure so that the user can enter their ejustice username
        Exit Sub
    End If
    
    'check the dispatch / shield has been entered
    If IsNull(Me.DispatchID.Value) = False Then
        'if there is an id entered, set the id variable to the id entered on the form
        dId = CInt(Me.DispatchID.Value)
    Else
        'if there is not an id entered, advise the user to enter their id
        MsgBox "Please enter your Dispatch ID or Shield #."
        'set the focus to the id field on the form
        Me.DispatchID.SetFocus
        'exit the procedure so that the user can enter their id
        Exit Sub
    End If
    
    'disable the sql injection warnings
    DoCmd.SetWarnings False
    'generate the insert sql to insert the new user
    insertSql = "INSERT INTO users (DispatchID, Username, EJusticeUsername, [Password], Email) VALUES (" & dId & ", '" & uName & "', '" & EJustice & "', '" & pass & "', '" & eMailAddy & "')"

    'execute the insert sql to insert the new user into the users table
    DoCmd.RunSQL insertSql
    'reenable the sql injection warnings
    DoCmd.SetWarnings True
    
    'send an email to the cad admins that a new user has registered and they need to be authorized to use the system
    SendRegistrationEmail uName, eMailAddy
    
    'advise the user that they have successfully registered to use the cad
    MsgBox "You have successfully registered to use the NYSPP CAD System.  Once your credentials have been verified, you will be authorized to use the system."
    
    'open the log in form
    DoCmd.OpenForm "LogIn"
    'close the registration form
    DoCmd.Close acForm, "Register"
End Sub
