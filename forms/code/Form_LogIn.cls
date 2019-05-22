Attribute VB_Name = "Form_LogIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'on form load, look up username based on windows profile currently logged in and fill in username field
Private Sub Form_Load()
    Me.Username.Value = GetUserName()
End Sub

Private Sub Label10_Click()
    'open change password form
    DoCmd.OpenForm "ChangePassword"
    'close log in form
    DoCmd.Close acForm, "LogIn"
End Sub

Private Sub Label8_Click()
    'open reset password form
    DoCmd.OpenForm "ResetPassword"
    'close log in form
    DoCmd.Close acForm, "LogIn"
End Sub

Private Sub Label9_Click()
    'open register for cad form
    DoCmd.OpenForm "Register"
    'close the log in form
    DoCmd.Close acForm, "LogIn"
End Sub

'log in button
Private Sub LogIn_Click()
    'set username to username field
    Username = Me.Username.Value
    'set password to password field
    Password = SHA1HASH(Me.Password.Value)
    
    Dim x As Integer
    Dim auth As Integer
    Dim Success As Integer
    Dim uName As String
    Dim uPass As String
    Dim errReas As String
    Dim CurDate As String
    
    'set success to 0 for new log in attempt
    Success = 0
    'set current date to today
    CurDate = Now()
    'set x to how many times the username is in the users table
    x = DCount("[DispatchID]", "users", "[Username] = '" & Username & "'")
        
    'check how many times the user is in the user table, should be 1
    If x = 1 Then
        'look up the password associated with the username entered
        uPass = DLookup("[Password]", "users", "[Username] = '" & Username & "'")
        'check that the password that was entered is equal to the password stored in the users table
        If Password = uPass Then
            'if a match
            'set authorized variable by looking up authorized value from users table for the username entered
            auth = DLookup("[Authorized]", "users", "[Username] = '" & Username & "'")
            'check if the user is authorized to use the system
            If auth = 1 Then
                'if user is authorized, set the success variable
                Success = 1
            Else
                'if not authorized, advise user via message box
                MsgBox "You are not yet authorized to use this system.  If you have just registered please wait until an administrator approves your access."
                'set error reason to not authorized
                errReas = "Not Authorized"
            End If
        Else
            'if passwords do not match, advise user
            MsgBox "Your password is incorrect, please try again or reset your password."
            'set the error reason to incorrect password
            errReas = "Incorrect Password"
        End If
    Else
        'if not 1, then user isn't registered for the cad, advise user of such
        MsgBox "You are not registered to use this system.  Please create a username and password and wait for authorization."
        'set the error reason as not registered
        errReas = "Not Registered"
    End If
    
    'set sql injection warnings to off
    DoCmd.SetWarnings False
    'execute sql to insert login information to the login table
    DoCmd.RunSQL "INSERT INTO logins (Username, LogInDate, ErrorReason) VALUES ('" & Username & "', #" & CurDate & "#, '" & errReas & "')"
    
    'check login success variable
    If Success = 1 Then
        'clear the ejustice logged in field in the users table
        DoCmd.RunSQL "UPDATE users SET LoggedInEJustice = 0, EJusticeLinkId = ''  WHERE Username = '" & Username & "'"

        'run the update assignment procedure
        UpdateAssignments

        DoCmd.Close acForm, "LogIn"
    Else
        'if unsuccessful, clear the password field so the user can retry logging in
        Me.Password.Value = Null
    End If
    
    'set the sql injection warnings to on
    'DoCmd.SetWarnings True
End Sub
