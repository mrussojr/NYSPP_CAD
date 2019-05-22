Attribute VB_Name = "Module3"
Option Compare Database

Function SendRegistrationEmail(uName As String, eMailAddy As String)
    Dim olApp As Object
    Dim objMail As Object
    Dim toAddys As String
    Dim dbs As DAO.Database
    Dim rsSQL As DAO.Recordset
    Dim strSql As String
    
    toAddys = ""

    Set dbs = CurrentDb

    'Open a snapshot-type Recordset based on an SQL statement
    strSql = "SELECT * FROM users WHERE Admin = 1"
    Set rsSQL = dbs.OpenRecordset(strSql, dbOpenSnapshot)
    
    Do While Not rsSQL.EOF
        toAddys = toAddys & rsSQL.Fields("Email").Value & ";"
        rsSQL.MoveNext
    Loop
    
    On Error Resume Next 'Keep going if there is an error
    
    Set olApp = GetObject(, "Outlook.Application") 'See if Outlook is open
    
    If Err Then 'Outlook is not open
        Set olApp = CreateObject("Outlook.Application") 'Create a new instance of Outlook
    End If
    
    'Create e-mail item
    Set objMail = olApp.CreateItem(olMailItem)
    
    With objMail
        'Set body format to HTML
        .BodyFormat = olFormatHTML
        .To = toAddys
        .Subject = "NYSPP CAD Registration"
        .HTMLBody = uName & " with email address " & eMailAddy & " is requesting access to the NYSPP CAD System.  Upon verification of this users credentials, please log in to the CAD System and authorize this user."
        .send
    End With
End Function

Function SendNewPassword(eMailAddy As String, NewPass As String)
    Dim olApp As Object
    Dim objMail As Object
    Dim toAddys As String
    Dim dbs As DAO.Database
    Dim rsSQL As DAO.Recordset
    Dim strSql As String
    
    toAddys = ""

    Set dbs = CurrentDb

    'Open a snapshot-type Recordset based on an SQL statement
    strSql = "SELECT * FROM users WHERE Admin = 1"
    Set rsSQL = dbs.OpenRecordset(strSql, dbOpenSnapshot)
    
    Do While Not rsSQL.EOF
        toAddys = toAddys & rsSQL.Fields("Email").Value & ";"
        rsSQL.MoveNext
    Loop
    
    On Error Resume Next 'Keep going if there is an error
    
    Set olApp = GetObject(, "Outlook.Application") 'See if Outlook is open
    
    If Err Then 'Outlook is not open
        Set olApp = CreateObject("Outlook.Application") 'Create a new instance of Outlook
    End If
    
    'Create e-mail item
    Set objMail = olApp.CreateItem(olMailItem)
    
    With objMail
        'Set body format to HTML
        .BodyFormat = olFormatHTML
        .To = eMailAddy
        .BCC = toAddys
        .Subject = "NYSPP CAD Registration"
        .HTMLBody = "You have requested to reset your password.  Your temporary password is " & NewPass & ", once you log in you can change your password."
        .send
    End With
End Function

Function SendAlarmEmail(eMailAddy As String, Narr As String, SubjectString As String)
    Dim olApp As Object
    Dim objMail As Object
    Dim toAddys As String
    
    toAddys = ""

    On Error Resume Next 'Keep going if there is an error
    
    Set olApp = GetObject(, "Outlook.Application") 'See if Outlook is open
    
    If Err Then 'Outlook is not open
        Set olApp = CreateObject("Outlook.Application") 'Create a new instance of Outlook
    End If
    
    'Create e-mail item
    Set objMail = olApp.CreateItem(olMailItem)
    
    With objMail
        'Set body format to HTML
        .BodyFormat = olFormatHTML
        .To = eMailAddy
        .Subject = SubjectString
        .HTMLBody = Narr
        .send
    End With
End Function

Function SendCourtEmail(eMailAddy As String)
    Dim olApp As Object
    Dim objMail As Object
    Dim toAddys As String
    
    toAddys = ""

    On Error Resume Next 'Keep going if there is an error
    
    Set olApp = GetObject(, "Outlook.Application") 'See if Outlook is open
    
    If Err Then 'Outlook is not open
        Set olApp = CreateObject("Outlook.Application") 'Create a new instance of Outlook
    End If
    
    'Create e-mail item
    Set objMail = olApp.CreateItem(olMailItem)
    
    With objMail
        .display
        'Set body format to HTML
        .BodyFormat = olFormatHTML
        .To = eMailAddy
    End With
End Function
