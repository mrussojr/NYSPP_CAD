Attribute VB_Name = "Module1"
Option Compare Database
Option Explicit

Public Username As String, Password As String

Declare Function WNetGetUser Lib "mpr.dll" _
     Alias "WNetGetUserA" (ByVal lpName As String, _
     ByVal lpUserName As String, lpnLength As Long) As Long

Const NoError = 0                    'The Function call was successful

Function GetUserName() As String

     Dim LUserName As String
     Const lpnLength As Integer = 255
     Dim Status As Integer
     Dim lpName

     ' Assign the buffer size constant to lpUserName.
     LUserName = Space$(lpnLength + 1)

     ' Get the log-on name of the person using product.
     Status = WNetGetUser(lpName, LUserName, lpnLength)

     ' See whether error occurred.
     If Status = NoError Then
          ' This line removes the null character. Strings in C are null-
          ' terminated. Strings in Visual Basic are not null-terminated.
          ' The null character must be removed from the C strings to be used
          ' cleanly in Visual Basic.
          LUserName = Left$(LUserName, InStr(LUserName, Chr(0)) - 1)

     Else
          ' An error occurred.
          MsgBox "Unable to get the name."
          End
     End If

     GetUserName = LUserName

End Function

Function GetDispatchId() As Integer
    Dim NAME As String
    NAME = GetUserName()
    GetDispatchId = DLookup("[DispatchID]", "users", "[Username] = '" & NAME & "'")
End Function

Function GetEjusticeUsername() As String
    Dim NAME As String
    NAME = GetUserName()
    GetEjusticeUsername = DLookup("[EJusticeUsername]", "users", "[Username] = '" & NAME & "'")
End Function

Function GetORI() As String
    Dim NAME As String
    NAME = GetUserName()
    GetORI = DLookup("[ORI]", "users", "[Username] = '" & NAME & "'")
End Function

Function RightNow() As String
    Dim currentTime As String
    Dim currentMinute As String
    Dim currentHour As String
    currentMinute = Minute(Now())
    If currentMinute < 10 Then
        currentMinute = "0" & currentMinute
    End If
    currentHour = Hour(Now())
    If currentHour < 10 Then
        currentHour = "0" & currentHour
    End If
    currentTime = currentHour & ":" & currentMinute
    RightNow = currentTime
End Function

Function ShortTime(d As Date) As String
    Dim currentTime As String
    Dim currentMinute As String
    Dim currentHour As String
    currentMinute = Minute(d)
    If currentMinute < 10 Then
        currentMinute = "0" & currentMinute
    End If
    currentHour = Hour(d)
    If currentHour < 10 Then
        currentHour = "0" & currentHour
    End If
    currentTime = currentHour & ":" & currentMinute
    ShortTime = currentTime
End Function

Function TodayLong() As String
    Dim currentMonth, currentDay As String
    currentDay = Day(Now())
    If currentDay < 10 Then
        currentDay = "0" & currentDay
    End If
    currentMonth = Month(Now())
    If currentMonth < 10 Then
        currentMonth = "0" & currentMonth
    End If
    TodayLong = currentMonth & "/" & currentDay & "/" & Year(Now())
End Function

Function DaysInMonth(MyDate)

   ' This function takes a date as an argument and returns
   ' the total number of days in the month.

   Dim NextMonth, EndOfMonth

   NextMonth = DateAdd("m", 1, MyDate)
   EndOfMonth = NextMonth - DatePart("d", NextMonth)
   DaysInMonth = DatePart("d", EndOfMonth)

End Function

Function CheckTextValue(SCall, UCall) As String
    SCall = CStr(SCall)
   
    If SCall = "JV" Or SCall = "MS" Then
        If IsNull(UCall) = False Then
            SCall = UCall
        End If
    Else
        SCall = SCall
    End If
    CheckTextValue = SCall
End Function

Function CheckWhoCalled(UCall) As Integer
    If IsNull(UCall) = True Or UCall = "" Then
        CheckWhoCalled = 0
    Else
        CheckWhoCalled = 1
    End If
End Function

Public Function EnterInService(pId As Integer) As Integer
    Dim Yes_No As String
    Dim offShld As Integer
            
    Yes_No = MsgBox("Make in service Entry?", vbYesNoCancel, "In Service Entry")
    
    If Yes_No = vbYes Then
        DoCmd.SetWarnings False
        offShld = DLookup("[offShield]", "parksIDs", "[pId] = " & pId)
        DoCmd.RunSQL "INSERT INTO [commlog] ([Date1], [Time1], [SourceCall], [Reason], [Dispatcher]) VALUES ('" & TodayLong() & "', '" & RightNow() & "', " & offShld & ", 'ISV', " & GetDispatchId & ")"
        EnterInService = 1
        DoCmd.SetWarnings True
    ElseIf Yes_No = vbCancel Then
        EnterInService = 0
    Else
        EnterInService = 1
    End If
End Function

Public Function EnterOutService(pId As Integer) As Integer
    Dim Yes_No As String
    Dim offShld As Integer
            
    Yes_No = MsgBox("Make out of service Entry?", vbYesNoCancel, "In Service Entry")
    
    If Yes_No = vbYes Then
        DoCmd.SetWarnings False
        offShld = DLookup("[offShield]", "parksIDs", "[pId] = " & pId)
        DoCmd.RunSQL "INSERT INTO [commlog] ([Date1], [Time1], [SourceCall], [Reason], [Dispatcher]) VALUES ('" & TodayLong() & "', '" & RightNow() & "', " & offShld & ", 'OSV', " & GetDispatchId & ")"
        EnterOutService = 1
        DoCmd.SetWarnings True
    ElseIf Yes_No = vbCancel Then
        EnterOutService = 0
    Else
        EnterOutService = 1
    End If
End Function

Public Function OpenEvent(pId As Integer, Optional flag As String)
    Dim AssId As Long, zoneId As Integer
    Dim secondary As Long
    Dim webAddy As String
        
    AssId = DLookup("[AssId]", "parksIDs", "[pId] = " & pId)
    secondary = DLookup("[SecondaryId]", "parksIDs", "[pId] = " & pId)
    
    'ZoneId = DLookup("[pZone]", "parksIDs", "[pId] = " & pId)
    
'    If AssId = 0 Then
'        AssId = ""
'    End If
    
    webAddy = "http://policeblotter/Events.aspx?slctEventId=&OfficerId=" & pId & "&AssignmentId=" & AssId & "&SecondaryId=" & secondary
    
    DoCmd.OpenForm "Events", , , , , , webAddy & flag
End Function

Public Function OpenExistingEvent(pId As Long, oID As Integer, Optional flag As String)
    Dim webAddy As String
    
    webAddy = "http://policeblotter/Events.aspx?slctEventId=" & pId & "&OfficerId=" & oID
    
    DoCmd.OpenForm "Events", , , , , , webAddy & flag
End Function

Public Function OpenParks(pId As Integer)
    Dim AssId, zoneId As Integer
    Dim webAddy As String
        
    AssId = DLookup("[AssId]", "parksIDs", "[pId] = " & pId)
    zoneId = DLookup("[pZone]", "parksIDs", "[pId] = " & pId)
    
    If AssId = 0 Then
        AssId = ""
    End If
    
    webAddy = "http://policeblotter/Assignment.aspx?slctAssignmentId=" & AssId & "&SecondaryId=&OfficerId=" & pId & "&ZoneId=" & zoneId & "&Page=Assignment"
    
    DoCmd.OpenForm "browser", , , , , , webAddy
    
End Function

Public Function CompareTime(Date1 As Date, Date2 As Variant) As Date
    If IsNull(Date2) = True Then
        CompareTime = Date1
    ElseIf CDate(Date2) > Date1 Then
        CompareTime = Date2
    Else
        CompareTime = Date1
    End If
End Function

'http://maps.googleapis.com/maps/api/staticmap?center=ny+46+and+dixon+dr,+western,+ny&zoom=14&size=400x200&sensor=false&visual_refresh=true&maptype=roadmap&markers=icon:http://goo.gl/6fmKFh%7Ccolor:red%7Clabel:X%7Cny+46+and+dixon+dr,+western,+ny

Function DateLong(Date1 As Date) As String
    Dim currentMonth, currentDay As String
    currentDay = Day(Date1)
    If currentDay < 10 Then
        currentDay = "0" & currentDay
    End If
    currentMonth = Month(Date1)
    If currentMonth < 10 Then
        currentMonth = "0" & currentMonth
    End If
    DateLong = currentMonth & "/" & currentDay & "/" & Year(Date1)
End Function

Function VerifyEmail(emailAddress As String) As Boolean
    If InStr(1, emailAddress, "@") > 1 Then
        If InStr(InStr(1, emailAddress, "@"), emailAddress, ".") > 1 Then
            VerifyEmail = True
        Else
            VerifyEmail = False
        End If
    Else
        VerifyEmail = False
    End If
End Function

Function RandomPassword(nLen As Integer) As String
    Dim i As Long
    Dim nRnd As Double
    Dim sPW As String
    Dim bAdd As Boolean
    
    Dim strTemp As String

    Randomize

    While Len(sPW) < nLen
        
        nRnd = Int(Rnd * 75) + 48
        bAdd = False
        Select Case nRnd
            Case 48 To 57    ' Numeric characters
                bAdd = True
            Case 65 To 90    ' Upper case characters
                bAdd = True
            'Case 97 To 122  ' Lower case characters
            '    bAdd = True
            Case Else        ' Useless characters
                bAdd = False
        End Select
        
        If bAdd Then
            sPW = sPW & Chr(nRnd)
        End If
        
    Wend

    RandomPassword = sPW
End Function

Function IsLoaded(ByVal strFormName As String) As Boolean
 ' Returns True if the specified form is open in Form view or Datasheet view.
    Dim oAccessObject As AccessObject

    Set oAccessObject = CurrentProject.AllForms(strFormName)
    If oAccessObject.IsLoaded Then
        If oAccessObject.CurrentView <> acCurViewDesign Then
            IsLoaded = True
        End If
    End If
End Function

Public Function Pause(NumberOfSeconds As Variant)
    On Error GoTo Error_GoTo

    Dim PauseTime As Variant
    Dim Start As Variant
    Dim Elapsed As Variant

    PauseTime = NumberOfSeconds
    Start = Timer
    Elapsed = 0
    Do While Timer < Start + PauseTime
        Elapsed = Elapsed + 1
        If Timer = 0 Then
            ' Crossing midnight
            PauseTime = PauseTime - Elapsed
            Start = 0
            Elapsed = 0
        End If
        DoEvents
    Loop

Exit_GoTo:
    On Error GoTo 0
    Exit Function
Error_GoTo:
    Debug.Print Err.Number, Err.Description, Erl
    GoTo Exit_GoTo
End Function

Public Function CheckLetters(str As String) As Boolean
    Dim Letters As String, Letter As String
    Dim x As Integer
    Letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    x = 1
    
    Do While x < 53
        Letter = Mid(Letters, x, 1)
        If InStr(1, str, Letter) > 0 Then
            CheckLetters = True
            Exit Function
        End If
    Loop
End Function

'function to convert current passwords to SHA1HASH
Public Function ConvertPasswords()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    'set the database to the current database
    Set db = CurrentDb
    'set the recordset to all of the passwords in the users table that have a length not equal to 40
    Set rs = db.OpenRecordset("SELECT Password FROM users WHERE NOT LEN(Password) = 40")
    
    'loop through all records in the recordset
    While Not rs.EOF
        With rs
            .Edit 'set editing enabled
            .Fields(0).Value = SHA1HASH(rs.Fields(0).Value) 'hash the password value
            .Update 'update the password in the table
            .MoveNext 'move to the next password
        End With
    Wend
    
    'close the recordset and set it to nothing
    rs.Close
    Set rs = Nothing
    
    'advise that passwords have been converted
    MsgBox "Process Complete: All passwords have been converted."
End Function

