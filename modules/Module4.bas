Attribute VB_Name = "Module4"
Option Compare Database

'Ejustice Functions

Private loggedIn As Boolean
Private DataRun As Boolean
Private LinkFound As Boolean
Private saveRec As Boolean
Public ParksPage As Integer

Function IsLoggedIn() As Boolean
    Dim loggedIn As Integer
    loggedIn = Nz(DLookup("[LoggedInEJustice]", "users", "[Username] = '" & GetUserName() & "'"), 0)
    If loggedIn = 0 Then
        IsLoggedIn = False
    ElseIf loggedIn = 1 Then
        IsLoggedIn = True
    End If
End Function

Function UpdateLogIn()
    DoCmd.SetWarnings False
    DoCmd.RunSQL "UPDATE users SET LoggedInEJustice = 1 WHERE Username = '" & GetUserName() & "'"
    DoCmd.SetWarnings True
End Function

Function HasDataBeenRun() As Boolean
    HasDataBeenRun = DataRun
End Function

Function RunData()
    DataRun = True
End Function

Function ClearData()
    DataRun = False
End Function

Function HasLinkBeenFound() As Boolean
    HasLinkBeenFound = LinkFound
End Function

Function FoundLink()
    LinkFound = True
End Function

Function OpenEJustice(Arg As String)
    If IsLoggedIn = False Then
        DoCmd.OpenForm "EJustice"
    Else
        If IsLoaded("EJustice") = True Then
            DoCmd.Close acForm, "EJustice"
        End If
        
        DoCmd.OpenForm "EJustice", , , , , , Arg
        Dim Addy As String, linkPt1 As String
        linkPt1 = DLookup("[EJusticeLinkId]", "users", "[Username] = '" & GetUserName() & "'")
        Addy = linkPt1 & "L2dJQSEvUUt3QS9ZQnZ3LzZfMzBHMjA4UkdJNTBCNDAyNTExNDRWMzAwSTY!/"
        
        'Forms!EJustice.WebBrowser0.Navigate Addy
    End If
End Function

'Event Entry Page Functions

Public Function ReadyToSaveRecord() As Boolean
    ReadyToSaveRecord = saveRec
End Function

Public Function OkToSaveRecord()
    saveRec = True
End Function

Public Function CreateNewRecord()
    saveRec = False
End Function

Public Function sanitize(Location As String) As String
    Dim removeCases(0 To 2) As String
    Dim checkCases(0 To 1, 0 To 10) As String
    Dim x As Integer
    
    removeCases(0) = "(LEASED)"
    removeCases(1) = "(CLOSED FOR 2010)"
    removeCases(2) = "(SATTELLITE OF SLSP)"
    
    checkCases(0, 0) = "STATE PARK"
    checkCases(1, 0) = "STATE PARK"
    checkCases(0, 1) = "MARINE PARK"
    checkCases(1, 1) = "BOAT LAUNCH"
    checkCases(0, 2) = "PARK"
    checkCases(1, 2) = "STATE PARK"
    checkCases(0, 3) = "HIST.SITE"
    checkCases(1, 3) = "HISTORIC SITE"
    checkCases(0, 4) = "HIST. SITE"
    checkCases(1, 4) = "HISTORIC SITE"
    checkCases(0, 5) = "HISTORIC SITE"
    checkCases(1, 5) = "HISTORIC SITE"
    checkCases(0, 6) = "STATE HISTORIC SITE"
    checkCases(1, 6) = "HISTORIC SITE"
    checkCases(0, 7) = "(HISTORIC SITE)"
    checkCases(1, 7) = "HISTORIC SITE"
    checkCases(0, 8) = "SITE"
    checkCases(1, 8) = "HISTORIC SITE"
    checkCases(0, 9) = "BOAT LAUNCH"
    checkCases(1, 9) = "BOAT LAUNCH"
    checkCases(0, 10) = "GOLF COURSE"
    checkCases(1, 10) = "GOLF COURSE"

    Location = UCase(Location)
    
    x = 0
    Do While x < 3
        If InStr(1, Location, removeCases(x)) > 0 Then
            Location = Replace(Location, removeCases(x), "")
        End If
        x = x + 1
    Loop
    
    x = 0
    Do While x < 11
        If InStr(1, Location, checkCases(0, x)) > 0 Then
            Location = Replace(Location, checkCases(0, x), checkCases(1, x))
            Exit Do
        End If
        x = x + 1
    Loop
    
    Location = Trim(Location)
    
    If x = 11 Then
        Location = Location + " STATE PARK"
    End If
    
    sanitize = Location
End Function

'Event Functions
Public Function UpdateLastTime(Time As String, dat As String, x As Integer)
    Dim tm As Date, dt As Date
    tm = CDate(Time)
    dt = CDate(dat)
    
    DoCmd.RunSQL "UPDATE events SET LastTime=#" & tm & "#, LastDate=#" & dt & "#, ColorCode=1 WHERE ID = " & x
    RqryChildren
End Function

Public Function ClearEvent(x As Integer)
    DoCmd.SetWarnings False
    DoCmd.RunSQL "UPDATE events SET Active=0 WHERE ID = " & x
    DoCmd.SetWarnings True
    RqryChildren
End Function

Public Function RqryChildren()
    Forms!CAD!Child0.Form!ActiveEventHolder.Form.Requery
    Forms!CAD.Child4.Requery
End Function

Public Function ReturnDateValue(Date1 As Date, Time1 As Date) As Long
    Dim Date2 As Date
    
    Date2 = CDate(Date1 & " " & Time1)
    
    ReturnDateValue = DateDiff("n", Date2, Now())
End Function

'get the schedule filename from the ScheduleDates table
Public Function GetFileName() As String
    Dim fName As String
    Dim sDate As Date, eDate As Date, tDate As Date
    Dim db As Database
    Dim recSet As DAO.Recordset
    
    'set todays date
    tDate = Date
    
    'set the database as the current database
    Set db = CurrentDb
    
    With db
        'set the recordset to the ScheduleDates table
        Set recSet = .OpenRecordset("ScheduleDates", dbOpenSnapshot)
        
        With recSet
            'set loop to run as long as there are records
            Do While Not .EOF
                'set the start date to the value in the begindate field
                sDate = .Fields("BeginDate").Value
                'set the end date to the value in the enddate field
                eDate = .Fields("EndDate").Value

                'check to see if today is between the begin and end dates
                If tDate >= sDate And today <= eDate Then
                    'set the filename variable to the value in the filename field
                    fName = .Fields("FileName").Value
                End If
                
                'move to the next record
                .MoveNext
            Loop
        End With
    End With

    'close the recordset
    recSet.Close
    
    'return the filename
    GetFileName = fName
End Function

Public Function StatusFunction()
    Dim accForm As Form
    Set accForm = Forms.Item("InputForm")
    
    'accForm.Command46_Click
    accForm.cmdSave_Click
End Function

'string matcher function to find if 2 strings are similar for intput form
Public Function StringMatcher(typedReason As String, tableReason As String) As Boolean
    Dim len1 As Integer, len2 As Integer, min As Integer, x As Integer, count As Integer
    Dim A() As Byte, B() As Byte
    
    'start count at 0
    count = 0
    
    'set the len variables to the 2 strings
    len1 = Len(typedReason)
    len2 = Len(tableReason)
    
    'change strings to byte arrays
    A = typedReason
    B = tableReason
    
    'set the min variable to the smaller of the 2 lengths
    min = MinNum(len1, len2)
    
    'loop through each character of each string and compare the 2
    For x = 0 To (min * 2) - 1 Step 2
        'check if characters match
        If A(x) = B(x) Then
            'if match, add 1 to count for number of matched characters
            count = count + 1
        End If
    Next x
    
    'check if count of characters between 2 strings is equal to len of smallest string minus 1
    If count = min - 1 Then
        'if count is close to original string set stringmatcher to true
        StringMatcher = True
    Else
        'else set to false
        StringMatcher = False
    End If
End Function

'function to find smaller of 2 numbers
Public Function MinNum(x As Integer, y As Integer) As Integer
    If x < y Then
      MinNum = x
    Else
      MinNum = y
    End If
End Function

'get assignments from assignments text file
Public Function GetAssignments()
    Dim DataLine As String, textline As String
    Dim assignments() As String, parts() As String, sql As String
    Dim x As Integer, count As Integer, oID As Integer, Avail As Integer, y As Integer, counter As Integer
    
    'count the number of assignments
    count = CountAssignments()
    counter = -1
    
    'set array size to number of assignments
    ReDim assignments(count, 6)
    
    'open the assignments text file
    Open "\\cenfile1\police$\Dispatching disk\databases\assignments.txt" For Input As #1
        
    'for every line in the assignments file
    Do While Not EOF(1)
        'grab one line at a time
        Line Input #1, textline
        counter = counter + 1
        'split the line by --
        parts = Split(textline, "--")
        
        'put each part of the assignment line into the array
        For y = 0 To UBound(parts)
            assignments(counter, y) = Nz(parts(y), "")
        Next y
    Loop
    
    'close the assignments text file
    Close #1

    DoCmd.SetWarnings False
        
    'loop through each assignment in the array
    For x = 0 To count - 1
        'if the officer id of the assignment matches an officer id in the parksids table
        If DCount("[id]", "parksIDs", "[pId] = " & CInt(assignments(x, 0))) > 0 Then
            'set the variable for whether the officer should be available or unavailable
            Select Case Trim(assignments(x, 5))
                Case "Sick Leave", "Annual Leave", "Other Leave", "Administration", "Desk", "Training", "Academy"
                    Avail = 0
                Case Else
                    Avail = 1
            End Select
            
            'sql to update the parksIDs table with the officer assignment information
            sql = "UPDATE parksIDs SET Isv=1, Avail=" & Avail & ", AssId=" & CLng(assignments(x, 1)) & _
               ", SecondaryId=" & CLng(assignments(x, 2)) & ", CarNo='" & assignments(x, 4) & "', StartTime='" & _
                assignments(x, 3) & "' WHERE pId = " & CInt(assignments(x, 0))
            
            DoCmd.RunSQL sql
        End If
    Next x
    
    DoCmd.SetWarnings True
    
    'check if the cad screen is loaded, if not open it
    If Not IsLoaded("CAD") Then
        DoCmd.OpenForm "CAD"
    End If
    
    'refresh the inservice officers section
    Forms!CAD!Child2.Form.Refresh
    'hide the updating message
    Forms!CAD!Label75.Visible = False
End Function

'counts the number of lines in the assignments text file
Public Function CountAssignments() As Integer
    Const BUFSIZE As Long = 100000
    Dim T0 As Single
    Dim LfAnsi As String
    Dim F As Integer
    Dim FileBytes As Long
    Dim BytesLeft As Long
    Dim Buffer() As Byte
    Dim strBuffer As String
    Dim BufPos As Long
    Dim LineCount As Long

    T0 = Timer()
    LfAnsi = StrConv(vbLf, vbFromUnicode)
    F = FreeFile(0)
    Open "\\cenfile1\police$\Dispatching disk\databases\assignments.txt" For Binary Access Read As #F
    FileBytes = LOF(F)
    ReDim Buffer(BUFSIZE - 1)
    BytesLeft = FileBytes
    Do Until BytesLeft = 0
        If BufPos = 0 Then
            If BytesLeft < BUFSIZE Then ReDim Buffer(BytesLeft - 1)
            Get #F, , Buffer
            strBuffer = Buffer 'Binary copy of bytes.
            BytesLeft = BytesLeft - LenB(strBuffer)
            BufPos = 1
        End If
        Do Until BufPos = 0
            BufPos = InStrB(BufPos, strBuffer, LfAnsi)
            If BufPos > 0 Then
                LineCount = LineCount + 1
                BufPos = BufPos + 1
            End If
        Loop
    Loop
    Close #F
    
    CountAssignments = LineCount
End Function

Public Sub UpdateAssignments()
    Dim sql As String
    
    DoCmd.SetWarnings False
    'generate sql to update the parksIds to delete all the assignment ids and cars
    sql = "UPDATE parksIDs SET Isv = 0, Avail = 0, AssId = 0, SecondaryId = 0, CarNo = '', StartTime = ''"
    'execute the sql
    DoCmd.RunSQL sql

    'check if the CAD form is open
    If IsLoaded("CAD") Then
        'change the label on the cad screen to show updating
        Forms!CAD!Label75.Visible = True
    End If
    
    'call the program that connects to the PARKS sql server and grabs the assignments
    Call Shell("\\cenfile1\police$\Dispatching disk\databases\pBlotCon.exe")
    'give the program time to run
    Pause 1
    'run the get assignments function to grab the assignments
    GetAssignments
End Sub

