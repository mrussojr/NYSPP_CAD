Attribute VB_Name = "Form_RunData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Location As String
Private caller As String
Private callId As Long
Public outOfState As Boolean

Private Sub Form_Load()
    'set a variable indicating this is a new data being run
    CreateNewRecord
    'set sql injection warnings false
    DoCmd.SetWarnings False
    'set the last name to have the focus
    Me.LastName.SetFocus
    'set out of state variable to false
    outOfState = False
    'check the opening arguments
    If IsNull(Me.OpenArgs) = False And Not Me.OpenArgs = "" Then
        'if opening arguments exist, call setLocation to set the location of the data being run
        SetLocation DLookup("[Location]", "events", "[ID] = " & Me.OpenArgs)
        'call setCaller to set the caller of the data being run
        SetCaller DLookup("[Off]", "events", "[ID] = " & Me.OpenArgs)
        'set the caller from the caller variable
        Me.Text8.Value = GetCaller()
        'set the location from the location variable
        Me.Text49.Value = GetLocation()
    Else
        'unlock the caller and location fields so the user can enter an officer and location
        Me.Text8.Enabled = True
        Me.Text8.Locked = False
        Me.Text49.Enabled = True
        Me.Text49.Locked = False
    End If
End Sub

'cancel button
Private Sub Command4_Click()
    'undo any changes
    Me.Undo
    'close the run data form
    DoCmd.Close acForm, "RunData", acSaveNo
End Sub

'run data by cid button
Private Sub CIDCheck_Click()
    Dim CID As String, State As String, sqlStr As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    'check to make sure an officer is entered in the officer field, advise user if not entered and exit the procedure without running the data
    If IsNull(Me.Text8.Value) = True Or Me.Text8.Value = "" Then
        MsgBox "Please enter an officer for the data to be run for."
        Exit Sub
    End If
    
    'check to make sure a location in the location field, advise user if not entered and exit the procedure without running the data
    If IsNull(Me.Text49.Value) = True Or Me.Text49.Value = "" Then
        MsgBox "Please enter a location where the officer is running the data."
        Exit Sub
    End If
    
    'disable the sql injection warnings
    DoCmd.SetWarnings False
    
    'check if the cid has been entered, advise if not and exit without running the data, else set the cid variable to the cid
    If IsNull(Me.CID.Value) = True Or Me.CID.Value = "" Then
        MsgBox "Please enter the CID number to continue."
        Exit Sub
    Else
        CID = Me.CID.Value
    End If
    
    'check if the state has been entered, advise if not and exit without running the data, else set the state variable to the state
    If IsNull(Me.StatePicker.Value) = True Or Me.StatePicker.Value = "" Then
        MsgBox "Please enter the State to continue."
        Exit Sub
    Else
        State = Me.StatePicker.Value
    End If
    
    'generate the insert sql to save the cid in the people table
    sqlStr = "INSERT INTO people (State, CID) VALUES ('" & State & "', '" & CID & "')"
    'execute the insert sql
    DoCmd.RunSQL sqlStr
    
    'check the opening arguments of the form
    If IsNull(Me.OpenArgs) = False And Not Me.OpenArgs = "" Then
        'count the number of people with the cid
        If DCount("[ID]", "people", "[CID] = '" & CID & "'") > 1 Then
            'if there are more than one, get the id from the most recent
            Set db = CurrentDb
            Set rs = db.OpenRecordset("SELECT * FROM people WHERE CID = '" & CID & "' ORDER BY ID ASC")
            
            rs.MoveLast
            pId = rs.Fields("ID")
            
            Set db = Nothing
            Set rs = Nothing
        Else
            'set the id of the cid just entered
            pId = DLookup("[ID]", "people", "[CID] = '" & CID & "'")
        End If
        
        'set the event id to the opening arguments of the form
        eId = Me.OpenArgs
        
        'generate the insert sql to associate the event and the person
        sqlStr = "INSERT INTO EventPersonAssoc (EventId, PersonId) VALUES (" & eId & ", " & pId & ")"
        'execute the sql for association
        DoCmd.RunSQL sqlStr
    End If
    
    'enter call in the call log requesting the data
    EnterDataCall (CID)
    'set the sql injection warnings to on
    DoCmd.SetWarnings True
    'open the ejustice form to run the cid
    OpenEJustice "RCID"
    'DoCmd.Close acForm, "RunData", acSaveNo
End Sub

'run data by date of birth button
Private Sub DOBCheck_Click()
    Dim DOB As String, State As String, fName As String, lName As String, mName As String, Sex As String, sqlStr As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    'check to make sure an officer is entered in the officer field, advise user if not entered and exit the procedure without running the data
    If IsNull(Me.Text8.Value) = True Or Me.Text8.Value = "" Then
        MsgBox "Please enter an officer for the data to be run for."
        Exit Sub
    End If
    
    'check to make sure a location is entered in the location field, advise user if not entered and exit the procedure without running the data
    If IsNull(Me.Text49.Value) = True Or Me.Text49.Value = "" Then
        MsgBox "Please enter a location where the officer is running the data."
        Exit Sub
    End If
    
    'disable the sql injection warnings
    DoCmd.SetWarnings False
    
    'check to make sure a date of birth is entered, advise if not and exit without running the data, else set the date of birth variable
    If IsNull(Me.DOB.Value) = True Or Me.DOB.Value = "" Then
        MsgBox "Please enter the DOB to continue."
        Exit Sub
    Else
        DOB = Me.DOB.Value
    End If
    
    'check to make sure a state has been entered, advise if no state has been entered and exit without running the data, else set the state variable
    If IsNull(Me.StatePicker.Value) = True Or Me.StatePicker.Value = "" Then
        MsgBox "Please enter the State to continue."
        Exit Sub
    Else
        State = Me.StatePicker.Value
    End If
    
    'check to make sure a last name has been entered, advise if no last name has been entered and exit without running the data, else set the last name variable
    If IsNull(Me.LastName.Value) = True Or Me.LastName.Value = "" Then
        MsgBox "Please enter the person's last name to continue."
        Exit Sub
    Else
        lName = Me.LastName.Value
    End If
    
    'check to make sure a first name has been entered, advise if no first name has been entered and exit without running the data, else set the first name variable
    If IsNull(Me.FirstName.Value) = True Or Me.FirstName.Value = "" Then
        MsgBox "Please enter the person's first name to continue."
        Exit Sub
    Else
        fName = Me.FirstName.Value
    End If
    
    'check to make sure the gender has been entered, advise if no gender selected and exit without running the data, else set the gender variable
    If IsNull(Me.Sex.Value) = True Or Me.Sex.Value = "" Then
        MsgBox "Please enter the person's gender to continue."
        Exit Sub
    Else
        Sex = Me.Sex.Value
    End If
        
    'check to make sure the middle initial has been entered, if not set to empty string, else set the variable
    If IsNull(Me.Middle.Value) = True Then
        mName = ""
    Else
        mName = Me.Middle.Value
    End If
    
    'generate the sql to insert the data into the people table
    sqlStr = "INSERT INTO people (State, DOB, LastName, FirstName, Middle, Sex) VALUES ('" & State & "', '" & DOB & "', '" & lName & "', '" & fName & "', '" & Middle & "', '" & Sex & "')"
    'execute the sql
    DoCmd.RunSQL sqlStr
    
    'check the open arguments of the form
    If IsNull(Me.OpenArgs) = False And Not Me.OpenArgs = "" Then
        'count the number of people in the people table, if more than one select the last one
        If DCount("[ID]", "people", "[DOB] = '" & DOB & "' AND [FirstName] = '" & fName & "' AND [LastName] = '" & lName & "'") > 1 Then
            Set db = CurrentDb
            Set rs = db.OpenRecordset("SELECT * FROM people WHERE CID = '" & CID & "' ORDER BY ID ASC")
            
            rs.MoveLast
            pId = rs.Fields("ID")
            
            Set db = Nothing
            Set rs = Nothing
        Else
            'set the id to the last person entered
            pId = DLookup("[ID]", "people", "[DOB] = '" & DOB & "' AND [FirstName] = '" & fName & "' AND [LastName] = '" & lName & "'")
        End If
        
        'set the event id from the opening arguments from the form
        eId = Me.OpenArgs
        
        'generate the sql to associate the person and the event
        sqlStr = "INSERT INTO EventPersonAssoc (EventId, PersonId) VALUES (" & eId & ", " & pId & ")"
        'execute the sql to associate the person and event
        DoCmd.RunSQL sqlStr
    End If

    'enter call into comm log requesting data
    EnterDataCall ("(" & State & ") " & lName & ", " & fName & " " & Middle & ", " & DOB)
    'set the sql injection warnings back on
    DoCmd.SetWarnings True
    'open the ejustice form to run the data by the name and date of birth
    OpenEJustice "RDOB"
    
    'open the parks search form
    Dim str As String
    str = "Persons/" & lName & "/" & fName
    DoCmd.OpenForm "SearchParks", , , , , , str
End Sub

'link to open the multiple state picker
Private Sub Label43_Click()
    DoCmd.OpenForm "MultState", acNormal, , , , , "RunData"
End Sub

'run data by plate button
Private Sub PlateCheck_Click()
    Dim Plate As String, State As String, Exp As String, sqlStr As String
    Dim plateType As Integer, pId As Integer, eId As Integer
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sId As Variant

    'check to make sure an officer is selected for the date to be run for, if none selected, advise user, and exit
    If IsNull(Me.Text8.Value) = True Or Me.Text8.Value = "" Then
        MsgBox "Please enter an officer for the data to be run for."
        Exit Sub
    End If
    
    'check to make sure a location is entered, if not, advise user and exit
    If IsNull(Me.Text49.Value) = True Or Me.Text49.Value = "" Then
        MsgBox "Please enter a location where the officer is running the data."
        Exit Sub
    End If

    'set the sql injection warnings off
    DoCmd.SetWarnings False

    'check to make sure a plate is entered, if not, advise user and exit, else set the plate number variable
    If IsNull(Me.PlateNum.Value) = True Or Me.PlateNum.Value = "" Then
        MsgBox "Please enter the plate number to continue."
        Exit Sub
    Else
        Plate = Me.PlateNum.Value
    End If

    'check to make sure a state is selected, if not, advise user and exit, else set the state variable
    If IsNull(Me.Combo35.Value) = True Or Me.Combo35.Value = "" Then
        MsgBox "Please enter the State to continue."
        Exit Sub
    Else
        State = Me.Combo35.Value
    End If
    
    'check to make sure the plate type is entered, if not, advise user and exit, else set the plate type variable
    If IsNull(Me.Type.Value) = True Or Me.Type.Value = "" Then
        MsgBox "Please enter the plate type to continue."
        Exit Sub
    Else
        plateType = Me.Type.Value
    End If

    'check to make sure the expiration is entered, if not, advise user and exit, else set the expiration variable
    If IsNull(Me.Expiration.Value) = True Or Me.Expiration.Value = "" Then
        MsgBox "Please enter the expiration year to continue."
        Exit Sub
    Else
        Exp = Me.Expiration.Value
    End If
    
    'generate the sql to insert the vehicle into the vehicle table
    sqlStr = "INSERT INTO vehicles (State, PlateNumber, Type, Expiration) VALUES ('" & State & "', '" & Plate & "', " & plateType & ", '" & Exp & "')"
    'execute the sql
    DoCmd.RunSQL sqlStr

    'check if the forms opening arguments exist
    If IsNull(Me.OpenArgs) = False And Not Me.OpenArgs = "" Then
        'count the number of vehicles with the same plate number and select the last one and set the vehicle id variable
        If DCount("[ID]", "vehicles", "[PlateNumber] = '" & Plate & "'") > 1 Then
            Set db = CurrentDb
            Set rs = db.OpenRecordset("SELECT * FROM vehicles WHERE PlateNumber = '" & Plate & "' ORDER BY ID ASC")

            rs.MoveLast
            pId = rs.Fields("ID")

            Set db = Nothing
            Set rs = Nothing
        Else
            'set the vehicle id
            pId = DLookup("[ID]", "vehicles", "[PlateNumber] = '" & Plate & "'")
        End If
    
        'set the event id variable from the forms opening arguments
        eId = Me.OpenArgs

        'generate the sql for vehicle and event association
        sqlStr = "INSERT INTO EventVehicleAssoc (EventId, VehId) VALUES (" & eId & ", " & pId & ")"
        'execute the sql
        DoCmd.RunSQL sqlStr
    End If

    'enter call into the comm log requesting the data
    EnterDataCall ("(" & State & ") " & Plate)
    'reenable the sql injection warnings
    DoCmd.SetWarnings True
    
    'if out of state, change plate type from int value to string value
    If outOfState Then
        sId = Me.Type.Value
        Me.Type.BoundColumn = 4
        Me.Type.Value = DLookup("[ShortCode]", "plateTypes", "[TypeNo] = " & sId)
    End If
    'open the ejustice form to run the plate data
    OpenEJustice "RREG"
    'DoCmd.Close acForm, "RunData", acSaveNo
    
    Dim str As String
    str = "LicensePlates/" & Plate
    DoCmd.OpenForm "SearchParks", , , , , , str
End Sub

'on exit state drop down
Private Sub Combo35_Exit(Cancel As Integer)
    'check the state value
    If Me.Combo35.Value = "NY" Then
        'if ny, lock the vehicle year and make fields as they are not required by ejustice
        Me.VehMake.Enabled = False
        Me.VehMake.Locked = True
        Me.VehYear.Enabled = False
        Me.VehYear.Locked = True
        Me.Type.BoundColumn = 3
    Else
        'else, unlock the vehicle year and make fields as they are required by ejustice to run out of state plates
        Me.VehMake.Enabled = True
        Me.VehMake.Locked = False
        Me.VehYear.Enabled = True
        Me.VehYear.Locked = False
        Me.Type.BoundColumn = 3
        outOfState = True
    End If
End Sub

Private Sub StatePicker_AfterUpdate()
    Me.StateSource2.Value = Me.StatePicker.Value
End Sub

'click the state source field
Private Sub StateSource2_Click()
    'set the focus to the state picker
    Me.StatePicker.SetFocus
    'hide the state source field
    Me.StateSource2.Visible = False
    'set the value of the state source field to an empty string
    Me.StateSource2.Value = ""
    'set the value of the state picker to ny
    Me.StatePicker.Value = "NY"
End Sub

'run data by vin button
Private Sub VINCheck_Click()
    Dim VIN As String, State As String, VehMake As String, sqlStr As String
    Dim VehYear As Integer, pId As Integer, eId As Integer
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    'check if the officer has been entered, if not, advise user and exit
    If IsNull(Me.Text8.Value) = True Or Me.Text8.Value = "" Then
        MsgBox "Please enter an officer for the data to be run for."
        Exit Sub
    End If
    
    'check if the location has been entered, if not, advise user and exit
    If IsNull(Me.Text49.Value) = True Or Me.Text49.Value = "" Then
        MsgBox "Please enter a location where the officer is running the data."
        Exit Sub
    End If
    
    'set the sql injection warnings off
    DoCmd.SetWarnings False
    
    'check if the vin has been entered, if not, advise user and exit, else set the vin variable
    If IsNull(Me.VIN.Value) = True Or Me.VIN.Value = "" Then
        MsgBox "Please enter the VIN to continue."
        Exit Sub
    Else
        VIN = Me.VIN.Value
    End If
    
    'check to see if the state has been entered, if not, advise user and exit, else set the state variable
    If IsNull(Me.Combo35.Value) = True Or Me.Combo35.Value = "" Then
        MsgBox "Please enter the State to continue."
        Exit Sub
    Else
        State = Me.Combo35.Value
    End If
    
    'set the vehicle make variable, if null set to empty string
    VehMake = Nz(Me.VehMake.Value, "")
    'set the vehicle year variable, if null set to 0
    VehYear = Nz(Me.VehYear.Value, 0)
    
    'generate sql to insert vehicle into the vehicle table
    sqlStr = "INSERT INTO vehicles (State, VIN, VehYear, VehMake) VALUES ('" & State & "', '" & VIN & "', " & VehYear & ", '" & VehMake & "')"
    'execute the sql
    DoCmd.RunSQL sqlStr
    
    'check the opening arguments of the form
    If IsNull(Me.OpenArgs) = False And Not Me.OpenArgs = "" Then
        'count the number of vehicles with the vin, and select the id from the last one entered
        If DCount("[ID]", "vehicles", "[VIN] = '" & VIN & "'") > 1 Then
            Set db = CurrentDb
            Set rs = db.OpenRecordset("SELECT * FROM vehicles WHERE VIN = '" & VIN & "' ORDER BY ID ASC")
            
            rs.MoveLast
            pId = rs.Fields("ID")
            
            Set db = Nothing
            Set rs = Nothing
        Else
            'set the id of the vehicle that matches the vin just entered
            pId = DLookup("[ID]", "vehicles", "[VIN] = '" & VIN & "'")
        End If
        
        'set the event id
        eId = Me.OpenArgs
        
        'generate the sql to associate the event and the vehicle
        sqlStr = "INSERT INTO EventVehicleAssoc (EventId, VehId) VALUES (" & eId & ", " & pId & ")"
        'execute the sql
        DoCmd.RunSQL sqlStr
    End If
    
    'enter the call for data in the comm log
    EnterDataCall ("(" & State & ") " & VIN)
    'reenable the sql warnings
    DoCmd.SetWarnings True
    'open the ejustice form to run the data by vin
    OpenEJustice "RVIN"
    'DoCmd.Close acForm, "RunData", acSaveNo
End Sub

'set the location variable
Public Function SetLocation(loc As String)
    Location = loc
End Function

'set the unit called variable
Public Function SetCaller(unit As String)
    caller = unit
End Function

'set the call id variable
Public Function SetCallId(x As Long)
    callId = x
End Function

'get the location from the location variable
Public Function GetLocation() As String
    GetLocation = Location
End Function

'get the caller from the caller variable
Public Function GetCaller() As String
    GetCaller = caller
End Function

'get the call id from the call id variable
Public Function GetCallId() As Long
    GetCallId = callId
End Function

'enter a call for data in the comm log
Public Function EnterDataCall(Narr As String)
    Dim Officer As String, CurDate As String, CurTime As String, sql As String, sql2 As String, sql3 As String
    Dim x As Long
    
    'generate the sql to insert the call into the comm log
    sql = "INSERT INTO commlog (Date1, Time1, SourceCall, Reason, Narrative, Dispatcher) VALUES ('" & DateLong(CDate(Me.Text0.Value)) & "', '" & ShortTime(CDate(Me.Text2.Value)) & "', '" & Me.Text8.Value & "', 'DATA', '@" & Me.Text49.Value & ". " & Narr & "', " & GetDispatchId() & ")"
    
    'execute the sql
    CurrentDb.Execute sql
    'generate the sql to get the id of the last call entered
    sql2 = "SELECT @@Identity"
    x = CurrentDb.OpenRecordset(sql2)(0)
    'set the call id variable to the call id of the last call entered
    SetCallId (x)
    
    'check if the forms opening arguments are entered or not
    If IsNull(Me.OpenArgs) = False And Not Me.OpenArgs = "" Then
        'generate the sql to insert the association between the call and the event
        sql3 = "INSERT INTO EventCallAssoc (EventID, CallID) VALUES (" & Me.OpenArgs & ", " & GetCallId() & ")"
        'execute the sql
        CurrentDb.Execute sql3
        'update the last time value of the event to be the time that the data was run
        UpdateLastTime Me.Text2.Value, Me.Text0.Value, Me.OpenArgs
    End If
    
    'refresh the call log
    Forms!CAD.Child4.Requery
End Function
