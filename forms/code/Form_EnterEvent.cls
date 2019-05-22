Attribute VB_Name = "Form_EnterEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public outOfState As Boolean

'save event button
Private Sub Command6_Click()
    Dim EvtDate As String, evtTime As String, unit As String, Disp As String, Narrative As String, strSql As String
    Dim callId As Long, eId As Integer
    
    'check to see if date, time, officer, location and event type are all entered, if not advise user and exit
    If IsNull(Me.Text0.Value) Or IsNull(Me.Text2.Value) Or IsNull(Me.Text8.Value) Or IsNull(Me.Text10.Value) Or IsNull(Me.Text19.Value) Then
        MsgBox "Please make sure that the date, time, officer, location, and event type are all filled in."
        Exit Sub
    End If

    'lookup the id of the dispatcher generating the call
    Disp = GetDispatchId() & ""
    'set the unit from the event entry form
    unit = Me.Text8.Value
    'set the date from the event entry form
    EvtDate = DateLong(CDate(Me.Text0.Value))
    'set the time from the event entry form
    evtTime = ShortTime(CDate(Me.Text2.Value))
    'set the narrative from the event entry form and set it to "" if it is blank
    Narrative = Nz(Me.Text17.Value, "")
    
    'set ok to save record variable to true to allow record to be saved
    OkToSaveRecord
    'run the save record command
    DoCmd.RunCommand acCmdSaveRecord
    'disable the sql injection warnings
    DoCmd.SetWarnings False
    
    'generate the insert sql for the event call to be entered in the comm log
    strSql = "INSERT INTO commlog (Date1, Time1, SourceCall, Reason, Narrative, Dispatcher) VALUES ('" & EvtDate & "', '" & evtTime & "', '" & unit & "', 'EVENT', '" & Narrative & "', '" & Disp & "')"
    'execute the insert sql
    CurrentDb.Execute strSql
    
    'look up the id of the last call to be used for event / call association
    callId = CurrentDb.OpenRecordset("SELECT @@Identity")(0)
    'set the id
    eId = Me.idField
    
    'generate the insert sql for the event / call association
    strSql = "INSERT INTO EventCallAssoc (EventID, CallID) VALUES (" & eId & ", " & callId & ")"
    'execute the insert sql for call and event association
    DoCmd.RunSQL strSql
    
    'advise the user via message box that the record has been saved successfully
    MsgBox "Event has been saved!"
    'set the sql injection warnings to true
    DoCmd.SetWarnings True
    'refresh the call log
    Forms!CAD!Child4.Form.Requery
    'refresh the active events
    Forms!CAD!Child0.Form!ActiveEventHolder.Form.Requery
    'close the event entry form
    DoCmd.Close acForm, "EnterEvent", acSaveNo
End Sub

'procedure to check if the save record command has been initiated, this allows the record to be cancelled and not be saved
Private Sub Form_BeforeInsert(Cancel As Integer)
    If Command = acCmdSaveRecord Then
        'check the ready to save record variable has been set
        If ReadyToSaveRecord = False Then
            'if the variable hasn't been set in the save button procedure, then cancel the save procedure
            Cancel = True
        Else
            'else save the record by following through with the save procedure
            Cancel = False
        End If
    End If
End Sub

Private Sub Form_Close()
    'on closing the form, undo any unsaved changes
    Me.Undo
    'set the sql injection warnings on
    DoCmd.SetWarnings True
End Sub

Private Sub Form_Load()
    'set a new record to be created, sets ok to save to false
    CreateNewRecord
    'set the sql injection warnings off
    DoCmd.SetWarnings False
    'set the cancel button to have focus
    Me.Command4.SetFocus
    'set the out of state variable to false
    outOfState = False
End Sub

'cancel button
Private Sub Command4_Click()
    'undo any unsaved changes
    Me.Undo
    'check to see if the event was saved
    If Not Me.idField.Value = "(New)" Then
        'run sql to delete the newly created record, since we are cancelling the creation of the event
        DoCmd.RunSQL "DELETE FROM events WHERE ID = " & Me.idField.Value
    End If
    'close the event entry form
    DoCmd.Close acForm, "EnterEvent", acSaveNo
End Sub

'open the muliple states dialog for running data
Private Sub Label43_Click()
    DoCmd.OpenForm "MultState", acNormal, , , , , "EnterEvent"
End Sub

Private Sub StatePicker_AfterUpdate()
    Me.StateSource2.Value = Me.StatePicker.Value
End Sub

'reset the state source drop down box
Private Sub StateSource2_Click()
    'set the state picker drop down box to have the focus
    Me.StatePicker.SetFocus
    'set the value of the state source input to be nothing
    Me.StateSource2.Value = ""
    'hide the state source input
    Me.StateSource2.Visible = False
    'set the initial value of the state picker back to NY
    Me.StatePicker.Value = "NY"
End Sub

'search the tables to pull up an address if a park is selected
Private Sub Text10_Exit(Cancel As Integer)
    Dim pType As String, pZone As String, tableName As String
    'check to make sure there is a location entered
    If Not IsNull(Me.Text10.Value) Then
        'look up the type of location that is currently entered, set to "N" if the location type cannot be found
        pType = Nz(DLookup("[type]", "AllLocations", "[Park] = '" & Me.Text10.Value & "'"), "N")
        'check to make sure the location type is "P" as this is only to be done when a park is selected
        If pType = "P" Then
            'look up the zone the park is located in, set to "N" if zone cannot be found
            pZone = Left(Nz(DLookup("[Region]", "AllLocations", "[Park] = '" & Me.Text10.Value & "'"), "N"), 1)
            'set a table name based on the zone returned above
            Select Case pZone
                Case "F"
                    'if finger lakes zone, set to finger lakes table
                    tableName = "PANIC_BOOK_-_FL"
                Case "C"
                    'if central zone, set to central table
                    tableName = "PANIC_BOOK_-_CEN"
                Case "T"
                    'if thousand islands zone, set to thousand islands table
                    tableName = "PANIC_BOOK_-_TI"
                Case Else
                    'if no zone, exit procedure
                    Exit Sub
            End Select
            'set the address field of the enter event form by looking up park address in table, set to empty string if park cannot be found
            Me.Combo12.Value = Nz(DLookup("[Copy_of_Address]", tableName, "[Park] = '" & Me.Text10.Value & "'"), "")
            'check if there are any line breaks in the address
            If InStr(1, Me.Combo12.Value, vbCrLf) > 0 Then
                'if a line break is found, replace it with a comma and a space
                Me.Combo12.Value = Replace(Me.Combo12.Value, vbCrLf, ", ")
            End If
        End If
    End If
End Sub

'procedure that initiates the create of the mini map upon exiting the cross street input box
Private Sub Text14_Exit(Cancel As Integer)
    Dim webAddress As String, Address As String, XStreet As String, loc As String, totAddy As String, locType As String
    
    If IsNull(Me.Text10.Value) = False Then 'check to make sure a park / municipality is selected
        'set the location variable to the selection of the park / municipality drop down box
        loc = Me.Text10.Value
        'look up the type of location that has been entered, if location doesn't exist, set the type as "N"
        locType = Nz(DLookup("[type]", "AllLocations", "[Park] = '" & loc & "'"), "N")
        If locType = "P" Then
            'replace misc words from the location
            loc = sanitize(loc)
            'replace any spaces with the +, required for compatibility with the google maps api
            loc = Replace(loc, " ", "+")
        ElseIf locType = "C" Or locType = "T" Or locType = "V" Then
            loc = Right(loc, Len(loc) - 2)
            'replace any spaces with the +, required for compatibility with the google maps api
            loc = Replace(loc, " ", "+")
            'add on the ny tag for municipalities to let the google maps api that the address is in ny
            loc = loc & ",+NY"
        Else
            'replace any spaces with the +, required for compatibility with the google maps api
            loc = Replace(loc, " ", "+")
        End If
    Else
        'if there is no selection made in the park / municipality drop down, then set the location variable to nothing
        loc = ""
    End If
    
    If IsNull(Me.Combo12.Value) = False Then 'check to make sure an address is entered
        'set the address variable to that of the address box on the event entry page
        Address = Me.Combo12.Value
        'replace any spaces with the +, required for compatibility with the google maps api
        Address = Replace(Address, " ", "+")
    Else
        'if there is no address entered on the event entry page, set the address variable to nothing
        Address = ""
    End If
    
    If IsNull(Me.Text14.Value) = False Then 'check to make sure a cross street is entered
        'set the cross street variable to the value in the event entry form
        XStreet = Me.Text14.Value
        'once again replace all spaces with the + for the google maps api
        XStreet = Replace(XStreet, " ", "+")
    Else
        'if no cross street is entered, set the variable to nothing
        XStreet = ""
    End If
    
    If loc = "" And Address = "" And XStreet = "" Then
        'if there is no location, address, or cross street, then exit the procedure without performing any calls to the google maps api
        Exit Sub
    Else
        'set the total address variable to be used in the call to the google maps api, comprised of the address, cross street, and location
        totAddy = Address & "+and+" & XStreet & ",+" & loc
        'set the webaddress for the map browser to navigate to
        webAddress = "http://maps.googleapis.com/maps/api/staticmap?center=" & totAddy & "&zoom=14&size=400x200&sensor=false&visual_refresh=true&maptype=roadmap&markers=icon:http://goo.gl/6fmKFh%7Ccolor:red%7Clabel:X%7C" & totAddy
         
        'navigate to the google maps api web address generated above
        Me.WebBrowser6.Navigate webAddress
        'set the map location field of the event entry form to the web address of the map to be stored in the events table
        Me.MapLoc.Value = webAddress
        'set the total address to a field on the event entry form, necessary for use with the interactive map
        Me.Text52.Value = totAddy
    End If
End Sub

'run data by cid button
Private Sub CIDCheck_Click()
    Dim CID As String, State As String, sqlStr As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    'disable sql injection warnings
    DoCmd.SetWarnings False
    
    'check to see if a cid has been entered
    If IsNull(Me.CID.Value) = True Or Me.CID.Value = "" Then
        'if no cid has been entered, advise the user by message box and exit the procedure
        MsgBox "Please enter the CID number to continue."
        Exit Sub
    Else
        'set the cid
        CID = Me.CID.Value
    End If
    
    'check to see if the state has been entered
    If IsNull(Me.StatePicker.Value) = True Or Me.StatePicker.Value = "" Then
        'if the state has not been entered, advise the user by message box and exit the procedure
        MsgBox "Please enter the State to continue."
        Exit Sub
    Else
        'set the state
        State = Me.StatePicker.Value
    End If
    
    'generate the insert sql to save the cid in the person table
    sqlStr = "INSERT INTO people (State, CID) VALUES ('" & State & "', '" & CID & "')"
    'execute the insert sql to save the cid
    DoCmd.RunSQL sqlStr
    
    'count the number of people in the people table with that cid
    If DCount("[ID]", "people", "[CID] = '" & CID & "'") > 1 Then
        'if there are more than one, get the id from the last
        Set db = CurrentDb
        Set rs = db.OpenRecordset("SELECT * FROM people WHERE CID = '" & CID & "' ORDER BY ID ASC")
        
        rs.MoveLast
        pId = rs.Fields("ID")
        
        Set db = Nothing
        Set rs = Nothing
    Else
        'set the id to the id of the peron with the cid
        pId = DLookup("[ID]", "people", "[CID] = '" & CID & "'")
    End If
    
    'set the event id
    eId = Me.idField.Value
    
    'generate the insert sql to associate the person and the event
    sqlStr = "INSERT INTO EventPersonAssoc (EventId, PersonId) VALUES (" & eId & ", " & pId & ")"
    'execute the insert sql to associate the person and the event
    DoCmd.RunSQL sqlStr
    
    'enter the data call into the comm log
    EnterDataCall (CID)
    'set the sql injection warnings to true
    DoCmd.SetWarnings True
    'open the ejustice page for running a data by cid
    OpenEJustice "CID"
End Sub

'run data by date of birth button
Private Sub DOBCheck_Click()
    Dim DOB As String, State As String, fName As String, lName As String, mName As String, Sex As String, sqlStr As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    'disable sql injection warnings
    DoCmd.SetWarnings False
    
    'check to see if date of birth has been entered
    If IsNull(Me.DOB.Value) = True Or Me.DOB.Value = "" Then
        'if there is no date of birth, advise user and exit
        MsgBox "Please enter the DOB to continue."
        Exit Sub
    Else
        'set date of birth value from event entry form
        DOB = Me.DOB.Value
    End If
    
    'check to see if the state has been entered
    If IsNull(Me.StatePicker.Value) = True Or Me.StatePicker.Value = "" Then
        'if the state has not been entered, advise user and exit
        MsgBox "Please enter the State to continue."
        Exit Sub
    Else
        'set the state value from the event entry form
        State = Me.StatePicker.Value
    End If
    
    'check to see if the person's last name has been entered
    If IsNull(Me.LastName.Value) = True Or Me.LastName.Value = "" Then
        'if no last name entered, advise user and exit
        MsgBox "Please enter the person's last name to continue."
        Exit Sub
    Else
        'set the last name from the form
        lName = Me.LastName.Value
    End If
    
    'check to see if the person's first name has been entered
    If IsNull(Me.FirstName.Value) = True Or Me.FirstName.Value = "" Then
        'if no last name entered, advise and exit
        MsgBox "Please enter the person's first name to continue."
        Exit Sub
    Else
        'set the first name from the form
        fName = Me.FirstName.Value
    End If
    
    'check to see if the person's gender has been entered
    If IsNull(Me.Sex.Value) = True Or Me.Sex.Value = "" Then
        'if the gender hasn't been entered, advise and exit
        MsgBox "Please enter the person's gender to continue."
        Exit Sub
    Else
        'set the gender
        Sex = Me.Sex.Value
    End If
        
    'check to see if the middle initial has been entered
    If IsNull(Me.Middle.Value) = True Then
        'if no middle initial is entered, set it to empty string, no need to advise user or exit since middle initial is not a required field on the ejustice form
        mName = ""
    Else
        'set the middle initial
        mName = Me.Middle.Value
    End If
    
    'generate the insert sql to insert the person into the people table
    sqlStr = "INSERT INTO people (State, DOB, LastName, FirstName, Middle, Sex) VALUES ('" & State & "', '" & DOB & "', '" & lName & "', '" & fName & "', '" & Middle & "', '" & Sex & "')"
    'execute the insert sql
    DoCmd.RunSQL sqlStr
    
    'count the number of people with the same name and date of birth
    If DCount("[ID]", "people", "[DOB] = '" & DOB & "' AND [FirstName] = '" & fName & "' AND [LastName] = '" & lName & "'") > 1 Then
        'if there are more than one entry for the same name and date of birth, go to the last record and get the id
        Set db = CurrentDb
        Set rs = db.OpenRecordset("SELECT * FROM people WHERE CID = '" & CID & "' ORDER BY ID ASC")
        
        rs.MoveLast
        pId = rs.Fields("ID")
        
        Set db = Nothing
        Set rs = Nothing
    Else
        'set the person id for the previously entered person
        pId = DLookup("[ID]", "people", "[DOB] = '" & DOB & "' AND [FirstName] = '" & fName & "' AND [LastName] = '" & lName & "'")
    End If
    
    'set the event id
    eId = Me.OpenArgs
    
    'generate the insert sql to associate the person with the event
    sqlStr = "INSERT INTO EventPersonAssoc (EventId, PersonId) VALUES (" & eId & ", " & pId & ")"
    'execute the insert sql for association
    DoCmd.RunSQL sqlStr
    
    'enter call in the comm log for request for data
    EnterDataCall ("(" & State & ") " & lName & ", " & fName & " " & Middle & ", " & DOB)
    'reenable the sql injection warnings
    DoCmd.SetWarnings True
    'open the ejustice form and run the data with the name and date of birth provided
    OpenEJustice "DOB"
End Sub

'registration check by plate button
Private Sub PlateCheck_Click()
    Dim Plate As String, State As String, Exp As String, sqlStr As String
    Dim plateType As Integer, pId As Integer, eId As Integer
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    'disable the sql injection warnings
    DoCmd.SetWarnings False
    
    'check if the plate number has been entered
    If IsNull(Me.PlateNum.Value) = True Or Me.PlateNum.Value = "" Then
        'if the plate number has not been entered, advise the user and exit
        MsgBox "Please enter the plate number to continue."
        Exit Sub
    Else
        'set the plate number to be run
        Plate = Me.PlateNum.Value
    End If
    
    'check to see if the state has been entered
    If IsNull(Me.Combo35.Value) = True Or Me.Combo35.Value = "" Then
        'if no state has been entered, advise and exit
        MsgBox "Please enter the State to continue."
        Exit Sub
    Else
        'set the state for the plate to be run in
        State = Me.Combo35.Value
    End If
    
    'check to see if the plate type has been entered
    If IsNull(Me.Type.Value) = True Or Me.Type.Value = "" Then
        'if the plate type has not been entered, then advise and exit only if state value is not NY
        If Not Me.Combo35.Value = "NY" Then
            MsgBox "Please enter the plate type to continue."
            Exit Sub
        End If
    Else
        'set the plate type to be checked
        plateType = Me.Type.Value
    End If
    
    'check whether the expiration date has been filled in
    If IsNull(Me.Expiration.Value) = True Or Me.Expiration.Value = "" Then
        'if the expiration year has not been filled in, advise and exit only if state value is not NY
        If Not Me.Combo35.Value = "NY" Then
            MsgBox "Please enter the expiration year to continue."
            Exit Sub
        End If
    Else
        'set the expiration
        Exp = Me.Expiration.Value
    End If
    
    'generate the insert sql to insert the vehicle informat into the vehicles table
    sqlStr = "INSERT INTO vehicles (State, PlateNumber, Type, Expiration) VALUES ('" & State & "', '" & Plate & "', " & plateType & ", '" & Exp & "')"
    'execute the insert sql
    DoCmd.RunSQL sqlStr
    
    'count the number of times the plate number is in the vehicles table
    If DCount("[ID]", "vehicles", "[PlateNumber] = '" & Plate & "'") > 1 Then
        'if the plate is in the table more than once then move to the last record and get the id
        Set db = CurrentDb
        Set rs = db.OpenRecordset("SELECT * FROM vehicles WHERE PlateNumber = '" & Plate & "' ORDER BY ID ASC")
        
        rs.MoveLast
        pId = rs.Fields("ID")
        
        Set db = Nothing
        Set rs = Nothing
    Else
        'set the plate id
        pId = DLookup("[ID]", "vehicles", "[PlateNumber] = '" & Plate & "'")
    End If
    
    'set the event id
    eId = Me.idField.Value
    
    'generate the insert sql for vehicle / event association
    sqlStr = "INSERT INTO EventVehicleAssoc (EventId, VehId) VALUES (" & eId & ", " & pId & ")"
    'execute the insert sql for association
    DoCmd.RunSQL sqlStr
    
    'enter the call for data in the comm log
    EnterDataCall ("(" & State & ") " & Plate)
    'set the sql injection warnings back on
    DoCmd.SetWarnings True
    
    'if out of state, change plate type from int value to string value
    If outOfState Then
        sId = Me.Type.Value
        Me.Type.BoundColumn = 4
        Me.Type.Value = DLookup("[ShortCode]", "plateTypes", "[TypeNo] = " & sId)
    End If
    'open the ejustice form to run the plate data
    OpenEJustice "REG"
End Sub

'on exit state drop down
Private Sub Combo35_Exit(Cancel As Integer)
    'check state value
    If Me.Combo35.Value = "NY" Then
        'if ny, lock and disable the vehicle year and make fields as they are not required by ejustice for ny plates
        Me.VehMake.Enabled = False
        Me.VehMake.Locked = True
        Me.VehYear.Enabled = False
        Me.VehYear.Locked = True
        Me.Type.BoundColumn = 3
    Else
        'if not ny, unlock and enable the vehicle year and make fields as they are required to run an out of state plate
        Me.VehMake.Enabled = True
        Me.VehMake.Locked = False
        Me.VehYear.Enabled = True
        Me.VehYear.Locked = False
        Me.Type.BoundColumn = 3
        outOfState = True
    End If
End Sub

'upon exiting the officer field
Private Sub Text8_Exit(Cancel As Integer)
    'set the form field text46 to the parks id of the officer that is selected
    Me.Text46.Value = DLookup("[pId]", "InServiceOfficers", "[Expr1] = '" & Me.Text8.Value & "'")
    'check to see if the parks id is successfully loaded into text 46
    If IsNull(Me.Text46.Value) = False And Not Me.Text46.Value = "" Then
        Dim zoneId As Integer
        'look up the zone id
        zoneId = DLookup("[pZone]", "parksIDs", "[pId] = " & Me.Text46.Value)
        'set a form field with the zone name
        Me.Text47.Value = DLookup("[zone]", "zones", "[pId] = " & zoneId)
        'refresh the park / municipality drop down to only show places in the selected officer's zone
        Me.Text10.Requery
    End If
End Sub

'registration check by vin button
Private Sub VINCheck_Click()
    Dim VIN As String, State As String, VehMake As String, sqlStr As String
    Dim VehYear As Integer, pId As Integer, eId As Integer
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    'disable the sql injection warnings
    DoCmd.SetWarnings False
    
    'check to see if the vin has been entered
    If IsNull(Me.VIN.Value) = True Or Me.VIN.Value = "" Then
        'if the vin has not been entered, advise the user and exit the procedure
        MsgBox "Please enter the VIN to continue."
        Exit Sub
    Else
        'set the vin to be checked
        VIN = Me.VIN.Value
    End If
    
    'check to see if the state has been entered
    If IsNull(Me.Combo35.Value) = True Or Me.Combo35.Value = "" Then
        'if the state has not been entered, advise and exit
        MsgBox "Please enter the State to continue."
        Exit Sub
    Else
        'set the state
        State = Me.Combo35.Value
    End If
    
    'set the vehicle make from the form, and an empty string if it isn't entered on the form
    VehMake = Nz(Me.VehMake.Value, "")
    'set the vehicle year from the form, and 0 if the year isn't entered on the form
    VehYear = Nz(Me.VehYear.Value, 0)
    
    'generate the insert sql to insert the vehicle info into the vehicles table
    sqlStr = "INSERT INTO vehicles (State, VIN, VehYear, VehMake) VALUES ('" & State & "', '" & VIN & "', " & VehYear & ", '" & VehMake & "')"
    'execute the sql
    DoCmd.RunSQL sqlStr
    
    'count the number of vehicles with the entered vin
    If DCount("[ID]", "vehicles", "[VIN] = '" & VIN & "'") > 1 Then
        'if more than one, go to last one entered and get the id
        Set db = CurrentDb
        Set rs = db.OpenRecordset("SELECT * FROM vehicles WHERE VIN = '" & VIN & "' ORDER BY ID ASC")
        
        rs.MoveLast
        pId = rs.Fields("ID")
        
        Set db = Nothing
        Set rs = Nothing
    Else
        'set the vehicle id
        pId = DLookup("[ID]", "vehicles", "[VIN] = '" & VIN & "'")
    End If
    
    'set the event id
    eId = Me.idField.Value
    
    'generate the insert sql to associate the vehicle and the event
    sqlStr = "INSERT INTO EventVehicleAssoc (EventId, VehId) VALUES (" & eId & ", " & pId & ")"
    'execute the insert sql
    DoCmd.RunSQL sqlStr
    
    'enter call in comm log requesting data be run
    EnterDataCall ("(" & State & ") " & VIN)
    'reenable sql injection warnings
    DoCmd.SetWarnings True
    'open ejustice to run data on vin
    OpenEJustice "VIN"
End Sub

'enter calls into the comm log when a data is run
Public Function EnterDataCall(Narr As String)
    Dim Officer As String, CurDate As String, CurTime As String, sql As String, sql2 As String, sql3 As String
    Dim x As Long
    
    'generate sql for data request in comm log
    sql = "INSERT INTO commlog (Date1, Time1, SourceCall, Reason, Narrative, Dispatcher) VALUES ('" & DateLong(CDate(Me.Text0.Value)) & "', '" & ShortTime(CDate(Me.Text2.Value)) & "', '" & Me.Text8.Value & "', 'DATA', '@" & Me.Text10.Value & " @" & Me.Combo12.Value & " and " & Me.Text14.Value & ". " & Narr & "', " & GetDispatchId() & ")"
    'generate sql to get id of last call entered
    sql2 = "SELECT @@Identity"
    
    'execute sql to insert call into comm log
    CurrentDb.Execute sql
    'get id of last call entered
    x = CurrentDb.OpenRecordset(sql2)(0)
    
    'generate sql for event call association
    sql3 = "INSERT INTO EventCallAssoc (EventID, CallID) VALUES (" & Me.id.Value & ", " & x & ")"
    'execute sql for association
    CurrentDb.Execute sql3
End Function

'open interactive map when mini map is brought into focus, requires user to click on map
Private Sub WebBrowser6_Enter()
    'MsgBox Me.Text52.Value
    a_testGeocode Me.Text52.Value
End Sub
