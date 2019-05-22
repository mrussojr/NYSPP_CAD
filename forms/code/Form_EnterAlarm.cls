Attribute VB_Name = "Form_EnterAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command48_Click()
    'Email button to send email to facility manager advising of alarm at said facility
    Dim SubStr As String, emailAddress As String, bodyStr As String, Park As String, emailSentStr As String
    
    'get the park the alarm was set off in
    Park = Me.Text10.Value
    'set the subject line of the email to park alarm
    SubStr = Park & " Alarm"
    
    'lookup the email address for the manager for the park that the alarm has gone off in
    emailAddress = DLookup("[Expr1]", "emails", "[Park] = '" & Park & "'")
    
    'set the body of the email
    bodyStr = "Alarm @ " & Me.Text10.Value & " @ " & Me.Combo12.Value & ".<br>" & Me.Text17.Value & "<br>" & emailAddress
    
    'Comment out the following line before going live
    'emailAddress = "michael.russo@parks.ny.gov"
    
    'call the procedure for sending the email and pass in variables for email address, body, and subject
    SendAlarmEmail emailAddress, bodyStr, SubStr
    
    'generate a string advising that the email has been sent to whom and at what time
    emailSentStr = "EMAIL SENT TO MANAGER (" & emailAddress & ") @ " & RightNow()
    
    If IsNull(Me.Text17.Value) = True Or Me.Text17.Value = "" Then
        'if there is no narrative then set the narrative to the email sent string
        Me.Text17.Value = emailSentStr
    Else
        'else set it to the current narrative and add the email sent string to the end
        Me.Text17.Value = Me.Text17.Value & vbCrLf & emailSentStr
    End If
End Sub

'save alarm button
Private Sub Command6_Click()
    Dim EvtDate As String, evtTime As String, unit As String, Disp As String, Narrative As String, strSql As String
    Dim callId As Long, eId As Long
    
    'lookup the id of the dispatcher taking the alarm call
    Disp = GetDispatchId() & ""
    'set the unit calling from the alarm entry form
    unit = Me.Text8.Value
    'set the date from the event entry form
    EvtDate = DateLong(CDate(Me.Text0.Value))
    'set the time from the event entry form
    evtTime = ShortTime(CDate(Me.Text2.Value))
    'set the narrative from the alarm entry form
    Narrative = Nz(Me.Text17.Value, "")
    
    'set the variable indicating that it is ok to save the record
    OkToSaveRecord
    'save the record
    DoCmd.RunCommand acCmdSaveRecord
    'set the sql injection warnings off
    DoCmd.SetWarnings False
    
    'generate the insert sql for inserting the call data into the call log
    strSql = "INSERT INTO commlog (Date1, Time1, SourceCall, Reason, Narrative, Dispatcher) VALUES ('" & EvtDate & "', '" & evtTime & "', '" & unit & "', 'ALARM', '" & Narrative & "', '" & Disp & "')"
    'execute the insert sql
    DoCmd.RunSQL strSql
    
    'look up the id of the call that was just entered for association with the alarm
    callId = DLookup("[id]", "commlog", "[Date1] = '" & EvtDate & "' AND [Time1] = '" & evtTime & "' AND [Dispatcher] = '" & Disp & "'")
    'set the id
    eId = Me.idField
    
    'generate the insert sql for call and alarm association
    strSql = "INSERT INTO EventCallAssoc (EventID, CallID) VALUES (" & eId & ", " & callId & ")"
    'execute the insert sql for call and alarm association
    DoCmd.RunSQL strSql
    
    'let the user know that the alarm has been saved
    MsgBox "Event has been saved!"
    'set the sql injection warnings back on
    DoCmd.SetWarnings True
    'refresh the call log
    Forms!CAD!Child4.Form.Requery
    'refresh the active events
    Forms!CAD!Child0.Form!ActiveEventHolder.Form.Requery
    'close the enter alarm form
    DoCmd.Close acForm, "EnterAlarm", acSaveNo
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
    'clear the form when it is closed
    Me.Undo
    'reset the sql injection warnings
    DoCmd.SetWarnings True
End Sub

Private Sub Form_Load()
    'set a variable that a new record is to be created on loading of the form
    CreateNewRecord
    'disable the sql injection warnings
    DoCmd.SetWarnings False
    'set the cancel button to have the focus when the form loads
    Me.Command4.SetFocus
End Sub

'cancel button
Private Sub Command4_Click()
    'undo any changes without saving the record
    Me.Undo
    'close the alarm entry page
    DoCmd.Close acForm, "EnterAlarm", acSaveNo
End Sub

'search the tables to pull up an address if a park is selected
Private Sub Text14_Exit(Cancel As Integer)
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
