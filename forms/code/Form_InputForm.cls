Attribute VB_Name = "Form_InputForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Command51_Click()
    'open the event form
    DoCmd.OpenForm "EnterEvent"
    
    'set the date from the call screen
    Forms!EnterEvent!Text0.Value = Nz(Me.Date1.Value, TodayLong())
    
    'set the time from the call screen
    Forms!EnterEvent!Text2.Value = Nz(Me.Time1.Value, RightNow())
    
    'set the caller from the call screen if entered
    If Not IsNull(Me.SourceCall.Value) Then
        Forms!EnterEvent!Text8.Value = Me.SourceCall.Value
    End If
    
    'set the narrative from the call screen
    Forms!EnterEvent!Text17.Value = Nz(Me.Narrative.Value, "")
    
    'undo the changes and don't save record
    Me.Undo
    'close the input form
    DoCmd.Close acForm, "InputForm"
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    'before saving record, check if the ready to save record variable has been changed to true
    If Command = acCmdSaveRecord Then
        If ReadyToSaveRecord = False Then
            'cancel the save command
            Cancel = True
        Else
            'don't cancel the save command
            Cancel = False
        End If
    End If
End Sub

'save button
Private Sub cmdSave_Click()
    Dim UnitCalled As String, sql As String
    Dim idNum As Integer, oID As Integer
    
    'disable sql injection warnings
    DoCmd.SetWarnings False
    
    'check if call is outgoing or incoming
    If Me.Text43.Value = "X" Then
        'if x then outgoing call, set unit called to source call
        Me.UnitCalled.Value = Me.SourceCall.Value
        'UnitCalled = Me.SourceCall.Value
        'set source call to jv
        Me.SourceCall.Value = "MS"
    ElseIf Me.Text43.Value = "" Or IsNull(Me.Text43.Value) = True Then
        'UnitCalled = ""
        'set unit called to empty string and leave source call as source call
        Me.UnitCalled.Value = ""
    End If
    
    'check if reason is filled out
    If IsNull(Me.Reason.Value) Or Me.Reason.Value = "" Then
        'cancel save command and advise that it needs to be filled out
        MsgBox "Please fill out the reason code!"
        Exit Sub
    End If
    
    'run the save record command
    DoCmd.RunCommand acCmdSaveRecord
    'DoCmd.RunSQL "INSERT INTO commlog (Date1, Time1, UnitCalled, SourceCall, Reason, Dispatcher) VALUES ('" & TodayLong() & "', '" & RightNow() & "', '" & UnitCalled & "', '" & Me.SourceCall.Value & "', '" & Me.Reason.Value & "', " & GetDispatchId & ")"
    
    'check if the opening arguments exist, if they exist this call will be associated with an event
    If IsNull(Me.OpenArgs) = False Or Not Me.OpenArgs = "" Then
        'if open arguments
        'run sql to insert association between call and event
        DoCmd.RunSQL "INSERT INTO EventCallAssoc (EventID, CallID) VALUES (" & Me.OpenArgs & ", " & Me.idField & ")"
        'update the last time in event table for the event
        UpdateLastTime Me.Time1.Value, Me.Date1.Value, Me.OpenArgs
        
        'if reason of call is clear, clear the event out of the active events
        If Me.Reason.Value = "CLR" Then
            ClearEvent Me.OpenArgs
        End If
        
        'refresh active events and call log
        RqryChildren
    End If
    
    If Me.Reason.Value = "ISV" Then
        oID = Nz(DLookup("[pId]", "parksIDs", "[offShield] = " & Me.SourceCall.Value & " And [Active] = 1"), 0)
        If Not oID = 0 Then
            OpenParks (oID)
        End If
    End If
    
    'Me.Undo
    'reset sql injection warnings
    DoCmd.SetWarnings True
    'refresh active events and call log
    RqryChildren
    'close the input call form
    DoCmd.Close acForm, "InputForm", acSaveNo
End Sub

Private Sub Form_Load()
'    Dim x As Integer, y As Integer
'
'    x = DCount("[id]", "parksIDs", "[Isv] = 1")
'    y = 1
'
'    Do While y < x
'        Me.SourceCall.AddItem DLookup("[CarNo]", "parksIDs", "[Isv] = 1")
'        y = y + 1
'    Loop
    
    Dim rs As DAO.Recordset
    Dim db As Database
    Dim strSql As String
    
    Set db = CurrentDb
    
    'generate sql to select all in service active officers
    strSql = "SELECT * FROM parksIDs WHERE Isv = 1 ORDER BY Avail DESC"
    
    Set rs = db.OpenRecordset(strSql)
    
    'for each record in the query, add it to the caller drop down
    Do While Not rs.EOF
        If IsNull(rs.Fields("CarNo")) = False And Not rs.Fields("CarNo") = "Not App" Then
            Me.SourceCall.AddItem rs!CarNo
        End If
        rs.MoveNext
    Loop
    
    'add the final values to the drop down box for caller
    Me.SourceCall.AddItem "STANLEY SEC"
    Me.SourceCall.AddItem "PHONE"
    Me.SourceCall.AddItem "EMER"
    
    'set the create new record variable to true
    CreateNewRecord
End Sub

'set the narrative to upper case when losing focus
Private Sub Narrative_Exit(Cancel As Integer)
    If IsNull(Me.Narrative.Value) = False And Not Me.Narrative.Value = "" Then
        Me.Narrative.Value = UCase(Me.Narrative.Value)
    End If
End Sub

'on exiting the reason code, validate against codes in reason table
Private Sub Reason_Exit(Cancel As Integer)
    Dim db As Database
    Dim rs As DAO.Recordset
    Dim typedReason As String
    Dim stringFound As Boolean
    
    'set string found to false for use in setting the field to OTHER if no string is matched
    stringFound = False
    
    'set to reason user typed, to uppercase
    typedReason = Nz(UCase(Me.Reason.Value), "")
    
    'check if reason is blank, then exit
    If IsNull(typedReason) Or typedReason = "" Then
        Exit Sub
    End If
    
    'open a recordset containing the reason codes from the reasoncodes tables for comparison
    Set db = CurrentDb
    Set rs = CurrentDb.OpenRecordset("SELECT [ReasonCode] From ReasonCodes")
    
    'loop through recordset to match reason code
    'do while there are records in the reason table
    While Not rs.EOF
        If rs.Fields(0).Value = typedReason Then
            'if reason code matches one in the table, exit procedure
            Exit Sub
        End If
    
        'move to the next record
        rs.MoveNext
    Wend
    
    'move the record back to the first
    rs.MoveFirst
    
    'loop through the recordset to use the string matcher function to find similar words
    While Not rs.EOF
        'use the string matcher to see if the typed reason matches a reason from the table
        If StringMatcher(typedReason, rs.Fields(0).Value) = True Then
          'if the typed reason is similar to a reason in the reason table then reset the reason field
          Me.Reason.Value = rs.Fields(0).Value
          stringFound = True
        End If
        
        'move to the next record
        rs.MoveNext
    Wend
    
    'if a string was not found, set to other
    If stringFound = False Then
        Me.Reason.Value = "OTHER"
    End If
    
    'set record set and database to nothing
    Set rs = Nothing
    Set db = Nothing
End Sub

'drop down the reason when a letter is pressed
Private Sub Reason_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode > 0 Then
        Reason.Dropdown
    End If
End Sub

'set the source caller to upper case when losing focus
Private Sub SourceCall_Exit(Cancel As Integer)
    If IsNull(Me.SourceCall.Value) = False And Not Me.SourceCall.Value = "" Then
        Me.SourceCall.Value = UCase(Me.SourceCall.Value)
    End If
End Sub

'pseudo check box out going call, set x on click and clear on click again
Private Sub Text43_Click()
    If Me.Text43.Value = "" Or IsNull(Me.Text43.Value) = True Then
        Me.Text43.Value = "X"
    ElseIf Me.Text43.Value = "X" Then
        Me.Text43.Value = ""
    End If
End Sub

'cancel button
Private Sub Command46_Click()
    'undo the changes and don't save record
    Me.Undo
    'close the input form
    DoCmd.Close acForm, "InputForm"
End Sub
