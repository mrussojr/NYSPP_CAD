Attribute VB_Name = "Form_ActiveEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub AddTo_Click()
    'Add calls to the associated event
    DoCmd.OpenForm "InputForm", acNormal, , , , , Me.idBox.Value
End Sub

Private Sub CloseEvent_Click()
    'Close out event
    Dim x As Integer
    x = Me.idBox.Value
    
    'Enter clear time into comm log
    EnterClearTime (x)
    'Make event inactive
    ClearEvent (x)
    'refresh active events
    Me.Requery
End Sub

Public Sub EnterClearTime(x As Integer)
    Dim caller As String, CurTime As String, today As String, Reason As String, evtType As String, sql As String, sql2 As String, sql3 As String
    Dim DispId As Integer, y As Long
       
    'disable sql injection warnings
    DoCmd.SetWarnings False
    
    'lookup officer id for caller of clear call
    caller = DLookup("[Off]", "events", "[ID] = " & x)
    'lookup event type for narrative of clear call
    evtType = DLookup("[Type]", "events", "[ID] = " & x)
    'clear call time is now
    CurTime = RightNow()
    'clear date is today
    today = TodayLong()
    'set reason of call to clear for call
    Reason = "CLR"
    'get id of dispatcher entering call for clear call
    DispId = GetDispatchId()
    
    'generate sql to insert the clear call into the comm log
    sql = "INSERT INTO commlog (Date1, Time1, SourceCall, Reason, Narrative, Dispatcher) VALUES ('" & today & "', '" & CurTime & "', '" & caller & "', '" & Reason & "', '" & evtType & "', " & DispId & ")"
    'generate sql to get the id of the call just generated to associate it with the event being cleared
    sql2 = "SELECT @@Identity"
    
    'execute insert sql
    CurrentDb.Execute sql
    'obtain id of last call entered
    y = CurrentDb.OpenRecordset(sql2)(0)
    
    'generate sql to insert id of last call and event id to associate them together
    sql3 = "INSERT INTO EventCallAssoc (EventID, CallID) VALUES (" & x & ", " & y & ")"
    'execute association sql
    CurrentDb.Execute sql3
    
    'update the last time field of the event to right now, today
    UpdateLastTime CurTime, today, x
    
    'reenable sql injection warnings
    DoCmd.SetWarnings True
End Sub

Private Sub Command14_Click()
    'run a data to be associated with this event, will pass the event id to the run data form
    DoCmd.OpenForm "RunData", acNormal, , , , , Me.idBox.Value
End Sub

Private Sub Command18_Click()
    'open the event for viewing, passes the event id to the view event form
    DoCmd.OpenForm "ViewEvent", acNormal, , , , , Me.idBox.Value
End Sub

Private Sub Command19_Click()
    'enter a status check for this event into the comm log
    StatusCheck (Me.idBox.Value)
End Sub

Public Sub StatusCheck(x As Integer)
    Dim caller As String, CurTime As String, today As String, Reason As String, evtType As String, sql As String, sql2 As String, sql3 As String
    Dim DispId As Integer, y As Long
           
    'disable sql injection warnings
    DoCmd.SetWarnings False
    
    'look up officer for status call
    caller = DLookup("[Off]", "events", "[ID] = " & x)
    'look up event type for narrative of status call
    evtType = DLookup("[Type]", "events", "[ID] = " & x)
    'create time of right now for status call
    CurTime = RightNow()
    'create date of today for status call
    today = TodayLong()
    'create reason of status for status call
    Reason = "STATUS"
    'look up dispatcher id for status call
    DispId = GetDispatchId()
    
    'generate sql for status call
    sql = "INSERT INTO commlog (Date1, Time1, UnitCalled, SourceCall, Reason, Narrative, Dispatcher) VALUES ('" & today & "', '" & CurTime & "', '" & caller & "', 'JV', '" & Reason & "', '" & evtType & "', " & DispId & ")"
    'generate sql to grab call id of status call for event / call association
    sql2 = "SELECT @@Identity"
    
    'execute insert sql for status call
    CurrentDb.Execute sql
    'execute sql to get status call id
    y = CurrentDb.OpenRecordset(sql2)(0)
    
    'generate sql to associate the status call and the event
    sql3 = "INSERT INTO EventCallAssoc (EventID, CallID) VALUES (" & x & ", " & y & ")"
    'execute the sql to associate the status call and the event
    CurrentDb.Execute sql3
    
    'update the last time field of the event to right now, today
    UpdateLastTime CurTime, today, x
    
    'refresh the active events to update the color if necessary
    Me.Refresh
    'set sql injection warnings back on
    DoCmd.SetWarnings True
End Sub

Private Sub Form_Timer()
'    With Me!Officer
'        .ForeColor = IIf((.ForeColor) = RGB(0, 183, 239), vbRed, RGB(0, 183, 239))
'    End With

    If Application.CurrentObjectType = acForm Then 'And Screen.ActiveForm.NAME = "CAD" And IsLoaded("CAD") Then
        Dim db As Database
        Dim rs As DAO.Recordset
        Dim x As Integer
        Dim s As String
        
        'set empty string to keep track whether or not all calls need action
        s = ""
        
        Set db = CurrentDb
        Set rs = db.OpenRecordset("ActiveEventsQuery")
        
        'loop through all records in active event query
        Do While Not rs.EOF
            'if the time the last time any action was completed on the call is over 5 minutes from right now then set the call to red for action
            If rs.Fields("Expr1") >= 5 Or rs.Fields("Expr1") < 0 Then
                'only set color code if not already set
                If Not rs.Fields("ColorCode") = 0 Then
                    With rs
                        .Edit
                            .Fields("ColorCode") = 0
                        .Update
                    End With
                End If
                s = s & "0"
            Else 'set color back to blue once action is completed and time is less than 5 minutes from right now
                'only set color code if not already set
                If Not rs.Fields("ColorCode") = 1 Then
                    With rs
                        .Edit
                            .Fields("ColorCode") = 1
                        .Update
                    End With
                End If
                s = s & "1"
            End If
    
            'move to next record as long as it isn't the last
            rs.MoveNext
        Loop
        
        rs.Close
        
        Set rs = Nothing
        
        'only refresh the active events window if any calls don't need action, if all calls need action do not refresh until action performed, increase performance
        If InStr(1, s, "1") > 0 Then
            Me.Requery
        End If
    End If
End Sub
