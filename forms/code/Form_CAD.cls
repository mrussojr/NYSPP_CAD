Attribute VB_Name = "Form_CAD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub CallEntry_Click()
    'open form to enter generic call
    DoCmd.OpenForm "InputForm"
End Sub

Private Sub Command14_Click()
    'open the menu to perform searches against the comm log
    DoCmd.OpenForm "SearchLog"
End Sub

Private Sub Command15_Click()
    'open the menu to generate reports for a specific month
    DoCmd.OpenForm "ReportGenerator"
End Sub

Private Sub Command28_Click()
    'opens the run data form for datas that are not going to be associated with an event
    DoCmd.OpenForm "RunData"
End Sub

Private Sub Command34_Click()
    'opens the change password form for the user to change their log in password
    DoCmd.OpenForm "ChangePassword", acNormal, , , , acWindowNormal, "CAD"
End Sub

Private Sub Command36_Click()
    'opens the administrative feature to edit officers that are in the CAD system, for event entry and attendance purposes
    DoCmd.OpenQuery "OffEdit", acViewNormal, acEdit
End Sub

Private Sub Command37_Click()
    'opens the search events page
    DoCmd.OpenForm "SearchEvents"
End Sub

Private Sub Command42_Click()
    'opens the panic book for the central zone
    DoCmd.OpenForm "ERG", acNormal, , , acFormReadOnly, acWindowNormal, "PANIC_BOOK_-_CEN"
End Sub

Private Sub Command43_Click()
    'opens the panic book for the ti zone
    DoCmd.OpenForm "ERG", acNormal, , , acFormReadOnly, acWindowNormal, "PANIC_BOOK_-_TI"
End Sub

Private Sub Command44_Click()
    'opens the panic book for the fl zone
    DoCmd.OpenForm "ERG", acNormal, , , acFormReadOnly, acWindowNormal, "PANIC_BOOK_-_FL"
End Sub

'open resources form
Private Sub Command45_Click()
    DoCmd.OpenForm "Resources"
End Sub

'enter a hang up call into the comm log
Private Sub Command48_Click()
    quickCallEntry "HANG UP", False
End Sub

'enter status checks for multiple units all at once
Private Sub Command49_Click()
    quickCallEntry "STATUS", True
End Sub

'enter call into comm log where no one is on the line
Private Sub Command50_Click()
    quickCallEntry "OTHER", False, "NO ONE ON THE LINE"
End Sub

'enter a wrong number call into the log
Private Sub Command51_Click()
    quickCallEntry "OTHER", False, "WRONG NUMBER"
End Sub

'enter a quick call into the comm log
Public Sub quickCallEntry(callReason As String, inputBoxReq As Boolean, Optional Narrative As String)
    Dim CurTime As String, today As String, Reason As String, sql As String, Narr As String, UnitCalled As String, SourceCall As String
    Dim DispId As Integer, y As Integer
    
    'disable sql injection warnings
    DoCmd.SetWarnings False
    
    'set the time of the call to right now
    CurTime = RightNow()
    'set the date of the call to today
    today = TodayLong()
    'set the reason to the supplied reason
    Reason = callReason
    'get the id of the dispatcher who entered the call
    DispId = GetDispatchId()
    
    'check if input box is required and check call reason
    If inputBoxReq And callReason = "STATUS" Then
        'if call reason is status
        'set the narrative to what the user enters in the input box
        Narr = InputBox("Please enter the units status checked.", "Status Check")
        'set the unit called to all, typical of nightly status checks
        UnitCalled = "ALL"
        'set the source of the call to MS, typical of nightly status checks
        SourceCall = "MS"
    
        'if the dialog is cancelled then exit here and do not enter any information
        If StrPtr(Narr) = 0 Then
            Exit Sub
        End If
    ElseIf inputBoxReq And callReason = "NEG RESP" Then
        'if call reason is negative response
        'set unit called to what the user enters in the input box
        UnitCalled = InputBox("Enter the unit that didn't respond.", "Negative Response")
        
        'if the unit called dialogis cancelled then exit here and do not enter call info in call log
        If StrPtr(UnitCalled) = 0 Then
            Exit Sub
        End If
        
        'set the source of the call to MS
        SourceCall = "MS"
        'set the narrative to what the user enters in the input box
        Narr = InputBox("Enter the method of communication used to contact the unit.", "Method of Communication")
        
        'if the method of communication dialog is cancelled then exit here and do not enter any information
        If StrPtr(Narr) = 0 Then
            Exit Sub
        End If
    Else
        'set the narrative to the supplied narrative
        Narr = Narrative
        'set the unit called to nothing
        UnitCalled = ""
        'set the source of the call to be phone
        SourceCall = "PHONE"
    End If
    
    'generate the sql to enter the call into the comm log
    sql = "INSERT INTO commlog (Date1, Time1, UnitCalled, SourceCall, Reason, Narrative, Dispatcher) VALUES ('" & today & "', '" & CurTime & "', '" & UnitCalled & "', '" & SourceCall & "', '" & Reason & "', '" & Narr & "', " & DispId & ")"
    
    'execute the sql to enter the call into the comm log
    CurrentDb.Execute sql
    
    'refresh the call log
    Forms!CAD.Child4.Requery
    
    'reset the sql injection warnings
    DoCmd.SetWarnings True
End Sub

'open the alarm entry form
Private Sub Command52_Click()
    DoCmd.OpenForm "EnterAlarm"
End Sub

'enter a negative reponse call into the comm log
Private Sub Command65_Click()
    quickCallEntry "NEG RESP", True
End Sub

'refresh the in service and out of service officers
Private Sub Command66_Click()
    'run the update assignments procedure
    UpdateAssignments
End Sub

'opens the ejustice portal and navigates to the inbox
Private Sub Command67_Click()
    DoCmd.OpenForm "EJustice", , , , , , "INBOX"
End Sub

'opens the previous events
Private Sub Command68_Click()
    DoCmd.OpenForm "PrevEventHolder"
End Sub

'open the schedule
Private Sub Command69_Click()
    Dim xlApp As Excel.Application
    Dim fileName As String
    
    fileName = Nz(GetFileName(), "")
    
    If fileName = "" Then
        Exit Sub
    End If
    
    'set xlApp to an instance of the excel application
    Set xlApp = CreateObject("Excel.Application")
    'set excel to be visible
    xlApp.Visible = True
    'open the schedule workbook
    xlApp.Workbooks.Open "\\oprhp-smb\oprhp_shared\park police\Midstate\Schedules\" & fileName
End Sub

'reset the ejustice login info
Private Sub Command70_Click()
    'set sql warnings off
    DoCmd.SetWarnings False
    'run sql to reset login
    DoCmd.RunSQL "UPDATE users SET LoggedInEJustice = 0 WHERE username = '" & GetUserName & "'"
    'set sql warnings back on
    DoCmd.SetWarnings True
    
    'advise user that login was reset
    MsgBox "Login reset!"
End Sub

'open the status check sheet
Private Sub Command71_Click()
    Dim xlApp As Excel.Application
    
    'set xlApp to an instance of the excel application
    Set xlApp = CreateObject("Excel.Application")
    'set excel tobe visible
    xlApp.Visible = True
    'open the schedule workbook
    xlApp.Workbooks.Open "\\cenfile1\police$\Dispatching disk\databases\Status Check Sheet\Status_Checker_2.0b.xlsm"
    xlApp.Run "Fill_Sheet"
End Sub

'open the form to lookup court information
Private Sub CourtLookup_Click()
    DoCmd.OpenForm "CourtLookUp"
End Sub

'open the ejustice form to log in, run data, or check inbox messages
Private Sub EJustice_Click()
    DoCmd.OpenForm "EJustice"
End Sub

'open the event entry page to enter an event or call for service
Private Sub EventEntry_Click()
    DoCmd.OpenForm "EnterEvent"
End Sub

Private Sub Form_Close()
    'close the inbox manager once the cad form is closed
    DoCmd.Close acForm, "InboxManager"
End Sub

Private Sub Form_Load()
    'if user is an administrator, load the admin features onto the cad, else make sure they are invisible to a regular user
    If DLookup("[Admin]", "users", "[Username] = '" & GetUserName() & "'") = 1 Then
        Me.Frame22.Visible = True
        Me.ApproveUsers.Visible = True
        Me.RemoveUsers.Visible = True
        Me.Command36.Visible = True
    Else
        Me.Frame22.Visible = False
        Me.ApproveUsers.Visible = False
        Me.RemoveUsers.Visible = False
        Me.Command36.Visible = False
    End If
    
    'if the inbox manager is not loaded, open it
    If Not CurrentProject.AllForms("InboxManager").IsLoaded Then
        DoCmd.OpenForm "InboxManager", acNormal, , , , acHidden
    End If
End Sub

Private Sub Form_Open(Cancel As Integer)
    'give focus to the active events section upon opening the cad system
    Me.Child0.SetFocus
End Sub

Private Sub Form_Timer()
    'refresh the call log section based on the form timer
    If Application.CurrentObjectType = acForm Then 'Screen.ActiveForm.NAME = "CAD" Then
        Me.Child4.Requery
    End If
End Sub

'open the approve users admin function form
Private Sub ApproveUsers_Click()
    DoCmd.OpenForm "AuthUsers"
End Sub

'open the remove user admin function form
Private Sub RemoveUsers_Click()
    DoCmd.OpenForm "RemUsers"
End Sub
