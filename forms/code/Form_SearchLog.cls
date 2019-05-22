Attribute VB_Name = "Form_SearchLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public dateFiltStr As String, timeFiltStr As String, calleeFiltStr As String
Public callerFiltStr As String, reasonFiltStr As String, idFiltStr As String
Public narrFiltStr As String

'time1 button
Private Sub Command6_Click()
    'set time1 value to right now
    Me.Time1.Value = RightNow()
End Sub

'time2 button
Private Sub Command9_Click()
    'set time2 value to right now
    Me.Time2.Value = RightNow()
End Sub

Private Sub Date1_BeforeUpdate(Cancel As Integer)
    If Not Me.Date1.Value = "" And Not IsNull(Me.Date1.Value) Then
        Me.Date2.Enabled = True
    Else
        Me.Date2.Enabled = False
    End If
End Sub

'on exit from date1 field, call build complete date function
Private Sub Date1_Exit(Cancel As Integer)
    BuildCompleteDate
End Sub

Private Sub Date2_BeforeUpdate(Cancel As Integer)
    BuildFilter
End Sub

'on exit from date2 field, call build complete date function
Private Sub Date2_Exit(Cancel As Integer)
    'BuildCompleteDate
    BuildFilter
End Sub

Private Sub Dispatcher_BeforeUpdate(Cancel As Integer)
    BuildFilter
End Sub

Private Sub Form_Load()
    Me.Date2.Enabled = False
    Me.Time2.Enabled = False
    Me.Command9.Enabled = False
    
    dateFiltStr = ""
    timeFiltStr = ""
    calleeFiltStr = ""
    callerFiltStr = ""
    reasonFiltStr = ""
    idFiltStr = ""
    narrFiltStr = ""
    
    With Me!SearchLog_subform.Form
        .Filter = ""
        .FilterOn = False
    End With
End Sub

Private Sub Narrative_BeforeUpdate(Cancel As Integer)
    BuildFilter
End Sub

Private Sub Reason_BeforeUpdate(Cancel As Integer)
    BuildFilter
End Sub

Private Sub SourceCall_BeforeUpdate(Cancel As Integer)
    BuildFilter
End Sub

Private Sub Time1_BeforeUpdate(Cancel As Integer)
    If Not Me.Time1.Value = "" And Not IsNull(Me.Time1.Value) Then
        Me.Time2.Enabled = True
        Me.Command9.Enabled = True
    Else
        Me.Time2.Enabled = False
        Me.Command9.Enabled = False
    End If
End Sub

'on exit from time1 field, call build complete date function
Private Sub Time1_Exit(Cancel As Integer)
    BuildCompleteDate
End Sub

Private Sub Time2_BeforeUpdate(Cancel As Integer)
    'BuildCompleteDate
    BuildFilter
End Sub

'on exit from time2 field, call build complete date function
Private Sub Time2_Exit(Cancel As Integer)
    'BuildCompleteDate
    BuildFilter
End Sub

'reset button
Private Sub Reset_Click()
    Dim cControl As control
        
    'loop through each control in the form
    For Each cControl In Me.Controls
        'check the type of control
        If TypeOf cControl Is TextBox Or TypeOf cControl Is ComboBox Then
            'if the control is a text box or combobox, set the value to null
            cControl = Null
        End If
        
    Next cControl
    
    With Me!SearchLog_subform.Form
        .Filter = ""
        .FilterOn = False
    End With
End Sub

'search button
Private Sub cmdSearch_Click()
    DoCmd.OpenReport "commlog", acViewPreview, , , , Me!SearchLog_subform.Form.Filter
End Sub

'function to build a complete date with the date and time
Private Sub BuildCompleteDate()
    Dim ddate1 As String, ddate2 As String, ttime1 As String, ttime2 As String
    
    'if date1 is empty, exit the function
    If IsNull(Me.Date1.Value) = True Then
        Exit Sub
    End If
    
    'set the date1 variable to the date1 value
    ddate1 = Me.Date1.Value
    'check the time1 value
    If IsNull(Me.Time1.Value) Then
        'if the time1 value is null
        'check the date2 value
        If IsNull(Me.Date2.Value) Then
            'if the date2 value is null
            'check the time2 value
            If IsNull(Me.Time2.Value) Then
                'if time2 value is null
                'set the time period between both dates to encompass one whole day
                'set time1 to 1200am
                ttime1 = "00:00"
                'set date2 to date1
                ddate2 = ddate1
                'set time2 to 1159pm
                ttime2 = "23:59"
            End If
        Else
            'if date2 is not null
            'check the time2 value
            If IsNull(Me.Time2.Value) Then
                'if time2 value is null
                'set the time period between both dates to start at 1200am on the first day to 1159pm on the second day
                'set time1 to 1200am
                ttime1 = "00:00"
                'set date2 to date2 value
                ddate2 = Me.Date2.Value
                'set time2 to 1159pm
                ttime2 = "23:59"
            End If
        End If
    Else
        'if the time1 value is not null
        'check the date2 value
        If IsNull(Me.Date2.Value) Then
            'if the date2 value is null
            'check the time2 value
            If IsNull(Me.Time2.Value) Then
                'if the time2 value is null
                'set the time period to search for to be time1 on date1
                'set time1 to the time1 value
                ttime1 = Me.Time1.Value
                'set date2 to the date1 value
                ddate2 = ddate1
                'set time2 to the time1 value
                ttime2 = ttime1
            Else
                'if the time2 value is not null
                'set the time period to be between time1 and time2 on date1
                'set time1 to the time1 value
                ttime1 = Me.Time1.Value
                'set date2 to the date1 value
                ddate2 = ddate1
                'set time2 to the time2 value
                ttime2 = Me.Time2.Value
            End If
        Else
            'if the date2 value is not null
            'check the time2 value
            If IsNull(Me.Time2.Value) Then
                'if the time2 value is null
                'set the time period to be between time1 on date1 to 1159pm on date2
                'set time1 to the time1 value
                ttime1 = Me.Time1.Value
                'set date2 to the date2 value
                ddate2 = Me.Date2.Value
                'set time2 to 1159pm
                ttime2 = "23:59"
            Else
                'if the time2 value is not null
                'set the time period to be between time1 on date1 and time2 and date2
                'set time1 to the time1 value
                ttime1 = Me.Time1.Value
                'set date2 to the date2 value
                ddate2 = Me.Date2.Value
                'set time2 to the time2 value
                ttime2 = Me.Time2.Value
            End If
        End If
    End If
        
    'set the complete date1 to be date1 and time1
    Me.completeDate1.Value = ddate1 & " " & ttime1
    'set the complete date2 to be date2 and time2
    Me.completeDate2.Value = ddate2 & " " & ttime2
End Sub

Public Sub BuildFilter()
    Dim filterString As String
    Dim filterCheckStr As String
    Dim filterAndClause As String
    Dim x As Integer
    Dim tempStr As String
    Dim filterArray As Variant
    
    setFilters
    
    filterArray = Array(dateFiltStr, timeFiltStr, calleeFiltStr, callerFiltStr, reasonFiltStr, idFiltStr, narrFiltStr)
    
    filterCheckStr = CheckFilters
    
    If filterCheckStr = "00000000" Then
        Me!SearchLog_subform.Form.FilterOn = False
        Exit Sub
    End If
        
    tempStr = filterCheckStr
    
    For x = LBound(filterArray) To UBound(filterArray)
        If x = UBound(filterArray) Then
            filterAndClause = ""
        Else
            If filterArray(x + 1) = "" Or IsNull(filterArray(x + 1)) Then
                filterAndClause = ""
            Else
                If filterString = "" Then
                    filterAndClause = ""
                Else
                    filterAndClause = " AND "
                End If
            End If
        End If
        
        filterString = filterString & filterArray(x) & filterAndClause
    Next
    
    Me.Text25.Value = filterString
    
    With Me!SearchLog_subform.Form
        .Filter = filterString
        .FilterOn = True
        .Refresh
        .Requery
    End With
End Sub

Public Function setFilters()
    dateFilter
    'timeFilter
    calleeFilter
    callerFilter
    reasonFilter
    idFilter
    narrFilter
End Function

Public Function dateFilter()
    If Not Me.Date2.Value = "" And Not IsNull(Me.Date2.Value) Then
        dateFiltStr = "Expr2 >= #" & DateLong(Me.Date1.Value) & "# AND Expr2 <= #" & DateLong(Me.Date2.Value) & "#"
    ElseIf Not Me.Date1.Value = "" And Not IsNull(Me.Date1.Value) Then
        dateFiltStr = "Expr2 = #" & DateLong(Me.Date1.Value) & "#"
    Else
        dateFiltStr = ""
    End If
End Function

Public Function timeFilter()
    If Not Me.Time2.Value = "" And Not IsNull(Me.Time2.Value) Then
        timeFiltStr = "Expr3 >= #" & CDate(Me.Time1.Value) & "# AND Expr3 <= #" & CDate(Me.Time2.Value) & "#"
    ElseIf Not Me.Time1.Value = "" And Not IsNull(Me.Time1.Value) Then
        timeFiltStr = "Expr3 = '" & CDate(Me.Time1.Value) & "'"
    Else
        timeFiltStr = ""
    End If
End Function

Public Function calleeFilter()
    If Not Me.UnitCalled.Value = "" And Not IsNull(Me.UnitCalled.Value) Then
        calleeFiltStr = "UnitCalled = '" & Me.UnitCalled.Value & "'"
    Else
        calleeFiltStr = ""
    End If
End Function

Public Function callerFilter()
    If Not Me.SourceCall.Value = "" And Not IsNull(Me.SourceCall.Value) Then
        callerFiltStr = "SourceCall = '" & Me.SourceCall.Value & "'"
    Else
        callerFiltStr = ""
    End If
End Function

Public Function reasonFilter()
    If Not Me.Reason.Value = "" And Not IsNull(Me.Reason.Value) Then
        reasonFiltStr = "Reason = '" & Me.Reason.Value & "'"
    Else
        reasonFiltStr = ""
    End If
End Function

Public Function idFilter()
    If Not Me.Dispatcher.Value = 0 And Not IsNull(Me.Dispatcher.Value) Then
        idFiltStr = "Dispatcher = '" & Me.Dispatcher.Value & "'"
    Else
        idFiltStr = ""
    End If
End Function

Public Function narrFilter()
    Dim strArray() As String
    Dim size As Integer, x As Integer
    Dim narrVal As String, concat As String
    
    narrVal = Nz(Me.Narrative.Value, "")
    narrFiltStr = ""
    
    If Not narrVal = "" Then
        size = (Len(narrVal) - Len(Replace(narrVal, ",", ""))) + 1
        If size <= 1 Then
            narrFiltStr = "Narrative ALike '%" & narrVal & "%'"
        Else
            strArray = Split(narrVal, ",")
            concat = " OR "
            
            For x = LBound(strArray) To UBound(strArray)
                strArray(x) = Trim(strArray(x))
                
                narrFiltStr = narrFiltStr & "Narrative ALike '%" & strArray(x) & "%'" & concat
                
                If x = UBound(strArray) - 1 Then
                    concat = ""
                End If
            Next
            narrFiltStr = "(" & narrFiltStr & ")"
        End If
    Else
        narrFiltStr = ""
    End If
        
End Function

Public Function CheckFilters() As String
    Dim str As String
    
    str = ""
    
    If IsNull(dateFiltStr) Or dateFiltStr = "" Then
        str = str + "0"
    Else
        str = str + "1"
    End If
    
    If IsNull(timeFiltStr) Or timeFiltStr = "" Then
        str = str + "0"
    Else
        str = str + "1"
    End If

    If IsNull(calleeFiltStr) Or calleeFiltStr = "" Then
        str = str + "0"
    Else
        str = str + "1"
    End If
    
    If IsNull(callerFiltStr) Or callerFiltStr = "" Then
        str = str + "0"
    Else
        str = str + "1"
    End If
    
    If IsNull(reasonFiltStr) Or reasonFiltStr = "" Then
        str = str + "0"
    Else
        str = str + "1"
    End If
    
    If IsNull(idFiltStr) Or idFiltStr = "" Then
        str = str + "0"
    Else
        str = str + "1"
    End If
    
    If IsNull(narrFiltStr) Or narrFiltStr = "" Then
        str = str + "0"
    Else
        str = str + "1"
    End If
    
    CheckFilters = str
End Function

Private Sub UnitCalled_BeforeUpdate(Cancel As Integer)
    BuildFilter
End Sub
