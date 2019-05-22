Attribute VB_Name = "Form_ReportGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'generate monthly report button
Private Sub Generate_Click()
    Dim mos, yr, Day, query As String
    Dim Date1, Date2 As Date
    
    'set the month variable to the value of the month drop down box
    mos = Me.Combo1.Value
    'set the year variable to the value of the year drop down box
    yr = Me.Combo2.Value
        
    'check whether the month and year have been filled in
    If IsNull(mos) = True Or IsNull(yr) = True Then
        'if the month or the year has not been entered, advise the user that both the month and year must be selected
        MsgBox ("Please be sure to select a month and year")
    Else
        'if the month and year are both selected
        'set the day variable to the number of days in the month of the year selected
        Day = DaysInMonth(mos & "/" & yr)
    
        'set the date1 variable to the first of the month and year selected
        Date1 = CDate(mos & "/01/" & yr)
        'set the date1 variable to midnight of the first day of the month and year selected
        ''Date1 = CDate(Date1 & " 00:01")
        'set the date2 variable to the last day of the month and year selected
        Date2 = CDate(mos & "/" & Day & "/" & yr)
        'set the date2 variable to 11:59 of the last day of the month and year selected
        ''Date2 = CDate(Date2 & " 23:59")
        
        'generate the where condition for the report button
        query = "(Expr2 BETWEEN #" & Date1 & "# AND #" & Date2 & "#)"
        'open the report with the where condition
        DoCmd.OpenReport "commlog", acViewPreview, , query
    End If
End Sub
