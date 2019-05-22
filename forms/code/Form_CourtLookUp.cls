Attribute VB_Name = "Form_CourtLookUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub County_Change()
    If IsNull(Me.County.Value) = True Or Me.County.Value = "" Then
        'remove the filter for the courts when the county is deselected
        Me.FilterOn = False
        'refresh the form without the filter on
        Me.Requery
    Else
        'set the filter on the court form so that only courts in the selected county are shown
        Me.Filter = "County = '" & Me.County.Value & "'"
        'set the filter to be on
        Me.FilterOn = True
        'refresh the form with the filter on
        Me.Requery
    End If
End Sub

'when a county is typed in that is not in the list, remove the filter and refresh to show all courts
Private Sub County_NotInList(NewData As String, Response As Integer)
    Me.FilterOn = False
    Me.Requery
End Sub

'when the form is loaded, make sure that no filters are on so that every court is shown
Private Sub Form_Load()
    Me.FilterOn = False
    Me.Filter = ""
End Sub

'close the court look up form
Private Sub Label12_Click()
    DoCmd.Close acForm, "CourtLookUp"
End Sub

Private Sub SearchBox_Change()
    If IsNull(Me.County.Value) = True Or Me.County.Value = "" Then 'set the filter to not include the county
        'set the filter off so it can be changed
        Me.FilterOn = False
        'set the filter to blank so it can be changed
        Me.Filter = ""
        'set the filter to return courts that match what the user has either typed or started typing into the court lookup form
        Me.Filter = "Jurisdiction ALike '%" & Me.SearchBox.Value & "%'"
        'turn the filter back on with the new requirements
        Me.FilterOn = True
        'refresh the court lookup form with the newly applied filter
        Me.Requery
        'set the search box to have the focus so that the user may continue to type if necessary
        Me.SearchBox.SetFocus
        'set the cursor to the end of the search box so the user doesn't need to retype what has already been entered
        Me.SearchBox.SelStart = Nz(Len(Me.SearchBox), 0)
        Me.SearchBox.SelLength = Nz(Len(Me.SearchBox), 0)
    Else 'set the filter to include the selected county
        'turn off the filter so it can be changed
        Me.FilterOn = False
        'set the filter to blank so it can be changed
        Me.Filter = ""
        'set the filter to return courts that match what the user has either typed or started typing into the court lookup form where the court is in the selected county
        Me.Filter = "Jurisdiction ALike '%" & Me.SearchBox.Value & "%' And County = '" & Me.County.Value & "'"
        'turn the filter back on with the new requirements
        Me.FilterOn = True
        'refresh the court look up form with the new filter
        Me.Requery
        'set the search box to have the focus so that the user may continue to type if necessary
        Me.SearchBox.SetFocus
        'set the cursor to the end of the search box so the user doesn't need to retype what has already been entered
        Me.SearchBox.SelStart = Nz(Len(Me.SearchBox), 0)
        Me.SearchBox.SelLength = Nz(Len(Me.SearchBox), 0)
    End If
End Sub

'open the court information report with the court's information that was clicked
Private Sub StateFull_Click()
    DoCmd.OpenReport "CourtInfo", acViewReport, , "Jurisdiction = '" & Me.StateFull.Value & "'"
End Sub
