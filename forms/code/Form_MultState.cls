Attribute VB_Name = "Form_MultState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
    'when the form is loaded, set the filter off and set it to an empty string
    Me.FilterOn = False
    Me.Filter = ""
End Sub

'clear search box link
Private Sub Label10_Click()
    'set the search box to an empty string
    Me.SearchBox.Value = ""
    'set the filter to off
    Me.FilterOn = False
    'set the filter to an empty string
    Me.Filter = ""
    'refresh the form to show all states
    Me.Requery
End Sub

'remove all selected states link
Private Sub Label11_Click()
    'remove any previously selected states
    Me.StateSource.Value = ""
    'change the label to show that no states have been selected
    Me.Label7.Caption = "Selected (0/5):"
End Sub

'done link
Private Sub Label12_Click()
    Dim formName As String
    
    'check the opening arguments of the form, form should only be loaded from the run data form or the enter event form
    If Not IsNull(Me.OpenArgs) Then
        'set the form name variable to be the opening arguments of the multiple state form, this will be the form that called the multiple state form
        formName = Me.OpenArgs
        'find the opening form and set the state picker value to the selected states
        Forms.Item(formName).StatePicker.Value = Me.StateSource.Value
        'find the opening form and set the state field value to the selected states
        Forms.Item(formName).StateSource2.Value = Me.StateSource.Value
        'find the opening form and set the state field to visible
        Forms.Item(formName).StateSource2.Visible = True
        'close the multiple state picker form, without saving any changes, assures that the filter is empty and off when opened the next time
        DoCmd.Close acForm, "MultState", acSaveNo
    End If
End Sub

Private Sub SearchBox_Change()
    'turn off the filter
    Me.FilterOn = False
    'set the filter to an empty string
    Me.Filter = ""
    'set the filter to be what is typed in the search box
    Me.Filter = "StateFull ALike '" & Me.SearchBox.Value & "%'"
    'turn the filter back on
    Me.FilterOn = True
    'refresh the form to apply the filter
    Me.Requery
    'set the focus back to the search box so the user can continue typing
    Me.SearchBox.SetFocus
    'set the cursor to the end of the search box so that the user can continue typing and not have to retype what has already been entered
    Me.SearchBox.SelStart = Nz(Len(Me.SearchBox), 0)
    Me.SearchBox.SelLength = Nz(Len(Me.SearchBox), 0)
End Sub

'procedure for when a state is clicked
Private Sub StateFull_Click()
    Dim count As Integer
    Dim stSource As String
    
    'set the count of states that have been picked to 0
    count = 0
    
    'check the state source field
    If IsNull(Me.StateSource.Value) Or Me.StateSource.Value = "" Then
        'if nothing has been entered, add the state that was clicked
        Me.StateSource.Value = Me.State.Value
        'set the count to 1
        count = 1
        'set the selected label to show that 1 state has been seleced
        Me.Label7.Caption = "Selected (" & count & "/5):"
    Else
        'set the state source variable to the state source field value
        stSource = Me.StateSource.Value
        
        'check to see if the state that was clicked is not in the state source field
        If InStr(1, stSource, Me.State.Value) = 0 Then
            'check how many states are in the state source by counting commas
            If InStr(1, stSource, ",") = 0 Then
                'if there are no commas this indicates that there is only one state selected, so set the count to 2
                count = 2
                'add the state that was clicked to the state source field
                Me.StateSource.Value = stSource & "," & Me.State.Value
                'change the states selected label to show that 2 states have been selected
                Me.Label7.Caption = "Selected (" & count & "/5):"
            Else
                'set the count of states to length of the state source field minus the length of the state source field with the commas replaced with empty strings
                'essentially this is counting the number of commas in the state source field, this will give a count of how many states have been selected
                count = Len(stSource) - Len(Replace(stSource, ",", ""))
                'need to add 1 to the count to get the correct amount of states that have been selected, then another 1 to add the next state
                count = count + 2
                'check that the count of states selected is 5 or less, max of 5 is required by ejustice
                If count <= 5 Then
                    'if count is less than or equal to 5, add the selected state to the states selected field
                    Me.StateSource.Value = stSource & "," & Me.State.Value
                    'set the states selected label to show the number of states selected
                    Me.Label7.Caption = "Selected (" & count & "/5):"
                Else
                    'if count is more than five, advise user that they can only select 5 states, ejustice max
                    MsgBox "You can only select 5 states or codes."
                End If
            End If
        End If
    End If
End Sub
