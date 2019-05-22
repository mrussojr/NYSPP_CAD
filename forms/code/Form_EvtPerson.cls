Attribute VB_Name = "Form_EvtPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command18_Click()
    'check that the parks event id is filled
    If IsNull(Forms!ViewEvent.Text38.Value) = False And Not Forms!ViewEvent.Text38.Value = 0 Then
        'if parks event id has been captured, set the person id field on the view even form to the id of the person being clicked
        Forms!ViewEvent.PersonID.Value = Me.Text19.Value
        'open existing event with the parks event id, parks officer id, and peep parameters, this will take the event page to the person tab to be auto filled
        OpenExistingEvent Forms!ViewEvent.Text38.Value, Forms!ViewEvent.Text36.Value, "PEEP"
    Else
        'advise the user that the parks event has not been created yet as the parks event id has not been captured
        MsgBox "An event has not been created in PARKS yet!"
        Exit Sub
    End If
End Sub
