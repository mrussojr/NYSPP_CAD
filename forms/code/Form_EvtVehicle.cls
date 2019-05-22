Attribute VB_Name = "Form_EvtVehicle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command18_Click()
    'check that the parks event id field is filled in
    If IsNull(Forms!ViewEvent.Text38.Value) = False And Not Forms!ViewEvent.Text38.Value = 0 Then
        'if the parks event id is filled in, set the vehicle id field to the id of the vehicle being clicked
        Forms!ViewEvent.VehicleID.Value = Me.Text18.Value
        'open the existing event form with the parks event id, parks officer id, and vehi parameters
        OpenExistingEvent Forms!ViewEvent.Text38.Value, Forms!ViewEvent.Text36.Value, "VEHI"
    Else
        'advise the user that the parks event has not been created
        MsgBox "An event has not been created in PARKS yet!"
        Exit Sub
    End If
End Sub
