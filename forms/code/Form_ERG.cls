Attribute VB_Name = "Form_ERG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

'when opening the panic book form, determination to which zone's info should be loaded comes from the forms open arguments
'these arguments are provided when the user clicks the specific button on the main cad form
Private Sub Form_Load()
    If Not Me.OpenArgs = "" Then
        Me.RecordSource = Me.OpenArgs
    End If
    
    DoCmd.SetOrderBy "[Park_Facility]"
End Sub
