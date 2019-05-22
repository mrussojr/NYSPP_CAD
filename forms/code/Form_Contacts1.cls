Attribute VB_Name = "Form_Contacts1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
    If Not IsNull(Me.OpenArgs) Then
        Me.Filter = "Shield = " & Me.OpenArgs
        Me.FilterOn = True
    End If
End Sub
