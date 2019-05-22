Attribute VB_Name = "Form_SearchEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command16_Click()
    Me.Subform1.Requery
End Sub

Private Sub zone_Exit(Cancel As Integer)
    Me.officerNam.Requery
End Sub


