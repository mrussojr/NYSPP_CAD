Attribute VB_Name = "Report_CourtInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub EMAIL_Click()
    If IsNull(Me.EMAIL.Value) = False And Not Me.EMAIL.Value = "" Then
        SendCourtEmail Me.EMAIL.Value
    End If
End Sub
