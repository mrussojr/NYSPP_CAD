Attribute VB_Name = "Form_Home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Command5_Click()
    DoCmd.OpenForm "SearchLog", acNormal, , , , acWindowNormal
End Sub

Private Sub Command6_Click()
    DoCmd.OpenForm "ReportGenerator", acNormal, , , , acWindowNormal
End Sub

Private Sub Enter_Click()
    DoCmd.OpenForm "Form1", acNormal, , , , acWindowNormal
End Sub

Private Sub Image2_Click()
    Dim openArg As String
    openArg = "http://albweb7/policeblotter/"
    DoCmd.OpenForm "browser", acNormal, , , , acWindowNormal, openArg
End Sub
