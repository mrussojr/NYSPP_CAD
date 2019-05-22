Attribute VB_Name = "Report_commlog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Report_Load()
    If Not IsNull(Me.OpenArgs) Then
        Me.Filter = Me.OpenArgs
        Me.FilterOn = True
        Me.FilterOnLoad = True
    End If
End Sub
