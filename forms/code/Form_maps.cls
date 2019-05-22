Attribute VB_Name = "Form_maps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

'open file on click
Private Sub mapName_Click()
    FollowHyperlink Me.MapLocation.Value
End Sub
