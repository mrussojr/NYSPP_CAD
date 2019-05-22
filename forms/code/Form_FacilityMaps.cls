Attribute VB_Name = "Form_FacilityMaps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public mapType As Integer
Public parkId As Integer

'set the park id to the selected value of the drowdown when it changes
Private Sub Combo7_AfterUpdate()
        parkId = 0
        If Not IsNull(Me.Combo7.Value) Then
            parkId = Me.Combo7.Value
        End If
        FilterMaps
End Sub

'set the park id to the selected value of the dropdown when it changes
Private Sub Combo7_Change()
        parkId = 0
        If Not IsNull(Me.Combo7.Value) Then
            parkId = Me.Combo7.Value
        End If
        FilterMaps
End Sub

Private Sub Form_Load()
    If IsNull(Me.OpenArgs) = False Then
        parkId = Me.OpenArgs
        mapType = 4
        Me.Combo7.Value = Me.OpenArgs
    End If
    FilterMaps
End Sub

'on tab change, change mapType to maptype id of maptype selected by tab
Private Sub TabCtl1_Change()
    Select Case Me.TabCtl1.Value
        Case 0
            mapType = 4
        Case 1
            mapType = 2
        Case 2
            mapType = 3
        Case 3
            mapType = 1
        Case 4
            mapType = 5
    End Select
    FilterMaps
End Sub

Private Sub FilterMaps()
    Dim filterString As String
    
    If IsNull(parkId) = True Then
        filterString = ""
    Else
        filterString = "parkId = " & parkId & " AND mapType = " & mapType
    End If
    
    With Me!maps.Form
        .Filter = filterString
        .FilterOn = True
        .Requery
        .Refresh
    End With
End Sub
