Attribute VB_Name = "Form_Members"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim zone As Integer

Private Sub TabCtl3_Change()
    Select Case Me.TabCtl3.Value
        Case 0
            zone = 5
        Case 1
            zone = 1
        Case 2
            zone = 6
    End Select
    
    If Me.Toggle25.Value = -1 Then
        Me!CenMem.Form.Filter = "pZone = " & zone & " AND Isv = 1"
    Else
        Me!CenMem.Form.Filter = "pZone = " & zone
    End If

    Me!CenMem.Form.FilterOn = "True"
    Me!CenMem.Form.Requery
End Sub

Private Sub Toggle25_Click()
    With Me.Toggle25
        If .Value = -1 Then
            .Caption = "Isv Only"
            Me!CenMem.Form.Filter = "pZone = " & zone & " AND Isv = 1"
        Else
            .Caption = "All"
            Me!CenMem.Form.Filter = "pZone = " & zone
        End If
        Me!CenMem.Form.FilterOn = "True"
        Me!CenMem.Form.Requery
    End With
End Sub
