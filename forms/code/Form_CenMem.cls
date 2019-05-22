Attribute VB_Name = "Form_CenMem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'populate call form with caller info
Private Sub Command127_Click()
    Dim Isv As Integer
    Dim caller As String
    
    Isv = DLookup("[Isv]", "parksIDs", "[pId] =" & Me.oID.Value)
    
    If Isv = 0 Then
        caller = CStr(DLookup("[offShield]", "parksIDs", "[pId] =" & Me.oID.Value))
    ElseIf Isv = 1 Then
        caller = DLookup("[CarNo]", "parksIDs", "[pId] =" & Me.oID.Value)
    Else
        Exit Sub
    End If
    
    DoCmd.OpenForm "InputForm"
    Forms!InputForm!SourceCall.Value = caller
End Sub

'open info form
Private Sub Command128_Click()
    Dim SHIELD As Integer
    
    SHIELD = DLookup("[offShield]", "parksIDs", "[pId] = " & Me.oID.Value)
    
    DoCmd.OpenForm "Contacts1", , , , , , SHIELD
End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'    If Me.Text.FontUnderline = True Then
'        Me.Text.FontUnderline = False
'    End If
End Sub

Private Sub Detail_Paint()
'    Dim x As Integer
'
'    x = 1
'
'    Do While x < Me.Controls.Count
'        Me.Controls(x).Tag = x
'        x = x + 1
'    Loop
End Sub

Public Sub Text_Click()
    If IsNull(Me.Isv.Value) = True Or Me.Isv.Value = 0 Then 'check to see if the clicked officer is in service
        If EnterInService(Me.oID.Value) = 0 Then 'create the enter in service dialog for the given officer id
            'if enter in service dialog is cancelled then exit without saving any information
            Exit Sub
        End If
        'open parks assignment page with given officer parks id so their assignement can be created
        OpenParks (Me.oID.Value)
        'refresh the call log to show most recent call
        Forms!CAD.Child4.Requery
    Else
        If EnterOutService(Me.oID.Value) = 0 Then 'create the enter out of service dialog for the given officer id
            'if enter out of service dialog is cancelled then exit without saving any information
            Exit Sub
        End If
        'open the parks assignment page with the given officer parks id so their assignment can be completed
        OpenParks (Me.oID.Value)
        'refresh the call log to show most recent call
        Forms!CAD.Child4.Requery
    End If
End Sub

Private Sub Text_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'    If Me.Text.FontUnderline = False Then
'        Me.Text.FontUnderline = True
'    End If
End Sub

Private Sub Event_Click()
    'open the parks event page for the given officer id, this will not make any entries into the comm log, used to back log any missed events
    OpenEvent (Me.oID.Value)
End Sub
