Attribute VB_Name = "Form_Calls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Narrative_Click()
    If IsNull(Me.Narrative.Value) = False Then 'check if there is a narrative to click on
        If Len(Me.Narrative.Value) > 20 Then
            'if the narrative is longer than 20 characters, open it in a message box
            MsgBox Me.Narrative, , Me.SourceCall
        End If
        Dim clipboard As MSForms.DataObject
        Set clipboard = New MSForms.DataObject
        'set the narrative text to the clipboard, available for pasting into other applications
        clipboard.SetText Me.Narrative.Value
        clipboard.PutInClipboard
    End If
End Sub

Private Sub SourceCall_Click()
    If IsNull(Me.SourceCall.Value) = False Then 'check if there a source caller to click on
        If Len(Me.SourceCall.Value) > 4 Then
            'if the source caller is longer than 4 characters, open it in a message box
            MsgBox (Me.SourceCall)
        End If
        Dim clipboard As MSForms.DataObject
        Set clipboard = New MSForms.DataObject
        'set the source called text to the clipboard, available for pasting into other applications
        clipboard.SetText Me.SourceCall.Value
        clipboard.PutInClipboard
    End If
End Sub
