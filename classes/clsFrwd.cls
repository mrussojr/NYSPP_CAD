Attribute VB_Name = "clsFrwd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Dim oObject As Object
Dim sMethod As String
Dim bInstantiated As Boolean

Private Sub Class_Initialize()
    bInstantiated = False
End Sub

Public Sub Set_Destination(oInObject As Object, sInMethod As String)
    Set oObject = oInObject
    sMethod = sInMethod
    bInstantiated = True
End Sub

Public Function forwardEvent()
    If bInstantiated Then
        CallByName oObject, sMethod, VbMethod
    End If
End Function