Attribute VB_Name = "Form_SearchParks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Sub Form_Load()
    Dim webAddress As String, domain As String
    Dim args() As String
    
    domain = "http://policeblotter/Search"
    
    args = Split(Me.OpenArgs, "/", -1, vbBinaryCompare)
    
    If args(0) = "Persons" Then
        webAddress = "Persons.aspx?sLastName=" & args(1) & "&sFirstName=" & args(2)
    ElseIf args(0) = "LicensePlates" Then
        webAddress = "LicensePlates.aspx?sPlateNo=" & args(1)
    Else
        DoCmd.Close acForm, "SearchParks"
    End If
    
    'navigate the web browser to the web address
    Me.WebBrowser0.Navigate domain & webAddress
End Sub

