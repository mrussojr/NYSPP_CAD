Attribute VB_Name = "Form_browser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim gotSecondary As Boolean

Public Function buttonClick(Index As Integer)
    On Error Resume Next ' just in case there is no page back or forward
    Select Case Index
        Case 0 'Go Back Button
        WebBrowser0.GoBack 'Go Back one Page
        Case 1 'Go Forward Button
        WebBrowser0.GoForward 'Go Forward one Page
        Case 2 'Stop Button
        WebBrowser0.Stop 'stop page
        Case 3 'Refresh Button
        WebBrowser0.Refresh 'refresh page
        Case 4 'Go Home Button
        WebBrowser0.GoHome 'Go to home page
        Case 5 'Search Button
        WebBrowser0.GoSearch 'Search
    End Select
End Function

Private Sub Command1_Click()
    buttonClick (0)
End Sub

Private Sub Command2_Click()
    buttonClick (1)
End Sub

Private Sub Command3_Click()
    buttonClick (3)
    Me.addressBar.Value = Me.WebBrowser0.LocationURL
End Sub

Private Sub Command4_Click()
    buttonClick (2)
End Sub

Private Sub Command7_Click()
    Dim webAddress, newAddress As String
    webAddress = Me.WebBrowser0.LocationURL
    newAddress = Me.addressBar.Value
   
    If Not newAddress = webAddress Then
        Me.WebBrowser0.Navigate newAddress
    End If
End Sub

Private Sub Command8_Click()
    Dim searchQuery, webAddress As String
    searchQuery = Me.searchBar.Value
    If Not searchQuery = Null Or Not searchQuery = "" Then
        webAddress = "http://www.google.com/search?q=" & searchQuery
        Me.WebBrowser0.Navigate webAddress
    End If
End Sub

Private Sub Form_Load()
    Dim webAddress As String
    'when form opens set the location to the arguments sent or google if opened independantly
    If IsNull(Me.OpenArgs) = False Then
        webAddress = Me.OpenArgs
    Else
        webAddress = "http://www.google.com"
    End If
    gotSecondary = False
    'navigate to address
    Me.WebBrowser0.Navigate webAddress
    'set the address bar value to the web address
    Me.addressBar.Value = webAddress
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'once the form is closed, open the UPD form to update the in-service officers
    'DoCmd.OpenForm "UPD"
End Sub

Private Sub searchBar_Enter()
    If Me.searchBar.Value = "Search..." Then
        Me.searchBar.Value = Null
    End If
End Sub
