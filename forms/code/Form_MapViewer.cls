Attribute VB_Name = "Form_MapViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public FrameAdded As Integer 'variable to make sure that the map is only loaded once

Private Sub Form_Load()
    'navigate the web browser to the location of the map web page
    Me.WebBrowser0.Navigate "\\cenfile1\police$\Dispatching disk\databases\TEST AREA\maps.htm"
    'set the variable that moniters whether the map was loaded or not
    FrameAdded = 0
    'silence any errors in the web browser
    Me.WebBrowser0.Silent = True
End Sub

Private Sub WebBrowser0_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
    Dim HTMLdoc As New HTMLDocument
    Dim pageSource As MSHTML.IHTMLElementCollection
    'Dim cElem2 As MSHTML.HTMLIFrame
    Dim cElem As MSHTML.IHTMLIFrameElement3
    
    'check if the web browser is not loaded and the map has not been loaded
    If Me.WebBrowser0.ReadyState = 3 And FrameAdded = 0 Then
        'MsgBox Me.WebBrowser0.ReadyState
        'grab the html document for parsing
        Set HTMLdoc = WebBrowser0.Document
        'get the iframe for manipulation
        Set pageSource = HTMLdoc.getElementsByName("CADMapViewer")
        Set cElem = pageSource.Item(0)
        'MsgBox cElem.src
        
        'set the source of the iframe to the google maps api with the opening arguments as the map location,
        'these arguments are passed in from the view event page when the mini map is clicked
        cElem.src = "https://www.google.com/maps/embed/v1/view?key=AIzaSyAxI0D74huu_Ecr04t1MoLjHttxrYDJrlQ&center=" & Me.OpenArgs & "&zoom=14&maptype=satellite"
        
        'MsgBox cElem.src
        'Set cElem = IHTMLIFrameElement3
        'HTMLdoc.replaceChild cElem, cElem
'        With HTMLdoc
'            Set cElem2 = .createElement("CADMapViewer")
'            With cElem2
'                'cElem.src = "https://www.google.com/maps/embed/v1/view?key=AIzaSyAxI0D74huu_Ecr04t1MoLjHttxrYDJrlQ&center=" & Me.OpenArgs & "&zoom=14&maptype=satellite"
'                .setAttribute "src", "google.com"
'                .src = "http://www.google.com"
'            End With
'            .replaceChild cElem2, cElem
'        End With
        'Me.WebBrowser0.Refresh
        
        'set the map loaded variable to show that the map has been loaded
        FrameAdded = 1
    End If
End Sub
