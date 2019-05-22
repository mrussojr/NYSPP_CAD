Attribute VB_Name = "Form_InboxManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Sub Form_Load()
    Dim webAddress As String
   
    'set the web address to the ejustice inbox alert page, using the ori of the currently logged in user and the ejustice user name of the currenly logged in user
    webAddress = "https://www.ejusticeny.ny.gov/ijapp/AlertMonitor/?userOri=" & GetORI & "&userId=" & GetEjusticeUsername & "&subGrpName="
    'navigate to the inbox alert address
    Me.WebBrowser0.Navigate webAddress
End Sub

Private Sub Form_Timer()
    'WebBrowser0.Refresh
    If Application.CurrentObjectType = acForm Then 'IsLoaded("CAD") And Screen.ActiveForm.NAME = "CAD" And Application.CurrentObjectType = acForm Then
        WebBrowser0.Silent = True
    End If
End Sub

Private Sub WebBrowser0_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error Resume Next
    If Me.WebBrowser0.ReadyState = 4 Then 'check if web browser is done loading
        'check if the cad is loaded and in normal view mode
        If CurrentProject.AllForms("CAD").IsLoaded = True And CurrentProject.AllForms("CAD").CurrentView = acCurViewFormBrowse Then
            Dim HTMLdoc As New HTMLDocument
            Dim pageSource As MSHTML.IHTMLElementCollection
            Dim cElem As MSHTML.IHTMLTable
            Dim cElem2 As MSHTML.IHTMLTableRow
            Dim cElem3 As MSHTML.IHTMLTableCell
    
            'get the html document for parsing
            Set HTMLdoc = WebBrowser0.Document
            'get the unread message boxes
            Set pageSource = HTMLdoc.getElementsByClassName("unreadMessageCounts")
    
            'get the unread priority message cell of the html table
            Set cElem = pageSource.Item(0)
            Set cElem2 = cElem.rows(1)
            Set cElem3 = cElem2.cells(0)
            
            'set the unread priority message field on the cad form to the cell value converted to an integer
            Forms!CAD!UnreadPriorityMsg.Value = CInt(cElem3.innerHTML)
            
            'get the unread message cell of the html table
            Set cElem = pageSource.Item(1)
            Set cElem2 = cElem.rows(1)
            Set cElem3 = cElem2.cells(0)
            
            'set the unread message field on the cad form to the cell value converted to an integer
            Forms!CAD!UnreadMsg.Value = CInt(cElem3.innerHTML)
        End If
    End If
End Sub

