Attribute VB_Name = "Form_Resources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
'shift sheet button
Private Sub Command16_Click()
    Dim xlApp As Excel.Application
    
    'set xlApp to an instance of the excel application
    Set xlApp = CreateObject("Excel.Application")
    'set excel to be visible
    xlApp.Visible = True
    'open the schedule workbook
    xlApp.Workbooks.Open "\\cenfile1\police$\Dispatching disk\databases\TEST AREA\2015 SHIFT WORKSHEET - Changes.xls", , 1
    'xlApp.Run "ClearSheet"
    xlApp.Run "Fill_sheet"
End Sub

Private Sub Command17_Click()
    Dim xlApp As Excel.Application
    
    'set xlApp to an instance of the excel application
    Set xlApp = CreateObject("Excel.Application")
    'set excel to be visible
    xlApp.Visible = True
    'open the schedule workbook
    xlApp.Workbooks.Open "\\cenfile1\police$\Dispatching disk\Other Resources\ERG2012.xls"
End Sub

Private Sub Command22_Click()
    FollowHyperlink "\\cenfile1\police$\CENTRAL REGION DRIVE\Dispatch Reference\EMERGENCY REFERENCE GUIDES\GUIDE COVERS\CEN.pdf"
End Sub

Private Sub Command23_Click()
    FollowHyperlink "\\cenfile1\police$\CENTRAL REGION DRIVE\Dispatch Reference\EMERGENCY REFERENCE GUIDES\GUIDE COVERS\TI.pdf"
End Sub

Private Sub Command24_Click()
    FollowHyperlink "\\cenfile1\police$\CENTRAL REGION DRIVE\Dispatch Reference\EMERGENCY REFERENCE GUIDES\GUIDE COVERS\FL.pdf"
End Sub

Private Sub Command27_Click()
    FollowHyperlink "\\cenfile1\police$\CENTRAL REGION DRIVE\Filing Cabinet\Communications\Telephone File\Current Telephone Number Listings\Regional Listing\CEN.pdf"
End Sub

Private Sub Command28_Click()
    FollowHyperlink "\\cenfile1\police$\CENTRAL REGION DRIVE\Filing Cabinet\Communications\Telephone File\Current Telephone Number Listings\Regional Listing\TI.pdf"
End Sub

Private Sub Command29_Click()
    FollowHyperlink "\\cenfile1\police$\CENTRAL REGION DRIVE\Filing Cabinet\Communications\Telephone File\Current Telephone Number Listings\Regional Listing\FL.pdf"
End Sub

Private Sub Command30_Click()
    FollowHyperlink "\\cenfile1\police$\CENTRAL REGION DRIVE\Dispatch Reference\CONTACT INFO\00-CEN-TI-FL CONTACT LIST.pdf"
End Sub

Private Sub Command31_Click()
    DoCmd.OpenForm "FacilityMaps"
End Sub

Private Sub Command32_Click()
    FollowHyperlink "\\cenfile1\police$\CENTRAL REGION DRIVE\Filing Cabinet\Communications\Telephone File\Current Telephone Number Listings\Agency Listing\2015-09 NYS Park Police Phone Directory.doc"
End Sub

Private Sub Command33_Click()
    FollowHyperlink "\\oprhp-smb\oprhp_shared\park police\Midstate\District Files\Contacts\Jefferson County - Medical Examiner Call Schedule.docx"
End Sub

Private Sub Label14_Click()
    'check if shift is down
    If IsShiftKeyDown(LeftKeyOrRightKey) Then
        'if shift is down, use chrome ;)
        Dim chromePath As String, URL As String

        chromePath = """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
        URL = """file:///P:\Dispatching disk\index.html"""
        
        Shell (chromePath & " -url " & URL)
    Else
        Dim IE As Object
        'create an ie object to navigate web pages
        Set IE = CreateObject("InternetExplorer.Application")
        
        With IE
            .Visible = True
            'open the nyspp online references on the p drive
            .Navigate "file:///P:\Dispatching disk\index.html" ', CLng(2048)
        End With
    End If
End Sub

Private Sub Label15_Click()
    'check if shift is down
    If IsShiftKeyDown(LeftKeyOrRightKey) Then
        'if shift is down, use chrome ;)
        Dim chromePath As String, URL As String

        chromePath = """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
        URL = """http://policeblotter"""
        
        Shell (chromePath & " -url " & URL)
    Else
        Dim IE As Object
        'create an ie object to navigate web pages
        Set IE = CreateObject("InternetExplorer.Application")
        
        With IE
            .Visible = True
            'open the nyspp online references on the p drive
            .Navigate "http://policeblotter" ', CLng(2048)
        End With
    End If
End Sub
