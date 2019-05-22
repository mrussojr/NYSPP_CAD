Attribute VB_Name = "Form_EJustice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private WithEvents p_Butt As MSHTML.HTMLInputButtonElement
Attribute p_Butt.VB_VarHelpID = -1
Public EventID As Long

'when the form loads set ejustice location based on log in status and forms opening arguments
Public Sub Form_Load()
    Dim webAddress As String
    'check if the user has logged into ejustice yet
    If IsLoggedIn() = True Then
        'if the user has logged into ejustice
        'check the opening arguments of the form
        If Me.OpenArgs = "INBOX" Then
            'if opening arguments = inbox, then send user to the portal inbox
            webAddress = DLookup("[EJusticeLinkId]", "users", "[username] = '" & GetUserName() & "'") & "L2dJQSEvUUt3QS80SmlFL1o2XzMwRzIwOFJHSVRGNEIwMlBCOUFTN04xRzQw/"
        Else
            'else send the user to the portal dispatch page
            'webAddress = "https://www.ejusticeny.ny.gov/wps/myportal"
            webAddress = DLookup("[EJusticeLinkId]", "users", "[username] = '" & GetUserName() & "'") & "L2dJQSEvUUt3QS80SmlFL1o2XzMwRzIwOFJHSTUwQjQwMjUxMTQ0VjMwMEk2/"
        End If
    Else
        'if the user has not logged into ejustice, take them to the login page
        webAddress = "https://www.ejusticeny.ny.gov/ijlogin/ijlogin.fcc"
    End If

    'navigate to the set web address
    Me.WebBrowser0.Navigate webAddress
    
    'set running data to false
    ClearData
End Sub

'refresh page every 30 minutes
Private Sub Form_Timer()
    WebBrowser0.Refresh
    WebBrowser0.Silent = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'check for personalized portion of ejustice address
    If Me.WebBrowser0.ReadyState = 4 And HasLinkBeenFound() = False Then
        Dim HTMLdoc As New HTMLDocument
        Dim pageSource As MSHTML.IHTMLElementCollection
        Dim cElem As MSHTML.IHTMLAnchorElement
        Dim link As String, temp As String
        
        'set htmldoc to the html of the current page
        Set HTMLdoc = WebBrowser0.Document
        'set pagesource for elements in the html with a class of child-link1, this is where the personalized portion of the ejustice address is located
        Set pageSource = HTMLdoc.getElementsByClassName("child-link1")

        If pageSource.length > 0 Then 'check to make sure there is at least one element with class child-link1
            'set celem to the second anchor element with class child-link1
            Set cElem = pageSource.Item(2)
            'grab the link from the above html anchor element
            link = cElem.href
            'find the personalized portion of the link
            temp = Left(link, (Len(link) - 1))
            temp = Left(temp, InStrRev(temp, "/"))
        
            'set sql injection warnings off
            DoCmd.SetWarnings False
            'run sql to update the personalized portion of the ejustice link to the users table for future use
            DoCmd.RunSQL "UPDATE users SET users.EJusticeLinkId = '" & temp & "' WHERE (users.Username) = '" & GetUserName() & "'"
            'reenable the sql injection warnings
            DoCmd.SetWarnings True
        
            'set a variable to show that the link has been found so this procedure doesn't need to be run every time
            FoundLink
        End If
    End If
End Sub

Private Sub WebBrowser0_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    Dim cF As clsFrwd
    Dim HTMLdoc As New HTMLDocument
    Dim formSource As MSHTML.IHTMLFormElement
    'Dim cElem As MSHTML.IHTMLInputButtonElement
    Dim pageSource As MSHTML.IHTMLElementCollection
    
    Set HTMLdoc = WebBrowser0.Document
    
    If HTMLdoc.Title = "Navigation Canceled" Or HTMLdoc.Title = "HTTP 404 Not Found" Then
        DoCmd.SetWarnings False
        DoCmd.RunSQL "UPDATE users SET LoggedInEJustice = 0 WHERE username = '" & GetUserName & "'"
        DoCmd.Close acForm, "EJustice"
        DoCmd.OpenForm "EJustice"
        DoCmd.SetWarnings True
        Exit Sub
    End If
    
    Set cF = New clsFrwd

    Set formSource = HTMLdoc.getElementById("viewns_Z7_5NO69B1A0GKL30IDSVA5OL30C4_:navForm")
    
    If Not IsNull(formSource) And Not formSource Is Nothing Then
        If InStr(1, formSource.innerHTML, "populateButton") < 1 Then
            formSource.innerHTML = formSource.innerHTML & "<input type='button' id='populateButton' name='populateButton' class='dispatchActionButton' value='populate' >"
    
            Set pageSource = HTMLdoc.getElementsByName("populateButton")
            'Set cElem = pageSource.Item(0)
            Set p_Butt = pageSource.Item(0)
            
            'cF.Set_Destination Me, "populateTable"
            'cElem.OnClick = cF.forwardEvent
        End If
    End If

End Sub

Private Function p_Butt_onclick() As Boolean
    DoCmd.OpenForm "ActiveEventsForEJustice"
    'populateTable
End Function

Private Sub WebBrowser0_FileDownload(ByVal ActiveDocument As Boolean, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub WebBrowser0_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
    Dim HTMLdoc As New HTMLDocument
    Dim pageSource As MSHTML.IHTMLElementCollection
    Dim cElem As MSHTML.IHTMLElement
    
    Dim buttonPresource As String
    
    buttonPresource = "viewns_Z7_5NO69B1A0GKL30IDSVA5OL30C4_"
    
    If Me.WebBrowser0.ReadyState = 4 Then
        If IsLoggedIn = False Then 'check whether the user has logged into ejustice yet
            'grab html document for sorting through
            Set HTMLdoc = WebBrowser0.Document
            'find elements with a name of user
            Set pageSource = HTMLdoc.getElementsByName("USER")
    
            'grab the first element with the name of user
            Set cElem = pageSource.Item(0)
    
            'look up the ejustice username for the currently logged in user and set it as the value of the user element
            cElem.Value = GetEjusticeUsername()
    
            'set a variable to show that the user has logged in
            UpdateLogIn
        Else
            If Not IsNull(Me.OpenArgs) And HasDataBeenRun() = False Then 'make sure a data has not been run yet and check that the user is running a registration from the enter event page
                'perform a dmv search based on the opening arguments
                eSearch Me.OpenArgs
                'set a variable indicating that a data has been run so this procedure doesn't continue to run and request the same data over and over
                RunData
            ElseIf HasDataBeenRun() Then
                
            End If
        End If
    End If
End Sub

Public Function populateTable()
    Dim HTMLdoc As New HTMLDocument
    Dim pageSource As MSHTML.IHTMLElementCollection
    Dim cElem As MSHTML.IHTMLElement
    
    Set HTMLdoc = Me.WebBrowser0.Document
    
    'set for ny cid
    Set cElem = HTMLdoc.getElementById("_id30_recordTable_0")
    If Not cElem Is Nothing Then
        'MsgBox cElem.innerHTML
        storeNYCID cElem.Children(0)
        MsgBox "Import Complete"
        Exit Function
    End If
    
    'set for ny dob
    Set cElem = HTMLdoc.getElementById("_id31_recordTable_0")
    If Not cElem Is Nothing Then
        'MsgBox cElem.innerHTML
        storeNYCID cElem.Children(0)
        MsgBox "Import Complete"
        Exit Function
    End If
    
    'set for ny plate or ny vin
    Set cElem = HTMLdoc.getElementById("dmvRegistrationResponse_recordTable_0")
    If Not cElem Is Nothing Then
        'MsgBox "storing ny plate"
        storeNYPlate cElem.Children(0)
        MsgBox "Import Complete"
        Exit Function
    End If
    
    'set for out of state data
    Set pageSource = HTMLdoc.getElementsByClassName("summaryDetail")
    If Not pageSource Is Nothing Then
        Set cElem = pageSource.Item(0)
        If Not cElem Is Nothing Then
            'MsgBox cElem.innerHTML
            MsgBox "Import currently only available for NYS data :("
            Exit Function
        End If
    End If
    
    MsgBox "Unable to capture data"
End Function

Public Function storeNYPlate(table As MSHTML.IHTMLTable)
    Dim oRow As MSHTML.IHTMLTableRow
    Dim oCell As MSHTML.IHTMLTableCell
    Dim x As Integer, y As Integer, z As Integer, vehType As Integer, VehYear As Integer
    Dim chk As String, vehPlate As String, vehStatus As String, vehExp As String, VehMake As String, vehModel As String, vehStyle As String, vehColor As String, vehVIN As String, tmp As String, sql As String, sql2 As String, altrow As String, altrow2 As String, altrow3 As String
    Dim vID As Long, assocId As Long
    
    'if data is for a boat, increase the rows by 1
    If InStr(1, table.innerHTML, "BOT - Boat") > 0 Then
        altrow = "1/1"
        altrow2 = "4/1"
        altrow3 = "5/1"
    Else
        altrow = "1/1"
        altrow2 = "3/1"
        altrow3 = "5/1"
    End If
    
    x = 0

    Do While x < table.rows.length
        y = 0
        Set oRow = table.rows.Item(x)
        Do While y < oRow.cells.length
            Set oCell = oRow.cells.Item(y)
            
            chk = x & "/" & y
            
            Select Case chk
                Case "0/1"
                    vehStatus = Left(oCell.innerHTML, InStr(1, oCell.innerHTML, "&nbsp;") - 1)
                    vehExp = Mid(oCell.innerHTML, InStr(1, oCell.innerHTML, "Expiration:") + 25, 10)
                    If vehExp = "<SPAN clas" Then
                        vehExp = Mid(oCell.innerHTML, InStr(1, oCell.innerHTML, "Expiration:") + 47, 10)
                    End If
                Case altrow
                    vehPlate = Left(oCell.innerHTML, InStr(1, oCell.innerHTML, "&nbsp;") - 1)
                    tmp = oCell.innerHTML
                    tmp = Mid(tmp, InStr(1, tmp, "(") + 1, InStr(1, tmp, ")") - InStr(1, tmp, "(") - 1)
                    vehType = CInt(tmp)
                Case altrow2
                    If vehType = 86 Or vehType = 36 Or vehType = 84 Or vehType = 90 Or vehType = 87 Then
                        tmp = oCell.innerHTML
                        VehYear = Left(tmp, InStr(1, tmp, ";") - 1)
                        tmp = Replace(tmp, VehYear, "", 1, 1)
                        tmp = Trim(Right(tmp, Len(tmp) - 1))
                        VehMake = Left(tmp, InStr(1, tmp, ";") - 1)
                        tmp = Replace(tmp, VehMake, "")
                        tmp = Trim(Right(tmp, Len(tmp) - 1))
                        vehModel = ""
'                        vehModel = Left(tmp, InStr(1, tmp, ";") - 1)
'                        tmp = Replace(tmp, vehModel, "")
'                        tmp = Trim(Right(tmp, Len(tmp) - 1))
                        If vehType = 90 Then
                            vehStyle = tmp
                            vehColor = ""
                        Else
                            vehStyle = Left(tmp, InStr(1, tmp, ";") - 1)
                            tmp = Replace(tmp, vehStyle, "")
                            vehColor = Trim(Right(tmp, Len(tmp) - 1))
                        End If
                    Else
                        tmp = oCell.innerHTML
                        VehYear = Left(tmp, InStr(1, tmp, ";") - 1)
                        tmp = Replace(tmp, VehYear, "", 1, 1)
                        tmp = Trim(Right(tmp, Len(tmp) - 1))
                        VehMake = Left(tmp, InStr(1, tmp, ";") - 1)
                        tmp = Replace(tmp, VehMake, "")
                        tmp = Trim(Right(tmp, Len(tmp) - 1))
                        vehModel = Left(tmp, InStr(1, tmp, ";") - 1)
                        tmp = Replace(tmp, vehModel, "")
                        tmp = Trim(Right(tmp, Len(tmp) - 1))
                        vehStyle = Left(tmp, InStr(1, tmp, ";") - 1)
                        tmp = Replace(tmp, vehStyle, "")
                        vehColor = Trim(Right(tmp, Len(tmp) - 1))
                    End If
                Case altrow3
                    vehVIN = oCell.innerHTML
            End Select
            
            y = y + 1
        Loop
        x = x + 1
    Loop
    
    vID = Nz(DLookup("[ID]", "vehicles", "PlateNumber = '" & vehPlate & "'"), 0)
    
    DoCmd.SetWarnings False
    
    If vID = 0 Then
        sql = "INSERT INTO vehicles (State, PlateNumber, Type, Expiration, VehYear, VehMake, VIN, Status, VehStyle, VehModel, VehColor) VALUES "
        sql = sql & "('NY', '" & vehPlate & "', " & vehType & ", '" & vehExp & "', " & VehYear & ", '" & VehMake & "', '" & vehVIN & "', '" & vehStatus & "', '" & vehStyle & "', '" & vehModel & "', '" & vehColor & "')"
        
        CurrentDb.Execute sql
        
        sql2 = "SELECT @@Identity"
        vID = CurrentDb.OpenRecordset(sql2)(0)
        
        sql = "INSERT INTO EventVehicleAssoc (EventId, VehId) VALUES (" & Me.EventID & "," & vID & ")"
        DoCmd.RunSQL sql
    Else
        sql = "UPDATE vehicles SET PlateNumber='" & vehPlate & "', Type=" & vehType & ", Expiration='" & vehExp & "', VehYear=" & VehYear & ", VehMake='" & VehMake & "', VIN='" & vehVIN & "', Status='" & vehStatus & "', VehStyle='" & vehStyle & "', VehModel='" & vehModel & "', VehColor='" & vehColor & "' WHERE ID = " & vID
        DoCmd.RunSQL sql
        
        assocId = DCount("[ID]", "EventVehicleAssoc", "[EventId] = " & Me.EventID & " AND [VehId] = " & vID)
        
        If assocId = 0 Then
            sql = "INSERT INTO EventVehicleAssoc (EventId, VehId) VALUES (" & Me.EventID & "," & vID & ")"
            DoCmd.RunSQL sql
        End If
    End If
    
    DoCmd.SetWarnings True
End Function

Public Function storeNYCID(table As MSHTML.IHTMLTable)
    Dim oRow As MSHTML.IHTMLTableRow
    Dim oCell As MSHTML.IHTMLTableCell
    Dim x As Integer, y As Integer, zipCode As Integer
    Dim chk As String, fullName As String, lName As String, fName As String, mName As String, dateOfBirth As String, gender As String, CID As String, licClass As String, licExp As String, eyeColor As String, streetAddress As String, County As String, muni As String, tmp As String, sql As String, sql2 As String, altrow As String, altrow2 As String, altrow3
    Dim pId As Long, assocId As Long
    
    'if data has a permit then increase the rows by 1
    If InStr(1, table.innerHTML, "Permit") > 0 Then
        altrow = "3/1"
        altrow2 = "4/1"
        altrow3 = "5/1"
    Else
        altrow = "2/1"
        altrow2 = "3/1"
        altrow3 = "4/1"
    End If
    
    x = 0

    Do While x < table.rows.length
        y = 0
        Set oRow = table.rows.Item(x)
        Do While y < oRow.cells.length
            Set oCell = oRow.cells.Item(y)
            
            chk = x & "/" & y
            
            tmp = oCell.innerHTML
            
            Select Case chk
                Case "1/1"
                    licClass = Replace(Left(tmp, InStrRev(tmp, "*")), "*", "")
                    licExp = Mid(oCell.innerHTML, InStr(1, oCell.innerHTML, "Expiration:") + 25, 10)
                Case altrow
                    fullName = Left(tmp, InStr(1, tmp, "&nbsp;") - 1)
                    tmp = Replace(tmp, fullName, "")
                    lName = Left(fullName, InStr(1, fullName, ",") - 1)
                    fullName = Trim(Right(Replace(fullName, lName, ""), Len(Replace(fullName, lName, "")) - 1))
                    If InStr(1, fullName, " ") > 0 Then
                        fName = Left(fullName, InStr(1, fullName, " ") - 1)
                        mName = Nz(Trim(Replace(fullName, fName, "")), "")
                    Else
                        fName = Trim(fullName)
                        mName = ""
                    End If
                    CID = Right(tmp, Len(tmp) - InStrRev(tmp, ";"))
                Case altrow2
                    dateOfBirth = Left(tmp, InStr(1, tmp, "</SPAN>") - 1)
                    dateOfBirth = Right(dateOfBirth, 10)
                    gender = Mid(tmp, InStr(1, tmp, "Sex Code"), InStrRev(tmp, "</SPAN>") - InStr(1, tmp, "Sex Code"))
                    gender = Right(gender, Len(gender) - InStrRev(gender, ">"))
                    tmp = Right(tmp, Len(tmp) - InStr(1, tmp, "Eye Color:"))
                    eyeColor = Right(tmp, Len(tmp) - InStr(1, tmp, "&nbsp;") + 1)
                    eyeColor = Replace(eyeColor, "&nbsp;", "")
                Case altrow3
                    streetAddress = Mid(tmp, InStr(1, tmp, ">") + 1, InStr(1, tmp, "<BR>") - InStr(1, tmp, ">") - 1)
                    tmp = Right(tmp, Len(tmp) - InStr(1, tmp, "<BR>") - 3)
                    muni = Left(tmp, InStr(1, tmp, ",") - 1)
                    zipCode = CInt(Mid(tmp, InStr(1, tmp, "</DIV>") - 5, 5))
                    County = Right(tmp, Len(tmp) - InStrRev(tmp, "&nbsp;") - 5)
                    County = Left(County, InStr(1, County, "<") - 1)
            End Select
            
            y = y + 1
        Loop
        
        x = x + 1
    Loop
    
    pId = Nz(DLookup("[ID]", "people", "[LastName] = '" & lName & "' And [FirstName] = '" & fName & "'"), 0)
    
    If pId = 0 Then
        pId = Nz(DLookup("[ID]", "people", "[CID] = '" & CID & "'"), 0)
    End If
    
    DoCmd.SetWarnings False
    
    If pId = 0 Then
        sql = "INSERT INTO people (LastName, FirstName, Middle, DOB, Sex, State, CID, Class, Expiration, EyeColor, StreetAddress, County, Municipality, ZipCode) VALUES "
        sql = sql & "('" & lName & "', '" & fName & "', '" & mName & "', '" & dateOfBirth & "', '" & gender & "', 'NY', '" & CID & "', '" & licClass & "', '" & licExp & "', '" & eyeColor & "', '" & streetAddress & "', '" & County & "', '" & muni & "', " & zipCode & ")"
        
        CurrentDb.Execute sql
        
        sql2 = "Select @@Identity"
        pId = CurrentDb.OpenRecordset(sql2)(0)
        
        sql = "INSERT INTO EventPersonAssoc (EventId, PersonId) VALUES (" & Me.EventID & ", " & pId & ")"
        DoCmd.RunSQL sql
    Else
        sql = "UPDATE people SET LastName='" & lName & "', FirstName='" & fName & "', Middle='" & mName & "', DOB='" & dateOfBirth & "', Sex='" & gender & "', State='NY', CID='" & CID & "', Class='" & licClass & "', Expiration='" & licExp & "', EyeColor='" & eyeColor & "', StreetAddress='" & streetAddress & "', County='" & County & "', Municipality='" & muni & "', ZipCode=" & zipCode & " WHERE ID=" & pId
        DoCmd.RunSQL sql
        
        assocId = DCount("[ID]", "EventPersonAssoc", "[EventId] = " & Me.EventID & " AND [PersonId] = " & pId)
        
        If assocId = 0 Then
            sql = "INSERT INTO EventPersonAssoc (EventId, PersonId) VALUES (" & Me.EventID & "," & pId & ")"
            DoCmd.RunSQL sql
        End If
    End If
    
    DoCmd.SetWarnings True
End Function

Public Function eSearch(formName As String)
    Dim dbs As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlStr As String, buttonPresource As String
    Dim HTMLdoc As New HTMLDocument
    Dim pageSource As MSHTML.IHTMLElementCollection
    Dim cElem As MSHTML.IHTMLElement
    Dim frm As Form
    Dim ctl As control
    
    buttonPresource = "viewns_Z7_5NO69B1A0GKL30IDSVA5OL30C4_"
    
    Set dbs = CurrentDb
    
    sqlStr = "SELECT * FROM EjusticeElements WHERE FormID = '" & formName & "' ORDER BY OrderId ASC"
    Set rs = dbs.OpenRecordset(sqlStr, dbOpenSnapshot)
    
    Set HTMLdoc = Me.WebBrowser0.Document
    
    Do While Not rs.EOF
        Set pageSource = HTMLdoc.getElementsByName(buttonPresource & rs.Fields("ElementName"))
        Set cElem = pageSource.Item(0)
        
        If rs.Fields("ElementType") = "Button" Then
            cElem.FireEvent "OnClick"
        ElseIf rs.Fields("ElementType") = "Text" Then
            Set frm = Forms.Item(rs.Fields("ValueFormName"))
            Set ctl = frm.Controls.Item(rs.Fields("ValueFieldName"))
            cElem.Value = ctl
        ElseIf rs.Fields("ElementType") = "Check" Then
            Dim chkBox As MSHTML.IHTMLInputElement
            Set chkBox = cElem
            chkBox.Checked = True
        End If
        
        Set pageSource = Nothing
        Set cElem = Nothing
        
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
End Function
