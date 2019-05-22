Attribute VB_Name = "Form_Events"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim loader As Integer
Dim Fill As Boolean
Dim FillStr As String

'on form close, procedure to capture parks event id to be saved for future use, this id is necessary for
'any of the fill functions from the view event form (ability to fill in parks from the cad)
Private Sub Form_Close()
    'check to see if this instance of the events form should be used to capture the parks id
    If InStr(1, Me.OpenArgs, "Fill") > 0 Then
        Dim webAddress As String, tmp As String, tmp2 As String
        Dim evtNumber As Long, evtId As Long
        
        'set web address variable to the current url of the webpage
        webAddress = Me.WebBrowser0.LocationURL
        
        'check if there is a current page url
        If InStr(1, webAddress, "slctEventId=") > 0 Then
            'find the location of "slctEventId=" in the current url and return everything from there after it to a temp variable
            tmp = Right(webAddress, Len(webAddress) - InStr(1, webAddress, "slctEventId=") + 1)
            'find the location of "&" in the remaining url and return only to that character
            tmp = Left(tmp, InStr(1, tmp, "&") - 1)
            'finally replace "slctEventId=" with an empty string, should leave only the parks event id that is to be captured
            tmp2 = Replace(tmp, "slctEventId=", "")
            
            'check if there is an event id or not
            If Not IsNull(tmp2) And Not tmp2 = "" Then
                'if there is an event id, convert it to a long integer
                evtNumber = CLng(tmp2)
                'pass the parks event id back to the view event form to be saved into the events table for future use
                Forms.Item("ViewEvent").Text38.Value = evtNumber
                
                DoCmd.SetWarnings False
                CurrentDb.Execute "UPDATE events SET pId=" & evtNumber & " WHERE ID=" & Forms.Item("ViewEvent").Text25.Value
                DoCmd.SetWarnings True
            End If
        End If
    End If
End Sub

Public Sub Form_Load()
    Dim webAddress As String

    loader = 1
    'set the fill variable to false, indicating that this instance of the event page is just for viewing
    Fill = False
    
    'check the forms open arguments
    If InStr(1, Me.OpenArgs, "Fill") > 0 Then
        'if the forms open arguments contain the word fill
        'set the web address to the open arguments minus the word fill
        webAddress = Left(Me.OpenArgs, Len(Me.OpenArgs) - 4)
        'set the fill variable to true, this allows the auto fill procedures to be called
        Fill = True
        'set the fill string variable to fill, this will tell which fill procedure to call
        FillStr = "Fill"
    ElseIf InStr(1, Me.OpenArgs, "PEEP") > 0 Then
        'if contains peep
        'set the web address to the opening arguments minus word peep
        webAddress = Left(Me.OpenArgs, Len(Me.OpenArgs) - 4)
        'set the fill variable to fill, this allows the auto fill procedures to be called
        Fill = True
        'set the fill string variable to peep, this will tell which fill procedure to call
        FillStr = "PEEP"
    ElseIf InStr(1, Me.OpenArgs, "VEHI") > 0 Then
        'if contains vehi
        'set the web address to the opening arguments minus word vehi
        webAddress = Left(Me.OpenArgs, Len(Me.OpenArgs) - 4)
        'set the fill variable to fill, allows the auto fill procedures to be called
        Fill = True
        'set the fill string variable to vehi, tells which auto fill procedure to call
        FillStr = "VEHI"
    ElseIf InStr(1, Me.OpenArgs, "CALL") > 0 Then
        'if contains call
        'set the web address to the opening arguments minus word call
        webAddress = Left(Me.OpenArgs, Len(Me.OpenArgs) - 4)
        'set the fill variable to fill, allows the auto fill procedures to be called
        Fill = True
        'set the fill string variable to cell, tells which auto fill procedure to call
        FillStr = "CALL"
    Else
        'if no fill keywords are found, set web address to the opening arguments
        webAddress = Me.OpenArgs
    End If

    'navigate the web browser to the web address
    Me.WebBrowser0.Navigate webAddress
End Sub

Public Sub WebBrowser0_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
    Dim HTMLdoc As New HTMLDocument
    Dim pageSource As MSHTML.IHTMLElementCollection
    Dim cElem As MSHTML.IHTMLElement
    
    'check if the web browser is loaded, if it should filled, and which fill procedure to use
    If Me.WebBrowser0.ReadyState = 4 And Fill = True And FillStr = "Fill" Then
        'if the fill string is fill
        'grab the html document for parsing
        Set HTMLdoc = Me.WebBrowser0.Document
        
        Dim dateFields() As String
        Dim timeFields() As String
        Dim i As Integer
        
        'splits the string of date fields by the pipeline character, creating an array of date fields for looping through
        dateFields() = Split("ctl00$Container$uctrlEvent$txtEventDate|ctl00$Container$uctrlEvent$txtAssignedDate|ctl00$Container$uctrlEvent$txtOccursStartDate", "|", -1, vbBinaryCompare)
        'splits the string of time fields by the pipeline character, creating an array of time fields for looping through
        timeFields() = Split("ctl00$Container$uctrlEvent$txtCallTime|ctl00$Container$uctrlEvent$txtArrivedTime|ctl00$Container$uctrlEvent$txtOccursStartTime", "|", -1, vbBinaryCompare)
        
        'loop through the date fields
        For i = LBound(dateFields()) To UBound(dateFields())
            'get the current date field html element and fill it with the current date value from the view event page
            Set pageSource = HTMLdoc.getElementsByName(dateFields(i))
            Set cElem = pageSource.Item(0)
            cElem.Value = Forms!ViewEvent.Text0.Value
        Next i
        
        'loop through the time fields
        For i = LBound(timeFields()) To UBound(timeFields())
            'get the current time field html element and fill it with the current time value from the view event page
            Set pageSource = HTMLdoc.getElementsByName(timeFields(i))
            Set cElem = pageSource.Item(0)
            cElem.Value = Replace(ShortTime(Forms!ViewEvent.Text2.Value), ":", "")
        Next i
        
        'get the time assigned field in parks and fill it with the time right now in short military time
        Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlEvent$txtAssignedTime")
        Set cElem = pageSource.Item(0)
        cElem.Value = Replace(ShortTime(RightNow()), ":", "")
        
        'check to see that the incident type on the view event page is filled in
        If IsNull(Forms!ViewEvent.Text19.Value) = False And Not Forms!ViewEvent.Text19.Value = "" Then
            Dim cat As String
            'set the category variable to the incident type from the view event page
            cat = Forms!ViewEvent.Text19.Value
            'count the event types in the event types table that match the category variable, by receiving a count greater than 0 it shows that category to exist
            If DCount("[ID]", "EvtTypes", "[Event_Type] = '" & cat & "'") > 0 Then
                Dim catParkId As Integer
                'look up the category id number from the event types table
                catParkId = DLookup("[ParkNum]", "EvtTypes", "[Event_Type] = '" & cat & "'")
                'get the category drop down from parks and fill it
                Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlEvent$lstCategory")
                Set cElem = pageSource.Item(0)
                cElem.Value = catParkId
                'fire the onchange event of the category drop down to simulate the actual interactions that would normally occur
                cElem.FireEvent "OnChange"
                'pause the code from running for 2 seconds, this allows the javascript in the document to run, necessary for populating following fields
                Pause 2
            End If
        End If
        
        'pause 2 seconds
        Pause 2
        
        'get the district drop down box, fill it, fire it's onchange event, and pause 2 seconds to allow the javascript on the page to run
        Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlEvent$lstDistrict")
        Set cElem = pageSource.Item(0)
        cElem.Value = 3687
        cElem.FireEvent "OnChange"
        Pause 2
        
        'get the zone drop down box, look up the zone id by the officer id from the view event form,
        'select the option for the zone, fire it's onchange event, and pause 2 seconds to allow counties to be loaded into county drop down box
        Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlEvent$lstZone")
        Set cElem = pageSource.Item(0)
        Dim zone As Integer
        zone = DLookup("[pZone]", "parksIDs", "[pId] = " & Forms!ViewEvent.Text36.Value)
        cElem.Value = zone
        cElem.FireEvent "OnChange"
        Pause 2
        
        'get the county drop down box, look up the county of the location selected, select the count, fire onchange event, pause 2 seconds
        Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlEvent$lstCounty")
        Set cElem = pageSource.Item(0)
        Dim CountyName As String
        Dim County As Integer
        CountyName = DLookup("[County]", "AllLocations", "[Park] ='" & Forms!ViewEvent.Text10.Value & "'")
        County = DLookup("[pID]", "AllCounties", "[County] = '" & CountyName & "'")
        'County = DLookup("[pCounty]", "parksIDs", "[pId] = " & Forms!ViewEvent.Text36.Value)
        cElem.Value = County
        cElem.FireEvent "OnChange"
        Pause 2
        
        'get the station drop down, look up station the location is located in, select the station, fire onchange event, and pause 2 seconds
        Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlEvent$lstStation")
        Set cElem = pageSource.Item(0)
        Dim stationName As String
        Dim station As Integer
        stationName = CStr(DLookup("[Station]", "AllLocations", "[Park] ='" & Forms!ViewEvent.Text10.Value & "'"))
        station = DLookup("[StationPID]", "Stations", "[Station] = '" & stationName & "'")
        'station = DLookup("[pStation]", "parksIDs", "[pId] = " & Forms!ViewEvent.Text36.Value)
        cElem.Value = station
        cElem.FireEvent "OnChange"
        Pause 2
        
        
        Dim loca As String
        'set the park or municipality to the location variable
        loca = Forms!ViewEvent.Text10.Value
        'check to make sure that a park or municipality was actually selected
        If IsNull(loca) = False And Not loca = "" Then
            Dim Addy As String
            'check to see whether the chosen location is a park or a municipality
            If DCount("[ID]", "AllParks", "[Park] = '" & loca & "'") > 0 Then
                'if location is a park
                'get the park drop down box, look up park id, select park, fire onchange event, pause 2 seconds
                Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlEvent$lstPark")
                Set cElem = pageSource.Item(0)
                Dim Park As Integer
                Park = DLookup("[ParksID]", "AllParks", "[Park] = '" & loca & "'")
                cElem.Value = Park
                cElem.FireEvent "OnChange"
                Pause 2
                
                'get the municipality drop down box, select the first item (page javascript loads only one choice), fire the onchange event, and pause
                Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlEvent$lstTown")
                Set cElem = pageSource.Item(0)
                cElem.selectedIndex = 1
                cElem.FireEvent "OnChange"
                Pause 2
                
                'get the address box
                Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlEvent$txtAddress")
                Set cElem = pageSource.Item(0)
                Dim reg As String
                'look up the zone for the selected park
                reg = DLookup("[Region]", "AllParks", "[Park] = '" & loca & "'")
                Dim table As String
                
                'get the table to look up the address based on the zone
                Select Case reg
                    Case "Finger Lakes":
                        table = "PANIC_BOOK_-_FL"
                    Case "Thousand Islands":
                        table = "PANIC_BOOK_-_TI"
                    Case "Central":
                        table = "PANIC_BOOK_-_CEN"
                    Case Else
                        table = "PANIC_BOOK_-_CEN"
                End Select
                
                'look up address of park, set address box value to park address
                Addy = DLookup("[Copy_of_Address]", table, "[Park] = '" & loca & "'")
                cElem.Value = Addy
                
                'get the business text box, fill it with the park
                Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlEvent$txtBusiness")
                Set cElem = pageSource.Item(0)
                cElem.Value = loca
            Else
                'if the location is a municipality
                'get the park drop down box, select the out of park option, fire the onchange event, and pause
                Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlEvent$lstPark")
                Set cElem = pageSource.Item(0)
                cElem.Value = 1471
                cElem.FireEvent "OnChange"
                Pause 2
                
                'get the municipality drop down box, look up the town id, select the town, fire the onchange event, and pause
                Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlEvent$lstTown")
                Set cElem = pageSource.Item(0)
                Dim town As Integer
                town = Nz(DLookup("[ParksID]", "locations", "[Location] = '" & loca & "'"), 0)
                cElem.Value = town
                cElem.FireEvent "OnChange"
                Pause 2
                
                'get the address text box
                Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlEvent$txtAddress")
                Set cElem = pageSource.Item(0)
                Dim str1 As String, str2 As String
                'set the string variables to address and cross street from the view event page
                str1 = Nz(Forms!ViewEvent.Combo12.Value, "")
                str2 = Nz(Forms!ViewEvent.Text14.Value, "")
                
                'check if the address is blank
                If Not str1 = "" Then
                    'if not blank, check if the cross street is blank
                    If Not str2 = "" Then
                        'if not blank, set the address variable to address and cross streed
                        Addy = str1 & " & " & str2
                    Else
                        'if the cross street is blank, set the address variable to address
                        Addy = str1
                    End If
                Else
                    'if the address is blank, set the address variable to an empty string
                    Addy = ""
                End If
                
                'set the value of the address text box to the address variable
                cElem.Value = Addy
            End If
        End If
        
        'get the status drop down box, select the pending classification status (all events are to start with this status), fire the onchange event, and pause
        Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlEvent$lstStatus")
        Set cElem = pageSource.Item(0)
        cElem.Value = 1418
        cElem.FireEvent "OnChange"
        Pause 2
        
        'get the narrative text box and fill it with the narrative from the view event page
        Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlEvent$txtNarrative")
        Set cElem = pageSource.Item(0)
        cElem.Value = Forms!ViewEvent.Text17.Value
        
        
'        Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$hdnEventId")
'        Set cElem = pageSource.Item(0)
'        If IsNull(cElem.Value) = False And Not cElem.Value = "" Then
'            Dim sql As String
'            sql = "UPDATE events SET pId=" & cElem.Value & " WHERE ID = " & Forms!ViewEvent.Text25.Value
'            DoCmd.SetWarnings False
'            DoCmd.RunSQL sql
'            DoCmd.SetWarnings True
'        End If

        'set fill variable to false so it doesn't continue to try and fill the parks page
        Fill = False
    ElseIf Me.WebBrowser0.ReadyState = 4 And Fill = True And FillStr = "PEEP" Then
        'if the fill string is peep
        'navigate to the people tab using the page javascript and pause to allow time for it to load
        Me.WebBrowser0.Navigate "javascript:__doPostBack('ctl00$Container$tab2','')"
        Pause 2
        
        'grab the html document for parsing
        Set HTMLdoc = Me.WebBrowser0.Document
        
        'get the last name text box, look up the last name in the person table using the person id field of the event form,
        'convert the last name to upper case, and enter it into the last name text box
        Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlPeople$txtLastName")
        Set cElem = pageSource.Item(0)
        cElem.Value = UCase(DLookup("[LastName]", "EvtPerson", "[ID] = " & Forms!ViewEvent.PersonID.Value))
        
        'same as last but with first name
        Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlPeople$txtFirstName")
        Set cElem = pageSource.Item(0)
        cElem.Value = UCase(DLookup("[FirstName]", "EvtPerson", "[ID] = " & Forms!ViewEvent.PersonID.Value))
        
        'same as last but with middle initial
        Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlPeople$txtMiddleInt")
        Set cElem = pageSource.Item(0)
        cElem.Value = Nz(UCase(DLookup("[Middle]", "EvtPerson", "[ID] = " & Forms!ViewEvent.PersonID.Value)), "")
        
        'get the address text box, look up street address and fill
        Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlPeople$txtAddress")
        Set cElem = pageSource.Item(0)
        cElem.Value = Nz(UCase(DLookup("[StreetAddress]", "EvtPerson", "[ID] = " & Forms!ViewEvent.PersonID.Value)), "")

        'get the city text box, look up the city, and fill
        Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlPeople$txtCity")
        Set cElem = pageSource.Item(0)
        cElem.Value = Nz(UCase(DLookup("[Municipality]", "EvtPerson", "[ID] = " & Forms!ViewEvent.PersonID.Value)), "")
        
        'get the state drop down, look up state, and select option for that state
        Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlPeople$lstState")
        Set cElem = pageSource.Item(0)
        cElem.Value = Nz(UCase(DLookup("[Expr4]", "EvtPerson", "[ID] = " & Forms!ViewEvent.PersonID.Value)), "")

        'get the zipcode text box, look up the persons zipcode and fill the box
        Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlPeople$txtZip")
        Set cElem = pageSource.Item(0)
        cElem.Value = Nz(DLookup("[ZipCode]", "EvtPerson", "[ID] = " & Forms!ViewEvent.PersonID.Value), 0)
    
        'get the gender drop down, look up gender, and select option in drop down for that gender
        Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlPeople$lstGender")
        Set cElem = pageSource.Item(0)
        cElem.Value = Nz(UCase(DLookup("[Expr2]", "EvtPerson", "[ID] = " & Forms!ViewEvent.PersonID.Value)), "")
    
        'get the eye color drop down, look up the persons eye color, and select the option for that list
        Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlPeople$lstEye")
        Set cElem = pageSource.Item(0)
        cElem.Value = Nz(DLookup("[Expr3]", "EvtPerson", "[ID] = " & Forms!ViewEvent.PersonID.Value), 3344)
        
        'get the comments box, look up the cid and lic exp, and fill in the textarea
        Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlPeople$uctrlCommentsModule$txtComments")
        Set cElem = pageSource.Item(0)
        cElem.Value = "CID: " & Nz(UCase(DLookup("[CID]", "EvtPerson", "[ID] = " & Forms!ViewEvent.PersonID.Value)), "") & vbNewLine & "Exp: " & Nz(DLookup("[Expiration]", "EvtPerson", "[ID] = " & Forms!ViewEvent.PersonID.Value), "")
        
        'same as last but with date of birth, fire onchange event, and pause
        Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlPeople$txtDateOfBirth")
        Set cElem = pageSource.Item(0)
        cElem.Value = Nz(UCase(DLookup("[DOB]", "EvtPerson", "[ID] = " & Forms!ViewEvent.PersonID.Value)), "")
        cElem.FireEvent "OnChange"
        'Pause 2
        
    ElseIf Me.WebBrowser0.ReadyState = 4 And Fill = True And FillStr = "VEHI" Then
        'if the fill string is vehi
        'navigate to the property tab using the page javascript and pause to allow page to load
        Me.WebBrowser0.Navigate "javascript:__doPostBack('ctl00$Container$tab6','')"
        Pause 2
        
        'navigate to the vehicle property tab using the page javascript and pause to allow page to load
        Me.WebBrowser0.Navigate "javascript:__doPostBack('ctl00$Container$uctrlProperty$lnkProperty2','')"
        Pause 2
        
        'grab html document for parsing
        Set HTMLdoc = Me.WebBrowser0.Document
        
        'get the drop down for vehicle status and select option for observed (most commonly used status)
        Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlProperty$lstVehicleStatus")
        Set cElem = pageSource.Item(0)
        cElem.Value = 1384
        
        'get the text box for the vehicle plate, look up plate using id on view event form, and fill in plate
        Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlProperty$txtVehiclePlate")
        Set cElem = pageSource.Item(0)
        cElem.Value = Nz(UCase(DLookup("[PlateNumber]", "EvtVehicle", "[ID] = " & Forms!ViewEvent.VehicleID.Value)), "")
        
        'get full plate check box, set to checked, fire on click event, and pause
        Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlProperty$chkFullPlate")
        Set cElem = pageSource.Item(0)
        cElem.Checked = True
        cElem.FireEvent "OnClick"
        Pause 2
        
        Dim ste As String
        'set state variable by looking up the state of the vehicle by id
        ste = Nz(UCase(DLookup("[State]", "EvtVehicle", "[ID] = " & Forms!ViewEvent.VehicleID.Value)), "")
        'check value of state variable
        If Not ste = "NY" And Not ste = "" Then
            'if state variable not ny or blank (not ny because default in parks is ny)
            'get state drop down and select state value
            Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlProperty$lstVehicleState")
            Set cElem = pageSource.Item(0)
            cElem.Value = ste
        End If
        
        'get plate expiration text box, look up expiration from vehicle table, set the value, and fire the onkeypress event
        Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlProperty$txtExpireYear")
        Set cElem = pageSource.Item(0)
        cElem.Value = Nz(DLookup("[Expiration]", "EvtVehicle", "[ID] = " & Forms!ViewEvent.VehicleID.Value), "")
        If Len(cElem.Value) > 4 Then
            cElem.Value = Right(cElem.Value, 4)
        End If
        cElem.FireEvent "OnKeyPress"
        
        'get vehicle year text box, look up vehicle year from vehicle table, set the vehicle year, and fire onkeypress event
        Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlProperty$txtVehicleYear")
        Set cElem = pageSource.Item(0)
        cElem.Value = Nz(DLookup("[VehYear]", "EvtVehicle", "[ID] = " & Forms!ViewEvent.VehicleID.Value), "")
        cElem.FireEvent "OnKeyPress"
        
        'check plate type of vehicle
        If Not Nz(DLookup("[Type]", "EvtVehicle", "[ID] = " & Forms!ViewEvent.VehicleID.Value), "") = "" Then
            Dim plateTypeString As String
            'look up plate type
            plateTypeString = DLookup("[Expr2]", "EvtVehicle", "[ID] = " & Forms!ViewEvent.VehicleID.Value)
            'get plate type text box and fill with 3 left most characters of the plate type
            Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlProperty$txtPlateType")
            Set cElem = pageSource.Item(0)
            cElem.Value = Left(plateTypeString, 3)
        End If
        
        'get vin text box, look up vin in vehicle table, and fill text box with vin
        Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlProperty$txtVIN")
        Set cElem = pageSource.Item(0)
        cElem.Value = Nz(UCase(DLookup("[VIN]", "EvtVehicle", "[ID] = " & Forms!ViewEvent.VehicleID.Value)), "")
        
        'get vehicle model text box, look up model in vehicle table, and fill text box with model
        Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlProperty$txtVehicleModel")
        Set cElem = pageSource.Item(0)
        cElem.Value = Nz(DLookup("[VehModel]", "EvtVehicle", "[ID] = " & Forms!ViewEvent.VehicleID.Value), "")
        
        'get vehicle style text box, look up style in vehicle table, and fill text box with style
        Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlProperty$txtVehicleStyle")
        Set cElem = pageSource.Item(0)
        cElem.Value = Nz(DLookup("[VehStyle]", "EvtVehicle", "[ID] = " & Forms!ViewEvent.VehicleID.Value), "")
        
        'get vehicle color text box, look up color in vehicle table, and fill text box with color
        Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlProperty$txtVehicleColor")
        Set cElem = pageSource.Item(0)
        cElem.Value = Nz(DLookup("[VehColor]", "EvtVehicle", "[ID] = " & Forms!ViewEvent.VehicleID.Value), "")
        
        'get vehicle make drop down box, look up vehicle make from vehicle table, select option for make, and fire onchange event
        Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlProperty$lstVehicleMake")
        Set cElem = pageSource.Item(0)
        cElem.Value = Nz(DLookup("[Expr1]", "EvtVehicle", "[ID] = " & Forms!ViewEvent.VehicleID.Value), 0)
        cElem.FireEvent "OnChange"
        
    ElseIf Me.WebBrowser0.ReadyState = 4 And Fill = True And FillStr = "CALL" Then
        Set HTMLdoc = Me.WebBrowser0.Document
        
        Dim calls As String, Narr As String, sql As String
        Dim db As Database
        Dim rs As DAO.Recordset
        Dim fld As DAO.Field
        
        calls = "*****" & vbNewLine
        sql = "SELECT [EventCallAssoc].[EventID], Year([Date1]) AS Expr3, commlog.Date1, commlog.Time1, " _
        & "commlog.UnitCalled, commlog.SourceCall, commlog.Reason, commlog.Narrative FROM commlog " _
        & "INNER JOIN EventCallAssoc ON commlog.id = EventCallAssoc.CallID WHERE (((EventCallAssoc.EventID) " _
        & "= " & Forms("ViewEvent")!Text25.Value & ")) ORDER BY Year([Date1]) DESC , commlog.Date1 DESC , commlog.Time1 DESC;"
        
        Set db = CurrentDb
        Set rs = db.OpenRecordset(sql)
        
        rs.MoveFirst
        
        While Not rs.EOF
            With rs
                calls = calls & .Fields(3) & " -- " & .Fields(4) & " -- " & .Fields(5) & " -- " & .Fields(6) & " -- " & .Fields(7) & vbNewLine
            End With
            
            rs.MoveNext
        Wend
        
        Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlEvent$txtNarrative")
        Set cElem = pageSource.Item(0)
        Narr = cElem.Value
        
        If InStr(1, Narr, "*****") > 0 Then
            Narr = Left(Narr, InStr(1, Narr, "*****") - 1)
        End If
        
        Narr = Narr & vbNewLine & calls
        
        cElem.Value = Narr
        
        Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$uctrlEvent$btnUpdate")
        Set cElem = pageSource.Item(0)
        
        cElem.Click
        
        Fill = False
    ElseIf Me.WebBrowser0.ReadyState = 4 And Fill = False Then
'        If catchEventID = True Then
'            Dim Address As String, sqlx As String
'            Dim parksID As Long
'            Dim x As Integer, y As Integer
'
''            Address = Me.WebBrowser0.LocationURL
''
''            If Not IsNull(InStr(1, Address, "slctEventId=&")) Then
''                x = InStr(1, Address, "slctEventId=") + 12
''                y = InStr(1, Address, "&")
''
''                If y - x > 0 Then
''                    parksID = Mid(Address, x, y - x)
''
''                    sqlx = "UPDATE events SET pId=" & parksID & " WHERE ID=" & Forms!ViewEvent!Text25.Value
''                    DoCmd.SetWarnings False
''                    DoCmd.RunSQL sqlx
''                    DoCmd.SetWarnings True
''                    catchEventID = True
''                End If
''            End If
'
'            Set HTMLdoc = Me.WebBrowser0.Document
'            Set pageSource = HTMLdoc.getElementsByName("ctl00$Container$hdnEventId")
'            Set cElem = pageSource.Item(0)
'
'            If Not IsNull(cElem.Value) And Not cElem.Value = 0 Then
'                parksID = cElem.Value
'                sqlx = "UPDATE events SET pId=" & parksID & " WHERE ID=" & Forms!ViewEvent!Text25.Value
'                DoCmd.SetWarnings False
'                DoCmd.RunSQL sqlx
'                DoCmd.SetWarnings True
'                catchEventID = True
'                MsgBox parksID
'            End If
'        End If
    End If
End Sub

