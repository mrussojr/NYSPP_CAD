Attribute VB_Name = "Form_ViewEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim saveRec As Integer ', mapLoaded As Integer

'address field exit event
Private Sub Combo12_Exit(Cancel As Integer)
    'if the address field exists, convert it to upper case
    If IsNull(Me.Combo12.Value) = False And Not Me.Combo12.Value = "" Then
        Me.Combo12.Value = UCase(Me.Combo12.Value)
    End If
End Sub

'print button
Private Sub Command34_Click()
    Dim myPath As String, strReportName As String
    
    'set the path of the pdf
    myPath = "C:\Temp\"
    'set the pdf name
    strReportName = Me.Text8.Value & " - " & Me.Text19.Value & ".pdf"
    
    'save the pdf with the name save and the path provided
    DoCmd.OutputTo acOutputForm, "", acFormatPDF, myPath & strReportName, True
End Sub

'fill parks event button
Private Sub Command35_Click()
    'check to see the parks event has already been created
    If IsNull(Me.Text38.Value) = True Or Me.Text38.Value = 0 Then
        'if the event has not been created, open the parks event page to populate it with the event info
        OpenEvent Me.Text36.Value, "Fill"
    Else
        'else open the event page with the parks id, to view or edit
        OpenExistingEvent Me.Text38.Value, Me.Text36.Value
    End If
End Sub

'add calls to parks event
Private Sub Command43_Click()
    If IsNull(Me.Text38.Value) = True Or Me.Text38.Value = 0 Then
        'if the event has not been created, open the parks event page to populate it with the event info
        'OpenEvent Me.Text36.Value, "Fill"
    Else
        'else open the event page with the parks id, to view or edit
        OpenExistingEvent Me.Text38.Value, Me.Text36.Value, "CALL"
    End If
End Sub

'maps available button
Private Sub Command44_Click()
    Dim loc As String
    Dim count As Integer, locId As Integer
    
    'set the location to the selected location of the event
    loc = Me.Text10.Value
    'lookup the location id
    locId = DLookup("[ID]", "AllLocations", "[Park] = '" & loc & "'")

    DoCmd.OpenForm "FacilityMaps", , , , , , locId
End Sub

'save event button
Private Sub Command6_Click()
    'set variable to allow saving of record
    saveRec = 1
    'perform save command
    DoCmd.RunCommand acCmdSaveRecord
    'advise user via message box that the event has been saved
    MsgBox "Event has been saved!"
    'close the view event form
    DoCmd.Close acForm, "ViewEvent", acSaveNo
End Sub

'before saving a record, check to see if the local saverecord variable has been set to 1
Private Sub Form_BeforeInsert(Cancel As Integer)
    If Command = acCmdSaveRecord Then
        If saveRec = 0 Then
            'if the saverecord variable is 0, cancel the save operation
            Cancel = True
        Else
            'else, save the record
            Cancel = False
        End If
    End If
End Sub

'when the form loads
Private Sub Form_Load()
    'set the save record variable to 0, disallowing the save operation automatically
    saveRec = 0
    'set the sql injection warnings false
    DoCmd.SetWarnings False
    'set the focuse of the form to the close button
    Me.Command4.SetFocus
    'check the opening arguments of the form
    If IsNull(Me.OpenArgs) = False And Not Me.OpenArgs = "" Then
        'if not null, set the filter of the form so that only the selected event is shown
        Me.Filter = "ID = " & Me.OpenArgs
        'turn the form filter on
        Me.FilterOnLoad = True
        Me.FilterOn = True
    End If
    
    'check the map location field, if it's not empty navigate to the mini map
    If IsNull(Me.MapLoc.Value) = False And Not Me.MapLoc.Value = "" Then
        Me.WebBrowser6.Navigate Me.MapLoc.Value
    End If
    
    Dim zoneId As Integer
    'check the officer id field, if its not empty look up the zone id and fill the zone id field
    If Not IsNull(Me.Text36.Value) = True Or Not Me.Text36.Value = "" Then
        zoneId = DLookup("[pZone]", "parksIDs", "[pId] = " & Me.Text36.Value)
        Me.Text37.Value = DLookup("[zone]", "zones", "[pId] =  " & zoneId)
        'refresh the park / municipality field to show only locations within the officers zone
        Me.Text10.Requery
    End If
    'call the exit procedure for the cross street to load the mini map
    Call Text14_Exit(1)
    'mapLoaded = 0
    
    'following section will show or hide the map button
    Dim loc As String
    Dim count As Integer, locId As Integer
    
    'set the location to the selected location of the event
    loc = Me.Text10.Value
    'lookup the location id
    locId = DLookup("[ID]", "AllLocations", "[Park] = '" & loc & "'")
    'use the id to count the number of maps in the map table for that location
    count = DCount("[id]", "maps", "[parkId] = " & locId)
    'if there are maps show the map button
    If count > 0 Then
        Me.Command44.Visible = True
    Else 'else hide the map button
        Me.Command44.Visible = False
    End If
'Checking to see if there is a PARKS id numberassociaed with the event and informing the user via PARKS_lbl
    Dim PARKS_id
    PARKS_id = DLookup("[pid]", "events", "[ID] = " & Me.Text25.Value)
    If PARKS_id = 0 Then
        PARKS_lbl.Caption = "A PARKS event has NOT been started!"
    ElseIf PARKS_id > 0 Then
        PARKS_lbl.Caption = "A PARKS event has been started."
    End If
End Sub

'cancel button
Private Sub Command4_Click()
    'undo any changes
    Me.Undo
    'close the viewevent form without saving
    DoCmd.Close acForm, "ViewEvent", acSaveNo
End Sub

'when the form is unloaded, reenable the sql injection warnings
Private Sub Form_Unload(Cancel As Integer)
    DoCmd.SetWarnings True
End Sub

'upon exiting the location field and it is filled it, set the location fields value to upper case
Private Sub Text10_Exit(Cancel As Integer)
    If IsNull(Me.Text10.Value) = False And Not Me.Text10.Value = "" Then
        Me.Text10.Value = UCase(Me.Text10.Value)
    End If
End Sub

'when exiting the cross street field
Private Sub Text14_Exit(Cancel As Integer)
    Dim webAddress As String, Address As String, XStreet As String, loc As String, totAddy As String
    
    'if the cross street is entered, change it to upper case
    If IsNull(Me.Text14.Value) = False And Not Me.Text14.Value = "" Then
        Me.Text14.Value = UCase(Me.Text14.Value)
    End If
    
    'check to see if the location was entered
    If IsNull(Me.Text10.Value) = False Then
        'if the location was entered, set the location variable to the location field
        loc = Me.Text10.Value
        'replace any spaces in the location to +, required by google maps api
        loc = Replace(loc, " ", "+")
        'add ny to the end of the location
        loc = loc & ",+NY"
    Else
        loc = ""
    End If
    
    'check to see if the address was entered
    If IsNull(Me.Combo12.Value) = False Then
        'if the address was entered, set the address variable to the address field
        Address = Me.Combo12.Value
        'replace any spaces in the address variable with +, required by google maps api
        Address = Replace(Address, " ", "+")
    Else
        'if the address is null, set it to an empty string, to be ignored when building the google maps api location string
        Address = ""
    End If
    
    'check to see if the cross street was entered
    If IsNull(Me.Text14.Value) = False Then
        'if the cross street was entered, set the cross street variable to the cross street field
        XStreet = Me.Text14.Value
        'replace any spaces with +, required by google maps api
        XStreet = Replace(XStreet, " ", "+")
    Else
        'if the cross street is empty, set the variable to an empty string to be ignored
        XStreet = ""
    End If
    
    'check to make sure that there is a location, address, or cross street, exit if all are null
    If loc = "" And Address = "" And XStreet = "" Then
        Exit Sub
    Else
        'set the total address string to be the address, cross street and location
        totAddy = Address & "+and+" & XStreet & ",+" & loc
        'set the web address to the address of the google maps api with the total address string injected
        webAddress = "http://maps.googleapis.com/maps/api/staticmap?center=" & totAddy & "&zoom=14&size=400x200&sensor=false&visual_refresh=true&maptype=roadmap&markers=icon:http://goo.gl/6fmKFh%7Ccolor:red%7Clabel:X%7C" & totAddy
                
        'navigate the web browser to the web address
        Me.WebBrowser6.Navigate webAddress
        'set the map location field to the web address
        Me.MapLoc.Value = webAddress
        'set the total address field
        Me.Text41.Value = totAddy
    End If
End Sub

'when exiting the narrative field, if it's not empty, convert it to upper case
Private Sub Text17_Exit(Cancel As Integer)
    If IsNull(Me.Text17.Value) = False And Not Me.Text17.Value = "" Then
        Me.Text17.Value = UCase(Me.Text17.Value)
    End If
End Sub

'when the officer field changes, refresh the location field to show only locations in that officers zone
Private Sub Text8_Change()
    Dim zoneId As Integer
    zoneId = DLookup("[pZone]", "parksIDs", "[pId] = " & Me.Text36.Value)
    Me.Text37.Value = DLookup("[zone]", "zones", "[pId] = " & zoneId)
    Me.Text10.Requery
End Sub

'when exiting the officer field, if it's not empty, convert it to upper case
Private Sub Text8_Exit(Cancel As Integer)
    If IsNull(Me.Text8.Value) = False And Not Me.Text8.Value = "" Then
        Me.Text8.Value = UCase(Me.Text8.Value)
    End If
End Sub

'when entering (clicking) the minimap, open the larger interactive map
Private Sub WebBrowser6_Enter()
    'MsgBox Me.Text41.Value
'    If mapLoaded = 0 Then
'        Call Text14_Exit(1)
'        mapLoaded = 1
'    End If
    a_testGeocode Me.Text41.Value
End Sub
