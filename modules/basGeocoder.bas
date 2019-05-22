Attribute VB_Name = "basGeocoder"
Option Compare Database
Option Explicit
'* v1.2 2014-01-26
'*   Added a wait function ("fGGCWait") to help deal with the Google-enforced request rate limit.
'* v1.1 2013-09-12
'*   New method for determining if an internet connection is available
#If VBA7 Then
    Private Declare PtrSafe Function InternetGetConnectedState Lib "wininet.dll" _
        (ByRef dwflags As LongPtr, ByVal dwReserved As Long) As Long
#Else
    Private Declare Function InternetGetConnectedState Lib "wininet.dll" _
        (ByRef dwflags As Long, ByVal dwReserved As Long) As Long
#End If

Type LatLng
    Status As String
    Lat As Double
    lng As Double
End Type


Function fIsConnectedToInternet() As Integer
'******************************************************************************
'* Peter De Baets, author
'* 5/23/2011
'* This routine checks to see if an internet connnection is available
'******************************************************************************

Dim Marker As Integer
Dim Rtn As Integer

On Error Resume Next
DoCmd.Hourglass True
Call SysCmd(4, "Checking for an internet connection...")
Marker = 1
Rtn = False

If InternetGetConnectedState(0&, 0&) Then
    '* We're connected to the internet
Else
    MsgBox "Internet connection not established.", vbCritical
    GoTo Exit_Section
End If

Rtn = True

Exit_Section:
    On Error Resume Next
    Call SysCmd(5)
    DoCmd.Hourglass False
    fIsConnectedToInternet = Rtn
    On Error GoTo 0
    Exit Function
End Function

Function ng_ggcAutoexec() As Integer
'* Runs at demo startup
Dim Marker As Integer
Dim Rtn As Integer

On Error GoTo Err_Section
Marker = 1

Select Case Val(SysCmd(7))
Case Is >= 12
    '*2007 Minimize navigation window
    DoCmd.SelectObject acForm, , True
    RunCommand acCmdWindowHide
Case Else
End Select

DoCmd.OpenForm "frmGGCMain"

Exit_Section:
    On Error Resume Next
    ng_ggcAutoexec = Rtn
    On Error GoTo 0
    Exit Function
Err_Section:
    Select Case Err
    Case Else
        Beep
        MsgBox "Error in ng_ggcAutoexec (" & Marker & "), object " & Err.Source & ": " & Err.Number & " - " & Err.Description
    End Select
    Err.Clear
    Resume Exit_Section
End Function


Sub a_testGeocode(Optional strAddr As String)
'Dim strAddr As String
Dim loc As LatLng

If fIsConnectedToInternet Then
Else
    Exit Sub
End If

'strAddr = "2400 Bayshore Parkway, Mountain View, CA"
'strAddr = Forms!EnterEvent!maplocation

loc = fGeocode(strAddr)
If loc.Lat = 0 Then
    MsgBox "Invalid results"
Else
    'MsgBox "Latitude and Longitude for '" & strAddr & "' is: " & loc.Lat & ", " & loc.lng
    If Not CurrentProject.AllForms("MapViewer").IsLoaded Then
        DoCmd.OpenForm "MapViewer", acNormal, , , , , loc.Lat & "," & loc.lng
    End If
End If
End Sub
Function fGetFullAddrNoApt( _
    pstrAddress As Variant, _
    pstrCity As Variant, _
    pstrState As Variant, _
    pstrZip As Variant, _
    pstrCountry As Variant, _
    Optional pintUseDefaultCityStateCountryIfBlank As Integer = True _
    ) As String
'******************************************************************************
'* Peter De Baets, author
'* 5/23/2011
'* Returns an address with default city, state and country filled in where blank,
'* (see inline comments below) and apartment numbers removed. The returned address
'* can be used to geocode the address location.
'******************************************************************************
Dim Rtn As String
Dim s As String
Dim strAddress As String
Dim strCity As String
Dim strState As String
Dim strZip As String
Dim strCountry As String
Dim ipos As Integer
Dim Marker As Integer

On Error GoTo Err_Section
Marker = 1

strCity = Trim("" & pstrCity)
strState = Trim("" & pstrState)
strCountry = Trim("" & pstrCountry)
If pintUseDefaultCityStateCountryIfBlank Then
    If strCity = "" Then strCity = "Los Angeles"    '<<<--- Put the default city here!!
    If strState = "" Then strState = "CA"       '<<<--- Put the default state here!!
    If strCountry = "" Then strCountry = "USA"       '<<<--- Put the default country here!!
End If
strZip = Trim("" & pstrZip)

s = ""
s = Trim("" & pstrAddress)

If Trim(s) = "" Then
Else
    '* Take out the apartment number info
    ipos = 0
    ipos = InStr(1, s, " apt")
    If ipos = 0 Then
        ipos = InStr(1, s, "#")
        If ipos > 1 Then ipos = ipos - 1
    Else
    End If
    If ipos = 0 Then
    Else
        s = Trim(Left(s, ipos))
    End If
    s = Replace(s, " N. ", " N ")
    s = Replace(s, " S. ", " S ")
    s = Replace(s, " E. ", " E ")
    s = Replace(s, " W. ", " W ")
End If
strAddress = Trim("" & s)

Rtn = strAddress
If Trim(Rtn) = "" Then
    If Trim("" & strCity) = "" Then
        If Trim("" & strState) = "" Then
            If Trim("" & strZip) = "" Then
                Rtn = Trim(strCountry)
            Else
                Rtn = Rtn & Trim("" & strZip)
                If Trim("" & strCountry) = "" Then
                Else
                    Rtn = Rtn & ", " & strCountry
                End If
            End If
        Else
            Rtn = Rtn & strState
            If Trim("" & strZip) = "" Then
                If Trim("" & strCountry) = "" Then
                Else
                    Rtn = Rtn & ", " & strCountry
                End If
            Else
                Rtn = Rtn & " " & strZip
                If Trim("" & strCountry) = "" Then
                Else
                    Rtn = Rtn & ", " & strCountry
                End If
            End If
        End If
    Else
        Rtn = Rtn & strCity
        If Trim("" & strState) = "" Then
            Rtn = Rtn & " " & pstrZip
            If Trim("" & strCountry) = "" Then
            Else
                Rtn = Rtn & ", " & strCountry
            End If
        Else
            Rtn = Rtn & ", " & strState
            If Trim("" & strZip) = "" Then
                If Trim("" & strCountry) = "" Then
                Else
                    Rtn = Rtn & ", " & strCountry
                End If
            Else
                Rtn = Rtn & " " & strZip
                If Trim("" & strCountry) = "" Then
                Else
                    Rtn = Rtn & ", " & strCountry
                End If
            End If
        End If
    End If
Else
    If Trim("" & strCity & strState & strZip & strCountry) = "" Then
    Else
        Rtn = Rtn & ", "
        If Trim("" & strCity) = "" Then
            If Trim("" & strState) = "" Then
                If Trim("" & strZip) = "" Then
                    Rtn = Rtn & strCountry
                Else
                    Rtn = Rtn & strZip
                    If Trim("" & strCountry) = "" Then
                    Else
                        Rtn = Rtn & ", " & strCountry
                    End If
                End If
            Else
                Rtn = Rtn & strState
                If Trim("" & strZip) = "" Then
                Else
                    Rtn = Rtn & " " & strZip
                End If
                If Trim("" & strCountry) = "" Then
                Else
                    Rtn = Rtn & ", " & strCountry
                End If
            End If
        Else
            Rtn = Rtn & strCity
            If Trim("" & strState) = "" Then
                Rtn = Trim(Rtn & " " & strZip)
                If Trim("" & strCountry) = "" Then
                Else
                    Rtn = Rtn & ", " & strCountry
                End If
            Else
                Rtn = Rtn & ", " & strState
                If Trim("" & strZip) = "" Then
                Else
                    Rtn = Rtn & " " & strZip
                End If
                If Trim("" & strCountry) = "" Then
                Else
                    Rtn = Rtn & ", " & strCountry
                End If
            End If
        End If
    End If
End If


Exit_Section:
    On Error Resume Next
    fGetFullAddrNoApt = Rtn
    On Error GoTo 0
    Exit Function
Err_Section:
    Select Case Err
    Case Else
        Beep
        MsgBox "Error in fGetFullAddrNoApt (" & Marker & "), object " & Err.Source & ": " & Err.Number & " - " & Err.Description
    End Select
    Err.Clear
    Resume Exit_Section
End Function


Function fGeocode( _
    pstrAddr As String _
    ) As LatLng
'******************************************************************************
'* Peter De Baets, author
'* 3/8/2011
'*
'******************************************************************************

Dim Marker As Integer
Dim Rtn As LatLng
Dim strAddr As String
Dim strXML As String

On Error GoTo Err_Section
Marker = 1

strAddr = Replace(pstrAddr, " ", "+")
strXML = fGetXMLText("http://maps.googleapis.com/maps/api/geocode/xml?address=" & strAddr & "&sensor=false")
'Debug.Print strXML
Rtn.Status = fGetXMLNodeText(strXML, "GeocodeResponse;status")
If Rtn.Status = "OK" Then
    Rtn.Lat = CDbl(fGetXMLNodeText(strXML, "GeocodeResponse;result;geometry;location;lat"))
    Rtn.lng = CDbl(fGetXMLNodeText(strXML, "GeocodeResponse;result;geometry;location;lng"))
Else
    Rtn.Lat = 0
    Rtn.lng = 0
End If

Exit_Section:
    On Error Resume Next
    fGeocode = Rtn
    On Error GoTo 0
    Exit Function
Err_Section:
    Select Case Err
    Case Else
        Beep
        MsgBox "Error in fGeocode (" & Marker & "), object " & Err.Source & ": " & Err.Number & " - " & Err.Description
    End Select
    Err.Clear
    Resume Exit_Section
End Function
Public Function fGGCWait(plngSeconds As Long)
Dim lngNow As Long
lngNow = Timer()
Do
    DoEvents
Loop While (Timer < lngNow + plngSeconds)
End Function
Private Function xg_CountDelimitedStrings(mainstr As String, delimiter As String) As Integer
'* Returns the number of strings in 'mainstr' that are delimited by 'delimiter'
'* Ex.: xg_CountDelimitedStrings("abc;def;ghi", ";") returns 3
Dim i As Integer
Dim ipos As Integer
Dim inewpos As Integer
Dim found As Integer
Dim Rtn As Integer
On Error GoTo Err_Section

Rtn = 0
If IsNull(mainstr) Then
Else
    ipos = 1
    found = True
    Do While found
        inewpos = InStr(ipos, mainstr, delimiter)
        If inewpos = 0 Then
            found = False
        Else
            If inewpos > ipos Then
                Rtn = Rtn + 1
            End If
            ipos = inewpos + 1
        End If
    Loop
    If Len(mainstr) > (ipos - 1) Then
        Rtn = Rtn + 1
    End If
End If
'MsgBox "# of delimited strings = " & rtn
xg_CountDelimitedStrings = Rtn
Exit Function

Err_Section:
    MsgBox "Error " & Err & " in function xg_CountDelimitedStrings " & Err.Description
    Exit Function

End Function
Function fGetXMLNodeText( _
    pstrXML As String, _
    pstrXMLFullNode As String _
    ) As String
'******************************************************************************
'* Peter De Baets, author
'* 9/4/2010
'*
'******************************************************************************

Dim Marker As Integer
Dim s As String
Dim i As Integer
Dim j As Integer
Dim strBranch As String
Dim intBranches As Integer
Dim strLastBranch As String
Dim objXML As New MSXML2.DOMDocument
Dim Rtn As String
Dim pn As IXMLDOMNode
Dim n As IXMLDOMNode

On Error GoTo Err_Section
Marker = 1

If Not objXML.loadXML(pstrXML) Then
    Err.Raise objXML.parseError.errorCode, , objXML.parseError.Reason
    GoTo Exit_Section
End If

'Debug.Print objXML.hasChildNodes
intBranches = xg_CountDelimitedStrings(pstrXMLFullNode, ";")
strLastBranch = xg_GetSubString(pstrXMLFullNode, intBranches, ";")
j = 1
strBranch = xg_GetSubString(pstrXMLFullNode, j, ";")
i = 0
Set pn = objXML
'Debug.Print pn.childNodes("GeocodeResponse").childNodes("result").childNodes("geometry").childNodes("location").childNodes("lat").Text
Do While True
    Set n = pn.childNodes(i)
    'Debug.Print "  Node " & i & ": " & n.nodeName
    If n.nodeName = strBranch Then
        '* We've found our branch
        If j = intBranches Then
            '* it is the final branch. Return the value
            Rtn = n.Text
            Exit Do
        Else
            j = j + 1
            strBranch = xg_GetSubString(pstrXMLFullNode, j, ";")
            Set pn = n
            i = -1
        End If
    Else
    End If
    If n.nodeName = strLastBranch Then
        Exit Do
    Else
        i = i + 1
    End If
Loop


'Debug.Print point.nodeName
'Debug.Print point.hasChildNodes
'Debug.Print point.selectSingleNode("GeocodeResponse").Text

Exit_Section:
    On Error Resume Next
    Set objXML = Nothing
    Set n = Nothing
    Set pn = Nothing
    fGetXMLNodeText = Rtn
    On Error GoTo 0
    Exit Function
Err_Section:
    Select Case Err
    Case Else
        Beep
        MsgBox "Error in fGetXMLNodeText (" & Marker & "), object " & Err.Source & ": " & Err.Number & " - " & Err.Description
    End Select
    Err.Clear
    Resume Exit_Section
End Function
Private Function xg_GetSubString(mainstr As String, n As Integer, delimiter As String) As String
'* Get the "n"-th substring from "mainstr" where strings are delimited by "delimiter"
    Dim i As Integer
    Dim substringcount As Integer
    Dim Pos As Integer
    Dim strx As String
    Dim val1 As Integer
    Dim W As String

On Error GoTo Err_xg_GetSubString

W = ""
substringcount = 0
i = 1
Pos = InStr(i, mainstr, delimiter)
Do While Pos <> 0
    strx = Mid(mainstr, i, Pos - i)
    substringcount = substringcount + 1
    If substringcount = n Then
        Exit Do
    End If
    'pddxxx In case the delimiter is more than one char
    'i = pos + 1
    i = Pos + Len(delimiter)
    Pos = InStr(i, mainstr, delimiter)
Loop

If substringcount = n Then
    xg_GetSubString = strx
Else
    strx = Mid(mainstr, i, Len(mainstr) + 1 - i)
    substringcount = substringcount + 1
    If substringcount = n Then
        xg_GetSubString = strx
    Else
        xg_GetSubString = ""
    End If
End If

On Error GoTo 0
Exit Function

Err_xg_GetSubString:
    MsgBox "xg_GetSubString " & Err & " " & Err.Description
    Resume Next

End Function
Function fGetXMLText( _
    pstrURL As String _
    ) As String
'******************************************************************************
'* Peter De Baets, author
'* 9/4/2010
'* Requires a referenc to "Microsoft XML, v3.0", or later
'******************************************************************************

Dim Marker As Integer
Dim Rtn As String
Dim s As String
Dim HttpReq As New MSXML2.XMLHTTP30
Dim objXML As New MSXML2.DOMDocument

On Error GoTo Err_Section
Marker = 1

HttpReq.Open "GET", pstrURL, False
HttpReq.send
Rtn = HttpReq.responseText

Exit_Section:
    On Error Resume Next
    Set HttpReq = Nothing
    Set objXML = Nothing
    fGetXMLText = Rtn
    On Error GoTo 0
    Exit Function
Err_Section:
    Select Case Err
    Case Else
        Beep
        MsgBox "Error in fGetXMLText (" & Marker & "), object " & Err.Source & ": " & Err.Number & " - " & Err.Description
    End Select
    Err.Clear
    Resume Exit_Section
End Function

