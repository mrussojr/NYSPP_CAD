Attribute VB_Name = "Rebuilder"
Option Compare Database

Public Sub LoadObjects()
    loadTables
    loadModules
    loadQueries
    loadFormsReports "forms"
    loadFormsReports "reports"
End Sub

Private Sub loadFormsReports(objType As String)
    Dim fs, oFolder
    Dim db As DAO.Database
    Dim qDef As DAO.QueryDef
    Dim sqlTxt As String, objNam As String, newFile As String
    Dim objAcTyp As AcObjectType
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set oFolder = fs.GetFolder(CurrentProject.Path & "\" & objType & "\properties")
    
    Dim objTypeStr As String
    
    If objType = "forms" Then
        objTypeStr = "Form"
        objAcTyp = acForm
    Else
        objTypeStr = "Report"
        objAcTyp = acReport
    End If
    
    For Each aFile In oFolder.Files
        objNam = Mid(aFile.NAME, 1, Len(aFile.NAME) - 4)
        
        Debug.Print objNam
        
        newFile = CurrentProject.Path & "\" & objType & "\" & objNam & "_combined.txt"
        
        Dim bFile As Object
        Dim cFile As Object
        
        Set cFile = fs.CreateTextFile(newFile)
        
        Open aFile For Input As #1
        
        Do While Not EOF(1)    ' Loop until end of file.
            Line Input #1, strTextLine
            cFile.writeLine strTextLine & vbCrLf
        Loop
        
        Close #1
        
        If fs.FileExists(CurrentProject.Path & "\" & objType & "\code\" & objTypeStr & "_" & objNam & ".cls") Then
            Set bFile = fs.OpenTextFile(CurrentProject.Path & "\" & objType & "\code\" & objTypeStr & "_" & objNam & ".cls")
            
            cFile.writeLine "CodeBehindForm" & vbCrLf
            
            cFile.writeLine bFile.readAll
            
            Set bFile = Nothing
        End If
        
        cFile.Close
        
        LoadFromText objAcTyp, objNam, newFile
        
        fs.deleteFile newFile
    Next aFile
    
    Set fs = Nothing
End Sub

Private Sub loadQueries()
    Dim fs, oFolder
    Dim db As DAO.Database
    Dim qDef As DAO.QueryDef
    Dim sqlTxt As String
    
    Set db = CurrentDb
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set oFolder = fs.GetFolder(CurrentProject.Path & "\queries")
    
    For Each aFile In oFolder.Files
        Set qDef = New DAO.QueryDef
        
        qDef.NAME = Mid(aFile.NAME, 1, Len(aFile.NAME) - 4)
        
        sqlTxt = ""
        
        Open aFile For Input As #1
        
        Do While Not EOF(1)    ' Loop until end of file.
            Line Input #1, strTextLine
            
            sqlText = sqlText & strTextLine & vbCrLf
        Loop
        
        Close #1
        
        sqlText = Trim(sqlText)
        sqlText = Left(sqlText, Len(sqlText) - 2)
        
        qDef.sql = sqlText
        
        db.QueryDefs.Append qDef
        
        Set qDef = Nothing
        
        sqlText = ""
    Next aFile
End Sub

Private Sub loadModules()
    Dim fs, oFolder
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set oFolder = fs.GetFolder(CurrentProject.Path & "\modules")
    
    For Each aFile In oFolder.Files
        LoadFromText acModule, Mid(aFile.NAME, 1, Len(aFile.NAME) - 4), aFile
    Next aFile
    
    Set oFolder = fs.GetFolder(CurrentProject.Path & "\classes")
    
    For Each aFile In oFolder.Files
        DoCmd.RunCommand acCmdNewObjectClassModule
        
        Application.ReplaceModule acModule, "Class1", aFile, 0
    Next aFile
    
End Sub

Private Sub loadTables()
    Dim fs, oFolder
    Dim db As DAO.Database
    
    Set db = CurrentDb
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set oFolder = fs.GetFolder(CurrentProject.Path & "\tables")
    
    For Each aFile In oFolder.Files
        Debug.Print aFile.NAME
        Dim tabNam As String
        Dim catch As Boolean
        
        tabNam = Mid(aFile.NAME, 1, Len(aFile.NAME) - 4)
        
        Dim tabDef As DAO.TableDef
        Dim fldDef As DAO.Field
        Dim idxDef As DAO.Index
        
        Set tabDef = New DAO.TableDef
        
        tabDef.NAME = tabNam
        
        Open aFile For Input As #1
        
        Dim lineNo As Integer
        
        lineNo = 0
        
        Do While Not EOF(1)    ' Loop until end of file.
            Line Input #1, strTextLine
            lineNo = lineNo + 1
            'Debug.Print lineNo & " - " & strTextLine
            
            If catch Then
                Dim fldArr() As String
                Dim tmpLine As String
                
                tmpLine = Trim(strTextLine)
                tmpLine = Replace(tmpLine, vbTab, "")
                
                fldArr() = Split(tmpLine)
                
                If UBound(fldArr) > 0 Then
                    Set fldDef = New DAO.Field
                    
                    Dim fldNam As String
                    
                    Dim fldArrIdx As Integer
                    Dim fldFlNamCaught As Boolean
                    
                    fldNam = ""
                    fldFlNamCaught = False
                    fldArrIdx = 0
                    
                    'Debug.Print UBound(fldArr)
                    
                    Do While (fldArrIdx < UBound(fldArr) And Not fldFlNamCaught)
                        fldNam = fldNam & fldArr(fldArrIdx)

                        If InStr(1, fldArr(fldArrIdx), "]") > 0 Then
                            fldFlNamCaught = True
                        End If
                        
                        fldArrIdx = fldArrIdx + 1
                    Loop
                    
                    If InStr(1, fldNam, ",") > 0 Then
                        tmpFldNam = Replace(fldNam, ",", "")
                        fldNam = tmpFldNam
                    End If
                    
                    fldDef.NAME = Replace(Replace(fldNam, "[", ""), "]", "")
                    
                    Select Case fldArr(fldArrIdx)
                        Case "Counter"
                            fldDef.Type = dbLong
                            fldDef.Attributes = DAO.FieldAttributeEnum.dbAutoIncrField
                        Case "Long"
                            fldDef.Type = dbLong
                        Case "Text"
                            fldDef.Type = dbText
                        Case "YesNo"
                            fldDef.Type = dbBoolean
                        Case "Byte"
                            fldDef.Type = dbByte
                        Case "Integer"
                            fldDef.Type = dbInteger
                        Case "Currency"
                            fldDef.Type = dbCurrency
                        Case "Single"
                            fldDef.Type = dbSingle
                        Case "Double"
                            fldDef.Type = dbDouble
                        Case "DateTime"
                            fldDef.Type = dbDate
                        Case "Binary"
                            fldDef.Type = dbBinary
                        Case "OLE Object"
                            fldDef.Type = dbLongBinary
                        Case "Memo"
                            fldDef.Type = dbMemo
                        Case "Hyperlink"
                            fldDef.Type = dbMemo
                            fldDef.Attributes = DAO.FieldAttributeEnum.dbHyperlinkField
                        Case "GUID"
                            fldDef.Type = dbGUID
                    End Select
                    
                    If UBound(fldArr) >= fldArrIdx + 1 Then
                        If Not fldArr(fldArrIdx + 1) = ")""" Then
                            fldDef.size = CInt(Replace(Replace(fldArr(fldArrIdx + 1), "(", ""), ")", ""))
                        End If
                    End If
                
                    tabDef.Fields.Append fldDef
                    
                    'Debug.Print "COL: " & fldDef.Name & "|" & fldDef.Type & "|" & fldDef.Size
                    
                    Set fldDef = Nothing
                    fldNam = ""
                End If
            End If
            
            If Right(strTextLine, 1) = """" And catch Then
                'Debug.Print vbTab & "End field catch"
                catch = False
            End If
            
            If Left(strTextLine, 7) = """CREATE" Then
                If Not Mid(strTextLine, 9, 5) = "TABLE" Then
                    'catch = True
                    
                    Dim idxArr() As String
                    Dim idxArrIdx As Integer
                    Dim nameCaught As Boolean
                    
                    nameCaught = False
                    
                    idxArrIdx = 0
                    
                    Set idxDef = New DAO.Index
                    
                    idxArr = Split(strTextLine)
                    
                    Do While idxArrIdx < UBound(idxArr)
                        'Debug.Print vbTab & "IDX_PARM" & vbTab & idxArr(idxArrIdx)
                        
                        If idxArr(idxArrIdx) = "UNIQUE" Then
                            idxDef.Unique = True
                        ElseIf Left(idxArr(idxArrIdx), 1) = "[" And Not nameCaught Then
                            idxDef.NAME = Replace(Replace(idxArr(idxArrIdx), "[", ""), "]", "")
                            nameCaught = True
                        ElseIf Left(idxArr(idxArrIdx), 2) = "([" Then
                            idxDef.Fields = Replace(Replace(idxArr(idxArrIdx), "([", ""), "])", "")
                        ElseIf idxArr(idxArrIdx) = "PRIMARY" Then
                            idxDef.Primary = True
                        ElseIf idxArr(idxArrIdx) = "DISALLOW" Then
                            idxDef.Required = True
                        ElseIf idxArr(idxArrIdx) = "IGNORE" Then
                            idxDef.IgnoreNulls = True
                        End If
                        
                        idxArrIdx = idxArrIdx + 1
                    Loop
                    
                    Debug.Print "IDX: " & idxDef.NAME
                    
                    tabDef.Indexes.Append idxDef
                    
                    Set idxDef = Nothing
                Else
                    catch = True
                End If
            End If
        Loop
        
        Close #1
        
        db.TableDefs.Append tabDef
        db.TableDefs.Refresh
        
        Set tabDef = Nothing
        'LoadFromText acTable, tabNam, aFile
    Next aFile
End Sub
