https://dev.azure.com/xxx/xxx/_apis/test/plans?api-version=5.0
https://dev.azure.com/xxx/xxx/_apis/test/Plans/TESTSUITEID/suites?$expand={$expand}&$skip={$skip}&$top={$top}&$asTreeView={$asTreeView}&api-version=5.0
https://dev.azure.com/xxx/xxx/_apis/test/Plans/TESTSUITEID/Suites/TESTCASEID/testcases
https://dev.azure.com/xxx/xxx/_apis/test/Plans/PLANID/Suites/SUITEID/points?api-version=6.0
https://github.com/xuanhavcdn/CypressAutomation_Interview_Test/blob/main/cypress/e2e/tests/login.cy.ts

  String filePath = getResourcePath(fileName);
        System.out.println("filePath forupload is "+ filePath);        
        WebElement element = Parq.getDriver().findElement(By.xpath("//input[@type='file']"));            				
		LocalFileDetector detector = new LocalFileDetector();
		((RemoteWebElement)element).setFileDetector(detector);		
		System.out.println("filePath is "+ filePath);
		element.sendKeys(detector.getLocalFile(filePath).getAbsolutePath());

'Option Explicit
Private p&, token, dic
Sub azure_TestPaln()

Dim startTime As Double, endTime As Double
startTime = Timer

DeleteRange = "A2:T90000"
On Error Resume Next
Sheets("Dictionary").Range(DeleteRange).ClearContents
Sheets("TestPlans").Range(DeleteRange).ClearContents
Sheets("TestSuites").Range(DeleteRange).ClearContents
Sheets("TestcaseLists").Range(DeleteRange).ClearContents
On Error GoTo 0


DeleteRange = "A2:T90000"
On Error Resume Next
Sheets("Dictionary").Range(DeleteRange).ClearContents
Sheets("TestPlans").Range(DeleteRange).ClearContents
Sheets("TestSuites").Range(DeleteRange).ClearContents
Sheets("TestcaseLists").Range(DeleteRange).ClearContents
On Error GoTo 0

TestCycleResult = Sheets("Macro").ComboBox1.Value

'ReadAZUREAPI URLS

TestPlanazureURL = Sheets("APIURLS").Cells(2, 3)
TestPlanSuiteURL = Sheets("APIURLS").Cells(3, 3)
TestcasesURL = Sheets("APIURLS").Cells(4, 3)
TestcasePointsLink = Sheets("APIURLS").Cells(5, 3)

''Read the

'------------------------------------------------------------------------------------------
''List out all testplans in Dictionary sheet
'------------------------------------------------------------------------------------------

Call callAPI(TestPlanazureURL, 1, json$)

Set dic = ParseJSON(json$)

For Each key In dic.Keys
    CurrentColRowscount = Sheets("Dictionary").Range("A" & Rows.Count).End(xlUp).Row + 1
    Sheets("Dictionary").Cells(CurrentColRowscount, 1) = "Key: " & key & " <BRvalue:> " & dic(key)
Next


CurrentDicRowscount = Sheets("Dictionary").Range("A" & Rows.Count).End(xlUp).Row + 1
columnIndex = 1
For intDicount = 2 To CurrentDicRowscount - 2
      dictValue = Sheets("Dictionary").Cells(intDicount, 1)
      splitdictvalue = Split(dictValue, "<BRvalue:>")
      If intDicount = 2 Then
       CurrentTestSuitescount = Sheets("TestPlans").Range("A" & Rows.Count).End(xlUp).Row + 1
      End If
      Sheets("TestPlans").Cells(CurrentTestSuitescount, columnIndex) = Trim(splitdictvalue(1))
      If InStr(splitdictvalue(0), "clientUrl") Then
         CurrentTestSuitescount = Sheets("TestPlans").Range("A" & Rows.Count).End(xlUp).Row + 1
         columnIndex = 0
      End If
      columnIndex = columnIndex + 1
Next
Set dic = Nothing

'--------------------------------------------------------------------------------------------
''List out all testcasereferences from test suites
'--------------------------------------------------------------------------------------------

  Set rngFound = Sheets("TestPlans").Columns("B").Find(TestCycleResult, Cells(Rows.Count, "B"), xlValues, xlWhole)
  RowNumber = rngFound.Row
  Set rngFound = Nothing
  TESTSUITEID = Sheets("TestPlans").Cells(RowNumber, 1)
  TestPlanSuiteURL = Replace(TestPlanSuiteURL, "TESTSUITEID", TESTSUITEID)
  
   Call callAPI(TestPlanSuiteURL, 2, json$)

    Set dic = ParseJSON(json$)
    For Each key In dic.Keys
        CurrentColRowscount = Sheets("Dictionary").Range("B" & Rows.Count).End(xlUp).Row + 1
        Sheets("Dictionary").Cells(CurrentColRowscount, 2) = "Key: " & key & " <BRvalue:> " & Trim(dic(key))
    Next
    Set dic = Nothing
  
CurrentDicRowscount = Sheets("Dictionary").Range("B" & Rows.Count).End(xlUp).Row + 1
columnIndex = 1
For intDicount = 2 To CurrentDicRowscount - 2
      dictValue = Sheets("Dictionary").Cells(intDicount, 2)
      splitdictvalue = Split(dictValue, "<BRvalue:>")
      If intDicount = 2 Then
       CurrentTestCasescount = Sheets("TestSuites").Range("A" & Rows.Count).End(xlUp).Row + 1
      End If
      Sheets("TestSuites").Cells(CurrentTestCasescount, columnIndex) = Trim(splitdictvalue(1))
      If InStr(splitdictvalue(0), "lastUpdatedDate") Then
         CurrentTestCasescount = Sheets("TestSuites").Range("A" & Rows.Count).End(xlUp).Row + 1
         columnIndex = 0
      End If
columnIndex = columnIndex + 1
Next
  
    
  
''--------------------------------------------------------------------------------------------
'''List out all testcaseReferences from test suites
''--------------------------------------------------------------------------------------------
'
  CurrentTestSuitesRowscount = Sheets("TestSuites").Range("A" & Rows.Count).End(xlUp).Row
  
  

   For intTestsuitecount = 3 To CurrentTestSuitesRowscount

    TestcaseID = Sheets("TestSuites").Cells(intTestsuitecount, 1)
    ParentFolderName = Sheets("TestSuites").Cells(intTestsuitecount, 11)
    TestCaseName = Sheets("TestSuites").Cells(intTestsuitecount, 2)
    TestcaseURl = Sheets("TestSuites").Cells(intTestsuitecount, 3) & "/testcases"

    TestcaseCount = Sheets("TestSuites").Cells(intTestsuitecount, 14)

    If Trim(ParentFolderName) <> TestCycleResult Or TestcaseCount > 0 Then

            Call callAPI(TestcaseURl, 3, json$)
            Set dic = ParseJSON(json$)
            URLFOUND = False
            For Each key In dic.Keys
                CurrentColRowscount = Sheets("TestcaseLists").Range("A" & Rows.Count).End(xlUp).Row + 1
                If Right(key, 12) = "testCase.url" Then
                 Sheets("TestcaseLists").Cells(CurrentColRowscount, 1) = TestcaseID
                 Sheets("TestcaseLists").Cells(CurrentColRowscount, 2) = ParentFolderName
                 Sheets("TestcaseLists").Cells(CurrentColRowscount, 3) = TestCaseName
                 Sheets("TestcaseLists").Cells(CurrentColRowscount, 4) = TestcaseURl
                 Sheets("TestcaseLists").Cells(CurrentColRowscount, 5) = dic(key)
                  URLFOUND = True
                End If
            Next
            CurrentColRowscount = Sheets("TestcaseLists").Range("A" & Rows.Count).End(xlUp).Row + 1
            If URLFOUND = False Then
               Sheets("TestcaseLists").Cells(CurrentColRowscount, 1) = TestcaseID
               Sheets("TestcaseLists").Cells(CurrentColRowscount, 2) = ParentFolderName
               Sheets("TestcaseLists").Cells(CurrentColRowscount, 3) = TestCaseName
               Sheets("TestcaseLists").Cells(CurrentColRowscount, 4) = TestcaseURl
               Sheets("TestcaseLists").Cells(CurrentColRowscount, 5) = "Testcases Not Exist"
            End If
     End If

    Set dic = Nothing
   Next
'''------------------------------------------------------------------------------------------------------------
'
  
  
 '--------------------------------------------------------------------------------------------
''List out all testcaseTitles from testcaseLinks
'--------------------------------------------------------------------------------------------
  
  CurrentTestCaseListscount = Sheets("TestcaseLists").Range("A" & Rows.Count).End(xlUp).Row

   For intTestcasescount = 2 To CurrentTestCaseListscount
     TestcaseLinkUrl = Sheets("TestcaseLists").Cells(intTestcasescount, 5)
      If TestcaseLinkUrl <> "Testcases Not Exist" Then
        Call callAPI(TestcaseLinkUrl, 5, json$)
        Set dic = ParseJSON(json$)
        For Each key In dic.Keys
           If Trim(Right(key, 5)) = "Title" Then
             CurrentColRowscount = Sheets("TestcaseLists").Range("F" & Rows.Count).End(xlUp).Row + 1
             Sheets("TestcaseLists").Cells(CurrentColRowscount, 6) = Trim(dic(key))
             Exit For
           End If
         Next
         Else
           CurrentColRowscount = Sheets("TestcaseLists").Range("F" & Rows.Count).End(xlUp).Row + 1
           Sheets("TestcaseLists").Cells(CurrentColRowscount, 6) = "Tests Not Exists"
      End If
   Next

Set dic = Nothing

'''add testpointkeys to dictionary to fetch testcase ID and result------------------------------
'
  CurrentTestSuitesRowscount = Sheets("TestSuites").Range("A" & Rows.Count).End(xlUp).Row

   For intTestsuitecount = 2 To CurrentTestSuitesRowscount

    SuiteID = Sheets("TestSuites").Cells(intTestsuitecount, 1)
    TestSuiteName = Sheets("TestSuites").Cells(intTestsuitecount, 2)
    PLANID = Sheets("TestSuites").Cells(intTestsuitecount, 7)

    TestPlanPointsLinkPLANID = Replace(TestcasePointsLink, "PLANID", PLANID)
    TestPlanPointsLink = Replace(TestPlanPointsLinkPLANID, "SUITEID", SuiteID)
   
    Call callAPI(TestPlanPointsLink, 4, json$)

    Set dic = ParseJSON(json$)
    For Each key In dic.Keys
        CurrentColRowscount = Sheets("Dictionary").Range("C" & Rows.Count).End(xlUp).Row + 1
        If Trim(Left(key, 9)) = "obj.value" Then
            splitkey = Split(key, ".")
            keyName = splitkey(UBound(splitkey))
            Else
            keyName = key
        End If
        
        Sheets("Dictionary").Cells(CurrentColRowscount, 3) = Trim(keyName)
        Sheets("Dictionary").Cells(CurrentColRowscount, 4) = Trim(dic(key))
        Sheets("Dictionary").Cells(CurrentColRowscount, 5) = Trim(TestSuiteName)
    Next
    Set dic = Nothing

  Next


''Replace all which tests Not exists tests

 Sheets("TestcaseLists").UsedRange.Replace What:="Tests Not Exists", Replacement:=""

''Assign the TetscaseID and Result to TestcaseLists


TestcaseCount = Sheets("TestcaseLists").Range("F" & Rows.Count).End(xlUp).Row + 1

   For intTestCaseCount = 2 To TestcaseCount
   
              TestCaseName = Sheets("TestcaseLists").Cells(intTestCaseCount, 6)
              FolderName = Sheets("TestcaseLists").Cells(intTestCaseCount, 3)
              TestcaseID = ""
              lastResultState = ""
              currentState = ""
              
              If TestCaseName <> "" Then
              
              
                 'Firstmatch Folder
                 
                 Set rngFoundFolder = Sheets("Dictionary").Columns("E").Find(Trim(FolderName), Cells(Rows.Count, "E"), xlValues, xlWhole)
                 rowFoundFolder = rngFoundFolder.Row
                 AftercellRange = "D" & rowFoundFolder - 1
                 
                 Set cell = Sheets("Dictionary").Columns("D").Find(Trim(TestCaseName), After:=Range(AftercellRange))
                 rowFound = cell.Row
        
                 
                 TestcaseID = Sheets("Dictionary").Cells(rowFound - 1, "C")
                 If Trim(TestcaseID) = "id" Then
                    TestcaseID = Sheets("Dictionary").Cells(rowFound - 1, "D")
                    Else
                    TestcaseID = "NotFound"
                 End If
                     
                 currentState = Sheets("Dictionary").Cells(rowFound - 3, "C")
                 If Trim(currentState) = "state" Then
                    currentState = Sheets("Dictionary").Cells(rowFound - 3, "D")
                    Else
                    currentState = "NotFound"
                 End If
                 
                 
                lastResultState = Sheets("Dictionary").Cells(rowFound - 4, "C")
                 If Trim(lastResultState) = "outcome" Then
                    lastResultState = Sheets("Dictionary").Cells(rowFound - 4, "D")
                    Else
                    lastResultState = "NotFound"
                 End If
                 
                 
                  Set rowFound = Nothing
                 ''Writeback result in
                   Sheets("TestcaseLists").Cells(intTestCaseCount, 7) = TestcaseID
                   Sheets("TestcaseLists").Cells(intTestCaseCount, 8) = lastResultState
                   Sheets("TestcaseLists").Cells(intTestCaseCount, 9) = currentState
              
              End If
   
   Next
   
   
 
''Defects lists

Call Module3.fetchdefects

''create PivotsheetData

Call Module4.createPivotData

'Refresh all pivots

Sheets("macro").PivotTables("PivotTable2").PivotCache.Refresh


Dim PivotWs As Worksheet
Set PivotWs = ActiveWorkbook.Worksheets("Macro")
    
'select pivot table and refresh
PivotWs.Select
Set PTSuppBase = ActiveSheet.PivotTables("PivotTable2")
PTSuppBase.RefreshTable

MinutesElapsed = Format((Timer - startTime) / 86400, "hh:mm:ss")
MsgBox "Completed in " & MinutesElapsed


End Sub

Function EncodeBase64(text As String)

  Dim arrData() As Byte
  arrData = StrConv(text, vbFromUnicode)

  Dim objXML As MSXML2.DOMDocument
  Dim objNode As MSXML2.IXMLDOMElement

  Set objXML = New MSXML2.DOMDocument
  Set objNode = objXML.createElement("b64")

  objNode.DataType = "bin.base64"
  objNode.nodeTypedValue = arrData
  EncodeBase64 = Application.Clean(objNode.text)

  Set objNode = Nothing
  Set objXML = Nothing
End Function


Public Function ReturnResult(ByVal responseText, ByVal key1, ByVal SheetName, ByRef strReturnResult)
    Dim objRegexp As RegExp
    Set objRegexp = New RegExp
    objRegexp.IgnoreCase = True
    objRegexp.Global = True

    strcheckexpression = "" & key1 & """" & ":([\s\S]+?),"
    
    objRegexp.Pattern = strcheckexpression
    Set matches = objRegexp.Execute(responseText)
    Index = 0
    For Each match In matches
        strReturnValue = match.Value
        strReturnValue = Replace(strReturnValue, key1, "")
        strReturnValue = Replace(strReturnValue, """", "")
        strReturnValue = Replace(strReturnValue, ":", "")
        strReturnValue = Replace(strReturnValue, ",", "")
        strReturnValue = Replace(strReturnValue, " ", "")
        strReturnValue = Replace(strReturnValue, "}", "")
        If key1 = "id" Then
         CurrentColRowscount = Sheets(SheetName).Range("A" & Rows.Count).End(xlUp).Row + 1
         Sheets(SheetName).Cells(CurrentColRowscount, 1) = strReturnValue
        End If
        If key1 = "name" And SheetName <> "TestSuiteKeyValue" Then
        CurrentColRowscount = Sheets(SheetName).Range("B" & Rows.Count).End(xlUp).Row + 1
         If InStr(strReturnValue, "Castle-MainProject") = False Then
            If Index > 0 Then
              Sheets(SheetName).Cells(CurrentColRowscount + 3, 2) = strReturnValue
            Else
              Sheets(SheetName).Cells(CurrentColRowscount, 2) = strReturnValue
            End If
            Index = Index + 1
         End If
        End If
        
        If (key1 = "name" And SheetName = "TestSuiteKeyValue") Then
        CurrentColRowscount = Sheets(SheetName).Range("B" & Rows.Count).End(xlUp).Row + 1
         If InStr(Trim(strReturnValue), "NFTCycle") = False Then
           If InStr(Trim(strReturnValue), "Castle-MainProject") = False Then
              If InStr(Trim(strReturnValue), "Windows") = False Then
                If Index > 0 Then
                 Sheets(SheetName).Cells(CurrentColRowscount + 3, 2) = strReturnValue
               End If
             End If
           End If
         End If
         If Index = 0 Then
            Sheets(SheetName).Cells(CurrentColRowscount, 2) = strReturnValue
         End If
          Index = Index + 1
        End If
       
        If InStr(strReturnValue, "succeeded") > 0 Then
           strFound = True
           Exit For
        End If
        If InStr(strReturnValue, "failed") > 0 Then
           strFound = True
           Exit For
        End If
    Next
    strReturnResult = strReturnValue

    Set objRegexp = Nothing
End Function

Sub callAPI(ByVal URL, ByVal rowno, ByRef responseText)
    UserName = "xxx"
    token = xxxx"
    sStatus = ""
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    objHTTP.Open "GET", URL, False
    objHTTP.setRequestHeader "Content-Type", "application/json"
    objHTTP.setRequestHeader "Accept", "application/json"
    objHTTP.setRequestHeader "Authorization", "Basic " & EncodeBase64(UserName & ":" & token)
    objHTTP.send
    responseText = objHTTP.responseText
    'Sheets("JsonResponse").Cells(rowno, 1) = responseText
    sStatus = objHTTP.Status & " | " & objHTTP.statusText
    Set objHTTP = Nothing
End Sub


'-------------------------------------------------------------------
' VBA JSON Parser
'-------------------------------------------------------------------

Function ParseJSON(json$, Optional key$ = "obj") As Object
    p = 1
    token = Tokenize(json)
    Set dic = CreateObject("Scripting.Dictionary")
    If token(p) = "{" Then ParseObj key Else ParseArr key
    Set ParseJSON = dic
     Set dic = Nothing
End Function
Function ParseObj(key$)
    Do: p = p + 1
        Select Case token(p)
            Case "]"
            Case "[":  ParseArr key
            Case "{"
                       If token(p + 1) = "}" Then
                           p = p + 1
                           dic.Add key, "null"
                       Else
                           ParseObj key
                       End If
                
            Case "}":  key = ReducePath(key): Exit Do
            Case ":":  key = key & "." & token(p - 1)
            Case ",":  key = ReducePath(key)
            Case Else: If token(p + 1) <> ":" Then dic.Add key, token(p)
        End Select
    Loop
End Function
Function ParseArr(key$)
    Dim e&
    Do: p = p + 1
        Select Case token(p)
            Case "}"
            Case "{":  ParseObj key & ArrayID(e)
            Case "[":  ParseArr key
            Case "]":  Exit Do
            Case ":":  key = key & ArrayID(e)
            Case ",":  e = e + 1
            Case Else: dic.Add key & ArrayID(e), token(p)
        End Select
    Loop
End Function
'-------------------------------------------------------------------
' Support Functions
'-------------------------------------------------------------------
Function Tokenize(s$)
    Const Pattern = """(([^""\\]|\\.)*)""|[+\-]?(?:0|[1-9]\d*)(?:\.\d*)?(?:[eE][+\-]?\d+)?|\w+|[^\s""']+?"
    Tokenize = RExtract(s, Pattern, True)
End Function
Function RExtract(s$, Pattern, Optional bGroup1Bias As Boolean, Optional bGlobal As Boolean = True)
  Dim c&, m, n, v
  With CreateObject("vbscript.regexp")
    .Global = bGlobal
    .MultiLine = False
    .IgnoreCase = True
    .Pattern = Pattern
    If .Test(s) Then
      Set m = .Execute(s)
      ReDim v(1 To m.Count)
      For Each n In m
        c = c + 1
        v(c) = n.Value
        If bGroup1Bias Then If Len(n.SubMatches(0)) Or n.Value = """""" Then v(c) = n.SubMatches(0)
      Next
    End If
  End With
  RExtract = v
End Function
Function ArrayID$(e)
    ArrayID = "(" & e & ")"
End Function
Function ReducePath$(key$)
    If InStr(key, ".") Then ReducePath = Left(key, InStrRev(key, ".") - 1) Else ReducePath = key
End Function
Function ListPaths(dic)
    Dim s$, v
    For Each v In dic
        s = s & v & " --> " & dic(v) & vbLf
    Next
    Debug.Print s
End Function
Function GetFilteredValues(dic, match)
    Dim c&, i&, v, w
    v = dic.Keys
    ReDim w(1 To dic.Count)
    For i = 0 To UBound(v)
        If v(i) Like match Then
            c = c + 1
            w(c) = dic(v(i))
        End If
    Next
    ReDim Preserve w(1 To c)
    GetFilteredValues = w
End Function
Function GetFilteredTable(dic, cols)
    Dim c&, i&, j&, v, w, z
    v = dic.Keys
    z = GetFilteredValues(dic, cols(0))
    ReDim w(1 To UBound(z), 1 To UBound(cols) + 1)
    For j = 1 To UBound(cols) + 1
         z = GetFilteredValues(dic, cols(j - 1))
         For i = 1 To UBound(z)
            w(i, j) = z(i)
         Next
    Next
    GetFilteredTable = w
End Function
Function OpenTextFile$(f)
    With CreateObject("ADODB.Stream")
        .Charset = "utf-8"
        .Open
        .LoadFromFile f
        OpenTextFile = .ReadText
    End With
End Function













Dictionary:


''
' Dictionary v1.4.1
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

#Const UseScriptingDictionaryIfAvailable = True

#If Mac Or Not UseScriptingDictionaryIfAvailable Then

' dict_KeyValue 0: FormattedKey, 1: OriginalKey, 2: Value
Private dict_pKeyValues As Collection
Private dict_pKeys() As Variant
Private dict_pItems() As Variant
Private dict_pObjectKeys As Collection
Private dict_pCompareMode As CompareMethod

#Else

Private dict_pDictionary As Object

#End If

' --------------------------------------------- '
' Types
' --------------------------------------------- '

Public Enum CompareMethod
    BinaryCompare = VBA.vbBinaryCompare
    TextCompare = VBA.vbTextCompare
    DatabaseCompare = VBA.vbDatabaseCompare
End Enum

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public Property Get CompareMode() As CompareMethod
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    CompareMode = dict_pCompareMode
#Else
    CompareMode = dict_pDictionary.CompareMode
#End If
End Property
Public Property Let CompareMode(Value As CompareMethod)
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    If Me.Count > 0 Then
        ' Can't change CompareMode for Dictionary that contains data
        ' http://msdn.microsoft.com/en-us/library/office/gg278481(v=office.15).aspx
        Err.Raise 5 ' Invalid procedure call or argument
    End If

    dict_pCompareMode = Value
#Else
    dict_pDictionary.CompareMode = Value
#End If
End Property

Public Property Get Count() As Long
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Count = dict_pKeyValues.Count
#Else
    Count = dict_pDictionary.Count
#End If
End Property

Public Property Get Item(key As Variant) As Variant
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Dim dict_KeyValue As Variant
    dict_KeyValue = dict_GetKeyValue(key)

    If Not IsEmpty(dict_KeyValue) Then
        If VBA.IsObject(dict_KeyValue(2)) Then
            Set Item = dict_KeyValue(2)
        Else
            Item = dict_KeyValue(2)
        End If
    Else
        ' Not found -> Returns Empty
    End If
#Else
    If VBA.IsObject(dict_pDictionary.Item(key)) Then
        Set Item = dict_pDictionary.Item(key)
    Else
        Item = dict_pDictionary.Item(key)
    End If
#End If
End Property
Public Property Let Item(key As Variant, Value As Variant)
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    If Me.Exists(key) Then
        dict_ReplaceKeyValue dict_GetKeyValue(key), key, Value
    Else
        dict_AddKeyValue key, Value
    End If
#Else
    dict_pDictionary.Item(key) = Value
#End If
End Property
Public Property Set Item(key As Variant, Value As Variant)
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    If Me.Exists(key) Then
        dict_ReplaceKeyValue dict_GetKeyValue(key), key, Value
    Else
        dict_AddKeyValue key, Value
    End If
#Else
    Set dict_pDictionary.Item(key) = Value
#End If
End Property

Public Property Let key(Previous As Variant, Updated As Variant)
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Dim dict_KeyValue As Variant
    dict_KeyValue = dict_GetKeyValue(Previous)

    If Not VBA.IsEmpty(dict_KeyValue) Then
        dict_ReplaceKeyValue dict_KeyValue, Updated, dict_KeyValue(2)
    End If
#Else
    dict_pDictionary.key(Previous) = Updated
#End If
End Property

' ============================================= '
' Public Methods
' ============================================= '

''
' Add an item with the given key
'
' @param {Variant} Key
' @param {Variant} Item
' --------------------------------------------- '
Public Sub Add(key As Variant, Item As Variant)
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    If Not Me.Exists(key) Then
        dict_AddKeyValue key, Item
    Else
        ' This key is already associated with an element of this collection
        Err.Raise 457
    End If
#Else
    dict_pDictionary.Add key, Item
#End If
End Sub

''
' Check if an item exists for the given key
'
' @param {Variant} Key
' @return {Boolean}
' --------------------------------------------- '
Public Function Exists(key As Variant) As Boolean
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Exists = Not IsEmpty(dict_GetKeyValue(key))
#Else
    Exists = dict_pDictionary.Exists(key)
#End If
End Function

''
' Get an array of all items
'
' @return {Variant}
' --------------------------------------------- '
Public Function Items() As Variant
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    If Me.Count > 0 Then
        Items = dict_pItems
    Else
        ' Split("") creates initialized empty array that matches Dictionary Keys and Items
        Items = VBA.Split("")
    End If
#Else
    Items = dict_pDictionary.Items
#End If
End Function

''
' Get an array of all keys
'
' @return {Variant}
' --------------------------------------------- '
Public Function Keys() As Variant
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    If Me.Count > 0 Then
        Keys = dict_pKeys
    Else
        ' Split("") creates initialized empty array that matches Dictionary Keys and Items
        Keys = VBA.Split("")
    End If
#Else
    Keys = dict_pDictionary.Keys
#End If
End Function

''
' Remove an item for the given key
'
' @param {Variant} Key
' --------------------------------------------- '
Public Sub Remove(key As Variant)
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Dim dict_KeyValue As Variant
    dict_KeyValue = dict_GetKeyValue(key)

    If Not VBA.IsEmpty(dict_KeyValue) Then
        dict_RemoveKeyValue dict_KeyValue
    Else
        ' Application-defined or object-defined error
        Err.Raise 32811
    End If
#Else
    dict_pDictionary.Remove key
#End If
End Sub

''
' Remove all items
' --------------------------------------------- '
Public Sub RemoveAll()
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Set dict_pKeyValues = New Collection

    Erase dict_pKeys
    Erase dict_pItems
#Else
    dict_pDictionary.RemoveAll
#End If
End Sub

' ============================================= '
' Private Functions
' ============================================= '

#If Mac Or Not UseScriptingDictionaryIfAvailable Then

Private Function dict_GetKeyValue(dict_Key As Variant) As Variant
    On Error Resume Next
    dict_GetKeyValue = dict_pKeyValues(dict_GetFormattedKey(dict_Key))
    Err.Clear
End Function

Private Sub dict_AddKeyValue(dict_Key As Variant, dict_Value As Variant, Optional dict_Index As Long = -1)
    If Me.Count = 0 Then
        ReDim dict_pKeys(0 To 0)
        ReDim dict_pItems(0 To 0)
    Else
        ReDim Preserve dict_pKeys(0 To UBound(dict_pKeys) + 1)
        ReDim Preserve dict_pItems(0 To UBound(dict_pItems) + 1)
    End If

    Dim dict_FormattedKey As String
    dict_FormattedKey = dict_GetFormattedKey(dict_Key)

    If dict_Index >= 0 And dict_Index < dict_pKeyValues.Count Then
        ' Shift keys/items after + including index into empty last slot
        Dim dict_i As Long
        For dict_i = UBound(dict_pKeys) To dict_Index + 1 Step -1
            dict_pKeys(dict_i) = dict_pKeys(dict_i - 1)
            If VBA.IsObject(dict_pItems(dict_i - 1)) Then
                Set dict_pItems(dict_i) = dict_pItems(dict_i - 1)
            Else
                dict_pItems(dict_i) = dict_pItems(dict_i - 1)
            End If
        Next dict_i

        ' Add key/item at index
        dict_pKeys(dict_Index) = dict_Key
        If VBA.IsObject(dict_Value) Then
            Set dict_pItems(dict_Index) = dict_Value
        Else
            dict_pItems(dict_Index) = dict_Value
        End If

        ' Add key-value at proper index
        dict_pKeyValues.Add Array(dict_FormattedKey, dict_Key, dict_Value), dict_FormattedKey, Before:=dict_Index + 1
    Else
        ' Add key-value as last item
        If VBA.IsObject(dict_Key) Then
            Set dict_pKeys(UBound(dict_pKeys)) = dict_Key
        Else
            dict_pKeys(UBound(dict_pKeys)) = dict_Key
        End If
        If VBA.IsObject(dict_Value) Then
            Set dict_pItems(UBound(dict_pItems)) = dict_Value
        Else
            dict_pItems(UBound(dict_pItems)) = dict_Value
        End If

        dict_pKeyValues.Add Array(dict_FormattedKey, dict_Key, dict_Value), dict_FormattedKey
    End If
End Sub

Private Sub dict_ReplaceKeyValue(dict_KeyValue As Variant, dict_Key As Variant, dict_Value As Variant)
    Dim dict_Index As Long
    Dim dict_i As Integer

    dict_Index = dict_GetKeyIndex(dict_KeyValue(1))

    ' Remove existing dict_Value
    dict_RemoveKeyValue dict_KeyValue, dict_Index

    ' Add new dict_Key dict_Value back
    dict_AddKeyValue dict_Key, dict_Value, dict_Index
End Sub

Private Sub dict_RemoveKeyValue(dict_KeyValue As Variant, Optional ByVal dict_Index As Long = -1)
    Dim dict_i As Long
    If dict_Index = -1 Then
        dict_Index = dict_GetKeyIndex(dict_KeyValue(1))
    End If

    If dict_Index >= 0 And dict_Index <= UBound(dict_pKeys) Then
        ' Shift keys/items after index down
        For dict_i = dict_Index To UBound(dict_pKeys) - 1
            dict_pKeys(dict_i) = dict_pKeys(dict_i + 1)

            If VBA.IsObject(dict_pItems(dict_i + 1)) Then
                Set dict_pItems(dict_i) = dict_pItems(dict_i + 1)
            Else
                dict_pItems(dict_i) = dict_pItems(dict_i + 1)
            End If
        Next dict_i

        ' Resize keys/items to remove empty slot
        If UBound(dict_pKeys) = 0 Then
            Erase dict_pKeys
            Erase dict_pItems
        Else
            ReDim Preserve dict_pKeys(0 To UBound(dict_pKeys) - 1)
            ReDim Preserve dict_pItems(0 To UBound(dict_pItems) - 1)
        End If
    End If

    dict_pKeyValues.Remove dict_KeyValue(0)
    dict_RemoveObjectKey dict_KeyValue(1)
End Sub

Private Function dict_GetFormattedKey(dict_Key As Variant) As String
    If VBA.IsObject(dict_Key) Then
        dict_GetFormattedKey = dict_GetObjectKey(dict_Key)
    ElseIf VarType(dict_Key) = VBA.vbBoolean Then
        dict_GetFormattedKey = IIf(dict_Key, "-1__-1", "0__0")
    ElseIf VarType(dict_Key) = VBA.vbString Then
        dict_GetFormattedKey = dict_Key

        If Me.CompareMode = CompareMethod.BinaryCompare Then
            ' Collection does not have method of setting key comparison
            ' So case-sensitive keys aren't supported by default
            ' -> Approach: Append lowercase characters to original key
            '    AbC -> AbC___b_, abc -> abc__abc, ABC -> ABC_____
            Dim dict_Lowercase As String
            dict_Lowercase = ""

            Dim dict_i As Integer
            Dim dict_Char As String
            Dim dict_Ascii As Integer
            For dict_i = 1 To VBA.Len(dict_GetFormattedKey)
                dict_Char = VBA.Mid$(dict_GetFormattedKey, dict_i, 1)
                dict_Ascii = VBA.Asc(dict_Char)
                If dict_Ascii >= 97 And dict_Ascii <= 122 Then
                    dict_Lowercase = dict_Lowercase & dict_Char
                Else
                    dict_Lowercase = dict_Lowercase & "_"
                End If
            Next dict_i

            If dict_Lowercase <> "" Then
                dict_GetFormattedKey = dict_GetFormattedKey & "__" & dict_Lowercase
            End If
        End If
    Else
        ' For numbers, add duplicate to distinguish from strings
        ' ->  123  -> "123__123"
        '    "123" -> "123"
        dict_GetFormattedKey = VBA.CStr(dict_Key) & "__" & CStr(dict_Key)
    End If
End Function

Private Function dict_GetObjectKey(dict_ObjKey As Variant) As String
    Dim dict_i As Integer
    For dict_i = 1 To dict_pObjectKeys.Count
        If dict_pObjectKeys.Item(dict_i) Is dict_ObjKey Then
            dict_GetObjectKey = "__object__" & dict_i
            Exit Function
        End If
    Next dict_i

    dict_pObjectKeys.Add dict_ObjKey
    dict_GetObjectKey = "__object__" & dict_pObjectKeys.Count
End Function

Private Sub dict_RemoveObjectKey(dict_ObjKey As Variant)
    Dim dict_i As Integer
    For dict_i = 1 To dict_pObjectKeys.Count
        If dict_pObjectKeys.Item(dict_i) Is dict_ObjKey Then
            dict_pObjectKeys.Remove dict_i
            Exit Sub
        End If
    Next dict_i
End Sub

Private Function dict_GetKeyIndex(dict_Key As Variant) As Long
    Dim dict_i As Long
    For dict_i = 0 To UBound(dict_pKeys)
        If VBA.IsObject(dict_pKeys(dict_i)) And VBA.IsObject(dict_Key) Then
            If dict_pKeys(dict_i) Is dict_Key Then
                dict_GetKeyIndex = dict_i
                Exit For
            End If
        ElseIf VBA.IsObject(dict_pKeys(dict_i)) Or VBA.IsObject(dict_Key) Then
            ' Both need to be objects to check equality, skip
        ElseIf dict_pKeys(dict_i) = dict_Key Then
            dict_GetKeyIndex = dict_i
            Exit For
        End If
    Next dict_i
End Function

#End If

Private Sub Class_Initialize()
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Set dict_pKeyValues = New Collection

    Erase dict_pKeys
    Erase dict_pItems
    Set dict_pObjectKeys = New Collection
#Else
    Set dict_pDictionary = CreateObject("Scripting.Dictionary")
#End If
End Sub

Private Sub Class_Terminate()
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Set dict_pKeyValues = Nothing
    Set dict_pObjectKeys = Nothing
#Else
    Set dict_pDictionary = Nothing
#End If
End Sub
Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    objHTTP.Open "GET", Url, False
    objHTTP.setRequestHeader "Content-Type", "application/json"
    objHTTP.setRequestHeader "Accept", "application/json"
    objHTTP.setRequestHeader "Authorization", "Basic " & EncodeBase64(UserName & ":" & token)
    objHTTP.send
    json$ = objHTTP.responseText

Dim http As Object, html As New HTMLDocument, topics As Object, titleElem As Object, detailsElem As Object, topic As HTMLHtmlElement
Dim i As Integer
Set request = CreateObject("MSXML2.ServerXMLHTTP")
    
    Dim converter As New ADODB.stream
    

    ' fetch page
    request.Open "GET", Url, False
    request.setRequestHeader "Authorization", "Basic " & EncodeBase64(UserName & ":" & token)
    request.send

    ' write raw bytes to the stream
    converter.Open
    converter.Type = adTypeBinary
    converter.Write request.responseBody
    
    converter.SaveToFile "location", 2

    ' read text characters from the stream
    converter.Position = 0
    converter.Type = adTypeText
    converter.Charset = "Windows-1251"

    ' set return value, close the stream
    GetHTML = converter.ReadText
    converter.Close



[1:34 PM] Ramu Gogurla
Function drawborder(ByVal strRange)
 
    Dim rRng As Range
 
    Set rRng = sheets("Generate Report").Range(strRange)

    rRng.WrapText = True
 
    'Clear existing

    rRng.Borders.LineStyle = xlNone
 
    'Apply new borders

    rRng.BorderAround xlContinuous

    rRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous

    rRng.Borders(xlInsideVertical).LineStyle = xlContinuous

    rRng.EntireColumn.AutoFit

    rRng.EntireRow.AutoFit


 Dim http As Object, html As New HTMLDocument, topics As Object, titleElem As Object, detailsElem As Object, topic As HTMLHtmlElement
 Dim i As Integer
 Dim oStream As Object
 Dim hasImage_column As Range

    Set http = CreateObject("MSXML2.ServerXMLHTTP")
    url = "https://xxe.png"
    http.Open "GET", "url ", False
    http.setRequestHeader "Authorization", "Basic " & EncodeBase64(UserName & ":" & token)
    http.send
    If http.Status = 200 Then
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write http.responseBody
        oStream.SaveToFile "c:\zz\test1.jpg", 2
        oStream.Close
    End If
    
    Set image_column = Worksheets(1).Range("A1:H22")
    
    
    Set shape = Sheets("Images").Shapes.AddPicture(locationToSave, msoFalse, msoTrue, 0, 0, 100, 100)
                
        With shape
            .Left = image_column.Cells(1).Left
            .Top = image_column.Cells(1).Top
            .Height = 500
            .Width = 500
            'image_column.Cells(1).EntireRow.RowHeight = .Height
        End With


Sub updateJsonRunTime()
    Dim rng As Range, Items As New Collection, myitem As New Dictionary, i As Integer, cell As Variant, myfile As String
    Dim FSO As New FileSystemObject
    Dim buss As String
    Dim JsonTS As TextStream
    'Set rng = Sheets("json").Cells(1, 1)
    Set JsonTS = FSO.OpenTextFile("xxx\json.json", ForReading)
    JsonText = JsonTS.ReadAll
    JsonTS.Close
     Dim JSON As Object
    Set JSON = ParseJSON(JsonText)
    'Debug.Print JsonText
    'Debug.Print json("root").Count
    
     Debug.Print JSON("data")("site")
    JSON("data")("site") = "9999"
    myfile = "xxx\output.json"
    Open myfile For Output As #1
    Print #1, JsonConverter.ConvertToJson(JSON, Whitespace:=2)
    Close #1
End Sub

    Set rRng = Nothing
 
End Function
[1:35 PM] Ramu Gogurla
generateReport.Range("B6:N50000").Interior.ColorIndex = 0
[1:35 PM] Ramu Gogurla
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual
ActiveSheet.DisplayPageBreaks = False

Sheets("xx").Cells(ResultupdateRowstart, 12) = "Pass " & sStatus
          Sheets("xx").Cells(ResultupdateRowstart, 12).Interior.ColorIndex = 4
          Else
          Sheets("x").Cells(ResultupdateRowstart, 12) = "Fail" & sStatus
          Sheets("xx").Cells(ResultupdateRowstart, 12).Interior.ColorIndex = 3

PivotWs.Select
Set PTSuppBase = ActiveSheet.PivotTables("PivotTable1")
PTSuppBase.RefreshTable
With PTSuppBase
    .DataBodyRange.HorizontalAlignment = xlLeft
    '.ColumnRange.HorizontalAlignment = xlLeft
    .TableStyle2 = "PivotStyleMedium20"
    .DataBodyRange.WrapText = True
    .TableRange1.Columns.AutoFit
End With
