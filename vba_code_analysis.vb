Option Explicit

'==================== Helper Functions ====================

Function GetStatusMessage(statusCode As Integer) As String
    Select Case statusCode
        Case 200: GetStatusMessage = "Success (200)"
        Case 201: GetStatusMessage = "Created (201)"
        Case 202: GetStatusMessage = "Accepted (202)"
        Case 204: GetStatusMessage = "No Content (204)"
        Case 300: GetStatusMessage = "Ambiguous Request (300)"
        Case 400: GetStatusMessage = "Bad Request (400)"
        Case 401: GetStatusMessage = "Unauthorized (401)"
        Case 403: GetStatusMessage = "Forbidden (403)"
        Case 404: GetStatusMessage = "Not Found (404)"
        Case 406: GetStatusMessage = "Not Acceptable (406)"
        Case 409: GetStatusMessage = "Conflict (409)"
        Case 415: GetStatusMessage = "Unsupported Media Type (415)"
        Case 422: GetStatusMessage = "Unprocessable Entity (422)"
        Case 429: GetStatusMessage = "Too Many Requests (429)"
        Case 500: GetStatusMessage = "Server Error (500)"
        Case Else: GetStatusMessage = "Unexpected status code (" & statusCode & ")"
    End Select
End Function
Function TokenizeCondition(conditionText As String) As Variant
    Dim temp As String
    temp = Replace(conditionText, "(", " ( ")
    temp = Replace(temp, ")", " ) ")
    temp = Replace(temp, " and ", " AND ")
    temp = Replace(temp, " or ", " OR ")
    TokenizeCondition = Split(temp)
End Function
Function BuildRuleSet(rules As Collection, mainRuleName As String) As Object
    Dim ruleSet As Object
    Set ruleSet = CreateObject("Scripting.Dictionary")
    ruleSet("Messages") = Null
    Set ruleSet("Data") = CreateObject("Scripting.Dictionary")
    ruleSet("Data")("Create") = CollectionToArray(rules)
    ruleSet("Data")("ETag") = Format(Now, "yyyy-mm-dd\THH:MM:SS")
    ruleSet("Data")("SysId") = CreateGUID()
    ruleSet("Data")("Name") = mainRuleName
    ruleSet("Data")("Description") = ""
    ruleSet("Data")("RuleType") = "Clause Assembly"
    ruleSet("Data")("IsLocked") = False
    ruleSet("Data")("LockedBy") = ""
    ruleSet("Data")("IsRuleSetActive") = True
    ruleSet("HasMoreData") = False
    ruleSet("PagingData") = Null
    Set BuildRuleSet = ruleSet
End Function

Function GenerateClauseAssemblyJSONUnified(wb As Workbook, ByRef mainRuleName As String, ByRef subRuleCount As Integer) As String
    Dim ws As Worksheet, rules As Collection
    Set rules = New Collection
    subRuleCount = 0
    mainRuleName = ""

    For Each ws In wb.Worksheets
        If ws.Name <> "Lookups" Then
            Dim sheetRules As Collection
            Set sheetRules = ProcessSheet(ws, mainRuleName)
            Dim ruleItem As Variant
            For Each ruleItem In sheetRules
                rules.Add ruleItem
                subRuleCount = subRuleCount + 1
            Next ruleItem
        End If
    Next ws

    Dim ruleSet As Object
    Set ruleSet = BuildRuleSet(rules, mainRuleName)
    GenerateClauseAssemblyJSONUnified = JsonConverter.ConvertToJson(ruleSet, Whitespace:=2)
End Function

Function ProcessSheet(ws As Worksheet, ByRef mainRuleName As String) As Collection
    Dim i As Long
    Dim rules As New Collection
    Dim ruleName As String, contractType As String
    Dim CTNAmeFound As Boolean, mainRuleFound As Boolean

    For i = 1 To 10
        If Trim(ws.Cells(i, 1).value) = "Contract Type Name" Then
            contractType = ws.Cells(i, 2).value
            CTNAmeFound = True
        End If
        If Trim(ws.Cells(i, 1).value) = "Main Rule Name" Or Trim(ws.Cells(i, 1).value) = "Rule Name" Then
            ruleName = ws.Cells(i, 2).value
            mainRuleFound = True
        End If
        If CTNAmeFound And mainRuleFound Then Exit For
    Next i
    mainRuleName = ruleName

    i = 1
    Do While i <= ws.UsedRange.Rows.Count
        If Trim(ws.Cells(i, 2).value) = "Sub-Rule Name" Or Trim(ws.Cells(i, 2).value) = "Sub-Rule" Then
            Dim subRuleName As String, conditionText As String
            Dim ruleDict As Object, expressions As Object, actions As Collection

            subRuleName = EscapeJSON(ws.Cells(i, 3).value)
            i = i + 1

            If Trim(ws.Cells(i, 2).value) = "Condition" Then
                conditionText = ws.Cells(i, 3).value
                Set expressions = BuildExpressionsAdvanced(conditionText)
                i = i + 1
            Else
                Set expressions = CreateObject("Scripting.Dictionary")
            End If

            If Trim(ws.Cells(i, 2).value) = "Select action" Then i = i + 1

            Set actions = New Collection
StartLoop:
            Do While i <= ws.UsedRange.Rows.Count
                If Trim(ws.Cells(i, 2).value) = "" Then
                    i = i + 1
                    GoTo StartLoop
                End If
                If Trim(ws.Cells(i, 2).value) = "Sub-Rule Name" Or Trim(ws.Cells(i, 2).value) = "Sub-Rule" Then Exit Do
                actions.Add BuildAction(ws, i)
                i = i + 1
            Loop

            Set ruleDict = CreateObject("Scripting.Dictionary")
            ruleDict("ConditionGroups") = Array(expressions)
            ruleDict("Actions") = CollectionToArray(actions)
            ruleDict("Id") = 1
            ruleDict("Name") = subRuleName
            ruleDict("StopProcessingMoreRules") = False
            ruleDict("IsActive") = True
            rules.Add ruleDict
        End If
        i = i + 1
    Loop

    Set ProcessSheet = rules
End Function



Function ParseConditionGroup(conditionText As String) As Object
    Dim group As Object
    Set group = CreateObject("Scripting.Dictionary")

    Dim andParts() As String, orParts() As String
    Dim subGroups As Collection
    Set subGroups = New Collection
    Dim part As Variant

    If InStr(conditionText, " or ") > 0 Then
        orParts = Split(conditionText, " or ")
        group("LogicalOperator") = "OR"
        For Each part In orParts
            subGroups.Add ParseSimpleExpression(part)
        Next part
    ElseIf InStr(conditionText, " and ") > 0 Then
        andParts = Split(conditionText, " and ")
        group("LogicalOperator") = "AND"
        For Each part In andParts
            If Left(Trim(part), 1) = "(" And Right(Trim(part), 1) = ")" Then
                subGroups.Add ParseConditionGroup(Mid(Trim(part), 2, Len(Trim(part)) - 2))
            Else
                subGroups.Add ParseSimpleExpression(part)
            End If
        Next part
    Else
        group("LogicalOperator") = "AND"
        subGroups.Add ParseSimpleExpression(conditionText)
    End If

    group("Expressions") = Array()
    group("ConditionGroups") = CollectionToArray(subGroups)
    Set ParseConditionGroup = group
End Function



Function BuildAction(ws As Worksheet, i As Long) As Object
    Dim actionType As String, clauseName As String, placeholderClause As String
    Dim clauseType As String, subClause As String, sequence As Variant
    Dim action As Object

    actionType = IIf(ws.Cells(i, 2).value Like "*Placeholder*", "SelectPlaceholderClause", "SelectClause")
    placeholderClause = EscapeJSON(ws.Cells(i, 3).value)
    clauseType = EscapeJSON(ws.Cells(i, 4).value)
    subClause = EscapeJSON(ws.Cells(i, 5).value)
    sequence = ws.Cells(i, 6).value
    If IsEmpty(sequence) Or sequence = "" Then sequence = Null
    clauseName = IIf(actionType = "SelectClause", subClause, "")

    Set action = CreateObject("Scripting.Dictionary")
    action("ActionType") = actionType
    action("ClauseName") = clauseName
    action("ConditionGroups") = Array()
    action("PlaceholderClause") = IIf(actionType = "SelectPlaceholderClause", placeholderClause, "")
    action("ClauseType") = clauseType
    action("SubClause") = IIf(actionType = "SelectPlaceholderClause", subClause, "")
    action("Sequence") = sequence

    Set BuildAction = action
End Function

Function EscapeJSON(text As Variant) As String
    If IsMissing(text) Or IsEmpty(text) Or IsNull(text) Then
        EscapeJSON = ""
    Else
        Dim temp As String
        temp = CStr(text)
        temp = Replace(temp, "", "\")
        temp = Replace(temp, """", """")
        temp = Replace(temp, vbCrLf, "\n")
        temp = Replace(temp, vbLf, "\n")
        EscapeJSON = temp
    End If
End Function

Function CleanText(ByVal inputText As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Pattern = "^\s+"
        .Global = True
    End With
    inputText = regex.Replace(inputText, "")
    With regex
        .Pattern = "\s+$"
        .Global = True
    End With
    inputText = regex.Replace(inputText, "")
    CleanText = inputText
End Function



Function CreateGUID() As String
    CreateGUID = "bbaa2d81-baaf-47f7-901c-1ccc5162b04c"
End Function

Sub openform()
    formRuleCreation.txtInstanceName.SetFocus
    formRuleCreation.Show
End Sub

Sub ProcessMultipleWorkbooks()
    Dim fDialog As FileDialog
    Dim selectedFile As String
    Dim instanceName As String
    Dim contractType As String
    Dim jwtToken As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim jsonBody As String
    Dim mainRuleName As String
    Dim subRuleCount As Integer

    instanceName = formRuleCreation.txtInstanceName.value
    contractType = "ICMStandaloneAgreement"
    jwtToken = formRuleCreation.txtToken.value

    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    fDialog.AllowMultiSelect = False
    fDialog.Title = "Select Excel File"
    fDialog.Filters.Clear
    fDialog.Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm"

    If fDialog.Show <> -1 Then Exit Sub

    selectedFile = fDialog.SelectedItems(1)
    Set wb = Workbooks.Open(selectedFile)

    For Each ws In wb.Worksheets
        If ws.Name <> "Lookups" Then
            jsonBody = GenerateClauseAssemblyJSONUnified(wb, mainRuleName, subRuleCount)
            Call SendClauseAssemblyRequest(jsonBody, instanceName, contractType, jwtToken, wb, ws.Name, mainRuleName, subRuleCount)
        End If
    Next ws

    wb.Save
    wb.Close False
End Sub

Sub SendClauseAssemblyRequest(jsonBody As String, instanceName As String, contractType As String, jwtToken As String, wb As Workbook, wsName As String, mainRuleName As String, subRuleCount As Integer)
    
 On Error GoTo httpError
 
    Dim http As Object
    Dim url As String
    Dim statusCode As Integer
    Dim responseText As String
    Dim logWb As Workbook
    Dim logSheet As Worksheet
    Dim nextRow As Long
    Dim logFilePath As String
    Dim startTime As Single, endTime As Single
    Dim executionTime As String, timeTaken As Double

    Set http = CreateObject("MSXML2.XMLHTTP")
    url = "https://" & instanceName & "-business-api.icertis.com/api/v1/ruleset/" & contractType & "/Clauseassembly"

    startTime = Timer
    With http
        .Open "POST", url, False
        .SetRequestHeader "ICMAuthToken", jwtToken
        .SetRequestHeader "Content-Type", "application/json"
        .Send jsonBody
        statusCode = .Status
        responseText = .responseText
    End With
    endTime = Timer
    timeTaken = Round(endTime - startTime, 2)
    executionTime = Format(Now, "yyyy-mm-dd HH:MM:SS")

    logFilePath = wb.Path & "\Execution_Log.xlsx"

    On Error Resume Next
    Set logWb = Workbooks.Open(logFilePath)
    If logWb Is Nothing Then
        Set logWb = Workbooks.Add
        Set logSheet = logWb.Sheets(1)
        logSheet.Name = "Execution Status"
        logSheet.Range("A1:H1").value = Array("CT", "Main Rule", "# of sub-rules", "Execution Status", "Error Description", "Execution Time", "Time Taken (s)", "JSON")
        logWb.SaveAs fileName:=logFilePath
    Else
        Set logSheet = logWb.Sheets("Execution Status")
    End If
    On Error GoTo 0

    nextRow = logSheet.Cells(logSheet.Rows.Count, "A").End(xlUp).Row + 1

    logSheet.Cells(nextRow, 1).value = wsName
    logSheet.Cells(nextRow, 2).value = mainRuleName
    logSheet.Cells(nextRow, 3).value = subRuleCount
    logSheet.Cells(nextRow, 4).value = GetStatusMessage(statusCode)
    logSheet.Cells(nextRow, 5).value = responseText
    logSheet.Cells(nextRow, 6).value = executionTime
    logSheet.Cells(nextRow, 7).value = timeTaken
    logSheet.Cells(nextRow, 8).value = jsonBody

    logWb.Save
    logWb.Close False
    
    Exit Sub
    
httpError:
    MsgBox Err.Description
    Debug.Print Err.Description
    
End Sub

'----

Function BuildExpressionsAdvanced(conditionText As String) As Object
   

    Dim tokens() As String
    tokens = TokenizeCondition(conditionText)
    Dim index As Long: index = 0
    Set BuildExpressionsAdvanced = ParseLogicalGroup(tokens, index)


End Function

Function ParseLogicalGroup(tokens() As String, ByRef index As Long) As Object
    Dim group As Object
    Set group = CreateObject("Scripting.Dictionary")
    group("LogicalOperator") = "AND"
    group("Expressions") = Array()
    group("ConditionGroups") = Array()

    Dim expressions As Collection
    Set expressions = New Collection

    Do While index <= UBound(tokens)
        Dim token As String
        token = Trim(tokens(index))

        If token = "(" Then
            index = index + 1
            Dim nestedGroup As Object
            Set nestedGroup = ParseLogicalGroup(tokens, index)
            group("ConditionGroups") = AppendToArray(group("ConditionGroups"), nestedGroup)
        ElseIf token = ")" Then
            Exit Do
        ElseIf token = "AND" Or token = "OR" Then
            group("LogicalOperator") = token
        Else
            Dim exprText As String
            exprText = token
            Do While index + 1 <= UBound(tokens)
                If tokens(index + 1) = "AND" Or tokens(index + 1) = "OR" Or tokens(index + 1) = "(" Or tokens(index + 1) = ")" Then Exit Do
                index = index + 1
                exprText = exprText & " " & tokens(index)
            Loop
            Dim expr As Object
            Set expr = ParseSimpleExpression(exprText)
            If Not expr Is Nothing Then expressions.Add expr
        End If
        index = index + 1
    Loop

    If expressions.Count > 0 Then
        group("Expressions") = CollectionToArray(expressions)
    End If

    ' If this group has no expressions and only one nested group, return that group directly
    If expressions.Count = 0 And UBound(group("ConditionGroups")) = 0 Then
        Set ParseLogicalGroup = group("ConditionGroups")(0)
    Else
        Set ParseLogicalGroup = group
    End If
End Function


Function ParseSimpleExpression(exprText As String) As Object
    Dim parts() As String
    parts = Split(exprText, "=")
    If UBound(parts) <> 1 Then Set ParseSimpleExpression = Nothing: Exit Function

    Dim expr As Object
    Set expr = CreateObject("Scripting.Dictionary")
    expr("Attribute") = CleanText(parts(0))
    expr("Operator") = "Eq"
    expr("Value") = CleanText(parts(1))
    expr("AggregationType") = Null
    expr("Selection") = Null

    Set ParseSimpleExpression = expr
End Function

Function AppendToArray(arr As Variant, item As Object) As Variant
    Dim result() As Variant
    Dim i As Long
    If IsEmpty(arr) Then
        ReDim result(0)
        Set result(0) = item
    Else
        ReDim result(0 To UBound(arr) + 1)
        For i = 0 To UBound(arr)
            Set result(i) = arr(i)
        Next i
        Set result(UBound(arr) + 1) = item
    End If
    AppendToArray = result
End Function

Function CollectionToArray(col As Collection) As Variant
    Dim arr() As Variant
    Dim i As Long
    If col.Count = 0 Then
        CollectionToArray = Array()
        Exit Function
    End If
    ReDim arr(0 To col.Count - 1)
    For i = 1 To col.Count
        Set arr(i - 1) = col(i)
    Next i
    CollectionToArray = arr
End Function