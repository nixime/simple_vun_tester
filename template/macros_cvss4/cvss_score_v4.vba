Option VBASupport 1
Option Explicit

' ==============================================================================
' CORE UTILITIES
' ==============================================================================
Public Function IsExcel() As Boolean
    Static cachedResult As Boolean
    Static hasRun As Boolean
    
    If Not hasRun Then
        On Error Resume Next
        Dim testApp As Object
        Set testApp = Application
        
        If Err.Number = 0 And Not testApp Is Nothing Then
            cachedResult = True
        Else
            cachedResult = False
        End If
        
        On Error GoTo 0
        hasRun = True
    End If
    
    IsExcel = cachedResult
End Function

Private Function CreateDictionary() As Object
    On Error Resume Next
    If IsExcel() Then
        Set CreateDictionary = CreateObject("Scripting.Dictionary")
    Else
        GlobalScope.BasicLibraries.loadLibrary("ScriptForge")
        Set CreateDictionary = CreateScriptService("Dictionary")
    End If
    On Error GoTo 0
End Function

Private Function UpdateDictValue(ByRef dict As Object, key As String, val As String)
    If IsExcel() Then
        dict(key) = val
    Else
        dict.ReplaceItem(key, val)
    End If
End Function

' ==============================================================================
' STATIC JSON DATA FUNCTIONS
' ==============================================================================
Function GetLevelValue(category As String, key As String) As Double
    ' Default return value if category or key is not found
    GetLevelValue = 0#
    
    Select Case UCase(category)
        Case "AV": Select Case UCase(key)
            Case "N": GetLevelValue = 0.0: Case "A": GetLevelValue = 0.1
            Case "L": GetLevelValue = 0.2: Case "P": GetLevelValue = 0.3
        End Select
        
        Case "PR": Select Case UCase(key)
            Case "N": GetLevelValue = 0.0: Case "L": GetLevelValue = 0.1
            Case "H": GetLevelValue = 0.2
        End Select
        
        Case "UI": Select Case UCase(key)
            Case "N": GetLevelValue = 0.0: Case "P": GetLevelValue = 0.1
            Case "A": GetLevelValue = 0.2
        End Select
        
        Case "AC": Select Case UCase(key)
            Case "L": GetLevelValue = 0.0: Case "H": GetLevelValue = 0.1
        End Select
        
        Case "AT": Select Case UCase(key)
            Case "N": GetLevelValue = 0.0: Case "P": GetLevelValue = 0.1
        End Select
        
        Case "VC", "VI", "VA"
            Select Case UCase(key)
                Case "H": GetLevelValue = 0.0: Case "L": GetLevelValue = 0.1
                Case "N": GetLevelValue = 0.2
            End Select
            
        Case "SC", "SI", "SA"
            Select Case UCase(key)
                Case "S": GetLevelValue = 0.0: Case "H": GetLevelValue = 0.1
                Case "L": GetLevelValue = 0.2: Case "N": GetLevelValue = 0.3
            End Select
            
        Case "CR", "IR", "AR"
            Select Case UCase(key)
                Case "H": GetLevelValue = 0.0: Case "M": GetLevelValue = 0.1
                Case "L": GetLevelValue = 0.2
            End Select
            
        Case "E": Select Case UCase(key)
            Case "U": GetLevelValue = 0.2: Case "P": GetLevelValue = 0.1
            Case "A": GetLevelValue = 0.0
        End Select
    End Select
End Function

Function GetMaxSeverity(category As String, key As Integer, Optional subkey As Integer) As Integer
    Select Case UCase(category)
        Case "EQ1"
            Select Case key
                Case 0: GetMaxSeverity = 1
                Case 1: GetMaxSeverity = 4
                Case 2: GetMaxSeverity = 5
            End Select
            
        Case "EQ2"
            Select Case key
                Case 0: GetMaxSeverity = 1
                Case 1: GetMaxSeverity = 2
            End Select

        Case "EQ3EQ6"
            Select Case key
                Case 0
                    Select Case subkey
                        Case 0: GetMaxSeverity = 7
                        Case 1: GetMaxSeverity = 6
                    End Select
                Case 1: GetMaxSeverity = 8
                Case 2: GetMaxSeverity = 10
            End Select

        Case "EQ4"
            Select Case key
                Case 0: GetMaxSeverity = 6
                Case 1: GetMaxSeverity = 5
                Case 2: GetMaxSeverity = 4
            End Select
    End Select
End Function

' ==============================================================================
' LOOKUP & PARSING LOGIC
' ==============================================================================
Private Function getSafeRangeSearch(rangeName As String, findString As String) As String
    Dim foundValue As String: foundValue = ""
    Dim targetRange As Object, r As Long, rowCount As Long, currentKey As String

    On Error Resume Next
    If Not IsExcel() Then
        Dim oDoc As Object
        Set oDoc = ThisComponent
        If oDoc.NamedRanges.hasByName(rangeName) Then
            Set targetRange = oDoc.NamedRanges.getByName(rangeName).getReferredCells()
            rowCount = targetRange.getRows().getCount()
            For r = 0 To rowCount - 1
                currentKey = Trim(targetRange.getCellByPosition(0, r).getString())
                If currentKey = findString Then
                    foundValue = CStr(targetRange.getCellByPosition(1, r).getValue())
                    Exit For
                End If
            Next r
        End If
    Else
        Set targetRange = Range(rangeName)
        rowCount = targetRange.Rows.Count
        For r = 1 To rowCount
            currentKey = Trim(CStr(targetRange.Cells(r, 1).Value))
            If currentKey = findString Then
                foundValue = CStr(targetRange.Cells(r, 2).Value)
                Exit For
            End If
        Next r
    End If
    getSafeRangeSearch = foundValue
End Function

Private Function FillDefaults(ByRef dict As Object)
    Dim keys As Variant, i As Integer
    keys = Array("AV", "AC", "AT", "PR", "UI", "VC", "VI", "VA", "SC", "SI", "SA", "E", _
                 "CR", "IR", "AR", "S", "AU", "R", "V", "RE", "U")
    For i = LBound(keys) To UBound(keys)
        dict.Add keys(i), "X"
    Next i
    
    UpdateDictValue dict, "AV", "N"
    UpdateDictValue dict, "AC", "L"
    UpdateDictValue dict, "AT", "N"
    UpdateDictValue dict, "PR", "N"
    UpdateDictValue dict, "UI", "N"
    UpdateDictValue dict, "VC", "L"
    UpdateDictValue dict, "VI", "L"
    UpdateDictValue dict, "VA", "L"
    UpdateDictValue dict, "SC", "N"
    UpdateDictValue dict, "SI", "N"
    UpdateDictValue dict, "SA", "N"
End Function

Private Function ApplyOverrides(ByRef dict As Object)
    If dict.Item("E") = "X" Then
        UpdateDictValue dict, "E", "A"
    End If
    If dict.Item("CR") = "X" Then
        UpdateDictValue dict, "CR", "H"
    End If
    If dict.Item("IR") = "X" Then
        UpdateDictValue dict, "IR", "H"
    End If
    If dict.Item("AR") = "X" Then
        UpdateDictValue dict, "AR", "H"
    End If
End Function

Private Function ParseVectorString(ByVal VectorStr As String) As Object
    Dim dict As Object: Set dict = CreateDictionary()
    Dim parts() As String, i As Long, pair() As String

    FillDefaults dict
    VectorStr = UCase(Trim(VectorStr))
    
    If InStr(1, VectorStr, "INVALID") > 0 Or VectorStr = "" Then Exit Function
    If Left(VectorStr, 9) = "CVSS:4.0/" Then VectorStr = Mid(VectorStr, 10)

    parts = Split(VectorStr, "/")
    For i = LBound(parts) To UBound(parts)
        pair = Split(parts(i), ":")
        If UBound(pair) = 1 Then UpdateDictValue dict, pair(0), pair(1)
    Next i

    ApplyOverrides dict
    Set ParseVectorString = dict
End Function

Private Function GetMacroVectorBaseScore(ByVal macroVector As String) As Double
    GetMacroVectorBaseScore = getSafeRangeSearch("MacroVectorKeyMap", macroVector)
End Function

' ==============================================================================
' CVSS 4.0 EQUATION LEVELS
' ==============================================================================
Private Function getEqLevel1(dict As Object) As Integer
    Dim av As String, pr As String, ui As String
    av = dict.Item("AV")
    pr = dict.Item("PR")
    ui = dict.Item("UI")
    
    If av = "N" And pr = "N" And ui = "N" Then
        getEqLevel1 = 0
    ElseIf (av = "N" Or pr = "N" Or ui = "N") And Not av = "P" Then
        getEqLevel1 = 1
    Else
        getEqLevel1 = 2
    End If
End Function

Private Function getEqLevel2(dict As Object) As Integer
    Dim ac As String, at As String
    ac = dict.Item("AC")
    at = dict.Item("AT")

    If ac = "L" And at = "N" Then
        getEqLevel2 = 0
    Else
        getEqLevel2 = 1
    End If
End Function

Private Function getEqLevel3(dict As Object) As Integer
    Dim vc As String, vi As String, va As String
    vc = dict.Item("VC")
    vi = dict.Item("VI")
    va = dict.Item("VA")

    If vc = "H" And vi = "H" Then
        getEqLevel3 = 0
    ElseIf vc = "H" Or vi = "H" Or va = "H" Then
        getEqLevel3 = 1
    Else
        getEqLevel3 = 2
    End If
End Function

Private Function getEqLevel4(dict As Object) As Integer
    Dim sc As String, si As String, sa As String
    sc = dict.Item("SC")
    si = dict.Item("SI")
    sa = dict.Item("SA")

    If si = "S" Or sa = "S" Then
        getEqLevel4 = 0
    ElseIf sc = "H" Or si = "H" Or sa = "H" Then
        getEqLevel4 = 1
    Else
        getEqLevel4 = 2
    End If
End Function

Private Function getEqLevel5(dict As Object) As Integer
    Dim e As String
    e = dict.Item("E")

    If e = "A" Or e = "X" Then
        getEqLevel5 = 0
    ElseIf e = "P" Then
        getEqLevel5 = 1
    Else
        getEqLevel5 = 2
    End If
End Function

Private Function getEqLevel6(dict As Object) As Integer
    Dim cr As String, vc As String, ir As String, vi As String, ar As String, va As String
    cr = dict.Item("CR")
    vc = dict.Item("VC")
    ir = dict.Item("IR")
    vi = dict.Item("VI")
    ar = dict.Item("AR")
    va = dict.Item("VA")

    If cr = "X" Then cr = "H"
    If ir = "X" Then ir = "H"
    If ar = "X" Then ar = "H"
    
    If (cr = "H" And vc = "H") Or (ir = "H" And vi = "H") Or (ar = "H" And va = "H") Then
        getEqLevel6 = 0
    Else
        getEqLevel6 = 1
    End If
End Function

Private Function getJointEqLevel36(dict As Object) As String
    Dim cr As String, vc As String, ir As String, vi As String, ar As String, va As String
    cr = dict.Item("CR")
    vc = dict.Item("VC")
    ir = dict.Item("IR")
    vi = dict.Item("VI")
    ar = dict.Item("AR")
    va = dict.Item("VA")
    
    If vc = "H" And vi = "H" And (cr = "H" Or ir = "H" Or (ar = "H" And va = "H")) Then
        getJointEqLevel36 = "00"
    ElseIf vc = "H" And vi = "H" And Not (cr = "H" Or ir = "H") And Not (ar = "H" And va = "H") Then
        getJointEqLevel36 = "01"
    ElseIf Not (vc = "H" And vi = "H") And (vc = "H" Or vi = "H" Or va = "H") And _
           ((cr = "H" And vc = "H") Or (ir = "H" And vi = "H") Or (ar = "H" And va = "H")) Then
        getJointEqLevel36 = "10"
    ElseIf (vc = "H" And vi = "H") And (vc = "H" Or vi = "H" Or va = "H") And _
           (Not (cr = "H" And vc = "H") And Not (ir = "H" And vi = "H") And Not (ar = "H" And va = "H")) Then
        getJointEqLevel36 = "11"
    ElseIf Not (vc = "H" Or vi = "H" Or va = "H") And _
           ((cr = "H" And vc = "H") Or (ir = "H" And vi = "H") Or (ar = "H" And va = "H")) Then
        getJointEqLevel36 = "20"
    Else
        getJointEqLevel36 = "21"
    End If
End Function

Function getEQMaxes(lookup As String, eqIndex As Integer, Optional subEqIndex As Integer) As Variant
    Dim lookupKey As String
    Dim lookupSubKey As String
    lookupKey = Mid(lookup, eqIndex, 1)
    getEQMaxes = Array()

    Select Case eqIndex
        Case 1
            Select Case lookupKey
                Case "0": getEQMaxes = Array("AV:N/PR:N/UI:N/")
                Case "1": getEQMaxes = Array("AV:A/PR:N/UI:N/", "AV:N/PR:L/UI:N/", "AV:N/PR:N/UI:P/")
                Case "2": getEQMaxes = Array("AV:P/PR:N/UI:N/", "AV:A/PR:L/UI:P/")
            End Select
            
        Case 2
            Select Case lookupKey
                Case "0": getEQMaxes = Array("AC:L/AT:N/")
                Case "1": getEQMaxes = Array("AC:H/AT:N/", "AC:L/AT:P/")
            End Select
            
        Case 3
            lookupSubKey = Mid(lookup, subEqIndex, 1)
            Select Case lookupKey
                Case "0"
                    If lookupSubKey = "0" Then
                        getEQMaxes = Array("VC:H/VI:H/VA:H/CR:H/IR:H/AR:H/")
                    ElseIf lookupSubKey = "1" Then
                        getEQMaxes = Array("VC:H/VI:H/VA:L/CR:M/IR:M/AR:H/", "VC:H/VI:H/VA:H/CR:M/IR:M/AR:M/")
                    End If
                Case "1"
                    If lookupSubKey = "0" Then
                        getEQMaxes = Array("VC:L/VI:H/VA:H/CR:H/IR:H/AR:H/", "VC:H/VI:L/VA:H/CR:H/IR:H/AR:H/")
                    ElseIf lookupSubKey = "1" Then
                        getEQMaxes = Array("VC:L/VI:H/VA:L/CR:H/IR:M/AR:H/", "VC:L/VI:H/VA:H/CR:H/IR:M/AR:M/", "VC:H/VI:L/VA:H/CR:M/IR:H/AR:M/", "VC:H/VI:L/VA:L/CR:M/IR:H/AR:H/", "VC:L/VI:L/VA:H/CR:H/IR:H/AR:M/")
                    End If
                Case "2"
                    If lookupSubKey = "1" Then
                        getEQMaxes = Array("VC:L/VI:L/VA:L/CR:H/IR:H/AR:H/")
                    End If
            End Select
            
        Case 4
            Select Case lookupKey
                Case "0": getEQMaxes = Array("SC:H/SI:S/SA:S/")
                Case "1": getEQMaxes = Array("SC:H/SI:H/SA:H/")
                Case "2": getEQMaxes = Array("SC:L/SI:L/SA:L/")
            End Select
            
        Case 5
            Select Case lookupKey
                Case "0": getEQMaxes = Array("E:A/")
                Case "1": getEQMaxes = Array("E:P/")
                Case "2": getEQMaxes = Array("E:U/")
            End Select
    End Select
End Function

Private Function get_severity_distance(ByRef metric As Object, ByRef max_vector_dict As Object, ByVal key As String) As Double
    get_severity_distance = GetLevelValue(key, metric.Item(key)) - GetLevelValue(key, max_vector_dict.Item(key))
End Function

' ==============================================================================
' PUBLIC INTERFACE
' ==============================================================================
Public Function GetMacroVector(ByVal VectorStr As String) As String
    Dim dict As Object: Set dict = ParseVectorString(VectorStr)
    GetMacroVector = CStr(getEqLevel1(dict)) & CStr(getEqLevel2(dict)) & _
                     CStr(getEqLevel3(dict)) & CStr(getEqLevel4(dict)) & _
                     CStr(getEqLevel5(dict)) & CStr(getEqLevel6(dict))
End Function

Public Function getExploit(epss As Double, KEV As String)
    Dim ex As String, ea As String, ep As String, eu As String
    ex = "Not Defined (X)"
    ea = "Attacked (A)"
    ep = "Proof of Concept (P)"
    eu = "Unreported (U)"

    getExploit = ex
    If KEV = "Yes (Y)" Then
        getExploit = ea
    ElseIf epss >= 0.9 Then
        getExploit = ea
    ElseIf epss >= 0.5 Then
        getExploit = ep
    ElseIf epss >= 0.1 Then
        getExploit = eu
    End If
End Function

Public Function cvss_score(ByVal VectorStr As String) As Double
    Dim metric As Object
    Dim IsNotApplicable As Boolean
    Dim macroVectorScore As Double
    Dim v As Variant

    Set metric = ParseVectorString(VectorStr)

    IsNotApplicable = True
    For Each v In Array("VC", "VI", "VA", "SC", "SI", "SA")
        If metric.Item(v) <> "N" Then
            IsNotApplicable = False
        End If
    Next v

    If IsNotApplicable Then
        cvss_score = 0.0
        Exit Function
    End If

    Dim macroValueResult As String
    Dim eq1 As Integer, eq2 As Integer, eq3 As Integer, eq4 As Integer, eq5 As Integer, eq6 As Integer
    macroValueResult = GetMacroVector(VectorStr)
    macroVectorScore = GetMacroVectorBaseScore(macroValueResult)

    eq1 = CInt(Mid(macroValueResult, 1, 1))
    eq2 = CInt(Mid(macroValueResult, 2, 1))
    eq3 = CInt(Mid(macroValueResult, 3, 1))
    eq4 = CInt(Mid(macroValueResult, 4, 1))
    eq5 = CInt(Mid(macroValueResult, 5, 1))
    eq6 = CInt(Mid(macroValueResult, 6, 1))

    Dim eq1_next_lower_macro As String
    Dim eq2_next_lower_macro As String
    eq1_next_lower_macro = CStr(eq1 + 1) & CStr(eq2) & CStr(eq3) & CStr(eq4) & CStr(eq5) & CStr(eq6)
    eq2_next_lower_macro = CStr(eq1) & CStr(eq2 + 1) & CStr(eq3) & CStr(eq4) & CStr(eq5) & CStr(eq6)

    Dim eq3eq6_next_lower_macro As String
    Dim eq3eq6_next_lower_macro_left As String
    Dim eq3eq6_next_lower_macro_right As String
    
    If eq3 = 1 And eq6 = 1 Then
        eq3eq6_next_lower_macro = CStr(eq1) & CStr(eq2) & CStr(eq3 + 1) & CStr(eq4) & CStr(eq5) & CStr(eq6)
    ElseIf eq3 = 0 And eq6 = 1 Then
        eq3eq6_next_lower_macro = CStr(eq1) & CStr(eq2) & CStr(eq3 + 1) & CStr(eq4) & CStr(eq5) & CStr(eq6)
    ElseIf eq3 = 1 And eq6 = 0 Then
        eq3eq6_next_lower_macro = CStr(eq1) & CStr(eq2) & CStr(eq3) & CStr(eq4) & CStr(eq5) & CStr(eq6 + 1)
    ElseIf eq3 = 0 And eq6 = 0 Then
        eq3eq6_next_lower_macro_left = CStr(eq1) & CStr(eq2) & CStr(eq3) & CStr(eq4) & CStr(eq5) & CStr(eq6 + 1)
        eq3eq6_next_lower_macro_right = CStr(eq1) & CStr(eq2) & CStr(eq3 + 1) & CStr(eq4) & CStr(eq5) & CStr(eq6)
    Else
        eq3eq6_next_lower_macro = CStr(eq1) & CStr(eq2) & CStr(eq3) & CStr(eq4) & CStr(eq5) & CStr(eq6 + 1)
    End If

    Dim eq4_next_lower_macro As String
    Dim eq5_next_lower_macro As String
    eq4_next_lower_macro = CStr(eq1) & CStr(eq2) & CStr(eq3) & CStr(eq4 + 1) & CStr(eq5) & CStr(eq6)
    eq5_next_lower_macro = CStr(eq1) & CStr(eq2) & CStr(eq3) & CStr(eq4) & CStr(eq5 + 1) & CStr(eq6)

    Dim score_eq1_next_lower_macro As Variant
    Dim score_eq2_next_lower_macro As Variant
    score_eq1_next_lower_macro = getSafeRangeSearch("MacroVectorKeyMap", eq1_next_lower_macro)
    score_eq2_next_lower_macro = getSafeRangeSearch("MacroVectorKeyMap", eq2_next_lower_macro)

    Dim score_eq3eq6_next_lower_macro As Variant
    Dim score_eq3eq6_next_lower_macro_left As Variant
    Dim score_eq3eq6_next_lower_macro_right As Variant
    
    If eq3 = 0 And eq6 = 0 Then
        score_eq3eq6_next_lower_macro_left = getSafeRangeSearch("MacroVectorKeyMap", eq3eq6_next_lower_macro_left)
        score_eq3eq6_next_lower_macro_right = getSafeRangeSearch("MacroVectorKeyMap", eq3eq6_next_lower_macro_right)

        If Val(score_eq3eq6_next_lower_macro_left) > Val(score_eq3eq6_next_lower_macro_right) Then
            score_eq3eq6_next_lower_macro = score_eq3eq6_next_lower_macro_left
        Else
            score_eq3eq6_next_lower_macro = score_eq3eq6_next_lower_macro_right
        End If
    Else
        score_eq3eq6_next_lower_macro = getSafeRangeSearch("MacroVectorKeyMap", eq3eq6_next_lower_macro)
    End If

    Dim score_eq4_next_lower_macro As Variant
    Dim score_eq5_next_lower_macro As Variant
    score_eq4_next_lower_macro = getSafeRangeSearch("MacroVectorKeyMap", eq4_next_lower_macro)
    score_eq5_next_lower_macro = getSafeRangeSearch("MacroVectorKeyMap", eq5_next_lower_macro)

    Dim eq1_maxes As Variant, eq2_maxes As Variant, eq3_eq6_maxes As Variant, eq4_maxes As Variant, eq5_maxes As Variant
    eq1_maxes = getEQMaxes(macroValueResult, 1)
    eq2_maxes = getEQMaxes(macroValueResult, 2)
    eq3_eq6_maxes = getEQMaxes(macroValueResult, 3, 6)
    eq4_maxes = getEQMaxes(macroValueResult, 4)
    eq5_maxes = getEQMaxes(macroValueResult, 5)

    Dim max_vectors() As String
    Dim max_vectors_size As Integer: max_vectors_size = 0
    Dim eq1_max As Variant, eq2_max As Variant, eq3_eq6_max As Variant, eq4_max As Variant, eq5max As Variant
    Dim push_val As String

    For Each eq1_max In eq1_maxes
        For Each eq2_max In eq2_maxes
            For Each eq3_eq6_max In eq3_eq6_maxes
                For Each eq4_max In eq4_maxes
                    For Each eq5max In eq5_maxes
                        ReDim Preserve max_vectors(0 To max_vectors_size)
                        push_val = eq1_max & eq2_max & eq3_eq6_max & eq4_max & eq5max
                        max_vectors(max_vectors_size) = push_val
                        max_vectors_size = max_vectors_size + 1
                    Next eq5max
                Next eq4_max
            Next eq3_eq6_max
        Next eq2_max
    Next eq1_max

    Dim severity_distance_AV As Double, severity_distance_PR As Double, severity_distance_UI As Double
    Dim severity_distance_AC As Double, severity_distance_AT As Double
    Dim severity_distance_VC As Double, severity_distance_VI As Double, severity_distance_VA As Double
    Dim severity_distance_SC As Double, severity_distance_SI As Double, severity_distance_SA As Double
    Dim severity_distance_CR As Double, severity_distance_IR As Double, severity_distance_AR As Double
    Dim max_vector As Variant, max_vector_dict As Object

    For Each max_vector In max_vectors
        Set max_vector_dict = ParseVectorString(max_vector)
        severity_distance_AV = get_severity_distance(metric, max_vector_dict, "AV")
        severity_distance_PR = get_severity_distance(metric, max_vector_dict, "PR")
        severity_distance_UI = get_severity_distance(metric, max_vector_dict, "UI")
        severity_distance_AC = get_severity_distance(metric, max_vector_dict, "AC")
        severity_distance_AT = get_severity_distance(metric, max_vector_dict, "AT")
        severity_distance_VC = get_severity_distance(metric, max_vector_dict, "VC")
        severity_distance_VI = get_severity_distance(metric, max_vector_dict, "VI")
        severity_distance_VA = get_severity_distance(metric, max_vector_dict, "VA")
        severity_distance_SC = get_severity_distance(metric, max_vector_dict, "SC")
        severity_distance_SI = get_severity_distance(metric, max_vector_dict, "SI")
        severity_distance_SA = get_severity_distance(metric, max_vector_dict, "SA")
        severity_distance_CR = get_severity_distance(metric, max_vector_dict, "CR")
        severity_distance_IR = get_severity_distance(metric, max_vector_dict, "IR")
        severity_distance_AR = get_severity_distance(metric, max_vector_dict, "AR")

        Dim less_than_zero As Boolean: less_than_zero = False
        Dim severity As Variant
        For Each severity In Array(severity_distance_AV, severity_distance_PR, severity_distance_UI, _
                                   severity_distance_AC, severity_distance_AT, severity_distance_VC, _
                                   severity_distance_VI, severity_distance_VA, severity_distance_SC, _
                                   severity_distance_SI, severity_distance_SA, severity_distance_CR, _
                                   severity_distance_IR, severity_distance_AR)
            If severity < 0 Then
                less_than_zero = True
            End If
        Next severity

        If Not less_than_zero Then
            Exit For
        End If
    Next max_vector

    Dim current_severity_distance_eq1 As Double
    Dim current_severity_distance_eq2 As Double
    Dim current_severity_distance_eq3eq6 As Double
    Dim current_severity_distance_eq4 As Double
    Dim current_severity_distance_eq5 As Double
    
    current_severity_distance_eq1 = severity_distance_AV + severity_distance_PR + severity_distance_UI
    current_severity_distance_eq2 = severity_distance_AC + severity_distance_AT
    current_severity_distance_eq3eq6 = severity_distance_VC + severity_distance_VI + severity_distance_VA + severity_distance_CR + severity_distance_IR + severity_distance_AR
    current_severity_distance_eq4 = severity_distance_SC + severity_distance_SI + severity_distance_SA
    current_severity_distance_eq5 = 0

    Dim available_distance_eq1 As Double
    Dim available_distance_eq2 As Double
    Dim available_distance_eq3eq6 As Double
    Dim available_distance_eq4 As Double
    Dim available_distance_eq5 As Double
    
    available_distance_eq1 = macroVectorScore - Val(score_eq1_next_lower_macro)
    available_distance_eq2 = macroVectorScore - Val(score_eq2_next_lower_macro)
    available_distance_eq3eq6 = macroVectorScore - Val(score_eq3eq6_next_lower_macro)
    available_distance_eq4 = macroVectorScore - Val(score_eq4_next_lower_macro)
    available_distance_eq5 = macroVectorScore - Val(score_eq5_next_lower_macro)

    Dim n_existing_lower As Double: n_existing_lower = 0
    Dim normalized_severity_eq1 As Double: normalized_severity_eq1 = 0
    Dim normalized_severity_eq2 As Double: normalized_severity_eq2 = 0
    Dim normalized_severity_eq3eq6 As Double: normalized_severity_eq3eq6 = 0
    Dim normalized_severity_eq4 As Double: normalized_severity_eq4 = 0
    Dim normalized_severity_eq5 As Double: normalized_severity_eq5 = 0

    Dim maxSeverity_eq1 As Double, maxSeverity_eq2 As Double, maxSeverity_eq3eq6 As Double, maxSeverity_eq4 As Double
    maxSeverity_eq1 = GetMaxSeverity("eq1", eq1) * 0.1
    maxSeverity_eq2 = GetMaxSeverity("eq2", eq2) * 0.1
    maxSeverity_eq3eq6 = GetMaxSeverity("eq3eq6", eq3, eq6) * 0.1
    maxSeverity_eq4 = GetMaxSeverity("eq4", eq4) * 0.1

    Dim percent_to_next_eq1_severity As Double
    Dim percent_to_next_eq2_severity As Double
    Dim percent_to_next_eq3eq6_severity As Double
    Dim percent_to_next_eq4_severity As Double
    Dim percent_to_next_eq5_severity As Double

    If score_eq1_next_lower_macro <> "" Then
        available_distance_eq1 = macroVectorScore - Val(score_eq1_next_lower_macro)
        n_existing_lower = n_existing_lower + 1
        percent_to_next_eq1_severity = (current_severity_distance_eq1) / maxSeverity_eq1
        normalized_severity_eq1 = available_distance_eq1 * percent_to_next_eq1_severity
    End If
    
    If score_eq2_next_lower_macro <> "" Then
        available_distance_eq2 = macroVectorScore - Val(score_eq2_next_lower_macro)
        n_existing_lower = n_existing_lower + 1
        percent_to_next_eq2_severity = (current_severity_distance_eq2) / maxSeverity_eq2
        normalized_severity_eq2 = available_distance_eq2 * percent_to_next_eq2_severity
    End If

    If score_eq3eq6_next_lower_macro <> "" Then
        available_distance_eq3eq6 = macroVectorScore - Val(score_eq3eq6_next_lower_macro)
        n_existing_lower = n_existing_lower + 1
        percent_to_next_eq3eq6_severity = (current_severity_distance_eq3eq6) / maxSeverity_eq3eq6
        normalized_severity_eq3eq6 = available_distance_eq3eq6 * percent_to_next_eq3eq6_severity
    End If

    If score_eq4_next_lower_macro <> "" Then
        available_distance_eq4 = macroVectorScore - Val(score_eq4_next_lower_macro)
        n_existing_lower = n_existing_lower + 1
        percent_to_next_eq4_severity = (current_severity_distance_eq4) / maxSeverity_eq4
        normalized_severity_eq4 = available_distance_eq4 * percent_to_next_eq4_severity
    End If
    
    If score_eq5_next_lower_macro <> "" Then
        available_distance_eq5 = macroVectorScore - Val(score_eq5_next_lower_macro)
        n_existing_lower = n_existing_lower + 1
        percent_to_next_eq5_severity = 0
        normalized_severity_eq5 = available_distance_eq5 * percent_to_next_eq5_severity
    End If

    Dim mean_distance As Double
    If n_existing_lower = 0 Then
        mean_distance = 0
    Else
        mean_distance = (normalized_severity_eq1 + normalized_severity_eq2 + normalized_severity_eq3eq6 + normalized_severity_eq4 + normalized_severity_eq5) / n_existing_lower
    End If

    macroVectorScore = macroVectorScore - mean_distance
    If macroVectorScore < 0 Then
        macroVectorScore = 0.0
    End If
    If macroVectorScore > 10 Then
        macroVectorScore = 10.0
    End If
    
    cvss_score = Round(macroVectorScore, 1)
End Function
