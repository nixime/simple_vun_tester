' ==============================================================================
' MODULE: CVSS v4.0 Calculation and Parsing Engine
' COMPATIBILITY: Microsoft Excel (VBA) & LibreOffice Calc (StarSuite Basic)
' ==============================================================================

Option Explicit
Option VBASupport 1 ' Crucial directive for LibreOffice Calc to support Excel VBA properties

'''
' Function: GET_CVSS_ELEMENT
' Purpose: Parses an official CVSS vector string and extracts a specific metric code,
'          returning its human-readable definition.
' Inputs:
'   - vectorString: The complete CVSS vector (e.g., "CVSS:4.0/AV:N/AC:L/AT:N...")
'   - metricCode: The exact code to look up (e.g., "AV", "PR", "UI")
' Outputs:
'   - String: The translated long name with its code, or "Not Found" if missing.
'''
Function GET_CVSS_ELEMENT(ByVal vectorString As String, ByVal metricCode As String) As String
    ' Local Variable Declarations (Enforced by Option Explicit)
    Dim pairs() As String    ' Array to hold split components of the vector string
    Dim pair As Variant      ' Iterator variable used for the For Each loop
    Dim searchKey As String  ' The exact metric substring we want to locate (e.g., "AV:")
    Dim val As String        ' Stores the extracted raw code value (e.g., "N")
    
    ' Construct the exact search prefix to avoid false-positive partial matches.
    ' For example, searching for "AC:" should never match "MAC:" (Modified Attack Complexity)
    searchKey = metricCode & ":"
    
    ' Split the full vector string into individual key-value pairs using the slash '/' delimiter
    pairs = Split(vectorString, "/")
    
    ' Iterate through every tokenized pair to extract the correct metric
    For Each pair In pairs
        ' Ensure the current component starts exactly with our target search key
        ' CStr() guarantees safety if the variant element undergoes comparison evaluation
        If Left(CStr(pair), Len(searchKey)) = searchKey Then
            ' Extract everything following the colon character (e.g., from "AV:N" grab "N")
            val = Mid(CStr(pair), Len(searchKey) + 1)
            Exit For ' Match found; safely terminate loop execution early
        End If
    Next pair
    
    ' If the requested metric was completely absent from the vector string, return a predictable error text
    If val = "" Then
        GET_CVSS_ELEMENT = "Not Found"
        Exit Function
    End If
    
    ' Translate the single-letter abbreviations into official CVSS v4.0 human-readable metrics
    Select Case metricCode
        Case "AV" ' --- Attack Vector ---
            Select Case val
                Case "N": GET_CVSS_ELEMENT = "Network (N)"
                Case "A": GET_CVSS_ELEMENT = "Adjacent (A)"
                Case "L": GET_CVSS_ELEMENT = "Local (L)"
                Case "P": GET_CVSS_ELEMENT = "Physical (P)"
                Case Else: GET_CVSS_ELEMENT = val
            End Select
            
        Case "AC" ' --- Attack Complexity ---
            Select Case val
                Case "L": GET_CVSS_ELEMENT = "Low (L)"
                Case "H": GET_CVSS_ELEMENT = "High (H)"
                Case Else: GET_CVSS_ELEMENT = val
            End Select
            
        Case "AT" ' --- Attack Requirements ---
            Select Case val
                Case "N": GET_CVSS_ELEMENT = "None (N)"
                Case "P": GET_CVSS_ELEMENT = "Present (P)"
                Case Else: GET_CVSS_ELEMENT = val
            End Select
            
        Case "PR" ' --- Privileges Required ---
            Select Case val
                Case "N": GET_CVSS_ELEMENT = "None (N)"
                Case "L": GET_CVSS_ELEMENT = "Low (L)"
                Case "H": GET_CVSS_ELEMENT = "High (H)"
                Case Else: GET_CVSS_ELEMENT = val
            End Select
            
        Case "UI" ' --- User Interaction ---
            Select Case val
                Case "N": GET_CVSS_ELEMENT = "None (N)"
                Case "P": GET_CVSS_ELEMENT = "Passive (P)"
                Case "A": GET_CVSS_ELEMENT = "Active (A)"
                Case Else: GET_CVSS_ELEMENT = val
            End Select
            
        Case "VC", "VI", "VA", "SC", "SI", "SA" ' --- Vulnerability & Subsequent System Impact Metrics ---
            Select Case val
                Case "N": GET_CVSS_ELEMENT = "None (N)"
                Case "L": GET_CVSS_ELEMENT = "Low (L)"
                Case "H": GET_CVSS_ELEMENT = "High (H)"
                Case Else: GET_CVSS_ELEMENT = val
            End Select
            
        Case Else
            ' Fallback mechanism for unmapped environmental, temporal, or supplemental metrics (e.g., E, CR, IR)
            GET_CVSS_ELEMENT = val
    End Select
End Function


'''
' Function: GET_CVSS4_VECTOR_STR
' Purpose: Takes individual spreadsheet cell parameters and structures them into a perfectly formatted,
'          specification-compliant CVSS v4.0 vector string.
' Inputs:
'   - AV through SA: Explicitly typed as Variants to handle cross-platform object behaviors gracefully.
'   - E: (Optional) Exploit Maturity parameter. Passes as Variant to allow IsMissing check.
' Outputs:
'   - String: Complete stitched vector (prefixed with CVSS:4.0/) or data error string.
'''
Function GET_CVSS4_VECTOR_STR( _
    ByVal AV As Variant, ByVal AC As Variant, ByVal AT As Variant, _
    ByVal PR As Variant, ByVal UI As Variant, ByVal VC As Variant, _
    ByVal VI As Variant, ByVal VA As Variant, ByVal SC As Variant, _
    ByVal SI As Variant, ByVal SA As Variant, _
    Optional ByVal E As Variant) As String
    
    ' Local Variable Declarations (Enforced by Option Explicit)
    Dim rawMetrics() As Variant      ' Container tracking raw cell/object references 
    Dim prefixes() As Variant        ' Matching structural specification prefixes for the metrics
    Dim metrics() As String          ' Final processed strings ready to be joined
    Dim i As Integer                 ' Loop index counter variable
    Dim size As Integer              ' Dynamically calculated limit of metrics to stitch together
    Dim isEOptProvided As Boolean    ' Tracking flag checking if Optional metric 'E' contains active data
    
    ' Step 1: Safely evaluate if the optional 12th parameter 'E' is active and populated.
    ' Under Option Explicit, we use IsMissing() to ensure an unpassed cell argument doesn't cause a runtime crash.
    isEOptProvided = False
    If Not IsMissing(E) Then
        If CStr(E) <> "" Then
            isEOptProvided = True
        End If
    End If
    
    ' Step 2: Establish array bounds dynamically.
    ' If Exploit Maturity (E) is present, process 12 metrics. Otherwise, stop at 11 base metrics.
    If isEOptProvided Then size = 12 Else size = 11
    
    ' Allocate the metrics array to match the dynamic size bounds explicitly
    ReDim metrics(1 To size)
    
    ' Step 3: Bundle tracking metrics into 1-indexed structures.
    ' The 'Empty' placeholder acts as a 0-index dummy value. This matches loop iteration index "i" flawlessly.
    ' If E is not provided, we pass 'Empty' into the array layout safely.
    rawMetrics = Array(Empty, AV, AC, AT, PR, UI, VC, VI, VA, SC, SI, SA, IIf(isEOptProvided, E, Empty))
    prefixes = Array(Empty, "AV:", "AC:", "AT:", "PR:", "UI:", "VC:", "VI:", "VA:", "SC:", "SI:", "SA:", "E:")
    
    ' Redirect execution to the error routine if an unhandled evaluation occurs
    On Error GoTo ErrorHandler

    ' Step 4: Iteratively unpack, normalize, and prefix every metric component
    For i = 1 To size
        Dim val As String
        
        ' --- CROSS-PLATFORM INTEROPERABILITY ENGINE ---
        ' Spreadsheet engines handle worksheet parameters differently when received inside a macro function:
        '   - Excel passes individual sheet cells as a Range Object (Internal VarType constant = 9).
        '   - LibreOffice Calc intercepts variables, converting individual cells to primitive variants or multi-dimensional data arrays.
        If VarType(rawMetrics(i)) = 9 Then
            ' Excel Mode: Safely pull the string equivalent out of the active Range Object's .Value property
            val = Trim(CStr(rawMetrics(i).Value))
        ElseIf IsArray(rawMetrics(i)) Then
            ' LibreOffice Mode Fallback: Safely pull value from the first row/column coordinate of the 2D variant matrix
            val = Trim(CStr(rawMetrics(i)(1, 1)))
        Else
            ' Direct Primitive Fallback: Treats variables as standard memory strings/integers directly
            val = Trim(CStr(rawMetrics(i)))
        End If
        
        ' Short-circuit validation check: If a cell tracking formula upstream outputs an invalid status string
        If InStr(1, UCase(val), "INVALID") > 0 Then
            GET_CVSS4_VECTOR_STR = "INVALID"
            Exit Function
        End If
        
        ' Step 5: Normalize human-readable descriptions back down to strict specification codes
        ' For example: converts "Network (N)" -> "N" by extracting characters enclosed within parenthetical boundaries.
        Dim startPos As Integer, endPos As Integer
        startPos = InStr(val, "(")
        endPos = InStr(val, ")")
        
        If startPos > 0 And endPos > startPos Then
            val = Mid(val, startPos + 1, endPos - startPos - 1)
        End If
        
        ' Construct and store the string element (e.g., "AV:" & "N" becomes "AV:N")
        metrics(i) = prefixes(i) & Trim(val)
    Next i
    
    ' Step 6: Assemble final vector string, complete with official specifications layout header
    GET_CVSS4_VECTOR_STR = "CVSS:4.0/" & Join(metrics, "/")
    Exit Function

ErrorHandler:
    ' Emergency fail-safe fallback routine tracking structure memory faults
    GET_CVSS4_VECTOR_STR = "MACRO_ERROR"
End Function
