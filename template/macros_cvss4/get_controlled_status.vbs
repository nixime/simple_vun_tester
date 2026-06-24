' ==============================================================================
' MODULE: Controlled Status Lookup Engine
' COMPATIBILITY: Microsoft Excel (VBA) & LibreOffice Calc (StarSuite Basic)
' ==============================================================================

Option Explicit
Option VBASupport 1 ' Crucial directive for LibreOffice Calc to support Excel VBA object properties

'''
' Function: GET_CONTROLLED_STATUS
' Purpose: Interrogates a named-range matrix ("_lookup_SafetyHeatMap") to find a risk status
'          based on a numeric CVSS Score and a string Safety Rating header.
' Inputs:
'   - cvssScore: The precise calculated vulnerability score.
'   - safetyRating: The column header text to match (e.g., "High", "Critical").
' Outputs:
'   - String: The mapped heat map status text, or "Not Found" if lookups fail.
'''
Function GET_CONTROLLED_STATUS(ByVal cvssScore As Double, ByVal safetyRating As String) As String
    ' Local Variable Declarations (Enforced by Option Explicit)
    Dim isExcel As Boolean         ' Flag tracking active application type (True = Excel, False = LibreOffice)
    Dim targetRange As Object      ' Abstract object referencing the resolved spreadsheet range matrix
    Dim foundStatus As String      ' Stores the final text extracted from the matched coordinate
    Dim colIdx As Long             ' Tracks the targeted column index matching the Safety Rating header
    Dim rowIdx As Long             ' Loop variable used to step down the range matrix rows
    Dim colIdxLoop As Long         ' Loop variable used to scan across the matrix header columns
    Dim rowCount As Long           ' Total vertical depth of the named range matrix
    Dim colCount As Long           ' Total horizontal width of the named range matrix
    Dim heatMapCVSS As Double      ' Temporary storage for parsing the lower-bound CVSS marker per matrix row
    
    ' --- STEP 1: ENVIRONMENT DETECTION ENGINE ---
    ' Dynamically inspects runtime availability to determine if host application is Microsoft Excel or LibreOffice Calc.
    isExcel = True
    On Error Resume Next
    ' "ThisComponent" is a global UNO API handle native exclusively to LibreOffice Calc modules.
    ' If accessing it fails or evaluates as empty, the script is running inside native Excel VBA.
    If Not IsEmpty(ThisComponent) Then isExcel = False
    Err.Clear
    On Error GoTo ErrorHandler

    ' Initialize default return string in case lookup coordinates fail to resolve
    foundStatus = "Not Found"

    ' --- STEP 2: CROSS-PLATFORM NAMED RANGE RESOLUTION ---
    ' Obtains a reference to the structural grid map regardless of the background Application runtime interface.
    If Not isExcel Then
        ' --- LIBREOFFICE CALC APPLICATION DESIGN (UNO API) ---
        Dim oDoc As Object, oNames As Object
        Set oDoc = ThisComponent
        Set oNames = oDoc.NamedRanges
        
        ' Verify the global defined expression name explicitly exists before extracting cells
        If oNames.hasByName("_lookup_SafetyHeatMap") Then
            Set targetRange = oNames.getByName("_lookup_SafetyHeatMap").getReferredCells()
        End If
    Else
        ' --- MICROSOFT EXCEL APPLICATION DESIGN (VBA API) ---
        On Error Resume Next
        ' Native Excel engine resolves string keys directly against the global application Workbook Range collection
        Set targetRange = Range("_lookup_SafetyHeatMap")
        On Error GoTo ErrorHandler
    End If

    ' --- STEP 3: PLATFORM-AGNOSTIC DATA INTERROGATION ---
    ' Executes grid calculations safely using uniform row/col abstractions once the object target resolves.
    If Not targetRange Is Nothing Then
        rowCount = targetRange.Rows.Count
        colCount = targetRange.Columns.Count
        
        ' Loop A: Extract the exact Column Index matching the 'safetyRating' argument.
        ' It checks the first row across every column inside the bounding matrix.
        colIdx = 0
        For colIdxLoop = 1 To colCount
            ' Trim string whitespaces and convert to lowercase to safeguard against data-entry formatting mismatches
            If LCase(Trim(GetCell(targetRange, 1, colIdxLoop, isExcel))) = LCase(Trim(safetyRating)) Then
                colIdx = colIdxLoop
                Exit For ' Match identified; immediately stop tracking loop columns
            End If
        Next colIdxLoop
        
        ' Loop B: Step vertically through rows to discover the corresponding numeric threshold match.
        ' Logic Assumptions: Matrix data headers occupy Row 1; Column 2 stores the matching minimum threshold CVSS score.
        If colIdx > 0 Then
            For rowIdx = 2 To rowCount Step 1
                ' Cast extracted string safely to double primitive format before attempting comparative logic evaluation
                heatMapCVSS = CDbl(GetCell(targetRange, rowIdx, 2, isExcel))
                
                ' Matrix sorting assumption: Ordered descending or checked sequentially for higher bounds.
                ' If the variable score exceeds or meets the threshold tier, capture target cell.
                If cvssScore >= heatMapCVSS Then
                    foundStatus = GetCell(targetRange, rowIdx, colIdx, isExcel)
                    Exit For ' Target tier located; truncate execution loop
                End If
            Next rowIdx
        End If
    End If

    ' Assign final output state back to the spreadsheet function calling framework
    GET_CONTROLLED_STATUS = foundStatus
ErrorHandler:
    GET_CONTROLLED_STATUS = "Unknown"
End Function


'''
' Function: GetCell
' Purpose: Abstraction helper bridging API coordinate queries across Excel and LibreOffice Calc.
' Inputs:
'   - rng: The resolved workbook block object matrix.
'   - r: 1-indexed row number.
'   - c: 1-indexed column number.
'   - isExcel: Execution route routing flag flag state.
' Outputs:
'   - String: Normalized string contents of the requested cell.
'''
Function GetCell(ByVal rng As Object, ByVal r As Long, ByVal c As Long, ByVal isExcel As Boolean) As String
    If isExcel Then
        ' --- Excel Processing ---
        ' Excel Range arrays are natively 1-indexed. Returns cell property contents.
        GetCell = CStr(rng.Cells(r, c).Value)
    Else
        ' --- LibreOffice Calc Processing ---
        ' UNO API requires absolute 0-indexing for positional cell grid pointers.
        ' Row 1, Col 1 corresponds to coordinate index target point (0, 0).
        GetCell = rng.getCellByPosition(c - 1, r - 1).getString()
    End If
End Function