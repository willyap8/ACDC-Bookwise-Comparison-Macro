Option Explicit

' ============================================================
'  ACDC Data Migration - Bookwise Daily Comparison Macro V3.1
'  Sub: BookwiseReconcileDailyReportV3
'  Paste this entire module into the add-in workbook (.xlam).
'  Run while the MASTER data worksheet is the active sheet.
'
'  V3.1 change: Date and Date of Birth COMPARISON is now
'  normalised to a canonical dd/mm/yyyy string via NormaliseDate()
'  before comparing, so a genuine Excel date value on one side and
'  a text date (e.g. "01/09/2026") on the other are treated as
'  equal. This prevents false "Modified" flags when MASTER and
'  DAILY store dates with different underlying types/formats.
'  Text dates are parsed explicitly as day/month/year (Australian);
'  unparseable values fall back to a literal text compare.
'
'  Procedure name kept as BookwiseReconcileDailyReportV3 so the
'  Quick Access Toolbar binding is not broken by the point release.
'
'  Prior changes retained:
'   - V3:   Date / Date of Birth copied to review sheets as verbatim
'           TEXT via CopyDateAsText() (prevents dd/mm <-> mm/dd
'           corruption on copy).
'   - v1.2: Start Time comparison normalised to hh:mm via
'           NormaliseTime().
' ============================================================

Public Sub BookwiseReconcileDailyReportV3()

    Dim wsMaster        As Worksheet
    Dim wsDaily         As Worksheet
    Dim wbDaily         As Workbook
    Dim wbMaster        As Workbook
    Dim wsMod           As Worksheet
    Dim wsCan           As Worksheet
    Dim wsNew           As Worksheet
    Dim wsAudit         As Worksheet
    Dim dailyDict       As Object
    Dim masterDict      As Object
    Dim dupDict         As Object
    Dim lo              As ListObject
    Dim masterTbl       As ListObject

    Dim dailyPath       As String
    Dim masterDataStart As Long
    Dim dailyDataStart  As Long
    Dim lastRowMaster   As Long
    Dim lastRowDaily    As Long
    Dim compareCol      As Long
    Dim tableLastCol    As Long
    Dim bookCol         As Long
    Dim patientCol      As Long
    Dim startTimeCol    As Long
    Dim medicareCol     As Long
    Dim consultantCol   As Long
    Dim postcodeCol     As Long
    Dim dateCol         As Long
    Dim dobCol          As Long
    Dim dailyBookCol    As Long

    Dim i               As Long
    Dim j               As Long
    Dim k               As Long
    Dim r               As Long
    Dim totalProcessed  As Long
    Dim modifiedCount   As Long
    Dim cancelledCount  As Long
    Dim newCount        As Long
    Dim totalRows       As Long
    Dim pct             As Long
    Dim auditNextRow    As Long
    Dim dailyRow        As Long

    Dim bookNo          As String
    Dim dKey            As String
    Dim masterVal       As String
    Dim dailyVal        As String
    Dim blankRows       As String
    Dim dupMsg          As String
    Dim hdrMaster       As String
    Dim hdrDaily        As String
    Dim modified        As Boolean
    Dim isDiff          As Boolean
    Dim isExcluded      As Boolean
    Dim hasOtherData    As Boolean
    Dim paramMatch      As Boolean
    Dim resizeOK        As Boolean
    Dim runTime         As Date

    runTime = Now()

    ' --------------------------------------------------------
    ' 0. SHEET NAME VALIDATION
    '    Absolute first check. No performance flags set yet.
    ' --------------------------------------------------------
    If ActiveSheet.Name <> "MASTER" Then
        MsgBox "This macro is only designed to run on the master sheet. " & _
               "Please recheck the sheet you are running on and try again.", _
               vbCritical, "Wrong Sheet"
        Exit Sub
    End If

    ' --------------------------------------------------------
    ' 0b. LOCKED SHEET DETECTION
    '     Runs before the file picker so the macro fails fast
    '     without opening any file or making any change.
    ' --------------------------------------------------------
    If ActiveSheet.ProtectContents Then
        MsgBox "The MASTER sheet is currently protected (locked)." & vbCrLf & vbCrLf & _
               "Please unlock the sheet before running this macro:" & vbCrLf & _
               "  Review tab  ->  Unprotect Sheet" & vbCrLf & vbCrLf & _
               "Once unlocked, run the macro again.", _
               vbCritical, "Sheet Is Locked"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo CleanFail

    ' --------------------------------------------------------
    ' 1. Reference MASTER workbook and active sheet
    '    ActiveSheet.Parent resolves the host workbook correctly
    '    when the macro is launched from an add-in (.xlam).
    ' --------------------------------------------------------
    Set wsMaster = ActiveSheet
    Set wbMaster = ActiveSheet.Parent

    ' --------------------------------------------------------
    ' 2. Determine column ranges
    '    compareCol   = comparison range, HARD CAP at 20 (Col T).
    '                   Nothing beyond Column T is ever read,
    '                   compared, highlighted, or copied.
    '    tableLastCol = full table range, detected dynamically
    '                   from header row 11 (NO cap) so comment
    '                   columns after T are included in the table
    '                   for correct filtering. Comment columns are
    '                   never used in the comparison.
    ' --------------------------------------------------------
    tableLastCol = wsMaster.Cells(11, wsMaster.Columns.Count).End(xlToLeft).Column
    If tableLastCol < 1 Then tableLastCol = 1

    compareCol = tableLastCol
    If compareCol > 20 Then compareCol = 20   ' Hard cap at Column T

    If SafeText(wsMaster.Cells(11, 1).Value) = "" Then
        MsgBox "Could not detect column headers in row 11 of this sheet." & vbCrLf & vbCrLf & _
               "Ensure you are running this macro while the MASTER data sheet is active.", _
               vbCritical, "Setup Error"
        GoTo CleanFail
    End If

    ' --------------------------------------------------------
    ' 3. FILE PICKER
    '    Runs before ANY structural changes to MASTER.
    '    Cancel routes to CleanExit - MASTER is untouched.
    ' --------------------------------------------------------
    dailyPath = Application.GetOpenFilename( _
        FileFilter:="Excel Files (*.xlsx;*.xlsm;*.xls;*.xlsb),*.xlsx;*.xlsm;*.xls;*.xlsb", _
        Title:="Select the Daily Report to Reconcile")

    If dailyPath = "False" Then
        MsgBox "No file selected. Macro cancelled.", vbExclamation, "Cancelled"
        GoTo CleanExit
    End If

    ' --------------------------------------------------------
    ' 4. Open DAILY workbook (read-only)
    ' --------------------------------------------------------
    Set wbDaily = Workbooks.Open(Filename:=dailyPath, ReadOnly:=True)
    Set wsDaily = wbDaily.ActiveSheet

    ' --------------------------------------------------------
    ' 5. SEARCH PARAMETER VALIDATION (rows 2-8, columns A and C)
    ' --------------------------------------------------------
    paramMatch = True
    For r = 2 To 8
        If SafeText(wsMaster.Cells(r, 1).Value) <> SafeText(wsDaily.Cells(r, 1).Value) Then
            paramMatch = False
            Exit For
        End If
        If SafeText(wsMaster.Cells(r, 3).Value) <> SafeText(wsDaily.Cells(r, 3).Value) Then
            paramMatch = False
            Exit For
        End If
    Next r

    If Not paramMatch Then
        MsgBox "The report search parameters appear to be different, please recheck the daily report and try again.", _
               vbCritical, "Parameter Mismatch"
        GoTo CleanFail
    End If

    ' --------------------------------------------------------
    ' 6. HEADER MATCH VALIDATION (row 11, columns A to T)
    '    Exact match required: spelling, case, and order.
    '    Comparison only ever uses A:T, so only A:T headers
    '    must match. Stops and names the first mismatched column.
    ' --------------------------------------------------------
    For j = 1 To compareCol
        hdrMaster = SafeText(wsMaster.Cells(11, j).Value)
        hdrDaily = SafeText(wsDaily.Cells(11, j).Value)
        If hdrMaster <> hdrDaily Then
            MsgBox "The column headers do not match between the MASTER and Daily report." & vbCrLf & vbCrLf & _
                   "First mismatch in column " & ColLetter(j) & " (header row 11):" & vbCrLf & _
                   "  MASTER: """ & hdrMaster & """" & vbCrLf & _
                   "  Daily:  """ & hdrDaily & """" & vbCrLf & vbCrLf & _
                   "Please recheck the daily report and try again.", _
                   vbCritical, "Header Mismatch"
            GoTo CleanFail
        End If
    Next j

    ' --------------------------------------------------------
    ' 7. MASTER PREP: Delete blank separator row 12 if present.
    '    masterDataStart is always set to 12 after this block.
    ' --------------------------------------------------------
    If Application.WorksheetFunction.CountA(wsMaster.Rows(12)) = 0 Then
        wsMaster.Rows(12).Delete Shift:=xlShiftUp
    End If
    masterDataStart = 12

    ' --------------------------------------------------------
    ' 8. Determine DAILY data start row (never modify DAILY)
    ' --------------------------------------------------------
    If Application.WorksheetFunction.CountA(wsDaily.Rows(12)) = 0 Then
        dailyDataStart = 13
    Else
        dailyDataStart = 12
    End If

    ' --------------------------------------------------------
    ' 9. Initial last-row estimate (for table range)
    ' --------------------------------------------------------
    lastRowMaster = wsMaster.Cells(wsMaster.Rows.Count, 1).End(xlUp).Row
    If lastRowMaster < masterDataStart Then
        MsgBox "No data rows found below the header in the MASTER sheet.", _
               vbExclamation, "No Data"
        GoTo CleanFail
    End If

    ' --------------------------------------------------------
    ' 10. CONVERT OR RESIZE MASTER TABLE to A11:[tableLastCol]
    '     Robust handling: the report-header merges above the
    '     table (A1:T1, C2:T2 ... A10:T10) stop at column T.
    '     A direct .Resize across the T->AG boundary can raise
    '     error 1004. We therefore try .Resize first inside a
    '     local trap; if it fails, we fall back to removing the
    '     table (data preserved via .Unlist) and recreating it
    '     at the full A11:[tableLastCol] range.
    ' --------------------------------------------------------
    Set masterTbl = Nothing
    For Each lo In wsMaster.ListObjects
        If lo.HeaderRowRange.Row = 11 Then
            Set masterTbl = lo
            Exit For
        End If
    Next lo

    resizeOK = False

    If masterTbl Is Nothing Then
        ' No table yet - create one directly
        On Error Resume Next
        Set masterTbl = wsMaster.ListObjects.Add( _
            xlSrcRange, _
            wsMaster.Range(wsMaster.Cells(11, 1), wsMaster.Cells(lastRowMaster, tableLastCol)), _
            , xlYes)
        If Not masterTbl Is Nothing Then
            masterTbl.Name = "MasterData"
            resizeOK = True
        End If
        On Error GoTo CleanFail
    Else
        ' Table exists - try a direct resize first
        On Error Resume Next
        masterTbl.Resize wsMaster.Range( _
            wsMaster.Cells(11, 1), wsMaster.Cells(lastRowMaster, tableLastCol))
        If Err.Number = 0 Then resizeOK = True
        Err.Clear
        On Error GoTo CleanFail

        ' Fallback: if direct resize failed, unlist and recreate
        If Not resizeOK Then
            On Error Resume Next
            masterTbl.Unlist                       ' converts back to a normal range, data preserved
            On Error GoTo CleanFail

            On Error Resume Next
            Set masterTbl = wsMaster.ListObjects.Add( _
                xlSrcRange, _
                wsMaster.Range(wsMaster.Cells(11, 1), wsMaster.Cells(lastRowMaster, tableLastCol)), _
                , xlYes)
            If Not masterTbl Is Nothing Then
                masterTbl.Name = "MasterData"
                resizeOK = True
            End If
            On Error GoTo CleanFail
        End If
    End If

    If Not resizeOK Then
        MsgBox "Unable to format the data range as a table covering columns A to " & _
               ColLetter(tableLastCol) & "." & vbCrLf & vbCrLf & _
               "This is often caused by merged cells overlapping the table area, " & _
               "or another table on the sheet. Please check the MASTER sheet layout " & _
               "and try again.", _
               vbCritical, "Table Setup Failed"
        GoTo CleanFail
    End If

    ' --------------------------------------------------------
    ' 11. FONT STANDARDISATION (neutralises Wingdings/symbol fonts)
    '     Scoped to the comparison range only (A:T).
    ' --------------------------------------------------------
    Call StandardiseFont(wsMaster, compareCol)
    Call StandardiseFont(wsDaily, compareCol)

    ' --------------------------------------------------------
    ' 12. REQUIRED / EXCLUDED COLUMNS (searched within A:T only)
    '     dateCol / dobCol drive both verbatim TEXT copying (V3)
    '     and canonical date-comparison normalisation (V3.1).
    ' --------------------------------------------------------
    bookCol = FindHeaderColumn(wsMaster, "Book No.", compareCol)
    patientCol = FindHeaderColumn(wsMaster, "Patient", compareCol)
    startTimeCol = FindHeaderColumn(wsMaster, "Start Time", compareCol)
    medicareCol = FindHeaderColumn(wsMaster, "Medicare Number", compareCol)
    consultantCol = FindHeaderColumn(wsMaster, "Consultant", compareCol)
    postcodeCol = FindHeaderColumn(wsMaster, "Postcode", compareCol)
    dateCol = FindHeaderColumn(wsMaster, "Date", compareCol)
    dobCol = FindHeaderColumn(wsMaster, "Date of Birth", compareCol)

    If bookCol = 0 Then
        MsgBox "Critical: Column 'Book No.' was not found in MASTER header row 11 within columns A to T." & vbCrLf & _
               "Verify the column name and try again.", _
               vbCritical, "Missing Column"
        GoTo CleanFail
    End If
    If patientCol = 0 Then
        MsgBox "Critical: Column 'Patient' was not found in MASTER header row 11 within columns A to T." & vbCrLf & _
               "Verify the column name and try again.", _
               vbCritical, "Missing Column"
        GoTo CleanFail
    End If

    ' --------------------------------------------------------
    ' 13. Refine lastRowMaster using the Book No. column
    ' --------------------------------------------------------
    lastRowMaster = wsMaster.Cells(wsMaster.Rows.Count, bookCol).End(xlUp).Row
    If lastRowMaster < masterDataStart Then
        MsgBox "No booking data found in the MASTER 'Book No.' column.", _
               vbExclamation, "No Data"
        GoTo CleanFail
    End If

    ' --------------------------------------------------------
    ' 14. DUPLICATE BOOK NO. CHECK
    ' --------------------------------------------------------
    Set dupDict = CreateObject("Scripting.Dictionary")
    dupDict.CompareMode = 1   ' vbTextCompare: case-insensitive

    dupMsg = ""
    For i = masterDataStart To lastRowMaster
        dKey = SafeText(wsMaster.Cells(i, bookCol).Value)
        If dKey <> "" Then
            If dupDict.exists(dKey) Then
                dupDict(dKey) = dupDict(dKey) & ", " & i
            Else
                dupDict.Add dKey, CStr(i)
            End If
        End If
    Next i

    Dim dk As Variant
    For Each dk In dupDict.Keys
        If InStr(CStr(dupDict(dk)), ",") > 0 Then
            dupMsg = dupMsg & "  Book No. " & dk & "  ->  Rows: " & dupDict(dk) & vbCrLf
        End If
    Next dk

    If Len(dupMsg) > 0 Then
        MsgBox "Data integrity error: Duplicate 'Book No.' values detected in the MASTER sheet." & vbCrLf & vbCrLf & _
               dupMsg & vbCrLf & _
               "All duplicate Booking Numbers must be resolved before running this macro.", _
               vbCritical, "Duplicate Booking Numbers Detected"
        GoTo CleanFail
    End If

    Set dupDict = Nothing

    ' --------------------------------------------------------
    ' 15. DATA INTEGRITY: Blank Book No. on relevant rows (A:T)
    ' --------------------------------------------------------
    blankRows = ""
    For i = masterDataStart To lastRowMaster
        If SafeText(wsMaster.Cells(i, bookCol).Value) = "" Then
            hasOtherData = False
            For k = 1 To compareCol
                If k <> bookCol Then
                    If SafeText(wsMaster.Cells(i, k).Value) <> "" Then
                        hasOtherData = True
                        Exit For
                    End If
                End If
            Next k
            If hasOtherData Then
                blankRows = blankRows & i & ", "
            End If
        End If
    Next i

    If Len(blankRows) > 0 Then
        blankRows = Left(blankRows, Len(blankRows) - 2)
        MsgBox "Data integrity error: Blank 'Book No.' values found in MASTER rows:" & vbCrLf & _
               blankRows & vbCrLf & vbCrLf & _
               "All blank Book No. values must be resolved before running this macro.", _
               vbCritical, "Data Integrity Error"
        GoTo CleanFail
    End If

    ' --------------------------------------------------------
    ' 16. SETUP REVIEW SHEETS (create or fully clear)
    '     Review sheets only carry the A:T comparison columns
    '     plus the timestamp column.
    ' --------------------------------------------------------
    Set wsMod = SetupReviewSheet(wbMaster, "Modified Bookings to Review", wsMaster, compareCol)
    Set wsCan = SetupReviewSheet(wbMaster, "Cancelled Bookings to Review", wsMaster, compareCol)
    Set wsNew = SetupReviewSheet(wbMaster, "New Bookings to Review", wsMaster, compareCol)

    ' --------------------------------------------------------
    ' 17. CLEAR ALL MASTER DATA HIGHLIGHTS (A:T only, fresh slate)
    ' --------------------------------------------------------
    wsMaster.Range( _
        wsMaster.Cells(masterDataStart, 1), _
        wsMaster.Cells(lastRowMaster, compareCol)).Interior.ColorIndex = xlNone

    ' --------------------------------------------------------
    ' 18. BUILD DAILY DICTIONARY (Book No. -> row number)
    ' --------------------------------------------------------
    Set dailyDict = CreateObject("Scripting.Dictionary")
    dailyDict.CompareMode = 1

    dailyBookCol = FindHeaderColumn(wsDaily, "Book No.", 20)
    If dailyBookCol = 0 Then
        MsgBox "Critical: Column 'Book No.' was not found in the DAILY report header row 11.", _
               vbCritical, "Missing Column"
        GoTo CleanFail
    End If

    lastRowDaily = wsDaily.Cells(wsDaily.Rows.Count, dailyBookCol).End(xlUp).Row

    For i = dailyDataStart To lastRowDaily
        dKey = SafeText(wsDaily.Cells(i, dailyBookCol).Value)
        If dKey <> "" Then
            If Not dailyDict.exists(dKey) Then
                dailyDict.Add dKey, i
            End If
        End If
    Next i

    ' --------------------------------------------------------
    ' 19. BUILD MASTER DICTIONARY (Book No. -> row number)
    ' --------------------------------------------------------
    Set masterDict = CreateObject("Scripting.Dictionary")
    masterDict.CompareMode = 1

    For i = masterDataStart To lastRowMaster
        dKey = SafeText(wsMaster.Cells(i, bookCol).Value)
        If dKey <> "" Then
            If Not masterDict.exists(dKey) Then
                masterDict.Add dKey, i
            End If
        End If
    Next i

    ' --------------------------------------------------------
    ' 20. RECONCILIATION LOOP (comparison uses A:T = compareCol)
    ' --------------------------------------------------------
    totalProcessed = 0
    modifiedCount = 0
    cancelledCount = 0
    newCount = 0
    totalRows = lastRowMaster - masterDataStart + 1

    For i = masterDataStart To lastRowMaster

        bookNo = SafeText(wsMaster.Cells(i, bookCol).Value)
        If bookNo = "" Then GoTo NextMasterRow

        totalProcessed = totalProcessed + 1

        If totalRows > 0 Then
            pct = CLng((CDbl(totalProcessed) / CDbl(totalRows)) * 100)
        Else
            pct = 100
        End If
        Application.StatusBar = "Reconciling... " & pct & "% " & _
                                 "(" & totalProcessed & " of " & totalRows & " rows)"

        If Not dailyDict.exists(bookNo) Then

            ' CANCELLED: Book No. absent from DAILY
            wsMaster.Range( _
                wsMaster.Cells(i, 1), _
                wsMaster.Cells(i, compareCol)).Interior.Color = vbRed

            Call CopyRowWithTimestamp(wsMaster, wsCan, i, runTime, _
                                      startTimeCol, compareCol, dateCol, dobCol)
            cancelledCount = cancelledCount + 1

        Else

            ' EXISTS IN DAILY: compare cell by cell (A:T only)
            dailyRow = CLng(dailyDict(bookNo))
            modified = False

            For j = 1 To compareCol

                isExcluded = False
                If medicareCol > 0 And j = medicareCol Then isExcluded = True
                If consultantCol > 0 And j = consultantCol Then isExcluded = True
                If postcodeCol > 0 And j = postcodeCol Then isExcluded = True
                If isExcluded Then GoTo NextCompareCol

                masterVal = SafeText(wsMaster.Cells(i, j).Value)
                dailyVal = SafeText(wsDaily.Cells(dailyRow, j).Value)

                isDiff = False
                If j = patientCol Then
                    If LCase(masterVal) <> LCase(dailyVal) Then isDiff = True
                ElseIf startTimeCol > 0 And j = startTimeCol Then
                    ' Normalise both to hh:mm so a true time value and a
                    ' text time (e.g. "10:45") are treated as equal.
                    If NormaliseTime(wsMaster.Cells(i, j).Value) <> _
                       NormaliseTime(wsDaily.Cells(dailyRow, j).Value) Then isDiff = True
                ElseIf (dateCol > 0 And j = dateCol) Or (dobCol > 0 And j = dobCol) Then
                    ' Normalise both to dd/mm/yyyy so a genuine date value
                    ' and a text date are treated as equal (V3.1).
                    If NormaliseDate(wsMaster.Cells(i, j).Value) <> _
                       NormaliseDate(wsDaily.Cells(dailyRow, j).Value) Then isDiff = True
                Else
                    If masterVal <> dailyVal Then isDiff = True
                End If

                If isDiff Then
                    wsMaster.Cells(i, j).Interior.Color = vbYellow
                    modified = True
                End If

NextCompareCol:
            Next j

            If modified Then
                Call CopyModifiedRow( _
                    wsDaily, wsMod, dailyRow, wsMaster, i, _
                    patientCol, runTime, startTimeCol, compareCol, _
                    medicareCol, consultantCol, postcodeCol, dateCol, dobCol)
                modifiedCount = modifiedCount + 1
            End If

        End If

NextMasterRow:
    Next i

    ' --------------------------------------------------------
    ' 21. NEW BOOKINGS: DAILY rows absent from MASTER
    ' --------------------------------------------------------
    Application.StatusBar = "Checking for new bookings in Daily report..."

    For i = dailyDataStart To lastRowDaily
        dKey = SafeText(wsDaily.Cells(i, dailyBookCol).Value)
        If dKey <> "" Then
            If Not masterDict.exists(dKey) Then
                Call CopyRowWithTimestamp(wsDaily, wsNew, i, runTime, _
                                          startTimeCol, compareCol, dateCol, dobCol)
                newCount = newCount + 1
            End If
        End If
    Next i

    ' --------------------------------------------------------
    ' 22. AUDIT LOG
    ' --------------------------------------------------------
    Set wsAudit = GetOrCreateAuditSheet(wbMaster)
    wsAudit.Unprotect

    auditNextRow = wsAudit.Cells(wsAudit.Rows.Count, 1).End(xlUp).Row + 1
    With wsAudit
        .Cells(auditNextRow, 1).Value = runTime
        .Cells(auditNextRow, 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
        .Cells(auditNextRow, 2).Value = Environ("USERNAME")
        .Cells(auditNextRow, 3).Value = dailyPath
        .Cells(auditNextRow, 4).Value = totalProcessed
        .Cells(auditNextRow, 5).Value = modifiedCount
        .Cells(auditNextRow, 6).Value = cancelledCount
        .Cells(auditNextRow, 7).Value = newCount
    End With

    wsAudit.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True

    ' --------------------------------------------------------
    ' 23. CLOSE DAILY WORKBOOK (without saving)
    ' --------------------------------------------------------
    wbDaily.Close SaveChanges:=False
    Set wbDaily = Nothing

    ' --------------------------------------------------------
    ' 24. SUCCESS MESSAGE
    ' --------------------------------------------------------
    MsgBox "Reconciliation Complete" & vbCrLf & vbCrLf & _
           "Total bookings processed:     " & totalProcessed & vbCrLf & _
           "Modified bookings to review:  " & modifiedCount & vbCrLf & _
           "Cancelled bookings to review: " & cancelledCount & vbCrLf & _
           "New bookings to review:       " & newCount, _
           vbInformation, "Reconciliation Summary - V3.1"

    GoTo CleanExit

' ============================================================
'  ERROR HANDLING
'  Captures the real error number/description FIRST, before any
'  cleanup (which uses On Error Resume Next and would clear Err).
' ============================================================
CleanFail:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description

    If Not wbDaily Is Nothing Then
        On Error Resume Next
        wbDaily.Close SaveChanges:=False
        On Error GoTo 0
        Set wbDaily = Nothing
    End If
    On Error Resume Next
    If Not wsAudit Is Nothing Then
        wsAudit.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    End If
    On Error GoTo 0

    If errNum <> 0 Then
        MsgBox "An unexpected error occurred during reconciliation." & vbCrLf & vbCrLf & _
               "Error " & errNum & ": " & errDesc & vbCrLf & vbCrLf & _
               "The Daily file has not been saved or modified.", _
               vbCritical, "Macro Error"
    End If

CleanExit:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

End Sub


' ============================================================
'  ColLetter
'  Converts a 1-based column number to its Excel letter(s).
'  Used in header mismatch and table failure messages.
' ============================================================
Public Function ColLetter(colNum As Long) As String
    Dim n As Long, s As String
    n = colNum
    Do While n > 0
        s = Chr(65 + ((n - 1) Mod 26)) & s
        n = (n - 1) \ 26
    Loop
    ColLetter = s
End Function


' ============================================================
'  StandardiseFont
'  Sets all fonts to Calibri from row 2 through last used row.
'  Scoped to columns 1 through colLimit (the A:T compare range).
' ============================================================
Public Sub StandardiseFont(ws As Worksheet, colLimit As Long)
    Dim lastR As Long
    On Error Resume Next
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    On Error GoTo 0
    If lastR < 2 Or colLimit < 1 Then Exit Sub
    ws.Range(ws.Cells(2, 1), ws.Cells(lastR, colLimit)).Font.Name = "Calibri"
End Sub


' ============================================================
'  SafeText
'  Safe variant-to-string conversion.
'  Returns "" for Null, Empty, or cell error values.
' ============================================================
Public Function SafeText(v As Variant) As String
    On Error Resume Next
    If IsNull(v) Or IsEmpty(v) Or IsError(v) Then
        SafeText = ""
    Else
        SafeText = CStr(v)
    End If
    On Error GoTo 0
End Function


' ============================================================
'  NormaliseTime
'  Returns a time normalised to "hh:mm" so a genuine time value
'  and a text representation (e.g. "10:45") are treated as equal.
'  Only the time-of-day portion is used; any date part is
'  stripped. Values that cannot be interpreted as a time are
'  returned as trimmed text (i.e. compared as-is).
' ============================================================
Public Function NormaliseTime(v As Variant) As String
    Dim t      As Double
    Dim s      As String
    Dim parsed As Date

    ' Null / Empty / cell error -> empty string
    On Error Resume Next
    If IsNull(v) Or IsEmpty(v) Or IsError(v) Then
        NormaliseTime = ""
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    Select Case VarType(v)

        Case vbDate
            t = CDbl(v)
            t = t - Int(t)                 ' keep time-of-day only
            NormaliseTime = Format(t, "hh:mm")

        Case vbSingle, vbDouble, vbInteger, vbLong, vbCurrency, vbDecimal
            t = CDbl(v)
            t = t - Int(t)
            NormaliseTime = Format(t, "hh:mm")

        Case vbString
            s = Trim(CStr(v))
            If s = "" Then
                NormaliseTime = ""
                Exit Function
            End If
            On Error Resume Next
            parsed = TimeValue(s)          ' handles "10:45", "10:45 AM"
            If Err.Number = 0 Then
                NormaliseTime = Format(parsed, "hh:mm")
                On Error GoTo 0
                Exit Function
            End If
            Err.Clear
            parsed = CDate(s)              ' handles full date/time text
            If Err.Number = 0 Then
                t = CDbl(parsed)
                t = t - Int(t)
                NormaliseTime = Format(t, "hh:mm")
                On Error GoTo 0
                Exit Function
            End If
            On Error GoTo 0
            ' Unparseable - fall back to raw trimmed text
            NormaliseTime = s

        Case Else
            NormaliseTime = Trim(SafeText(v))

    End Select
End Function


' ============================================================
'  NormaliseDate
'  Returns a date normalised to "dd/mm/yyyy" so a genuine Excel
'  date value and a text date (e.g. "01/09/2026") are treated as
'  equal for COMPARISON purposes only (does not alter any cell).
'
'  Rules (clinical-integrity conscious - never guesses):
'   - Genuine date value -> Format dd/mm/yyyy (date part only).
'   - Numeric serial      -> treated as a date serial, dd/mm/yyyy.
'   - Text                -> parsed EXPLICITLY as day/month/year
'                            (Australian order), NOT via locale.
'                            Requires a 4-digit year, valid month
'                            (1-12) and day (1-31), and a real
'                            calendar date (rejects 31/02 etc.).
'   - Anything that cannot be safely parsed as an unambiguous
'     dd/mm/yyyy date falls back to the trimmed/normalised text,
'     so the comparison degrades to a literal string match (no
'     worse than before, and no incorrect date assumptions).
' ============================================================
Public Function NormaliseDate(v As Variant) As String
    Dim s     As String
    Dim parts As Variant
    Dim dd    As Long
    Dim mm    As Long
    Dim yy    As Long
    Dim tmp   As Date

    ' Null / Empty / cell error -> empty string
    On Error Resume Next
    If IsNull(v) Or IsEmpty(v) Or IsError(v) Then
        NormaliseDate = ""
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    Select Case VarType(v)

        Case vbDate
            NormaliseDate = Format(v, "dd/mm/yyyy")

        Case vbSingle, vbDouble, vbInteger, vbLong, vbCurrency, vbDecimal
            On Error Resume Next
            NormaliseDate = Format(CDate(Int(CDbl(v))), "dd/mm/yyyy")
            If Err.Number <> 0 Then NormaliseDate = Trim(CStr(v))
            On Error GoTo 0

        Case vbString
            s = Trim(CStr(v))
            If s = "" Then
                NormaliseDate = ""
                Exit Function
            End If

            ' Keep only the date portion (drop any trailing time text)
            If InStr(s, " ") > 0 Then s = Trim(Left(s, InStr(s, " ") - 1))

            ' Normalise common delimiters to "/"
            s = Replace(s, "-", "/")
            s = Replace(s, ".", "/")

            parts = Split(s, "/")
            If UBound(parts) = 2 Then
                If IsNumeric(parts(0)) And IsNumeric(parts(1)) And IsNumeric(parts(2)) Then
                    dd = CLng(parts(0))
                    mm = CLng(parts(1))
                    yy = CLng(parts(2))
                    ' Only accept an unambiguous dd/mm/yyyy text date:
                    ' 4-digit year, month 1-12, day 1-31, real calendar date.
                    If Len(Trim(CStr(parts(2)))) = 4 And _
                       mm >= 1 And mm <= 12 And dd >= 1 And dd <= 31 Then
                        On Error Resume Next
                        tmp = DateSerial(yy, mm, dd)
                        If Err.Number = 0 Then
                            If Day(tmp) = dd And Month(tmp) = mm And Year(tmp) = yy Then
                                NormaliseDate = Format(tmp, "dd/mm/yyyy")
                                On Error GoTo 0
                                Exit Function
                            End If
                        End If
                        On Error GoTo 0
                    End If
                End If
            End If

            ' Could not safely parse -> literal text fallback (normalised)
            NormaliseDate = s

        Case Else
            NormaliseDate = Trim(SafeText(v))

    End Select
End Function


' ============================================================
'  CopyDateAsText
'  Writes a date value to a destination cell as TEXT, verbatim,
'  with NO coercion - protecting against Excel reinterpreting a
'  dd/mm/yyyy text string as a US (mm/dd/yyyy) real date on copy.
'
'  - Destination cell is pre-formatted as Text ("@") so Excel
'    cannot convert the incoming string to a serial date.
'  - If the SOURCE is already text, the exact characters are
'    copied across unchanged (the normal case for these files).
'  - If the SOURCE is a genuine Excel date value, it is written
'    as dd/mm/yyyy text explicitly (locale-proof), rather than
'    letting Excel/locale decide the order.
'  - Blank / error source cells produce an empty destination.
' ============================================================
Public Sub CopyDateAsText(src As Worksheet, srcRow As Long, srcCol As Long, _
                          dest As Worksheet, destRow As Long, destCol As Long)
    Dim v As Variant

    ' Force the destination cell to Text BEFORE writing anything.
    dest.Cells(destRow, destCol).NumberFormat = "@"

    v = src.Cells(srcRow, srcCol).Value

    If IsError(v) Then
        dest.Cells(destRow, destCol).Value = ""
    ElseIf IsNull(v) Or IsEmpty(v) Then
        dest.Cells(destRow, destCol).Value = ""
    ElseIf VarType(v) = vbDate Then
        ' Genuine date value -> explicit dd/mm/yyyy text (locale-proof)
        dest.Cells(destRow, destCol).Value = Format(v, "dd/mm/yyyy")
    Else
        ' Text (normal case) or other -> copy exact characters, no coercion.
        ' Leading apostrophe is NOT added to the value; the "@" format
        ' alone keeps it as text without altering the displayed string.
        dest.Cells(destRow, destCol).Value = CStr(v)
    End If
End Sub


' ============================================================
'  FindHeaderColumn
'  Searches row 11 up to colLimit for headerName (case-insensitive).
'  Returns the column index, or 0 if not found within the limit.
' ============================================================
Public Function FindHeaderColumn(ws As Worksheet, headerName As String, _
                                  colLimit As Long) As Long
    Dim c As Long
    FindHeaderColumn = 0
    For c = 1 To colLimit
        If LCase(SafeText(ws.Cells(11, c).Value)) = LCase(headerName) Then
            FindHeaderColumn = c
            Exit Function
        End If
    Next c
End Function


' ============================================================
'  SetupReviewSheet
'  Creates or fully clears the named review sheet in wb.
'  Copies MASTER headers (row 11, cols 1 to compareCol) to row 1.
'  Appends "Dt/Tm Added by Macro" as the final column header.
'  Data rows always start at row 2.
' ============================================================
Public Function SetupReviewSheet(wb As Workbook, sheetName As String, _
                                  wsMaster As Worksheet, compareCol As Long) As Worksheet
    Dim ws     As Worksheet
    Dim s      As Worksheet
    Dim exists As Boolean
    Dim j      As Long

    exists = False
    For Each s In wb.Worksheets
        If s.Name = sheetName Then
            Set ws = s
            exists = True
            Exit For
        End If
    Next s

    If Not exists Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If

    For j = 1 To compareCol
        ws.Cells(1, j).Value = SafeText(wsMaster.Cells(11, j).Value)
    Next j

    ws.Cells(1, compareCol + 1).Value = "Dt/Tm Added by Macro"
    ws.Rows(1).Font.Bold = True

    Set SetupReviewSheet = ws
End Function


' ============================================================
'  CopyRowWithTimestamp
'  Copies srcRow from src to dest as values only (A:T only).
'  Date and Date of Birth columns are copied as verbatim TEXT
'  via CopyDateAsText (prevents dd/mm <-> mm/dd corruption).
'  Applies hh:mm format to Start Time column if found.
'  Appends run timestamp in the "Dt/Tm Added by Macro" column.
'  Used for Cancelled Bookings and New Bookings sheets.
' ============================================================
Public Sub CopyRowWithTimestamp(src As Worksheet, dest As Worksheet, _
                                 srcRow As Long, ts As Date, _
                                 startTimeCol As Long, compareCol As Long, _
                                 dateCol As Long, dobCol As Long)
    Dim nextRow As Long
    Dim j       As Long
    Dim cellVal As Variant

    nextRow = dest.Cells(dest.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    For j = 1 To compareCol

        If (dateCol > 0 And j = dateCol) Or (dobCol > 0 And j = dobCol) Then
            ' Date columns: verbatim text copy, no coercion.
            Call CopyDateAsText(src, srcRow, j, dest, nextRow, j)
        Else
            cellVal = src.Cells(srcRow, j).Value
            If IsError(cellVal) Then
                dest.Cells(nextRow, j).Value = ""
            Else
                dest.Cells(nextRow, j).Value = cellVal
            End If
            If j = startTimeCol And startTimeCol > 0 Then
                dest.Cells(nextRow, j).NumberFormat = "hh:mm"
            End If
        End If

    Next j

    dest.Cells(nextRow, compareCol + 1).Value = ts
    dest.Cells(nextRow, compareCol + 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
End Sub


' ============================================================
'  CopyModifiedRow
'  Copies dailyRow from wsDaily into wsMod (values only, A:T).
'  Date and Date of Birth columns are copied as verbatim TEXT
'  via CopyDateAsText (prevents dd/mm <-> mm/dd corruption).
'  Applies hh:mm format to Start Time column if found.
'  Highlights yellow in review sheet only where DAILY differs
'  from MASTER - applying same Patient case rule, Start Time
'  normalisation, and Date/DOB normalisation as the main loop.
'  Excluded columns (Medicare Number, Consultant, Postcode)
'  are copied but never highlighted yellow in the review sheet.
'  Appends run timestamp in the "Dt/Tm Added by Macro" column.
' ============================================================
Public Sub CopyModifiedRow(wsDaily As Worksheet, wsMod As Worksheet, _
                            dailyRow As Long, wsMaster As Worksheet, _
                            masterRow As Long, patientCol As Long, _
                            ts As Date, startTimeCol As Long, compareCol As Long, _
                            medicareCol As Long, consultantCol As Long, _
                            postcodeCol As Long, dateCol As Long, dobCol As Long)
    Dim nextRow    As Long
    Dim j          As Long
    Dim cellVal    As Variant
    Dim masterVal  As String
    Dim dailyVal   As String
    Dim isDiff     As Boolean
    Dim isExcluded As Boolean

    nextRow = wsMod.Cells(wsMod.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    For j = 1 To compareCol

        If (dateCol > 0 And j = dateCol) Or (dobCol > 0 And j = dobCol) Then
            ' Date columns: verbatim text copy from DAILY, no coercion.
            Call CopyDateAsText(wsDaily, dailyRow, j, wsMod, nextRow, j)
        Else
            cellVal = wsDaily.Cells(dailyRow, j).Value
            If IsError(cellVal) Then
                wsMod.Cells(nextRow, j).Value = ""
            Else
                wsMod.Cells(nextRow, j).Value = cellVal
            End If
            If j = startTimeCol And startTimeCol > 0 Then
                wsMod.Cells(nextRow, j).NumberFormat = "hh:mm"
            End If
        End If

        isExcluded = False
        If medicareCol > 0 And j = medicareCol Then isExcluded = True
        If consultantCol > 0 And j = consultantCol Then isExcluded = True
        If postcodeCol > 0 And j = postcodeCol Then isExcluded = True

        If Not isExcluded Then
            masterVal = SafeText(wsMaster.Cells(masterRow, j).Value)
            dailyVal = SafeText(wsDaily.Cells(dailyRow, j).Value)

            isDiff = False
            If j = patientCol Then
                If LCase(masterVal) <> LCase(dailyVal) Then isDiff = True
            ElseIf startTimeCol > 0 And j = startTimeCol Then
                If NormaliseTime(wsMaster.Cells(masterRow, j).Value) <> _
                   NormaliseTime(wsDaily.Cells(dailyRow, j).Value) Then isDiff = True
            ElseIf (dateCol > 0 And j = dateCol) Or (dobCol > 0 And j = dobCol) Then
                If NormaliseDate(wsMaster.Cells(masterRow, j).Value) <> _
                   NormaliseDate(wsDaily.Cells(dailyRow, j).Value) Then isDiff = True
            Else
                If masterVal <> dailyVal Then isDiff = True
            End If

            If isDiff Then
                wsMod.Cells(nextRow, j).Interior.Color = vbYellow
            End If
        End If

    Next j

    wsMod.Cells(nextRow, compareCol + 1).Value = ts
    wsMod.Cells(nextRow, compareCol + 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
End Sub


' ============================================================
'  GetOrCreateAuditSheet
'  Returns the visible, protected audit log sheet.
'  Creates and formats it on first run if it does not exist.
'  Main sub unprotects before writing and re-protects after.
' ============================================================
Public Function GetOrCreateAuditSheet(wb As Workbook) As Worksheet
    Const AUDIT_NAME As String = "Reconcile Audit Log"
    Dim ws As Worksheet
    Dim s  As Worksheet

    For Each s In wb.Worksheets
        If s.Name = AUDIT_NAME Then
            Set GetOrCreateAuditSheet = s
            Exit Function
        End If
    Next s

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = AUDIT_NAME
    ws.Visible = xlSheetVisible

    With ws
        .Cells(1, 1).Value = "Run Date/Time"
        .Cells(1, 2).Value = "Run By (Username)"
        .Cells(1, 3).Value = "Daily File Path"
        .Cells(1, 4).Value = "Total Processed"
        .Cells(1, 5).Value = "Modified"
        .Cells(1, 6).Value = "Cancelled"
        .Cells(1, 7).Value = "New Bookings"
        .Rows(1).Font.Bold = True
        .Columns(1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
        .Columns(1).ColumnWidth = 20
        .Columns(2).ColumnWidth = 22
        .Columns(3).ColumnWidth = 80
        .Columns(4).ColumnWidth = 16
        .Columns(5).ColumnWidth = 12
        .Columns(6).ColumnWidth = 12
        .Columns(7).ColumnWidth = 14
    End With

    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True

    Set GetOrCreateAuditSheet = ws
End Function