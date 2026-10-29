Attribute VB_Name = "HicAutomation"
Option Explicit

' CONFIGURATION. Import this module into every monthly workbook after review.
' Launcher reads NETWORK_FOLDER, YEAR_SUBFOLDERS and REFRESH_TIMEOUT_SECONDS
' from this exported module and checks them against the workbook at runtime.
Public Const NETWORK_FOLDER As String = "<INTERNAL_NETWORK_PATH>"
Public Const YEAR_SUBFOLDERS As Boolean = True
Public Const REFRESH_TIMEOUT_SECONDS As Long = 900
Public Const LOG_FOLDER As String = "" ' Blank: shared .hic-state folder
Public Const EMAIL_TO As String = "<STAKEHOLDER_DISTRIBUTION_LIST>"
Public Const EMAIL_CC As String = "<CC_RECIPIENTS>"
Public Const EMAIL_FROM As String = "<SENDER_ADDRESS>"
Public Const EMAIL_SUBJECT As String = "HIC Update - {MONTH}"
Public Const EMAIL_BODY As String = "Hi All,{NL}{NL}Please find attached the updated HIC file for the period {START} to {END}.{NL}{NL}Regards,{NL}BR Support"
Public Const AUTO_SEND As Boolean = True
Public Const SCHEDULE_TIME_ZONE As String = "India Standard Time"
' IT must provision a least-privilege Windows-authenticated SQL account/driver.
' Never put a password in this distributable workbook or environment variable.
Public Const SQL_CONNECTION As String = "<APPROVED_SQL_CONNECTION_STRING>"
Public Const VERSION As String = "HIC-2026-09-03.1"

Private mActive As Boolean
Private mRows As Long
Private mDeadline As Date
Private mLog As String
Private mPhase As String

Public Function HicConfiguration() As String
    HicConfiguration = VERSION & "|" & NETWORK_FOLDER & "|" & CStr(YEAR_SUBFOLDERS) & "|" & CStr(REFRESH_TIMEOUT_SECONDS)
End Function

Public Function HicSqlConnectionString() As String
    HicSqlConnectionString = SQL_CONNECTION
End Function

Private Sub Fail(ByVal description As String)
    Err.Raise vbObjectError + 2100, "HIC", description
End Sub

Private Sub CheckDeadline()
    If Now >= mDeadline Then Fail "Refresh timeout exceeded."
End Sub

Private Function MonthLabel(ByVal d As Date) As String
    Dim months As Variant
    months = Split("January February March April May June July August September October November December", " ")
    MonthLabel = months(Month(d) - 1) & " " & CStr(Year(d))
End Function

Private Function ParseDateISO(ByVal value As String) As Date
    Dim d As Date
    If Not value Like "####-##-##" Then Fail "Run date must be yyyy-mm-dd."
    d = DateSerial(CInt(Left$(value, 4)), CInt(Mid$(value, 6, 2)), CInt(Right$(value, 2)))
    If Format$(d, "yyyy-mm-dd") <> value Then Fail "Invalid run date."
    ParseDateISO = d
End Function

Private Sub LogLine(ByVal message As String)
    Dim f As Integer
    f = FreeFile
    Open mLog For Append Access Write Lock Write As #f
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & Environ$("COMPUTERNAME") & vbTab & Environ$("USERNAME") & vbTab & mPhase & vbTab & Replace(Replace(message, vbCr, " "), vbLf, " ")
    Close #f
End Sub

Private Sub EnsureFolder(ByVal folder As String)
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Not fs.FolderExists(folder) Then
        On Error Resume Next
        fs.CreateFolder folder
        On Error GoTo 0
        If Not fs.FolderExists(folder) Then Fail "Cannot create/access folder: " & folder
    End If
End Sub

' Entry point is intended to be called by Launch-Hic.ps1, including manual runs.
' Mode Test: saves updated workbook but creates no email. Review: displays email.
' Mode Scheduled: uses AUTO_SEND. Every mode validates the complete report.
Public Function HicRun(ByVal runISO As String, ByVal mode As String) As String
    Dim wb As Workbook, ws As Worksheet, lo As ListObject, pt As PivotTable
    Dim runDate As Date, startDate As Date, endDate As Date
    Dim oldScreen As Boolean, oldAlerts As Boolean, oldEvents As Boolean, oldCalc As XlCalculation
    Dim lockNumber As Integer, lockHeld As Boolean, stateFolder As String, intent As String
    Dim fs As Object, marker As Object, ol As Object, mail As Object, attachment As Object
    Dim errorNumber As Long, errorText As String, expectedFolder As String, buttonAction As String
    If mActive Then Fail "A run is already active in this workbook."
    oldScreen = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts
    oldEvents = Application.EnableEvents
    oldCalc = Application.Calculation
    On Error GoTo Failed
    mActive = True
    Set wb = ThisWorkbook
    mPhase = "START"
    mLog = ""
    runDate = ParseDateISO(runISO)
    If mode <> "Test" And mode <> "Review" And mode <> "Scheduled" Then Fail "Unsupported mode."
    startDate = DateSerial(Year(runDate), Month(runDate) - 1, 15)
    endDate = DateSerial(Year(runDate), Month(runDate) + 1, 0)
    If startDate >= endDate Then Fail "Invalid date range."
    mDeadline = DateAdd("s", REFRESH_TIMEOUT_SECONDS, Now)
    If Left$(NETWORK_FOLDER, 2) <> "\\" Then Fail "A UNC report folder is required."
    expectedFolder = NETWORK_FOLDER
    If YEAR_SUBFOLDERS Then expectedFolder = expectedFolder & "\" & CStr(Year(runDate))
    If StrComp(wb.Path, expectedFolder, vbTextCompare) <> 0 Then Fail "Workbook is outside the configured monthly folder."
    If StrComp(wb.Name, "Follow Up Resultados HIC_Model-" & MonthLabel(runDate) & ".xlsm", vbTextCompare) <> 0 _
        And StrComp(wb.Name, "Follow Up Resultados HIC_Model-" & Replace(MonthLabel(runDate), " ", "-") & ".xlsm", vbTextCompare) <> 0 Then Fail "Wrong monthly workbook."
    If wb.ReadOnly Or wb.FileFormat <> xlOpenXMLWorkbookMacroEnabled Then Fail "Workbook is read-only or is not .xlsm."
    ' Inspected workbook uses ADODB, not workbook connections or Power Query.
    ' Fail closed if its data architecture changes: idle is not proof of success.
    If wb.Connections.count <> 0 Then Fail "Workbook connections were added; refresh validation must be reviewed."
    stateFolder = NETWORK_FOLDER & "\.hic-state"
    EnsureFolder stateFolder
    If LOG_FOLDER = "" Then
        mLog = stateFolder
    Else
        EnsureFolder LOG_FOLDER
        mLog = LOG_FOLDER
    End If
    mLog = mLog & "\HIC-" & runISO & "-" & Environ$("COMPUTERNAME") & "-" & Format$(Now, "hhnnss") & ".log"
    LogLine "Workbook=" & wb.FullName & "; RunDate=" & runISO & "; Start=" & Format$(startDate, "mm/dd/yyyy") & "; End=" & Format$(endDate, "mm/dd/yyyy") & "; Mode=" & mode
    ' SMB byte/file locking is shared across machines. Never delete this lock file.
    lockNumber = FreeFile
    Open stateFolder & "\HIC.lock" For Binary Access Read Write Lock Read Write As #lockNumber
    lockHeld = True
    intent = stateFolder & "\HIC-" & runISO & ".intent"
    Set fs = CreateObject("Scripting.FileSystemObject")
    If mode <> "Test" And fs.FileExists(intent) Then Fail "An email attempt/review already exists for this run date. Reconcile its status before retrying."
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Set ws = wb.Worksheets("Query")
    If ws.QueryTables.count <> 0 Then Fail "QueryTables were added; refresh validation must be reviewed."
    If ws.ProtectContents Then Fail "Query sheet is protected."
    If ws.Range("I4").MergeCells Or ws.Range("J4").MergeCells Then Fail "Date cells must be unmerged."
    If ws.Range("I4").HasFormula Or ws.Range("J4").HasFormula Then Fail "Date cells unexpectedly contain formulas."
    If ws.Range("H6").Value2 <> 2 Then Fail "Expected Sample Created Date (Query!H6=2), matching the inspected workbook."
    buttonAction = ws.Shapes("Button 1").OnAction
    If InStrRev(buttonAction, "!") > 0 Then buttonAction = Mid$(buttonAction, InStrRev(buttonAction, "!") + 1)
    If StrComp(buttonAction, "ThisWorkbook.MainSub", vbTextCompare) <> 0 Then Fail "Get Responses button assignment changed; review required."
    ' .Value receives a VBA Date, so Excel handles either date system correctly.
    ws.Range("I4").value = startDate
    ws.Range("J4").value = endDate
    ws.Range("I4:J4").NumberFormat = "mm/dd/yyyy"
    ' Preserve the visible effect of the existing Worksheet_Change handler.
    ws.Range("I7:J7").Font.Strikethrough = True
    wb.Activate ' Existing SQL builder uses unqualified Sheets references.
    mRows = 0
    mPhase = "REFRESH"
    ' Public bridge invokes the exact private procedure assigned to the button.
    Application.Run "'" & Replace(wb.Name, "'", "''") & "'!ThisWorkbook.HicGetResponses"
    WaitForRefresh wb
    Set lo = ws.ListObjects("Query")
    If mRows < 1 Then Fail "No successful SQL extraction was recorded. Install the InjectSQL patch."
    If lo.DataBodyRange Is Nothing Then Fail "Result table is empty."
    If lo.ListRows.count <> mRows Then Fail "Result row count does not match the SQL extraction."
    If Application.WorksheetFunction.CountA(lo.ListColumns(3).DataBodyRange) <> mRows Then Fail "Sample IDs are missing."
    Application.CalculateFull
    WaitForRefresh wb
    CheckErrors lo.DataBodyRange, "Query results (including SAP lookups)"
    Set pt = wb.Worksheets("Pivot").PivotTables("PivotTable1")
    If pt.PivotCache.SourceType <> xlDatabase Then Fail "Unexpected PivotTable source type."
    ' Preserve the existing cache, fields and filters; use the dynamic result table.
    pt.PivotCache.SourceData = lo.Name
    pt.PivotCache.Refresh
    If Not pt.RefreshTable Then Fail "PivotTable1 refresh returned failure."
    If pt.PivotCache.recordCount <> mRows Then Fail "Pivot cache record count differs from refreshed SQL results."
    WaitForRefresh wb
    ' Existing adjacent status formulas end at J897. Do not silently omit new rows.
    If pt.TableRange2.Row + pt.TableRange2.Rows.count - 1 > 897 Then Fail "Pivot exceeds the existing status formula area; extend it before deployment."
    Application.CalculateFull
    WaitForRefresh wb
    CheckErrors wb.Worksheets("Pivot").UsedRange, "Pivot report"
    LogLine "Refresh and validation succeeded; Records=" & CStr(mRows)
    mPhase = "SAVE"
    CheckDeadline
    wb.Save
    If Not wb.Saved Then Fail "Excel did not confirm the workbook was saved."
    If Not fs.FileExists(wb.FullName) Then Fail "Saved workbook is missing."
    If fs.GetFile(wb.FullName).Size = 0 Then Fail "Saved workbook is empty."
    LogLine "Saved .xlsm successfully."
    If mode = "Test" Then
        HicRun = "TEST_OK"
        GoTo Finished
    End If
    mPhase = "EMAIL"
    If Len(Trim$(EMAIL_TO)) = 0 Then Fail "To recipients are missing."
    Set ol = CreateObject("Outlook.Application")
    Set mail = ol.CreateItem(0)
    With mail
        .To = EMAIL_TO
        .CC = EMAIL_CC
        If EMAIL_FROM <> "" Then .SentOnBehalfOfName = EMAIL_FROM
        .Subject = Replace(EMAIL_SUBJECT, "{MONTH}", MonthLabel(runDate))
        .body = Replace(Replace(Replace(EMAIL_BODY, "{NL}", vbCrLf), "{START}", Format$(startDate, "mm/dd/yyyy")), "{END}", Format$(endDate, "mm/dd/yyyy"))
        If Not .Recipients.ResolveAll Then Fail "One or more Outlook recipients could not be resolved."
        Set attachment = .Attachments.Add(wb.FullName, 1)
        If .Attachments.count <> 1 Or attachment.Size < 1 Then Fail "Attachment verification failed."
        .Save
    End With
    CheckDeadline
    ' Durable intent BEFORE Send: never automatically retry an uncertain send.
    ' A crash may suppress an unsent email; operator reconciliation is required.
    Set marker = fs.CreateTextFile(intent, False, False)
    marker.WriteLine "PREPARED " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " Mode=" & mode & " AutoSend=" & CStr(AUTO_SEND)
    marker.WriteLine "Workbook=" & wb.FullName
    marker.WriteLine "OutlookEntryID=" & mail.EntryID
    marker.Close
    Set marker = Nothing
    If AUTO_SEND And mode = "Scheduled" Then
        mail.Send
        HicRun = "SUBMITTED"
        LogLine "Outlook accepted Send; delivery is not confirmed."
    Else
        mail.Display False
        HicRun = "DISPLAYED"
        LogLine "Displayed for review; automation did not send."
    End If
    Set marker = fs.OpenTextFile(intent, 8, False)
    marker.WriteLine HicRun & " " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    marker.Close
    Set marker = Nothing
Finished:
    mPhase = "END"
    LogLine "Completed: " & HicRun
    GoTo Cleanup
Failed:
    errorNumber = Err.Number
    errorText = Err.description
    On Error Resume Next
    If mLog <> "" Then LogLine "FAILED: " & CStr(errorNumber) & " " & errorText
    On Error GoTo 0
Cleanup:
    On Error Resume Next
    If Not marker Is Nothing Then marker.Close
    Set attachment = Nothing
    Set mail = Nothing
    Set ol = Nothing ' Never quit the user's Outlook session.
    Set fs = Nothing
    Set pt = Nothing
    Set lo = Nothing
    Set ws = Nothing
    Set wb = Nothing
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Application.ScreenUpdating = oldScreen
    If lockHeld Then Close #lockNumber
    mActive = False
    On Error GoTo 0
    If errorNumber <> 0 Then Err.Raise errorNumber, "HIC " & mPhase, errorText
End Function

Private Sub WaitForRefresh(ByVal wb As Workbook)
    Dim ws As Worksheet, qt As QueryTable, connection As WorkbookConnection, busy As Boolean
    ' This workbook has no workbook connections. Added connection types require review.
    Do
        CheckDeadline
        busy = False
        For Each connection In wb.Connections
            Select Case connection.Type
                Case xlConnectionTypeOLEDB
                    If connection.OLEDBConnection.Refreshing Then busy = True
                Case xlConnectionTypeODBC
                    If connection.ODBCConnection.Refreshing Then busy = True
                Case Else
                    Fail "Unreviewed connection type: " & connection.Name
            End Select
        Next connection
        For Each ws In wb.Worksheets
            For Each qt In ws.QueryTables
                If qt.Refreshing Then busy = True
            Next qt
        Next ws
        If Application.CalculationState = xlCalculating Then busy = True
        If Not busy Then Exit Do
        DoEvents
    Loop
    ' Can block inside COM; the launcher supplies an independent hard timeout.
    Application.CalculateUntilAsyncQueriesDone
    CheckDeadline
End Sub

Private Sub CheckErrors(ByVal target As Range, ByVal label As String)
    Dim values As Variant, r As Long, c As Long
    values = target.Value2
    If Not IsArray(values) Then
        If IsError(values) Then Fail label & " contains an Excel error."
        Exit Sub
    End If
    For r = LBound(values, 1) To UBound(values, 1)
        For c = LBound(values, 2) To UBound(values, 2)
            If IsError(values(r, c)) Then Fail label & " contains an Excel error at " & target.Cells(r, c).Address(False, False)
        Next c
        If r Mod 1000 = 0 Then CheckDeadline
    Next r
End Sub

' Called by the replacement ThisWorkbook.InjectSQL. SQL construction is unchanged.
' Fetch completely before modifying cells; preserve the nine calculated columns.
Public Sub HicInjectSQL(ByVal sql As String, ByVal rowPos As Integer, ByVal colPos As Integer, ByVal objectName As String)
    Dim cn As Object, rs As Object, lo As ListObject, formulas(32 To 40) As Variant
    Dim copied As Long, n As Long, c As Long, oldRows As Long, e As Long, description As String
    Dim expanding As Range
    On Error GoTo Failed
    If Not mActive Then mDeadline = DateAdd("s", REFRESH_TIMEOUT_SECONDS, Now)
    If rowPos <> 16 Or colPos <> 1 Or objectName <> "Query" Then Fail "Unexpected SQL output destination."
    Set lo = ThisWorkbook.Worksheets("Query").ListObjects(objectName)
    If lo.ListColumns.count <> 40 Or lo.HeaderRowRange.Row <> 15 Then Fail "Query table schema changed."
    If lo.DataBodyRange Is Nothing Then Fail "Query table must retain a formula template row."
    For c = 32 To 40
        If Not lo.ListColumns(c).DataBodyRange.Cells(1, 1).HasFormula Then Fail "Missing formula template in column " & CStr(c)
        formulas(c) = lo.ListColumns(c).DataBodyRange.Cells(1, 1).FormulaR1C1
    Next c
    Set cn = CreateObject("ADODB.Connection")
    cn.ConnectionTimeout = 30
    cn.CommandTimeout = REFRESH_TIMEOUT_SECONDS
    cn.Open SQL_CONNECTION
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3 ' Client cursor: fully materialized before changing table.
    rs.Open sql, cn, 3, 1 ' Static, read-only, synchronous.
    CheckDeadline
    If rs.Fields.count <> 31 Then Fail "SQL did not return the expected 31 source columns."
    If rs.EOF Then Fail "SQL returned zero records; workbook will not be saved or emailed."
    n = rs.recordCount
    If n < 1 Or n > lo.Parent.Rows.count - 15 Then Fail "Invalid result record count."
    oldRows = lo.ListRows.count
    If n > oldRows Then
        Set expanding = lo.Parent.Cells(16 + oldRows, 1).Resize(n - oldRows, 40)
        If Application.WorksheetFunction.CountA(expanding) > 0 Then Fail "Expanding Query table would overwrite existing cells."
    End If
    ' Clear only the prior table body, including stale formula rows when shrinking.
    lo.DataBodyRange.ClearContents
    lo.Resize lo.HeaderRowRange.Resize(n + 1, 40)
    copied = lo.DataBodyRange.Cells(1, 1).CopyFromRecordset(rs, n, 31)
    If copied <> n Then Fail "Incomplete SQL record copy."
    For c = 32 To 40
        lo.ListColumns(c).DataBodyRange.FormulaR1C1 = formulas(c)
    Next c
    CheckDeadline
    mRows = copied
    GoTo Cleanup
Failed:
    e = Err.Number
    ' Do not propagate provider messages which can contain authentication details.
    If Err.source = "HIC" Then
        description = Err.description
    Else
        description = "SQL extraction failed (error " & CStr(e) & "). Ask IT to inspect SQL connectivity and permissions."
    End If
Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State <> 0 Then rs.Close
    End If
    If Not cn Is Nothing Then
        If cn.State <> 0 Then cn.Close
    End If
    Set rs = Nothing
    Set cn = Nothing
    Set expanding = Nothing
    Set lo = Nothing
    On Error GoTo 0
    If e <> 0 Then Err.Raise e, "HIC SQL", description
End Sub
