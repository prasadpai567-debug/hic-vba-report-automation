Attribute VB_Name = "Module4"
Option Explicit

' Import ONLY into your separate HIC UPDATE.xlsm controller workbook.
' All new automation lives here. Calls only the report's existing button macro.
Private Const DEFAULT_ROOT As String = "<INTERNAL_NETWORK_PATH>"
Private Const REFRESH_TIMEOUT_SECONDS As Long = 900
Private Const EMAIL_TO As String = "<STAKEHOLDER_DISTRIBUTION_LIST>"
Private Const EMAIL_CC As String = "<CC_RECIPIENTS>"
Private Const EMAIL_FROM As String = "<SENDER_ADDRESS>"
Private Const EMAIL_SUBJECT As String = "HIC Update"

Private mDeadline As Date
Private Const PANEL As String = "HIC Controller"
Private Const PANEL_ID As String = "HIC Controller - Manual Update"
Private mRunning As Boolean
Private mLogFile As String

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

#If VBA7 Then
    Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (ByRef value As SYSTEMTIME)
    Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As LongPtr, ByRef processId As Long) As Long
#Else
    Private Declare Sub GetSystemTime Lib "kernel32" (ByRef value As SYSTEMTIME)
    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, ByRef processId As Long) As Long
#End If

Private Sub StopWith(ByVal message As String)
    Err.Raise vbObjectError + 2400, "HIC Controller", message
End Sub

Private Function IndiaDate() As Date
    Dim utc As SYSTEMTIME, instant As Date
    GetSystemTime utc
    instant = DateSerial(utc.wYear, utc.wMonth, utc.wDay) + TimeSerial(utc.wHour, utc.wMinute, utc.wSecond)
    IndiaDate = DateValue(DateAdd("n", 330, instant))
End Function

Private Function EnglishMonth(ByVal value As Date) As String
    Dim names As Variant
    names = Split("January February March April May June July August September October November December", " ")
    EnglishMonth = names(Month(value) - 1) & " " & CStr(Year(value))
End Function

Private Function GetPanelSheet() As Worksheet
    On Error Resume Next
    Set GetPanelSheet = ThisWorkbook.Worksheets(PANEL)
    On Error GoTo 0
    If GetPanelSheet Is Nothing Then StopWith "Run SetupHicController first (Alt+F8)."
    If CStr(GetPanelSheet.Range("A1").Value2) <> PANEL_ID Then StopWith "Controller sheet does not have the expected layout."
End Function

Public Sub SetupHicController()
    Dim ws As Worksheet, button As shape, i As Long, names As Variant, labels As Variant
    On Error GoTo Failed
    If mRunning Then StopWith "Wait for the current update to finish."
    If ThisWorkbook.FileFormat <> xlOpenXMLWorkbookMacroEnabled Then StopWith "Save this controller as HIC UPDATE.xlsm first."
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PANEL)
    On Error GoTo Failed
    If Not ws Is Nothing Then
        If CStr(ws.Range("A1").Value2) <> PANEL_ID Then StopWith "An unrelated sheet named HIC Controller already exists. Rename it before setup."
        ws.Activate
        RefreshDraftLabels ws
        MsgBox "Controller buttons updated. Its settings were preserved.", vbInformation
        Exit Sub
    End If
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    ws.Name = PANEL
    ws.Range("A1").value = PANEL_ID
    ws.Range("A1:B1").Merge
    ws.Range("A1").Font.Size = 20
    ws.Range("A1").Font.Bold = True
    ws.Range("A1:B1").Interior.Color = RGB(23, 83, 103)
    ws.Range("A1:B1").Font.Color = vbWhite
    ws.Rows(1).RowHeight = 38
    ws.Columns("A").ColumnWidth = 28
    ws.Columns("B").ColumnWidth = 95
    ws.Range("A3").value = "Report network root"
    ws.Range("B3").value = DEFAULT_ROOT
    ws.Range("A4").value = "Use year subfolder?"
    ws.Range("B4").value = True
    ws.Range("B3:B4").Interior.Color = RGB(255, 244, 204)
    ws.Range("B3").WrapText = True
    ws.Rows(3).RowHeight = 44
    ws.Range("A6").value = "Run date (India time)"
    ws.Range("B6").value = IndiaDate()
    ws.Range("B6").NumberFormat = "mm/dd/yyyy"
    ws.Range("A7").value = "Status"
    ws.Range("B7").value = "Ready - start with Validate Setup."
    ws.Range("A8").value = "Selected workbook"
    ws.Range("A9").value = "Controller log"
    ws.Range("B7:B9").WrapText = True
    ws.Rows("7:9").RowHeight = 42
    names = Array("RunHicValidate", "RunHicTest", "RunHicReview")
    labels = Array("1. Validate Setup", "2. Update Without Email", "3. Update and Save Draft")
    For i = 0 To 2
        Set button = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("B11").Left, ws.Range("B11").Top + i * 55, 330, 44)
        button.Name = "HIC_Controller_" & CStr(i + 1)
        button.TextFrame.Characters.Text = labels(i)
        button.TextFrame.Characters.Font.Color = vbWhite
        button.TextFrame.Characters.Font.Size = 12
        button.Fill.ForeColor.RGB = RGB(23, 83, 103)
        button.Line.Visible = msoFalse
        button.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!HicController." & names(i)
    Next i
    ws.Range("A22:B22").Merge
    ws.Range("A22").value = "Updates run in hidden Excel and save the report. The email is saved to Outlook Drafts and is never sent automatically."
    ws.Range("A22").WrapText = True
    ws.Rows(22).RowHeight = 32
    ws.Range("A24:B24").Merge
    ws.Range("A24").value = "Yellow cells select the report folder. Email settings are at the top of HicController. Save after setup."
    ws.Range("A24").WrapText = True
    ws.Rows(24).RowHeight = 32
    ws.Activate
    MsgBox "Buttons created. Check B3/B4, then save the controller. Start with Validate Setup.", vbInformation
    Exit Sub
Failed:
    MsgBox "Controller setup failed: " & Err.description, vbExclamation
End Sub

Public Sub RunHicValidate()
    RunController "Validate"
End Sub

Public Sub RunHicTest()
    RunController "Test"
End Sub

Public Sub RunHicReview()
    ' Retain the existing macro name so previously assigned buttons still work.
    RunHicDraft
End Sub

' All routines below operate on the opened report object, never ThisWorkbook.
' The only report macro invoked is the original Get Responses button action.
Private Sub CheckTimeLimit()
    If Now >= mDeadline Then StopWith "The HIC update exceeded its configured timeout."
End Sub

Private Sub WaitForReport(ByVal xl As Object, ByVal book As Object)
    Dim sheet As Object, qt As Object, connection As Object, busy As Boolean
    ' RefreshAll may start background OLEDB/ODBC/QueryTable refreshes.
    xl.CalculateUntilAsyncQueriesDone
    Do
        CheckTimeLimit
        busy = (xl.CalculationState = xlCalculating)
        For Each connection In book.Connections
            Select Case connection.Type
                Case 1 ' OLEDB
                    If connection.OLEDBConnection.Refreshing Then busy = True
                Case 2 ' ODBC
                    If connection.ODBCConnection.Refreshing Then busy = True
            End Select
        Next connection
        For Each sheet In book.Worksheets
            For Each qt In sheet.QueryTables
                If qt.Refreshing Then busy = True
            Next qt
        Next sheet
        If Not busy Then Exit Do
        DoEvents
    Loop
    CheckTimeLimit
End Sub

Private Sub UpdateOpenedReport(ByVal xl As Object, ByVal book As Object, ByVal runDate As Date)
    Dim ws As Object, pivotSheet As Object, table As Object
    Dim startDate As Date, endDate As Date, action As String
    Dim oldCalc As Long, oldEvents As Boolean, oldAlerts As Boolean, oldScreen As Boolean
    Dim e As Long, message As String, recordCount As Long
    oldCalc = xl.Calculation
    oldEvents = xl.EnableEvents
    oldAlerts = xl.DisplayAlerts
    oldScreen = xl.ScreenUpdating
    On Error GoTo Failed
    If book.ReadOnly Or book.FileFormat <> 52 Then StopWith "Report is read-only or is not .xlsm."
    Set ws = book.Worksheets("Query")
    Set pivotSheet = book.Worksheets("Pivot")
    If ws.ProtectContents Then StopWith "Query sheet is protected."
    If ws.Range("I4").MergeCells Or ws.Range("J4").MergeCells Then StopWith "Date cells must be unmerged."
    If ws.Range("I4").HasFormula Or ws.Range("J4").HasFormula Then StopWith "Date cells unexpectedly contain formulas."
    action = CStr(ws.Shapes("Button 1").OnAction)
    If InStrRev(action, "!") > 0 Then action = Mid$(action, InStrRev(action, "!") + 1)
    If StrComp(action, "ThisWorkbook.MainSub", vbTextCompare) <> 0 Then StopWith "Get Responses button assignment changed. Review it before running."
    startDate = DateSerial(Year(runDate), Month(runDate) - 1, 15)
    endDate = DateSerial(Year(runDate), Month(runDate) + 1, 0)
    If startDate >= endDate Then StopWith "Invalid reporting period."
    xl.EnableEvents = False
    xl.DisplayAlerts = False
    xl.ScreenUpdating = False
    xl.Calculation = xlCalculationAutomatic
    book.Activate
    ws.Activate
    ' Only write the two date inputs. Leave the existing macro to manage data.
    ws.Range("I4").value = startDate
    ws.Range("J4").value = endDate
    ws.Range("I4:J4").NumberFormat = "mm/dd/yyyy"
    WriteLog "STEP 1: dates updated: " & Format$(startDate, "mm/dd/yyyy") & " to " & Format$(endDate, "mm/dd/yyyy")
    WriteLog "STEP 2: invoking existing Get Responses action: " & action
    xl.Run "'" & Replace(book.Name, "'", "''") & "'!" & action
    WaitForReport xl, book
    xl.Calculate
    WaitForReport xl, book
    ' This is the workbook-object equivalent of Data > Refresh All.
    ' Selecting individual cells is not needed for Refresh All.
    pivotSheet.Activate
    WriteLog "STEP 3: Pivot sheet active; running Workbook.RefreshAll."
    book.RefreshAll
    WaitForReport xl, book
    xl.Calculate
    WaitForReport xl, book
    ' Basic nonempty-results check only. Formula errors are left as-is,
    ' matching the requested manual workflow. No source/table/formula edits.
    Set table = ws.ListObjects("Query")
    If table.DataBodyRange Is Nothing Then StopWith "Query results are empty. No save or draft was created."
    recordCount = xl.WorksheetFunction.CountA(table.ListColumns("Sample ID").DataBodyRange)
    If recordCount < 1 Then StopWith "Query contains no Sample IDs. No save or draft was created."
    WriteLog "REFRESH RETURNED: " & CStr(recordCount) & " nonempty Sample IDs; existing formula errors are not a blocking check."
    CheckTimeLimit
    book.Save
    If Not book.Saved Then StopWith "Excel did not confirm the report save."
    WriteLog "SAVED: " & book.FullName
    GoTo Cleanup
Failed:
    e = Err.Number
    message = Err.description
Cleanup:
    On Error Resume Next
    xl.Calculation = oldCalc
    xl.ScreenUpdating = oldScreen
    xl.DisplayAlerts = oldAlerts
    xl.EnableEvents = oldEvents
    Set table = Nothing
    Set pivotSheet = Nothing
    Set ws = Nothing
    On Error GoTo 0
    If e <> 0 Then Err.Raise e, "HIC report update", message
End Sub

Private Sub SaveOutlookDraft(ByVal reportPath As String, ByVal runDate As Date, ByVal stateFolder As String)
    Dim fs As Object, outlook As Object, mail As Object, attachment As Object, marker As Object
    Dim intent As String, e As Long, message As String, stage As String
    Dim signatureAttachments As Long
    On Error GoTo Failed
    stage = "Checking saved attachment and time limit"
    WriteLog "EMAIL STEP: " & stage
    CheckTimeLimit
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Not fs.FileExists(reportPath) Then StopWith "Saved report attachment is missing."
    If fs.GetFile(reportPath).Size = 0 Then StopWith "Saved report attachment is empty."
    If Len(Trim$(EMAIL_TO)) = 0 Then StopWith "The email recipient is empty."
    ' Each run has its own audit record. Earlier sends never block a later run.
    Do
        intent = fs.BuildPath(stateFolder, "HIC-" & Format$(runDate, "yyyy-mm-dd") & "-" & fs.GetTempName & ".intent")
    Loop While fs.FileExists(intent)
    stage = "Connecting to classic Outlook"
    WriteLog "EMAIL STEP: " & stage
    Set outlook = CreateObject("Outlook.Application")
    stage = "Creating email item"
    WriteLog "EMAIL STEP: " & stage
    Set mail = outlook.CreateItem(0)
    stage = "Composing message with GCC signature"
    WriteLog "EMAIL STEP: " & stage
    mail.BodyFormat = 2
    mail.HTMLBody = HicEmailHtml(fs.GetParentFolderName(reportPath))
    stage = "Setting To recipients"
    WriteLog "EMAIL STEP: " & stage
    mail.To = EMAIL_TO
    stage = "Setting CC recipients"
    WriteLog "EMAIL STEP: " & stage
    mail.CC = EMAIL_CC
    stage = "Setting sender"
    WriteLog "EMAIL STEP: " & stage
    If EMAIL_FROM <> "" Then mail.SentOnBehalfOfName = EMAIL_FROM
    stage = "Setting subject"
    WriteLog "EMAIL STEP: " & stage
    mail.Subject = Replace(EMAIL_SUBJECT, "{MONTH}", EnglishMonth(runDate))
    stage = "Setting message body"
    WriteLog "EMAIL STEP: " & stage

    ' Body and signature were prepared together before setting the shared sender.
    WriteLog "DRAFT: configured recipients retained for review before manual sending."
    stage = "Adding report attachment"
    WriteLog "EMAIL STEP: " & stage
    signatureAttachments = mail.Attachments.count
    Set attachment = mail.Attachments.Add(reportPath, 1)
    stage = "Verifying attachment"
    WriteLog "EMAIL STEP: " & stage
    If mail.Attachments.count <> signatureAttachments + 1 Or attachment.Size < 1 Then StopWith "Attachment verification failed."
    stage = "Checking saved attachment and time limit"
    WriteLog "EMAIL STEP: " & stage
    CheckTimeLimit
    ' Record this individual attempt before Save, including ambiguous failures.
    stage = "Recording email attempt"
    WriteLog "EMAIL STEP: " & stage
    Set marker = fs.CreateTextFile(intent, False, False)
    marker.WriteLine "PREPARED " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " Mode=Draft"
    marker.WriteLine "Workbook=" & reportPath
    marker.Close
    Set marker = Nothing
    stage = "Saving email to Outlook Drafts"
    WriteLog "EMAIL STEP: " & stage
    mail.Save
    If Len(CStr(mail.EntryID)) = 0 Then StopWith "Outlook did not confirm a saved draft ID."
    WriteLog "DRAFTED: Outlook saved the email without sending it."
    stage = "Recording draft status"
    WriteLog "EMAIL STEP: " & stage
    Set marker = fs.OpenTextFile(intent, 8, False)
    marker.WriteLine "DRAFTED " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    marker.Close
    Set marker = Nothing
    WriteLog "EMAIL COMPLETE: draft saved in Outlook; message was not sent."
    GoTo Cleanup
Failed:
    e = Err.Number
    message = "Report already saved. Email step failed: " & stage & " (error " & CStr(e) & "). " & Err.description
Cleanup:
    On Error Resume Next
    If Not marker Is Nothing Then marker.Close
    Set marker = Nothing
    Set attachment = Nothing
    Set mail = Nothing
    Set outlook = Nothing
    Set fs = Nothing
    On Error GoTo 0
    If e <> 0 Then Err.Raise e, "HIC email", message
End Sub

Private Function EscapeHicHtml(ByVal value As String) As String
    value = Replace(value, "&", "&amp;")
    value = Replace(value, "<", "&lt;")
    value = Replace(value, ">", "&gt;")
    EscapeHicHtml = Replace(value, Chr$(34), "&quot;")
End Function

Private Function HicEmailHtml(ByVal folderPath As String) As String
    Dim link As String, html As String
    link = Replace(folderPath, "%", "%25")
    link = Replace(Replace(link, "#", "%23"), " ", "%20")
    link = "file:" & Replace(link, "\", "/")
    html = "<html><body style=""font-family:Arial,sans-serif;font-size:11pt;color:#222;line-height:1.35;"">"
    html = html & "<p style=""margin:0 0 14pt 0;"">Hi All,</p>"
    html = html & "<p style=""margin:0 0 14pt 0;"">Please find the HIC Update File.</p>"
    html = html & "<p style=""margin:0 0 24pt 0;"">LINK: <a href=""" & EscapeHicHtml(link) & """>" & EscapeHicHtml(folderPath) & "</a></p>"
    html = html & "<div style=""margin:0;color:#173f38;"">Best regards,<br><br><b><ORGANIZATION_NAME></b><br><br>"
    html = html & "Email: <a href=""mailto:<SENDER_ADDRESS>""><SENDER_ADDRESS></a><br><br>"
    html = html & "<ORGANIZATION_BRAND_AND_ADDRESS>"
    html = html & "Find us here: <a href=""<ORGANIZATION_WEBSITE>""><ORGANIZATION_WEBSITE_LABEL></a>&nbsp;&nbsp; "
    html = html & "<a href=""<ORGANIZATION_LINKEDIN>"">LinkedIn</a>&nbsp;&nbsp; "
    html = html & "<a href=""<ORGANIZATION_FACEBOOK>"">Facebook</a>&nbsp;&nbsp; "
    html = html & "<a href=""<ORGANIZATION_INSTAGRAM>"">Instagram</a>&nbsp;&nbsp; "
    html = html & "<a href=""<ORGANIZATION_YOUTUBE>"">YouTube</a></div>"
    html = html & "<div style=""margin-top:24pt;max-width:760px;font-size:8pt;line-height:1.25;color:#888;font-style:italic;"">"
    html = html & "<p><ORGANIZATION_LEGAL_FOOTER></p>"
    html = html & "<p>This e-mail (including any attachments) is for the intended addressee(s) only and may contain confidential and/or proprietary information protected by law. You are hereby notified that any unauthorized reading, disclosure, copying or distribution of this e-mail or use of information herein is strictly prohibited. If you are not an intended recipient, you should delete this e-mail immediately. Thank you. Your personal data will be processed in accordance with our <a href=""<ORGANIZATION_WEBSITE>/privacy-policy"">privacy policy</a>.</p>"
    HicEmailHtml = html & "</div></body></html>"
End Function

Public Sub RunHicDraft()
    RunController "Draft"
End Sub

' One-click full workflow: validation, refresh, save and Outlook draft.
Public Sub RunHicAll()
    RunHicDraft
End Sub

Private Sub RefreshDraftLabels(ByVal ws As Worksheet)
    Dim shape As shape
    For Each shape In ws.Shapes
        If shape.Name = "HIC_Controller_3" Then
            shape.TextFrame.Characters.Text = "3. Update and Save Draft"
            shape.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!HicController.RunHicDraft"
        End If
    Next shape
    ws.Range("A22").value = "Updates run in hidden Excel and save the report. The email is saved to Outlook Drafts and is never sent automatically."
    ws.Range("A24").value = "Yellow cells select the report folder. Email settings are at the top of HicController. Save after setup."
End Sub

Private Sub WriteLog(ByVal message As String)
    Dim f As Integer
    f = FreeFile
    Open mLogFile For Append Access Write Lock Write As #f
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & Environ$("COMPUTERNAME") & vbTab & Environ$("USERNAME") & vbTab & Replace(Replace(message, vbCr, " "), vbLf, " ")
    Close #f
End Sub

Private Function FindReport(ByVal root As String, ByVal yearFolder As Boolean, ByVal d As Date) As String
    Dim fs As Object, folder As String, file As Object, label As String, shortLabel As String, count As Long
    Set fs = CreateObject("Scripting.FileSystemObject")
    folder = root
    If yearFolder Then folder = fs.BuildPath(root, CStr(Year(d)))
    label = EnglishMonth(d)
    shortLabel = Left$(label, 3) & " " & CStr(Year(d))
    If Not fs.FolderExists(folder) Then Exit Function
    For Each file In fs.GetFolder(folder).Files
        If StrComp(file.Name, "Follow Up Resultados HIC_Model-" & label & ".xlsm", vbTextCompare) = 0 _
            Or StrComp(file.Name, "Follow Up Resultados HIC_Model-" & Replace(label, " ", "-") & ".xlsm", vbTextCompare) = 0 _
            Or StrComp(file.Name, "Follow Up Resultados HIC_Model-" & shortLabel & ".xlsm", vbTextCompare) = 0 _
            Or StrComp(file.Name, "Follow Up Resultados HIC_Model-" & Replace(shortLabel, " ", "-") & ".xlsm", vbTextCompare) = 0 Then
            count = count + 1
            FindReport = file.Path
        End If
    Next file
    If count > 1 Then StopWith "Multiple " & label & " reports in " & folder & ". Keep one monthly report before running."
End Function

Private Function EnsureMonthlyReport(ByVal root As String, ByVal yearFolder As Boolean, ByVal d As Date, ByVal createMissing As Boolean) As String
    Dim fs As Object, current As String, source As String, folder As String, destination As String
    Dim temporary As String, sourceHandle As Integer, sourceHeld As Boolean
    Dim e As Long, message As String
    On Error GoTo Failed
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Not fs.FolderExists(root) Then StopWith "Report root is unavailable: " & root
    current = FindReport(root, yearFolder, d)
    If current <> "" Then
        EnsureMonthlyReport = current
        Exit Function
    End If
    source = FindReport(root, yearFolder, DateSerial(Year(d), Month(d) - 1, 1))
    If source = "" Then StopWith "Current report is missing and no previous-month report was found for " & EnglishMonth(DateSerial(Year(d), Month(d) - 1, 1)) & "."
    folder = root
    If yearFolder Then folder = fs.BuildPath(root, CStr(Year(d)))
    destination = fs.BuildPath(folder, "Follow Up Resultados HIC_Model-" & EnglishMonth(d) & ".xlsm")
    If Not createMissing Then
        WriteLog "MONTH PLAN: copy " & source & " to " & destination
        EnsureMonthlyReport = source ' Validation probes the template without creating anything.
        Exit Function
    End If
    ' Hold a write-denying lock while copying so the previous report cannot change.
    sourceHandle = FreeFile
    Open source For Input Access Read Lock Write As #sourceHandle
    sourceHeld = True
    If fs.GetFile(source).Size = 0 Then StopWith "Previous-month report is empty."
    If Not fs.FolderExists(folder) Then fs.CreateFolder folder
    temporary = fs.BuildPath(folder, "HIC-new-" & fs.GetTempName)
    fs.CopyFile source, temporary, False
    If fs.GetFile(temporary).Size <> fs.GetFile(source).Size Then StopWith "Monthly copy size verification failed."
    Close #sourceHandle
    sourceHeld = False
    ' Publish only the completed copy; MoveFile refuses an existing destination.
    If FindReport(root, yearFolder, d) <> "" Then StopWith "A current-month report appeared during copying. No report was overwritten."
    fs.MoveFile temporary, destination
    WriteLog "MONTH CREATED: " & destination & "; copied from " & source
    EnsureMonthlyReport = destination
    Exit Function
Failed:
    e = Err.Number
    message = Err.description
    On Error Resume Next
    If sourceHeld Then Close #sourceHandle
    On Error GoTo 0
    If temporary <> "" Then message = message & " Temporary copy, if present: " & temporary
    Err.Raise e, "HIC monthly file", message
End Function

Private Sub RunController(ByVal mode As String)
    Dim panelSheet As Worksheet, fs As Object, xl As Object, book As Object
    Dim root As String, stateFolder As String, reportPath As String
    Dim d As Date, yearFolder As Boolean, lockFile As Integer, probe As Integer
    Dim lockHeld As Boolean, probeHeld As Boolean, owned As Boolean, result As String
    Dim errorText As String, errorNumber As Long, oldStatus As Variant, pid As Long
    Dim processes As Object, process As Object, existing As Object, logFolder As String
    If mRunning Then
        MsgBox "An HIC run is already in progress.", vbInformation
        Exit Sub
    End If
    mRunning = True
    oldStatus = Application.StatusBar
    mLogFile = ""
    On Error GoTo Failed
    Set panelSheet = GetPanelSheet()
    RefreshDraftLabels panelSheet
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Len(Environ$("LOCALAPPDATA")) = 0 Then StopWith "Windows LOCALAPPDATA folder is unavailable."
    logFolder = fs.BuildPath(Environ$("LOCALAPPDATA"), "HIC-Controller-Logs")
    If Not fs.FolderExists(logFolder) Then fs.CreateFolder logFolder
    mLogFile = fs.BuildPath(logFolder, "HIC-" & Format$(Now, "yyyymmdd-hhnnss") & "-" & Environ$("COMPUTERNAME") & ".log")
    panelSheet.Range("B9").value = mLogFile
    WriteLog "START Mode=" & mode
    root = Trim$(CStr(panelSheet.Range("B3").Value2))
    If Right$(root, 1) = "\" Then root = Left$(root, Len(root) - 1)
    If Left$(root, 2) <> "\\" Then StopWith "B3 must contain a UNC network folder starting with \\."
    If UCase$(CStr(panelSheet.Range("B4").Value2)) <> "TRUE" And UCase$(CStr(panelSheet.Range("B4").Value2)) <> "FALSE" Then StopWith "B4 must be TRUE or FALSE."
    yearFolder = CBool(panelSheet.Range("B4").Value2)
    d = IndiaDate()
    panelSheet.Range("B6").value = d
    panelSheet.Range("B7").value = "Checking current-month report..."
    Application.StatusBar = "HIC: checking report..."
    stateFolder = fs.BuildPath(root, ".hic-state")
    If mode <> "Validate" Then
        If Not fs.FolderExists(stateFolder) Then fs.CreateFolder stateFolder
        lockFile = FreeFile
        Open fs.BuildPath(stateFolder, "launcher.lock") For Binary Access Read Write Lock Read Write As #lockFile
        lockHeld = True
    End If
    reportPath = EnsureMonthlyReport(root, yearFolder, d, mode <> "Validate")
    panelSheet.Range("B8").value = reportPath
    WriteLog "Workbook=" & reportPath & "; RunDate=" & Format$(d, "yyyy-mm-dd") & "; Start=" & Format$(DateSerial(Year(d), Month(d) - 1, 15), "mm/dd/yyyy") & "; End=" & Format$(DateSerial(Year(d), Month(d) + 1, 0), "mm/dd/yyyy")
    ' Probe the selected report (or previous-month template during validation).
    probe = FreeFile
    Open reportPath For Input Access Read Lock Read Write As #probe
    probeHeld = True
    Close #probe
    probeHeld = False
    stateFolder = fs.BuildPath(root, ".hic-state")
    If mode = "Validate" Then
        result = "VALIDATED: report or previous-month template is accessible. See log for any monthly creation plan. No report was created or updated; Outlook was not run."
        GoTo Finished
    End If
    ' Identify existing Excel processes before CreateObject. Never close one of them.
    Set existing = CreateObject("Scripting.Dictionary")
    Set processes = GetObject("winmgmts:\\.\root\cimv2").ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name='EXCEL.EXE'")
    For Each process In processes
        existing(CStr(process.processId)) = True
    Next process
    Set process = Nothing
    Set processes = Nothing
    Set xl = CreateObject("Excel.Application")
    GetWindowThreadProcessId xl.hWnd, pid
    If pid = 0 Or existing.Exists(CStr(pid)) Then StopWith "Cannot verify a new Excel instance. No existing Excel instance will be closed."
    owned = True
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.Caption = "HIC REPORT WORKER - leave this window alone while updating"
    If xl.Workbooks.count <> 0 Then
        xl.Visible = True ' Expose unexpected startup workbooks for manual recovery.
        owned = False
        StopWith "The new Excel instance loaded other workbooks. Close that instance manually and check startup add-ins."
    End If
    xl.EnableEvents = False
    xl.AskToUpdateLinks = False
    xl.AutomationSecurity = 2 ' Respect existing macro policy; never force-enable.
    panelSheet.Range("B7").value = "Opening report in separate Excel..."
    Set book = xl.Workbooks.Open(reportPath, 0, False)
    If book.ReadOnly Or book.FileFormat <> 52 Then StopWith "Report opened read-only or is not .xlsm."
    panelSheet.Range("B7").value = "Updating report. Do not edit or close the report worker."
    Application.StatusBar = "HIC: updating report in separate Excel..."
    WriteLog "WorkerPID=" & CStr(pid) & "; timeout=" & CStr(REFRESH_TIMEOUT_SECONDS) & " seconds; no independent hard watchdog"
    mDeadline = DateAdd("s", REFRESH_TIMEOUT_SECONDS, Now)
    UpdateOpenedReport xl, book, d
    If mode = "Draft" Then
        SaveOutlookDraft book.FullName, d, stateFolder
        result = "DRAFTED"
    Else
        result = "TEST_OK"
    End If
    If mode = "Test" And result <> "TEST_OK" Then StopWith "Unexpected test result: " & result
    If mode = "Draft" And result <> "DRAFTED" Then StopWith "Unexpected draft result: " & result
Finished:
    WriteLog "END " & result
    GoTo Cleanup
Failed:
    errorNumber = Err.Number
    errorText = Err.description
    On Error Resume Next
    If mLogFile <> "" Then WriteLog "FAILED " & CStr(errorNumber) & " " & errorText
    On Error GoTo 0
Cleanup:
    On Error Resume Next
    If probeHeld Then Close #probe
    If owned Then
        If Not book Is Nothing Then book.Close False
        ' Leave any additional workbook opened by a user/add-in alone.
        If xl.Workbooks.count = 0 Then
            xl.Quit
        Else
            xl.Visible = True ' Do not leave unexpected additional workbooks hidden.
            If errorText = "" Then errorText = "Other workbooks are now open in the worker; it was left open."
        End If
    End If
    Set book = Nothing
    Set xl = Nothing
    Set existing = Nothing
    Set processes = Nothing
    Set process = Nothing
    Set fs = Nothing
    If lockHeld Then Close #lockFile
    Application.StatusBar = oldStatus
    mRunning = False
    If Not panelSheet Is Nothing Then
        If errorText <> "" Then
            panelSheet.Range("B7").value = "Stopped: " & errorText
        Else
            panelSheet.Range("B7").value = result
        End If
    End If
    On Error GoTo 0
    ' Show a message only when the complete requested process succeeds.
    ' Failures remain visible in controller cell B7 and the log path in B9.
    If errorText = "" Then MsgBox "Process completed successfully.", vbInformation
End Sub


