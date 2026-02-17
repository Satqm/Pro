# Module 1
' ==================== MAIN MODULE ====================
Option Explicit

' Global variable to track correction mode
Public isCorrectionMode As Boolean


' ==================== MAIN LOGIN FUNCTION ====================
Public Sub LoginUserFromForm(enterpriseID As String)
    Dim ws As Worksheet
    Dim userSheetFound As Boolean
    Dim excludeSheets As Variant
    Dim wsDV As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim validUserSheets As Object
    Dim targetSheet As Worksheet
    Dim isValidUser As Boolean

    On Error GoTo ErrorHandler

    excludeSheets = Array("DV", "Admin", "Data base", "Form", "Login")
    userSheetFound = False
    isValidUser = False
    Set validUserSheets = CreateObject("Scripting.Dictionary")

    ' STEP 1: Get all valid Enterprise IDs from DV sheet
    On Error Resume Next
    Set wsDV = ThisWorkbook.Sheets("DV")
    On Error GoTo ErrorHandler

    If wsDV Is Nothing Then
        MsgBox "DV sheet not found!", vbCritical
    End If

    If Not wsDV Is Nothing Then
        lastRow = wsDV.Cells(wsDV.Rows.count, "F").End(xlUp).Row
        For i = 2 To lastRow
            Dim entID As String
            entID = Trim(CStr(wsDV.Cells(i, "F").value))
            If entID <> "" Then
                validUserSheets(LCase(entID)) = True
                ' Check if the current user is valid
                If LCase(entID) = LCase(enterpriseID) Then
                    isValidUser = True
                End If
            End If
        Next i
    End If

    ' STEP 2: Find target sheet
    For Each ws In ThisWorkbook.Sheets
        If LCase(ws.Name) = LCase(enterpriseID) Then
            userSheetFound = True
            Set targetSheet = ws
            Exit For
        End If
    Next ws

    ' STEP 3: Set target sheet if not found or user not valid
    If Not userSheetFound Or Not isValidUser Then
        Set targetSheet = ThisWorkbook.Sheets("Admin")
        userSheetFound = False ' Reset this for proper messaging
    End If

    ' STEP 4: Make Excel visible and setup everything
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Make Excel visible FIRST
    Application.Visible = True
    Application.WindowState = xlMaximized

    ' Prepare target sheet
    targetSheet.Visible = xlSheetVisible
    If userSheetFound And isValidUser Then
        Call SetupUserSheet(targetSheet, enterpriseID)
    End If

    ' Hide Login sheet
    On Error Resume Next
    ThisWorkbook.Sheets("Login").Visible = xlSheetVeryHidden
    On Error GoTo ErrorHandler

    ' Manage other sheet visibility
    For Each ws In ThisWorkbook.Sheets
        Select Case ws.Name
            Case "DV", "Data base", "Form"
                ' Make these sheets very hidden
                ws.Visible = xlSheetVeryHidden
            Case targetSheet.Name, "Admin", "Login"
                ' Keep these sheets as they are (already handled elsewhere)
            Case Else
                ' Hide other sheets that are not valid user sheets
                If Not validUserSheets.Exists(LCase(ws.Name)) Then
                    ws.Visible = xlSheetHidden
                End If
        End Select
    Next ws

    ' STEP 5: Activate the target sheet
    targetSheet.Activate
    targetSheet.Select

    ' **ENHANCEMENT: Ensure the correct cell is selected with professional scroll**
    On Error Resume Next
    If userSheetFound And isValidUser Then
        targetSheet.Range("C6").Select
    Else
        ' For Admin sheet, start at A1
        targetSheet.Range("A1").Select
    End If

    ' **NEW: Scroll to top-left for professional appearance**
    If Not ActiveWindow Is Nothing Then
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
    End If
    On Error GoTo ErrorHandler

    ' Final settings
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculate

    ' Small delay to ensure everything is rendered
    DoEvents
    Application.Wait Now + TimeValue("00:00:01")

    ' STEP 6: Show appropriate message based on login result
    If userSheetFound And isValidUser Then
        ' Successful login
        MsgBox "Welcome, " & Application.userName & "!" & vbNewLine & _
               "Enterprise ID '" & enterpriseID & "' logged in successfully.", _
               vbInformation, "Login Successful"
    Else
        ' User not found - Enhanced messaging
        Dim userResponse As VbMsgBoxResult

        If Not isValidUser Then
            ' User ID not found in valid users list
            userResponse = MsgBox("User not found!" & vbNewLine & vbNewLine & _
                                "Enterprise ID '" & enterpriseID & "' is not authorized." & vbNewLine & _
                                "You are being redirected to the Admin page." & vbNewLine & vbNewLine & _
                                "Please contact your administrator for access.", _
                                vbExclamation + vbOKOnly, "Access Denied")
        Else
            ' User exists but no corresponding sheet found
            userResponse = MsgBox("User sheet not found!" & vbNewLine & vbNewLine & _
                                "No worksheet exists for Enterprise ID '" & enterpriseID & "'." & vbNewLine & _
                                "You are being redirected to the Admin page." & vbNewLine & vbNewLine & _
                                "Please contact your administrator.", _
                                vbExclamation + vbOKOnly, "Sheet Not Found")
        End If
    End If

    Exit Sub

ErrorHandler:
    ' Ensure Excel becomes visible even if there's an error
    Application.Visible = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.WindowState = xlMaximized

    On Error Resume Next
    ThisWorkbook.Activate
    AppActivate Application.Caption
    On Error GoTo 0

    MsgBox "Error during login: " & Err.Description & vbNewLine & vbNewLine & _
           "Please try again or contact your administrator.", _
           vbCritical, "Login Error"
End Sub



' ==================== HELPER FUNCTIONS ====================
Private Sub SetupUserSheet(ws As Worksheet, entID As String)
    On Error Resume Next

    ' Unprotect sheet
      ' Unprotect sheet
    ws.Unprotect

      ' Initialize next code in BA1 (max existing +1, or 1 if none)
    Dim nextCode As Long
    nextCode = GetNextAvailableCodeForUser(entID)
    ws.Range("BA1").value = nextCode
    ws.Range("BA1").Locked = False

    ' Set enterprise ID and lock it
    ws.Range("W6").value = entID
    ws.Range("W6").Locked = True

    ' Set protection with proper ranges
    ws.Cells.Locked = True
    ws.Range("F2,W7").Locked = True
    ws.Range("C6,E6,H6,H1,B10:J14").Locked = False
    ws.Range("B3000:T3080").Locked = False

    ' Protect sheet and set zoom
    ws.Protect UserInterfaceOnly:=True, AllowFiltering:=True
    ActiveWindow.Zoom = 80

    On Error GoTo 0
End Sub

Private Function HasValidationInModule(rng As Range) As Boolean
    On Error Resume Next
    HasValidationInModule = (rng.Validation.Type <> xlValidateInputOnly)
    On Error GoTo 0
End Function







' ==================== GET NEXT LOCAL CODE (CACHED VERSION) ====================
Private Function GetNextLocalCode(ws As Worksheet) As String
    ' Read the current next code from cell BA1
    Dim nextCode As Long
    On Error Resume Next
    nextCode = Val(ws.Range("BA1").value)
    If nextCode < 1 Then
        ' BA1 is invalid – compute the real next code (one-time recovery)
        Dim entID As String
        entID = ws.Range("W6").value
        nextCode = GetNextAvailableCodeForUser(entID)
        ' Update BA1 so future calls are fast
        ws.Range("BA1").value = nextCode
    End If
    On Error GoTo 0
    
    GetNextLocalCode = CStr(nextCode)
End Function
' ==================== SUBMIT TIMESHEET WITH EMAIL CONFIRMATION ====================
Public Sub SubmitTimesheet()
    On Error GoTo ErrorHandler

    Dim wsForm As Worksheet, wsDB As Worksheet
    Dim nextRow As Long
    Dim timesheetType As String, enterpriseID As String, uniqueCode As String
    Dim employeeName As String, employeeEmail As String, monthName As String, weekName As String, teamName As String
    Dim submittedBy As String
    Dim isCorrection As Boolean
    Dim i As Integer
    Dim auditCount As Integer, nonAuditCount As Integer
    Dim baseCode As Long, codeCounter As Long
    Dim submittedRecords As Collection

    Set wsForm = ActiveSheet
    Set wsDB = ThisWorkbook.Sheets("Data base")
    Set submittedRecords = New Collection

    ' Gather basic form values
    enterpriseID = Trim(wsForm.Range("W6").value)
    employeeName = Trim(wsForm.Range("F2").value)
    employeeEmail = Trim(wsForm.Range("W7").value)
    monthName = Trim(wsForm.Range("C6").value)
    weekName = Trim(wsForm.Range("E6").value)
    teamName = Trim(wsForm.Range("H6").value)
    submittedBy = Application.userName

    ' Determine if this is a correction
    isCorrection = isCorrectionMode

    ' Basic validations
    If Not ValidateBasicFields(wsForm, enterpriseID, monthName, weekName, teamName) Then GoTo CleanUp

    ' Validate email address
    If employeeEmail = "" Then
        MsgBox "Employee email address is required (Cell W7).", vbExclamation, "Missing Email"
        GoTo CleanUp
    End If

    ' Validate hours columns
    If Not ValidateHoursColumns(wsForm) Then GoTo CleanUp

    ' Count and validate engagement records
    Dim recordCount As Integer
    recordCount = CountValidEngagementRecords(wsForm, auditCount, nonAuditCount)
    If recordCount = 0 Then
        MsgBox "Please fill at least one engagement record (either Audit or Non-Audit).", vbExclamation, "No Records Found"
        GoTo CleanUp
    End If

    ' Validate individual rows for completeness
    If Not ValidateEngagementRows(wsForm) Then GoTo CleanUp

    ' ==================== VALIDATION: Check for Non-Audit Engagement without Remarks ====================
    For i = 10 To 14
        If IsEngagementRowValid(wsForm, i) Then
            ' Check if there is any value in column H (Non-Audit Engagement)
            If Trim(wsForm.Cells(i, 8).value) <> "" Then
                ' Check if corresponding Remark in column J is empty
                If Trim(wsForm.Cells(i, 10).value) = "" Then
                    MsgBox "Please fill your Remark in Row " & i & "." & vbCrLf & _
                           "Remark is required when Non-Audit Engagement is filled.", _
                           vbExclamation, "Missing Remark"
                    wsForm.Range("J" & i).Select
                    GoTo CleanUp
                End If
            End If
        End If
    Next i

    ' Check for duplicates ONLY if NOT in correction mode
    If Not isCorrection Then
        If HasDuplicateEntries(wsForm, wsDB, enterpriseID, monthName, weekName, teamName) Then GoTo CleanUp
    End If

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    wsForm.Unprotect

    ' ==================== CORRECTION LOGIC ====================
    If isCorrection Then
        ' Mark existing records for deletion
        Dim r As Long
        For r = 3000 To 3200
            If wsForm.Cells(r, 3).value = enterpriseID And _
               wsForm.Cells(r, 8).value = monthName And _
               wsForm.Cells(r, 9).value = weekName And _
               wsForm.Cells(r, 10).value = teamName Then
                If wsForm.Cells(r, 2).value <> "" Then
                    wsForm.Cells(r, 20).value = "Marked for Deletion"
                End If
            End If
        Next r

        baseCode = CLng(GetNextLocalCode(wsForm))
        codeCounter = baseCode
    Else
        baseCode = CLng(GetNextLocalCode(wsForm))
        codeCounter = baseCode
    End If

    ' Confirmation message
    Dim confirmMsg As String
    If isCorrection Then
        confirmMsg = "You are Submitting Correction For " & auditCount & " Audit and " & nonAuditCount & " Non-Audit Engagements" & vbCrLf & _
                    "Week: " & weekName & vbCrLf & _
                    "Do you want to proceed with the corrections?"
    Else
        confirmMsg = "You are Recording " & auditCount & " Audit and " & nonAuditCount & " Non-Audit Engagements" & vbCrLf & _
                    "Week: " & weekName & vbCrLf & _
                    "Do you want to proceed with the submission?"
    End If

    If MsgBox(confirmMsg, vbQuestion + vbYesNo, "Confirm Submission") = vbNo Then
        MsgBox "Submission cancelled.", vbInformation
        GoTo CleanUp
    End If

    ' Write each engagement record and store for email
    For i = 10 To 14
        If IsEngagementRowValid(wsForm, i) Then
            nextRow = GetNextStorageRow(wsForm)
            uniqueCode = CStr(codeCounter)

            ' Write the record
            With wsForm
                .Cells(nextRow, 2).value = uniqueCode                  ' Unique Code
                .Cells(nextRow, 3).value = enterpriseID                ' Enterprise ID
                .Cells(nextRow, 4).value = employeeName                ' Employee Name
                .Cells(nextRow, 5).value = employeeEmail               ' Employee Email
                .Cells(nextRow, 6).value = Format(Now, "yyyy-mm-dd hh:mm:ss") ' Date and time
                .Cells(nextRow, 7).value = IIf(isCorrection, "Timesheet Correction", "New Timesheet Entry")
                .Cells(nextRow, 8).value = monthName                   ' Month
                .Cells(nextRow, 9).value = weekName                    ' Week
                .Cells(nextRow, 10).value = teamName                   ' Team Name
                .Cells(nextRow, 11).value = .Cells(i, 2).value        ' Region
                .Cells(nextRow, 12).value = .Cells(i, 3).value        ' Audit Engagement ID
                .Cells(nextRow, 13).value = .Cells(i, 4).value        ' Engagement Activity
                .Cells(nextRow, 14).value = .Cells(i, 5).value        ' Engagement Hours
                .Cells(nextRow, 15).value = .Cells(i, 6).value        ' Remark 1
                .Cells(nextRow, 16).value = .Cells(i, 8).value        ' Non-Audit Engagement
                .Cells(nextRow, 17).value = .Cells(i, 9).value        ' Non-Audit Hours
                .Cells(nextRow, 18).value = .Cells(i, 10).value       ' Remark 2
                .Cells(nextRow, 19).value = submittedBy               ' User Submitted Record

                ' Set font color to white (hidden)
                .Range(.Cells(nextRow, 2), .Cells(nextRow, 19)).Font.Color = RGB(255, 255, 255)

                ' Store record for email
                Dim recordData As Object
                Set recordData = CreateObject("Scripting.Dictionary")
                recordData("UniqueCode") = uniqueCode
                recordData("EnterpriseID") = enterpriseID
                recordData("EmployeeName") = employeeName
                recordData("Email") = employeeEmail
                recordData("DateTime") = Format(Now, "yyyy-mm-dd hh:mm:ss")
                recordData("Type") = IIf(isCorrection, "Timesheet Correction", "New Timesheet Entry")
                recordData("Month") = monthName
                recordData("Week") = weekName
                recordData("Team") = teamName
                recordData("Region") = .Cells(i, 2).value
                recordData("AuditEngagement") = .Cells(i, 3).value
                recordData("EngagementActivity") = .Cells(i, 4).value
                recordData("EngagementHours") = .Cells(i, 5).value
                recordData("Remark1") = .Cells(i, 6).value
                recordData("NonAuditEngagement") = .Cells(i, 8).value
                recordData("NonAuditHours") = .Cells(i, 9).value
                recordData("Remark2") = .Cells(i, 10).value
                recordData("SubmittedBy") = submittedBy

                submittedRecords.Add recordData
            End With

            codeCounter = codeCounter + 1
        End If
    Next i
    ThisWorkbook.Save

    ' Send email confirmation
    Call SendTimesheetConfirmationEmail(employeeName, employeeEmail, submittedRecords, isCorrection, weekName, enterpriseID, auditCount, nonAuditCount)

    ' Success message
    Dim successMsg As String
    If Not isCorrection Then
        successMsg = "Your Timesheet has been Successfully Submitted with " & auditCount & " Audit and " & nonAuditCount & " Non-Audit Engagements" & vbCrLf & _
                    "Week starting: " & weekName & vbCrLf & _
                    "Enterprise ID: " & enterpriseID & vbCrLf & _
                    "A confirmation email has been sent to: " & employeeEmail & vbCrLf & _
                    "Records will be synced to database by Admin."
    Else
        successMsg = "Your Timesheet Correction has been Successfully Submitted!" & vbCrLf & _
                    auditCount & " Audit and " & nonAuditCount & " Non-Audit Engagements" & vbCrLf & _
                    "Week starting: " & weekName & vbCrLf & _
                    "Previous records marked for deletion." & vbCrLf & _
                    "A confirmation email has been sent to: " & employeeEmail
    End If

  MsgBox successMsg, vbInformation, "Submission Complete"

    ' Update next code in BA1
    wsForm.Range("BA1").value = codeCounter

    ' Reset correction mode and form
    isCorrectionMode = False
    ResetFormInputsOnly wsForm
    GoTo CleanUp

ErrorHandler:
    MsgBox "ERROR in SubmitTimesheet:" & vbCrLf & _
           "Error: " & Err.Description & vbCrLf & _
           "Number: " & Err.Number, vbCritical, "System Error"

CleanUp:
    On Error Resume Next
    wsForm.Cells.Locked = True
    wsForm.Range("F2,W7").Locked = True
    wsForm.Range("C6,E6,H6,H1,B10:J14").Locked = False
    wsForm.Range("B3000:T3080").Locked = False
    wsForm.Protect UserInterfaceOnly:=True, AllowFiltering:=True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Call UpdatePendingCounts
    wsForm.Activate
End Sub
' ==================== GET NEXT AVAILABLE CODE FOR USER (ONE-TIME SCAN) ====================
Private Function GetNextAvailableCodeForUser(entID As String) As Long
    Dim maxCode As Long
    Dim ws As Worksheet
    Dim r As Long
    Dim wsDB As Worksheet
    Dim lastRow As Long
    
    maxCode = 0
    
    ' --- 1. Check the current user sheet if it already exists ---
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(entID)
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        For r = 3000 To 3200
            If ws.Cells(r, 2).value <> "" And IsNumeric(ws.Cells(r, 2).value) Then
                If CLng(ws.Cells(r, 2).value) > maxCode Then maxCode = CLng(ws.Cells(r, 2).value)
            End If
        Next r
    End If
    
    ' --- 2. Check the database for this user ---
    On Error Resume Next
    Set wsDB = ThisWorkbook.Sheets("Data base")
    On Error GoTo 0
    
    If Not wsDB Is Nothing Then
        lastRow = wsDB.Cells(wsDB.Rows.count, 1).End(xlUp).Row
        For r = 2 To lastRow
            If wsDB.Cells(r, 2).value = entID Then
                If IsNumeric(wsDB.Cells(r, 1).value) Then
                    If CLng(wsDB.Cells(r, 1).value) > maxCode Then maxCode = CLng(wsDB.Cells(r, 1).value)
                End If
            End If
        Next r
    End If
    
    ' Return the next available code
    GetNextAvailableCodeForUser = maxCode + 1
End Function
' ==================== SEND TIMESHEET CONFIRMATION EMAIL ====================
Private Sub SendTimesheetConfirmationEmail(employeeName As String, employeeEmail As String, _
                                          submittedRecords As Collection, isCorrection As Boolean, _
                                          weekName As String, enterpriseID As String, _
                                          auditCount As Integer, nonAuditCount As Integer)
    On Error GoTo EmailError

    Dim OutApp As Object
    Dim OutMail As Object
    Dim emailBody As String
    Dim recordData As Object
    Dim submissionType As String

    ' Determine submission type
    submissionType = IIf(isCorrection, "corrected", "submitted")

    ' Create Outlook application
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    ' Build email body with HTML formatting
    emailBody = BuildEmailBody(employeeName, submittedRecords, submissionType, weekName, enterpriseID, auditCount, nonAuditCount)

    ' Configure email
    With OutMail
        .To = employeeEmail
        .CC = "IAData.Analytics@sunlife.com"
        .Subject = "Timesheet " & IIf(isCorrection, "Correction", "Submission") & " Confirmation - Week " & weekName
        .htmlBody = emailBody
        .Display ' Use .Send to send automatically, .Display to show before sending
        ' For automatic sending, change .Display to .Send
        .Send
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing

    Exit Sub

EmailError:
    MsgBox "Error sending confirmation email: " & Err.Description & vbCrLf & _
           "However, your timesheet has been successfully submitted.", vbExclamation, "Email Error"
    Resume Next
End Sub

' ==================== BUILD EMAIL HTML BODY ====================
Private Function BuildEmailBody(employeeName As String, submittedRecords As Collection, _
                                submissionType As String, weekName As String, _
                                enterpriseID As String, auditCount As Integer, _
                                nonAuditCount As Integer) As String

    Dim htmlBody As String
    Dim recordData As Object
    Dim tableRows As String
    Dim i As Integer

    ' Build table rows from submitted records
    tableRows = ""
    For i = 1 To submittedRecords.count
        Set recordData = submittedRecords(i)
        tableRows = tableRows & BuildTableRow(recordData)
    Next i

    ' Build complete HTML email body
    htmlBody = "<!DOCTYPE html>" & vbCrLf
    htmlBody = htmlBody & "<html>" & vbCrLf
    htmlBody = htmlBody & "<head>" & vbCrLf
    htmlBody = htmlBody & "<style>" & vbCrLf
    htmlBody = htmlBody & "body { font-family: Arial, sans-serif; font-size: 14px; }" & vbCrLf
    htmlBody = htmlBody & "table { border-collapse: collapse; width: 100%; margin-top: 20px; }" & vbCrLf
    htmlBody = htmlBody & "th { background-color: #FDB913; color: white; padding: 10px; text-align: left; border: 1px solid #ddd; font-weight: bold; }" & vbCrLf
    htmlBody = htmlBody & "td { padding: 8px; border: 1px solid #ddd; }" & vbCrLf
    htmlBody = htmlBody & "tr:nth-child(even) { background-color: #f9f9f9; }" & vbCrLf
    htmlBody = htmlBody & ".summary { margin: 20px 0; padding: 15px; background-color: #f0f0f0; border-left: 4px solid #FDB913; }" & vbCrLf
    htmlBody = htmlBody & "</style>" & vbCrLf
    htmlBody = htmlBody & "</head>" & vbCrLf
    htmlBody = htmlBody & "<body>" & vbCrLf

    ' Greeting
    htmlBody = htmlBody & "<p>Hi <strong>" & employeeName & "</strong>,</p>" & vbCrLf

    ' Main message
    htmlBody = htmlBody & "<p>Your timesheet has been successfully <strong>" & submissionType & "</strong> with the following details:</p>" & vbCrLf

    ' Summary box
    htmlBody = htmlBody & "<div class='summary'>" & vbCrLf
    htmlBody = htmlBody & "<strong>Summary:</strong><br>" & vbCrLf
    htmlBody = htmlBody & "Week: " & weekName & "<br>" & vbCrLf
    htmlBody = htmlBody & "Enterprise ID: " & enterpriseID & "<br>" & vbCrLf
    htmlBody = htmlBody & "Audit Engagements: " & auditCount & "<br>" & vbCrLf
    htmlBody = htmlBody & "Non-Audit Engagements: " & nonAuditCount & vbCrLf
    htmlBody = htmlBody & "</div>" & vbCrLf

    ' Data table
    htmlBody = htmlBody & "<table>" & vbCrLf
    htmlBody = htmlBody & "<thead>" & vbCrLf
    htmlBody = htmlBody & "<tr>" & vbCrLf
    htmlBody = htmlBody & "<th>Unique Code</th>" & vbCrLf
    htmlBody = htmlBody & "<th>Enterprise ID</th>" & vbCrLf
    htmlBody = htmlBody & "<th>Employee Name</th>" & vbCrLf
    htmlBody = htmlBody & "<th>Email</th>" & vbCrLf
    htmlBody = htmlBody & "<th>Date and Time</th>" & vbCrLf
    htmlBody = htmlBody & "<th>Type</th>" & vbCrLf
    htmlBody = htmlBody & "<th>Month</th>" & vbCrLf
    htmlBody = htmlBody & "<th>Week</th>" & vbCrLf
    htmlBody = htmlBody & "<th>Team Name</th>" & vbCrLf
    htmlBody = htmlBody & "<th>Region</th>" & vbCrLf
    htmlBody = htmlBody & "<th>Audit Engagement Name</th>" & vbCrLf
    htmlBody = htmlBody & "<th>Engagement Activity</th>" & vbCrLf
    htmlBody = htmlBody & "<th>Engagement Actual Hours</th>" & vbCrLf
    htmlBody = htmlBody & "<th>Remark 1</th>" & vbCrLf
    htmlBody = htmlBody & "<th>Other than Engagement?</th>" & vbCrLf
    htmlBody = htmlBody & "<th>Others Actual Hours</th>" & vbCrLf
    htmlBody = htmlBody & "<th>Remark 2</th>" & vbCrLf
    htmlBody = htmlBody & "<th>User Submitted Record</th>" & vbCrLf
    htmlBody = htmlBody & "</tr>" & vbCrLf
    htmlBody = htmlBody & "</thead>" & vbCrLf
    htmlBody = htmlBody & "<tbody>" & vbCrLf
    htmlBody = htmlBody & tableRows
    htmlBody = htmlBody & "</tbody>" & vbCrLf
    htmlBody = htmlBody & "</table>" & vbCrLf

    ' Footer
    htmlBody = htmlBody & "<p style='margin-top: 30px;'>Thanks!</p>" & vbCrLf
    htmlBody = htmlBody & "<p style='font-size: 12px; color: #666;'><em>This is an automated confirmation email. Please do not reply to this email.</em></p>" & vbCrLf

    htmlBody = htmlBody & "</body>" & vbCrLf
    htmlBody = htmlBody & "</html>"

    BuildEmailBody = htmlBody
End Function

' ==================== BUILD TABLE ROW ====================
Private Function BuildTableRow(recordData As Object) As String
    Dim rowHTML As String

    rowHTML = "<tr>" & vbCrLf
    rowHTML = rowHTML & "<td>" & NullToEmpty(recordData("UniqueCode")) & "</td>" & vbCrLf
    rowHTML = rowHTML & "<td>" & NullToEmpty(recordData("EnterpriseID")) & "</td>" & vbCrLf
    rowHTML = rowHTML & "<td>" & NullToEmpty(recordData("EmployeeName")) & "</td>" & vbCrLf
    rowHTML = rowHTML & "<td>" & NullToEmpty(recordData("Email")) & "</td>" & vbCrLf
    rowHTML = rowHTML & "<td>" & NullToEmpty(recordData("DateTime")) & "</td>" & vbCrLf
    rowHTML = rowHTML & "<td>" & NullToEmpty(recordData("Type")) & "</td>" & vbCrLf
    rowHTML = rowHTML & "<td>" & NullToEmpty(recordData("Month")) & "</td>" & vbCrLf
    rowHTML = rowHTML & "<td>" & NullToEmpty(recordData("Week")) & "</td>" & vbCrLf
    rowHTML = rowHTML & "<td>" & NullToEmpty(recordData("Team")) & "</td>" & vbCrLf
    rowHTML = rowHTML & "<td>" & NullToEmpty(recordData("Region")) & "</td>" & vbCrLf
    rowHTML = rowHTML & "<td>" & NullToEmpty(recordData("AuditEngagement")) & "</td>" & vbCrLf
    rowHTML = rowHTML & "<td>" & NullToEmpty(recordData("EngagementActivity")) & "</td>" & vbCrLf
    rowHTML = rowHTML & "<td>" & NullToEmpty(recordData("EngagementHours")) & "</td>" & vbCrLf
    rowHTML = rowHTML & "<td>" & NullToEmpty(recordData("Remark1")) & "</td>" & vbCrLf
    rowHTML = rowHTML & "<td>" & NullToEmpty(recordData("NonAuditEngagement")) & "</td>" & vbCrLf
    rowHTML = rowHTML & "<td>" & NullToEmpty(recordData("NonAuditHours")) & "</td>" & vbCrLf
    rowHTML = rowHTML & "<td>" & NullToEmpty(recordData("Remark2")) & "</td>" & vbCrLf
    rowHTML = rowHTML & "<td>" & NullToEmpty(recordData("SubmittedBy")) & "</td>" & vbCrLf
    rowHTML = rowHTML & "</tr>" & vbCrLf

    BuildTableRow = rowHTML
End Function

' ==================== HELPER FUNCTION ====================
Private Function NullToEmpty(value As Variant) As String
    If IsNull(value) Or IsEmpty(value) Then
        NullToEmpty = ""
    Else
        NullToEmpty = CStr(value)
    End If
End Function



' ==================== HELPER FUNCTIONS ====================
Private Function GetNextStorageRow(ws As Worksheet) As Long
    Dim r As Long
    For r = 3000 To 3080
        If ws.Cells(r, 2).value = "" And ws.Cells(r, 20).value <> "Marked for Deletion" Then
            GetNextStorageRow = r
            Exit Function
        End If
    Next r
    GetNextStorageRow = 3000
End Function

Private Function ValidateBasicFields(wsForm As Worksheet, enterpriseID As String, monthName As String, weekName As String, teamName As String) As Boolean
    ValidateBasicFields = True
    If enterpriseID = "" Then
        MsgBox "Enterprise ID is required.", vbExclamation: ValidateBasicFields = False: Exit Function
    End If
    If monthName = "" Then
        MsgBox "Please select Month.", vbExclamation: ValidateBasicFields = False: wsForm.Range("C6").Select: Exit Function
    End If
    If weekName = "" Then
        MsgBox "Please select Week", vbExclamation: ValidateBasicFields = False: wsForm.Range("E6").Select: Exit Function
    End If
    If teamName = "" Then
        MsgBox "Please select Your Team.", vbExclamation: ValidateBasicFields = False: wsForm.Range("H6").Select: Exit Function
    End If
End Function

Private Function ValidateHoursColumns(ws As Worksheet) As Boolean
    Dim i As Integer
    Dim cellValue As String
    Dim numValue As Double

    ValidateHoursColumns = True

    For i = 10 To 14
        ' Check Audit Hours (Column E)
        cellValue = Trim(CStr(ws.Cells(i, 5).value))
        If cellValue <> "" Then
            If Not IsNumeric(cellValue) Then
                MsgBox "Row " & i & ": Audit Hours (Column E) must be a valid number.", vbExclamation, "Invalid Hours Entry"
                ws.Cells(i, 5).Select
                ValidateHoursColumns = False
                Exit Function
            End If
        End If

        ' Check Non-Audit Hours (Column I)
        cellValue = Trim(CStr(ws.Cells(i, 9).value))
        If cellValue <> "" Then
            If Not IsNumeric(cellValue) Then
                MsgBox "Row " & i & ": Non-Audit Hours (Column I) must be a valid number.", vbExclamation, "Invalid Hours Entry"
                ws.Cells(i, 9).Select
                ValidateHoursColumns = False
                Exit Function
            End If
        End If
    Next i
End Function

Private Function CountValidEngagementRecords(ws As Worksheet, ByRef auditCount As Integer, ByRef nonAuditCount As Integer) As Integer
    Dim count As Integer, i As Integer
    count = 0
    auditCount = 0
    nonAuditCount = 0

    For i = 10 To 14
        If IsEngagementRowValid(ws, i) Then
            count = count + 1
            If Len(Trim(CStr(ws.Cells(i, 3).value))) > 0 Then
                auditCount = auditCount + 1
            End If
            If Len(Trim(CStr(ws.Cells(i, 8).value))) > 0 Then
                nonAuditCount = nonAuditCount + 1
            End If
        End If
    Next i

    CountValidEngagementRecords = count
End Function

Private Function IsEngagementRowValid(ws As Worksheet, rowNum As Integer) As Boolean
    Dim hasAuditComplete As Boolean, hasNonAuditComplete As Boolean
    Dim auditFieldsFilled As Boolean, nonAuditFieldsFilled As Boolean

    auditFieldsFilled = (Len(Trim(CStr(ws.Cells(rowNum, 2).value))) > 0 Or _
                        Len(Trim(CStr(ws.Cells(rowNum, 3).value))) > 0 Or _
                        Len(Trim(CStr(ws.Cells(rowNum, 4).value))) > 0 Or _
                        Len(Trim(CStr(ws.Cells(rowNum, 5).value))) > 0 Or _
                        Len(Trim(CStr(ws.Cells(rowNum, 6).value))) > 0)

    nonAuditFieldsFilled = (Len(Trim(CStr(ws.Cells(rowNum, 8).value))) > 0 Or _
                           Len(Trim(CStr(ws.Cells(rowNum, 9).value))) > 0 Or _
                           Len(Trim(CStr(ws.Cells(rowNum, 10).value))) > 0)

    If Not (auditFieldsFilled Or nonAuditFieldsFilled) Then
        IsEngagementRowValid = False
        Exit Function
    End If

    If auditFieldsFilled Then
        hasAuditComplete = (Len(Trim(CStr(ws.Cells(rowNum, 2).value))) > 0 And _
                           Len(Trim(CStr(ws.Cells(rowNum, 3).value))) > 0 And _
                           Len(Trim(CStr(ws.Cells(rowNum, 4).value))) > 0 And _
                           Len(Trim(CStr(ws.Cells(rowNum, 5).value))) > 0)
    End If

    If nonAuditFieldsFilled Then
        hasNonAuditComplete = (Len(Trim(CStr(ws.Cells(rowNum, 8).value))) > 0 And _
                              Len(Trim(CStr(ws.Cells(rowNum, 9).value))) > 0)
    End If

    IsEngagementRowValid = (hasAuditComplete Or hasNonAuditComplete)
End Function

Private Function ValidateEngagementRows(ws As Worksheet) As Boolean
    Dim i As Integer
    Dim hasAuditComplete As Boolean, hasNonAuditComplete As Boolean
    Dim auditFieldsFilled As Boolean, nonAuditFieldsFilled As Boolean

    ValidateEngagementRows = True

    For i = 10 To 14
        auditFieldsFilled = (Len(Trim(CStr(ws.Cells(i, 2).value))) > 0 Or _
                            Len(Trim(CStr(ws.Cells(i, 3).value))) > 0 Or _
                            Len(Trim(CStr(ws.Cells(i, 4).value))) > 0 Or _
                            Len(Trim(CStr(ws.Cells(i, 5).value))) > 0 Or _
                            Len(Trim(CStr(ws.Cells(i, 6).value))) > 0)

        nonAuditFieldsFilled = (Len(Trim(CStr(ws.Cells(i, 8).value))) > 0 Or _
                               Len(Trim(CStr(ws.Cells(i, 9).value))) > 0 Or _
                               Len(Trim(CStr(ws.Cells(i, 10).value))) > 0)

        If auditFieldsFilled Or nonAuditFieldsFilled Then
            If auditFieldsFilled Then
                hasAuditComplete = (Len(Trim(CStr(ws.Cells(i, 2).value))) > 0 And _
                                   Len(Trim(CStr(ws.Cells(i, 3).value))) > 0 And _
                                   Len(Trim(CStr(ws.Cells(i, 4).value))) > 0 And _
                                   Len(Trim(CStr(ws.Cells(i, 5).value))) > 0)
            Else
                hasAuditComplete = False
            End If

            If nonAuditFieldsFilled Then
                hasNonAuditComplete = (Len(Trim(CStr(ws.Cells(i, 8).value))) > 0 And _
                                      Len(Trim(CStr(ws.Cells(i, 9).value))) > 0)
            Else
                hasNonAuditComplete = False
            End If

            If auditFieldsFilled And Not hasAuditComplete Then
                MsgBox "Row " & i & ": Please complete all required Audit Engagement fields.", vbExclamation, "Incomplete Audit Entry"
                ws.Cells(i, 2).Select
                ValidateEngagementRows = False
                Exit Function
            End If

            If nonAuditFieldsFilled And Not hasNonAuditComplete Then
                MsgBox "Row " & i & ": Please complete all required Non-Audit Engagement fields.", vbExclamation, "Incomplete Non-Audit Entry"
                ws.Cells(i, 8).Select
                ValidateEngagementRows = False
                Exit Function
            End If

            If Not (hasAuditComplete Or hasNonAuditComplete) Then
                MsgBox "Row " & i & ": Please complete either Audit OR Non-Audit engagement fields completely.", vbExclamation, "Incomplete Entry"
                ValidateEngagementRows = False
                Exit Function
            End If
        End If
    Next i
End Function

' ==================== ENHANCED DUPLICATE CHECKING WITH ROW NAVIGATION ====================
Private Function HasDuplicateEntries(wsForm As Worksheet, wsDB As Worksheet, enterpriseID As String, _
                                   monthName As String, weekName As String, teamName As String) As Boolean
    Dim i As Integer
    HasDuplicateEntries = False

    For i = 10 To 14
        If IsEngagementRowValid(wsForm, i) Then
            Dim engagementID As String
            Dim region As String
            Dim engagementActivity As String
            Dim engagementType As String
            Dim remarks As String

            If Len(Trim(CStr(wsForm.Cells(i, 3).value))) > 0 Then
                ' Audit engagement
                engagementID = Trim(CStr(wsForm.Cells(i, 3).value))
                region = Trim(CStr(wsForm.Cells(i, 2).value))
                engagementActivity = Trim(CStr(wsForm.Cells(i, 4).value))
                engagementType = "Audit"
                remarks = "" ' Remarks not used for Audit duplicate checking
            Else
                ' Non-audit engagement
                engagementID = Trim(CStr(wsForm.Cells(i, 8).value))
                region = ""
                engagementActivity = ""
                engagementType = "Non-Audit"
                ' IMPORTANT: Adjust column number (9) to match your actual Remarks column for Non-Audit
                remarks = Trim(CStr(wsForm.Cells(i, 9).value))
            End If

            Dim duplicateLocation As String
            duplicateLocation = ""

            ' Check for duplicates and get location info
            If CheckDuplicateInLocal(wsForm, enterpriseID, engagementID, monthName, weekName, region, teamName, engagementActivity, engagementType, remarks, duplicateLocation) Or _
               CheckDuplicateInDB(wsDB, enterpriseID, engagementID, monthName, weekName, region, teamName, engagementActivity, engagementType, remarks, duplicateLocation) Then

                ' Navigate to the problematic cell
                NavigateToErrorCell wsForm, i, engagementType

                ' Enhanced error message with row information
                Dim rowLabel As String
                Select Case i
                    Case 10: rowLabel = "Row 1"
                    Case 11: rowLabel = "Row 2"
                    Case 12: rowLabel = "Row 3"
                    Case 13: rowLabel = "Row 4"
                    Case 14: rowLabel = "Row 5"
                End Select

                Dim response As VbMsgBoxResult
                response = MsgBox("DUPLICATE ENTRY DETECTED" & vbCrLf & vbCrLf & _
                                "Location: " & rowLabel & " (Excel Row " & i & ")" & vbCrLf & _
                                "Type: " & engagementType & " Engagement" & vbCrLf & _
                                "Engagement ID: " & engagementID & vbCrLf & _
                                IIf(region <> "", "Region: " & region & vbCrLf, "") & _
                                IIf(engagementActivity <> "", "Activity: " & engagementActivity & vbCrLf, "") & _
                                IIf(engagementType = "Non-Audit" And remarks <> "", "Remarks: " & remarks & vbCrLf, "") & _
                                "Week: " & weekName & vbCrLf & _
                                "Team: " & teamName & vbCrLf & _
                                "Found in: " & duplicateLocation & vbCrLf & vbCrLf & _
                                "This entry already exists for the selected week." & vbCrLf & vbCrLf & _
                                "Would you like to make a CORRECTION to the existing entry instead?", _
                                vbYesNo + vbExclamation, "Duplicate Entry Found - " & rowLabel)

                If response = vbYes Then
                    MsgBox "Please use the 'View/Correction' button to modify existing entries." & vbCrLf & vbCrLf & _
                           "Tip: The correction mode will load your existing data for editing.", _
                           vbInformation, "Use Correction Mode"
                    HasDuplicateEntries = True
                    Exit Function
                Else
                    MsgBox "Submission cancelled." & vbCrLf & vbCrLf & _
                           "Please review " & rowLabel & " and modify the duplicate entry.", _
                           vbInformation, "Submission Cancelled"
                    HasDuplicateEntries = True
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

' ==================== NAVIGATION HELPER FUNCTION ====================
Private Sub NavigateToErrorCell(wsForm As Worksheet, rowNum As Integer, engagementType As String)
    On Error Resume Next

    ' Navigate to the appropriate cell based on engagement type
    If engagementType = "Audit" Then
        ' Focus on Audit Engagement ID cell (Column C)
        wsForm.Cells(rowNum, 3).Select
        ' Highlight the entire audit section for this row
        wsForm.Range(wsForm.Cells(rowNum, 2), wsForm.Cells(rowNum, 6)).Interior.Color = RGB(255, 200, 200)
    Else
        ' Focus on Non-Audit Engagement cell (Column H)
        wsForm.Cells(rowNum, 8).Select
        ' Highlight the non-audit section for this row
        wsForm.Range(wsForm.Cells(rowNum, 8), wsForm.Cells(rowNum, 10)).Interior.Color = RGB(255, 200, 200)
    End If

    ' Scroll to make sure the cell is visible
    ActiveWindow.ScrollRow = rowNum - 5

    On Error GoTo 0
End Sub

' ==================== ENHANCED Local Storage Duplicate Check ====================
Private Function CheckDuplicateInLocal(ws As Worksheet, enterpriseID As String, engagementID As String, _
                                     monthName As String, weekName As String, region As String, teamName As String, _
                                     engagementActivity As String, engagementType As String, remarks As String, _
                                     ByRef duplicateLocation As String) As Boolean
    Dim r As Long

    For r = 3000 To 3200
        If ws.Cells(r, 3).value = enterpriseID And _
           ws.Cells(r, 8).value = monthName And _
           ws.Cells(r, 9).value = weekName And _
           ws.Cells(r, 10).value = teamName And _
           ((region <> "" And ws.Cells(r, 11).value = region) Or region = "") And _
           (ws.Cells(r, 12).value = engagementID Or ws.Cells(r, 16).value = engagementID) And _
           ((engagementActivity <> "" And ws.Cells(r, 13).value = engagementActivity) Or engagementActivity = "") And _
           ws.Cells(r, 20).value <> "Marked for Deletion" Then

           ' For Non-Audit engagements, check if Remarks are different
           If engagementType = "Non-Audit" Then
               ' IMPORTANT: Adjust column number (17) to match your actual Remarks column in local storage
               Dim storedRemarks As String
               storedRemarks = Trim(CStr(ws.Cells(r, 17).value))

               ' If Remarks are different, this is NOT a duplicate
               If storedRemarks <> remarks Then
                   ' Continue checking other rows
                   GoTo ContinueLoop
               End If
           End If

           duplicateLocation = "Local Storage (Pending Sync)"
           CheckDuplicateInLocal = True
           Exit Function
        End If
ContinueLoop:
    Next r
    CheckDuplicateInLocal = False
End Function

' ==================== ENHANCED Database Duplicate Check ====================
Private Function CheckDuplicateInDB(wsDB As Worksheet, enterpriseID As String, engagementID As String, _
                                  monthName As String, weekName As String, region As String, teamName As String, _
                                  engagementActivity As String, engagementType As String, remarks As String, _
                                  ByRef duplicateLocation As String) As Boolean
    Dim r As Long, lastRow As Long

    lastRow = wsDB.Cells(wsDB.Rows.count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If wsDB.Cells(r, 2).value = enterpriseID And _
           wsDB.Cells(r, 7).value = monthName And _
           wsDB.Cells(r, 8).value = weekName And _
           wsDB.Cells(r, 9).value = teamName And _
           ((region <> "" And wsDB.Cells(r, 10).value = region) Or region = "") And _
           (wsDB.Cells(r, 11).value = engagementID Or wsDB.Cells(r, 15).value = engagementID) And _
           ((engagementActivity <> "" And wsDB.Cells(r, 12).value = engagementActivity) Or engagementActivity = "") Then

           ' For Non-Audit engagements, check if Remarks are different
           If engagementType = "Non-Audit" Then
               ' IMPORTANT: Adjust column number (16) to match your actual Remarks column in database
               Dim storedRemarks As String
               storedRemarks = Trim(CStr(wsDB.Cells(r, 16).value))

               ' If Remarks are different, this is NOT a duplicate
               If storedRemarks <> remarks Then
                   ' Continue checking other rows
                   GoTo ContinueLoop2
               End If
           End If

           duplicateLocation = "Database (Already Synced)"
           CheckDuplicateInDB = True
           Exit Function
        End If
ContinueLoop2:
    Next r
    CheckDuplicateInDB = False
End Function


' ==================== CLEAR HIGHLIGHTING FUNCTION ====================
Public Sub ClearRowHighlighting()
    ' Call this function to clear any error highlighting
    On Error Resume Next
    Dim wsForm As Worksheet
    Set wsForm = ActiveSheet

    wsForm.Unprotect
    ' Clear highlighting from the engagement entry area
    wsForm.Range("B10:J14").Interior.ColorIndex = xlNone

    ' Restore protection
    wsForm.Cells.Locked = True
    wsForm.Range("F2,W7").Locked = True
    wsForm.Range("C6,E6,H6,H1,B10:J14").Locked = False
    wsForm.Range("B3000:T3080").Locked = False
    wsForm.Protect UserInterfaceOnly:=True, AllowFiltering:=True
    On Error GoTo 0
End Sub


' ==================== VIEW/CORRECTION FUNCTION ====================
Public Sub ViewCorrection()
    Dim wsForm As Worksheet
    Set wsForm = ActiveSheet
    Dim weekName As String, entID As String

    weekName = Trim(wsForm.Range("E6").value)
    entID = Trim(wsForm.Range("W6").value)

    If weekName = "" Then
        MsgBox "Please select a week first before viewing/correcting entries.", vbExclamation, "Week Required"
        wsForm.Range("E6").Select
        Exit Sub
    End If

    ' Set correction mode
    isCorrectionMode = True

    ' Clear the entry area first
    ClearEntryArea wsForm

    ' Load week data into the entry area
    LoadWeekDataToEntryArea wsForm, weekName, entID

    MsgBox "Week entries loaded in the timesheet area. You can now edit and submit corrections." & vbCrLf & _
           "Note: Duplicate checking is disabled in correction mode.", vbInformation, "View/Correction Mode"
End Sub

Private Sub ClearEntryArea(ws As Worksheet)
    On Error Resume Next
    ws.Unprotect
    ' Clear the entry area (rows 10-14)
    ws.Range("B10:J14").ClearContents
    ws.Range("B10:J14").Interior.ColorIndex = xlNone
    
    ' Clear the unique code storage (column Z)
    ws.Range("Z10:Z14").ClearContents

    ws.Cells.Locked = True
    ws.Range("F2,W7").Locked = True
    ws.Range("C6,E6,H6,H1,B10:J14").Locked = False
    ws.Range("B3000:T3080").Locked = False
    ws.Protect UserInterfaceOnly:=True, AllowFiltering:=True
    On Error GoTo 0
End Sub

Private Sub LoadWeekDataToEntryArea(ws As Worksheet, weekName As String, entID As String)
    Dim r As Long
    Dim targetRow As Long
    Dim foundCount As Long
    
    targetRow = 10
    
    ' Search in local storage (from row 3000 to 3200)
    For r = 3000 To 3200
        If ws.Cells(r, 3).value = entID And _
           ws.Cells(r, 9).value = weekName And _
           targetRow <= 14 Then
            
            ' Copy data to entry area
            ws.Cells(targetRow, 2).value = ws.Cells(r, 11).value  ' Region
            ws.Cells(targetRow, 3).value = ws.Cells(r, 12).value  ' Audit Engagement
            ws.Cells(targetRow, 4).value = ws.Cells(r, 13).value  ' Engagement Activity
            ws.Cells(targetRow, 5).value = ws.Cells(r, 14).value  ' Hours
            ws.Cells(targetRow, 6).value = ws.Cells(r, 15).value  ' Remark 1
            ws.Cells(targetRow, 8).value = ws.Cells(r, 16).value  ' Non-Audit Engagement
            ws.Cells(targetRow, 9).value = ws.Cells(r, 17).value  ' Non-Audit Hours
            ws.Cells(targetRow, 10).value = ws.Cells(r, 18).value ' Remark 2
            
            ' Store the unique code in column Z for reference
            ws.Cells(targetRow, 26).value = ws.Cells(r, 2).value  ' Unique Code in column Z
            
            targetRow = targetRow + 1
            foundCount = foundCount + 1
        End If
    Next r
    
    ' Also search in database
    Dim wsDB As Worksheet
    Set wsDB = ThisWorkbook.Sheets("Data base")
    Dim lastRow As Long
    lastRow = wsDB.Cells(wsDB.Rows.count, 1).End(xlUp).Row
    
    For r = 2 To lastRow
        If wsDB.Cells(r, 2).value = entID And _
           wsDB.Cells(r, 7).value = weekName And _
           targetRow <= 14 Then
            
            ' Copy data to entry area
            ws.Cells(targetRow, 2).value = wsDB.Cells(r, 9).value   ' Region
            ws.Cells(targetRow, 3).value = wsDB.Cells(r, 10).value  ' Audit Engagement
            ws.Cells(targetRow, 4).value = wsDB.Cells(r, 11).value  ' Engagement Activity
            ws.Cells(targetRow, 5).value = wsDB.Cells(r, 12).value  ' Hours
            ws.Cells(targetRow, 6).value = wsDB.Cells(r, 13).value  ' Remark 1
            ws.Cells(targetRow, 8).value = wsDB.Cells(r, 14).value  ' Non-Audit Engagement
            ws.Cells(targetRow, 9).value = wsDB.Cells(r, 15).value  ' Non-Audit Hours
            ws.Cells(targetRow, 10).value = wsDB.Cells(r, 16).value ' Remark 2
            
            ' Store the unique code in column Z for reference
            ws.Cells(targetRow, 26).value = wsDB.Cells(r, 1).value  ' Unique Code in column Z
            
            targetRow = targetRow + 1
            foundCount = foundCount + 1
        End If
    Next r
    
    If foundCount = 0 Then
        MsgBox "No records found for week: " & weekName, vbInformation, "No Records"
        isCorrectionMode = False
    End If
End Sub

' ==================== TOGGLE MY DATABASE RECORDS (OPTIMIZED) ====================
Public Sub ToggleMyDatabaseRecords()
    Dim wsF As Worksheet, wsDB As Worksheet
    Dim entID As String, i As Long, j As Long, lastRow As Long, outRow As Long
    Dim startCell As Range, tblArea As Range
    Dim isDataPresent As Boolean
    Dim recordsArray As Variant, recordCount As Long
    Dim tempArray() As Variant

    ' Performance optimization
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Set wsF = ActiveSheet
    Set wsDB = ThisWorkbook.Sheets("Data base")
    Set startCell = wsF.Range("B25")
    Set tblArea = wsF.Range("B25:P297")

    entID = Trim(wsF.Range("W6").value)
    If entID = "" Then
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        MsgBox "Please verify Enterprise ID is populated.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    wsF.Unprotect
    On Error GoTo 0

    isDataPresent = WorksheetFunction.CountA(tblArea.Offset(1)) > 0

    If isDataPresent Then
        ' Clear all content and formatting completely
        With tblArea
            .ClearContents
            .ClearFormats
            .Interior.ColorIndex = xlNone
            .Borders.LineStyle = xlNone
            .Font.Bold = False
            .WrapText = False
            .HorizontalAlignment = xlGeneral
        End With

        On Error Resume Next
        wsF.Cells.Locked = True
        wsF.Range("F2,W7").Locked = True
        wsF.Range("C6,E6,H6,H1,B10:J14").Locked = False
        wsF.Range("B3000:T3080").Locked = False
        wsF.Protect UserInterfaceOnly:=True, AllowFiltering:=True
        On Error GoTo 0

        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        MsgBox "Your past timesheet records have been hidden.", vbInformation, "Toggle Complete"
        Exit Sub
    End If

    ' Headers
    startCell.Resize(1, 15).value = Array("Unique Code", "Employee Name", "Date and Time", "Type", _
        "Week", "Region", "Team Name", "Engagement Activity", "Audit Engagement", _
        "Engagement Hours", "Remark 1", "Non-Audit Eng", "Non-Audit Hours", "Remark 2", "Submitted By")

    ' Minimal header formatting - light color, bold, center, wrap text, thin border
    With startCell.Resize(1, 15)
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 240)  ' Light blue-gray
        .HorizontalAlignment = xlCenter
        .WrapText = True
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    ' Initialize array
    ReDim tempArray(1 To 1000, 1 To 15)
    recordCount = 0

    ' Collect database records
    lastRow = wsDB.Cells(wsDB.Rows.count, 1).End(xlUp).Row
    If lastRow >= 2 Then
        For i = 2 To lastRow
            If wsDB.Cells(i, 2).value = entID Then
                recordCount = recordCount + 1
                tempArray(recordCount, 1) = wsDB.Cells(i, 1).value
                tempArray(recordCount, 2) = wsDB.Cells(i, 3).value
                tempArray(recordCount, 3) = wsDB.Cells(i, 5).value
                tempArray(recordCount, 4) = wsDB.Cells(i, 6).value
                tempArray(recordCount, 5) = wsDB.Cells(i, 8).value
                tempArray(recordCount, 7) = wsDB.Cells(i, 9).value
                tempArray(recordCount, 6) = wsDB.Cells(i, 10).value
                tempArray(recordCount, 9) = wsDB.Cells(i, 11).value
                tempArray(recordCount, 8) = wsDB.Cells(i, 12).value
                tempArray(recordCount, 10) = wsDB.Cells(i, 13).value
                tempArray(recordCount, 11) = wsDB.Cells(i, 14).value
                tempArray(recordCount, 12) = wsDB.Cells(i, 15).value
                tempArray(recordCount, 13) = wsDB.Cells(i, 16).value
                tempArray(recordCount, 14) = wsDB.Cells(i, 17).value
                tempArray(recordCount, 15) = wsDB.Cells(i, 18).value
            End If
        Next i
    End If

    ' Collect local pending records
    For i = 3000 To 3200
        If wsF.Cells(i, 3).value = entID And wsF.Cells(i, 20).value <> "Marked for Deletion" Then
            recordCount = recordCount + 1
            tempArray(recordCount, 1) = wsF.Cells(i, 2).value
            tempArray(recordCount, 2) = wsF.Cells(i, 4).value
            tempArray(recordCount, 3) = wsF.Cells(i, 6).value
            tempArray(recordCount, 4) = wsF.Cells(i, 7).value
            tempArray(recordCount, 5) = wsF.Cells(i, 9).value
            tempArray(recordCount, 7) = wsF.Cells(i, 10).value
            tempArray(recordCount, 6) = wsF.Cells(i, 11).value
            tempArray(recordCount, 9) = wsF.Cells(i, 12).value
            tempArray(recordCount, 8) = wsF.Cells(i, 13).value
            tempArray(recordCount, 10) = wsF.Cells(i, 14).value
            tempArray(recordCount, 11) = wsF.Cells(i, 15).value
            tempArray(recordCount, 12) = wsF.Cells(i, 16).value
            tempArray(recordCount, 13) = wsF.Cells(i, 17).value
            tempArray(recordCount, 14) = wsF.Cells(i, 18).value
            tempArray(recordCount, 15) = wsF.Cells(i, 19).value
        End If
    Next i

    If recordCount > 0 Then
        ' Resize array to actual size
        ReDim recordsArray(1 To recordCount, 1 To 15)
        For i = 1 To recordCount
            For j = 1 To 15
                recordsArray(i, j) = tempArray(i, j)
            Next j
        Next i

        ' Sort records
        Call SortRecordsByWeekAndDateRevised(recordsArray, recordCount)

        ' Write data in one operation
        startCell.Offset(1).Resize(recordCount, 15).value = recordsArray

        ' Minimal data formatting - wrap text and thin borders only
        With startCell.Offset(1).Resize(recordCount, 15)
            .HorizontalAlignment = xlLeft
            .WrapText = True
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With

        MsgBox recordCount & " Record(s) displayed for " & entID & " (sorted by week and date)", vbInformation
    Else
        MsgBox "No records found for Enterprise ID: " & entID, vbInformation, "No Records"
    End If

    On Error Resume Next
    wsF.Cells.Locked = True
    wsF.Range("F2,W7").Locked = True
    wsF.Range("C6,E6,H6,H1,B10:J14").Locked = False
    wsF.Range("B3000:T3080").Locked = False
    wsF.Protect UserInterfaceOnly:=True, AllowFiltering:=True
    On Error GoTo 0

    ' Restore settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

Private Sub SortRecordsByWeekAndDateRevised(ByRef arr As Variant, ByVal count As Long)
    Dim i As Long, j As Long, k As Long
    Dim tempRow As Variant
    Dim weekDate1 As Date, weekDate2 As Date
    Dim dateTime1 As Date, dateTime2 As Date

    For i = 1 To count - 1
        For j = 1 To count - i
            weekDate1 = ExtractWeekStartDate(CStr(arr(j, 5)))
            weekDate2 = ExtractWeekStartDate(CStr(arr(j + 1, 5)))
            dateTime1 = CDate(arr(j, 3))
            dateTime2 = CDate(arr(j + 1, 3))

            If weekDate1 < weekDate2 Or (weekDate1 = weekDate2 And dateTime1 > dateTime2) Then
                ReDim tempRow(1 To 15)
                For k = 1 To 15
                    tempRow(k) = arr(j, k)
                    arr(j, k) = arr(j + 1, k)
                    arr(j + 1, k) = tempRow(k)
                Next k
            End If
        Next j
    Next i
End Sub

Private Function ExtractWeekStartDate(weekString As String) As Date
    Dim dateStr As String
    Dim startPos As Long

    startPos = InStr(weekString, "starting ") + 9
    dateStr = Mid(weekString, startPos)

    On Error Resume Next
    ExtractWeekStartDate = CDate(dateStr)
    If Err.Number <> 0 Then
        ExtractWeekStartDate = DateValue("1/1/1900")
        Err.Clear
    End If
    On Error GoTo 0
End Function





Public Sub SyncAllUserSheetsToDatabase()
    On Error GoTo ErrorHandler

    Dim pass As String
    pass = InputBox("Enter sync password:", "Admin Sync Required")
    If pass <> "123" Then
        MsgBox "Incorrect password. Access denied.", vbCritical, "Access Denied"
        Exit Sub
    End If

    Dim ws As Worksheet, wsDB As Worksheet, wsAdmin As Worksheet
    Dim r As Long, nextDBRow As Long
    Dim recordCount As Long, userCount As Long, deletedCount As Long
    Dim syncSuccess As Boolean
    Dim backupCollection As Collection
    
    recordCount = 0
    userCount = 0
    deletedCount = 0
    
    Set wsDB = ThisWorkbook.Sheets("Data base")
    Set wsAdmin = ThisWorkbook.Sheets("Admin")
    Set backupCollection = New Collection

    ' --- STEP 1: DISABLE EVENTS FOR SPEED ---
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    wsDB.Unprotect

    ' --- STEP 2: PROCESS DELETIONS IN DATABASE ---
    Dim lastDBRow As Long
    lastDBRow = wsDB.Cells(wsDB.Rows.count, 1).End(xlUp).Row
    For r = lastDBRow To 2 Step -1
        If wsDB.Cells(r, 20).value = "Marked for Deletion" Then
            wsDB.Rows(r).Delete
            deletedCount = deletedCount + 1
        End If
    Next r

    ' --- STEP 3: FIRST PASS - COPY TO DATABASE (BUT DON'T CLEAR YET) ---
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Data base" And ws.Name <> "Login" And ws.Name <> "Admin" And ws.Name <> "DV" And ws.Name <> "Form" Then
            
            For r = 3000 To 3200
                If ws.Cells(r, 2).value <> "" Then
                    If ws.Cells(r, 20).value <> "Marked for Deletion" Then
                        ' Write to database
                        nextDBRow = wsDB.Cells(wsDB.Rows.count, 1).End(xlUp).Row + 1
                        wsDB.Cells(nextDBRow, 1).Resize(1, 18).value = ws.Cells(r, 2).Resize(1, 18).value
                        recordCount = recordCount + 1
                        
                        ' Store for potential rollback (backup)
                        backupCollection.Add Array(ws.Name, r)
                    End If
                End If
            Next r
        End If
    Next ws

    ' --- STEP 4: SAVE DATABASE FIRST (COMMIT PHASE 1) ---
    wsDB.Protect UserInterfaceOnly:=True
    ThisWorkbook.Save
    
    ' --- STEP 5: SECOND PASS - NOW SAFE TO CLEAR LOCAL (COMMIT PHASE 2) ---
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Data base" And ws.Name <> "Login" And ws.Name <> "Admin" And ws.Name <> "DV" And ws.Name <> "Form" Then
            Dim hasRecords As Boolean
            hasRecords = False
            
            ' Check if this sheet had records
            For r = 3000 To 3200
                If ws.Cells(r, 2).value <> "" And ws.Cells(r, 20).value <> "Marked for Deletion" Then
                    hasRecords = True
                    Exit For
                End If
            Next r
            
            If hasRecords Then
                ' Clear local storage
                ws.Unprotect
                ws.Range("B3000:T3080").ClearContents
                ws.Range("B3000:T3080").Interior.ColorIndex = xlNone
                ws.Protect UserInterfaceOnly:=True, AllowFiltering:=True
                userCount = userCount + 1
            End If
        End If
    Next ws

    ' --- STEP 6: FINALIZE ---
    wsAdmin.Range("G5").value = Format(Now, "yyyy-mm-dd hh:mm:ss")
    ThisWorkbook.Save
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "SYNC COMPLETED SUCCESSFULLY!" & vbCrLf & _
           "Users processed: " & userCount & vbCrLf & _
           "Records synced: " & recordCount & vbCrLf & _
           "Records deleted: " & deletedCount & vbCrLf & _
           "All data safely committed to database." & vbCrLf & _
           "Last sync: " & Format(Now, "yyyy-mm-dd hh:mm:ss"), _
           vbInformation, "Sync Complete"

    Call UpdatePendingCounts
    Exit Sub

ErrorHandler:
    MsgBox "SYNC FAILED - NO DATA LOSS" & vbCrLf & _
           "Error: " & Err.Description & vbCrLf & vbCrLf & _
           "Local data has NOT been cleared. Please try again." & vbCrLf & _
           "Contact administrator if problem persists.", _
           vbCritical, "Sync Failed"
    
    ' Restore settings
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    On Error Resume Next
    wsDB.Protect UserInterfaceOnly:=True
    ThisWorkbook.Save
End Sub

Private Sub ResetFormInputsOnly(f As Worksheet)
    Application.EnableEvents = False
    f.Unprotect

    ' Reset correction mode
    isCorrectionMode = False

    ' Clear engagement data (rows 10-14) and basic selections
    f.Range("B10:J14").ClearContents
    f.Range("B10:J14").Interior.ColorIndex = xlNone
    f.Range("C6,E6,H6").ClearContents

    ' Clear unique code storage
    f.Range("Z10:Z14").ClearContents

    ' Set proper protection
    f.Cells.Locked = True
    f.Range("F2,W7").Locked = True
    f.Range("C6,E6,H6,H1,B10:J14").Locked = False
    f.Range("B3000:T3080").Locked = False

    f.Protect UserInterfaceOnly:=True, AllowFiltering:=True
    Application.EnableEvents = True
End Sub

' ==================== OTHER FUNCTIONS ====================
Public Sub LogoutUser()
    Dim response As VbMsgBoxResult
    Dim wsActive As Worksheet
    Dim tblArea As Range
    Dim sheetName As String

    Set wsActive = ActiveSheet
    sheetName = wsActive.Name

    response = MsgBox("Are you sure you want to logout?" & vbCrLf & _
                     "This will hide your sheet.", vbYesNo + vbQuestion, "Confirm Logout")

    If response = vbYes Then
        isCorrectionMode = False

        ' Clear toggle data area before logout (except for Admin, DV, Data base sheets)
        If sheetName <> "Admin" And sheetName <> "DV" And sheetName <> "Data base" Then
            On Error Resume Next
            wsActive.Unprotect
            On Error GoTo 0

            Set tblArea = wsActive.Range("B25:P297")

            ' Check if data exists in toggle area
            If WorksheetFunction.CountA(tblArea) > 0 Then
                ' Clear all content and formatting
                With tblArea
                    .ClearContents
                    .ClearFormats
                    .Interior.ColorIndex = xlNone
                    .Borders.LineStyle = xlNone
                    .Font.Bold = False
                    .WrapText = False
                    .HorizontalAlignment = xlGeneral
                End With
            End If

            ' Re-protect the sheet
            On Error Resume Next
            wsActive.Cells.Locked = True
            wsActive.Range("F2,W7").Locked = True
            wsActive.Range("C6,E6,H6,H1,B10:J14").Locked = False
            wsActive.Range("B3000:T3080").Locked = False
            wsActive.Protect UserInterfaceOnly:=True, AllowFiltering:=True
            On Error GoTo 0
        End If

        ' Hide sheet and navigate to Admin
        wsActive.Visible = xlSheetVeryHidden
        Sheets("Admin").Activate
        Sheets("Admin").Range("F24").Select

        MsgBox "Logged out successfully." & vbCrLf & _
               "Toggle data cleared to optimize file size." & vbCrLf & _
               "Please close this Excel file.", vbInformation, "Logout Complete"
    End If
End Sub



Public Sub UpdatePendingCounts()
    Dim ws As Worksheet
    Dim totalPending As Long
    Dim userCount As Long
    Dim sheetPending As Long

    totalPending = 0
    userCount = 0

    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Login" And ws.Name <> "Data base" And ws.Name <> "DV" And ws.Name <> "Admin" And ws.Name <> "Form" Then
            sheetPending = 0
            Dim r As Long
            For r = 3000 To 3200
                If ws.Cells(r, 2).value <> "" And ws.Cells(r, 20).value <> "Marked for Deletion" Then
                    sheetPending = sheetPending + 1
                End If
            Next r
            
            If sheetPending > 0 Then
                totalPending = totalPending + sheetPending
                userCount = userCount + 1
            End If
        End If
    Next ws

    With ThisWorkbook.Sheets("Admin")
        .Range("G3").value = totalPending
        .Range("G4").value = userCount
    End With
End Sub

Public Sub LogoutAllUserSheets()
    Dim ws As Worksheet, count As Long
    count = 0

    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Login" And ws.Name <> "Data base" And ws.Name <> "Admin" And ws.Name <> "DV" And ws.Name <> "Form" Then
            ws.Visible = xlSheetVeryHidden
            count = count + 1
        End If
    Next ws

    Sheets("Admin").Activate
End Sub

' ==================== USER FORM DE OPEN FUNCTION ====================
Public Sub OpenTimesheetCorrectionForm()
    On Error GoTo ErrorHandler

    If Trim(ActiveSheet.Range("W6").value) = "" Then
        MsgBox "Please login first. Enterprise ID not found in cell W6.", vbExclamation, "Login Required"
        Exit Sub
    End If

    Dim correctionForm As UserFormDE
    Set correctionForm = New UserFormDE

    correctionForm.LoadEmployeeName

    correctionForm.Show

    Exit Sub

ErrorHandler:
    MsgBox "Error opening correction form: " & Err.Description, vbCritical, "Error"
    On Error Resume Next
    Set correctionForm = Nothing
End Sub



Public Sub ShowAllUserSheets()
    If InputBox("Enter admin password:") <> "123" Then
        MsgBox "Access denied": Exit Sub
    End If
    
    Dim ws As Worksheet, count As Long
    count = 0
    
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Login" And ws.Name <> "Data base" Then
            ws.Visible = xlSheetVisible
            count = count + 1
        End If
    Next ws
    
    MsgBox count & " user sheets are now visible.", vbInformation
End Sub

# Module 2
' ==================== GET NEXT AVAILABLE CODE FOR USER (ONE-TIME SCAN) ====================
Private Function GetNextAvailableCodeForUser(entID As String) As Long
    Dim maxCode As Long
    Dim ws As Worksheet
    Dim r As Long
    Dim wsDB As Worksheet
    Dim lastRow As Long

    maxCode = 0

    ' --- 1. Check the current user sheet if it already exists ---
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(entID)
    On Error GoTo 0

    If Not ws Is Nothing Then
        For r = 3000 To 3200
            If ws.Cells(r, 2).value <> "" And IsNumeric(ws.Cells(r, 2).value) Then
                If CLng(ws.Cells(r, 2).value) > maxCode Then maxCode = CLng(ws.Cells(r, 2).value)
            End If
        Next r
    End If

    ' --- 2. Check the database for this user ---
    On Error Resume Next
    Set wsDB = ThisWorkbook.Sheets("Data base")
    On Error GoTo 0

    If Not wsDB Is Nothing Then
        lastRow = wsDB.Cells(wsDB.Rows.count, 1).End(xlUp).Row
        For r = 2 To lastRow
            If wsDB.Cells(r, 2).value = entID Then
                If IsNumeric(wsDB.Cells(r, 1).value) Then
                    If CLng(wsDB.Cells(r, 1).value) > maxCode Then maxCode = CLng(wsDB.Cells(r, 1).value)
                End If
            End If
        Next r
    End If

    ' Return the next available code
    GetNextAvailableCodeForUser = maxCode + 1
End Function

' ==================== USER MANAGEMENT FUNCTIONS (FIXED) ====================
Public Sub AddUsersFromTable()
    On Error GoTo ErrorHandler

    Dim userRange As Range
    Dim userRow As Range
    Dim enterpriseID As String, employeeName As String, employeeEmail As String
    Dim successCount As Integer, failCount As Integer
    Dim incompleteCount As Integer, existingCount As Integer
    Dim wsDV As Worksheet
    Dim errorDetails As String

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Use the predefined MUser range
    Set userRange = ThisWorkbook.Names("MUser").RefersToRange

    If userRange Is Nothing Then
        MsgBox "MUser range not found. Please check the Name Manager.", vbExclamation
        Exit Sub
    End If

    ' Initialize counters
    successCount = 0
    failCount = 0
    incompleteCount = 0
    existingCount = 0
    errorDetails = ""

    ' Get and unprotect DV sheet
    On Error Resume Next
    Set wsDV = ThisWorkbook.Sheets("DV")
    On Error GoTo ErrorHandler

    If wsDV Is Nothing Then
        MsgBox "Critical Error: DV sheet not found! Please contact admin.", vbCritical, "Missing Sheet"
        GoTo CleanUp
    End If
    wsDV.Unprotect

    ' Loop through each row in the MUser range (skip header row)
    For Each userRow In userRange.Rows
        If userRow.Row > userRange.Row Then ' Skip header row
            enterpriseID = Trim(userRow.Cells(1, 1).value)
            employeeName = Trim(userRow.Cells(1, 2).value)
            employeeEmail = Trim(userRow.Cells(1, 3).value)

            ' Only process rows where all fields are filled
            If enterpriseID <> "" And employeeName <> "" And employeeEmail <> "" Then
                ' Call AddNewUser function with error handling
                On Error Resume Next
                Call AddNewUser(enterpriseID, employeeName, employeeEmail)

                If Err.Number = 0 Then
                    successCount = successCount + 1
                ElseIf Err.Number = vbObjectError + 1 Then
                    existingCount = existingCount + 1
                    Err.Clear
                Else
                    failCount = failCount + 1
                    errorDetails = errorDetails & "- " & enterpriseID & ": " & Err.Description & vbNewLine
                    Err.Clear
                End If
                On Error GoTo ErrorHandler
            ElseIf enterpriseID <> "" Or employeeName <> "" Or employeeEmail <> "" Then
                ' If any field is filled but not all, count as incomplete
                incompleteCount = incompleteCount + 1
            End If
        End If
    Next userRow

    ' Update the Login dropdown after adding all users
    UpdateLoginDropdown

    ' Clear MUser range, leaving the header
    ClearMUserRange userRange

    ' Display results in a single message box
    Dim message As String
    message = "Process completed." & vbNewLine & vbNewLine

    If successCount > 0 Then
        message = message & "? Users added successfully: " & successCount & vbNewLine
    End If

    If incompleteCount > 0 Then
        message = message & "? Users with incomplete details: " & incompleteCount & vbNewLine
    End If

    If existingCount > 0 Then
        message = message & "? Users already existing: " & existingCount & vbNewLine
    End If

    If failCount > 0 Then
        message = message & "? Failed attempts: " & failCount & vbNewLine
        If errorDetails <> "" Then
            message = message & vbNewLine & "Error Details:" & vbNewLine & errorDetails
        End If
    End If

    If successCount = 0 And failCount = 0 And incompleteCount = 0 And existingCount = 0 Then
        message = message & "No users were processed."
    End If

    MsgBox message, vbInformation, "Add Users Result"

    GoTo CleanUp

ErrorHandler:
    MsgBox "Critical error in AddUsersFromTable: " & Err.Description & vbNewLine & _
           "Error Number: " & Err.Number, vbCritical, "Error"

CleanUp:
    ' Re-protect DV sheet
    If Not wsDV Is Nothing Then
        On Error Resume Next
        wsDV.Protect UserInterfaceOnly:=True
        On Error GoTo 0
    End If

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Private Sub ClearMUserRange(userRange As Range)
    ' Clear the MUser range, leaving the header
    If userRange.Rows.count > 1 Then
        userRange.Offset(1).ClearContents
    End If
End Sub

Public Sub AddNewUser(enterpriseID As String, employeeName As String, employeeEmail As String)
    On Error GoTo ErrorHandler

    Dim wsLogin As Worksheet, wsForm As Worksheet, wsDV As Worksheet, wsNew As Worksheet
    Dim nextDVRow As Long
    Dim nextCode As Long

    ' === Check if user already exists ===
    If SheetExists(enterpriseID) Then
        Err.Raise vbObjectError + 1, , "User '" & enterpriseID & "' already exists!"
    End If

    ' === Get Template Form (even if very hidden) ===
    Set wsForm = GetVeryHiddenSheet("Form")
    If wsForm Is Nothing Then
        Err.Raise vbObjectError + 2, , "Template sheet 'Form' not found!"
    End If

    ' === Get DV sheet ===
    On Error Resume Next
    Set wsDV = ThisWorkbook.Sheets("DV")
    On Error GoTo ErrorHandler

    If wsDV Is Nothing Then
        Err.Raise vbObjectError + 3, , "DV sheet not found!"
    End If

    ' === Create new sheet by copying Form ===
    Dim originalVisibility As XlSheetVisibility
    originalVisibility = wsForm.Visible
    wsForm.Visible = xlSheetVisible  ' Temporarily make Form visible
    wsForm.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
    wsForm.Visible = originalVisibility  ' Set Form back to original visibility
    Set wsNew = ActiveSheet

    ' === Rename to Enterprise ID ===
    On Error Resume Next
    wsNew.Name = enterpriseID
    If Err.Number <> 0 Then
        Application.DisplayAlerts = False
        wsNew.Delete
        Application.DisplayAlerts = True
        Err.Raise vbObjectError + 4, , "Cannot rename sheet to '" & enterpriseID & "'. Please use a valid sheet name."
    End If
    On Error GoTo ErrorHandler

    ' === Setup new sheet data with FIXED RANGES ===
    With wsNew
        .Unprotect

        ' Clear any existing data in key cells
        .Range("W6,F2,W7,AZ1,BA1").ClearContents

        ' Fill user data with new cell references
        .Range("W6").value = enterpriseID        ' Enterprise ID in W6
        .Range("F2").value = employeeName        ' Name in F2
        .Range("W7").value = employeeEmail       ' Email in W7
        .Range("AZ1").value = "Added: " & Format(Now, "yyyy-mm-dd hh:mm:ss")

        ' **FIXED: Initialize next code in BA1 using wsNew instead of ws**
        nextCode = GetNextAvailableCodeForUser(enterpriseID)
        .Range("BA1").value = nextCode
        .Range("BA1").Locked = False

        ' Remove fill formatting from ID and Email cells
        .Range("W6,W7").Interior.ColorIndex = xlNone

        ' Set protection with FIXED ranges
        .Cells.Locked = True

        ' Lock the ID and Email cells (they should not be editable)
        .Range("W6,W7").Locked = True

        ' Unlock data entry fields as per new structure
        .Range("F2").Locked = False          ' Name
        .Range("C6").Locked = False          ' Month
        .Range("E6").Locked = False          ' Week
        .Range("H6").Locked = False          ' Team name
        .Range("B10:B14").Locked = False     ' Region
        .Range("C10:C14").Locked = False     ' Audit Engagement ID
        .Range("D10:D14").Locked = False     ' Engagement Activity
        .Range("E10:E14").Locked = False     ' Hours
        .Range("F10:F14").Locked = False     ' Remark
        .Range("H10:H14").Locked = False     ' Non-Audit Engagement
        .Range("I10:I14").Locked = False     ' Hours
        .Range("J10:J14").Locked = False     ' Remark
        .Range("AT1").Locked = False         ' Additional field
        .Range("B3000:T3080").Locked = False ' FIXED STORAGE AREA

        .Protect UserInterfaceOnly:=True, AllowFiltering:=True
    End With

    ' === Add to DV sheet ===
    wsDV.Unprotect
    nextDVRow = wsDV.Cells(wsDV.Rows.count, "F").End(xlUp).Row + 1
    If nextDVRow < 2 Then nextDVRow = 2 ' Start from row 2 if sheet is empty
    wsDV.Range("F" & nextDVRow).value = enterpriseID
    wsDV.Range("G" & nextDVRow).value = employeeName
    wsDV.Range("H" & nextDVRow).value = employeeEmail
    wsDV.Protect UserInterfaceOnly:=True

    ' === Try to copy VBA code (optional - might fail if security high) ===
    On Error Resume Next
    CopySheetMacrosSafe wsForm, wsNew
    On Error GoTo ErrorHandler

    ' === Update counters if function exists ===
    On Error Resume Next
    Call UpdatePendingCounts
    On Error GoTo ErrorHandler

    ' === Clear specified cells on the new user's sheet ===
    ClearNewUserFields wsNew

    Exit Sub

ErrorHandler:
    ' Clean up if error occurred
    If Not wsNew Is Nothing Then
        On Error Resume Next
        Application.DisplayAlerts = False
        wsNew.Delete
        Application.DisplayAlerts = True
        On Error GoTo 0
    End If

    Err.Raise Err.Number, , Err.Description
End Sub

Private Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Sub CopySheetMacrosSafe(sourceSheet As Worksheet, targetSheet As Worksheet)
    On Error Resume Next

    Dim sourceCode As Object, targetCode As Object
    Set sourceCode = ThisWorkbook.VBProject.VBComponents(sourceSheet.CodeName)
    Set targetCode = ThisWorkbook.VBProject.VBComponents(targetSheet.CodeName)

    If Not sourceCode Is Nothing And Not targetCode Is Nothing Then
        If sourceCode.CodeModule.CountOfLines > 0 Then
            targetCode.CodeModule.DeleteLines 1, targetCode.CodeModule.CountOfLines
            targetCode.CodeModule.AddFromString sourceCode.CodeModule.Lines(1, sourceCode.CodeModule.CountOfLines)
        End If
    End If

    ' If error occurs, just continue - VBA copying is optional
    On Error GoTo 0
End Sub

Private Sub UpdateLoginDropdown()
    On Error GoTo ErrorHandler

    Dim wsDV As Worksheet
    Dim lastRow As Long, i As Long
    Dim itemValues As Variant

    ' Check if DV sheet exists
    If Not SheetExists("DV") Then
        MsgBox "DV sheet not found. Please check your workbook structure.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set wsDV = ThisWorkbook.Worksheets("DV")
    On Error GoTo ErrorHandler

    If wsDV Is Nothing Then
        MsgBox "DV worksheet not found.", vbExclamation
        Exit Sub
    End If

    ' Temporarily unprotect DV sheet
    wsDV.Unprotect

    ' Clear existing items in ComboBox
    UserFormLOGIN.ComboBoxID.Clear

    ' Find last row with data in column F
    lastRow = wsDV.Cells(wsDV.Rows.count, "F").End(xlUp).Row

    If lastRow < 2 Then
        ' No data in DV sheet, leave ComboBox empty
        GoTo CleanUp
    End If

    ' Get the values from DV sheet
    itemValues = wsDV.Range("F2:F" & lastRow).value

    ' Populate ComboBox with values
    If IsArray(itemValues) Then
        For i = 1 To UBound(itemValues, 1)
            If itemValues(i, 1) <> "" Then ' Only add non-empty values
                UserFormLOGIN.ComboBoxID.AddItem itemValues(i, 1)
            End If
        Next i
    Else
        ' Single value case
        If itemValues <> "" Then
            UserFormLOGIN.ComboBoxID.AddItem itemValues
        End If
    End If

    GoTo CleanUp

ErrorHandler:
    MsgBox "An error occurred while updating the Login dropdown: " & Err.Description, vbCritical, "Error"

CleanUp:
    ' Re-protect DV sheet
    If Not wsDV Is Nothing Then
        On Error Resume Next
        wsDV.Protect UserInterfaceOnly:=True
        On Error GoTo 0
    End If

    ' Set ComboBox style (optional)
    With UserFormLOGIN.ComboBoxID
        .Style = fmStyleDropDownCombo ' Allow typing or selection
        ' .Style = fmStyleDropDownList ' Only allow selection from list
    End With
End Sub

Private Function GetVeryHiddenSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = sheetName Then
            Set GetVeryHiddenSheet = ws
            Exit Function
        End If
    Next ws
End Function

Private Sub ClearNewUserFields(ws As Worksheet)
    Application.EnableEvents = False
    ws.Unprotect

    ' Clear all data entry fields according to new structure
    ws.Range("C6,E6,H6").ClearContents                    ' Month, Week, Team name
    ws.Range("B10:J14").ClearContents                     ' All engagement fields

    ' Set protection with FIXED ranges
    ws.Cells.Locked = True

    ' Lock ID and Email (not editable)
    ws.Range("W6,W7").Locked = True

    ' Unlock all data entry fields with FIXED ranges
    ws.Range("F2").Locked = False          ' Name
    ws.Range("C6").Locked = False          ' Month
    ws.Range("E6").Locked = False          ' Week
    ws.Range("H6").Locked = False          ' Team name
    ws.Range("B10:J14").Locked = False     ' Engagement area
    ws.Range("AT1").Locked = False         ' Additional field
    ws.Range("B3000:T3080").Locked = False ' FIXED STORAGE AREA

    ws.Protect UserInterfaceOnly:=True, AllowFiltering:=True
    Application.EnableEvents = True
End Sub


# Module 3
Public Sub HideAllUserSheets()
    If InputBox("Enter admin password:") <> "123" Then
        MsgBox "Access denied": Exit Sub
    End If
    
    Dim ws As Worksheet, count As Long
    count = 0
    
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Admin" And ws.Name <> "Data base" Then
            ws.Visible = xlSheetVeryHidden
            count = count + 1
        End If
    Next ws
    
    Sheets("Admin").Activate
    MsgBox count & " user sheets have been hidden.", vbInformation
End Sub


' Version with error handling
Sub ShowLogin()
    On Error Resume Next
    UserFormLOGIN.Show vbModal
    On Error GoTo 0
End Sub

' Version with confirmation
Sub DisplayLoginForm()
    UserFormLOGIN.Show
End Sub



Option Explicit

' ==================== DV SHEET DROPDOWN FUNCTIONS ====================

Public Sub PopulateComboFromDVColumn(comboBox As MSForms.comboBox, ColumnLetter As String)
    ' Populate combobox with unique values from specified column in DV sheet
    On Error GoTo ErrorHandler
    
    comboBox.Clear
    comboBox.Style = fmStyleDropDownCombo ' Allow typing and dropdown
    
    Dim wsDV As Worksheet
    Set wsDV = ThisWorkbook.Worksheets("DV")
    
    If wsDV Is Nothing Then
        MsgBox "DV worksheet not found.", vbExclamation
        Exit Sub
    End If
    
    Dim lastRow As Long
    lastRow = wsDV.Cells(wsDV.Rows.count, ColumnLetter).End(xlUp).Row
    
    If lastRow < 2 Then
        Exit Sub ' No data
    End If
    
    ' Collection to store unique values
    Dim uniqueItems As Collection
    Set uniqueItems = New Collection
    
    Dim i As Long
    For i = 2 To lastRow
        Dim cellValue As String
        cellValue = Trim(wsDV.Cells(i, ColumnLetter).value)
        
        If cellValue <> "" Then
            ' Add only unique values using error handling
            On Error Resume Next
            uniqueItems.Add cellValue, CStr(cellValue)
            On Error GoTo 0
        End If
    Next i
    
    ' Add unique items to combobox
    Dim item As Variant
    For Each item In uniqueItems
        comboBox.AddItem item
    Next item
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error populating combobox from DV sheet: " & Err.Description, vbExclamation
    On Error GoTo 0
End Sub

Public Sub PopulateComboFromDVWithCollection(comboBox As MSForms.comboBox, ColumnLetter As String, originalItems As Collection)
    ' Populate combobox with unique values and store in collection
    On Error GoTo ErrorHandler
    
    comboBox.Clear
    comboBox.Style = fmStyleDropDownCombo
    
    Set originalItems = New Collection
    
    Dim wsDV As Worksheet
    Set wsDV = ThisWorkbook.Worksheets("DV")
    
    If wsDV Is Nothing Then
        MsgBox "DV worksheet not found.", vbExclamation
        Exit Sub
    End If
    
    Dim lastRow As Long
    lastRow = wsDV.Cells(wsDV.Rows.count, ColumnLetter).End(xlUp).Row
    
    If lastRow < 2 Then
        Exit Sub ' No data
    End If
    
    ' Collection to store unique values
    Dim uniqueItems As Collection
    Set uniqueItems = New Collection
    
    Dim i As Long
    For i = 2 To lastRow
        Dim cellValue As String
        cellValue = Trim(wsDV.Cells(i, ColumnLetter).value)
        
        If cellValue <> "" Then
            ' Add only unique values using error handling
            On Error Resume Next
            uniqueItems.Add cellValue, CStr(cellValue)
            On Error GoTo 0
        End If
    Next i
    
    ' Add unique items to combobox and collection
    Dim item As Variant
    For Each item In uniqueItems
        comboBox.AddItem item
        originalItems.Add item
    Next item
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error populating combobox from DV sheet: " & Err.Description, vbExclamation
    On Error GoTo 0
End Sub

Public Sub PopulateDependentEngagementCombo(comboBox As MSForms.comboBox, teamValue As String, regionValue As String)
    ' Populate engagement combobox based on selected Team and Region
    On Error GoTo ErrorHandler
    
    comboBox.Clear
    comboBox.Style = fmStyleDropDownCombo
    
    ' Exit if either team or region is empty
    If teamValue = "" Or regionValue = "" Then
        Exit Sub
    End If
    
    Dim wsDV As Worksheet
    Set wsDV = ThisWorkbook.Worksheets("DV")
    
    If wsDV Is Nothing Then
        MsgBox "DV worksheet not found.", vbExclamation
        Exit Sub
    End If
    
    Dim lastRow As Long
    lastRow = wsDV.Cells(wsDV.Rows.count, "N").End(xlUp).Row
    
    If lastRow < 2 Then
        Exit Sub ' No data
    End If
    
    ' Collection to store unique engagement values
    Dim uniqueItems As Collection
    Set uniqueItems = New Collection
    
    Dim i As Long
    For i = 2 To lastRow
        Dim currentTeam As String
        Dim currentRegion As String
        Dim engagementValue As String
        
        currentTeam = Trim(wsDV.Cells(i, "N").value)
        currentRegion = Trim(wsDV.Cells(i, "O").value)
        engagementValue = Trim(wsDV.Cells(i, "P").value)
        
        ' Match both Team AND Region (case-insensitive)
        If StrComp(currentTeam, teamValue, vbTextCompare) = 0 And _
           StrComp(currentRegion, regionValue, vbTextCompare) = 0 And _
           engagementValue <> "" Then
            
            ' Add only unique values using error handling
            On Error Resume Next
            uniqueItems.Add engagementValue, CStr(engagementValue)
            On Error GoTo 0
        End If
    Next i
    
    ' Add unique items to combobox
    Dim item As Variant
    For Each item In uniqueItems
        comboBox.AddItem item
    Next item
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error loading dependent engagement dropdown: " & Err.Description, vbExclamation
    On Error GoTo 0
End Sub

' ==================== VALIDATION FUNCTIONS ====================

Public Function ValidateComboBoxValue(comboBox As MSForms.comboBox, value As String) As Boolean
    ' Check if value exists in combobox list
    ValidateComboBoxValue = False
    
    Dim i As Integer
    For i = 0 To comboBox.ListCount - 1
        If StrComp(comboBox.List(i), value, vbTextCompare) = 0 Then
            ValidateComboBoxValue = True
            Exit Function
        End If
    Next i
End Function

Public Function GetEngagementListForTeamRegion(teamValue As String, regionValue As String) As Collection
    ' Returns collection of engagement IDs for given Team and Region
    Set GetEngagementListForTeamRegion = New Collection
    
    On Error GoTo ErrorHandler
    
    Dim wsDV As Worksheet
    Set wsDV = ThisWorkbook.Worksheets("DV")
    
    If wsDV Is Nothing Then
        Exit Function
    End If
    
    ' Exit if either team or region is empty
    If teamValue = "" Or regionValue = "" Then
        Exit Function
    End If
    
    Dim lastRow As Long
    lastRow = wsDV.Cells(wsDV.Rows.count, "N").End(xlUp).Row
    
    If lastRow < 2 Then
        Exit Function ' No data
    End If
    
    Dim i As Long
    For i = 2 To lastRow
        Dim currentTeam As String
        Dim currentRegion As String
        Dim engagementValue As String
        
        currentTeam = Trim(wsDV.Cells(i, "N").value)
        currentRegion = Trim(wsDV.Cells(i, "O").value)
        engagementValue = Trim(wsDV.Cells(i, "P").value)
        
        ' Match both Team AND Region (case-insensitive)
        If StrComp(currentTeam, teamValue, vbTextCompare) = 0 And _
           StrComp(currentRegion, regionValue, vbTextCompare) = 0 And _
           engagementValue <> "" Then
            
            ' Add only unique values using error handling
            On Error Resume Next
            GetEngagementListForTeamRegion.Add engagementValue, CStr(engagementValue)
            On Error GoTo 0
        End If
    Next i
    
    Exit Function
    
ErrorHandler:
    Set GetEngagementListForTeamRegion = New Collection
    On Error GoTo 0
End Function

' ==================== LEGACY FUNCTIONS (for compatibility) ====================

Public Sub PopulateComboFromNameManagerUnique(comboBox As MSForms.comboBox, nameManagerName As String)
    ' Legacy function - now uses DV sheet for Team, Region, Engagement
    ' Falls back to named range if needed for other dropdowns
    
    On Error Resume Next
    
    comboBox.Clear
    comboBox.Style = fmStyleDropDownCombo
    
    ' Check if this is for Team, Region, or Engagement
    If nameManagerName = "Team_Name" Then
        ' Use DV sheet for Team
        PopulateComboFromDVColumn comboBox, "N"
        Exit Sub
    ElseIf nameManagerName = "Region" Then
        ' Use DV sheet for Region
        PopulateComboFromDVColumn comboBox, "O"
        Exit Sub
    End If
    
    ' For other named ranges, use original logic
    Dim nm As Name
    Set nm = ThisWorkbook.Names(nameManagerName)
    
    If Not nm Is Nothing Then
        Dim rng As Range
        Set rng = Range(nm.RefersTo)
        
        Dim uniqueItems As Collection
        Set uniqueItems = New Collection
        
        Dim cell As Range
        For Each cell In rng
            If cell.value <> "" Then
                ' Add only unique values
                On Error Resume Next
                uniqueItems.Add cell.value, CStr(cell.value)
                On Error GoTo 0
            End If
        Next cell
        
        ' Add unique items to combobox
        Dim item As Variant
        For Each item In uniqueItems
            comboBox.AddItem item
        Next item
    End If
    
    On Error GoTo 0
End Sub

Public Sub PopulateComboFromNameManagerUniqueWithCollection(comboBox As MSForms.comboBox, nameManagerName As String, originalItems As Collection)
    ' Legacy function with collection support
    
    On Error Resume Next
    
    comboBox.Clear
    comboBox.Style = fmStyleDropDownCombo
    Set originalItems = New Collection
    
    ' Check if this is for Engagement (region-based in old system)
    If InStr(1, nameManagerName, "_", vbTextCompare) > 0 Then
        ' This was originally for region-based engagements
        ' Now we use Team+Region dependency, so skip this
        Exit Sub
    End If
    
    Dim nm As Name
    Set nm = ThisWorkbook.Names(nameManagerName)
    
    If Not nm Is Nothing Then
        Dim rng As Range
        Set rng = Range(nm.RefersTo)
        
        Dim uniqueItems As Collection
        Set uniqueItems = New Collection
        
        Dim cell As Range
        For Each cell In rng
            If cell.value <> "" Then
                ' Add only unique values
                On Error Resume Next
                uniqueItems.Add cell.value, CStr(cell.value)
                On Error GoTo 0
            End If
        Next cell
        
        ' Add unique items to combobox and collection
        Dim item As Variant
        For Each item In uniqueItems
            comboBox.AddItem item
            originalItems.Add item
        Next item
    End If
    
    On Error GoTo 0
End Sub



# Module 4
Option Explicit

' ============================================
' WORKSHEET ACTIVATE EVENT HANDLER
' ============================================
Public Sub HandleActivate(ws As Worksheet)
    If ws.Range("W6").value <> "" Then ws.Range("C6").Select
End Sub

' ============================================
' WORKSHEET CHANGE EVENT HANDLER
' ============================================
Public Sub HandleChange(ws As Worksheet, Target As Range)
    On Error GoTo CleanUp
    Application.EnableEvents = False
    ws.Unprotect

    Dim cell As Range, targetRow As Long

    ' Region changed (B10:B14) - clear corresponding Engagement (C10:C14)
    If Not Intersect(Target, ws.Range("B10:B14")) Is Nothing Then
        For Each cell In Intersect(Target, ws.Range("B10:B14"))
            targetRow = cell.Row
            If Trim(ws.Range("C" & targetRow).value) <> "" Then
                ws.Range("C" & targetRow).ClearContents
                MsgBox "Region changed in row " & targetRow & ". Engagement cleared.", vbInformation
            End If
        Next cell
    End If

    ' Engagement requires Region validation
    If Not Intersect(Target, ws.Range("C10:C14")) Is Nothing Then
        For Each cell In Intersect(Target, ws.Range("C10:C14"))
            If Trim(cell.value) <> "" Then
                targetRow = cell.Row
                If Trim(ws.Range("B" & targetRow).value) = "" Then
                    MsgBox "Please select Region first for row " & targetRow & ".", vbExclamation
                    cell.ClearContents
                    ws.Range("B" & targetRow).Select
                    GoTo CleanUp
                End If
            End If
        Next cell
    End If

CleanUp:
    ws.Protect UserInterfaceOnly:=True, AllowFiltering:=True
    Application.EnableEvents = True
End Sub

' ============================================
' WORKSHEET SELECTION CHANGE EVENT HANDLER
' ============================================
Public Sub HandleSelectionChange(ws As Worksheet, Target As Range)
    Static lastTarget As Range, lastActiveCell As Range

    ' Clear previous highlighting
    If Not lastTarget Is Nothing Then
        On Error Resume Next
        lastTarget.Interior.ColorIndex = xlNone
        Set lastTarget = Nothing
        On Error GoTo 0
    End If

    If Not lastActiveCell Is Nothing Then
        On Error Resume Next
        lastActiveCell.Interior.ColorIndex = xlNone
        Set lastActiveCell = Nothing
        On Error GoTo 0
    End If

    If Target.Cells.count = 1 Then
        On Error Resume Next

        ' Special header cells - light green
        If Target.Address = "$C$6" Or Target.Address = "$E$6" Or Target.Address = "$H$6" Then
            Target.Interior.Color = RGB(221, 230, 225)
            Set lastActiveCell = Target

        ' Engagement entry area (B10:J14, excluding G)
        ElseIf Target.Row >= 10 And Target.Row <= 14 And Target.Column >= 2 And Target.Column <= 10 Then
            ws.Range("B" & Target.Row & ":F" & Target.Row & ",H" & Target.Row & ":J" & Target.Row).Interior.Color = RGB(255, 248, 224)
            Set lastTarget = ws.Range("B" & Target.Row & ":F" & Target.Row & ",H" & Target.Row & ":J" & Target.Row)

            If Target.Column <> 7 Then
                Target.Interior.Color = RGB(110, 210, 223)
                Set lastActiveCell = Target
            End If
        End If

        On Error GoTo 0
    End If
End Sub

' ============================================
' WORKSHEET DEACTIVATE EVENT HANDLER
' ============================================
Public Sub HandleDeactivate(ws As Worksheet)
    On Error Resume Next
    ws.Range("B10:F14,H10:J14,C6,E6,H6").Interior.ColorIndex = xlNone
    On Error GoTo 0
End Sub


# This Workbook
Option Explicit

' ============================================================================
' WINDOWS API DECLARATIONS FOR TOPMOST WINDOW
' ============================================================================

Private Declare PtrSafe Function SetWindowPos Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal hWndInsertAfter As LongPtr, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String) As LongPtr

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2

' ============================================================================
' CONSTANTS FOR DROPDOWN STORAGE
' ============================================================================

Private Const DROPDOWN_START_COLUMN As Long = 61  ' Column BI (61st column)
Private Const DROPDOWN_START_ROW As Long = 2



' ============================================================================
' WORKBOOK EVENT HANDLERS
' ============================================================================

Private Sub Workbook_Open()
    Dim ws As Worksheet
    Dim excelHwnd As LongPtr

    On Error Resume Next

    Application.Visible = False
    Application.ScreenUpdating = False

    ' Hide Login sheet
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "Login" Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws

    Application.ScreenUpdating = True

    ' Make Excel window stay on top ONLY for login form
    excelHwnd = Application.hwnd
    If excelHwnd <> 0 Then
        SetWindowPos excelHwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    End If

    On Error GoTo 0

    ' Show form modally to keep it on top
    UserFormLOGIN.Show vbModal

    ' Remove topmost status immediately after login form closes
    On Error Resume Next
    If excelHwnd <> 0 Then
        SetWindowPos excelHwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    End If
    On Error GoTo 0
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    On Error Resume Next
    If Sh.Name = "Admin" Then
        Call UpdatePendingCounts
    End If
End Sub

' ============================================================================
' SELECTION CHANGE EVENT - NO SCREEN FLICKER
' ============================================================================

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    Const TEAM_CELL As String = "H6"
    Const REGION_START As String = "B10"
    Const ENGAGEMENT_START As String = "C10"

    Dim engagementRange As Range
    Dim targetRowOffset As Long
    Dim regionCell As Range
    Dim engagementCell As Range
    Dim teamValue As String
    Dim regionValue As String
    Dim originalScreenUpdating As Boolean
    Dim originalCalculation As XlCalculation
    Dim wasProtected As Boolean

    On Error GoTo ErrorHandler

    ' Only process user sheets
    If Sh.Name = "DV" Or Sh.Name = "Data base" Or Sh.Name = "Admin" Or Sh.Name = "Login" Then Exit Sub

    ' Check if user selected an engagement cell (C10:C14)
    Set engagementRange = Sh.Range("C10:C14")

    If Not Intersect(Target, engagementRange) Is Nothing Then
        ' Only process single cell selection
        If Target.Cells.count > 1 Then Exit Sub

        Set engagementCell = Target

        ' Calculate which row offset this engagement cell is
        targetRowOffset = engagementCell.Row - Sh.Range(ENGAGEMENT_START).Row

        If targetRowOffset >= 0 And targetRowOffset <= 4 Then
            ' Skip if it's a header cell
            If IsHeaderCell(engagementCell) Then Exit Sub

            ' Get the region from THIS SPECIFIC ROW
            Set regionCell = Sh.Range(REGION_START).Offset(targetRowOffset, 0)
            regionValue = Trim(regionCell.value)

            ' Get Team value
            teamValue = Trim(Sh.Range(TEAM_CELL).value)

            ' Only refresh if both Team and Region exist
            If teamValue <> "" And regionValue <> "" Then
                ' Store original settings
                originalScreenUpdating = Application.ScreenUpdating
                originalCalculation = Application.Calculation

                ' Disable updates for smooth operation
                Application.ScreenUpdating = False
                Application.Calculation = xlCalculationManual
                Application.EnableEvents = False

                ' Check if sheet is protected
                wasProtected = Sh.ProtectContents

                If wasProtected Then
                    UnprotectSheet Sh
                End If

                ' Ensure this engagement cell has the correct validation
                Call EnsureEngagementValidation(Sh, engagementCell, teamValue, regionValue, targetRowOffset)

                ' Only reprotect if it was protected before
                If wasProtected Then
                    ProtectSheet Sh
                End If

                ' Restore settings
                Application.EnableEvents = True
                Application.Calculation = originalCalculation
                Application.ScreenUpdating = originalScreenUpdating
            End If
        End If
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "SheetSelectionChange Error: " & Err.Number & " - " & Err.Description
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

' ============================================================================
' SHEET CHANGE EVENT - SMART PROTECTION
' ============================================================================

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Const YEAR_CELL As String = "H1"
    Const MONTH_CELL As String = "C6"
    Const TEAM_CELL As String = "H6"
    Const REGION_RANGE As String = "B10:B14"

    Dim originalScreenUpdating As Boolean
    Dim originalCalculation As XlCalculation
    Dim needsProcessing As Boolean
    Dim wasProtected As Boolean

       On Error GoTo ErrorHandler

    ' Skip system sheets
    If Sh.Name = "DV" Or Sh.Name = "Data base" Or Sh.Name = "Admin" Or Sh.Name = "Login" Then Exit Sub

    ' Determine if this change needs our processing
    needsProcessing = False

    ' Check if change is in monitored ranges
    If Not (Intersect(Target, Sh.Range(YEAR_CELL)) Is Nothing And _
            Intersect(Target, Sh.Range(MONTH_CELL)) Is Nothing) Then
        needsProcessing = True
    End If

    ' Check Team/Region changes (skip system sheets)
    If Sh.Name <> "DV" And Sh.Name <> "Data base" And Sh.Name <> "Admin" And Sh.Name <> "Login" Then
        If Not Intersect(Target, Sh.Range(TEAM_CELL)) Is Nothing Or _
           Not Intersect(Target, Sh.Range(REGION_RANGE)) Is Nothing Then
            needsProcessing = True
        End If
    End If

    ' If this change doesn't need our processing, exit without protecting
    If Not needsProcessing Then Exit Sub

    ' Store original settings
    originalScreenUpdating = Application.ScreenUpdating
    originalCalculation = Application.Calculation

    ' Disable updates for smooth operation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Check if sheet was protected
    wasProtected = Sh.ProtectContents

    If wasProtected Then
        UnprotectSheet Sh
    End If

    ' Handle Year/Month changes
    If Not (Intersect(Target, Sh.Range(YEAR_CELL)) Is Nothing And _
            Intersect(Target, Sh.Range(MONTH_CELL)) Is Nothing) Then
        HandleYearMonthChange Sh, Target
    End If

    ' Handle Team/Region changes (skip system sheets)
    If Sh.Name <> "DV" And Sh.Name <> "Data base" And Sh.Name <> "Admin" And Sh.Name <> "Login" Then
        If Not Intersect(Target, Sh.Range(TEAM_CELL)) Is Nothing Or _
           Not Intersect(Target, Sh.Range(REGION_RANGE)) Is Nothing Then
            HandleTeamRegionChange Sh, Target
        End If
    End If

    ' Only reprotect if sheet was protected before
    If wasProtected Then
        ProtectSheet Sh
    End If

    ' Restore settings
    Application.EnableEvents = True
    Application.Calculation = originalCalculation
    Application.ScreenUpdating = originalScreenUpdating

    Exit Sub

ErrorHandler:
    Debug.Print "Workbook_SheetChange Error " & Err.Number & ": " & Err.Description
    ' Only protect if it was protected before
    If wasProtected Then
        ProtectSheet Sh
    End If
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim userName As String
    Dim ws As Worksheet
    Dim excelHwnd As LongPtr

    On Error Resume Next

    Application.ScreenUpdating = False

    ' Ensure topmost is removed before closing (safety check)
    excelHwnd = Application.hwnd
    If excelHwnd <> 0 Then
        SetWindowPos excelHwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    End If

    userName = Environ("USERNAME")
    For Each ws In ThisWorkbook.Worksheets
        If LCase(ws.Name) = LCase(userName) Then
            If ws.Visible <> xlSheetHidden Then
                ws.Visible = xlSheetHidden
            End If
            ' Cleanup named ranges for this user sheet
            Call CleanupSheetNamedRanges(ws)
            Exit For
        End If
    Next ws

    Application.ScreenUpdating = True

    On Error GoTo 0
End Sub

' ============================================================================
' YEAR & MONTH CHANGE HANDLER
' ============================================================================

Private Sub HandleYearMonthChange(Sh As Worksheet, Target As Range)
    Const YEAR_CELL As String = "H1"
    Const MONTH_CELL As String = "C6"
    Const WEEK_CELL As String = "E6"

    Dim selectedYear As Integer
    Dim selectedMonth As String
    Dim firstDayOfMonth As Date
    Dim lastDayOfMonth As Date
    Dim currentDate As Date
    Dim weekList As New Collection
    Dim weekString As String
    Dim weekArray() As String
    Dim i As Integer

    On Error GoTo ErrorHandler

    selectedYear = Sh.Range(YEAR_CELL).value
    selectedMonth = Sh.Range(MONTH_CELL).value

    If selectedYear = 0 Or selectedMonth = "" Then
        Sh.Range(WEEK_CELL).value = ""
        If HasValidation(Sh.Range(WEEK_CELL)) Then
            Sh.Range(WEEK_CELL).Validation.Delete
        End If
        Exit Sub
    End If

    firstDayOfMonth = DateSerial(selectedYear, Month(selectedMonth & " 1"), 1)
    lastDayOfMonth = DateSerial(selectedYear, Month(selectedMonth & " 1") + 1, 0)

    currentDate = firstDayOfMonth
    Do While currentDate <= lastDayOfMonth
        If Weekday(currentDate, vbMonday) = 1 Then
            weekString = "Week starting " & Format(currentDate, "mmmm d yyyy")
            weekList.Add weekString
        End If
        currentDate = currentDate + 1
    Loop

    If weekList.count > 0 Then
        ReDim weekArray(1 To weekList.count)
        For i = 1 To weekList.count
            weekArray(i) = weekList(i)
        Next i

        If HasValidation(Sh.Range(WEEK_CELL)) Then
            Sh.Range(WEEK_CELL).Validation.Delete
        End If

        Sh.Range(WEEK_CELL).Validation.Add _
            Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Formula1:=Join(weekArray, ",")

        Sh.Range(WEEK_CELL).value = ""
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "HandleYearMonthChange Error: " & Err.Number & " - " & Err.Description
End Sub

' ============================================================================
' TEAM & REGION CHANGE HANDLER - CONFLICT-FREE VERSION
' ============================================================================

Private Sub HandleTeamRegionChange(Sh As Worksheet, Target As Range)
    Const TEAM_CELL As String = "H6"
    Const REGION_START As String = "B10"
    Const ENGAGEMENT_START As String = "C10"

    Dim teamValue As String
    Dim regionCell As Range
    Dim engagementCell As Range
    Dim targetRowOffset As Long
    Dim i As Long
    Dim regionValue As String
    Dim changedInRegionRange As Boolean
    Dim isRegionCleared As Boolean

    On Error GoTo ErrorHandler

    teamValue = Trim(Sh.Range(TEAM_CELL).value)
    changedInRegionRange = Not Intersect(Target, Sh.Range("B10:B14")) Is Nothing
    If changedInRegionRange Then
        ' ===== REGION CHANGED - UPDATE CORRESPONDING ENGAGEMENT =====
        Dim cell As Range
        Dim teamMissing As Boolean
        teamMissing = False
        
        For Each cell In Intersect(Target, Sh.Range("B10:B14"))
            targetRowOffset = cell.Row - Sh.Range(REGION_START).Row

            If targetRowOffset >= 0 And targetRowOffset <= 4 Then
                Set regionCell = Sh.Range(REGION_START).Offset(targetRowOffset, 0)
                Set engagementCell = Sh.Range(ENGAGEMENT_START).Offset(targetRowOffset, 0)

                regionValue = Trim(regionCell.value)
                isRegionCleared = (regionValue = "")

                If Not IsHeaderCell(engagementCell) Then
                    engagementCell.value = ""
                    If HasValidation(engagementCell) Then
                        engagementCell.Validation.Delete
                    End If

                    If isRegionCleared Then
                        Call BlockCellInput(engagementCell)
                    ElseIf teamValue = "" Then
                        Call BlockCellInput(engagementCell)
                        teamMissing = True
                    Else
                        ' Write dropdown data to USER SHEET (not DV)
                        Call ApplyEngagementValidation(Sh, engagementCell, teamValue, regionValue, targetRowOffset)
                    End If
                End If
            End If
        Next cell
        
        ' Show one message if team was missing for any changed region
        If teamMissing And Target.Cells.count = 1 Then
            Application.ScreenUpdating = True
            MsgBox "Please select Team first to choose Engagement ID_Name", vbExclamation, "Team Required"
            Application.ScreenUpdating = False
        End If
    Else
        ' ===== TEAM CHANGED - UPDATE ALL REGIONS =====
        If teamValue = "" Then
            For i = 0 To 4
                Set engagementCell = Sh.Range(ENGAGEMENT_START).Offset(i, 0)

                If Not IsHeaderCell(engagementCell) Then
                    engagementCell.value = ""
                    If HasValidation(engagementCell) Then
                        engagementCell.Validation.Delete
                    End If
                    Call BlockCellInput(engagementCell)
                End If
            Next i
            Application.ScreenUpdating = True
            MsgBox "Please select Team first to choose Engagement ID_Name", vbExclamation, "Team Required"
            Application.ScreenUpdating = False
            Exit Sub
        End If

        For i = 0 To 4
            Set regionCell = Sh.Range(REGION_START).Offset(i, 0)
            Set engagementCell = Sh.Range(ENGAGEMENT_START).Offset(i, 0)
            regionValue = Trim(regionCell.value)

            If Not IsHeaderCell(engagementCell) Then
                engagementCell.value = ""
                If HasValidation(engagementCell) Then
                    engagementCell.Validation.Delete
                End If

                If regionValue <> "" Then
                    ' Write dropdown data to USER SHEET (not DV)
                    Call ApplyEngagementValidation(Sh, engagementCell, teamValue, regionValue, i)
                Else
                    Call BlockCellInput(engagementCell)
                End If
            End If
        Next i
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "HandleTeamRegionChange Error: " & Err.Number & " - " & Err.Description
    Application.ScreenUpdating = True
    MsgBox "An error occurred while updating engagements: " & Err.Description, vbCritical, "Error"
    Application.ScreenUpdating = False
End Sub

' ============================================================================
' ENSURE ENGAGEMENT VALIDATION
' ============================================================================

Private Sub EnsureEngagementValidation(Sh As Worksheet, engagementCell As Range, _
                                       teamValue As String, regionValue As String, _
                                       rowOffset As Long)
    On Error GoTo ErrorHandler

    ' Always refresh to ensure correct data
    Call ApplyEngagementValidation(Sh, engagementCell, teamValue, regionValue, rowOffset)

    Exit Sub

ErrorHandler:
    Debug.Print "EnsureEngagementValidation Error: " & Err.Number & " - " & Err.Description
End Sub

' ============================================================================
' APPLY ENGAGEMENT VALIDATION - WRITES TO USER SHEET (CONFLICT-FREE)
' ============================================================================
Private Sub ApplyEngagementValidation(Sh As Worksheet, engagementCell As Range, _
                                      teamValue As String, regionValue As String, _
                                      rowOffset As Long)
    Dim wsDV As Worksheet
    Dim lastRowDV As Long
    Dim matchedValues As Collection
    Dim matchedValue As Variant
    Dim i As Long
    Dim rangeName As String
    Dim dropdownCol As Long
    Dim dropdownStartRow As Long
    Dim dropdownEndRow As Long
    Dim validationRange As Range
    Dim uniqueKey As String
    Dim valueCount As Long

    On Error GoTo ErrorHandler
    
    ' Prevent recursive events while writing to sheet
    Application.EnableEvents = False

    Set wsDV = ThisWorkbook.Worksheets("DV")
    Set matchedValues = New Collection
    lastRowDV = wsDV.Cells(wsDV.Rows.count, "N").End(xlUp).Row

    Debug.Print "=== ApplyEngagementValidation START ==="
    Debug.Print "Sheet: " & Sh.Name & " | Team: " & teamValue & " | Region: " & regionValue & " | Row Offset: " & rowOffset

    ' Collect matching engagement values from DV sheet (READ ONLY)
    For i = 2 To lastRowDV
        If Trim(wsDV.Cells(i, "N").value) = teamValue And _
           Trim(wsDV.Cells(i, "O").value) = regionValue Then
            uniqueKey = CStr(wsDV.Cells(i, "P").value)
            On Error Resume Next
            matchedValues.Add Trim(wsDV.Cells(i, "P").value), uniqueKey
            If Err.Number = 0 Then
                Debug.Print "Added: " & Trim(wsDV.Cells(i, "P").value)
            End If
            Err.Clear
            On Error GoTo ErrorHandler
        End If
    Next i

    Debug.Print "Total matched values: " & matchedValues.count

    ' Check if any engagements found
    If matchedValues.count = 0 Then
        Call BlockCellInput(engagementCell)
        Application.ScreenUpdating = True
        MsgBox "There are no engagements tagged to Team- " & teamValue & " in the Region " & regionValue & "." & vbCrLf & vbCrLf & _
               "Please reach out to PPG team and get your team added to the given engagement.", _
               vbExclamation, "No Engagements Found"
        Application.ScreenUpdating = False
        Exit Sub
    End If

    ' ===== WRITE DROPDOWN DATA TO USER SHEET (AFTER COLUMN BI) =====
    dropdownCol = DROPDOWN_START_COLUMN + rowOffset
    dropdownStartRow = DROPDOWN_START_ROW

    Debug.Print "Writing to Column: " & dropdownCol & " (" & ColumnLetter(dropdownCol) & ") | Starting Row: " & dropdownStartRow

    ' Clear previous dropdown data in this column (clear more rows to be safe)
    Sh.Range(Sh.Cells(dropdownStartRow, dropdownCol), _
             Sh.Cells(dropdownStartRow + 999, dropdownCol)).ClearContents

    ' Write engagement values to USER SHEET
    valueCount = 0
    For Each matchedValue In matchedValues
        Sh.Cells(dropdownStartRow + valueCount, dropdownCol).value = CStr(matchedValue)
        Debug.Print "Written to " & Sh.Cells(dropdownStartRow + valueCount, dropdownCol).Address & ": " & CStr(matchedValue)
        valueCount = valueCount + 1
    Next matchedValue

    dropdownEndRow = dropdownStartRow + valueCount - 1

    Debug.Print "Data written from " & Sh.Cells(dropdownStartRow, dropdownCol).Address & " to " & Sh.Cells(dropdownEndRow, dropdownCol).Address

    ' ===== CREATE SHEET-SCOPED NAMED RANGE =====
    rangeName = "Eng_R" & rowOffset

    ' Delete old named range if exists
    On Error Resume Next
    Sh.Names(rangeName).Delete
    Debug.Print "Deleted old named range: " & rangeName & " (if existed)"
    Err.Clear
    On Error GoTo ErrorHandler

    ' Define the validation range on USER SHEET
    Set validationRange = Sh.Range(Sh.Cells(dropdownStartRow, dropdownCol), _
                                    Sh.Cells(dropdownEndRow, dropdownCol))

    Debug.Print "Validation Range: " & validationRange.Address(True, True, xlA1, True)

    ' Create SHEET-SCOPED named range
    Sh.Names.Add Name:=rangeName, RefersTo:=validationRange
    Debug.Print "Created named range: " & Sh.Name & "!" & rangeName

    ' ===== APPLY DATA VALIDATION =====
    On Error Resume Next
    engagementCell.Validation.Delete
    Debug.Print "Deleted old validation on " & engagementCell.Address
    Err.Clear
    On Error GoTo ErrorHandler

    ' Apply validation using sheet-scoped range
    engagementCell.Validation.Add _
        Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, _
        Formula1:="=" & rangeName

    Debug.Print "Applied validation with formula: =" & rangeName

    With engagementCell.Validation
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
        .ErrorTitle = "Invalid Selection"
        .ErrorMessage = "Please select a valid Engagement ID from the list."
    End With
    Debug.Print "=== ApplyEngagementValidation SUCCESS ==="

    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    Debug.Print "!!! ERROR in ApplyEngagementValidation !!!"
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
    Debug.Print "Row Offset: " & rowOffset
    Call BlockCellInput(engagementCell)
    Application.EnableEvents = True
End Sub

' ============================================================================
' HELPER FUNCTION - GET COLUMN LETTER
' ============================================================================

Private Function ColumnLetter(colNum As Long) As String
    Dim colLetter As String
    Dim tempNum As Long

    tempNum = colNum
    Do While tempNum > 0
        colLetter = Chr(65 + ((tempNum - 1) Mod 26)) & colLetter
        tempNum = (tempNum - 1) \ 26
    Loop

    ColumnLetter = colLetter
End Function

' ============================================================================
' HELPER FUNCTIONS
' ============================================================================

Private Function IsHeaderCell(targetCell As Range) As Boolean
    Dim cellValue As String

    On Error Resume Next
    cellValue = Trim(UCase(targetCell.value))
    On Error GoTo 0

    IsHeaderCell = False

    If cellValue = "" Then Exit Function
    If InStr(cellValue, "ENGAGEMENT") > 0 Then IsHeaderCell = True
    If Left(cellValue, 2) = "5." Then IsHeaderCell = True
    If cellValue = "5. ENGAGEMENT" Or cellValue = "5.ENGAGEMENT" Then IsHeaderCell = True
End Function

Private Sub BlockCellInput(targetCell As Range)
    On Error Resume Next

    targetCell.Validation.Delete

    targetCell.Validation.Add _
        Type:=xlValidateCustom, _
        AlertStyle:=xlValidAlertStop, _
        Formula1:="=LEN(" & targetCell.Address(False, False) & ")=0"

    With targetCell.Validation
        .IgnoreBlank = True
        .ShowError = True
        .ErrorTitle = "Input Not Allowed"
        .ErrorMessage = "Please select a valid Team and Region first to enable this field."
    End With

    On Error GoTo 0
End Sub

Private Sub CleanupSheetNamedRanges(ws As Worksheet)
    Dim nm As Name

    On Error Resume Next
    For Each nm In ws.Names
        ' Delete all named ranges starting with "Eng_R" on this sheet
        If InStr(nm.Name, "Eng_R") > 0 Then
            Debug.Print "Cleaning up named range: " & nm.Name
            nm.Delete
        End If
    Next nm
    On Error GoTo 0
End Sub

Private Function HasValidation(rng As Range) As Boolean
    On Error Resume Next
    HasValidation = (rng.Validation.Type <> xlValidateInputOnly)
    On Error GoTo 0
End Function

Private Sub UnprotectSheet(ws As Worksheet)
    On Error Resume Next
    ws.Unprotect Password:=""
    On Error GoTo 0
End Sub

Private Sub ProtectSheet(ws As Worksheet)
    On Error Resume Next
    ws.Protect Password:="", UserInterfaceOnly:=True, AllowFiltering:=True
    On Error GoTo 0
End Sub




# user Sheet code
Option Explicit

Private Sub Worksheet_Activate()
    HandleActivate Me
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    HandleChange Me, Target
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    HandleSelectionChange Me, Target
End Sub

Private Sub Worksheet_Deactivate()
    HandleDeactivate Me
End Sub



# UserFormDE code

Option Explicit
Option Compare Text ' For case-insensitive comparisons

' Module-level variables
Const YEAR_CELL As String = "H1"
Const MAX_LOCAL_ROW As Long = 3200
Const MIN_LOCAL_ROW As Long = 3000

Private currentUserID As String
Private currentUseremail As String
Private isCorrection As Boolean
Private correctionSource As String
Private correctionRowNumber As Long
Private isUpdatingFilters As Boolean
Private isLoadingRecord As Boolean
Private correctionYear As Integer


' ==================== FORM INITIALIZATION ====================

Private Sub UserForm_Initialize()
    On Error GoTo InitializeError
    
    ' Initialize variables
isCorrection = False
correctionSource = ""
correctionRowNumber = 0
isUpdatingFilters = False
isLoadingRecord = False
correctionYear = 0

    
    ' Initialize form controls
    Call InitializeComboBoxes
    Call ClearForm(False)
    Call SetupListView
    
    ' Center the form
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2
    
    ' Set focus to ComboMonth
    Me.ComboMonth.SetFocus
    
    Exit Sub
    
InitializeError:
    MsgBox "Error initializing form: " & Err.Description & vbCrLf & _
           "Please contact support.", vbCritical, "Initialization Error"
    Unload Me
End Sub

' ==================== EMPLOYEE LOADING ====================

Public Sub LoadEmployeeName()
    On Error GoTo LoadEmployeeError
    
    Dim wsData As Worksheet
    Dim foundCell As Range
    Dim empID As String
    Dim empEmail As String
    
    ' Get user ID and email
    empID = Trim(ActiveSheet.Range("W6").value)
    empEmail = Trim(ActiveSheet.Range("W7").value)
    
    ' Validate user ID
    If empID = "" Then
        MsgBox "Enterprise ID not found in cell W6." & vbCrLf & _
               "Please ensure you're logged in properly.", vbExclamation, "Login Required"
        Unload Me
        Exit Sub
    End If
    
    currentUserID = empID
    currentUseremail = empEmail
    
    ' Search for employee name in DV sheet
    Set wsData = ThisWorkbook.Sheets("DV")
    
    If wsData Is Nothing Then
        MsgBox "DV sheet not found.", vbExclamation
        Me.TextboxEmpname.value = "Employee Name Not Found"
        Exit Sub
    End If
    
    Set foundCell = wsData.Range("F:F").Find(What:=currentUserID, _
                                             LookIn:=xlValues, _
                                             LookAt:=xlWhole, _
                                             MatchCase:=False)
    
    If Not foundCell Is Nothing Then
        Me.TextboxEmpname.value = Trim(wsData.Cells(foundCell.Row, "G").value)
    Else
        Me.TextboxEmpname.value = "Employee Name Not Found"
    End If
    
    ' Load all user records initially
    Call FilterRecords
    
    ' Set focus to ComboMonth
    Me.ComboMonth.SetFocus
    
    Exit Sub
    
LoadEmployeeError:
    MsgBox "Error loading employee information: " & Err.Description, vbCritical
    Me.TextboxEmpname.value = "Error Loading Name"
End Sub

' ==================== COMBOBOX INITIALIZATION ====================

Private Sub InitializeComboBoxes()
    On Error GoTo ComboInitError
    
    ' Initialize all comboboxes for Excel-like behavior
    With Me.ComboMonth
        .Style = fmStyleDropDownCombo ' Allow typing
        .MatchEntry = fmMatchEntryComplete ' Auto-complete
        .MatchRequired = False
    End With
    
    With Me.ComboRegion
        .Style = fmStyleDropDownCombo
        .MatchEntry = fmMatchEntryComplete
        .MatchRequired = False
    End With
    
    With Me.ComboTeam
        .Style = fmStyleDropDownCombo
        .MatchEntry = fmMatchEntryComplete
        .MatchRequired = False
    End With
    
    With Me.ComboEng
        .Style = fmStyleDropDownCombo
        .MatchEntry = fmMatchEntryComplete
        .MatchRequired = False
    End With
    
    With Me.ComboActi
        .Style = fmStyleDropDownCombo
        .MatchEntry = fmMatchEntryComplete
        .MatchRequired = False
    End With
    
    With Me.ComboNE
        .Style = fmStyleDropDownCombo
        .MatchEntry = fmMatchEntryComplete
        .MatchRequired = False
    End With
    
    ' Load month combobox
    Call LoadMonthComboBox
    
    ' Load independent dropdowns from DV sheet
    Call LoadTeamComboFromDV
    Call LoadRegionComboFromDV
    Call LoadActivityDropdown
    Call LoadNEDropdown
    
    ' Engagement dropdown will be loaded when both Team and Region are selected
    Me.ComboEng.Clear
    
    Exit Sub
    
ComboInitError:
    MsgBox "Error initializing dropdowns: " & Err.Description, vbExclamation
End Sub

' ==================== DV SHEET DROPDOWN FUNCTIONS ====================

Private Sub LoadTeamComboFromDV()
    ' Load unique Team values from DV sheet Column N
    On Error GoTo ErrorHandler
    
    Me.ComboTeam.Clear
    
    Dim wsDV As Worksheet
       On Error Resume Next
    Set wsDV = ThisWorkbook.Worksheets("DV")
    On Error GoTo ErrorHandler
    If wsDV Is Nothing Then
        MsgBox "DV worksheet not found.", vbExclamation
        Exit Sub
    End If
    
    If wsDV Is Nothing Then
        MsgBox "DV worksheet not found.", vbExclamation
        Exit Sub
    End If
    
    Dim lastRow As Long
    lastRow = wsDV.Cells(wsDV.Rows.count, "N").End(xlUp).Row
    
    If lastRow < 2 Then
        Exit Sub ' No data
    End If
    
    ' Collection to store unique values
    Dim uniqueItems As Collection
    Set uniqueItems = New Collection
    
    Dim i As Long
    For i = 2 To lastRow
        Dim teamValue As String
        teamValue = Trim(wsDV.Cells(i, "N").value)
        
        If teamValue <> "" Then
            ' Add only unique values using error handling
            On Error Resume Next
            uniqueItems.Add teamValue, CStr(teamValue)
            On Error GoTo 0
        End If
    Next i
    
    ' Add unique items to combobox
    Dim item As Variant
    For Each item In uniqueItems
        Me.ComboTeam.AddItem item
    Next item
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error loading Team dropdown: " & Err.Description, vbExclamation
End Sub

Private Sub LoadRegionComboFromDV()
    ' Load unique Region values from DV sheet Column O
    On Error GoTo ErrorHandler
    
    Me.ComboRegion.Clear
    
    Dim wsDV As Worksheet
    Set wsDV = ThisWorkbook.Worksheets("DV")
    
    If wsDV Is Nothing Then
        MsgBox "DV worksheet not found.", vbExclamation
        Exit Sub
    End If
    
    Dim lastRow As Long
    lastRow = wsDV.Cells(wsDV.Rows.count, "O").End(xlUp).Row
    
    If lastRow < 2 Then
        Exit Sub ' No data
    End If
    
    ' Collection to store unique values
    Dim uniqueItems As Collection
    Set uniqueItems = New Collection
    
    Dim i As Long
    For i = 2 To lastRow
        Dim regionValue As String
        regionValue = Trim(wsDV.Cells(i, "O").value)
        
        If regionValue <> "" Then
            ' Add only unique values using error handling
            On Error Resume Next
            uniqueItems.Add regionValue, CStr(regionValue)
            On Error GoTo 0
        End If
    Next i
    
    ' Add unique items to combobox
    Dim item As Variant
    For Each item In uniqueItems
        Me.ComboRegion.AddItem item
    Next item
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error loading Region dropdown: " & Err.Description, vbExclamation
End Sub

Private Sub LoadEngagementComboFromDV()
    ' Load Engagement values from column CI (user-specific engagement list)
    ' This prevents conflicts between multiple users
    On Error GoTo ErrorHandler

    Me.ComboEng.Clear

    ' Clear and update engagement list if both Team and Region are selected
    If Me.ComboTeam.value = "" Or Me.ComboRegion.value = "" Then
        Exit Sub
    End If

    ' Update the engagement list in column CI for this user
    Call UpdateEngagementListInColumnCI

    Exit Sub

ErrorHandler:
    MsgBox "Error loading Engagement dropdown: " & Err.Description, vbExclamation
    Me.ComboEng.Clear
End Sub


' ==================== OTHER DROPDOWN FUNCTIONS ====================

Private Sub LoadMonthComboBox()
    On Error GoTo LoadMonthError
    
    Me.ComboMonth.Clear
    
    With Me.ComboMonth
        .AddItem "January"
        .AddItem "February"
        .AddItem "March"
        .AddItem "April"
        .AddItem "May"
        .AddItem "June"
        .AddItem "July"
        .AddItem "August"
        .AddItem "September"
        .AddItem "October"
        .AddItem "November"
        .AddItem "December"
    End With
    
    Exit Sub
    
LoadMonthError:
    MsgBox "Error loading month dropdown: " & Err.Description, vbExclamation
End Sub

Private Sub LoadActivityDropdown()
    On Error GoTo LoadActivityError
    
    Me.ComboActi.Clear
    
    ' Still using named range for Activity
    Dim actRange As Range
    On Error Resume Next
    Set actRange = ThisWorkbook.Names("EngActivity").RefersToRange
    On Error GoTo 0
    
    If actRange Is Nothing Then
        ' Try to load from DV sheet as fallback
        LoadActivityFromDV
        Exit Sub
    End If
    
    Dim cell As Range
    For Each cell In actRange
        If Trim(cell.value) <> "" Then
            If Not ItemExistsInComboBox(Me.ComboActi, Trim(cell.value)) Then
                Me.ComboActi.AddItem Trim(cell.value)
            End If
        End If
    Next cell
    
    Exit Sub
    
LoadActivityError:
    MsgBox "Error loading activity dropdown: " & Err.Description, vbExclamation
End Sub

Private Sub LoadActivityFromDV()
    ' Fallback method to load Activity from DV sheet if named range doesn't exist
    On Error GoTo ErrorHandler
    
    Dim wsDV As Worksheet
        On Error Resume Next
    Set wsDV = ThisWorkbook.Worksheets("DV")
    On Error GoTo ErrorHandler
    If wsDV Is Nothing Then
        MsgBox "DV worksheet not found.", vbExclamation
        Exit Sub
    End If
    
    If wsDV Is Nothing Then Exit Sub
    
    Dim lastRow As Long
    lastRow = wsDV.Cells(wsDV.Rows.count, "Q").End(xlUp).Row ' Assuming Activity in Column Q
    
    If lastRow < 2 Then Exit Sub
    
    Dim uniqueItems As Collection
    Set uniqueItems = New Collection
    
    Dim i As Long
    For i = 2 To lastRow
        Dim activityValue As String
        activityValue = Trim(wsDV.Cells(i, "Q").value)
        
        If activityValue <> "" Then
            On Error Resume Next
            uniqueItems.Add activityValue, CStr(activityValue)
            On Error GoTo 0
        End If
    Next i
    
    Dim item As Variant
    For Each item In uniqueItems
        Me.ComboActi.AddItem item
    Next item
    
    Exit Sub
    
ErrorHandler:
    ' Silent fail - just don't load items
End Sub

Private Sub LoadNEDropdown()
    On Error GoTo LoadNEError
    
    Me.ComboNE.Clear
    
    ' Still using named range for Non-Engagement
    Dim neRange As Range
    On Error Resume Next
    Set neRange = ThisWorkbook.Names("Other_than_Eng").RefersToRange
    On Error GoTo 0
    
    If neRange Is Nothing Then
        ' Try to load from DV sheet as fallback
        LoadNEFromDV
        Exit Sub
    End If
    
    Dim cell As Range
    For Each cell In neRange
        If Trim(cell.value) <> "" Then
            If Not ItemExistsInComboBox(Me.ComboNE, Trim(cell.value)) Then
                Me.ComboNE.AddItem Trim(cell.value)
            End If
        End If
    Next cell
    
    Exit Sub
    
LoadNEError:
    MsgBox "Error loading non-engagement dropdown: " & Err.Description, vbExclamation
End Sub

Private Sub LoadNEFromDV()
    ' Fallback method to load Non-Engagement from DV sheet
    On Error GoTo ErrorHandler
    
    Dim wsDV As Worksheet
      On Error Resume Next
    Set wsDV = ThisWorkbook.Worksheets("DV")
    On Error GoTo ErrorHandler
    If wsDV Is Nothing Then
        MsgBox "DV worksheet not found.", vbExclamation
        Exit Sub
    End If
    
    If wsDV Is Nothing Then Exit Sub
    
    Dim lastRow As Long
    lastRow = wsDV.Cells(wsDV.Rows.count, "R").End(xlUp).Row ' Assuming Non-Engagement in Column R
    
    If lastRow < 2 Then Exit Sub
    
    Dim uniqueItems As Collection
    Set uniqueItems = New Collection
    
    Dim i As Long
    For i = 2 To lastRow
        Dim neValue As String
        neValue = Trim(wsDV.Cells(i, "R").value)
        
        If neValue <> "" Then
            On Error Resume Next
            uniqueItems.Add neValue, CStr(neValue)
            On Error GoTo 0
        End If
    Next i
    
    Dim item As Variant
    For Each item In uniqueItems
        Me.ComboNE.AddItem item
    Next item
    
    Exit Sub
    
ErrorHandler:
    ' Silent fail - just don't load items
End Sub

' ==================== HELPER FUNCTIONS ====================

Private Function ItemExistsInComboBox(cmb As MSForms.comboBox, itemText As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim i As Long
    For i = 0 To cmb.ListCount - 1
        If StrComp(cmb.List(i), itemText, vbTextCompare) = 0 Then
            ItemExistsInComboBox = True
            Exit Function
        End If
    Next i
    
    ItemExistsInComboBox = False
    Exit Function
    
ErrorHandler:
    ItemExistsInComboBox = False
End Function

Private Function IsValidMonth(monthName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim monthNum As Integer
    Dim testDate As Date
    
    ' Try to parse the month name
    monthName = Trim(monthName)
    If monthName = "" Then
        IsValidMonth = False
        Exit Function
    End If
    
    ' Check if it's in the list first (case-insensitive)
    Dim i As Long
    For i = 0 To Me.ComboMonth.ListCount - 1
        If StrComp(Me.ComboMonth.List(i), monthName, vbTextCompare) = 0 Then
            IsValidMonth = True
            Exit Function
        End If
    Next i
    
    ' Try to convert to date (accepts "Jan", "January", etc.)
    On Error Resume Next
    testDate = DateValue("1 " & monthName & " 2000")
    monthNum = Month(testDate)
    On Error GoTo ErrorHandler
    
    If monthNum >= 1 And monthNum <= 12 Then
        IsValidMonth = True
    Else
        IsValidMonth = False
    End If
    
    Exit Function
    
ErrorHandler:
    IsValidMonth = False
End Function

Private Function ValueExistsInComboBox(cmb As MSForms.comboBox, valueText As String) As Boolean
    ' Validates if a value exists in the combobox dropdown list
    ' Returns True if found, False if not found or combobox is empty
    On Error GoTo ErrorHandler

    ValueExistsInComboBox = False

    ' Check if combobox has items
    If cmb.ListCount = 0 Then
        Exit Function
    End If

    ' Check if value is empty
    If Trim(valueText) = "" Then
        Exit Function
    End If

    ' Search for exact match (case-insensitive)
    Dim i As Long
    For i = 0 To cmb.ListCount - 1
        If StrComp(Trim(cmb.List(i)), Trim(valueText), vbTextCompare) = 0 Then
            ValueExistsInComboBox = True
            Exit Function
        End If
    Next i

    Exit Function

ErrorHandler:
    ValueExistsInComboBox = False
End Function

' ==================== COMBOBOX EVENT HANDLERS ====================

Private Sub ComboMonth_Change()
    On Error GoTo ErrorHandler
    
    If Not isUpdatingFilters And Not isLoadingRecord Then
        ' Validate month before proceeding
        If Me.ComboMonth.value <> "" Then
            If Not IsValidMonth(Me.ComboMonth.value) Then
                MsgBox "Invalid month name. Please select a valid month from the list.", vbExclamation, "Invalid Month"
                Me.ComboMonth.value = ""
                Me.ComboMonth.SetFocus
                Exit Sub
            End If
        End If
        
        Call UpdateWeekDropdown
        Call FilterRecords
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in month selection: " & Err.Description, vbExclamation
End Sub

Private Sub ComboTeam_Change()
    On Error GoTo ErrorHandler
    
    If Not isUpdatingFilters And Not isLoadingRecord Then
        ' Clear and reload engagement dropdown based on new team selection
        Call LoadEngagementComboFromDV
        Call FilterRecords
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in team selection: " & Err.Description, vbExclamation
End Sub

Private Sub ComboRegion_Change()
    On Error GoTo ErrorHandler
    
    If Not isUpdatingFilters And Not isLoadingRecord Then
        ' Clear and reload engagement dropdown based on new region selection
        Call LoadEngagementComboFromDV
        Call FilterRecords
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in region selection: " & Err.Description, vbExclamation
End Sub

Private Sub ComboEng_Change()
    If Not isUpdatingFilters And Not isLoadingRecord Then Call FilterRecords
End Sub

Private Sub ComboActi_Change()
    If Not isUpdatingFilters And Not isLoadingRecord Then Call FilterRecords
End Sub

Private Sub ComboNE_Change()
    If Not isUpdatingFilters And Not isLoadingRecord Then Call FilterRecords
End Sub

Private Sub ComboWeek_Change()
    If Not isUpdatingFilters And Not isLoadingRecord Then Call FilterRecords
End Sub

Private Sub TextboxEmpname_Change()
    If Not isUpdatingFilters And Not isLoadingRecord Then Call FilterRecords
End Sub

' ==================== WEEK DROPDOWN UPDATES ====================

Private Sub UpdateWeekDropdown()
    On Error GoTo WeekUpdateError
    
    If Me.ComboMonth.value = "" Then
        Me.ComboWeek.Clear
        Exit Sub
    End If
    
    isUpdatingFilters = True ' Prevent filtering during week update
    
    Dim currentWeek As String
    currentWeek = Me.ComboWeek.value ' Store current selection
    
    Me.ComboWeek.Clear
    
    Dim selectedMonth As String
    Dim selectedYear As Integer
    Dim firstDayOfMonth As Date
    Dim lastDayOfMonth As Date
    Dim currentDate As Date
    Dim weekString As String
    Dim monthNum As Integer
    
    selectedMonth = Trim(Me.ComboMonth.value)
    

' Get year with improved logic
On Error Resume Next

' Priority 1: If correcting a record, use the correction year
If isCorrection And correctionYear > 0 Then
    selectedYear = correctionYear
    Debug.Print "CORRECTION MODE: Using correction year: " & selectedYear & " from correctionYear variable"
ElseIf Not isCorrection And correctionYear > 0 Then
    ' We have a year from loading but not in correction mode - still use it
    selectedYear = correctionYear
    Debug.Print "LOADING MODE: Using year: " & selectedYear & " from correctionYear variable"

' Priority 2: Use year from cell H1
ElseIf ActiveSheet.Range(YEAR_CELL).value <> "" Then
    selectedYear = CInt(ActiveSheet.Range(YEAR_CELL).value)
    If Err.Number <> 0 Or selectedYear < 2000 Or selectedYear > 2100 Then
        selectedYear = Year(Date)
    End If
' Priority 3: Use current year as fallback
Else
    selectedYear = Year(Date)
End If

On Error GoTo WeekUpdateError


    
    ' Validate and get month number
    If Not IsValidMonth(selectedMonth) Then
        MsgBox "Invalid month: " & selectedMonth, vbExclamation
        Me.ComboMonth.value = ""
        isUpdatingFilters = False
        Exit Sub
    End If
    
    ' Convert month name to number safely
    On Error Resume Next
    monthNum = Month(DateValue("1 " & selectedMonth & " 2000"))
    If Err.Number <> 0 Then
        MsgBox "Invalid month name. Please select from the list.", vbExclamation
        Me.ComboMonth.value = ""
        isUpdatingFilters = False
        Exit Sub
    End If
    On Error GoTo WeekUpdateError
    
    ' Calculate dates
    firstDayOfMonth = DateSerial(selectedYear, monthNum, 1)
    lastDayOfMonth = DateSerial(selectedYear, monthNum + 1, 0)
    
    currentDate = firstDayOfMonth
    Do While currentDate <= lastDayOfMonth
        If Weekday(currentDate, vbMonday) = 1 Then
            weekString = "Week starting " & Format(currentDate, "mmmm d yyyy")
            Me.ComboWeek.AddItem weekString
        End If
        currentDate = DateAdd("d", 1, currentDate)
    Loop
    
    ' Restore week selection if it still exists
    If currentWeek <> "" Then
        Dim i As Integer
        For i = 0 To Me.ComboWeek.ListCount - 1
            If Me.ComboWeek.List(i) = currentWeek Then
                Me.ComboWeek.ListIndex = i
                Exit For
            End If
        Next i
    End If
    
    isUpdatingFilters = False
    Exit Sub
    
WeekUpdateError:
    MsgBox "Error updating week dropdown: " & Err.Description & vbCrLf & _
           "Please check the month name and try again.", vbExclamation
    Me.ComboWeek.Clear
    isUpdatingFilters = False
End Sub
' ==================== ENGAGEMENT LIST IN COLUMN CI ====================

Private Sub UpdateEngagementListInColumnCI()
    ' Updates column CI with filtered engagement list based on Team and Region
    ' This ensures each user has their own engagement list without conflicts
    On Error GoTo ErrorHandler

    Dim wsUser As Worksheet
    Set wsUser = ActiveSheet

    ' Clear existing engagement list in column CI
    Dim lastRowCI As Long
    lastRowCI = wsUser.Cells(wsUser.Rows.count, "CI").End(xlUp).Row
    If lastRowCI >= 2 Then
        wsUser.Range("CI2:CI" & lastRowCI).ClearContents
    End If

    ' Exit if either team or region is empty
    If Me.ComboTeam.value = "" Or Me.ComboRegion.value = "" Then
        Me.ComboEng.Clear
        Exit Sub
    End If

    Dim wsDV As Worksheet
    On Error Resume Next
    Set wsDV = ThisWorkbook.Worksheets("DV")
    On Error GoTo ErrorHandler

    If wsDV Is Nothing Then
        MsgBox "DV worksheet not found.", vbExclamation
        Exit Sub
    End If

    Dim lastRowDV As Long
    lastRowDV = wsDV.Cells(wsDV.Rows.count, "N").End(xlUp).Row

    If lastRowDV < 2 Then
        Me.ComboEng.Clear
        Exit Sub
    End If

    Dim selectedTeam As String
    Dim selectedRegion As String
    selectedTeam = Trim(Me.ComboTeam.value)
    selectedRegion = Trim(Me.ComboRegion.value)

    ' Collection to store unique engagement values
    Dim uniqueEngagements As Collection
    Set uniqueEngagements = New Collection

    Dim i As Long
    For i = 2 To lastRowDV
        Dim currentTeam As String
        Dim currentRegion As String
        Dim engagementValue As String

        currentTeam = Trim(wsDV.Cells(i, "N").value)
        currentRegion = Trim(wsDV.Cells(i, "O").value)
        engagementValue = Trim(wsDV.Cells(i, "P").value)

        ' Match both Team AND Region (case-insensitive)
        If StrComp(currentTeam, selectedTeam, vbTextCompare) = 0 And _
           StrComp(currentRegion, selectedRegion, vbTextCompare) = 0 And _
           engagementValue <> "" Then

            ' Add only unique values
            On Error Resume Next
            uniqueEngagements.Add engagementValue, CStr(engagementValue)
            On Error GoTo 0
        End If
    Next i

    ' Write unique engagements to column CI starting from row 2
    Dim rowNum As Long
    rowNum = 2

    Dim item As Variant
    For Each item In uniqueEngagements
        wsUser.Cells(rowNum, "CI").value = item
        rowNum = rowNum + 1
    Next item

    ' Add header if not present
    If wsUser.Cells(1, "CI").value = "" Then
        wsUser.Cells(1, "CI").value = "Engagement List"
    End If

    ' Now populate the ComboEng from column CI
    Me.ComboEng.Clear

    For i = 2 To rowNum - 1
        If wsUser.Cells(i, "CI").value <> "" Then
            Me.ComboEng.AddItem wsUser.Cells(i, "CI").value
        End If
    Next i

    Exit Sub

ErrorHandler:
    MsgBox "Error updating engagement list: " & Err.Description, vbExclamation
    Me.ComboEng.Clear
End Sub


' ==================== LISTVIEW SETUP ====================

Private Sub SetupListView()
    On Error GoTo ListViewError
    
    With Me.ListViewRecord
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .HideSelection = False
        .LabelWrap = False
        .MultiSelect = False
        
        ' Clear existing headers
        .ColumnHeaders.Clear
        
        ' Add column headers (18 columns total)
        .ColumnHeaders.Add , , "Unique Code", 80
        .ColumnHeaders.Add , , "Enterprise ID", 100
        .ColumnHeaders.Add , , "Employee Name", 150
        .ColumnHeaders.Add , , "Employee Email", 120
        .ColumnHeaders.Add , , "Date Time", 140
        .ColumnHeaders.Add , , "Type", 120
        .ColumnHeaders.Add , , "Month", 80
        .ColumnHeaders.Add , , "Week", 180
        .ColumnHeaders.Add , , "Team Name", 120
        .ColumnHeaders.Add , , "Region", 80
        .ColumnHeaders.Add , , "Audit Engagement", 120
        .ColumnHeaders.Add , , "Engagement Activity", 130
        .ColumnHeaders.Add , , "Engagement Hr", 110
        .ColumnHeaders.Add , , "Remark 1", 100
        .ColumnHeaders.Add , , "Non-Audit Eng", 120
        .ColumnHeaders.Add , , "Non-Audit Hr", 90
        .ColumnHeaders.Add , , "Remark 2", 100
        .ColumnHeaders.Add , , "Submitted By", 120
    End With
    
    Exit Sub
    
ListViewError:
    MsgBox "Error setting up list view: " & Err.Description, vbCritical
End Sub

' ==================== RECORD SELECTION ====================

Private Sub ListViewRecord_DblClick()
    On Error GoTo ErrorHandler
    
    ' Prevent any filtering during record loading
    isLoadingRecord = True
    isUpdatingFilters = True
    
    If Me.ListViewRecord.selectedItem Is Nothing Then
        isLoadingRecord = False
        isUpdatingFilters = False
        Exit Sub
    End If
    
    ' Store the selected item data immediately to avoid reference issues
    Dim selectedKey As String
    Dim subItemsArray() As String
    Dim i As Integer
    
    selectedKey = Me.ListViewRecord.selectedItem.Key
    
    ' Store all subitems
    ReDim subItemsArray(0 To Me.ListViewRecord.selectedItem.ListSubItems.count)
    subItemsArray(0) = Me.ListViewRecord.selectedItem.text
    
    For i = 1 To Me.ListViewRecord.selectedItem.ListSubItems.count
        subItemsArray(i) = Me.ListViewRecord.selectedItem.SubItems(i)
    Next i
    
    ' Extract source and row number from key
    Dim keyParts As Variant
    keyParts = Split(selectedKey, "_")
    
    If UBound(keyParts) < 1 Then
        MsgBox "Invalid record key.", vbExclamation
        isLoadingRecord = False
        isUpdatingFilters = False
        Exit Sub
    End If
    
    correctionSource = keyParts(0)
    correctionRowNumber = CLng(keyParts(1))
    
    ' Load record for correction using stored data
    If LoadRecordFromStoredData(subItemsArray) Then
        isCorrection = True
        Me.Caption = "Timesheet Data Entry - CORRECTION MODE (" & correctionSource & ")"
        MsgBox "Record loaded for correction. Modify fields as needed and click 'Correct Record'.", vbInformation
        Me.ComboMonth.SetFocus
    Else
        MsgBox "Error loading record for correction.", vbExclamation
    End If
    
    isLoadingRecord = False
    isUpdatingFilters = False
    Exit Sub
    
ErrorHandler:
    MsgBox "Error selecting record: " & Err.Description, vbCritical
    isLoadingRecord = False
    isUpdatingFilters = False
End Sub

Private Function LoadRecordFromStoredData(subItemsArray() As String) As Boolean
    On Error GoTo ErrorHandler
    
    LoadRecordFromStoredData = False
    
    ' Clear form first
    Call ClearForm(True)
    
    ' Load values from stored array with bounds checking
    Dim arraySize As Integer
    arraySize = UBound(subItemsArray)
    
   If arraySize >= 2 Then Me.TextboxEmpname.value = subItemsArray(2)  ' Employee Name

' Load Team and Region FIRST (needed for engagement dropdown)
If arraySize >= 8 Then Me.ComboTeam.value = subItemsArray(8)       ' Team Name
If arraySize >= 9 Then Me.ComboRegion.value = subItemsArray(9)     ' Region

' Load month and extract year from week
If arraySize >= 6 Then Me.ComboMonth.value = subItemsArray(6)      ' Month

' Extract year from week string BEFORE populating week dropdown
If arraySize >= 7 Then
    Dim weekText As String
    weekText = subItemsArray(7)

    ' Extract year from week string
    If weekText <> "" And Len(weekText) > 4 Then
        On Error Resume Next
        ' Extract the year (last 4 digits of the string)
        Dim yearPart As String
        yearPart = Right(Trim(weekText), 4)
        correctionYear = CInt(yearPart)

        ' Validate year
        If Err.Number <> 0 Or correctionYear < 2000 Or correctionYear > 2100 Then
            correctionYear = Year(Date)
        End If
        Err.Clear
        On Error GoTo ErrorHandler

        Debug.Print "Extracted correction year: " & correctionYear & " from week: " & weekText
    Else
        correctionYear = Year(Date)
    End If

    ' Update week dropdown with the correct year
    Call UpdateWeekDropdown

    ' NOW set the ComboWeek value AFTER dropdown is populated with correct year
    Me.ComboWeek.value = weekText
End If

    If arraySize >= 9 Then Me.ComboRegion.value = subItemsArray(9)     ' Region
    If arraySize >= 10 Then Me.ComboEng.value = subItemsArray(10)      ' Audit Engagement
    If arraySize >= 11 Then Me.ComboActi.value = subItemsArray(11)     ' Engagement Activity
    If arraySize >= 12 Then Me.TextBoxHour1.value = subItemsArray(12)  ' Engagement Hr
    If arraySize >= 13 Then Me.TextBoxR1.value = subItemsArray(13)     ' Remark 1
    If arraySize >= 14 Then Me.ComboNE.value = subItemsArray(14)       ' Non-Audit Eng
    If arraySize >= 15 Then Me.TextBoxH2.value = subItemsArray(15)     ' Non-Audit Hr
    If arraySize >= 16 Then Me.TextBoxR2.value = subItemsArray(16)     ' Remark 2
    
  
    
    ' Load engagement dropdown after setting team and region
    If arraySize >= 8 And Me.ComboTeam.value <> "" And arraySize >= 9 And Me.ComboRegion.value <> "" Then
        Call LoadEngagementComboFromDV
        ' Re-select the engagement
        If arraySize >= 10 Then Me.ComboEng.value = subItemsArray(10)
    End If
    
    LoadRecordFromStoredData = True
    Exit Function
    
ErrorHandler:
    MsgBox "Error loading record data: " & Err.Description, vbCritical
    LoadRecordFromStoredData = False
End Function

' ==================== FILTERING FUNCTIONS ====================

Private Sub Commandfilter_Click()
    Call FilterRecords
End Sub

Private Sub Commclearfilter_Click()
    On Error GoTo ErrorHandler
    
    ' Prevent filtering during clear operation
    isUpdatingFilters = True
    
    ' Clear all filter controls except employee name (keep it as it's loaded from login)
    Dim employeeName As String
    employeeName = Me.TextboxEmpname.value ' Preserve employee name
    
    Me.ComboMonth.value = ""
    Me.ComboWeek.value = ""
    Me.ComboRegion.value = ""
    Me.ComboTeam.value = ""
    Me.ComboEng.value = ""
    Me.ComboActi.value = ""
    Me.ComboNE.value = ""
    
    ' Restore employee name
    Me.TextboxEmpname.value = employeeName
    
    ' Clear engagement dropdown since team or region might be cleared
    Me.ComboEng.Clear
    
    ' Enable filtering again
    isUpdatingFilters = False
    
    ' Load all records (no filters applied)
    Call FilterRecords
    
    MsgBox "All filters cleared. Showing all records.", vbInformation, "Filters Cleared"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error clearing filters: " & Err.Description, vbCritical
    isUpdatingFilters = False
End Sub

Private Sub ClearALL()
    On Error GoTo ErrorHandler
    
    ' Prevent filtering during clear operation
    isUpdatingFilters = True
    
    ' Clear all filter controls except employee name (keep it as it's loaded from login)
    Dim employeeName As String
    employeeName = Me.TextboxEmpname.value ' Preserve employee name
    
    Me.ComboMonth.value = ""
    Me.ComboWeek.value = ""
    Me.ComboRegion.value = ""
    Me.ComboTeam.value = ""
    Me.ComboEng.value = ""
    Me.ComboActi.value = ""
    Me.ComboNE.value = ""
    
    ' Restore employee name
    Me.TextboxEmpname.value = employeeName
    
    ' Clear engagement dropdown since team or region might be cleared
    Me.ComboEng.Clear
    
    ' Enable filtering again
    isUpdatingFilters = False
    
    ' Load all records (no filters applied)
    Call FilterRecords
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error clearing filters: " & Err.Description, vbCritical
    isUpdatingFilters = False
End Sub

Private Sub FilterRecords()
    On Error GoTo FilterError
    
    If currentUserID = "" Then
        MsgBox "Enterprise ID not available.", vbExclamation
        Exit Sub
    End If
    
    ' Don't filter if we're loading a record
    If isLoadingRecord Then Exit Sub
    
    ' Clear ListView
    Me.ListViewRecord.ListItems.Clear
    
    Dim recordCount As Long
    Dim totalAuditHours As Double
    Dim totalNonAuditHours As Double
    Dim dbAuditHours As Double
    Dim dbNonAuditHours As Double
    Dim localAuditHours As Double
    Dim localNonAuditHours As Double
    
    recordCount = 0
    totalAuditHours = 0
    totalNonAuditHours = 0
    
    ' Check if any filter is applied
    Dim hasFilters As Boolean
    hasFilters = (Me.ComboMonth.value <> "" Or _
                  Me.ComboWeek.value <> "" Or _
                  Me.ComboRegion.value <> "" Or _
                  Me.ComboTeam.value <> "" Or _
                  Me.ComboEng.value <> "" Or _
                  Me.ComboActi.value <> "" Or _
                  Me.ComboNE.value <> "")
    
    If hasFilters Then
        ' Filter records from Database
        recordCount = recordCount + FilterRecordsFromDatabase(dbAuditHours, dbNonAuditHours)
        totalAuditHours = totalAuditHours + dbAuditHours
        totalNonAuditHours = totalNonAuditHours + dbNonAuditHours
        
        ' Filter records from Local storage
        recordCount = recordCount + FilterRecordsFromLocal(localAuditHours, localNonAuditHours)
        totalAuditHours = totalAuditHours + localAuditHours
        totalNonAuditHours = totalNonAuditHours + localNonAuditHours
    Else
        ' Load all records if no filters
        recordCount = recordCount + LoadAllRecordsFromDatabase(dbAuditHours, dbNonAuditHours)
        totalAuditHours = totalAuditHours + dbAuditHours
        totalNonAuditHours = totalNonAuditHours + dbNonAuditHours
        
        recordCount = recordCount + LoadAllRecordsFromLocal(localAuditHours, localNonAuditHours)
        totalAuditHours = totalAuditHours + localAuditHours
        totalNonAuditHours = totalNonAuditHours + localNonAuditHours
    End If
    
    ' Update caption with count and hours
    If recordCount > 0 Then
        Me.Caption = "Timesheet Data Entry - " & recordCount & " Record(s) | " & _
                     "Audit Hours: " & Format(totalAuditHours, "0.00") & " | " & _
                     "Non-Audit Hours: " & Format(totalNonAuditHours, "0.00")
    Else
        Me.Caption = "Timesheet Data Entry - No Records Found"
    End If
    
    Exit Sub
    
FilterError:
    MsgBox "Error filtering records: " & Err.Description, vbCritical
End Sub

Private Function LoadAllRecordsFromDatabase(ByRef auditHoursSum As Double, ByRef nonAuditHoursSum As Double) As Long
    On Error GoTo ErrorHandler
    
    Dim wsDB As Worksheet
    Set wsDB = ThisWorkbook.Sheets("Data base")
    
    Dim lastRow As Long
    Dim r As Long
    Dim recordCount As Long
    Dim col As Integer
    Dim auditHours As Double
    Dim nonAuditHours As Double
    
    recordCount = 0
    auditHoursSum = 0
    nonAuditHoursSum = 0
    lastRow = wsDB.Cells(wsDB.Rows.count, 1).End(xlUp).Row
    
    For r = 2 To lastRow
        If wsDB.Cells(r, 2).value = currentUserID And wsDB.Cells(r, 20).value <> "Marked for Deletion" Then
            Dim listItem As listItem
            Set listItem = Me.ListViewRecord.ListItems.Add(, , wsDB.Cells(r, 1).value)
            listItem.Key = "DB_" & r
            
            ' Add subitems for remaining 17 columns
            For col = 2 To 18
                listItem.SubItems(col - 1) = CStr(wsDB.Cells(r, col).value)
            Next col
            
            ' Sum hours (Column 13 = Audit Hours, Column 16 = Non-Audit Hours)
            auditHours = 0
            nonAuditHours = 0
            
            On Error Resume Next
            auditHours = CDbl(wsDB.Cells(r, 13).value)
            nonAuditHours = CDbl(wsDB.Cells(r, 16).value)
            On Error GoTo 0
            
            auditHoursSum = auditHoursSum + auditHours
            nonAuditHoursSum = nonAuditHoursSum + nonAuditHours
            
            recordCount = recordCount + 1
        End If
    Next r
    
    LoadAllRecordsFromDatabase = recordCount
    Exit Function
    
ErrorHandler:
    MsgBox "Error loading database records: " & Err.Description, vbExclamation
    LoadAllRecordsFromDatabase = 0
End Function

Private Function LoadAllRecordsFromLocal(ByRef auditHoursSum As Double, ByRef nonAuditHoursSum As Double) As Long
    On Error GoTo ErrorHandler
    
    Dim wsUser As Worksheet
    Set wsUser = ActiveSheet
    
    Dim r As Long
    Dim recordCount As Long
    Dim col As Integer
    Dim auditHours As Double
    Dim nonAuditHours As Double
    
    recordCount = 0
    auditHoursSum = 0
    nonAuditHoursSum = 0
    
    ' Check within fixed range (3000-3200)
    For r = MIN_LOCAL_ROW To MAX_LOCAL_ROW
        If wsUser.Cells(r, 3).value = currentUserID And wsUser.Cells(r, 20).value <> "Marked for Deletion" Then
            Dim listItem As listItem
            Set listItem = Me.ListViewRecord.ListItems.Add(, , wsUser.Cells(r, 2).value)
            listItem.Key = "LOCAL_" & r
            
            ' Add subitems for columns 3 to 19 (17 columns total)
            For col = 3 To 19
                listItem.SubItems(col - 2) = CStr(wsUser.Cells(r, col).value)
            Next col
            
            ' Sum hours (Column 14 = Audit Hours, Column 17 = Non-Audit Hours in local)
            auditHours = 0
            nonAuditHours = 0
            
            On Error Resume Next
            auditHours = CDbl(wsUser.Cells(r, 14).value)
            nonAuditHours = CDbl(wsUser.Cells(r, 17).value)
            On Error GoTo 0
            
            auditHoursSum = auditHoursSum + auditHours
            nonAuditHoursSum = nonAuditHoursSum + nonAuditHours
            
            recordCount = recordCount + 1
        End If
    Next r
    
    LoadAllRecordsFromLocal = recordCount
    Exit Function
    
ErrorHandler:
    MsgBox "Error loading local records: " & Err.Description, vbExclamation
    LoadAllRecordsFromLocal = 0
End Function

Private Function FilterRecordsFromDatabase(ByRef auditHoursSum As Double, ByRef nonAuditHoursSum As Double) As Long
    On Error GoTo ErrorHandler
    
    Dim wsDB As Worksheet
    Set wsDB = ThisWorkbook.Sheets("Data base")
    
    Dim lastRow As Long
    Dim r As Long
    Dim recordCount As Long
    Dim includeRecord As Boolean
    Dim col As Integer
    Dim auditHours As Double
    Dim nonAuditHours As Double
    
    recordCount = 0
    auditHoursSum = 0
    nonAuditHoursSum = 0
    lastRow = wsDB.Cells(wsDB.Rows.count, 1).End(xlUp).Row
    
    For r = 2 To lastRow
        If wsDB.Cells(r, 2).value = currentUserID And wsDB.Cells(r, 20).value <> "Marked for Deletion" Then
            includeRecord = True
            
            ' Apply filters based on form selections (case-insensitive)
            If Me.ComboMonth.value <> "" And UCase(wsDB.Cells(r, 7).value) <> UCase(Me.ComboMonth.value) Then
                includeRecord = False
            End If
            If Me.ComboWeek.value <> "" And UCase(wsDB.Cells(r, 8).value) <> UCase(Me.ComboWeek.value) Then
                includeRecord = False
            End If
            If Me.ComboRegion.value <> "" And UCase(wsDB.Cells(r, 10).value) <> UCase(Me.ComboRegion.value) Then
                includeRecord = False
            End If
            If Me.ComboTeam.value <> "" And UCase(wsDB.Cells(r, 9).value) <> UCase(Me.ComboTeam.value) Then
                includeRecord = False
            End If
            If Me.ComboEng.value <> "" And UCase(wsDB.Cells(r, 11).value) <> UCase(Me.ComboEng.value) Then
                includeRecord = False
            End If
            If Me.ComboActi.value <> "" And UCase(wsDB.Cells(r, 12).value) <> UCase(Me.ComboActi.value) Then
                includeRecord = False
            End If
            If Me.ComboNE.value <> "" And UCase(wsDB.Cells(r, 15).value) <> UCase(Me.ComboNE.value) Then
                includeRecord = False
            End If
            
            If includeRecord Then
                Dim listItem As listItem
                Set listItem = Me.ListViewRecord.ListItems.Add(, , wsDB.Cells(r, 1).value)
                listItem.Key = "DB_" & r
                
                For col = 2 To 18
                    listItem.SubItems(col - 1) = CStr(wsDB.Cells(r, col).value)
                Next col
                
                ' Sum hours (Column 13 = Audit Hours, Column 16 = Non-Audit Hours)
                auditHours = 0
                nonAuditHours = 0
                
                On Error Resume Next
                auditHours = CDbl(wsDB.Cells(r, 13).value)
                nonAuditHours = CDbl(wsDB.Cells(r, 16).value)
                On Error GoTo 0
                
                auditHoursSum = auditHoursSum + auditHours
                nonAuditHoursSum = nonAuditHoursSum + nonAuditHours
                
                recordCount = recordCount + 1
            End If
        End If
    Next r
    
    FilterRecordsFromDatabase = recordCount
    Exit Function
    
ErrorHandler:
    MsgBox "Error filtering database records: " & Err.Description, vbExclamation
    FilterRecordsFromDatabase = 0
End Function

Private Function FilterRecordsFromLocal(ByRef auditHoursSum As Double, ByRef nonAuditHoursSum As Double) As Long
    On Error GoTo ErrorHandler
    
    Dim wsUser As Worksheet
    Set wsUser = ActiveSheet
    
    Dim r As Long
    Dim recordCount As Long
    Dim includeRecord As Boolean
    Dim col As Integer
    Dim auditHours As Double
    Dim nonAuditHours As Double
    
    recordCount = 0
    auditHoursSum = 0
    nonAuditHoursSum = 0
    
    ' Check within fixed range (3000-3200)
    For r = MIN_LOCAL_ROW To MAX_LOCAL_ROW
        If wsUser.Cells(r, 3).value = currentUserID And wsUser.Cells(r, 20).value <> "Marked for Deletion" Then
            includeRecord = True
            
            ' Apply filters based on form selections (case-insensitive)
            If Me.ComboMonth.value <> "" And UCase(wsUser.Cells(r, 8).value) <> UCase(Me.ComboMonth.value) Then
                includeRecord = False
            End If
            If Me.ComboWeek.value <> "" And UCase(wsUser.Cells(r, 9).value) <> UCase(Me.ComboWeek.value) Then
                includeRecord = False
            End If
            If Me.ComboRegion.value <> "" And UCase(wsUser.Cells(r, 11).value) <> UCase(Me.ComboRegion.value) Then
                includeRecord = False
            End If
            If Me.ComboTeam.value <> "" And UCase(wsUser.Cells(r, 10).value) <> UCase(Me.ComboTeam.value) Then
                includeRecord = False
            End If
            If Me.ComboEng.value <> "" And UCase(wsUser.Cells(r, 12).value) <> UCase(Me.ComboEng.value) Then
                includeRecord = False
            End If
            If Me.ComboActi.value <> "" And UCase(wsUser.Cells(r, 13).value) <> UCase(Me.ComboActi.value) Then
                includeRecord = False
            End If
            If Me.ComboNE.value <> "" And UCase(wsUser.Cells(r, 16).value) <> UCase(Me.ComboNE.value) Then
                includeRecord = False
            End If
            
            If includeRecord Then
                Dim listItem As listItem
                Set listItem = Me.ListViewRecord.ListItems.Add(, , wsUser.Cells(r, 2).value)
                listItem.Key = "LOCAL_" & r
                
                For col = 3 To 19
                    listItem.SubItems(col - 2) = CStr(wsUser.Cells(r, col).value)
                Next col
                
                ' Sum hours (Column 14 = Audit Hours, Column 17 = Non-Audit Hours in local)
                auditHours = 0
                nonAuditHours = 0
                
                On Error Resume Next
                auditHours = CDbl(wsUser.Cells(r, 14).value)
                nonAuditHours = CDbl(wsUser.Cells(r, 17).value)
                On Error GoTo 0
                
                auditHoursSum = auditHoursSum + auditHours
                nonAuditHoursSum = nonAuditHoursSum + nonAuditHours
                
                recordCount = recordCount + 1
            End If
        End If
    Next r
    
    FilterRecordsFromLocal = recordCount
    Exit Function
    
ErrorHandler:
    MsgBox "Error filtering local records: " & Err.Description, vbExclamation
    FilterRecordsFromLocal = 0
End Function

' ==================== CORRECTION FUNCTIONS ====================

Private Sub Commandcorrec_Click()
    On Error GoTo ErrorHandler
    
    If Not isCorrection Then
        MsgBox "Please select a record to correct by double-clicking it in the list.", vbExclamation
        Exit Sub
    End If
    
    If ValidateForm() Then
        Call SaveCorrectedRecord
       isCorrection = False
correctionSource = ""
correctionRowNumber = 0
correctionYear = 0
Me.Caption = "Timesheet Data Entry"

        Call ClearForm(False)
        Call FilterRecords
        Call ClearALL
        Me.ComboMonth.SetFocus
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error processing correction: " & Err.Description, vbCritical
End Sub

Private Function ValidateForm() As Boolean
    On Error GoTo ErrorHandler

    ValidateForm = True

    ' Validate month - must be from dropdown list
    If Me.ComboMonth.value = "" Then
        MsgBox "Please select Month.", vbExclamation
        ValidateForm = False
        Me.ComboMonth.SetFocus
        Exit Function
    End If

    If Not IsValidMonth(Me.ComboMonth.value) Then
        MsgBox "Please select a valid Month from the list.", vbExclamation
        ValidateForm = False
        Me.ComboMonth.SetFocus
        Exit Function
    End If

    ' Validate that month exists in dropdown
    If Not ValueExistsInComboBox(Me.ComboMonth, Me.ComboMonth.value) Then
        MsgBox "Invalid Month. Please select from the dropdown list only." & vbCrLf & _
               "Manual entry is not allowed.", vbExclamation, "Invalid Selection"
        ValidateForm = False
        Me.ComboMonth.value = ""
        Me.ComboMonth.SetFocus
        Exit Function
    End If

    
    ' Validate week
    If Me.ComboWeek.value = "" Then
        MsgBox "Please select Week.", vbExclamation
        ValidateForm = False
        Me.ComboWeek.SetFocus
        Exit Function
    End If

    ' Validate that week exists in dropdown
    If Not ValueExistsInComboBox(Me.ComboWeek, Me.ComboWeek.value) Then
        MsgBox "Invalid Week. Please select from the dropdown list only." & vbCrLf & _
               "Manual entry is not allowed.", vbExclamation, "Invalid Selection"
        ValidateForm = False
        Me.ComboWeek.value = ""
        Me.ComboWeek.SetFocus
        Exit Function
    End If

    ' Validate team
    If Me.ComboTeam.value = "" Then
        MsgBox "Please select Team.", vbExclamation
        ValidateForm = False
        Me.ComboTeam.SetFocus
        Exit Function
    End If
    
    ' Check if at least one engagement is filled
    Dim eng1Filled As Boolean
    Dim eng2Filled As Boolean
    
    eng1Filled = (Me.ComboEng.value <> "" And Me.TextBoxHour1.value <> "")
    eng2Filled = (Me.ComboNE.value <> "" And Me.TextBoxH2.value <> "")
    
    If Not (eng1Filled Or eng2Filled) Then
        MsgBox "Please fill at least one engagement section.", vbExclamation
        ValidateForm = False
        Exit Function
    End If
    
' Validate engagement section if filled
If eng1Filled Then
    ' Validate team (required for audit engagement)
    If Me.ComboTeam.value = "" Then
        MsgBox "Please select Team.", vbExclamation
        ValidateForm = False
        Me.ComboTeam.SetFocus
        Exit Function
    End If

    ' Validate that team exists in dropdown
    If Not ValueExistsInComboBox(Me.ComboTeam, Me.ComboTeam.value) Then
        MsgBox "Invalid Team. Please select from the dropdown list only." & vbCrLf & _
               "Manual entry is not allowed.", vbExclamation, "Invalid Selection"
        ValidateForm = False
        Me.ComboTeam.value = ""
        Me.ComboTeam.SetFocus
        Exit Function
    End If

    ' Validate region (required for audit engagement)
    If Me.ComboRegion.value = "" Then
        MsgBox "Please select Region.", vbExclamation
        ValidateForm = False
        Me.ComboRegion.SetFocus
        Exit Function
    End If

    ' Validate that region exists in dropdown
    If Not ValueExistsInComboBox(Me.ComboRegion, Me.ComboRegion.value) Then
        MsgBox "Invalid Region. Please select from the dropdown list only." & vbCrLf & _
               "Manual entry is not allowed.", vbExclamation, "Invalid Selection"
        ValidateForm = False
        Me.ComboRegion.value = ""
        Me.ComboRegion.SetFocus
        Exit Function
    End If

    ' Validate that engagement exists for this team/region combination
    If Me.ComboEng.value = "" Then
        MsgBox "Please select Engagement ID.", vbExclamation
        ValidateForm = False
        Me.ComboEng.SetFocus
        Exit Function
    End If

    ' Validate that engagement exists in dropdown
    If Not ValueExistsInComboBox(Me.ComboEng, Me.ComboEng.value) Then
        MsgBox "Invalid Engagement. Please select from the dropdown list only." & vbCrLf & _
               "This engagement does not exist for the selected Team and Region.", vbExclamation, "Invalid Selection"
        ValidateForm = False
        Me.ComboEng.value = ""
        Me.ComboEng.SetFocus
        Exit Function
    End If

    ' Validate engagement activity
    If Me.ComboActi.value = "" Then
        MsgBox "Please select Engagement Activity.", vbExclamation
        ValidateForm = False
        Me.ComboActi.SetFocus
        Exit Function
    End If

    ' Validate that activity exists in dropdown
    If Not ValueExistsInComboBox(Me.ComboActi, Me.ComboActi.value) Then
        MsgBox "Invalid Activity. Please select from the dropdown list only." & vbCrLf & _
               "Manual entry is not allowed.", vbExclamation, "Invalid Selection"
        ValidateForm = False
        Me.ComboActi.value = ""
        Me.ComboActi.SetFocus
        Exit Function
    End If

        
        ' Validate hour is numeric
        If Not IsNumeric(Me.TextBoxHour1.value) Then
            MsgBox "Engagement Hours must be a number.", vbExclamation
            ValidateForm = False
            Me.TextBoxHour1.SetFocus
            Exit Function
        End If
        
        ' Validate hour is positive
        If CDbl(Me.TextBoxHour1.value) <= 0 Then
            MsgBox "Engagement Hours must be greater than 0.", vbExclamation
            ValidateForm = False
            Me.TextBoxHour1.SetFocus
            Exit Function
        End If
    End If
    
    ' Validate non-audit section if filled
    If eng2Filled Then
        ' Validate that non-audit engagement exists in dropdown
        If Not ValueExistsInComboBox(Me.ComboNE, Me.ComboNE.value) Then
            MsgBox "Invalid Non-Audit Engagement. Please select from the dropdown list only." & vbCrLf & _
                   "Manual entry is not allowed.", vbExclamation, "Invalid Selection"
            ValidateForm = False
            Me.ComboNE.value = ""
            Me.ComboNE.SetFocus
            Exit Function
        End If

        ' Validate hour is numeric
        If Not IsNumeric(Me.TextBoxH2.value) Then
            MsgBox "Non-Audit Hours must be a number.", vbExclamation
            ValidateForm = False
            Me.TextBoxH2.SetFocus
            Exit Function
        End If

        ' Validate hour is positive
        If CDbl(Me.TextBoxH2.value) <= 0 Then
            MsgBox "Non-Audit Hours must be greater than 0.", vbExclamation
            ValidateForm = False
            Me.TextBoxH2.SetFocus
            Exit Function
        End If
    End If

    
    ' Check if ComboNE has any value and TextBoxR2 is empty
    If Trim(Me.ComboNE.value) <> "" Then
        If Trim(Me.TextBoxR2.value) = "" Then
            MsgBox "Please fill your Remark before Submit Correction." & vbCrLf & _
                   "Remark is required when Non-Audit Engagement is filled.", _
                   vbExclamation, "Remark Required"
            ValidateForm = False
            Me.TextBoxR2.SetFocus
            Exit Function
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "Error in validation: " & Err.Description, vbCritical
    ValidateForm = False
End Function

Private Sub SaveCorrectedRecord()
    On Error GoTo ErrorHandler

    Application.EnableEvents = False

    Dim recordData As Object
    Set recordData = CreateObject("Scripting.Dictionary")

    ' Store the corrected record data before saving
    recordData("UniqueCode") = ""  ' Will be filled from source

    recordData("EnterpriseID") = currentUserID
    recordData("EmployeeName") = Me.TextboxEmpname.value
    recordData("Email") = currentUseremail
    recordData("DateTime") = Format(Now, "yyyy-mm-dd hh:mm:ss")
    recordData("Type") = "Timesheet Correction"
    recordData("Month") = Me.ComboMonth.value
    recordData("Week") = Me.ComboWeek.value
    recordData("Team") = Me.ComboTeam.value
    recordData("Region") = Me.ComboRegion.value
    recordData("AuditEngagement") = Me.ComboEng.value
    recordData("EngagementActivity") = Me.ComboActi.value
    recordData("EngagementHours") = IIf(Me.TextBoxHour1.value <> "" And IsNumeric(Me.TextBoxHour1.value), Me.TextBoxHour1.value, 0)
    recordData("Remark1") = Me.TextBoxR1.value
    recordData("NonAuditEngagement") = Me.ComboNE.value
    recordData("NonAuditHours") = IIf(Me.TextBoxH2.value <> "" And IsNumeric(Me.TextBoxH2.value), Me.TextBoxH2.value, 0)
    recordData("Remark2") = Me.TextBoxR2.value
    recordData("SubmittedBy") = Application.userName

    ' Update the existing record in place
    If correctionSource = "DB" Then
        Call UpdateDatabaseRecord(recordData)
    Else
        Call UpdateLocalRecord(recordData)
    End If

    Application.EnableEvents = True

    ThisWorkbook.Save

     ' Send correction email
    Call SendCorrectionConfirmationEmail(recordData)

    ' NOW reset correction variables AFTER everything is saved
    isCorrection = False
    correctionSource = ""
    correctionRowNumber = 0
    correctionYear = 0
    Me.Caption = "Timesheet Data Entry"

    MsgBox "SUCCESS! CORRECTION completed!" & vbCrLf & _
           "A confirmation email has been sent to: " & currentUseremail, _
           vbInformation, "Correction Complete"

    Exit Sub

ErrorHandler:
    MsgBox "Error saving correction: " & Err.Description, vbCritical
    On Error Resume Next
    Application.EnableEvents = True
    ' Reset variables even on error
    isCorrection = False
    correctionSource = ""
    correctionRowNumber = 0
    correctionYear = 0
    Me.Caption = "Timesheet Data Entry"
End Sub



Private Sub UpdateDatabaseRecord(recordData As Object)
    On Error GoTo ErrorHandler

    Dim wsDB As Worksheet
    Set wsDB = ThisWorkbook.Sheets("Data base")

    wsDB.Unprotect

    With wsDB
        ' Store the unique code before marking for deletion
        recordData("UniqueCode") = .Cells(correctionRowNumber, 1).value

        ' Mark old record for deletion
        .Cells(correctionRowNumber, 20).value = "Marked for Deletion"

        ' Add new record with corrections
        Dim nextRow As Long
        nextRow = .Cells(.Rows.count, 1).End(xlUp).Row + 1

        .Cells(nextRow, 1).value = recordData("UniqueCode")              ' Same Unique Code
        .Cells(nextRow, 2).value = currentUserID                         ' Enterprise ID
        .Cells(nextRow, 3).value = Me.TextboxEmpname.value               ' Employee Name
        .Cells(nextRow, 4).value = currentUseremail                      ' Email
        .Cells(nextRow, 5).value = recordData("DateTime")                ' Date Time
        .Cells(nextRow, 6).value = "Timesheet Correction"                ' Type
        .Cells(nextRow, 7).value = Me.ComboMonth.value                   ' Month
        .Cells(nextRow, 8).value = Me.ComboWeek.value                    ' Week
        .Cells(nextRow, 9).value = Me.ComboTeam.value                    ' Team Name
        .Cells(nextRow, 10).value = Me.ComboRegion.value                 ' Region
        .Cells(nextRow, 11).value = Me.ComboEng.value                    ' Audit Engagement
        .Cells(nextRow, 12).value = Me.ComboActi.value                   ' Engagement Activity
        .Cells(nextRow, 13).value = CDbl(recordData("EngagementHours"))  ' Engagement Hr
        .Cells(nextRow, 14).value = Me.TextBoxR1.value                   ' Remark 1
        .Cells(nextRow, 15).value = Me.ComboNE.value                     ' Non-Audit Eng
        .Cells(nextRow, 16).value = CDbl(recordData("NonAuditHours"))    ' Non-Audit Hr
        .Cells(nextRow, 17).value = Me.TextBoxR2.value                   ' Remark 2
        .Cells(nextRow, 18).value = Application.userName                 ' Submitted By

        ' Update old record's Submitted By field
        .Cells(correctionRowNumber, 18).value = Application.userName
    End With

    wsDB.Protect UserInterfaceOnly:=True

    Exit Sub

ErrorHandler:
    MsgBox "Error updating database record: " & Err.Description, vbCritical
    On Error Resume Next
    wsDB.Protect UserInterfaceOnly:=True
End Sub


Private Sub UpdateLocalRecord(recordData As Object)
    On Error GoTo ErrorHandler

    Dim wsUser As Worksheet
    Set wsUser = ActiveSheet

    wsUser.Unprotect

    With wsUser
        ' Store the unique code before marking for deletion
        recordData("UniqueCode") = .Cells(correctionRowNumber, 2).value

        ' Mark old record for deletion
        .Cells(correctionRowNumber, 20).value = "Marked for Deletion"

        ' Update old record's Submitted By
        .Cells(correctionRowNumber, 19).value = Application.userName

        ' Add new record with corrections in next available row
        Dim nextRow As Long
        For nextRow = MIN_LOCAL_ROW To MAX_LOCAL_ROW
            If .Cells(nextRow, 2).value = "" Then
                .Cells(nextRow, 2).value = recordData("UniqueCode")              ' Same Unique Code
                .Cells(nextRow, 3).value = currentUserID                         ' Enterprise ID
                .Cells(nextRow, 4).value = Me.TextboxEmpname.value               ' Employee Name
                .Cells(nextRow, 5).value = currentUseremail                      ' Email
                .Cells(nextRow, 6).value = recordData("DateTime")                ' Date Time
                .Cells(nextRow, 7).value = "Timesheet Correction"                ' Type
                .Cells(nextRow, 8).value = Me.ComboMonth.value                   ' Month
                .Cells(nextRow, 9).value = Me.ComboWeek.value                    ' Week
                .Cells(nextRow, 10).value = Me.ComboTeam.value                   ' Team Name
                .Cells(nextRow, 11).value = Me.ComboRegion.value                 ' Region
                .Cells(nextRow, 12).value = Me.ComboEng.value                    ' Audit Engagement
                .Cells(nextRow, 13).value = Me.ComboActi.value                   ' Engagement Activity
                .Cells(nextRow, 14).value = CDbl(recordData("EngagementHours"))  ' Engagement Hr
                .Cells(nextRow, 15).value = Me.TextBoxR1.value                   ' Remark 1
                .Cells(nextRow, 16).value = Me.ComboNE.value                     ' Non-Audit Eng
                .Cells(nextRow, 17).value = CDbl(recordData("NonAuditHours"))    ' Non-Audit Hr
                .Cells(nextRow, 18).value = Me.TextBoxR2.value                   ' Remark 2
                .Cells(nextRow, 19).value = Application.userName                 ' Submitted By

                ' Set font color to white
                .Range(.Cells(nextRow, 2), .Cells(nextRow, 19)).Font.Color = RGB(255, 255, 255)
                Exit For
            End If
        Next nextRow
    End With

    wsUser.Protect UserInterfaceOnly:=True

    Exit Sub

ErrorHandler:
    MsgBox "Error updating local record: " & Err.Description, vbCritical
    On Error Resume Next
    wsUser.Protect UserInterfaceOnly:=True
End Sub
' ==================== SEND CORRECTION CONFIRMATION EMAIL ====================
Private Sub SendCorrectionConfirmationEmail(recordData As Object)
    On Error GoTo EmailError

    Dim OutApp As Object
    Dim OutMail As Object
    Dim emailBody As String

    ' Create Outlook application
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    ' Build email body with HTML formatting
    emailBody = BuildCorrectionEmailBody(recordData)

    ' Configure email
    With OutMail
        .To = currentUseremail
        .CC = "IAData.Analytics@sunlife.com"
        .Subject = "Timesheet Correction Confirmation - " & recordData("Week")
        .htmlBody = emailBody
        .Send ' Use .Display to review before sending, or .Send for automatic sending
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing

    Exit Sub

EmailError:
    MsgBox "Error sending confirmation email: " & Err.Description & vbCrLf & _
           "However, your correction has been successfully saved.", vbExclamation, "Email Error"
    Resume Next
End Sub

' ==================== BUILD CORRECTION EMAIL HTML BODY ====================
Private Function BuildCorrectionEmailBody(recordData As Object) As String

    Dim htmlBody As String
    Dim tableRow As String
    Dim auditHours As Double
    Dim nonAuditHours As Double

    ' Get hours for summary
    auditHours = 0
    nonAuditHours = 0

    On Error Resume Next
    auditHours = CDbl(recordData("EngagementHours"))
    nonAuditHours = CDbl(recordData("NonAuditHours"))
    On Error GoTo 0

    ' Build complete HTML email body
    htmlBody = "<!DOCTYPE html>" & vbCrLf
    htmlBody = htmlBody & "<html>" & vbCrLf
    htmlBody = htmlBody & "<head>" & vbCrLf
    htmlBody = htmlBody & "<style>" & vbCrLf
    htmlBody = htmlBody & "body { font-family: Arial, sans-serif; font-size: 14px; }" & vbCrLf
    htmlBody = htmlBody & "table { border-collapse: collapse; width: 100%; margin-top: 20px; }" & vbCrLf
    htmlBody = htmlBody & "th { background-color: #FDB913; color: white; padding: 10px; text-align: left; border: 1px solid #ddd; font-weight: bold; }" & vbCrLf
    htmlBody = htmlBody & "td { padding: 8px; border: 1px solid #ddd; }" & vbCrLf
    htmlBody = htmlBody & "tr:nth-child(even) { background-color: #f9f9f9; }" & vbCrLf
    htmlBody = htmlBody & ".summary { margin: 20px 0; padding: 15px; background-color: #fff3cd; border-left: 4px solid #ff9800; }" & vbCrLf
    htmlBody = htmlBody & ".correction-badge { display: inline-block; background-color: #ff9800; color: white; padding: 5px 10px; border-radius: 3px; font-weight: bold; }" & vbCrLf
    htmlBody = htmlBody & "</style>" & vbCrLf
    htmlBody = htmlBody & "</head>" & vbCrLf
    htmlBody = htmlBody & "<body>" & vbCrLf

    ' Greeting
    htmlBody = htmlBody & "<p>Hi <strong>" & NullToEmpty(recordData("EmployeeName")) & "</strong>,</p>" & vbCrLf

    ' Main message with correction badge
    htmlBody = htmlBody & "<p><span class='correction-badge'>CORRECTION</span> Your timesheet correction has been successfully submitted with the following details:</p>" & vbCrLf

    ' Summary box with correction styling
    htmlBody = htmlBody & "<div class='summary'>" & vbCrLf
    htmlBody = htmlBody & "<strong>Correction Summary:</strong><br>" & vbCrLf
    htmlBody = htmlBody & "Unique Code: " & NullToEmpty(recordData("UniqueCode")) & "<br>" & vbCrLf
    htmlBody = htmlBody & "Week: " & NullToEmpty(recordData("Week")) & "<br>" & vbCrLf
    htmlBody = htmlBody & "Enterprise ID: " & NullToEmpty(recordData("EnterpriseID")) & "<br>" & vbCrLf
    htmlBody = htmlBody & "Audit Hours: " & Format(auditHours, "0.00") & "<br>" & vbCrLf
    htmlBody = htmlBody & "Non-Audit Hours: " & Format(nonAuditHours, "0.00") & "<br>" & vbCrLf
    htmlBody = htmlBody & "Corrected By: " & NullToEmpty(recordData("SubmittedBy")) & "<br>" & vbCrLf
    htmlBody = htmlBody & "Correction Date/Time: " & NullToEmpty(recordData("DateTime")) & vbCrLf
    htmlBody = htmlBody & "</div>" & vbCrLf

    ' Data table
    htmlBody = htmlBody & "<table>" & vbCrLf
    htmlBody = htmlBody & "<thead>" & vbCrLf
    htmlBody = htmlBody & "<tr>" & vbCrLf
    htmlBody = htmlBody & "<th>Unique Code</th>" & vbCrLf
    htmlBody = htmlBody & "<th>Enterprise ID</th>" & vbCrLf
    htmlBody = htmlBody & "<th>Employee Name</th>" & vbCrLf
    htmlBody = htmlBody & "<th>Email</th>" & vbCrLf
    htmlBody = htmlBody & "<th>Date and Time</th>" & vbCrLf
    htmlBody = htmlBody & "<th>Type</th>" & vbCrLf
    htmlBody = htmlBody & "<th>Month</th>" & vbCrLf
    htmlBody = htmlBody & "<th>Week</th>" & vbCrLf
    htmlBody = htmlBody & "<th>Team Name</th>" & vbCrLf
    htmlBody = htmlBody & "<th>Region</th>" & vbCrLf
    htmlBody = htmlBody & "<th>Audit Engagement Name</th>" & vbCrLf
    htmlBody = htmlBody & "<th>Engagement Activity</th>" & vbCrLf
    htmlBody = htmlBody & "<th>Engagement Actual Hours</th>" & vbCrLf
    htmlBody = htmlBody & "<th>Remark 1</th>" & vbCrLf
    htmlBody = htmlBody & "<th>Other than Engagement?</th>" & vbCrLf
    htmlBody = htmlBody & "<th>Others Actual Hours</th>" & vbCrLf
    htmlBody = htmlBody & "<th>Remark 2</th>" & vbCrLf
    htmlBody = htmlBody & "<th>User Submitted Record</th>" & vbCrLf
    htmlBody = htmlBody & "</tr>" & vbCrLf
    htmlBody = htmlBody & "</thead>" & vbCrLf
    htmlBody = htmlBody & "<tbody>" & vbCrLf

    ' Build table row from record data
    tableRow = "<tr>" & vbCrLf
    tableRow = tableRow & "<td>" & NullToEmpty(recordData("UniqueCode")) & "</td>" & vbCrLf
    tableRow = tableRow & "<td>" & NullToEmpty(recordData("EnterpriseID")) & "</td>" & vbCrLf
    tableRow = tableRow & "<td>" & NullToEmpty(recordData("EmployeeName")) & "</td>" & vbCrLf
    tableRow = tableRow & "<td>" & NullToEmpty(recordData("Email")) & "</td>" & vbCrLf
    tableRow = tableRow & "<td>" & NullToEmpty(recordData("DateTime")) & "</td>" & vbCrLf
    tableRow = tableRow & "<td><strong>" & NullToEmpty(recordData("Type")) & "</strong></td>" & vbCrLf
    tableRow = tableRow & "<td>" & NullToEmpty(recordData("Month")) & "</td>" & vbCrLf
    tableRow = tableRow & "<td>" & NullToEmpty(recordData("Week")) & "</td>" & vbCrLf
    tableRow = tableRow & "<td>" & NullToEmpty(recordData("Team")) & "</td>" & vbCrLf
    tableRow = tableRow & "<td>" & NullToEmpty(recordData("Region")) & "</td>" & vbCrLf
    tableRow = tableRow & "<td>" & NullToEmpty(recordData("AuditEngagement")) & "</td>" & vbCrLf
    tableRow = tableRow & "<td>" & NullToEmpty(recordData("EngagementActivity")) & "</td>" & vbCrLf
    tableRow = tableRow & "<td>" & Format(auditHours, "0.00") & "</td>" & vbCrLf
    tableRow = tableRow & "<td>" & NullToEmpty(recordData("Remark1")) & "</td>" & vbCrLf
    tableRow = tableRow & "<td>" & NullToEmpty(recordData("NonAuditEngagement")) & "</td>" & vbCrLf
    tableRow = tableRow & "<td>" & Format(nonAuditHours, "0.00") & "</td>" & vbCrLf
    tableRow = tableRow & "<td>" & NullToEmpty(recordData("Remark2")) & "</td>" & vbCrLf
    tableRow = tableRow & "<td>" & NullToEmpty(recordData("SubmittedBy")) & "</td>" & vbCrLf
    tableRow = tableRow & "</tr>" & vbCrLf

    htmlBody = htmlBody & tableRow
    htmlBody = htmlBody & "</tbody>" & vbCrLf
    htmlBody = htmlBody & "</table>" & vbCrLf

    ' Important note about correction
    htmlBody = htmlBody & "<div style='margin-top: 20px; padding: 10px; background-color: #f0f0f0; border-left: 4px solid #ff9800;'>" & vbCrLf
    htmlBody = htmlBody & "<strong>Note:</strong> This correction has replaced your previous timesheet entry for this week. " & vbCrLf
    htmlBody = htmlBody & "The original record has been marked for deletion and this corrected version will be used for reporting." & vbCrLf
    htmlBody = htmlBody & "</div>" & vbCrLf

    ' Footer
    htmlBody = htmlBody & "<p style='margin-top: 30px;'>Thanks!</p>" & vbCrLf
    htmlBody = htmlBody & "<p style='font-size: 12px; color: #666;'><em>This is an automated confirmation email. Please do not reply to this email.</em></p>" & vbCrLf

    htmlBody = htmlBody & "</body>" & vbCrLf
    htmlBody = htmlBody & "</html>"

    BuildCorrectionEmailBody = htmlBody
End Function

' ==================== HELPER FUNCTION ====================
Private Function NullToEmpty(value As Variant) As String
    If IsNull(value) Or IsEmpty(value) Then
        NullToEmpty = ""
    Else
        NullToEmpty = CStr(value)
    End If
End Function


' ==================== FORM CLEANUP ====================

Private Sub ClearForm(ClearALL As Boolean)
    On Error GoTo ErrorHandler
    
    isUpdatingFilters = True ' Prevent filtering during clear
    
    ' Clear these fields
    Me.ComboEng.value = ""
    Me.ComboActi.value = ""
    Me.TextBoxHour1.value = ""
    Me.TextBoxR1.value = ""
    Me.ComboNE.value = ""
    Me.TextBoxH2.value = ""
    Me.TextBoxR2.value = ""
    Me.ComboRegion.value = ""
    Me.ComboTeam.value = ""
    
    If ClearALL Then
        Me.TextboxEmpname.value = ""
        Me.ComboMonth.value = ""
        Me.ComboWeek.value = ""
    End If
    
    ' Clear engagement dropdown when team or region is cleared
    Me.ComboEng.Clear
    
    isUpdatingFilters = False
    Me.ComboMonth.SetFocus
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error clearing form: " & Err.Description, vbExclamation
    isUpdatingFilters = False
End Sub

Private Sub CommandClose_Click()
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Dim response As VbMsgBoxResult
        response = MsgBox("Are you sure you want to close the correction form?", vbYesNo + vbQuestion, "Confirm Close")
        
        If response = vbNo Then
            Cancel = True
        End If
    End If
End Sub







# UserFormLOGIN code

Option Explicit

Private Sub UserForm_Initialize()
    On Error Resume Next

    ' Center form
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2

    ' Populate Enterprise IDs
    PopulateEnterpriseIDs

    ' Set focus if form is visible
    If Me.Visible Then
        Me.ComboBoxID.SetFocus
    End If

    ' Initially disable login button
    Me.CommandButtonLogin.Enabled = False

    On Error GoTo 0
End Sub

Private Sub UserForm_Activate()
End Sub

Private Sub PopulateEnterpriseIDs()
    Dim wsDV As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim dict As Object
    Dim entID As String

    On Error GoTo ErrorHandler

    ' Validate worksheet exists
    On Error Resume Next
    Set wsDV = ThisWorkbook.Sheets("DV")
    On Error GoTo ErrorHandler

    If wsDV Is Nothing Then
        MsgBox "Critical Error: DV sheet not found!", vbCritical, "Sheet Error"
        Exit Sub
    End If

    Set dict = CreateObject("Scripting.Dictionary")
    lastRow = wsDV.Cells(wsDV.Rows.count, "F").End(xlUp).Row

    Me.ComboBoxID.Clear
    Me.ComboBoxID.AddItem ""

    For i = 2 To lastRow
        entID = Trim(CStr(wsDV.Cells(i, "F").value))
        If entID <> "" And Not dict.Exists(entID) Then
            dict.Add entID, True
            Me.ComboBoxID.AddItem entID
        End If
    Next i

CleanUp:
    Set dict = Nothing
    Set wsDV = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error loading Enterprise IDs: " & Err.Description, vbCritical, "Load Error"
    Resume CleanUp
End Sub

Private Sub CommandButtonLogin_Click()
    Dim selectedID As String
    Dim normalizedID As String
    Dim actualSheetName As String
    Dim requiresPassword As Boolean
    Dim targetSheet As Worksheet

    On Error GoTo ErrorHandler

    If Me.ComboBoxID.value = "" Then
        MsgBox "Please select an Enterprise ID", vbExclamation, "Login Required"
        Me.ComboBoxID.SetFocus
        Exit Sub
    End If

    ' Store value before any form operations
    selectedID = Trim(Me.ComboBoxID.value)
    normalizedID = UCase(Replace(selectedID, " ", ""))

    ' Determine if password is required and map to actual sheet name
    requiresPassword = False
    actualSheetName = selectedID

    Select Case normalizedID
        Case "DV"
            requiresPassword = True
            actualSheetName = "DV"
        Case "FORM"
            requiresPassword = True
            actualSheetName = "Form"
        Case "DATABASE"
            requiresPassword = True
            actualSheetName = "Data base"
    End Select

    ' Check if the selected ID requires password authentication
    If requiresPassword Then
        Dim enteredPassword As String
        Dim maxAttempts As Integer
        Dim attempts As Integer

        maxAttempts = 3
        attempts = 0

        Do While attempts < maxAttempts
            enteredPassword = InputBox("This Sheet requires authentication." & vbCrLf & vbCrLf & _
                                      "Please enter the password to continue:" & vbCrLf & _
                                      "(Attempt " & (attempts + 1) & " of " & maxAttempts & ")", _
                                      "Password Required - " & actualSheetName, "")

            ' Check if user clicked Cancel
            If enteredPassword = "" Then
                MsgBox "Login cancelled.", vbInformation, "Login Cancelled"
                Exit Sub
            End If

            ' Validate password
            If enteredPassword = "123" Then
                Exit Do
            Else
                attempts = attempts + 1
                If attempts < maxAttempts Then
                    MsgBox "Incorrect password. Please try again." & vbCrLf & vbCrLf & _
                           "Remaining attempts: " & (maxAttempts - attempts), _
                           vbExclamation, "Authentication Failed"
                Else
                    MsgBox "Maximum login attempts exceeded." & vbCrLf & _
                           "Access denied for Sheet: " & actualSheetName, _
                           vbCritical, "Access Denied"
                    Me.ComboBoxID.SetFocus
                    Exit Sub
                End If
            End If
        Loop

        ' Password authentication successful - Navigate to protected sheet
        On Error Resume Next
        Set targetSheet = ThisWorkbook.Sheets(actualSheetName)
        On Error GoTo ErrorHandler

        If targetSheet Is Nothing Then
            MsgBox "Error: Sheet '" & actualSheetName & "' not found.", vbCritical, "Sheet Error"
            Exit Sub
        End If

        ' **ENHANCED: Professional sheet navigation**
        Application.ScreenUpdating = False
        Application.EnableEvents = False

        ' Unhide and activate sheet
        If targetSheet.Visible <> xlSheetVisible Then
            targetSheet.Visible = xlSheetVisible
        End If

        targetSheet.Activate

        ' **Navigate to first cell (A1) for professional appearance**
        On Error Resume Next
        targetSheet.Range("A1").Select
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
        On Error GoTo ErrorHandler

        Application.ScreenUpdating = True
        Application.EnableEvents = True

        If Err.Number <> 0 Then
            Application.ScreenUpdating = True
            Application.EnableEvents = True
            MsgBox "Could not activate sheet: " & Err.Description, vbCritical, "Activation Error"
            On Error GoTo 0
            Exit Sub
        End If
        On Error GoTo ErrorHandler

        ' Make Excel visible and maximize
        Application.Visible = True
        Application.WindowState = xlMaximized

        ' Disable button and close form
        Me.CommandButtonLogin.Enabled = False
        Me.Hide
        DoEvents

        ' Show success message after form is hidden
        MsgBox "Successfully logged in to " & actualSheetName & " sheet.", _
               vbInformation, "Login Successful"

        Unload Me
        Exit Sub
    End If

    ' **Normal login flow for regular Enterprise IDs**
    Me.CommandButtonLogin.Enabled = False
    Me.Hide
    DoEvents

    ' Call the login function with the Enterprise ID
    Call LoginUserFromForm(selectedID)

    ' Close the form
    Unload Me
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
    Me.CommandButtonLogin.Enabled = True
End Sub

Private Sub CommandButtonCancel_Click()
    Dim response As VbMsgBoxResult

    On Error Resume Next

    response = MsgBox("Are you sure you want to cancel login?" & vbCrLf & _
                     "This will close the workbook.", vbYesNo + vbQuestion, "Confirm Cancel")

    If response = vbYes Then
        ' Proper cleanup sequence
        Me.Hide
        DoEvents
        Unload Me
        DoEvents

        Application.DisplayAlerts = False
        Application.EnableEvents = False
        Application.Visible = True
        Application.WindowState = xlMaximized
        ThisWorkbook.Close SaveChanges:=False
        Application.DisplayAlerts = True
        Application.EnableEvents = True
    End If

    On Error GoTo 0
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Dim response As VbMsgBoxResult

    On Error Resume Next

    If CloseMode = vbFormControlMenu Then
        response = MsgBox("Are you sure you want to cancel login?" & vbCrLf & _
                         "This will close the workbook.", vbYesNo + vbQuestion, "Confirm Cancel")

        If response = vbNo Then
            Cancel = True
        Else
            Me.Hide
            DoEvents

            Application.DisplayAlerts = False
            Application.EnableEvents = False
            Application.Visible = True
            Application.WindowState = xlMaximized
            ThisWorkbook.Close SaveChanges:=False
            Application.DisplayAlerts = True
            Application.EnableEvents = True
        End If
    End If

    On Error GoTo 0
End Sub

Private Sub ComboBoxID_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then ' Enter key
        Call CommandButtonLogin_Click
    End If
    On Error GoTo 0
End Sub

Private Sub ComboBoxID_Change()
    On Error Resume Next
    If Me.ComboBoxID.value <> "" Then
        Me.CommandButtonLogin.Enabled = True
    Else
        Me.CommandButtonLogin.Enabled = False
    End If
    On Error GoTo 0
End Sub





