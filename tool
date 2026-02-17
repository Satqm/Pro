We’ve identified two main issues causing slow performance and an annoying email pop‑up:

1. Email confirmation appears instead of being sent silently – because the code uses .Display (shows the email) followed by .Send.
2. Duplicate checking scans the entire database for each row – this is the main reason submission feels slow when the database grows.

Below are simple find‑and‑replace steps to fix both.
Important: Copy the exact surrounding lines so you know where to make the change.
Only the lines shown need to be replaced; everything else stays the same.

---

Fix 1 – Send email without displaying it

Where to look:
In Module 1, find the procedure SendTimesheetConfirmationEmail.
Inside that procedure, locate the block that starts with With OutMail and ends with End With.

Find these exact lines (including the comment lines):

```vb
        .Display ' Use .Send to send automatically, .Display to show before sending
        ' For automatic sending, change .Display to .Send
        .Send
```

Replace them with:

```vb
        .Send ' Send email automatically without displaying
```

After this change, the email will be sent directly to the recipient’s inbox – no window will pop up.

---

Fix 2 – Speed up duplicate checking (major performance boost)

The current code checks every row in the database five times (once for each timesheet row).
We will replace the whole HasDuplicateEntries function with a new version that scans the database only once and uses a fast lookup table.

Where to look:
In Module 1, find the function HasDuplicateEntries. It starts with:

```vb
' ==================== ENHANCED DUPLICATE CHECKING WITH ROW NAVIGATION ====================
Private Function HasDuplicateEntries(wsForm As Worksheet, wsDB As Worksheet, enterpriseID As String, _
                                   monthName As String, weekName As String, teamName As String) As Boolean
```

Replace the entire function (from that line down to the matching End Function) with the new code below.
The new function is longer, but it will make submission much faster, especially when many records exist.

New HasDuplicateEntries function (paste this in place of the old one):

```vb
' ==================== OPTIMIZED DUPLICATE CHECKING ====================
Private Function HasDuplicateEntries(wsForm As Worksheet, wsDB As Worksheet, enterpriseID As String, _
                                   monthName As String, weekName As String, teamName As String) As Boolean
    Dim localKeys As Object, dbKeys As Object
    Dim r As Long, i As Integer
    Dim key As String
    Dim lastRow As Long
    Dim region As String, engagementID As String, activity As String
    Dim nonAuditEng As String, nonAuditRemarks As String
    Dim inputKey As String

    ' Create dictionaries for fast lookup
    Set localKeys = CreateObject("Scripting.Dictionary")
    Set dbKeys = CreateObject("Scripting.Dictionary")

    ' ---- Load all existing records for this user & week from local storage ----
    For r = 3000 To 3200
        If wsForm.Cells(r, 3).value = enterpriseID And _
           wsForm.Cells(r, 9).value = weekName And _
           wsForm.Cells(r, 20).value <> "Marked for Deletion" Then

            key = BuildDuplicateKey( _
                wsForm.Cells(r, 3).value, _  ' EnterpriseID
                wsForm.Cells(r, 8).value, _  ' Month
                wsForm.Cells(r, 9).value, _  ' Week
                wsForm.Cells(r, 10).value, _ ' Team
                wsForm.Cells(r, 11).value, _ ' Region
                wsForm.Cells(r, 12).value, _ ' Audit Engagement
                wsForm.Cells(r, 16).value, _ ' Non-Audit Engagement
                wsForm.Cells(r, 13).value, _ ' Engagement Activity
                wsForm.Cells(r, 17).value)   ' Non-Audit Remarks (Remark 2)
            localKeys(key) = True
        End If
    Next r

    ' ---- Load all existing records for this user & week from the database ----
    lastRow = wsDB.Cells(wsDB.Rows.count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If wsDB.Cells(r, 2).value = enterpriseID And _
           wsDB.Cells(r, 8).value = weekName Then

            key = BuildDuplicateKey( _
                wsDB.Cells(r, 2).value, _  ' EnterpriseID
                wsDB.Cells(r, 7).value, _  ' Month
                wsDB.Cells(r, 8).value, _  ' Week
                wsDB.Cells(r, 9).value, _  ' Team
                wsDB.Cells(r, 10).value, _ ' Region
                wsDB.Cells(r, 11).value, _ ' Audit Engagement
                wsDB.Cells(r, 15).value, _ ' Non-Audit Engagement
                wsDB.Cells(r, 12).value, _ ' Engagement Activity
                wsDB.Cells(r, 16).value)   ' Non-Audit Remarks (Remark 2)
            dbKeys(key) = True
        End If
    Next r

    ' ---- Check each row the user is trying to submit ----
    For i = 10 To 14
        If IsEngagementRowValid(wsForm, i) Then
            ' Determine if this is an Audit or Non-Audit row
            If Len(Trim(wsForm.Cells(i, 3).value)) > 0 Then
                ' Audit engagement
                region = Trim(wsForm.Cells(i, 2).value)
                engagementID = Trim(wsForm.Cells(i, 3).value)
                activity = Trim(wsForm.Cells(i, 4).value)
                nonAuditEng = ""
                nonAuditRemarks = ""
            Else
                ' Non-audit engagement
                region = ""
                engagementID = Trim(wsForm.Cells(i, 8).value)
                activity = ""
                nonAuditEng = engagementID
                nonAuditRemarks = Trim(wsForm.Cells(i, 9).value)
            End If

            ' Build a key for this input row
            inputKey = BuildDuplicateKey( _
                enterpriseID, monthName, weekName, teamName, _
                region, engagementID, nonAuditEng, activity, nonAuditRemarks)

            ' Check both dictionaries
            If localKeys.Exists(inputKey) Or dbKeys.Exists(inputKey) Then
                ' Duplicate found – handle it exactly as before
                Dim duplicateLocation As String
                duplicateLocation = IIf(localKeys.Exists(inputKey), "Local Storage (Pending Sync)", "Database (Already Synced)")

                ' Highlight the row
                NavigateToErrorCell wsForm, i, IIf(Len(Trim(wsForm.Cells(i, 3).value)) > 0, "Audit", "Non-Audit")

                ' Show message (same as original)
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
                                "Type: " & IIf(Len(Trim(wsForm.Cells(i, 3).value)) > 0, "Audit", "Non-Audit") & " Engagement" & vbCrLf & _
                                "Engagement ID: " & engagementID & vbCrLf & _
                                IIf(region <> "", "Region: " & region & vbCrLf, "") & _
                                IIf(activity <> "", "Activity: " & activity & vbCrLf, "") & _
                                IIf(nonAuditRemarks <> "", "Remarks: " & nonAuditRemarks & vbCrLf, "") & _
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
                Else
                    MsgBox "Submission cancelled." & vbCrLf & vbCrLf & _
                           "Please review " & rowLabel & " and modify the duplicate entry.", _
                           vbInformation, "Submission Cancelled"
                End If

                HasDuplicateEntries = True
                Exit Function
            End If
        End If
    Next i

    HasDuplicateEntries = False
End Function
```

After replacing the function, add this new helper function right after it (anywhere after the End Function of HasDuplicateEntries and before the next function):

```vb
' ==================== Helper to build a unique key for duplicate detection ====================
Private Function BuildDuplicateKey(entID As Variant, monthVal As Variant, weekVal As Variant, _
                                  teamVal As Variant, regionVal As Variant, auditEng As Variant, _
                                  nonAuditEng As Variant, activityVal As Variant, remarksVal As Variant) As String
    ' Combine all fields with a pipe separator; empty values become empty strings
    BuildDuplicateKey = entID & "|" & monthVal & "|" & weekVal & "|" & teamVal & "|" & _
                        regionVal & "|" & auditEng & "|" & nonAuditEng & "|" & activityVal & "|" & remarksVal
End Function
```

---

What these changes do

· Email fix – removes the pop‑up so the email is sent silently in the background.
· Duplicate check rewrite – scans the database once instead of five times, and uses an in‑memory lookup for all rows. This will dramatically speed up submission, especially as more data accumulates.

After making these two replacements, save the workbook (as macro‑enabled) and test the submission process. You should notice a much faster response and no more email windows interrupting your work.

If you still experience lag when multiple users are working simultaneously, please let us know – we can suggest additional optimizations for the real‑time dropdown updates.
