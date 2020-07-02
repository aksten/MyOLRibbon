'Private Sub cboEmployee_FK_AfterUpdate()
'    If (Me.txtEmployeeEmail = vbNullString) Or Me.txtEmployeeEmail = "" Then
'        Me.txtTo = Nz(Me.txtForemanEmail, "")
'    Else
'        If IsNull(Me.txtForemanEmail) Or Me.txtForemanEmail = "" Then
'            Me.txtTo = ""
'        ElseIf Me.txtEmployeeEmail = "Jfissel@tlsokc.com" Or Me.txtEmployeeEmail = "Rjimenez@tlsokc.com" Or Me.txtEmployeeEmail = "Ldelgado@tlsokc.com" Then
'            Me.txtTo = "JFissel@tlsokc.com;RJimenez@tlsokc.com;LDelgado@tlsokc.com"
'        ElseIf Me.txtEmployeeEmail = "efrench@tlsokc.com" Or Me.txtEmployeeEmail = "jmudge@tlsokc.com" Or Me.txtEmployeeEmail = "sstrode@tlsokc.com" Or Me.txtEmployeeEmail = "alittlefield@tlsokc.com" _
'            Or Me.txtEmployeeEmail = "tgreen@tlsokc.com" Or Me.txtEmployeeEmail = "scruz@tlsokc.com" Or Me.txtEmployeeEmail = "jguerrero@tlsokc.com" Or Me.txtEmployeeEmail = "nvandalsem@tlsokc.com" Or Me.txtEmployeeEmail = "nvandalsem@tlsokc.com" Then
'            Me.txtTo = "TLS Mechanic Shop"
'        ElseIf Me.txtEmployeeEmail = Me.txtForemanEmail Then
'            Me.txtTo = Me.txtEmployeeEmail
'        Else
'            Me.txtTo = Me.txtEmployeeEmail & ";" & Me.txtForemanEmail
'        End If
'    End If
'End Sub


'Private Sub chkAllDay_AfterUpdate()
'If Me.chkAllDay = True Then
'    If IsNull(Me.txtEndDate = Me.txtStartDate) Then
'        Me.txtStartTime = TimeValue("0:00")
'    End If
'    Me.txtEndTime = TimeValue("23:59")
'    Me.txtSubject.Visible = True
'End If
'If Me.txtStartDate = Me.txtEndDate And Me.txtStartTime > TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = False Then
'        'Debug.Print "OneDayLeavingEarlyUnpaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be leaving at " & Format(Me.txtStartTime, "h:nn ampm") & " on " & Me.txtStartDate & " - No PTO"
'    ElseIf Me.txtStartDate = Me.txtEndDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime < TimeValue("23:59") And Me.PaidYN = False Then
'        'Debug.Print "OneDayArrivingLateUnpaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be coming in at " & Format(Me.txtEndTime, "h:nn ampm") & " on " & Me.txtStartDate & " - No PTO"
'    ElseIf Me.txtEndDate > Me.txtStartDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime < TimeValue("23:59") And Me.PaidYN = False Then
'        'Debug.Print "OutMultiDaysReturningMidDayUnpaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be out from " & Me.txtStartDate & " to " & Me.txtEndDate & " at " & Format(Me.txtEndTime, "h:nn ampm") & " - No PTO"
'    ElseIf Me.txtEndDate > Me.txtStartDate And Me.txtStartTime > TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = False Then
'        ''Debug.Print "OutMultiDaysLeavingEarlyUnPaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be leaving at " & Format(Me.txtStartTime, "h:nn ampm") & " on " & Me.txtStartDate & " out through " & Me.txtEndDate & " - No PTO"
'    ElseIf Me.txtEndDate > Me.txtStartDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = False Then
'        'Debug.Print "OutMultipleFullDaysUnPaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be out " & Me.txtStartDate & " to " & Me.txtEndDate & " - No PTO"
'    ElseIf Me.txtStartDate = Me.txtEndDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = False Then
'        'Debug.Print "OneFullDayOutUnPaid"
'        Me.txtSubject = Me.EmployeeName & " will be out " & Me.txtStartDate & " - No PTO"
'    ElseIf Me.txtStartDate = Me.txtEndDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime < TimeValue("23:59") And Me.PaidYN = True Then
'        'Debug.Print "OneDayArrivingLatePaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be coming in at " & Format(Me.txtEndTime, "h:nn ampm") & " on " & Me.txtStartDate & " - PTO"
'    ElseIf Me.txtStartDate = Me.txtEndDate And Me.txtStartTime > TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = True Then
'        'Debug.Print "OneDayLeavingEarlyPaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be leaving at " & Format(Me.txtStartTime, "h:nn ampm") & " on " & Me.txtStartDate & " - PTO"
'    ElseIf Me.txtEndDate > Me.txtStartDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime < TimeValue("23:59") And Me.PaidYN = True Then
'        'Debug.Print "OutMultiDaysReturningMidDayPaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be out from " & Me.txtStartDate & " to " & Me.txtEndDate & " at " & Format(Me.txtEndTime, "h:nn ampm") & " - PTO"
'    ElseIf Me.txtEndDate > Me.txtStartDate And Me.txtStartTime > TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = True Then
'        'Debug.Print "OutMultiDaysLeavingEarlyPaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be leaving at " & Format(Me.txtStartTime, "h:nn ampm") & " on " & Me.txtStartDate & " out through " & Me.txtEndDate & " - PTO"
'    ElseIf Me.txtEndDate > Me.txtStartDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = True Then
'        'Debug.Print "OutMultipleFullDaysPaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be out " & Me.txtStartDate & " to " & Me.txtEndDate & " - PTO"
'    ElseIf Me.txtStartDate = Me.txtEndDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = True Then
'        'Debug.Print "OneFullDayOutPaid"
'        Me.txtSubject = Me.EmployeeName & " will be out " & Me.txtStartDate & " - PTO"
'    ElseIf Me.txtStartDate = Me.txtEndDate And Me.txtStartTime > TimeValue("0:00") And Me.txtEndTime < TimeValue("23:59") And Me.PaidYN = True Then
'        'Debug.Print "OneDayLeaveAndReturnPaid"
'        Me.txtSubject = Me.EmployeeName & " will be leaving at " & Format(Me.txtStartTime, "h:nn ampm") & " and returning at " & Format(Me.txtEndTime, "h:nn ampm") & " on " & Me.txtStartDate & " - PTO"
'    ElseIf Me.txtStartDate = Me.txtEndDate And Me.txtStartTime > TimeValue("0:00") And Me.txtEndTime < TimeValue("23:59") And Me.PaidYN = False Then
'        'Debug.Print "OneDayLeaveAndReturnUnPaid"
'        Me.txtSubject = Me.EmployeeName & " will be leaving at " & Format(Me.txtStartTime, "h:nn ampm") & " and returning at " & Format(Me.txtEndTime, "h:nn ampm") & " on " & Me.txtStartDate & " - No PTO"
'    End If
'End Sub

'Private Sub chkPaid_AfterUpdate()
'    If Me.txtStartDate = Me.txtEndDate And Me.txtStartTime > TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = False Then
'        'Debug.Print "OneDayLeavingEarlyUnpaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be leaving at " & Format(Me.txtStartTime, "h:nn ampm") & " on " & Me.txtStartDate & " - No PTO"
'    ElseIf Me.txtStartDate = Me.txtEndDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime < TimeValue("23:59") And Me.PaidYN = False Then
'        'Debug.Print "OneDayArrivingLateUnpaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be coming in at " & Format(Me.txtEndTime, "h:nn ampm") & " on " & Me.txtStartDate & " - No PTO"
'    ElseIf Me.txtEndDate > Me.txtStartDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime < TimeValue("23:59") And Me.PaidYN = False Then
'        'Debug.Print "OutMultiDaysReturningMidDayUnpaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be out from " & Me.txtStartDate & " to " & Me.txtEndDate & " at " & Format(Me.txtEndTime, "h:nn ampm") & " - No PTO"
'    ElseIf Me.txtEndDate > Me.txtStartDate And Me.txtStartTime > TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = False Then
'        'Debug.Print "OutMultiDaysLeavingEarlyUnPaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be leaving at " & Format(Me.txtStartTime, "h:nn ampm") & " on " & Me.txtStartDate & " out through " & Me.txtEndDate & " - No PTO"
'    ElseIf Me.txtEndDate > Me.txtStartDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = False Then
'        'Debug.Print "OutMultipleFullDaysUnPaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be out " & Me.txtStartDate & " to " & Me.txtEndDate & " - No PTO"
'    ElseIf Me.txtStartDate = Me.txtEndDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = False Then
'        'Debug.Print "OneFullDayOutUnPaid"
'        Me.txtSubject = Me.EmployeeName & " will be out " & Me.txtStartDate & " - No PTO"
'    ElseIf Me.txtStartDate = Me.txtEndDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime < TimeValue("23:59") And Me.PaidYN = True Then
'        'Debug.Print "OneDayArrivingLatePaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be coming in at " & Format(Me.txtEndTime, "h:nn ampm") & " on " & Me.txtStartDate & " - PTO"
'    ElseIf Me.txtStartDate = Me.txtEndDate And Me.txtStartTime > TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = True Then
'        'Debug.Print "OneDayLeavingEarlyPaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be leaving at " & Format(Me.txtStartTime, "h:nn ampm") & " on " & Me.txtStartDate & " - PTO"
'    ElseIf Me.txtEndDate > Me.txtStartDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime < TimeValue("23:59") And Me.PaidYN = True Then
'        'Debug.Print "OutMultiDaysReturningMidDayPaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be out from " & Me.txtStartDate & " to " & Me.txtEndDate & " at " & Format(Me.txtEndTime, "h:nn ampm") & " - PTO"
'    ElseIf Me.txtEndDate > Me.txtStartDate And Me.txtStartTime > TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = True Then
'        'Debug.Print "OutMultiDaysLeavingEarlyPaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be leaving at " & Format(Me.txtStartTime, "h:nn ampm") & " on " & Me.txtStartDate & " out through " & Me.txtEndDate & " - PTO"
'    ElseIf Me.txtEndDate > Me.txtStartDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = True Then
'        'Debug.Print "OutMultipleFullDaysPaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be out " & Me.txtStartDate & " to " & Me.txtEndDate & " - PTO"
'    ElseIf Me.txtStartDate = Me.txtEndDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = True Then
'        'Debug.Print "OneFullDayOutPaid"
'        Me.txtSubject = Me.EmployeeName & " will be out " & Me.txtStartDate & " - PTO"
'    ElseIf Me.txtStartDate = Me.txtEndDate And Me.txtStartTime > TimeValue("0:00") And Me.txtEndTime < TimeValue("23:59") And Me.PaidYN = True Then
'        'Debug.Print "OneDayLeaveAndReturnPaid"
'        Me.txtSubject = Me.EmployeeName & " will be leaving at " & Format(Me.txtStartTime, "h:nn ampm") & " and returning at " & Format(Me.txtEndTime, "h:nn ampm") & " on " & Me.txtStartDate & " - PTO"
'    ElseIf Me.txtStartDate = Me.txtEndDate And Me.txtStartTime > TimeValue("0:00") And Me.txtEndTime < TimeValue("23:59") And Me.PaidYN = False Then
'        'Debug.Print "OneDayLeaveAndReturnUnPaid"
'        Me.txtSubject = Me.EmployeeName & " will be leaving at " & Format(Me.txtStartTime, "h:nn ampm") & " and returning at " & Format(Me.txtEndTime, "h:nn ampm") & " on " & Me.txtStartDate & " - No PTO"
'    End If
'End Sub

'Private Sub cmdCheck_Click()
'    If CheckExists(Me.txtSubject) = True Then
'        MessageBox.Show("Meeting exists", vbOKOnly)
'    Else
'        MessageBox.Show("Meeting does not exist", vbOKOnly)
'    End If
'End Sub

'Private Sub cmdCleanUp_Click()
'    Call myPowershell("C:\Scripts\AdobeStamps\CleanUpMovedFiles.ps1")
'End Sub

''Private Sub cmdMove_Click()
''    Dim myFullEmpPath As String
''    Dim myOldFileName As String
''    Dim myFileName As String
''    myOldFileName = Me.txtFileName & " - Copy2.pdf"
''    myFileName = Me.txtFileName & ".pdf"
''    myFullEmpPath = DLookup("EmployeeFolderPathBase", "tblPaths", "[PathID] = 1") & DLookup("EmployeeFolderPath", "tblEmployees", "[EmployeeID] = " & Me.Employee_FK) & "\" & myFileName
''    'Debug.Print myFullEmpPath
''    Name "\\tlsok.local\userdata\TulsaUsers\ASmedley\Desktop\PTO\" & myOldFileName As myFullEmpPath
''End Sub

''Private Sub cmdCopy_Click()
''    Call myPowershell("C:\Scripts\AdobeStamps\CopyPTOPDF2Requests.ps1")
''End Sub

'Sub cmdCreateMeeting_Click()

'    On Error GoTo Err_Handler

'    oSendTo = Me.txtTo
'    oSubject = Me.txtSubject
'    oBody = Me.txtBody
'    oLocation = "Tulsa Division"
'    oStartDate = Me.txtStartDate
'    oStartTime = Me.txtStartTime
'    oEndDate = Me.txtEndDate
'    oEndTime = Me.txtEndTime
'    'Debug.Print Me.txtStartDate & ": " & oStartDate
'    'Debug.Print Me.txtStartTime & ": " & oStartTime
'    'Debug.Print Me.txtEndDate & ": " & oEndDate
'    'Debug.Print Me.txtEndTime & ": " & oEndTime
'    If Me.chkAllDay = True Then
'        oFullDay = True
'    Else
'        oFullDay = False
'    End If

'    If Len(Me.txtAttachments) > 0 Then
'        oAtt = Me.txtAttachments
'    End If
'    Call GenerateMeeting

'Exit_Handler:
'    Exit Sub

'Err_Handler:
'    If Err.Number = 0 Then
'        Resume Next
'    Else
'        MessageBox.Show Err.Number & ": " & Err.Description & vbCrLf & Err.Source, vbOKOnly, "ERROR"
'Resume Exit_Handler
'    End If

'End Sub

'Sub cmdEmail_Click()
'    Dim myOldFileName As String
'    Dim myFullE2OKCPath As String
'    Dim myNewFileName As String
'    myOldFileName = "\\tlsok.local\userdata\TulsaUsers\ASmedley\Desktop\PTO\" & Me.txtFileName & " - Copy.pdf"
'    myNewFileName = Me.txtFileName & ".pdf"
'    myFullE2OKCPath = DLookup("E2OKCPath", "tblPaths", "[PathID] = 1") & myNewFileName
'    On Error GoTo Err_Handler

'    oSendTo = "tcarothers@tlsokc.com"
'    If Me.txtSubject Like "*was*" Or Me.txtSubject Like "*left*" Then
'        oSubject = "TIME OFF REQUEST(S): " & Left(Me.txtSubject, InStr(1, Me.txtSubject, " was") - 1)
'    Else
'        oSubject = "TIME OFF REQUEST(S): " & Left(Me.txtSubject, InStr(1, Me.txtSubject, " will") - 1)
'    End If
'    oBody = Me.txtBody
'    If Len(Me.txtAttachments) > 0 Then
'        oAtt = Me.txtAttachments
'    End If

'    Call GenerateEmail

'    Name myOldFileName As myFullE2OKCPath

'Exit_Handler:
'    Exit Sub

'Err_Handler:
'    If Err.Number = 0 Then
'        Resume Next
'    Else
'        MessageBox.Show Err.Number & ": " & Err.Description & vbCrLf & Err.Source, vbOKOnly, "ERROR"
'Resume Exit_Handler
'    End If

'End Sub

'Private Sub cmdSaveClose_Click()
'    DoCmd.RunCommand acCmdSaveRecord
'Application.Quit
'End Sub

'Private Sub cmdSaveNew_Click()
'    DoCmd.RunCommand acCmdSaveRecord
'DoCmd.GoToRecord , , acNewRec
'DoCmd.GoToControl "cboEmployee_FK"
'End Sub

'Private Sub cmdStampPTO_Click()
'    Call myPowershell("C:\Scripts\AdobeStamps\StampPTOPDFs.ps1")
'End Sub

'Private Sub Form_Current()
'    If Me.NewRecord = True Then
'        Me.txtTo = ""
'        Me.txtFileName = ""
'        Exit Sub
'    Else
'        Me.txtFileName = ""
'        If IsNull(Me.txtEmployeeEmail) Or Me.txtEmployeeEmail = "" Then
'            Me.txtTo = Nz(Me.txtForemanEmail, "")
'        Else
'            If IsNull(Me.txtForemanEmail) Or Me.txtForemanEmail = "" Then
'                Me.txtTo = ""
'            ElseIf Me.txtEmployeeEmail = "Jfissel@tlsokc.com" Or Me.txtEmployeeEmail = "Rjimenez@tlsokc.com" Or Me.txtEmployeeEmail = "Ldelgado@tlsokc.com" Then
'                Me.txtTo = "JFissel@tlsokc.com;RJimenez@tlsokc.com;LDelgado@tlsokc.com"
'            ElseIf Me.txtEmployeeEmail = "efrench@tlsokc.com" Or Me.txtEmployeeEmail = "jmudge@tlsokc.com" Or Me.txtEmployeeEmail = "sstrode@tlsokc.com" Or Me.txtEmployeeEmail = "alittlefield@tlsokc.com" _
'            Or Me.txtEmployeeEmail = "tgreen@tlsokc.com" Or Me.txtEmployeeEmail = "scruz@tlsokc.com" Or Me.txtEmployeeEmail = "jguerrero@tlsokc.com" Or Me.txtEmployeeEmail = "nvandalsem@tlsokc.com" Or Me.txtEmployeeEmail = "nvandalsem@tlsokc.com" Then
'                Me.txtTo = "TLS Mechanic Shop"
'            ElseIf Me.txtEmployeeEmail = Me.txtForemanEmail Then
'                Me.txtTo = Me.txtEmployeeEmail
'            Else
'                Me.txtTo = Me.txtEmployeeEmail & ";" & Me.txtForemanEmail
'            End If
'        End If
'    End If
'End Sub

'Private Sub Form_Load()
'    DoCmd.GoToRecord , , acNewRec
'End Sub


'Private Sub txtEndDate_AfterUpdate()
'    Me.txtSubject.Visible = True

'    If Me.txtStartDate = Me.txtEndDate And Me.txtStartTime > TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = False Then
'        'Debug.Print "OneDayLeavingEarlyUnpaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be leaving at " & Format(Me.txtStartTime, "h:nn ampm") & " on " & Me.txtStartDate & " - No PTO"
'    ElseIf Me.txtStartDate = Me.txtEndDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime < TimeValue("23:59") And Me.PaidYN = False Then
'        'Debug.Print "OneDayArrivingLateUnpaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be coming in at " & Format(Me.txtEndTime, "h:nn ampm") & " on " & Me.txtStartDate & " - No PTO"
'    ElseIf Me.txtEndDate > Me.txtStartDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime < TimeValue("23:59") And Me.PaidYN = False Then
'        'Debug.Print "OutMultiDaysReturningMidDayUnpaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be out from " & Me.txtStartDate & " to " & Me.txtEndDate & " at " & Format(Me.txtEndTime, "h:nn ampm") & " - No PTO"
'    ElseIf Me.txtEndDate > Me.txtStartDate And Me.txtStartTime > TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = False Then
'        'Debug.Print "OutMultiDaysLeavingEarlyUnPaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be leaving at " & Format(Me.txtStartTime, "h:nn ampm") & " on " & Me.txtStartDate & " out through " & Me.txtEndDate & " - No PTO"
'    ElseIf Me.txtEndDate > Me.txtStartDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = False Then
'        'Debug.Print "OutMultipleFullDaysUnPaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be out " & Me.txtStartDate & " to " & Me.txtEndDate & " - No PTO"
'    ElseIf Me.txtStartDate = Me.txtEndDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = False Then
'        'Debug.Print "OneFullDayOutUnPaid"
'        Me.txtSubject = Me.EmployeeName & " will be out " & Me.txtStartDate & " - No PTO"
'    ElseIf Me.txtStartDate = Me.txtEndDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime < TimeValue("23:59") And Me.PaidYN = True Then
'        'Debug.Print "OneDayArrivingLatePaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be coming in at " & Format(Me.txtEndTime, "h:nn ampm") & " on " & Me.txtStartDate & " - PTO"
'    ElseIf Me.txtStartDate = Me.txtEndDate And Me.txtStartTime > TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = True Then
'        'Debug.Print "OneDayLeavingEarlyPaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be leaving at " & Format(Me.txtStartTime, "h:nn ampm") & " on " & Me.txtStartDate & " - PTO"
'    ElseIf Me.txtEndDate > Me.txtStartDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime < TimeValue("23:59") And Me.PaidYN = True Then
'        'Debug.Print "OutMultiDaysReturningMidDayPaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be out from " & Me.txtStartDate & " to " & Me.txtEndDate & " at " & Format(Me.txtEndTime, "h:nn ampm") & " - PTO"
'    ElseIf Me.txtEndDate > Me.txtStartDate And Me.txtStartTime > TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = True Then
'        'Debug.Print "OutMultiDaysLeavingEarlyPaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be leaving at " & Format(Me.txtStartTime, "h:nn ampm") & " on " & Me.txtStartDate & " out through " & Me.txtEndDate & " - PTO"
'    ElseIf Me.txtEndDate > Me.txtStartDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = True Then
'        'Debug.Print "OutMultipleFullDaysPaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be out " & Me.txtStartDate & " to " & Me.txtEndDate & " - PTO"
'    ElseIf Me.txtStartDate = Me.txtEndDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = True Then
'        'Debug.Print "OneFullDayOutPaid"
'        Me.txtSubject = Me.EmployeeName & " will be out " & Me.txtStartDate & " - PTO"
'    ElseIf Me.txtStartDate = Me.txtEndDate And Me.txtStartTime > TimeValue("0:00") And Me.txtEndTime < TimeValue("23:59") And Me.PaidYN = True Then
'        'Debug.Print "OneDayLeaveAndReturnPaid"
'        Me.txtSubject = Me.EmployeeName & " will be leaving at " & Format(Me.txtStartTime, "h:nn ampm") & " and returning at " & Format(Me.txtEndTime, "h:nn ampm") & " on " & Me.txtStartDate & " - PTO"
'    ElseIf Me.txtStartDate = Me.txtEndDate And Me.txtStartTime > TimeValue("0:00") And Me.txtEndTime < TimeValue("23:59") And Me.PaidYN = False Then
'        'Debug.Print "OneDayLeaveAndReturnUnPaid"
'        Me.txtSubject = Me.EmployeeName & " will be leaving at " & Format(Me.txtStartTime, "h:nn ampm") & " and returning at " & Format(Me.txtEndTime, "h:nn ampm") & " on " & Me.txtStartDate & " - No PTO"
'    End If
'End Sub

'Private Sub txtEndTime_AfterUpdate()
'    Me.txtSubject.Visible = True

'    If Me.txtStartDate = Me.txtEndDate And Me.txtStartTime > TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = False Then
'        'Debug.Print "OneDayLeavingEarlyUnpaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be leaving at " & Format(Me.txtStartTime, "h:nn ampm") & " on " & Me.txtStartDate & " - No PTO"
'    ElseIf Me.txtStartDate = Me.txtEndDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime < TimeValue("23:59") And Me.PaidYN = False Then
'        'Debug.Print "OneDayArrivingLateUnpaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be coming in at " & Format(Me.txtEndTime, "h:nn ampm") & " on " & Me.txtStartDate & " - No PTO"
'    ElseIf Me.txtEndDate > Me.txtStartDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime < TimeValue("23:59") And Me.PaidYN = False Then
'        'Debug.Print "OutMultiDaysReturningMidDayUnpaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be out from " & Me.txtStartDate & " to " & Me.txtEndDate & " at " & Format(Me.txtEndTime, "h:nn ampm") & " - No PTO"
'    ElseIf Me.txtEndDate > Me.txtStartDate And Me.txtStartTime > TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = False Then
'        'Debug.Print "OutMultiDaysLeavingEarlyUnPaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be leaving at " & Format(Me.txtStartTime, "h:nn ampm") & " on " & Me.txtStartDate & " out through " & Me.txtEndDate & " - No PTO"
'    ElseIf Me.txtEndDate > Me.txtStartDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = False Then
'        'Debug.Print "OutMultipleFullDaysUnPaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be out " & Me.txtStartDate & " to " & Me.txtEndDate & " - No PTO"
'    ElseIf Me.txtStartDate = Me.txtEndDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = False Then
'        'Debug.Print "OneFullDayOutUnPaid"
'        Me.txtSubject = Me.EmployeeName & " will be out " & Me.txtStartDate & " - No PTO"
'    ElseIf Me.txtStartDate = Me.txtEndDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime < TimeValue("23:59") And Me.PaidYN = True Then
'        'Debug.Print "OneDayArrivingLatePaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be coming in at " & Format(Me.txtEndTime, "h:nn ampm") & " on " & Me.txtStartDate & " - PTO"
'    ElseIf Me.txtStartDate = Me.txtEndDate And Me.txtStartTime > TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = True Then
'        'Debug.Print "OneDayLeavingEarlyPaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be leaving at " & Format(Me.txtStartTime, "h:nn ampm") & " on " & Me.txtStartDate & " - PTO"
'    ElseIf Me.txtEndDate > Me.txtStartDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime < TimeValue("23:59") And Me.PaidYN = True Then
'        'Debug.Print "OutMultiDaysReturningMidDayPaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be out from " & Me.txtStartDate & " to " & Me.txtEndDate & " at " & Format(Me.txtEndTime, "h:nn ampm") & " - PTO"
'    ElseIf Me.txtEndDate > Me.txtStartDate And Me.txtStartTime > TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = True Then
'        'Debug.Print "OutMultiDaysLeavingEarlyPaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be leaving at " & Format(Me.txtStartTime, "h:nn ampm") & " on " & Me.txtStartDate & " out through " & Me.txtEndDate & " - PTO"
'    ElseIf Me.txtEndDate > Me.txtStartDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = True Then
'        'Debug.Print "OutMultipleFullDaysPaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be out " & Me.txtStartDate & " to " & Me.txtEndDate & " - PTO"
'    ElseIf Me.txtStartDate = Me.txtEndDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = True Then
'        'Debug.Print "OneFullDayOutPaid"
'        Me.txtSubject = Me.EmployeeName & " will be out " & Me.txtStartDate & " - PTO"
'    ElseIf Me.txtStartDate = Me.txtEndDate And Me.txtStartTime > TimeValue("0:00") And Me.txtEndTime < TimeValue("23:59") And Me.PaidYN = True Then
'        'Debug.Print "OneDayLeaveAndReturnPaid"
'        Me.txtSubject = Me.EmployeeName & " will be leaving at " & Format(Me.txtStartTime, "h:nn ampm") & " and returning at " & Format(Me.txtEndTime, "h:nn ampm") & " on " & Me.txtStartDate & " - PTO"
'    ElseIf Me.txtStartDate = Me.txtEndDate And Me.txtStartTime > TimeValue("0:00") And Me.txtEndTime < TimeValue("23:59") And Me.PaidYN = False Then
'        'Debug.Print "OneDayLeaveAndReturnUnPaid"
'        Me.txtSubject = Me.EmployeeName & " will be leaving at " & Format(Me.txtStartTime, "h:nn ampm") & " and returning at " & Format(Me.txtEndTime, "h:nn ampm") & " on " & Me.txtStartDate & " - No PTO"
'    End If
'End Sub


'Private Sub txtFileName_AfterUpdate()
'    If Len(Me.txtFileName) > 0 Then
'        Me.txtAttachments = "\\TLS-FILE\TUL Administrative\PTO Request Forms\" & Me.txtFileName & ".pdf"
'    End If
'End Sub


'Private Sub txtMeetingID_Change()
'    If IsNull(Me.txtEmployeeEmail) Or Me.txtEmployeeEmail = "" Then
'        Me.txtTo = Nz(Me.txtForemanEmail, "")
'    Else
'        If IsNull(Me.txtForemanEmail) Or Me.txtForemanEmail = "" Then
'            Me.txtTo = ""
'        ElseIf Me.txtEmployeeEmail = "Jfissel@tlsokc.com" Or Me.txtEmployeeEmail = "Rjimenez@tlsokc.com" Or Me.txtEmployeeEmail = "Ldelgado@tlsokc.com" Then
'            Me.txtTo = "JFissel@tlsokc.com;RJimenez@tlsokc.com;LDelgado@tlsokc.com"
'        ElseIf Me.txtEmployeeEmail = "efrench@tlsokc.com" Or Me.txtEmployeeEmail = "jmudge@tlsokc.com" Or Me.txtEmployeeEmail = "sstrode@tlsokc.com" Or Me.txtEmployeeEmail = "alittlefield@tlsokc.com" _
'            Or Me.txtEmployeeEmail = "tgreen@tlsokc.com" Or Me.txtEmployeeEmail = "scruz@tlsokc.com" Or Me.txtEmployeeEmail = "jguerrero@tlsokc.com" Or Me.txtEmployeeEmail = "nvandalsem@tlsokc.com" Or Me.txtEmployeeEmail = "nvandalsem@tlsokc.com" Then
'            Me.txtTo = "TLS Mechanic Shop"
'        ElseIf Me.txtEmployeeEmail = Me.txtForemanEmail Then
'            Me.txtTo = Me.txtEmployeeEmail
'        Else
'            Me.txtTo = Me.txtEmployeeEmail & ";" & Me.txtForemanEmail
'        End If
'    End If
'End Sub

'Private Sub txtStartDate_AfterUpdate()
'    Me.txtEndDate = Me.txtStartDate
'End Sub

'Private Sub txtStartTime_AfterUpdate()
'    Me.txtSubject.Visible = True

'    If Me.txtStartDate = Me.txtEndDate And Me.txtStartTime > TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = False Then
'        'Debug.Print "OneDayLeavingEarlyUnpaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be leaving at " & Format(Me.txtStartTime, "h:nn ampm") & " on " & Me.txtStartDate & " - No PTO"
'    ElseIf Me.txtStartDate = Me.txtEndDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime < TimeValue("23:59") And Me.PaidYN = False Then
'        'Debug.Print "OneDayArrivingLateUnpaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be coming in at " & Format(Me.txtEndTime, "h:nn ampm") & " on " & Me.txtStartDate & " - No PTO"
'    ElseIf Me.txtEndDate > Me.txtStartDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime < TimeValue("23:59") And Me.PaidYN = False Then
'        'Debug.Print "OutMultiDaysReturningMidDayUnpaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be out from " & Me.txtStartDate & " to " & Me.txtEndDate & " at " & Format(Me.txtEndTime, "h:nn ampm") & " - No PTO"
'    ElseIf Me.txtEndDate > Me.txtStartDate And Me.txtStartTime > TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = False Then
'        'Debug.Print "OutMultiDaysLeavingEarlyUnPaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be leaving at " & Format(Me.txtStartTime, "h:nn ampm") & " on " & Me.txtStartDate & " out through " & Me.txtEndDate & " - No PTO"
'    ElseIf Me.txtEndDate > Me.txtStartDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = False Then
'        'Debug.Print "OutMultipleFullDaysUnPaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be out " & Me.txtStartDate & " to " & Me.txtEndDate & " - No PTO"
'    ElseIf Me.txtStartDate = Me.txtEndDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = False Then
'        'Debug.Print "OneFullDayOutUnPaid"
'        Me.txtSubject = Me.EmployeeName & " will be out " & Me.txtStartDate & " - No PTO"
'    ElseIf Me.txtStartDate = Me.txtEndDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime < TimeValue("23:59") And Me.PaidYN = True Then
'        'Debug.Print "OneDayArrivingLatePaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be coming in at " & Format(Me.txtEndTime, "h:nn ampm") & " on " & Me.txtStartDate & " - PTO"
'    ElseIf Me.txtStartDate = Me.txtEndDate And Me.txtStartTime > TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = True Then
'        'Debug.Print "OneDayLeavingEarlyPaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be leaving at " & Format(Me.txtStartTime, "h:nn ampm") & " on " & Me.txtStartDate & " - PTO"
'    ElseIf Me.txtEndDate > Me.txtStartDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime < TimeValue("23:59") And Me.PaidYN = True Then
'        'Debug.Print "OutMultiDaysReturningMidDayPaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be out from " & Me.txtStartDate & " to " & Me.txtEndDate & " at " & Format(Me.txtEndTime, "h:nn ampm") & " - PTO"
'    ElseIf Me.txtEndDate > Me.txtStartDate And Me.txtStartTime > TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = True Then
'        'Debug.Print "OutMultiDaysLeavingEarlyPaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be leaving at " & Format(Me.txtStartTime, "h:nn ampm") & " on " & Me.txtStartDate & " out through " & Me.txtEndDate & " - PTO"
'    ElseIf Me.txtEndDate > Me.txtStartDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = True Then
'        'Debug.Print "OutMultipleFullDaysPaid"
'        Me.txtSubject = Me.txtEmployeeName & " will be out " & Me.txtStartDate & " to " & Me.txtEndDate & " - PTO"
'    ElseIf Me.txtStartDate = Me.txtEndDate And Me.txtStartTime = TimeValue("0:00") And Me.txtEndTime = TimeValue("23:59") And Me.PaidYN = True Then
'        'Debug.Print "OneFullDayOutPaid"
'        Me.txtSubject = Me.EmployeeName & " will be out " & Me.txtStartDate & " - PTO"
'    ElseIf Me.txtStartDate = Me.txtEndDate And Me.txtStartTime > TimeValue("0:00") And Me.txtEndTime < TimeValue("23:59") And Me.PaidYN = True Then
'        'Debug.Print "OneDayLeaveAndReturnPaid"
'        Me.txtSubject = Me.EmployeeName & " will be leaving at " & Format(Me.txtStartTime, "h:nn ampm") & " and returning at " & Format(Me.txtEndTime, "h:nn ampm") & " on " & Me.txtStartDate & " - PTO"
'    ElseIf Me.txtStartDate = Me.txtEndDate And Me.txtStartTime > TimeValue("0:00") And Me.txtEndTime < TimeValue("23:59") And Me.PaidYN = False Then
'        'Debug.Print "OneDayLeaveAndReturnUnPaid"
'        Me.txtSubject = Me.EmployeeName & " will be leaving at " & Format(Me.txtStartTime, "h:nn ampm") & " and returning at " & Format(Me.txtEndTime, "h:nn ampm") & " on " & Me.txtStartDate & " - No PTO"
'End If
'End Sub
