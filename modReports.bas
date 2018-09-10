Attribute VB_Name = "modReports"
''This function sends the WSR report to Sudheendran.
''It extracts the tasks that I save in the excel sheet at the location "E:\Amar\Office\2. Weekly Reports\WSR", and list them project wise
'
'Public Sub SendWSRReport()
'
'    On Error GoTo err_SendWSRReport
'
'    'Constants for column name in excel
'    Const xlDateColumn As String = "A"
'    Const xlTaskDescriptionColumn As String = "B"
'    Const xlProjectNameColumn As String = "C"
'    Const xlTaskForColumn As String = "D"
'
'    Dim lRowCounter As Long
'    Dim dtStartDate As Date, dtEndDate As Date
'    Dim sICICIProjectTasks As String, sNetsProjectTask As String
'    Dim iICICITaskCounter As Integer, iNetsTaskCounter As Integer
'    Dim sSubject As String, sBody As String
'    Dim oWorkBook As Excel.Workbook
'
'    Set oWorkBook = Excel.Workbooks.Open(sWSRFilePath)
'
'    lRowCounter = 2
'    dtStartDate = CDate(GetWeekStartDate)
'    dtEndDate = CDate(GetWeekEndDate)
'    sICICIProjectTasks = "ICICI" & vbCrLf & vbCrLf
'    sNetsProjectTask = "NETS" & vbCrLf & vbCrLf
'    iICICITaskCounter = 1
'    iNetsTaskCounter = 1
'
'    With oWorkBook.ActiveSheet
'        While .Range("A" & Trim$(Str$(lRowCounter))).Value <> vbNullString
'            If CDate(.Range(xlDateColumn & Trim$(Str$(lRowCounter))).Value) >= dtStartDate And _
'                CDate(.Range(xlDateColumn & Trim$(Str$(lRowCounter))).Value) <= dtEndDate Then
'
'                If UCase$(Trim$(.Range(xlTaskForColumn & Trim$(Str$(lRowCounter))).Value)) = "WSR" Or _
'                    UCase$(Trim$(.Range(xlTaskForColumn & Trim$(Str$(lRowCounter))).Value)) = "BOTH" Then
'
'                    If UCase$(Trim$(.Range(xlProjectNameColumn & Trim$(Str$(lRowCounter))).Value)) = "ICICI" Then
'                        sICICIProjectTasks = sICICIProjectTasks & Trim$(Str$(iICICITaskCounter)) & ". " & _
'                            Trim$(.Range(xlTaskDescriptionColumn & Trim$(Str$(lRowCounter))).Value) & vbCrLf
'
'                        iICICITaskCounter = iICICITaskCounter + 1
'
'                    ElseIf UCase$(Trim$(.Range(xlProjectNameColumn & Trim$(Str$(lRowCounter))).Value)) = "NETS" Then
'                        sNetsProjectTask = sNetsProjectTask & Trim$(Str$(iICICITaskCounter)) & ". " & _
'                            Trim$(.Range(xlTaskDescriptionColumn & Trim$(Str$(lRowCounter))).Value) & vbCrLf
'
'                        iNetsTaskCounter = iNetsTaskCounter + 1
'
'                    End If 'Project name check
'                End If 'Task for check
'            End If 'Date check
'
'            lRowCounter = lRowCounter + 1
'        Wend
'
'        'Closing the workbook
'        .Close
'    End With
'
'    'Releasing the memory
'    Set oWorkBook = Nothing
'
'    'Setting the subject of the mail
'    sSubject = "WSR for NETS and ICICI for the week " & GetWeekStartDate & " to " & GetWeekEndDate
'
'    'Setting the body of the email
'    sBody = "Hi Sudheendran, Please find below the WSR for Nets and ICICI projects;" & vbCrLf & _
'        sICICIProjectTasks & vbCrLf & _
'        sNetsProjectTask
'
'    'sending the mail
'    PrepareMail "sudheendran.tl@c-sam.com", sSubject, sBody
'
'    Exit Sub
'
'err_SendWSRReport:
'    MsgBox "Error Number: " & Err.Number & vbCrLf & _
'        "Error Description: " & Err.Description & vbCrLf & _
'        "Error Source: " & Err.Source, vbOKOnly, "C-SAM Solutions"
'End Sub
