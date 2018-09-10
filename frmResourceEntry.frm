VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmResourceEntry 
   Caption         =   "Resource Review Entry"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5730
   OleObjectBlob   =   "frmResourceEntry.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmResourceEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Const xlResourceColumn As String = "A"
Private Const xlProjectNameColumn As String = "B"
Private Const xlPCRNumberColumn As String = "C"
Private Const xlActualQAReleaseDate As String = "G"
Private Const xlActualUATReleaseDate As String = "I"
Private Const xlBlockerBugColumn As String = "J"
Private Const xlMajorBugColumn As String = "K"
Private Const xlMinorBugColumn As String = "L"
Private Const xlTrivalBugColumn As String = "M"
Private Const xlUATBugColumn As String = "N"
Private Const xlCommentsColumn As String = "AA"

Dim oWorkbook As Excel.Workbook

Private Sub cmdCancel_Click()
    frmResourceEntry.Hide
End Sub

Private Sub cmdOk_Click()
    
    On Error GoTo err_cmdOk_Click
    
    Dim lCtr As Long
    Dim sComments As String
    Dim boolDataAdded As Boolean
    
    boolDataAdded = False
    lCtr = 4

    With oWorkbook.ActiveSheet
        While .Range(xlResourceColumn & Trim$(Str$(lCtr))).Value <> vbNullString
            If Trim$(.Range(xlResourceColumn & Trim$(Str$(lCtr))).Value) = cmbResourceName.Text Then
                If Trim$(.Range(xlProjectNameColumn & Trim$(Str$(lCtr))).Value) = cmbProjectName.Text Then
                    If Trim$(.Range(xlPCRNameColumn & Trim$(Str$(lCtr))).Value) = UCase$(txtPCRNumber.Text) Then
                        
                        'Updating Actual Release Dates
                        If Trim$(txtActualQAReleaseDate.Text) <> vbNullString Then
                            .Range(xlActualQAReleaseDate & Trim$(Str$(lCtr))).Value = Format$(txtActualQAReleaseDate.Text, "dd-mmm-yyyy")
                        End If
                        
                        If Trim$(txtActualUATReleaseDate.Text) <> vbNullString Then
                            .Range(xlActualUATReleaseDate & Trim$(Str$(lCtr))).Value = Format$(txtActualUATReleaseDate.Text, "dd-mmm-yyyy")
                        End If
                        
                        'Updating bugs data
                        .Range(xlBlockerBugColumn & Trim$(Str$(lCtr))).Value = _
                            Val(Trim$(.Range(xlBlockerBugColumn & Trim$(Str$(lCtr))).Value)) + Val(Trim$(txtBlockerBug.Text))
                        .Range(xlMajorBugColumn & Trim$(Str$(lCtr))).Value = _
                            Val(Trim$(.Range(xlMajorBugColumn & Trim$(Str$(lCtr))).Value)) + Val(Trim$(txtMajorBug.Text))
                        .Range(xlMinorBugColumn & Trim$(Str$(lCtr))).Value = _
                            Val(Trim$(.Range(xlMinorBugColumn & Trim$(Str$(lCtr))).Value)) + Val(Trim$(txtMinorBug.Text))
                        .Range(xlTrivalBugColumn & Trim$(Str$(lCtr))).Value = _
                            Val(Trim$(.Range(xlTrivalBugColumn & Trim$(Str$(lCtr))).Value)) + Val(Trim$(txtTrivial.Text))
                        .Range(xlUATBugColumn & Trim$(Str$(lCtr))).Value = _
                            Val(Trim$(.Range(xlUATBugColumn & Trim$(Str$(lCtr))).Value)) + Val(Trim$(txtUAT.Text))
                            
                        'Updating Comments
                        sComments = Trim$(Range(xlCommentsColumn & Trim$(Str$(lCtr))).Value)
                        
                        'Adding a line feed if some comments already exists
                        If Trim$(sComments) <> vbNullString Then
                            sComments = sComments & Chr(10) & Chr(10) & txtDate.Text & ":" & Chr(10) & Trim$(txtRemarks.Text)
                        Else
                            sComments = sComments & txtDate.Text & ":" & Chr(10) & Trim$(txtRemarks.Text)
                        End If
                        
                        .Range(xlCommentsColumn & Trim$(Str$(lCtr))).Value = sComments
                        
                        'Setting the flag to true
                        boolDataAdded = True
                        
                        'exiting the loop cause the data is added.
                        GoTo exit_for_loop
                    End If 'PCR Name
                End If 'Project Name
            End If 'Resource Name
    
            lCtr = lCtr + 1
        Wend
    End With

exit_for_loop:

    If boolDataAdded = False Then
        MsgBox "Comment not added!!!", vbOKOnly, "C-SAM Solutions"

    Else
        frmResourceEntry.Hide
        MsgBox "Comment added successfully!!!", vbOKOnly, "C-SAM Solutions"

        oWorkbook.Save
        oWorkbook.Close
        
        Set oWorkbook = Nothing
    End If
    
    Exit Sub
    
err_cmdOk_Click:
    MsgBox "Error Source: " & Err.Source & vbCrLf & _
        "Error Description: " & Err.DESCRIPTION & vbCrLf & _
        "Error Number: " & Err.Number, vbOKOnly, "C-SAM Solutions"
End Sub

Private Sub cmdSearchPCRDetails_Click()
    
    On Error GoTo err_cmdSearchPCRDetails_Click
    
    If FetchPCRRecord = False Then
        lblConfirmation.Caption = "No Records Found!!!"
    End If
    
    Exit Sub
    
err_cmdSearchPCRDetails_Click:
    MsgBox "Error Source: " & Err.Source & vbCrLf & _
        "Error Description: " & Err.DESCRIPTION & vbCrLf & _
        "Error Number: " & Err.Number, vbOKOnly, "C-SAM Solutions"
End Sub

Private Sub txtBlockerBug_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = isNumericValues(KeyAscii)
End Sub

Private Sub txtMajorBug_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = isNumericValues(KeyAscii)
End Sub

Private Sub txtMinorBug_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = isNumericValues(KeyAscii)
End Sub

Private Sub txtTrivial_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = isNumericValues(KeyAscii)
End Sub

Private Sub txtUAT_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = isNumericValues(KeyAscii)
End Sub

Private Sub UserForm_Activate()
    Clear_Controls
    
    'Populating Project Names
    cmbProjectName.AddItem "ICICI", 0
    cmbProjectName.AddItem "Nets", 1
    
    cmbProjectName.ListIndex = 0
    
    'Populating Resource Names
    cmbResourceName.AddItem "Pritam P.", 0
    cmbResourceName.AddItem "Mangesh Y.", 1
    cmbResourceName.AddItem "Juned A.", 2
    cmbResourceName.AddItem "Vishal S.", 3
    cmbResourceName.AddItem "Lalit P.", 4
    cmbResourceName.AddItem "Priti S.", 5
    cmbResourceName.AddItem "Sumeet P.", 6
    cmbResourceName.AddItem "Samip S.", 7
    
    cmbResourceName.ListIndex = 0
    
    'Populating Today's date
    txtDate.Text = Format$(Now(), "dd-mmm-yyyy")
    
    'Opening the review sheet
    Set oWorkbook = Excel.Workbooks.Open("D:\Amar\Office\4. Review\Review-Tracking-Sheet.xlsx")
End Sub

Private Sub Clear_Controls()
    cmbProjectName.Clear
    cmbResourceName.Clear
    txtPCRNumber.Text = "PCR-"
    txtDate.Text = vbNullString
    txtRemarks.Text = vbNullString
    txtBlockerBug.Text = vbNullString
    txtMajorBug.Text = vbNullString
    txtMinorBug.Text = vbNullString
    txtTrivial.Text = vbNullString
    txtUAT.Text = vbNullString
End Sub

Private Function isNumericValues(ByVal KeyAscii As Integer) As Integer
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        isNumericValues = KeyAscii
    Else
        isNumericValues = 0
        MsgBox "Please enter numbers only!!", vbOKOnly, "C-SAM Solutions"
    End If
End Function

Private Function FetchPCRRecord() As Boolean
    
    
    On Error GoTo err_FetchPCRRecord
    
    Dim lCtr As Long 'Used as counter in the While Loop used for searching records in the excel sheet.
    Dim boolRecordFound As Boolean 'Declaring this variable to set the message in the lable. As a confirmation if the record exist in the excel sheet or not.
    
    boolRecordFound = False 'Since the record is not found as yet.
    lCtr = 4
    
    With oWorkbook.ActiveSheet
        While .Range(xlResourceColumn & Trim$(Str$(lCtr))).Value <> vbNullString
            If Trim$(.Range(xlResourceColumn & Trim$(Str$(lCtr))).Value) = cmbResourceName.Text Then
                If Trim$(.Range(xlProjectNameColumn & Trim$(Str$(lCtr))).Value) = cmbProjectName.Text Then
                    If Trim$(.Range(xlPCRNumberColumn & Trim$(Str$(lCtr))).Value) = UCase$(txtPCRNumber.Text) Then
                        
                        'Fetching the actual release dates
                        txtActualQAReleaseDate.Text = Format$(.Range(xlActualQAReleaseDate & Trim$(Str$(lCtr))).Value, "dd-mmm-yyyy")
                        txtActualUATReleaseDate.Text = Format$(.Range(xlActualUATReleaseDate & Trim$(Str$(lCtr))).Value, "dd-mmm-yyyy")
                        
                        'Fetching the data of the bug count
                        txtBlockerBug.Text = .Range(xlBlockerBugColumn & Trim$(Str$(lCtr))).Value
                        txtMajorBug.Text = .Range(xlMajorBugColumn & Trim$(Str$(lCtr))).Value
                        txtMinorBug.Text = .Range(xlMinorBugColumn & Trim$(Str$(lCtr))).Value
                        txtTrivial.Text = .Range(xlTrivalBugColumn & Trim$(Str$(lCtr))).Value
                        txtUAT.Text = .Range(xlUATBugColumn & Trim$(Str$(lCtr))).Value
                        
                        'Since the record is found!!! :)
                        boolRecordFound = True
                        
                        'Ending the while loop cause the required record is found
                        GoTo end_while_loop
                    End If 'PCR Name
                End If 'Project Name
            End If 'Resource Name
            
            lCtr = lCtr + 1
        Wend
    End With
    
end_while_loop:

    FetchPCRRecord = boolRecordFound
    
    Exit Function
    
err_FetchPCRRecord:
    MsgBox "Error Source: " & Err.Source & vbCrLf & _
        "Error Description: " & Err.DESCRIPTION & vbCrLf & _
        "Error Number: " & Err.Number, vbOKOnly, "C-SAM Solutions"

End Function

