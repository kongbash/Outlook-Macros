VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddOrUpdatePCRDetails 
   Caption         =   "Add / Update Project Details"
   ClientHeight    =   8940
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5880
   OleObjectBlob   =   "frmAddOrUpdatePCRDetails.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAddOrUpdatePCRDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'ADODB Connection
'The database is placed at D:\Amar\Office\1. Projects\4. Project Management\Central Database

Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

'Constants for column name in excel
Private Const xlResourceColumn As String = "A"
Private Const xlProjectNameColumn As String = "B"
Private Const xlPCRNumberColumn As String = "C"
Private Const xlPCRName As String = "D"
Private Const xlProjectStartDate As String = "E"
Private Const xlPlannedQAReleaseDate As String = "F"
Private Const xlActualQAReleaseDate As String = "G"
Private Const xlPlannedUATReleaseDate As String = "H"
Private Const xlActualUATReleaseDate As String = "I"
Private Const xlCommentsColumn As String = "AA"

Dim oWorkbook As Excel.Workbook

Private Sub cmbPCRName_Change()

    If Trim$(UCase$(cmbPCRName.Text)) = "ADD NEW PCR" Then
        txtNewPCRNo.Enabled = True
        
    ElseIf Trim$(cmbPCRName.Text) <> vbNullString Then
        txtNewPCRNo.Enabled = False
        
        'Populating PCR Name dropdown box
        iCtr = 0
        Set rs = con.Execute("SELECT PCR_Name, Description, Planned_Start_Date, Planned_QA_Release_Date, Planned_UAT_Release_Date " & _
            "FROM PCR_Master WHERE PCR_NO = '" & cmbPCRName.Text & "'")
            
        With rs
            txtPCRName.Text = Trim$(.Fields("PCR_Name"))
            txtStartDate.Text = Trim$(Format$(.Fields("Planned_Start_Date"), "dd-mmm-yyyy"))
            txtPlannedQAReleaseDate.Text = Trim$(Format$(.Fields("Planned_QA_Release_Date"), "dd-mmm-yyyy"))
            txtPlannedUATReleaseDate.Text = Trim$(Format$(.Fields("Planned_UAT_Release_Date"), "dd-mmm-yyyy"))
            txtDescription.Text = Trim$(.Fields("Description"))
                
            .Close
        End With
    End If
End Sub

Private Sub cmbProjectName_Change()
    PopulatePCRNo
    ClearControls
    
    txtNewPCRNo.Enabled = False
End Sub

Private Sub cmdCancel_Click()
    frmAddOrUpdatePCRDetails.Hide
    Set frmAddOrUpdatePCRDetails = Nothing
End Sub

Private Sub cmdOk_Click()
    On Error GoTo err_cmdOk_Click
    
    Dim lCtr As Long 'Used as counter in the While Loop used for searching records in the excel sheet.
    Dim boolRecordFound As Boolean 'Declaring this variable to set the message in the lable. As a confirmation if the record exist in the excel sheet or not.
    Dim sComments As String
    
    boolRecordFound = False 'Since the record is not found as yet.
    lCtr = 4
    
    With oWorkbook.ActiveSheet
        While .Range(xlResourceColumn & Trim$(Str$(lCtr))).Value <> vbNullString
            lCtr = lCtr + 1
        Wend
        
        'Adding 1 to the row counter to move 1 row below the last one
        'lCtr = lCtr + 1
        
        'Adding the new record on the new row
        .Range(xlResourceColumn & Trim$(Str$(lCtr))).Value = cmbResourceName.Text
        .Range(xlProjectNameColumn & Trim$(Str$(lCtr))).Value = cmbProjectName.Text
        .Range(xlPCRNumberColumn & Trim$(Str$(lCtr))).Value = txtPCRNumber.Text
        .Range(xlPCRName & Trim$(Str$(lCtr))).Value = txtPCRName.Text
        
        .Range(xlProjectStartDate & Trim$(Str$(lCtr))).Value = Format$(txtStartDate.Text, "dd-mmm-yyyy")
        .Range(xlPCRName & Trim$(Str$(lCtr))).Value = txtPCRName.Text
        
        .Range(xlPlannedQAReleaseDate & Trim$(Str$(lCtr))).Value = Format$(txtPlannedQAReleaseDate.Text, "dd-mmm-yyyy")
        .Range(xlPlannedUATReleaseDate & Trim$(Str$(lCtr))).Value = Format$(txtPlannedUATReleaseDate.Text, "dd-mmm-yyyy")
        
        'Adding a line feed if some comments already exists
        sComments = .Range(xlCommentsColumn & Trim$(Str$(lCtr))).Value
        
        If Trim$(sComments) <> vbNullString Then
            sComments = sComments & Chr(10) & Chr(10) & Trim$(txtRemarks.Text)
        Else
            sComments = Trim$(txtRemarks.Text)
        End If
        
        .Range(xlCommentsColumn & Trim$(Str$(lCtr))).Value = sComments
        
        oWorkbook.Save
    End With
    
    MsgBox "Details added successfully!!!", vbOKOnly, "C-SAM Solutions"
    ClearControls
    
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

Private Sub txtPCRName_Change()
    lblConfirmation.Caption = vbNullString
End Sub

Private Sub UserForm_Initialize()
    
    Dim iCtr As Integer
    
    'Initializing the database connection
    con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Amar\Office\1. Projects\4. Project Management\Central Database\C-SAM.accdb;" & _
        "Persist Security Info=False;"

    con.Open
    
    'Setting the cursor location as client so that record count property can be accessed
    rs.CursorLocation = adUseClient
    
    'Populating Project Name
    iCtr = 0
    Set rs = con.Execute("SELECT Project_ID, ProjectName, Active FROM ProjectMaster")
    
    With rs
        While .EOF <> True
            cmbProjectName.AddItem .Fields(PROJECTNAME), iCtr
            cmbProjectName.Column(2, iCtr) = .Fields(PROJECT_ID)
            
            .MoveNext
            iCtr = iCtr + 1
        Wend
        
        cmbProjectName.ListIndex = 0
        .Close
    End With
    
    Set rs = Nothing
    
    'Populating Resource Name dropdown box
    iCtr = 0
    Set rs = con.Execute("SELECT Resource_ID, FirstName, LastName, Email_id FROM TeamMembers")
    
    With rs
        While rs.EOF <> True
            lstResourceName.AddItem .Fields(FIRSTNAME) & " " & .Fields(LASTNAME), iCtr
            lstResourceName.Column(2, iCtr) = .Fields(EMAIL_ID)
            
            .MoveNext
            iCtr = iCtr + 1
        Wend
        
        lstResourceName.ListIndex = 0
        .Close
    End With
    
    Set rs = Nothing
    
    'Populating PCR Name dropdown box
    PopulatePCRNo
    
    Set rs = Nothing
    
    'Opening the review sheet
    Set oWorkbook = Excel.Workbooks.Open("D:\Amar\Office\4. Review\Review-Tracking-Sheet.xlsx")
End Sub

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
                        txtStartDate.Text = Format$(.Range(xlProjectStartDate & Trim$(Str$(lCtr))).Value, "dd-mmm-yyyy")
                        txtPCRName.Text = .Range(xlPCRName & Trim$(Str$(lCtr))).Value
                        
                        txtPlannedQAReleaseDate.Text = Format$(.Range(xlPlannedQAReleaseDate & Trim$(Str$(lCtr))).Value, "dd-mmm-yyyy")
                        txtPlannedUATReleaseDate.Text = Format$(.Range(xlPlannedUATReleaseDate & Trim$(Str$(lCtr))).Value, "dd-mmm-yyyy")
                        
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

Private Sub UserForm_Terminate()
    'closing the excel workbook and releasing the memory
    oWorkbook.Close
    Set oWorkbook = Nothing
    
    'Closing the database connection
    con.Close
    Set con = Nothing
End Sub

Private Sub ClearControls()
    'Clearing values from all the controls
    txtNewPCRNo.Text = "PCR-"
    txtPCRName.Text = vbNullString
    txtStartDate.Text = vbNullString
    txtPlannedQAReleaseDate.Text = vbNullString
    txtPlannedUATReleaseDate.Text = vbNullString
    txtDescription.Text = vbNullString
    
    lblConfirmation.Caption = vbNullString
    
End Sub

Private Sub PopulatePCRNo()
    Dim iCtr As Integer
    
    iCtr = 0
    
    cmbPCRName.Clear
    Set rs = con.Execute("SELECT PCR_ID, Project_ID, PCR_No, PCR_Name FROM PCR_Master WHERE Project_ID = " & cmbProjectName.Column(2, cmbProjectName.ListIndex))
    
    With rs
        While rs.EOF <> True
            cmbPCRName.AddItem .Fields(PCR_NO), iCtr
            cmbPCRName.Column(2, iCtr) = .Fields(PCR_ID)
            
            .MoveNext
            iCtr = iCtr + 1
        Wend
        
        cmbPCRName.AddItem "Add New PCR", iCtr
        cmbPCRName.Column(2, iCtr) = "Add New"
        
        .Close
    End With
End Sub
