VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSendReviewForm 
   Caption         =   "Send Review"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5265
   OleObjectBlob   =   "frmSendReviewForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSendReviewForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




'Constants for column name in excel
Private Const xlResourceColumn As String = "A"
Private Const xlProjectNameColumn As String = "B"
Private Const xlPCRNameColumn As String = "C"
Private Const xlProjectStartDate As String = "E"
Private Const xlBlockerBugColumn As String = "J"
Private Const xlMajorBugColumn As String = "K"
Private Const xlMinorBugColumn As String = "L"
Private Const xlTrivalBugColumn As String = "M"
Private Const xlUATBugColumn As String = "N"
Private Const xlCommentsColumn As String = "AA"

'Constant for the path where the review file will be saved
Private Const sReviewFilePath As String = "D:\Amar\Office\4. Review\2013-14\Quarterly Review\"

Dim oStartDate As Date
Dim oEndDate As Date

Private Sub cmdCancel_Click()
    frmSendReviewForm.Hide
    Set frmSendReviewForm = Nothing
End Sub

Private Sub cmdOk_Click()
    
    On Error GoTo err_cmdOk_Click
    
    Dim iCtr As Integer
    Dim sComments As String
    Dim boolCommentAdded As Boolean
    Dim oWorkbook As Excel.Workbook
    
    If CheckYear(txtYear.Text) = True Then
    
        'Opening the review tracking sheet
        Set oWorkbook = Excel.Workbooks.Open("D:\Amar\Office\4. Review\Review-Tracking-Sheet.xlsx")
        
        'Setting the start and end date as per the quarter
        Select Case cmbQuarter
        Case "Q1 - April-June":
            oStartDate = CDate("04/01/" & Trim$(txtYear.Text))
            oEndDate = CDate("06/30/" & Trim$(txtYear.Text))
            
        Case "Q1 - July-September":
            oStartDate = CDate("07/01/" & Trim$(txtYear.Text))
            oEndDate = CDate("06/30/" & Trim$(txtYear.Text))
            
        Case "Q1 - October-December":
            oStartDate = CDate("10/01/" & Trim$(txtYear.Text))
            oEndDate = CDate("12/31/" & Trim$(txtYear.Text))
            
        Case "Q1 - January-March":
            oStartDate = CDate("01/01/" & Trim$(txtYear.Text))
            oEndDate = CDate("03/31/" & Trim$(txtYear.Text))
            
        End Select
        
        If cmbResourceName.Text = "All" Then
            For iCtr = 1 To cmbResourceName.ListCount - 1
                lblProgressBar.Caption = "Processing data for " & cmbResourceName.List(iCtr) & " (" & iCtr & "/" & (cmbResourceName.ListCount - 1) & ")"
                
                Set oWorkbook = Workbooks.Open("D:\Amar\Office\4. Review\Review-Tracking-Sheet.xlsx")
                
                If UpdateWorkBook(oWorkbook, cmbResourceName.List(iCtr)) Then
                    SendMail oWorkbook.FullName, cmbResourceName.Column(2, iCtr)
                    
                Else
                    'Deleting the excel file for which record is not generated
                    'Shell "del " & """" & sReviewFilePath & cmbQuarter.Text & "\" & cmbResourceName.List(iCtr) & "- review sheet for " & cmbQuarter.Text & ".xlsx" & """"
                End If
                
                'Closing the workbook here
                oWorkbook.Close
            Next iCtr
        Else
            If UpdateWorkBook(oWorkbook, cmbResourceName.Text) Then
                SendMail oWorkbook.FullName, cmbResourceName.Column(2, cmbResourceName.ListIndex)
            
            Else
                MsgBox "No records found for " & cmbResourceName.Text & "!!!", vbOKOnly, "C-SAM Solutions"
                
                'Deleting the excel file
                'Shell "del " & """" & sReviewFilePath & cmbQuarter.Text & "\" & cmbResourceName.List(iCtr) & "- review sheet for " & cmbQuarter.Text & ".xlsx" & """"
            End If
        End If
        
        MsgBox "Done!!!", vbOKOnly, "C-SAM Solutions"
        lblProgressBar.Caption = vbNullString
        
    End If 'CheckYear
    
    oWorkbook.Close
    Set oWorkbook = Nothing
    
    Exit Sub
    
err_cmdOk_Click:
    MsgBox "Error Number: " & Err.Number & vbCrLf & _
        "Error Description: " & Err.DESCRIPTION & vbCrLf & _
        "Error Source: " & Err.Source, vbOKOnly, "C-SAM Solutions"
End Sub

Private Sub txtYear_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = isNumericValues(KeyAscii)
End Sub

Private Sub UserForm_Initialize()
    'Populating Resource Names
    cmbResourceName.AddItem "All", 0
    cmbResourceName.Column(2, 0) = "All"
    
    cmbResourceName.AddItem "Pritam P.", 1
    cmbResourceName.Column(2, 1) = "pritam.patil@c-sam.com"
    
    cmbResourceName.AddItem "Mangesh Y.", 2
    cmbResourceName.Column(2, 2) = "mangesh.yadav@c-sam.com"
    
    cmbResourceName.AddItem "Juned A.", 3
    cmbResourceName.Column(2, 3) = "juned.ahmed@c-sam.com"
    
    cmbResourceName.AddItem "Vishal S.", 4
    cmbResourceName.Column(2, 4) = "vishal.shelar@c-sam.com"
    
    cmbResourceName.AddItem "Lalit P.", 5
    cmbResourceName.Column(2, 5) = "lalit.patil@c-sam.com"
    
    cmbResourceName.AddItem "Priti S.", 6
    cmbResourceName.Column(2, 6) = "priti.sankhe@c-sam.com"
    
    cmbResourceName.AddItem "Sumeet P.", 7
    cmbResourceName.Column(2, 7) = "sumeet.panchal@c-sam.com"
    
    cmbResourceName.AddItem "Samip S.", 8
    cmbResourceName.Column(2, 8) = "samip.shah@c-sam.com"
    
    cmbResourceName.ListIndex = 0
    
    'Populating Quarter
    cmbQuarter.AddItem "Q1 - April-June", 0
    cmbQuarter.AddItem "Q2 - July-September", 1
    cmbQuarter.AddItem "Q3 - October-December", 2
    cmbQuarter.AddItem "Q4 - January-March", 3
    
    cmbQuarter.ListIndex = 0
    
    'Populating Year
    'Deducting 35 days so that the value of last month is fetched. This is useful in Jan-Mar quarter cause in Jan we need to send the review of last quarter
    txtYear.Text = Format$((Now() - 35), "YYYY")
    
    'Populating comments box with current date
    txtComments.Text = Format$(Now(), "dd-mmm-yyyy") & ": " & vbCr
End Sub

Private Function UpdateWorkBook(ByRef oWorkbook As Excel.Workbook, sResourceName As String) As Boolean
    
    On Error GoTo err_UpdateWorkBook
    
    Dim lCtr As Long
    Dim boolWereRecordsAvailable As Boolean
    
    'Setting the counter at 4 as the data start from 4
    lCtr = 4
    
    'Assuming that no records are available
    boolWereRecordsAvailable = False
    
    'Renaiming the original file with the name of the resource
    oWorkbook.SaveAs sReviewFilePath & cmbQuarter.Text & "\" & sResourceName & "- review sheet for " & cmbQuarter.Text & ".xlsx"
    
    'Searching for the record that does not belong to him/her and deleting it
    With oWorkbook.ActiveSheet
        While .Range(xlResourceColumn & Trim$(Str$(lCtr))).Value <> vbNullString
            If Trim$(.Range(xlResourceColumn & Trim$(Str$(lCtr))).Value) <> sResourceName Then
                Rows(Trim$(Str$(lCtr)) & ":" & Trim$(Str$(lCtr))).Select
                Selection.Delete Shift:=xlUp
                
                'The boolean flag is not touched here cause it is already set to false.
                'Also if it is changed in every loop then it will not give a correct picture. cause if its set to true below then it will be set to false here
                'indicating that no records are found for this person
            ElseIf Trim$(.Range(xlResourceColumn & Trim$(Str$(lCtr))).Value) = sResourceName Then
            
                If CDate(.Range(xlProjectStartDate & Trim$(Str$(lCtr))).Value) >= oStartDate And _
                    CDate(.Range(xlProjectStartDate & Trim$(Str$(lCtr))).Value) <= oEndDate Then
                    'Incrementing the counter only if the record is not found. Cause a row is deleted if the record is found, which will bring the next record up.
                    lCtr = lCtr + 1
                    
                    'Setting boolean flag to true cause a record is found.
                    boolWereRecordsAvailable = True
                    
                Else
                    Rows(Trim$(Str$(lCtr)) & ":" & Trim$(Str$(lCtr))).Select
                    Selection.Delete Shift:=xlUp
                    
                    'Again the boolean flag is not touched here cause it is already set to false.
                End If
                
            End If 'Resource Name
        Wend
    End With
    
    'Saving and closing the excel sheet
    oWorkbook.Save
    'oWorkBook.Close
    
    'Returning the boolean value
    UpdateWorkBook = boolWereRecordsAvailable
    
    Exit Function
    
err_UpdateWorkBook:
    MsgBox "Error Number: " & Err.Number & vbCrLf & _
        "Error Description: " & Err.DESCRIPTION & vbCrLf & _
        "Error Source: " & Err.Source, vbOKOnly, "C-SAM Solutions"
End Function

Private Sub SendMail(sWorkBookNameWithPath As String, sEmailId As String)

    On Error GoTo err_SendMail
    
    'Outlook variables
    Dim oApp As New Outlook.Application
    Dim oMailItem As Outlook.MailItem
    
    Set oMailItem = oApp.CreateItem(olMailItem)
    
    With oMailItem
        .To = sEmailId
        .Subject = "Review sheet for " & cmbQuarter.Text
        .Attachments.Add sWorkBookNameWithPath
        .Body = "Please go through the sheet and let me know in case of any issues"
        .Send
    End With
    
    Set oMailItem = Nothing
    Set oApp = Nothing
    
    Exit Sub
    
err_SendMail:
    MsgBox "Error Number: " & Err.Number & vbCrLf & _
        "Error Description: " & Err.DESCRIPTION & vbCrLf & _
        "Error Source: " & Err.Source, vbOKOnly, "C-SAM Solutions"
End Sub

Private Function isNumericValues(ByVal KeyAscii As Integer) As Integer
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        isNumericValues = KeyAscii
    Else
        isNumericValues = 0
        MsgBox "Please enter numbers only!!", vbOKOnly, "C-SAM Solutions"
    End If
End Function

Private Function CheckYear(sYear As String) As Boolean

    Dim boolIsValidYear As Boolean
    
    boolIsValidYear = True
    
    If Len(sYear) < 4 Then
        MsgBox "Invalid Year!!!", vbOKOnly, "C-SAM Solutions"
        boolIsValidYear = False
        
    ElseIf Val(sYear) < 2013 Then
        MsgBox "Years are supported from 2013 only!!!", vbOKOnly, "C-SAM Solutions"
        boolIsValidYear = False
        
    ElseIf Val(sYear) > Val(Format$(Now(), "yyyy")) Then
        MsgBox "Please enter correct year!!!", vbOKOnly, "C-SAM Solutions"
        boolIsValidYear = False
        
    End If
    
    CheckYear = boolIsValidYear
    
End Function
