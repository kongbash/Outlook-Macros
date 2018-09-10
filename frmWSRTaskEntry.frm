VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmWSRTaskEntry 
   Caption         =   "WSR Task Entry"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7680
   OleObjectBlob   =   "frmWSRTaskEntry.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmWSRTaskEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Constants for column name in excel
Private Const xlDateColumn As String = "A"
Private Const xlTaskDescriptionColumn As String = "B"
Private Const xlProjectNameColumn As String = "C"
Private Const xlTaskForColumn As String = "D"

Private Sub cmdAdd_Click()
    Dim lRowCounter As Long
    Dim sComments As String
    Dim boolCommentAdded As Boolean
    Dim oWorkbook As Excel.Workbook
    
    If CheckControls = True Then
        Set oWorkbook = Excel.Workbooks.Open(sWSRFilePath)
        lRowCounter = GetLastRow(oWorkbook) + 1
        
        With oWorkbook
            .ActiveSheet.Range(xlDateColumn & Trim$(Str$(lRowCounter))).Value = txtDate.Text
            .ActiveSheet.Range(xlTaskDescriptionColumn & Trim$(Str$(lRowCounter))).Value = txtTaskDescription.Text
            .ActiveSheet.Range(xlProjectNameColumn & Trim$(Str$(lRowCounter))).Value = cmbProjectName.Value
            
            If optWSR.Value = 1 Then
                .ActiveSheet.Range(xlTaskForColumn & Trim$(Str$(lRowCounter))).Value = "WSR"
                
            ElseIf optTimeSheet.Value = 1 Then
                .ActiveSheet.Range(xlTaskForColumn & Trim$(Str$(lRowCounter))).Value = "TIMESHEET"
                
            Else
                .ActiveSheet.Range(xlTaskForColumn & Trim$(Str$(lRowCounter))).Value = "BOTH"
                
            End If
            
            'Saving and Closing the Excel workbook
            .Save
            .Close
            
            'Closing the form
            frmWSRTaskEntry.Hide
            
        End With
    End If
    
    Set oWorkbook = Nothing
    
End Sub

Private Sub cmdCancel_Click()
    frmWSRTaskEntry.Hide
End Sub

Private Sub UserForm_Activate()
    txtTaskDescription.Text = vbNullString
End Sub

Private Sub UserForm_Initialize()
    txtDate.Text = Format$(Now(), "dd-mmm-yyyy")
    
    cmbProjectName.AddItem "Nets", 0
    cmbProjectName.AddItem "ICICI", 1
    
    cmbProjectName.ListIndex = 0
    
    optWSR.Value = 1
End Sub

Private Function CheckControls() As Boolean
    Dim boolAreAllValuesEnteredCorrectly As Boolean
    
    boolAreAllValuesEnteredCorrectly = True
    
    If Trim$(txtDate.Text) = vbNullString Then
        boolAreAllValuesEnteredCorrectly = False
        MsgBox "Please enter a date!!", vbOKOnly + vbInformation, "C-SAM Solutions"
        
    ElseIf Trim$(txtTaskDescription.Text) = vbNullString Then
        boolAreAllValuesEnteredCorrectly = False
        MsgBox "Please enter task description!!", vbOKOnly + vbInformation, "C-SAM Solutions"
        
    ElseIf (optWSR.Value = 0 And optTimeSheet.Value = 0 And optBoth.Value = 0) Then
        boolAreAllValuesEnteredCorrectly = False
        MsgBox "Please select where this task should be reported!!", vbOKOnly + vbInformation, "C-SAM Solutions"
    
    End If
    
    CheckControls = boolAreAllValuesEnteredCorrectly
    
End Function
