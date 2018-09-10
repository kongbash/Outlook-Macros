VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDatePicker 
   Caption         =   "Pick date"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2595
   OleObjectBlob   =   "frmDatePicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim sDate As String

Private Sub mvDatePicker_DateClick(ByVal DateClicked As Date)
    Dim olApp As Outlook.Application
    Dim oMailItem As Outlook.MailItem
    Dim oInspector As Outlook.Inspector
    Dim oWordDoc As Word.Document
    
    Set olApp = New Outlook.Application
    Set oMailItem = olApp.ActiveInspector.CurrentItem   'get the object of the active mail item window in which we need to add the date
    Set oInspector = oMailItem.GetInspector
    
    If oInspector.EditorType = olEditorWord Then
        Set oWordDoc = oInspector.WordEditor
        oWordDoc.Application.Selection.TypeText Format$(DateClicked, "dd-mmm-yyyy")
    End If
    
    frmDatePicker.Hide
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
mvDatePicker.Value = Date
End Sub
