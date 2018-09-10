VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub AddTask()
    'Outlook variables
    Dim olTask As Outlook.TaskItem
    Dim olItem As Object
    Dim olExp As Outlook.Explorer
    Dim fldCurrent As Outlook.MAPIFolder
    Dim olApp As Outlook.Application
    
    Set olApp = New Outlook.Application
    Set olTask = olApp.CreateItem(olTaskItem)
    Set olExp = olApp.ActiveExplorer
    
    Dim cntSelection As Integer
    cntSelection = olExp.Selection.Count
    
    For i = 1 To cntSelection
        Set olItem = olExp.Selection.Item(i)
        olTask.Body = Format$(Now(), "dd-mmm-yyyy") & vbCrLf
        olTask.Attachments.Add olItem
        olTask.Subject = olItem.ConversationTopic
    Next
    
    olTask.Display
    
    Set olApp = Nothing
    Set olTask = Nothing
    Set olExp = Nothing
End Sub

