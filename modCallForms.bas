Attribute VB_Name = "modCallForms"

Public Sub Call_Review_Form()
    frmResourceEntry.Show
End Sub

Public Sub Call_SendReview_Form()
    frmSendReviewForm.Show
End Sub

Public Sub Call_AddPCRDetails()
    frmAddOrUpdatePCRDetails.Show
End Sub

Public Sub Call_CRImpact()
    frmCRImpact.Show
End Sub

Public Sub Call_DatePicker()
    frmDatePicker.Show
End Sub


Public Sub AddTask()
    'Outlook variables
    Dim olTask As Outlook.TaskItem
    Dim olItem As Object
    Dim olExp As Outlook.Explorer
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
    Set olItem = Nothing
End Sub

Public Sub AddMeetingRequest()
    'Outlook variables
    Dim olMeetingItem As Outlook.AppointmentItem
    Dim olItem As Object
    Dim olExp As Outlook.Explorer
    Dim olApp As Outlook.Application
    
    Set olApp = New Outlook.Application
    Set olMeetingItem = olApp.CreateItem(olAppointmentItem)
    Set olExp = olApp.ActiveExplorer
    
    Dim cntSelection As Integer
    cntSelection = olExp.Selection.Count
    
    For i = 1 To cntSelection
        Set olItem = olExp.Selection.Item(i)
        olMeetingItem.Body = "AT&T Access Code: 8708850"
        olMeetingItem.Body = Format$(Now(), "dd-mmm-yyyy") & vbCrLf
        olMeetingItem.Attachments.Add olItem
        olMeetingItem.Subject = olItem.ConversationTopic
    Next
    
    olMeetingItem.Display
    
    Set olApp = Nothing
    Set olMeetingItem = Nothing
    Set olExp = Nothing
    Set olItem = Nothing
End Sub

Public Sub CreateWSRMail()
    Dim sWeek As String
    Dim sBodyText As String
    Dim sAttachmentPath As String
    
    sWeek = GetWeekStartDate("dd-mmm-yyyy") & " to " & GetWeekEndDate("dd-mmm-yyyy")
    sBodyText = "Hi Arvind / Janna, " & vbCrLf & vbCrLf & "Please find attached the WSR for the week " & sWeek
    sAttachmentPath = "C:\Amar\Office\6. Weekly Reports\WSR\FIGs\WSR_FIGs_" & GetWeekEndDate("yyyymmdd") & ".docx"

    
    PrepareMail "Pekar, Janna <Janna.Pekar@mastercard.com>; Ramamoorthy, Arvind <Arvind.Ramamoorthy@mastercard.com>", _
        "WSR for the week " & sWeek, _
        sBodyText, _
        "", _
        sAttachmentPath

End Sub

Public Sub Taxi_Service()
    Dim sWeek As String
    Dim sBodyText As String
    Dim sAttachmentPath As String
    
    sBodyText = "Hi Pushpendra, " & Chr(13) & _
        "Request you to kindly book a cab for the below employee. " & vbCrLf & _
        "a.      Name and Employee ID  : " & vbCrLf & _
        "b.      Time when drop is required : " & vbCrLf & _
        "c.      Residential address : " & vbCrLf & _
        "d.      Contact Number : " & vbCrLf & vbCrLf & _
        "Please let me know if any more details are required from my end."

    
    PrepareMail "MCBaroda_traveldesk@avis.co.in", _
        "Cab booking for ", _
        sBodyText
End Sub

'Public Sub MyReplyAll()
'    'Outlook variables
'    Dim olItem As Object
'    Dim olExp As Outlook.Explorer
'    Dim olApp As Outlook.Application
'
'    Set olApp = New Outlook.Application
'    Set olExp = olApp.ActiveExplorer
'
'    Dim cntSelection As Integer
'    cntSelection = olExp.Selection.Count
'
'    For i = 1 To cntSelection
'        Set olItem = olExp.Selection.Item(i)
'        olItem.ReplyAll
'    Next
'
'    Set olApp = Nothing
'    Set olExp = Nothing
'    Set olItem = Nothing
'End Sub
