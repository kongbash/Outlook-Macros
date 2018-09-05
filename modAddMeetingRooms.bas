Public Sub AddAllMeetingRooms()
    Dim oMailItem As Outlook.MailItem
    Dim oAppointItem As Outlook.AppointmentItem
    Dim oRecepient As Outlook.Recipient
    Dim oObj As Object
    
    Set oObj = Outlook.Application.ActiveInspector.CurrentItem
    'MsgBox oObj.MessageClass
    
    If (oObj.MessageClass = "IPM.Appointment") Then       ' Calendar Entry.
        Set oAppointItem = oObj
        Set oRecepient = oAppointItem.Recipients.Add("meeting-room-1@your-organization.com;" & vbCrLf & _
            "meeting-room-2your-organization.com" & vbCrLf & _
            "meeting-room-3@your-organization.com" & vbCrLf & _
            "meeting-room-4@your-organization.com" & vbCrLf & _
            "meeting-room-n@your-organization.com")
        
        oAppointItem.Recipients.ResolveAll
            
    Else
        MsgBox "You cannot perform this task here.", vbOKOnly + vbInformation, "MasterCard"
        
    End If
End Sub