Attribute VB_Name = "modMeetingInvites"
Public Sub MeetingInviteForQFeedback()
    ''
    ''This function is used for sending Meeting invites for Quarterly feedback.
    ''It reads an excel file (Attendence for daily stand-up.xls) from the desktop (C:\Users\e050078\Desktop)
    ''
    
    On Error GoTo err_MeetingInviteForQuarterlyFeedback
    
    'Outlook variables
    Dim oApp As New Outlook.Application
    Dim oAppointmentItem As Outlook.AppointmentItem
    Dim oRecepient As Outlook.Recipient
    
    ''Excel variables
    Dim oWorkSheet As Excel.Worksheet
    
    ''Variables to prepare email body
    Dim sBodyStart As String, sHTMLBody As String, sTo As String
    Dim iPosition As Integer, iLenghtOfBodyTag As Integer, iCtr As Integer
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''Opening the excel sheet to read name and email id of the resource
    Set oWorkSheet = Excel.Application.Workbooks.Open("C:\Users\e050078\Desktop\review-sheets.xlsx").Sheets(1)
    
    ''Setting the counter as 2, cause the values in the excel sheet are starting from 2nd row
    iCtr = 2
    
    
    While Trim$(oWorkSheet.Range("A" & Trim$(Str$(iCtr))).Value) <> vbNullString
        ''Creating an Appointment Item
        Set oAppointmentItem = oApp.CreateItem(olAppointmentItem)
        
        With oAppointmentItem
            .MeetingStatus = olMeeting
            .Subject = Trim$(oWorkSheet.Range("D1").Value)
            
            sTo = Trim$(oWorkSheet.Range("B" & Trim$(Str$(iCtr))).Value) & "; "
            
            If Trim$(oWorkSheet.Range("C" & Trim$(Str$(iCtr))).Value) <> vbNullString Then
                sTo = sTo & Trim$(oWorkSheet.Range("C" & Trim$(Str$(iCtr))).Value) & "; "
            End If
                        
            Set oRecepient = oAppointmentItem.Recipients.Add(sTo & vbCrLf & _
                "Rosalind Franklin/SE/Pune 3rd Floor <RosalindFranklin.SE.Pune3rdFloor@mastercard.com>; " & vbCrLf & _
                "Caroline Herschel/SE/Pune 3rd Floor <CarolineHerschel/SE/Pune3rdFloor@mastercard.com>; " & vbCrLf & _
                "Ada Lovelace/SE/Pune 3rd Floor <AdaLovelace/SE/Pune3rdFloor@mastercard.com>; " & vbCrLf & _
                "Anandibai Joshi/SE/Pune 3rd Floor <AnandibaiJoshi/SE/Pune3rdFloor@mastercard.com>; " & vbCrLf & _
                "Shakuntala Devi/SE/Pune 3rd Floor <ShakuntalaDevi/SE/Pune3rdFloor@mastercard.com>; " & vbCrLf & _
                "Tessy Thomas /SW/Pune 3rd Floor <TessyThomas/SW/Pune3rdFloor@mastercard.com>; " & vbCrLf & _
                "J.Manjula /NW/Pune 3rd Floor <J.Manjula/NW/Pune3rdFloor@mastercard.com>; " & vbCrLf & _
                "Julia Robinson/NW/Pune 3rd Floor <JuliaRobinson/NW/Pune3rdFloor@mastercard.com>; " & vbCrLf & _
                "Sheryl Sandberg/NW/Pune 3rd Floor <SherylSandberg/NW/Pune3rdFloor@mastercard.com>; " & vbCrLf & _
                "Kiran Shaw/NE/Pune 3rd Floor <KiranShaw/NE/Pune3rdFloor@mastercard.com>; " & vbCrLf & _
                "P.T. Usha/NE/Pune 3rd Floor <P.T.Usha/NE/Pune3rdFloor@mastercard.com>; " & vbCrLf & _
                "Arunima Sinha/NE/Pune 3rd Floor <ArunimaSinha/NE/Pune3rdFloor@mastercard.com>; " & vbCrLf & _
                "Grace Hopper/SE/Pune 3rd Floor <GraceHopper/SE/Pune3rdFloor@mastercard.com>; " & vbCrLf & _
                "Mary Kom/NW/Pune 3rd Floor <MaryKom/NW/Pune3rdFloor@mastercard.com>")
        End With
        
        oAppointmentItem.Recipients.ResolveAll
        oAppointmentItem.Display
        
        ''Destroying the Appointment Item
        Set oAppointmentItem = Nothing
        Set oRecepient = Nothing
        
        iCtr = iCtr + 1
    Wend
    
    Set oWorkSheet = Nothing
    Set oApp = Nothing
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Exit Sub
    
err_MeetingInviteForQuarterlyFeedback:
    MsgBox "Error Number: " & Err.Number & vbCrLf & _
        "Error Description: " & Err.DESCRIPTION & vbCrLf & _
        "Error Source: " & Err.Source, vbOKOnly, "C-SAM Solutions"
End Sub

