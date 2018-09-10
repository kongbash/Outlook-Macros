Attribute VB_Name = "modCommonFunctions"
'Constant for the path where the WSR file is saved
Public Const sWSRFilePath As String = "D:\Amar\Office\2. Weekly Reports\WSR\Weekly-Task-Details.xlsm"
Public Const sReviewFileAndPath As String = "D:\Amar\Office\4. Review\Review-Tracking-Sheet.xlsx"

Public Function GetWeekStartDate(sDateFormat As String) As String
    GetWeekStartDate = Format$((Now() - Weekday(Now(), vbUseSystemDayOfWeek) + 2), sDateFormat)
End Function

Public Function GetWeekEndDate(sDateFormat As String) As String
    GetWeekEndDate = Format$((Now() - Weekday(Now(), vbUseSystemDayOfWeek) + 6), sDateFormat)
End Function

'Public Function GetLastRow(oWorkBook As Excel.Workbook) As Long
'    Dim lRowCount As Long
'
'    With oWorkBook
'        If Trim$(.ActiveSheet.Range("A2").Value) <> vbNullString Then
'            lRowCount = .Application.ActiveCell.Row
'
'        Else
'            lRowCount = 2
'        End If
'    End With
'
'    GetLastRow = lRowCount
'End Function

Public Sub PrepareMail(sTo As String, sSubject As String, sBody As String, Optional sCC As String, Optional sAttachment As String)
    'This function is used for sending mails.
    
    On Error GoTo err_SendMail
    
    'Outlook variables
    Dim oApp As New Outlook.Application
    Dim oMailItem As Outlook.MailItem
    
    Dim sBodyStart As String, sHTMLBody As String
    Dim iPosition As Integer, iLenghtOfBodyTag As Integer
    
    'The body tag begins with this tag
    sBodyStart = ":.5in'><div class=WordSection1><p class=MsoNormal><o:p>"
    
    Set oMailItem = oApp.CreateItem(olMailItem)
    
    With oMailItem
        .To = sTo
        
        If Trim$(sCC) <> vbNullString Then
            .CC = sCC
        End If
        
        .Subject = sSubject
        .Display
        
'        sHTMLBody = .HTMLBody
'
'        iLenghtOfBodyTag = Len(sBodyStart)
'        iPosition = (InStr(sHTMLBody, sBodyStart) + iLenghtOfBodyTag)
'
'        If iPosition = iLenghtOfBodyTag Then
'            MsgBox "Body tag not found.", vbOKOnly + vbCritical, "MasterCard"
'        End If
'
'        .HTMLBody = Replace$(sHTMLBody, sBodyStart, sBodyStart & sBody)
        
        If sAttachment <> vbNullString Then
            .Attachments.Add sAttachment
        End If
        
        '.Display
    End With
    
    Set oMailItem = Nothing
    Set oApp = Nothing
    
    Exit Sub
    
err_SendMail:
    MsgBox "Error Number: " & Err.Number & vbCrLf & _
        "Error Description: " & Err.DESCRIPTION & vbCrLf & _
        "Error Source: " & Err.Source, vbOKOnly, "C-SAM Solutions"
End Sub


Public Sub Add_Related_Mail()
    
    On Error GoTo err_Add_Related_Mail
    
    Dim olTask As Outlook.TaskItem
    Dim olItem As Object
    Dim olExp As Outlook.Explorer
    Dim olApp As Outlook.Application
    
    Set olApp = New Outlook.Application
    Set olTask = olApp.CreateItem(olTaskItem)
    Set olExp = olApp.ActiveExplorer
    
    'Only the 1st selected item will be taken
    If olExp.Selection.Count > 1 Then
        MsgBox "Please select only 1 mail item.", vbOKOnly + vbCritical, "MasterCard"
    
    ElseIf olExp.Selection.Count = 1 Then
        Set olItem = olExp.Selection.Item(1)
        
    ElseIf olExp.Selection.Count = 0 Then
        MsgBox "Please select a mail item.", vbOKOnly + vbCritical, "MasterCard"
        
    End If
    
    Exit Sub
    
Add_Related_Mail:
    MsgBox "Error Number: " & Err.Number & vbCrLf & _
        "Error Description: " & Err.DESCRIPTION & vbCrLf & _
        "Error Source: " & Err.Source, vbOKOnly, "C-SAM Solutions"
End Sub


Public Sub AddToList()
    Dim oMailItem As Outlook.MailItem
    Dim oAppointItem As Outlook.AppointmentItem
    Dim oRecepient As Outlook.Recipient
    Dim oObj As Object
    
    Set oObj = Outlook.Application.ActiveInspector.CurrentItem
    'Debug.Print oObj.ActiveControl
    
    If (oObj.MessageClass = "IPM.Appointment") Then       ' Calendar Entry.
        Set oAppointItem = oObj
        Set oRecepient = oAppointItem.Recipients.Add("Raja, Rashmin <Rashmin.Raja@mastercard.com>; " & _
            "Patil, Preetam <Preetam.Patil@mastercard.com>;  " & _
            "Sinha, Ajay <Ajay.Sinha@mastercard.com>; " & _
            "Bhattacharya, Siddhartha <Siddhartha.Bhattacharya@mastercard.com>;  " & _
            "Yadav, Manoj <Manoj.Yadav@mastercard.com>;  " & _
            "Jadhav, Vaibhav Vitthal <VaibhavVitthal.Jadhav@mastercard.com>;  ")

        oAppointItem.Recipients.ResolveAll

    ElseIf (oObj.MessageClass = "IPM.Note") Then          ' Mail Entry.
        Set oMailItem = oObj
        
        oMailItem.To = Trim$(oMailItem.To) & "; Raja, Rashmin <Rashmin.Raja@mastercard.com>; " & _
            "Patil, Preetam <Preetam.Patil@mastercard.com>;  " & _
            "Sinha, Ajay <Ajay.Sinha@mastercard.com>; " & _
            "Bhattacharya, Siddhartha <Siddhartha.Bhattacharya@mastercard.com>;  " & _
            "Yadav, Manoj <Manoj.Yadav@mastercard.com>;  " & _
            "Jadhav, Vaibhav Vitthal <VaibhavVitthal.Jadhav@mastercard.com>;  "

    Else
        MsgBox "You cannot perform this task here.", vbOKOnly + vbInformation, "MasterCard"

    End If
End Sub

Public Sub AddPeopleMgmtTeam()
    Dim oMailItem As Outlook.MailItem
    Dim oAppointItem As Outlook.AppointmentItem
    Dim oRecepient As Outlook.Recipient
    Dim oObj As Object
    
    Set oObj = Outlook.Application.ActiveInspector.CurrentItem
    'Debug.Print oObj.ActiveControl
    
    If (oObj.MessageClass = "IPM.Appointment") Then       ' Calendar Entry.
        Set oAppointItem = oObj
        Set oRecepient = oAppointItem.Recipients.Add("Raja, Rashmin <Rashmin.Raja@mastercard.com>; " & _
            "Patil, Preetam <Preetam.Patil@mastercard.com>; " & _
            "Yadav, Manoj <Manoj.Yadav@mastercard.com>")

        oAppointItem.Recipients.ResolveAll

    ElseIf (oObj.MessageClass = "IPM.Note") Then          ' Mail Entry.
        Set oMailItem = oObj
        
        
        'Debug.Print oMailItem.To.ActiveControl
        
        oMailItem.To = Trim$(oMailItem.To) & "; Raja, Rashmin <Rashmin.Raja@mastercard.com>; " & _
            "Patil, Preetam <Preetam.Patil@mastercard.com>; " & _
            "Yadav, Manoj <Manoj.Yadav@mastercard.com>"

    Else
        MsgBox "You cannot perform this task here.", vbOKOnly + vbInformation, "MasterCard"

    End If
End Sub

Public Sub AddMeetingRoomsWithProjector()
    Dim oMailItem As Outlook.MailItem
    Dim oAppointItem As Outlook.AppointmentItem
    Dim oRecepient As Outlook.Recipient
    Dim oObj As Object
    
    Set oObj = Outlook.Application.ActiveInspector.CurrentItem
    'MsgBox oObj.MessageClass
    
    If (oObj.MessageClass = "IPM.Appointment") Then       ' Calendar Entry.
        Set oAppointItem = oObj
        Set oRecepient = oAppointItem.Recipients.Add("Chandrasekhar/SE/Pune 4th Floor <Chandrasekhar_SE_Pune@mastercard.com> " & vbCrLf & _
            "JC Bose/SE/Pune 4th Floor <JC_Bose_SE_Pune@mastercard.com> " & vbCrLf & _
            "Kalpana Chawla/NE/Pune 4th Floor <Kalpana_Chawla_NE_Pune@mastercard.com> " & vbCrLf & _
            "Visvesvaraya/NE/Pune 4th Floor <Visvesvaraya_NE_Pune@mastercard.com> " & vbCrLf & _
            "Khorana/NW/Pune 4th Floor <Khorana_NW_Pune@mastercard.com> " & vbCrLf & _
            "Raja Ramanna/SE/Pune 4th Floor <Raja_Ramanna_SE_Pune@mastercard.com> " & vbCrLf & _
            "Newton/NE/Pune 8th Floor <Newton_NE_Pune_8th_Floor@mastercard.com> " & vbCrLf & _
            "Galileo/SW/Pune 8th Floor <Galileo_SW_Pune_8th_Floor@mastercard.com> " & vbCrLf & _
            "Darwin/SW/Pune 8th Floor <Darwin_SW_Pune_8th_Floor@mastercard.com> " & vbCrLf & _
            "Faraday/SE/Pune 8th Floor <Faraday_SE_Pune_8th_Floor@mastercard.com> " & vbCrLf & _
            "Raman/SE/Pune 8th Floor <Raman_SE_Pune_8th_Floor@mastercard.com> " & vbCrLf & _
            "Nobel North/SE/Pune 8th Floor <Nobel_North_SE_Pune_8th_Floor@mastercard.com> " & vbCrLf & _
            "Amazon/NW/Pune 10th Floor <Amazon_Pune_10th_Floor@mastercard.com> " & vbCrLf & _
            "Thames/NW/Pune 10th Floor <Thames_Pune_10th_Floor@mastercard.com> " & vbCrLf & _
            "Yarra/NW/Pune 10th Floor <Yarra_Pune_10th_Floor@mastercard.com> " & vbCrLf & _
            "Mississippi/NW/Pune 10th Floor <Mississippi_Pune_10th_Floor@mastercard.com> " & vbCrLf & _
            "Tigris/NW/Pune 10th Floor <Tigris_Pune_10th_Floor@mastercard.com> " & vbCrLf & _
            "Gange South/NW/Pune 10th Floor <Ganga_-_South_Pune_10th_Floor@mastercard.com> " & vbCrLf & _
            "Gange North/NW/Pune 10th Floor <Ganga_-_North_Pune_10th_Floor@mastercard.com>")
        
        oAppointItem.Recipients.ResolveAll
            
    Else
        MsgBox "You cannot perform this task here.", vbOKOnly + vbInformation, "MasterCard"
        
    End If
End Sub

Public Sub AddMeetingRoomsWithVideoCalling()
    Dim oMailItem As Outlook.MailItem
    Dim oAppointItem As Outlook.AppointmentItem
    Dim oRecepient As Outlook.Recipient
    Dim oObj As Object
    
    Set oObj = Outlook.Application.ActiveInspector.CurrentItem
    'MsgBox oObj.MessageClass
    
    If (oObj.MessageClass = "IPM.Appointment") Then       ' Calendar Entry.
        Set oAppointItem = oObj
        Set oRecepient = oAppointItem.Recipients.Add("Ramanujan/SE/Pune 4th Floor <Ramanujan_SE_Pune@mastercard.com> " & vbCrLf & _
            "Abdul Kalam/SW/Pune 4th Floor <Abdul_Kalam_SW_Pune@mastercard.com> " & vbCrLf & _
            "Homi Bhabha/NE/Pune 4th Floor <Homi_Bhabha_Chawla_NE_Pune@mastercard.com> " & vbCrLf & _
            "Raja Ramanna/SE/Pune 4th Floor <Raja_Ramanna_SE_Pune@mastercard.com> " & vbCrLf & _
            "Einstein/NW/Pune 8th Floor <Einstein_NW_Pune_8th_Floor@mastercard.com> " & vbCrLf & _
            "Nobel South/SE/Pune 8th Floor <Nobel_South_SE_Pune_8th_Floor@mastercard.com> " & vbCrLf & _
            "Telepresence/Pune 9th Floor <Telepresence_Pune_9th_Floor@mastercard.com> " & vbCrLf & _
            "Amazon/NW/Pune 10th Floor <Amazon_Pune_10th_Floor@mastercard.com>")
        
        oAppointItem.Recipients.ResolveAll
            
    Else
        MsgBox "You cannot perform this task here.", vbOKOnly + vbInformation, "MasterCard"
        
    End If
End Sub

Public Sub AddAllMeetingRooms()
    Dim oMailItem As Outlook.MailItem
    Dim oAppointItem As Outlook.AppointmentItem
    Dim oRecepient As Outlook.Recipient
    Dim oObj As Object
    
    Set oObj = Outlook.Application.ActiveInspector.CurrentItem
    'MsgBox oObj.MessageClass
    
    If (oObj.MessageClass = "IPM.Appointment") Then       ' Calendar Entry.
        Set oAppointItem = oObj
        Set oRecepient = oAppointItem.Recipients.Add("Ramanujan/SE/Pune 4th Floor <Ramanujan_SE_Pune@mastercard.com>; JC Bose/SE/Pune 4th Floor <JC_Bose_SE_Pune@mastercard.com>; " & vbCrLf & _
            "Chandrasekhar/SE/Pune 4th Floor <Chandrasekhar_SE_Pune@mastercard.com>; Abdul Kalam/SW/Pune 4th Floor <Abdul_Kalam_SW_Pune@mastercard.com>; " & vbCrLf & _
            "Kalpana Chawla/NE/Pune 4th Floor <Kalpana_Chawla_NE_Pune@mastercard.com>; Visvesvaraya/NE/Pune 4th Floor <Visvesvaraya_NE_Pune@mastercard.com>; " & vbCrLf & _
            "Khorana/NW/Pune 4th Floor <Khorana_NW_Pune@mastercard.com>; Homi Bhabha/NE/Pune 4th Floor <Homi_Bhabha_Chawla_NE_Pune@mastercard.com>; " & vbCrLf & _
            "Raja Ramanna/SE/Pune 4th Floor <Raja_Ramanna_SE_Pune@mastercard.com>; Einstein/NW/Pune 8th Floor <Einstein_NW_Pune_8th_Floor@mastercard.com>; " & vbCrLf & _
            "Newton/NE/Pune 8th Floor <Newton_NE_Pune_8th_Floor@mastercard.com>; Curie/NE/Pune 8th Floor <Curie_NE_Pune_8th_Floor@mastercard.com>; " & vbCrLf & _
            "Galileo/SW/Pune 8th Floor <Galileo_SW_Pune_8th_Floor@mastercard.com>; Darwin/SW/Pune 8th Floor <Darwin_SW_Pune_8th_Floor@mastercard.com>; " & vbCrLf & _
            "Archimedes/SW/Pune 8th Floor <Archimedes_SW_Pune_8th_Floor@mastercard.com>; Faraday/SE/Pune 8th Floor <Faraday_SE_Pune_8th_Floor@mastercard.com>; " & vbCrLf & _
            "Raman/SE/Pune 8th Floor <Raman_SE_Pune_8th_Floor@mastercard.com>; Edison/SE/Pune 8th Floor <Edison_SE_Pune_8th_Floor@mastercard.com>; " & vbCrLf & _
            "Nobel South/SE/Pune 8th Floor <Nobel_South_SE_Pune_8th_Floor@mastercard.com>; Nobel North/SE/Pune 8th Floor <Nobel_North_SE_Pune_8th_Floor@mastercard.com>; " & vbCrLf & _
            "Amur/SE/Pune 9th Floor <Amur_Pune_9th_Floor@mastercard.com>; Ottawa/NW/Pune 9th Floor <Ottawa_Pune_9th_Floor@mastercard.com>; " & vbCrLf & _
            "Nile/NE/Pune 9th Floor <Nile_Pune_9th_Floor@mastercard.com>; Volga/SW/Pune 9th Floor <Volga_Pune_9th_Floor@mastercard.com>; " & vbCrLf & _
            "Niger/SE/Pune 9th Floor <Niger_Pune_9th_Floor@mastercard.com>; Congo/SW/Pune 9th Floor <Congo_Pune_9th_Floor@mastercard.com>; " & vbCrLf & _
            "Telepresence/Pune 9th Floor <Telepresence_Pune_9th_Floor@mastercard.com>; Amazon/NW/Pune 10th Floor <Amazon_Pune_10th_Floor@mastercard.com>; " & vbCrLf & _
            "Thames/NW/Pune 10th Floor <Thames_Pune_10th_Floor@mastercard.com>; Yarra/NW/Pune 10th Floor <Yarra_Pune_10th_Floor@mastercard.com>; " & vbCrLf & _
            "Mississippi/NW/Pune 10th Floor <Mississippi_Pune_10th_Floor@mastercard.com>; Tigris/NW/Pune 10th Floor <Tigris_Pune_10th_Floor@mastercard.com>; " & vbCrLf & _
            "Yukon/NW/Pune 10th Floor <Yukon_Pune_10th_Floor@mastercard.com>; Indus North/SE/Pune 9th Floor <Indus_-_North_Pune_9th_Floor@mastercard.com>; " & vbCrLf & _
            "Gange South/NW/Pune 10th Floor <Ganga_-_South_Pune_10th_Floor@mastercard.com>; Gange North/NW/Pune 10th Floor <Ganga_-_North_Pune_10th_Floor@mastercard.com>; " & vbCrLf & _
            "Rosalind Franklin/SE/Pune 3rd Floor <RosalindFranklin.SE.Pune3rdFloor@mastercard.com>; Caroline Herschel/SE/Pune 3rd Floor <CarolineHerschel/SE/Pune3rdFloor@mastercard.com>; " & vbCrLf & _
            "Ada Lovelace/SE/Pune 3rd Floor <AdaLovelace/SE/Pune3rdFloor@mastercard.com>; Grace Hopper/SE/Pune 3rd Floor <GraceHopper_SE_Pune3rdFloor@mastercard.com>; " & vbCrLf & _
            "Anandibai Joshi/SE/Pune 3rd Floor <AnandibaiJoshi/SE/Pune3rdFloor@mastercard.com>; Shakuntala Devi/SE/Pune 3rd Floor <ShakuntalaDevi/SE/Pune3rdFloor@mastercard.com>; " & vbCrLf & _
            "Tessy Thomas /SW/Pune 3rd Floor <TessyThomas/SW/Pune3rdFloor@mastercard.com>; J.Manjula /NW/Pune 3rd Floor <J.Manjula/NW/Pune3rdFloor@mastercard.com>; " & vbCrLf & _
            "Julia Robinson/NW/Pune 3rd Floor <JuliaRobinson/NW/Pune3rdFloor@mastercard.com>; Sheryl Sandberg/NW/Pune 3rd Floor <SherylSandberg/NW/Pune3rdFloor@mastercard.com>; " & vbCrLf & _
            "Kiran Shaw/NE/Pune 3rd Floor <KiranShaw/NE/Pune3rdFloor@mastercard.com>; P.T. Usha/NE/Pune 3rd Floor <P.T.Usha/NE/Pune3rdFloor@mastercard.com>; " & vbCrLf & _
            "Arunima Sinha/NE/Pune 3rd Floor <ArunimaSinha_NE_Pune3rdFloor@mastercard.com>; Mary Kom/NW/Pune 3rd Floor <MaryKom_NW_Pune3rdFloor@mastercard.com>")


        
        oAppointItem.Recipients.ResolveAll
            
    Else
        MsgBox "You cannot perform this task here.", vbOKOnly + vbInformation, "MasterCard"
        
    End If
End Sub

Public Sub Add4thFloorMeetingRooms()
    Dim oMailItem As Outlook.MailItem
    Dim oAppointItem As Outlook.AppointmentItem
    Dim oRecepient As Outlook.Recipient
    Dim oObj As Object
    
    Set oObj = Outlook.Application.ActiveInspector.CurrentItem
    'MsgBox oObj.MessageClass
    
    If (oObj.MessageClass = "IPM.Appointment") Then       ' Calendar Entry.
        Set oAppointItem = oObj
        Set oRecepient = oAppointItem.Recipients.Add("Ramanujan/SE/Pune 4th Floor <Ramanujan_SE_Pune@mastercard.com>; " & vbCrLf & _
            "JC Bose/SE/Pune 4th Floor <JC_Bose_SE_Pune@mastercard.com>; " & vbCrLf & _
            "Chandrasekhar/SE/Pune 4th Floor <Chandrasekhar_SE_Pune@mastercard.com>; " & vbCrLf & _
            "Abdul Kalam/SW/Pune 4th Floor <Abdul_Kalam_SW_Pune@mastercard.com>; " & vbCrLf & _
            "Kalpana Chawla/NE/Pune 4th Floor <Kalpana_Chawla_NE_Pune@mastercard.com>; " & vbCrLf & _
            "Visvesvaraya/NE/Pune 4th Floor <Visvesvaraya_NE_Pune@mastercard.com>; " & vbCrLf & _
            "Khorana/NW/Pune 4th Floor <Khorana_NW_Pune@mastercard.com>; " & vbCrLf & _
            "Homi Bhabha/NE/Pune 4th Floor <Homi_Bhabha_Chawla_NE_Pune@mastercard.com>; " & vbCrLf & _
            "Raja Ramanna/SE/Pune 4th Floor <Raja_Ramanna_SE_Pune@mastercard.com>;")
        
        oAppointItem.Recipients.ResolveAll
            
    Else
        MsgBox "You cannot perform this task here.", vbOKOnly + vbInformation, "MasterCard"
        
    End If
End Sub

Public Sub Add3rdFloorMeetingRooms()
    Dim oMailItem As Outlook.MailItem
    Dim oAppointItem As Outlook.AppointmentItem
    Dim oRecepient As Outlook.Recipient
    Dim oObj As Object
    
    Set oObj = Outlook.Application.ActiveInspector.CurrentItem
    'MsgBox oObj.MessageClass
    
    If (oObj.MessageClass = "IPM.Appointment") Then
        Set oAppointItem = oObj
        Set oRecepient = oAppointItem.Recipients.Add("Rosalind Franklin/SE/Pune 3rd Floor <RosalindFranklin.SE.Pune3rdFloor@mastercard.com>; " & vbCrLf & _
            "Caroline Herschel/SE/Pune 3rd Floor <CarolineHerschel.SE.Pune3rdFloor@mastercard.com>; " & vbCrLf & _
            "Ada Lovelace/SE/Pune 3rd Floor <AdaLovelace.SE.Pune3rdFloor@mastercard.com>; " & vbCrLf & _
            "Anandibai Joshi/SE/Pune 3rd Floor <AnandibaiJoshi.SE.Pune3rdFloor@mastercard.com>; " & vbCrLf & _
            "Shakuntala Devi/SE/Pune 3rd Floor <ShakuntalaDevi/SE/Pune3rdFloor@mastercard.com>; " & vbCrLf & _
            "J.Manjula /NW/Pune 3rd Floor <J.Manjula/NW/Pune3rdFloor@mastercard.com>; " & vbCrLf & _
            "Julia Robinson/NW/Pune 3rd Floor <JuliaRobinson.NW.Pune3rdFloor@mastercard.com>; " & vbCrLf & _
            "Sheryl Sandberg/NW/Pune 3rd Floor <SherylSandberg/NW/Pune3rdFloor@mastercard.com>; " & vbCrLf & _
            "Kiran Shaw/NE/Pune 3rd Floor <KiranShaw.NE.Pune3rdFloor@mastercard.com>; " & vbCrLf & _
            "P.T. Usha/NE/Pune 3rd Floor <P.T.Usha/NE/Pune3rdFloor@mastercard.com>; " & vbCrLf & _
            "Arunima Sinha/NE/Pune 3rd Floor <ArunimaSinha.NE.Pune3rdFloor@mastercard.com>; " & vbCrLf & _
            "Grace Hopper/SE/Pune 3rd Floor <GraceHopper_SE_Pune3rdFloor@mastercard.com>; " & vbCrLf & _
            "Mary Kom/NW/Pune 3rd Floor <MaryKom.NW.Pune3rdFloor@mastercard.com>")


        
        oAppointItem.Recipients.ResolveAll
            
    Else
        MsgBox "You cannot perform this task here.", vbOKOnly + vbInformation, "MasterCard"
        
    End If
End Sub

Public Sub test()
    Dim o As New ViewCtl
    
End Sub

