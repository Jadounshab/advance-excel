# Project on VBA
/ This is a project based on VBA in which I try to automate some excel work related to HR purpose.
In this project I bulit a program using VBA which goes to define cell and run the loop and creating salary slips of the respected employee.
After Making salary slip it automaticaly save in a define folder with thier employee code and name Ex: (Name = Rony, code = 34246) so file name would be 34246_Rony.
And then I configure the outlook with this program.
So. when running a loop it creates a mail and type a receiver email from the give data in TO box and type a subject automaticaly and also type a mail body and attached a attachment of salary slip in PDF format and send a email one by one till given data./  

/Here below I shared VBA code./

Sub GenPaySlip()
    LastRow = Sheet1.Range("C" & Rows.Count).End(xlUp).Row
    For EmpRow = 2 To LastRow
        Sheet2.Range("D14") = Sheet1.Range("C" & EmpRow)
        Filename = "C:\Users\99299\OneDrive\Desktop\SVG work\Slip_folder\" & Sheet1.Range("C" & EmpRow) & Sheet2.Range("H14") & ".pdf"
        Sheet2.ExportAsFixedFormat xlTypePDF, Filename
        SendEmail
    Next EmpRow
End Sub



Sub SendEmail()
    Dim OutApp As Outlook.Application
    Dim OutMail As Outlook.MailItem
    Set OutApp = New Outlook.Application
    Set OutMail = OutApp.CreateItem(olMailItem)
    EmpMail = Sheet1.Range("D:D").Find(Sheet2.Range("D15"), , xlValues, xlWhole).Row
    Filename = "C:\Users\99299\OneDrive\Desktop\SVG work\Slip_folder\" & Sheet2.Range("D14") & Sheet2.Range("H14") & ".pdf"
    With OutMail
        .To = Sheet1.Range("AS" & EmpMail)
        .Subject = "Salary slip for the month of " & Sheet2.Range("H14") & "-" & Sheet2.Range("D15")
        .Body = "Dear Employee," & vbCrLf & "Please Find attached your Salary Slip"
        .Attachments.Add Filename
        .Send
    End With
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub

