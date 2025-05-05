
Function SendAuthEmail(Recipient, Subject, Body, AttachmentPath)

    Dim objMessage, objConfig, Fields, senderEmail, senderPass
    Dim OutlookApp, MailItem

  '##### Input your email credentials #######
    senderEmail = "your.user.email@triumph.com"
    senderPass  = "YourEmailPassword"
  '###########################################

    Set OutlookApp = Nothing
    Set OutlookApp = CreateObject("Outlook.Application")
    Set MailItem = OutlookApp.CreateItem(0)
    Set objMessage = CreateObject("CDO.Message")
    Set objConfig = CreateObject("CDO.Configuration")

    Set Fields = objConfig.Fields
    With Fields
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2  
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.office365.com"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 587
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = senderEmail
        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = senderPass
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
        .Update
    End With

     MailItem.To = Recipient
     MailItem.Subject = Subject
     MailItem.HTMLBody = "<html><body><pre>" & Body & "</pre></body></html>"
      If AttachmentPath <> "" Then
        MailItem.Attachments.Add AttachmentPath
      End If
    MailItem.Send

    Set MailItem = Nothing
    Set OutlookApp = Nothing
    Set objMessage = Nothing
    Set objConfig = Nothing
    Set Fields = Nothing
    WScript.Sleep 350

End Function

Function SendEmail(Recipient, Subject, Body, AttachmentPath)

    Dim OutlookApp, MailItem
    Set OutlookApp = Nothing
    Set OutlookApp = CreateObject("Outlook.Application")
    Set MailItem = OutlookApp.CreateItem(0)

     MailItem.To = Recipient
     MailItem.Subject = Subject
     MailItem.HTMLBody = "<html><body><pre>" & Body & "</pre></body></html>"
      If AttachmentPath <> "" Then
        MailItem.Attachments.Add AttachmentPath
      End If
    MailItem.Send

    Set MailItem = Nothing
    Set OutlookApp = Nothing
    WScript.Sleep 350

End Function
'                                                                     --------  Main  --------

'Option Explicit
Dim fso, inputFolder , oFolder, oFile, fileExtension
Dim dvs, matches, extractedID, shell
Dim dvsCustPoExcel, dvsCustPoWB, dvsCustPoWSheet, custPoLastRow, iOrder, order, custPO
Dim inputCustPoFile, configFolder, dictEmails, dvsExcel, dvsWorkbook, dvsWorksheet
Dim emailLastRow, iEmail, emailName, emailValue, emailSubject, emailBody, fileForEmailSubject
Set shell = CreateObject("WScript.Shell")
rootFolder = shell.CurrentDirectory & "\"
'emailSubject = "Order confirmation"
inputFolder  = "F:\ABSDIE~1\ABNSDI~1"
configFolder = rootFolder & "Config\"
inputCustPoFile = configFolder & "CustomerOrdersExport1.csv"
Set dvs = CreateObject("VBScript.RegExp")
dvs.Pattern = "_(.*?)(?:_|\.pdf$)" 
dvs.Global = False
dvs.IgnoreCase = True
Set fso = CreateObject("Scripting.FileSystemObject")
folderExist = fso.FolderExists(configFolder)
    If Not folderExist then 
        fso.CreateFolder(configFolder)
        WScript.Echo "Check if config files are placed in 'Config' folder"
        WScript.Quit
    Else
        If Not fso.FileExists(configFolder & "SubjectText.txt") then 
            WScript.Echo "File 'SubjectText.txt' is missing from the 'Config' folder. Please check."
            WScript.Quit
        End if
        If Not fso.FileExists(configFolder & "Emails.xlsx") then 
            WScript.Echo "File 'Emails.xlsx' with list of client emails is missing from the 'Config' folder. Please check."
            WScript.Quit
        End if
    End if
Set fileForEmailSubject = fso.OpenTextFile(configFolder & "SubjectText.txt", 1)
emailBody = fileForEmailSubject.ReadAll

'############################################# Add Emails #############################################
Set dvsExcel = CreateObject("Excel.Application")
Set dvsWorkbook = dvsExcel.Workbooks.Open(configFolder & "Emails.xlsx")
Set dvsWorksheet = dvsWorkbook.Sheets(1)
dvsExcel.Visible = False
dvsExcel.DisplayAlerts=False
emailLastRow = dvsWorksheet.UsedRange.Rows.Count
Set dictEmails = CreateObject("Scripting.Dictionary")

For iEmail = 1 to emailLastRow
    customerName = Trim(CStr(dvsWorksheet.Cells(iEmail, 1).Value))
    emailValue = Trim(CStr(dvsWorksheet.Cells(iEmail, 2).Value))

    If Not dictEmails.Exists(customerName) Then
        dictEmails.Add customerName, emailValue
    End If
    
Next

dvsExcel.Quit
Set dvsExcel = Nothing
Set dvsWorkbook = Nothing
Set dvsWorksheet = Nothing

'############################################# Combined Mapping #############################################
Set dvsCustPoExcel = CreateObject("Excel.Application")
Set dvsCustPoWB = dvsCustPoExcel.Workbooks.Open(inputCustPoFile)
Set dvsCustPoWSheet = dvsCustPoWB.Sheets(1)

Set mainDict = CreateObject("Scripting.Dictionary")
mainDict.RemoveAll()
dvsCustPoExcel.Visible = False
dvsCustPoExcel.DisplayAlerts=False
custPoLastRow = dvsCustPoWSheet.UsedRange.Rows.Count

For iOrder = 1 to custPoLastRow
    order = Trim(CStr(dvsCustPoWSheet.Cells(iOrder, 1).Value))
    customerName = Trim(CStr(dvsCustPoWSheet.Cells(iOrder, 4).Value))

    If order <> "" And customerName <> "" Then
        Set info = CreateObject("Scripting.Dictionary")
        info.Add "CustomerName", customerName
        emailFound = False
        For Each emailCustomerName In dictEmails.Keys
            If InStr(1, customerName, emailCustomerName, vbTextCompare) > 0 Then
                info.Add "Email", dictEmails(emailCustomerName)
                emailFound = True
                Exit For
            End If
        Next

        If Not emailFound Then
            info.Add "Email", ""  ' No match found
        End If

        mainDict.Add order, info
    End If
Next

dvsCustPoExcel.Quit
Set dvsCustPoExcel = Nothing
Set dvsCustPoWB = Nothing
Set dvsCustPoWSheet = Nothing

'############################################# Main part #############################################
Dim reportFile, writeToReport
reportFile = rootFolder & "Report_" & year(Now()) & month(Now()) & day(Now()) & ".txt"
Set writeToReport = fso.OpenTextFile(reportFile,2,True)
writeToReport.WriteLine  Time & "- Process started!"

If fso.FolderExists(inputFolder ) Then
    On Error Resume Next
    Set oFolder = fso.GetFolder(inputFolder )
    If Err.Number <> 0 Then
         WScript.Echo "ERROR: Could not get the Folder object even though it exists." & vbCrLf & "Check folder permissions. Error: " & Err.Description
         Set dvs = Nothing : Set fso = Nothing
         WScript.Quit
    End If
    On Error GoTo 0

    If oFolder.Files.Count > 0 Then
        For Each oFile In oFolder.Files
          fileExtension = fso.GetExtensionName(oFile.Path)

            If LCase(fileExtension) = "pdf" Then
                extractedID = "ID_Not_Found"
                Set matches = dvs.Execute(oFile.Name)
                If matches.Count > 0 Then
                    extractedID = matches(0).SubMatches(0)
                    'WScript.Echo "Extracted ID: " & extractedID
                Else
                    'WScript.Echo "No ID found matching pattern '" & dvs.Pattern & "'."
                End If
                Set matches = Nothing
                emailSubject =  "Order confirmation - " & oFile.Name
                    If mainDict.Exists(extractedID) Then
                        Set data = mainDict(extractedID)
                        sendTo = data("Email")

                        If sendTo = "Email not found" OR sendTo = "" Then
                            writeToReport.WriteLine Time & "- Missing email from the table for " & extractedID
                        Else
                             'SendAuthEmail sendTo, emailSubject, emailBody, xlsxFile
                             SendEmail sendTo, emailSubject, emailBody, oFile.Path
                            writeToReport.WriteLine Time & "- Email sent to " & sendTo & " for order " & extractedID
                        End If

                    Else
                        writeToReport.WriteLine Time & "- Order ID not found: " & extractedID
                    End If
            Else
                 ' Skipping non-PDF files
            End If
        Next

    Else
        WScript.Echo "The folder exists but contains no files."
    End If
    Set oFolder = Nothing
Else
    WScript.Echo "ERROR: Short path does not exist or cannot be accessed." & vbCrLf & _
                 "Check:" & vbCrLf & _
                 "1. The short path '" & inputFolder  & "' is EXACTLY correct (use DIR /X in Command Prompt)." & vbCrLf & _
                 "2. You have permissions to access the folder."
End If

WScript.Echo "Finished scanning files."

'############################################# Clean up #############################################
Set dvs = Nothing
Set fso = Nothing
Set custPO_dict = Nothing
Set writeToReport = Nothing
