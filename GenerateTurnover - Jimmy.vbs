run_macro

WScript.Sleep 300

SendEmail

WScript.Sleep 100

quitscript

sub quitscript() 
    Wscript.quit 
end sub 

Sub run_macro()
    Dim objExcel
    Dim xlBook
    Dim FolderFromPath
	Dim dteNow
	
	dteNow = Date()
	dteNow = DateAdd("d", dteNow, -2)

    Set objExcel = CreateObject("Excel.application")
	
	objExcel.DisplayAlerts = False
	
    set xlBook = objExcel.Workbooks.Open("C:\Users\user.name\Desktop\turnover_report\report.xlsm")
	
    objExcel.Application.run "'" & xlBook.Name & "'!ReportMacro"
	
  '    objExcel.ActiveWorkbook.SaveAs "C:\Users\user.name\Desktop\turnover_report\" &  Right("0" & Day(dteNow), 2) & "_" & Right("0" & Month(dteNow), 2) & "_" _
   ' & Year(dteNow) & "_TO" & Year(dteNow) & ".xlsx", 51

	  objExcel.ActiveWorkbook.SaveAs "C:\Users\user.name\Desktop\turnover_report\" &  "PQC_CurrYear_" & Year(dteNow) & Right("0" & Month(dteNow), 2) & Right("0" & Day(dteNow), 2) & ".xlsx", 51
	
	
    objExcel.Application.Quit
	
	'call quitscript
	
End Sub


Sub SendEmail()

    Dim signature
    Dim objOL
    Dim objMI
	
	Dim dteNow 
	dteNow = Date()
	dteNow = DateAdd("d", dteNow, -2)

    Set objOL = CreateObject("Outlook.Application")
	
	'what_address = "username@abc"
	
	what_address = "user.name@abc.com.hk"
	
	subject_line = "turnover report - " & Right("0" & Day(dteNow), 2) & "_" & Right("0" & Month(dteNow), 2) & "_" & Year(dteNow) 
	
	'mail_body = "Hi Jimmy<br><br>It is turnover report. <br><br> Thanks<br><br>Jimmy"
	
	mail_body = "Dear Abc, " & vbNewLine & vbNewLine & "It is turnover report." _
		& vbNewLine & vbNewLine & "Thanks" _
		& vbNewLine & vbNewLine & "Jimmy"
	 
    ' Create e-mail message

    Set objMI = objOL.CreateItem(0) ' 0 = olMailItem

   
    With objMI

        .Display

        '.SentOnBehalfOfName = "lara.w"

        '.SentOnBehalfOfName = "abc-info"

        .To = what_address

        .Subject = subject_line 

        '.HTMLBody = mail_body & .HTMLBody
		
		.Body = mail_body
		
        ' File Attachment

        .Attachments.Add "C:\Users\user.name\Desktop\turnover_report\" &  "PQC_CurrYear_" & Year(dteNow) & Right("0" & Month(dteNow), 2) & Right("0" & Day(dteNow), 2) & ".xlsx"

        .Send       

    End With

    ' Quit Outlook

    ' objOL.Quit

    Set signature = Nothing

    Set objOL = Nothing

    Set objMI = Nothing

End Sub