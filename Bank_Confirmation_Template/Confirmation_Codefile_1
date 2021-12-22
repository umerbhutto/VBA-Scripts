Sub bankconfirmationtemplate()

' Creates the last row that will be used
lRow = ThisWorkbook.Sheets("Input").Cells(Rows.Count, 1).End(xlUp).Row
' Loop through all the rows

' Calls the necessary macros for each template based on region
For k = 3 To lRow
Select Case Cells(k, 16)

    Case "UAE"
        Call bankuaetemplate(k)
    
    Case "DIFC"
        Call bankdifctemplate(k)
    
    Case "ADGM"
        Call bankadgmtemplate(k)
        
    Case ""
        MsgBox ("No region was selected for row " & Str(k))
        Exit Sub
        
End Select

Next k

ThisWorkbook.Sheets("Input").Activate
MsgBox (Str(k - 3) & " bank confirmations have been created and saved in the following folder: " & ThisWorkbook.Path & "\Confirmations.")


End Sub

Sub legalconfirmationtemplate()

' Workbook name - Update this when applicable
workbook_name = "Excel - Bank Confirmation_v1.0.xlsm"

lRow = ThisWorkbook.Sheets("Input_").Cells(Rows.Count, 1).End(xlUp).Row
' Creates the last row that will be used

For i = 3 To lRow

Dim wordApp As Word.Application
Dim wDoc As Word.Document

' Create Word Application using Excel
Set wordApp = CreateObject("word.application")
Set wDoc = wordApp.Documents.Open(Workbooks(workbook_name).Path & "\Templates\legal_confirmation_template_uae.docm")
wordApp.Visible = False 'Used to keep the word application visible or not

''' Following are the controls that pertain to the Word Document by order of control
''' e.g. control 1 refers to date of letter etc.

' Control 1 - Date of Letter
wDoc.ContentControls(1).Range.Text = Workbooks(workbook_name).Sheets("Input_").Cells(i, 2)

' Control 2 - Legal Contact Name
wDoc.ContentControls(2).Range.Text = Workbooks(workbook_name).Sheets("Input_").Cells(i, 15)

' Control 3 - Legal Firm Name
wDoc.ContentControls(3).Range.Text = Workbooks(workbook_name).Sheets("Input_").Cells(i, 9)

' Control 4 - Adress Line 1
wDoc.ContentControls(4).Range.Text = Workbooks(workbook_name).Sheets("Input_").Cells(i, 10)

'Control 5 - Adress Line 2
wDoc.ContentControls(5).Range.Text = Workbooks(workbook_name).Sheets("Input_").Cells(i, 11)

'Control 6 - PO Box
wDoc.ContentControls(6).Range.Text = "P.O. Box " & Workbooks(workbook_name).Sheets("Input_").Cells(i, 12) & ", " & Workbooks(workbook_name).Sheets("Input_").Cells(i, 13) & ", " & Workbooks(workbook_name).Sheets("Input_").Cells(i, 14)

'Control 7 - Legal Firm Telephone No.
wDoc.ContentControls(7).Range.Text = "Tel No: +" & Workbooks(workbook_name).Sheets("Input_").Cells(i, 17)

'Control 8 - Email
wDoc.ContentControls(8).Range.Text = "Email: " & Workbooks(workbook_name).Sheets("Input_").Cells(i, 16)

'Control 9/19 - Client Name
wDoc.ContentControls(9).Range.Text = Workbooks(workbook_name).Sheets("Input_").Cells(i, 8)
wDoc.ContentControls(19).Range.Text = Workbooks(workbook_name).Sheets("Input_").Cells(i, 8)

'Control 10/11 - Period End Date / Period Beg Date
wDoc.ContentControls(10).Range.Text = DateValue("1 January " & DatePart("yyyy", Workbooks(workbook_name).Sheets("Input_").Cells(i, 4)))
wDoc.ContentControls(11).Range.Text = Workbooks(workbook_name).Sheets("Input_").Cells(i, 4)

'Control 12/16 - PwC Contact
wDoc.ContentControls(12).Range.Text = Workbooks(workbook_name).Sheets("Input_").Cells(i, 1)
wDoc.ContentControls(16).Range.Text = Workbooks(workbook_name).Sheets("Input_").Cells(i, 1)

'Control 13/18 - PwC Contact Email
wDoc.ContentControls(13).Range.Text = Workbooks(workbook_name).Sheets("Input_").Cells(i, 3)
wDoc.ContentControls(18).Range.Text = Workbooks(workbook_name).Sheets("Input_").Cells(i, 3)

' Control 14 - Deadline Date
wDoc.ContentControls(14).Range.Text = Workbooks(workbook_name).Sheets("Input_").Cells(i, 5)

' Control 15 - Expected Date
wDoc.ContentControls(15).Range.Text = Workbooks(workbook_name).Sheets("Input_").Cells(i, 6)

'Control 17 - PwC Contact Telephone No.
wDoc.ContentControls(17).Range.Text = "+" & Workbooks(workbook_name).Sheets("Input_").Cells(i, 7)

' Create the file name and save and close the file
file_name = Application.WorksheetFunction.Trim(Str(i - 2) & "-LegalConf-" & Workbooks(workbook_name).Sheets("Input_").Cells(i, 8) & "-" & Workbooks(workbook_name).Sheets("Input_").Cells(i, 9) & ".doc")
wDoc.SaveAs2 (Workbooks(workbook_name).Path & "\Confirmations\" & file_name)
wDoc.Close

'Closes all instances of word application (not needed anymore)
wordApp.Quit
Set wordApp = Nothing
''' Contact Umer Bhutto for any issues

Next i

ThisWorkbook.Sheets("Input_").Activate
MsgBox (Str(i - 3) & " legal confirmations have been created and saved in the following folder: " & ThisWorkbook.Path & "\Confirmations.")


End Sub

Sub custodianconfirmationtemplate()

' Workbook name - Update this when applicable
workbook_name = "Excel - Bank Confirmation_v1.0.xlsm"

lRow = ThisWorkbook.Sheets("Input_C").Cells(Rows.Count, 1).End(xlUp).Row
' Creates the last row that will be used

For i = 3 To lRow

Dim wordApp As Word.Application
Dim wDoc As Word.Document

' Create Word Application using Excel
Set wordApp = CreateObject("word.application")
Set wDoc = wordApp.Documents.Open(Workbooks(workbook_name).Path & "\Templates\custodian_confirmation_uae.docm")
wordApp.Visible = False 'Used to keep the word application visible or not

''' Following are the controls that pertain to the Word Document by order of control
''' e.g. control 1 refers to date of letter etc.

' Control 1 - Date of Letter
wDoc.ContentControls(1).Range.Text = Workbooks(workbook_name).Sheets("Input_C").Cells(i, 2)

' Control 2 - Custodian Contact Name
wDoc.ContentControls(2).Range.Text = Workbooks(workbook_name).Sheets("Input_C").Cells(i, 12)

' Control 3 - Custodian Name
wDoc.ContentControls(3).Range.Text = Workbooks(workbook_name).Sheets("Input_C").Cells(i, 6)

' Control 4 - Adress Line 1
wDoc.ContentControls(4).Range.Text = Workbooks(workbook_name).Sheets("Input_C").Cells(i, 7)

'Control 5 - Adress Line 2
wDoc.ContentControls(5).Range.Text = Workbooks(workbook_name).Sheets("Input_C").Cells(i, 8)

'Control 6 - PO Box, City, Country
wDoc.ContentControls(6).Range.Text = "P.O. Box " & Workbooks(workbook_name).Sheets("Input_C").Cells(i, 9) & ", " & Workbooks(workbook_name).Sheets("Input_C").Cells(i, 10) & ", " & Workbooks(workbook_name).Sheets("Input_C").Cells(i, 11)

'Control 7 - Custodian Contact Telephone No.
wDoc.ContentControls(7).Range.Text = "Tel No: +" & Workbooks(workbook_name).Sheets("Input_C").Cells(i, 14)

'Control 8 - Custodian Contact Email
wDoc.ContentControls(8).Range.Text = "Email: " & Workbooks(workbook_name).Sheets("Input_C").Cells(i, 13)

'Control 9 - Client Name
wDoc.ContentControls(9).Range.Text = Workbooks(workbook_name).Sheets("Input_C").Cells(i, 5)

'Control 10 - Period End Date
wDoc.ContentControls(10).Range.Text = Workbooks(workbook_name).Sheets("Input_C").Cells(i, 4)

'Control 11
wDoc.ContentControls(11).Range.Text = Workbooks(workbook_name).Sheets("Input_C").Cells(i, 1)

'Control 12 - PwC Contact Email
wDoc.ContentControls(12).Range.Text = Workbooks(workbook_name).Sheets("Input_C").Cells(i, 3)

' Create the file name and save and close the file
' Confirmation No. - CustodianConf - Client Name - Custodian Name
file_name = Application.WorksheetFunction.Trim(Str(i - 2) & "-CustodianConf-" & Workbooks(workbook_name).Sheets("Input_C").Cells(i, 5) & "-" & Workbooks(workbook_name).Sheets("Input_C").Cells(i, 6) & ".docm")
wDoc.SaveAs2 (Workbooks(workbook_name).Path & "\Confirmations\" & file_name)
wDoc.Close

'Closes all instances of word application (not needed anymore)
wordApp.Quit
Set wordApp = Nothing
''' Contact Umer Bhutto for any issues

Next i

ThisWorkbook.Sheets("Input_C").Activate
MsgBox (Str(i - 3) & " Custodian confirmations have been created and saved in the following folder: " & ThisWorkbook.Path & "\Confirmations.")

End Sub
