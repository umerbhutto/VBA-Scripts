Sub bankuaetemplate(i)

Dim wordApp As Word.Application
Dim wDoc As Word.Document

workbook_name = "Excel - Bank Confirmation_v1.0.xlsm"

' Create Word Application using Excel
Set wordApp = CreateObject("word.application")
Set wDoc = wordApp.Documents.Open(Workbooks(workbook_name).Path & "\Templates\bank_confirmation_template_uae.docm")
wordApp.Visible = False 'Used to keep the word application visible or not

''' Following are the controls that pertain to the Word Document by order of control
''' e.g. control 1 refers to date of letter etc.

' Control 1/21 - Date of Letter
wDoc.ContentControls(1).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 2)
wDoc.ContentControls(21).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 2)

' Control 2/14 - Bank Contact Name
wDoc.ContentControls(2).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 13)
wDoc.ContentControls(14).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 13)

' Control 3/15 - Bank Name
wDoc.ContentControls(3).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 7)
wDoc.ContentControls(15).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 7)

' Control 4/16 - Adress Line 1
wDoc.ContentControls(4).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 8)
wDoc.ContentControls(16).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 8)

'Control 5/17 - Adress Line 2
wDoc.ContentControls(5).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 9)
wDoc.ContentControls(17).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 9)

'Control 6/18 - PO Box
wDoc.ContentControls(6).Range.Text = "P.O. Box " & Workbooks(workbook_name).Sheets("Input").Cells(i, 10) & ", " & Workbooks(workbook_name).Sheets("Input").Cells(i, 11) & ", " & Workbooks(workbook_name).Sheets("Input").Cells(i, 12)
wDoc.ContentControls(18).Range.Text = "P.O. Box " & Workbooks(workbook_name).Sheets("Input").Cells(i, 10) & ", " & Workbooks(workbook_name).Sheets("Input").Cells(i, 11) & ", " & Workbooks(workbook_name).Sheets("Input").Cells(i, 12)

'Control 7/19 - Banker Telephone No.
wDoc.ContentControls(7).Range.Text = "Tel No: +" & Workbooks(workbook_name).Sheets("Input").Cells(i, 15)
wDoc.ContentControls(19).Range.Text = "Tel No: +" & Workbooks(workbook_name).Sheets("Input").Cells(i, 15)

'Control 8/20 - Email
wDoc.ContentControls(8).Range.Text = "Email: " & Workbooks(workbook_name).Sheets("Input").Cells(i, 14)
wDoc.ContentControls(20).Range.Text = "Email: " & Workbooks(workbook_name).Sheets("Input").Cells(i, 14)

'Control 9/22 - Client Name
wDoc.ContentControls(9).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 6)
wDoc.ContentControls(22).Range.Text = UCase(Workbooks(workbook_name).Sheets("Input").Cells(i, 6))

'Control 10/23/24/25/26 - Period End Date
wDoc.ContentControls(10).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 4)
wDoc.ContentControls(23).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 4)
wDoc.ContentControls(24).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 4)
wDoc.ContentControls(26).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 4)
''' This relates to the starting period and is based on the period end date
wDoc.ContentControls(25).Range.Text = DateValue("1 January " & DatePart("yyyy", Workbooks(workbook_name).Sheets("Input").Cells(i, 4)))

'Control 11/27 - PwC Contact
wDoc.ContentControls(11).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 1)
wDoc.ContentControls(27).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 1)

'Control 12/28 - PwC Contact Email
wDoc.ContentControls(12).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 3)
wDoc.ContentControls(28).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 3)

'Control 13 - PwC Contact Telephone No.
wDoc.ContentControls(13).Range.Text = "+" & Workbooks(workbook_name).Sheets("Input").Cells(i, 5)

' Update Headers from page 3 to page 5
For j = 3 To 5
With wDoc.Sections(j).Headers(wdHeaderFooterPrimary).Range
    .InsertAfter Text:=vbCrLf & UCase(Workbooks(workbook_name).Sheets("Input").Cells(i, 6))
    .InsertAfter vbTab
    .InsertAfter Text:=vbCrLf & vbCrLf & UCase(Workbooks(workbook_name).Sheets("Input").Cells(i, 7))
    .InsertAfter vbTab
    .InsertAfter Text:=vbCrLf & vbCrLf & ("At close of business on 31 December " & DatePart("yyyy", Workbooks(workbook_name).Sheets("Input").Cells(i, 4)))
End With
Next j

' Close the header / footer pane
wDoc.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
wDoc.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument

' Create the file name and save and close the file
file_name = Application.WorksheetFunction.Trim(Str(i - 2) & "-BankConf-" & Workbooks(workbook_name).Sheets("Input").Cells(i, 6) & "-" & Workbooks(workbook_name).Sheets("Input").Cells(i, 7) & ".doc")
wDoc.SaveAs2 (Workbooks(workbook_name).Path & "\Confirmations\" & file_name)
wDoc.Close

'Closes all instances of word application (not needed anymore)
wordApp.Quit
Set wordApp = Nothing
''' Contact Umer Bhutto for any issues


End Sub

Sub bankdifctemplate(i)

Dim wordApp As Word.Application
Dim wDoc As Word.Document

workbook_name = "Excel - Bank Confirmation_v1.0.xlsm"

' Create Word Application using Excel
Set wordApp = CreateObject("word.application")
Set wDoc = wordApp.Documents.Open(Workbooks(workbook_name).Path & "\Templates\bank_confirmation_template_difc.docm")
wordApp.Visible = False 'Used to keep the word application visible or not

''' Following are the controls that pertain to the Word Document by order of control
''' e.g. control 1 refers to date of letter etc.

' Control 1/21 - Date of Letter
wDoc.ContentControls(1).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 2)
wDoc.ContentControls(21).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 2)

' Control 2/14 - Bank Contact Name
wDoc.ContentControls(2).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 13)
wDoc.ContentControls(14).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 13)

' Control 3/15 - Bank Name
wDoc.ContentControls(3).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 7)
wDoc.ContentControls(15).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 7)

' Control 4/16 - Adress Line 1
wDoc.ContentControls(4).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 8)
wDoc.ContentControls(16).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 8)

'Control 5/17 - Adress Line 2
wDoc.ContentControls(5).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 9)
wDoc.ContentControls(17).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 9)

'Control 6/18 - PO Box
wDoc.ContentControls(6).Range.Text = "P.O. Box " & Workbooks(workbook_name).Sheets("Input").Cells(i, 10) & ", " & Workbooks(workbook_name).Sheets("Input").Cells(i, 11) & ", " & Workbooks(workbook_name).Sheets("Input").Cells(i, 12)
wDoc.ContentControls(18).Range.Text = "P.O. Box " & Workbooks(workbook_name).Sheets("Input").Cells(i, 10) & ", " & Workbooks(workbook_name).Sheets("Input").Cells(i, 11) & ", " & Workbooks(workbook_name).Sheets("Input").Cells(i, 12)

'Control 7/19 - Banker Telephone No.
wDoc.ContentControls(7).Range.Text = "Tel No: +" & Workbooks(workbook_name).Sheets("Input").Cells(i, 15)
wDoc.ContentControls(19).Range.Text = "Tel No: +" & Workbooks(workbook_name).Sheets("Input").Cells(i, 15)

'Control 8/20 - Email
wDoc.ContentControls(8).Range.Text = "Email: " & Workbooks(workbook_name).Sheets("Input").Cells(i, 14)
wDoc.ContentControls(20).Range.Text = "Email: " & Workbooks(workbook_name).Sheets("Input").Cells(i, 14)

'Control 9/22 - Client Name
wDoc.ContentControls(9).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 6)
wDoc.ContentControls(22).Range.Text = UCase(Workbooks(workbook_name).Sheets("Input").Cells(i, 6))

'Control 10/23/24/25/26 - Period End Date
wDoc.ContentControls(10).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 4)
wDoc.ContentControls(23).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 4)
wDoc.ContentControls(24).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 4)
wDoc.ContentControls(26).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 4)
''' This relates to the starting period and is based on the period end date
wDoc.ContentControls(25).Range.Text = DateValue("1 January " & DatePart("yyyy", Workbooks(workbook_name).Sheets("Input").Cells(i, 4)))

'Control 11/27 - PwC Contact
wDoc.ContentControls(11).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 1)
wDoc.ContentControls(27).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 1)

'Control 12/28 - PwC Contact Email
wDoc.ContentControls(12).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 3)
wDoc.ContentControls(28).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 3)

'Control 13 - PwC Contact Telephone No.
wDoc.ContentControls(13).Range.Text = "+" & Workbooks(workbook_name).Sheets("Input").Cells(i, 5)

' Update Headers from page 3 to page 5
For j = 3 To 5
With wDoc.Sections(j).Headers(wdHeaderFooterPrimary).Range
    .InsertAfter Text:=vbCrLf & UCase(Workbooks(workbook_name).Sheets("Input").Cells(i, 6))
    .InsertAfter vbTab
    .InsertAfter Text:=vbCrLf & vbCrLf & UCase(Workbooks(workbook_name).Sheets("Input").Cells(i, 7))
    .InsertAfter vbTab
    .InsertAfter Text:=vbCrLf & vbCrLf & ("At close of business on 31 December " & DatePart("yyyy", Workbooks(workbook_name).Sheets("Input").Cells(i, 4)))
End With
Next j

' Close the header / footer pane
wDoc.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
wDoc.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument

' Create the file name and save and close the file
file_name = Application.WorksheetFunction.Trim(Str(i - 2) & "-BankConf-" & Workbooks(workbook_name).Sheets("Input").Cells(i, 6) & "-" & Workbooks(workbook_name).Sheets("Input").Cells(i, 7) & ".doc")
wDoc.SaveAs2 (Workbooks(workbook_name).Path & "\Confirmations\" & file_name)
wDoc.Close

'Closes all instances of word application (not needed anymore)
wordApp.Quit
Set wordApp = Nothing
''' Contact Umer Bhutto for any issues


End Sub

Sub bankadgmtemplate(i)

Dim wordApp As Word.Application
Dim wDoc As Word.Document

workbook_name = "Excel - Bank Confirmation_v1.0.xlsm"

' Create Word Application using Excel
Set wordApp = CreateObject("word.application")
Set wDoc = wordApp.Documents.Open(Workbooks(workbook_name).Path & "\Templates\bank_confirmation_template_difc.docm")
wordApp.Visible = False 'Used to keep the word application visible or not

''' Following are the controls that pertain to the Word Document by order of control
''' e.g. control 1 refers to date of letter etc.

' Control 1/21 - Date of Letter
wDoc.ContentControls(1).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 2)
wDoc.ContentControls(21).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 2)

' Control 2/14 - Bank Contact Name
wDoc.ContentControls(2).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 13)
wDoc.ContentControls(14).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 13)

' Control 3/15 - Bank Name
wDoc.ContentControls(3).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 7)
wDoc.ContentControls(15).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 7)

' Control 4/16 - Adress Line 1
wDoc.ContentControls(4).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 8)
wDoc.ContentControls(16).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 8)

'Control 5/17 - Adress Line 2
wDoc.ContentControls(5).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 9)
wDoc.ContentControls(17).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 9)

'Control 6/18 - PO Box
wDoc.ContentControls(6).Range.Text = "P.O. Box " & Workbooks(workbook_name).Sheets("Input").Cells(i, 10) & ", " & Workbooks(workbook_name).Sheets("Input").Cells(i, 11) & ", " & Workbooks(workbook_name).Sheets("Input").Cells(i, 12)
wDoc.ContentControls(18).Range.Text = "P.O. Box " & Workbooks(workbook_name).Sheets("Input").Cells(i, 10) & ", " & Workbooks(workbook_name).Sheets("Input").Cells(i, 11) & ", " & Workbooks(workbook_name).Sheets("Input").Cells(i, 12)

'Control 7/19 - Banker Telephone No.
wDoc.ContentControls(7).Range.Text = "Tel No: +" & Workbooks(workbook_name).Sheets("Input").Cells(i, 15)
wDoc.ContentControls(19).Range.Text = "Tel No: +" & Workbooks(workbook_name).Sheets("Input").Cells(i, 15)

'Control 8/20 - Email
wDoc.ContentControls(8).Range.Text = "Email: " & Workbooks(workbook_name).Sheets("Input").Cells(i, 14)
wDoc.ContentControls(20).Range.Text = "Email: " & Workbooks(workbook_name).Sheets("Input").Cells(i, 14)

'Control 9/22 - Client Name
wDoc.ContentControls(9).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 6)
wDoc.ContentControls(22).Range.Text = UCase(Workbooks(workbook_name).Sheets("Input").Cells(i, 6))

'Control 10/23/24/25/26 - Period End Date
wDoc.ContentControls(10).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 4)
wDoc.ContentControls(23).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 4)
wDoc.ContentControls(24).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 4)
wDoc.ContentControls(26).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 4)
''' This relates to the starting period and is based on the period end date
wDoc.ContentControls(25).Range.Text = DateValue("1 January " & DatePart("yyyy", Workbooks(workbook_name).Sheets("Input").Cells(i, 4)))

'Control 11/27 - PwC Contact
wDoc.ContentControls(11).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 1)
wDoc.ContentControls(27).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 1)

'Control 12/28 - PwC Contact Email
wDoc.ContentControls(12).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 3)
wDoc.ContentControls(28).Range.Text = Workbooks(workbook_name).Sheets("Input").Cells(i, 3)

'Control 13 - PwC Contact Telephone No.
wDoc.ContentControls(13).Range.Text = "+" & Workbooks(workbook_name).Sheets("Input").Cells(i, 5)

' Update Headers from page 3 to page 5
For j = 3 To 5
With wDoc.Sections(j).Headers(wdHeaderFooterPrimary).Range
    .InsertAfter Text:=vbCrLf & UCase(Workbooks(workbook_name).Sheets("Input").Cells(i, 6))
    .InsertAfter vbTab
    .InsertAfter Text:=vbCrLf & vbCrLf & UCase(Workbooks(workbook_name).Sheets("Input").Cells(i, 7))
    .InsertAfter vbTab
    .InsertAfter Text:=vbCrLf & vbCrLf & ("At close of business on 31 December " & DatePart("yyyy", Workbooks(workbook_name).Sheets("Input").Cells(i, 4)))
End With
Next j

' Close the header / footer pane
wDoc.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
wDoc.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument

' Create the file name and save and close the file
file_name = Application.WorksheetFunction.Trim(Str(i - 2) & "-BankConf-" & Workbooks(workbook_name).Sheets("Input").Cells(i, 6) & "-" & Workbooks(workbook_name).Sheets("Input").Cells(i, 7) & ".doc")
wDoc.SaveAs2 (Workbooks(workbook_name).Path & "\Confirmations\" & file_name)
wDoc.Close

'Closes all instances of word application (not needed anymore)
wordApp.Quit
Set wordApp = Nothing
''' Contact Umer Bhutto for any issues

End Sub

