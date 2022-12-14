Sub Parser()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    ActiveWindow.FreezePanes = False
    
    Dim Master As ThisWorkbook
    Dim ProductDatabase As Workbook
    Dim user As String
    Dim RemovedLocation As String
    Dim QuoteLocation As String
    Set Master = ThisWorkbook
    
    user = "Andrei"
    RemovedLocation = "C:\Users\ABB5350\OneDrive - Wolseley Group\Desktop\Product Cleaner\Removed\"
    QuoteLocation = "C:\Users\ABB5350\OneDrive - Wolseley Group\Desktop\Product Cleaner\Wolcen Quotes\"
    
    Master.Sheets("Data").Columns.EntireColumn.Hidden = False
    Master.Sheets("Data").Rows.EntireRow.Hidden = False

    If Master.Sheets("Data").AutoFilterMode = True Then
        Master.Sheets("Data").AutoFilterMode = False
    End If
    
    LastRowMaster = Master.Sheets("Data").Range("A" & Rows.Count).End(xlUp).Row
    
    If LastRowMaster > 1 Then
        Branchcode = InputBox("Branch Code:")
        AccountNumber = InputBox("Account Number:")
        Master.Sheets("Data").Columns("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
        Master.Sheets("Data").Columns("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
        Master.Sheets("Data").Range("C2:C" & LastRowMaster).Formula = "=TRUNC(@B:B,3)"
        Master.Sheets("Data").Range("C2:C" & LastRowMaster).Copy
        Master.Sheets("Data").Range("B2:B" & LastRowMaster).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Master.Sheets("Data").Range("C2:C" & LastRowMaster).ClearContents
        
        Set ProductDatabase = Workbooks.Open(Application.ActiveWorkbook.Path & "\Product Database\Product Database.xlsb", ReadOnly:=True)
        ProductDatabase.Sheets(1).Columns("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
        Master.Sheets("Data").Range("C2:C" & LastRowMaster).Formula = "=IFERROR(VLOOKUP(A2,'[Product Database.xlsb]Product File (Pyr1)'!$A:$I,9,FALSE),""No Data"")"
        Master.Sheets("Data").Range("C2:C" & LastRowMaster).Copy
        Master.Sheets("Data").Range("C2:C" & LastRowMaster).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Master.Activate
        Master.Sheets("Data").Range("A1:C" & LastRowMaster).AutoFilter
        Master.Sheets("Data").Range("A1:C" & LastRowMaster).AutoFilter Field:=3, Criteria1:=Array("EOL", "OBS", "No Data"), Operator:=xlFilterValues
        If Master.Sheets("Data").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Master.Sheets.Add.Name = "Removed Codes"
            Master.Sheets("Removed Codes").Range("A1").Value = "Wolseley Code"
            Master.Sheets("Removed Codes").Range("B1").Value = "Description"
            Master.Sheets("Removed Codes").Range("C1").Value = "Removal Reason"
            Master.Sheets("Data").Range("A2:A" & LastRowMaster).Copy
            Master.Sheets("Removed Codes").Range("A2").PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            Master.Sheets("Data").Range("C2:C" & LastRowMaster).Copy
            Master.Sheets("Removed Codes").Range("C2").PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            LastRowRemoved = Master.Sheets("Removed Codes").Range("A" & Rows.Count).End(xlUp).Row
            Master.Sheets("Removed Codes").Range("B2:B" & LastRowRemoved).Formula = "=IFERROR(VLOOKUP(@A:A,'[Product Database.xlsb]Product File (Pyr1)'!$A:$B,2,FALSE),""No Data"")"
            Master.Sheets("Removed Codes").Range("B2:B" & LastRowRemoved).Copy
            Master.Sheets("Removed Codes").Range("B2:B" & LastRowRemoved).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            Master.Sheets("Removed Codes").Copy
            ActiveWorkbook.SaveAs RemovedLocation & user & " - " & Branchcode & " - " & AccountNumber & " FPE Terms - Removed.xlsx"
            ActiveWorkbook.Close False
            answer = MsgBox("Do you want to Email the removed codes?", vbQuestion + vbYesNo + vbDefaultButton2, "Attention")
            If answer = vbYes Then
                EmailToSend = InputBox("Email:")
                
                Dim MyOutlook As Object
                Set MyOutlook = CreateObject("Outlook.Application")
                Dim MyMail As Object
                Set MyMail = MyOutlook.CreateItem(olMailItem)
                
                MyMail.To = EmailToSend
                MyMail.Subject = Branchcode & " - " & AccountNumber & " - OBS/EOL/No Data"
                MyMail.Body = "Recent email regarding setting up terms for " & AccountNumber & ", there are " & LastRowRemoved - 1 & " product codes which cannot be uploaded due to being OBS/EOL or Incorrect Code."
                
                Attached_File = RemovedLocation & user & " - " & Branchcode & " - " & AccountNumber & " FPE Terms - Removed.xlsx"
                MyMail.Attachments.Add Attached_File
                
                MyMail.Send
                
                MsgBox "Email sent to " & EmailToSend & ", " & LastRowRemoved - 1 & " products were removed!."
            Else
                MsgBox LastRowRemoved - 1 & " Products were Removed! You have copy in your removals folder!"
            End If
            Master.Sheets("Removed Codes").Delete
            Master.Sheets("Data").AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
            Master.Sheets("Data").AutoFilter.ShowAllData
        End If
        ProductDatabase.Close False
        Master.Sheets("Data").AutoFilter.ShowAllData
        LastRowMaster = Master.Sheets("Data").Range("A" & Rows.Count).End(xlUp).Row
        If LastRowMaster > 1 Then
            Master.Sheets.Add.Name = "Temporary"
            
            Master.Sheets("Temporary").Range("A1").Value = "Q"
            Master.Sheets("Temporary").Range("B1").Value = "Quote"
            Master.Sheets("Temporary").Range("C1").Value = Branchcode & "/" & AccountNumber
            Master.Sheets("Temporary").Range("D1").Value = "Customer - " & AccountNumber
            Master.Sheets("Temporary").Range("E1").Value = AccountNumber
            Master.Sheets("Temporary").Range("F1").Value = "q0"
            Master.Sheets("Temporary").Range("G1").Value = "39352"
            Master.Sheets("Temporary").Range("H1").Value = "14:45:46"
            Master.Sheets("Temporary").Range("I1").Value = "almir"
            
            Master.Sheets("Temporary").Range("A2").Value = "Item"
            Master.Sheets("Temporary").Range("B2").Value = "Description"
            Master.Sheets("Temporary").Range("C2").Value = "Discount 1"
            Master.Sheets("Temporary").Range("D2").Value = "Discount 2"
            Master.Sheets("Temporary").Range("E2").Value = "Quantity"
            Master.Sheets("Temporary").Range("F2").Value = "Per(Qty)"
            Master.Sheets("Temporary").Range("G2").Value = "Price"
            Master.Sheets("Temporary").Range("H2").Value = "Per(Price)"
            
            Master.Sheets("Data").Range("A2:A" & LastRowMaster).Copy
            Master.Sheets("Temporary").Range("A3").PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            Master.Sheets("Data").Range("B2:B" & LastRowMaster).Copy
            Master.Sheets("Temporary").Range("G3").PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            LastRowTemp = Master.Sheets("Temporary").Range("A" & Rows.Count).End(xlUp).Row
            Master.Sheets("Temporary").Range("E3:E" & LastRowTemp).Value = "1"
            
            Master.Sheets("Temporary").Copy
            ActiveWorkbook.SaveAs QuoteLocation & Branchcode & " - " & AccountNumber & " - " & user & " - FPE Terms.csv", FileFormat:=xlCSV
            ActiveWorkbook.Close True
            

            Master.Sheets("Data").Range("A2:C" & LastRowMaster).ClearContents
            Master.Sheets("Temporary").Delete
            MsgBox "Done!!!!"
        Else
            MsgBox "There are no more codes left. All codes appeared to be EOL/OBS/No Data (If Not sure, Contact: andrei.polin@wolseley.co.uk via Teams)"
        End If
        
    
    Else
        MsgBox "Please Input WUKCodes and Prices into parser, Then Try again."
    End If
    If Master.Sheets("Data").AutoFilterMode = True Then
        Master.Sheets("Data").AutoFilterMode = False
    End If
    Master.Save
End Sub
