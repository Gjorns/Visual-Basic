Sub margin_sheetv7()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Dim MarginSheet As Workbook
    Dim PriceFile As Workbook
    Dim PriceFileCopper As Workbook
    Dim ControlPanel As Workbook
    Dim Estimator As String
    Dim AccountID As String
    Dim QuoteID As String
    Dim SavePath As String
    Dim NacetPath As String
    Dim UsePriceFile As Boolean
    
    Set MarginSheet = ActiveWorkbook
    NacetPath = "\\WUKRLS00FP001\RLS_Data\Departments\NACET\Price File Maintenance\"
    
    Estimator = "Andrei Polin"
    UsePriceFile = True
    
    If MarginSheet.Sheets(1).Range("A1") = "Branch" Then
        AccountID = MarginSheet.Sheets(1).Range("B2").Value
        QuoteID = MarginSheet.Sheets(1).Range("C5").Value
        For Each Sheet In MarginSheet.Worksheets
             If Sheet.Name = "Finance Margins" Then
                  Sheet.Delete
             End If
        Next Sheet
        MarginSheet.Sheets(1).Name = "PricePoint"
        MarginSheet.Worksheets.Add(After:=Sheets(Sheets.Count)).Name = "Finance Margins"
        LastRowPricePoint = MarginSheet.Sheets("PricePoint").Range("A" & Rows.Count).End(xlUp).Row
        
        MarginSheet.Sheets("Finance Margins").Range("A9").Value = "Finance Margin Sheet - Generated via PricePoint"
        MarginSheet.Sheets("Finance Margins").Range("I9").Value = "Finance Margins"
        MarginSheet.Sheets("Finance Margins").Range("P9").Value = "Basket Totals"
        MarginSheet.Sheets("Finance Margins").Range("S9").Value = "Cost Plus Calculator"
        MarginSheet.Sheets("Finance Margins").Range("V9").Value = "Discount Off Calculator"
        MarginSheet.Sheets("Finance Margins").Range("Y9").Value = "Product Details"
        MarginSheet.Sheets("Finance Margins").Range("AB9").Value = "Pricing Details"
        MarginSheet.Sheets("Finance Margins").Range("A1").Value = "Account ID:"
        MarginSheet.Sheets("Finance Margins").Range("A2").Value = "Account Name:"
        MarginSheet.Sheets("Finance Margins").Range("A3").Value = "Sheet Generator:"
        MarginSheet.Sheets("Finance Margins").Range("A4").Value = "Quote ID:"
        MarginSheet.Sheets("Finance Margins").Range("A5").Value = "Date Generated:"
        MarginSheet.Sheets("Finance Margins").Range("F1").Value = "Total Sell:"
        MarginSheet.Sheets("Finance Margins").Range("F2").Value = "Total Cost:"
        MarginSheet.Sheets("Finance Margins").Range("F3").Value = "Basket Margin:"
        MarginSheet.Sheets("Finance Margins").Range("F4").Value = "Total Profit:"
        MarginSheet.Sheets("Finance margins").Range("A10").Value = "Estimator Comments"
        MarginSheet.Sheets("Finance Margins").Range("B10").Value = "Supplier Part No."
        MarginSheet.Sheets("Finance Margins").Range("C10").Value = "Wolseley Code"
        MarginSheet.Sheets("Finance Margins").Range("D10").Value = "Product Description"
        MarginSheet.Sheets("Finance Margins").Range("E10").Value = "Terms Price (£)"
        MarginSheet.Sheets("Finance Margins").Range("F10").Value = "Terms Discount (%)"
        MarginSheet.Sheets("Finance Margins").Range("G10").Value = "Rebated Price (£)"
        MarginSheet.Sheets("Finance Margins").Range("H10").Value = "Quantity"
        MarginSheet.Sheets("Finance Margins").Range("I10").Value = "Margin (%)"
        MarginSheet.Sheets("Finance Margins").Range("J10").Value = "Rebated Margin (%)"
        MarginSheet.Sheets("Finance Margins").Range("K10").Value = "Cust Rebate (%)"
        MarginSheet.Sheets("Finance Margins").Range("L10").Value = "Support (£)"
        MarginSheet.Sheets("Finance Margins").Range("M10").Value = "WUK Support Ref."
        MarginSheet.Sheets("Finance Margins").Range("N10").Value = "Nett Cost (£)"
        MarginSheet.Sheets("Finance Margins").Range("O10").Value = "Nett Cost NRA (£)"
        MarginSheet.Sheets("Finance Margins").Range("P10").Value = "Total Sell (£)"
        MarginSheet.Sheets("Finance Margins").Range("Q10").Value = "Total Cost (£)"
        MarginSheet.Sheets("Finance Margins").Range("R10").Value = "Total Profit (£)"
        MarginSheet.Sheets("Finance Margins").Range("S10").Value = "Set Margin (%)"
        MarginSheet.Sheets("Finance Margins").Range("T10").Value = "New Sell (£)"
        MarginSheet.Sheets("Finance Margins").Range("U10").Value = "New Discount (%)"
        MarginSheet.Sheets("Finance Margins").Range("V10").Value = "Set Discount (%)"
        MarginSheet.Sheets("Finance Margins").Range("W10").Value = "New Sell (£)"
        MarginSheet.Sheets("Finance Margins").Range("X10").Value = "New Margin (%)"
        MarginSheet.Sheets("Finance Margins").Range("Y10").Value = "Wolseley SPG"
        MarginSheet.Sheets("Finance Margins").Range("Z10").Value = "Supplier Name"
        MarginSheet.Sheets("Finance Margins").Range("AA10").Value = "Supplier HoS"
        MarginSheet.Sheets("Finance Margins").Range("AB10").Value = "Current Invoice (£)"
        MarginSheet.Sheets("Finance Margins").Range("AC10").Value = "Future Invoice (£)"
        MarginSheet.Sheets("Finance Margins").Range("AD10").Value = "Invoice Change (%)"
        MarginSheet.Sheets("Finance Margins").Range("AE10").Value = "Current Trade (£)"
        MarginSheet.Sheets("Finance Margins").Range("AF10").Value = "Future Trade (£)"
        MarginSheet.Sheets("Finance Margins").Range("AG10").Value = "Trade Change (%)"
        MarginSheet.Sheets("Finance Margins").Range("AH10").Value = "Current Branch (£)"
        MarginSheet.Sheets("Finance Margins").Range("AI10").Value = "Future Branch (£)"
        MarginSheet.Sheets("Finance Margins").Range("AJ10").Value = "Branch Change (%)"
        MarginSheet.Sheets("Finance Margins").Range("AK10").Value = "Current SLP (£)"
        MarginSheet.Sheets("Finance Margins").Range("AL10").Value = "Future SLP (£)"
        MarginSheet.Sheets("Finance Margins").Range("AM10").Value = "SLP Change (%)"
        MarginSheet.Sheets("Finance Margins").Range("AN10").Value = "Current Real Cost (£)"
        MarginSheet.Sheets("Finance Margins").Range("AO10").Value = "Future Real Cost (£)"
        MarginSheet.Sheets("Finance Margins").Range("AP10").Value = "Real Change (%)"
        MarginSheet.Sheets("Finance Margins").Range("AQ10").Value = "Future Date"
        MarginSheet.Sheets("Finance Margins").Range("AR10").Value = "Rebate Impacted"
        MarginSheet.Sheets("Finance Margins").Range("AS10").Value = "Product Lifecycle"
        MarginSheet.Sheets("Finance Margins").Range("AT10").Value = "Product Narrative"
        MarginSheet.Sheets("Finance Margins").Range("AU10").Value = "Last Sale Date"
        MarginSheet.Sheets("Finance Margins").Range("AV10").Value = "Last Sale Price (£)"
        MarginSheet.Sheets("Finance Margins").Range("H1").Value = "No. of Increases:"
        MarginSheet.Sheets("Finance Margins").Range("H2").Value = "No. of OBS/EOL:"
        MarginSheet.Sheets("Finance Margins").Range("H3").Value = "No. of Loss Making:"
        MarginSheet.Sheets("Finance Margins").Range("H4").Value = "No. of Supports:"
    
        MarginSheet.Sheets("PricePoint").Range("H15:H" & LastRowPricePoint).Copy
        MarginSheet.Sheets("Finance Margins").Range("C11").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Lastrowfinancemargins = MarginSheet.Sheets("Finance Margins").Range("C" & Rows.Count).End(xlUp).Row
        
        MarginSheet.Sheets("Finance Margins").Range("B11:B" & Lastrowfinancemargins).Formula = "=INDEX(PricePoint!E:E,MATCH('Finance Margins'!C11,PricePoint!H:H,0))"
        MarginSheet.Sheets("Finance Margins").Range("D11:D" & Lastrowfinancemargins).Formula = "=INDEX(PricePoint!I:I,MATCH('Finance Margins'!C11,PricePoint!H:H,0))"
        MarginSheet.Sheets("Finance Margins").Range("E11:E" & Lastrowfinancemargins).Formula = "=INDEX(PricePoint!AB:AB,MATCH('Finance Margins'!C11,PricePoint!H:H,0))"
        MarginSheet.Sheets("Finance Margins").Range("H11:H" & Lastrowfinancemargins).Value = "1"
        MarginSheet.Sheets("Finance Margins").Range("K11:K" & Lastrowfinancemargins).Formula = "=INDEX(PricePoint!AQ:AQ,MATCH('Finance Margins'!C11,PricePoint!H:H,0))"
        MarginSheet.Sheets("Finance Margins").Range("L11:L" & Lastrowfinancemargins).Formula = "=INDEX(PricePoint!AN:AN,MATCH('Finance Margins'!C11,PricePoint!H:H,0))"
        MarginSheet.Sheets("Finance Margins").Range("M11:M" & Lastrowfinancemargins).Formula = "=IF(INDEX(PricePoint!AJ:AJ,MATCH('Finance Margins'!C11,PricePoint!H:H,0))=0,"""",INDEX(PricePoint!AJ:AJ,MATCH('Finance Margins'!C11,PricePoint!H:H,0)))"
        MarginSheet.Sheets("Finance Margins").Range("Y11:Y" & Lastrowfinancemargins).Formula = "=INDEX(PricePoint!G:G,MATCH('Finance Margins'!C11,PricePoint!H:H,0))"
        MarginSheet.Sheets("Finance Margins").Range("Z11:Z" & Lastrowfinancemargins).Formula = "=INDEX(PricePoint!D:D,MATCH('Finance Margins'!C11,PricePoint!H:H,0))"
        MarginSheet.Sheets("Finance Margins").Range("AB11:AB" & Lastrowfinancemargins).Formula = "=INDEX(PricePoint!Q:Q,MATCH('Finance Margins'!C11,PricePoint!H:H,0))"
        MarginSheet.Sheets("Finance Margins").Range("AE11:AE" & Lastrowfinancemargins).Formula = "=INDEX(PricePoint!O:O,MATCH('Finance Margins'!C11,PricePoint!H:H,0))"
        MarginSheet.Sheets("Finance Margins").Range("AH11:AH" & Lastrowfinancemargins).Formula = "=INDEX(PricePoint!P:P,MATCH('Finance Margins'!C11,PricePoint!H:H,0))"
        MarginSheet.Sheets("Finance Margins").Range("AK11:AK" & Lastrowfinancemargins).Formula = "=INDEX(PricePoint!R:R,MATCH('Finance Margins'!C11,PricePoint!H:H,0))"
        MarginSheet.Sheets("Finance Margins").Range("AN11:AN" & Lastrowfinancemargins).Formula = "=INDEX(PricePoint!S:S,MATCH('Finance Margins'!C11,PricePoint!H:H,0))"
        MarginSheet.Sheets("Finance Margins").Range("AC11:AC" & Lastrowfinancemargins).Formula = "=IF(INDEX(PricePoint!W:W,MATCH('Finance Margins'!C11,PricePoint!H:H,0))=0,@AB:AB,INDEX(PricePoint!W:W,MATCH('Finance Margins'!C11,PricePoint!H:H,0)))"
        MarginSheet.Sheets("Finance Margins").Range("AF11:AF" & Lastrowfinancemargins).Formula = "=IF(INDEX(PricePoint!U:U,MATCH('Finance Margins'!C11,PricePoint!H:H,0))=0,@AE:AE,INDEX(PricePoint!U:U,MATCH('Finance Margins'!C11,PricePoint!H:H,0)))"
        MarginSheet.Sheets("Finance Margins").Range("AI11:AI" & Lastrowfinancemargins).Formula = "=IF(INDEX(PricePoint!V:V,MATCH('Finance Margins'!C11,PricePoint!H:H,0))=0,@AH:AH,INDEX(PricePoint!V:V,MATCH('Finance Margins'!C11,PricePoint!H:H,0)))"
        MarginSheet.Sheets("Finance Margins").Range("AL11:AL" & Lastrowfinancemargins).Formula = "=IF(INDEX(PricePoint!X:X,MATCH('Finance Margins'!C11,PricePoint!H:H,0))=0,@AK:AK,INDEX(PricePoint!X:X,MATCH('Finance Margins'!C11,PricePoint!H:H,0)))"
        MarginSheet.Sheets("Finance Margins").Range("AO11:AO" & Lastrowfinancemargins).Formula = "=IF(INDEX(PricePoint!Y:Y,MATCH('Finance Margins'!C11,PricePoint!H:H,0))=0,@AN:AN,INDEX(PricePoint!Y:Y,MATCH('Finance Margins'!C11,PricePoint!H:H,0)))"
        MarginSheet.Sheets("Finance Margins").Range("AQ11:AQ" & Lastrowfinancemargins).Formula = "=IF(INDEX(PricePoint!T:T,MATCH('Finance Margins'!C11,PricePoint!H:H,0))=0,""No Increases"",INDEX(PricePoint!T:T,MATCH('Finance Margins'!C11,PricePoint!H:H,0)))"
        MarginSheet.Sheets("Finance Margins").Range("AR11:AR" & Lastrowfinancemargins).Formula = "=IF(INDEX(PricePoint!F:F,MATCH('Finance Margins'!C11,PricePoint!H:H,0))=0,""No Data"",INDEX(PricePoint!F:F,MATCH('Finance Margins'!C11,PricePoint!H:H,0)))"
        MarginSheet.Sheets("Finance Margins").Range("AS11:AS" & Lastrowfinancemargins).Formula = "=IF(INDEX(PricePoint!J:J,MATCH('Finance Margins'!C11,PricePoint!H:H,0))=0,""No Data"",INDEX(PricePoint!J:J,MATCH('Finance Margins'!C11,PricePoint!H:H,0)))"
        MarginSheet.Sheets("Finance Margins").Range("AT11:AT" & Lastrowfinancemargins).Formula = "=IF(INDEX(PricePoint!M:M,MATCH('Finance Margins'!C11,PricePoint!H:H,0))=0,""No Data"",INDEX(PricePoint!M:M,MATCH('Finance Margins'!C11,PricePoint!H:H,0)))"
        MarginSheet.Sheets("Finance Margins").Range("AU11:AU" & Lastrowfinancemargins).Formula = "=IF(INDEX(PricePoint!K:K,MATCH('Finance Margins'!C11,PricePoint!H:H,0))=0,""No Sales"",INDEX(PricePoint!K:K,MATCH('Finance Margins'!C11,PricePoint!H:H,0)))"
        MarginSheet.Sheets("Finance Margins").Range("AV11:AV" & Lastrowfinancemargins).Formula = "=IF(INDEX(PricePoint!L:L,MATCH('Finance Margins'!C11,PricePoint!H:H,0))=0,""No Sales"",INDEX(PricePoint!L:L,MATCH('Finance Margins'!C11,PricePoint!H:H,0)))"
        MarginSheet.Sheets("Finance Margins").Range("B1").Value = MarginSheet.Sheets("PricePoint").Range("B2").Value
        MarginSheet.Sheets("Finance Margins").Range("B2").Value = MarginSheet.Sheets("PricePoint").Range("B3").Value
        MarginSheet.Sheets("Finance Margins").Range("B3").Value = Estimator
        MarginSheet.Sheets("Finance Margins").Range("B4").Value = MarginSheet.Sheets("PricePoint").Range("B4").Value
        MarginSheet.Sheets("Finance Margins").Range("B5").Value = Now()
        
        MarginSheet.Sheets("Finance Margins").Range("A11:AV" & Lastrowfinancemargins).Copy
        MarginSheet.Sheets("Finance Margins").Range("A11:AV" & Lastrowfinancemargins).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        
        MarginSheet.Sheets("Finance Margins").Range("F11:F" & Lastrowfinancemargins).Formula = "=(@AE:AE-@E:E)/@AE:AE"
        MarginSheet.Sheets("Finance Margins").Range("G11:G" & Lastrowfinancemargins).Formula = "=@E:E*(1-@K:K)"
        MarginSheet.Sheets("Finance Margins").Range("I11:I" & Lastrowfinancemargins).Formula = "=(@E:E-@O:O)/@E:E"
        MarginSheet.Sheets("Finance Margins").Range("J11:J" & Lastrowfinancemargins).Formula = "=(@G:G-@O:O)/@G:G"
        MarginSheet.Sheets("Finance Margins").Range("N11:N" & Lastrowfinancemargins).Formula = "=@AN:AN"
        MarginSheet.Sheets("Finance Margins").Range("O11:O" & Lastrowfinancemargins).Formula = "=IF(@AR:AR=""Y"",@N:N,@N:N-@L:L)"
        MarginSheet.Sheets("Finance Margins").Range("P11:P" & Lastrowfinancemargins).Formula = "=@G:G*@H:H"
        MarginSheet.Sheets("Finance Margins").Range("Q11:Q" & Lastrowfinancemargins).Formula = "=@O:O*@H:H"
        MarginSheet.Sheets("Finance Margins").Range("R11:R" & Lastrowfinancemargins).Formula = "=@P:P-@Q:Q"
        MarginSheet.Sheets("Finance Margins").Range("T11:T" & Lastrowfinancemargins).Formula = "=@O:O/(1-@S:S)"
        MarginSheet.Sheets("Finance Margins").Range("U11:U" & Lastrowfinancemargins).Formula = "=(@T:T/@AE:AE)-1"
        MarginSheet.Sheets("Finance Margins").Range("W11:W" & Lastrowfinancemargins).Formula = "=@AE:AE*(1-@V:V)"
        MarginSheet.Sheets("Finance Margins").Range("X11:X" & Lastrowfinancemargins).Formula = "=(@W:W-@O:O)/@W:W"
        MarginSheet.Sheets("Finance Margins").Range("AD11:AD" & Lastrowfinancemargins).Formula = "=IFERROR((@AC:AC/@AB:AB)-1,"""")"
        MarginSheet.Sheets("Finance Margins").Range("AG11:AG" & Lastrowfinancemargins).Formula = "=IFERROR((@AF:AF/@AE:AE)-1,"""")"
        MarginSheet.Sheets("Finance Margins").Range("AJ11:AJ" & Lastrowfinancemargins).Formula = "=IFERROR((@AI:AI/@AH:AH)-1,"""")"
        MarginSheet.Sheets("Finance Margins").Range("AM11:AM" & Lastrowfinancemargins).Formula = "=IFERROR((@AL:AL/@AK:AK)-1,"""")"
        MarginSheet.Sheets("Finance Margins").Range("AP11:AP" & Lastrowfinancemargins).Formula = "=IFERROR((@AO:AO/@AN:AN)-1,"""")"
        MarginSheet.Sheets("Finance Margins").Range("G1").Formula = "=SUM(P11:P" & Lastrowfinancemargins & ")"
        MarginSheet.Sheets("Finance Margins").Range("G2").Formula = "=SUM(Q11:Q" & Lastrowfinancemargins & ")"
        MarginSheet.Sheets("Finance Margins").Range("G3").Formula = "=IFERROR((G1-G2)/G1,0)"
        MarginSheet.Sheets("Finance Margins").Range("G4").Formula = "=G1*G3"
        MarginSheet.Sheets("Finance Margins").Range("I1").Formula = "=COUNTIF(AQ11:AQ" & Lastrowfinancemargins & ",""<>No Increases"") & "" Products"""
        MarginSheet.Sheets("Finance Margins").Range("I2").Formula = "=COUNTIF(AS11:AS" & Lastrowfinancemargins & ",""*EOL*"")+COUNTIF(AS11:AS" & Lastrowfinancemargins & ",""*OBS*"") & "" Products"""
        MarginSheet.Sheets("Finance Margins").Range("I3").Formula = "=COUNTIF(J11:J" & Lastrowfinancemargins & ",""<0"") & "" Products"""
        MarginSheet.Sheets("Finance Margins").Range("I4").Formula = "=COUNTIF(L11:L" & Lastrowfinancemargins & ","">0"") & "" Products"""
        MarginSheet.Sheets("Finance Margins").Range("C5").Formula = "=TRIM(RIGHT(SUBSTITUTE(B4,""."",REPT("" "",100)),100))"
        
        MarginSheet.Sheets("Finance Margins").Range("A9:H9,I9:O9,P9:R9,S9:U9,V9:X9,Y9:AA9,AB9:AV9").Merge
        MarginSheet.Sheets("Finance Margins").Range("E11:E" & Lastrowfinancemargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("G11:G" & Lastrowfinancemargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("L11:L" & Lastrowfinancemargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("N11:N" & Lastrowfinancemargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("O11:O" & Lastrowfinancemargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("P11:P" & Lastrowfinancemargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("Q11:Q" & Lastrowfinancemargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("R11:R" & Lastrowfinancemargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("R11:R" & Lastrowfinancemargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("T11:T" & Lastrowfinancemargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("W11:W" & Lastrowfinancemargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("AB11:AB" & Lastrowfinancemargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("AC11:AC" & Lastrowfinancemargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("AE11:AE" & Lastrowfinancemargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("AF11:AF" & Lastrowfinancemargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("AH11:AH" & Lastrowfinancemargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("AI11:AI" & Lastrowfinancemargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("AK11:AK" & Lastrowfinancemargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("AL11:AL" & Lastrowfinancemargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("AN11:AN" & Lastrowfinancemargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("AO11:AO" & Lastrowfinancemargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("AV11:AV" & Lastrowfinancemargins).NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("G1").NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("G2").NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("G4").NumberFormat = "£#,##0.00"
        MarginSheet.Sheets("Finance Margins").Range("F11:F" & Lastrowfinancemargins).NumberFormat = "0.00%"
        MarginSheet.Sheets("Finance Margins").Range("I11:I" & Lastrowfinancemargins).NumberFormat = "0.00%"
        MarginSheet.Sheets("Finance Margins").Range("J11:J" & Lastrowfinancemargins).NumberFormat = "0.00%"
        MarginSheet.Sheets("Finance Margins").Range("K11:K" & Lastrowfinancemargins).NumberFormat = "0.00%"
        MarginSheet.Sheets("Finance Margins").Range("S11:S" & Lastrowfinancemargins).NumberFormat = "0.00%"
        MarginSheet.Sheets("Finance Margins").Range("V11:V" & Lastrowfinancemargins).NumberFormat = "0.00%"
        MarginSheet.Sheets("Finance Margins").Range("AD11:AD" & Lastrowfinancemargins).NumberFormat = "0.00%"
        MarginSheet.Sheets("Finance Margins").Range("AG11:AG" & Lastrowfinancemargins).NumberFormat = "0.00%"
        MarginSheet.Sheets("Finance Margins").Range("AJ11:AJ" & Lastrowfinancemargins).NumberFormat = "0.00%"
        MarginSheet.Sheets("Finance Margins").Range("AM11:AM" & Lastrowfinancemargins).NumberFormat = "0.00%"
        MarginSheet.Sheets("Finance Margins").Range("AP11:AP" & Lastrowfinancemargins).NumberFormat = "0.00%"
        MarginSheet.Sheets("Finance Margins").Range("X11:X" & Lastrowfinancemargins).NumberFormat = "0.00%"
        MarginSheet.Sheets("Finance Margins").Range("U11:U" & Lastrowfinancemargins).NumberFormat = "0.00%"
        MarginSheet.Sheets("Finance Margins").Range("G3").NumberFormat = "0.00%"
        MarginSheet.Sheets("Finance Margins").Range("AQ11:AQ" & Lastrowfinancemargins).NumberFormat = "dd/mm/yyyy"
        MarginSheet.Sheets("Finance Margins").Range("AU11:AU" & Lastrowfinancemargins).NumberFormat = "dd/mm/yyyy"
        MarginSheet.Sheets("Finance Margins").Range("A9:AV" & Lastrowfinancemargins).HorizontalAlignment = xlCenter
        MarginSheet.Sheets("Finance Margins").Range("A9:AV" & Lastrowfinancemargins).VerticalAlignment = xlCenter
        MarginSheet.Sheets("Finance Margins").Range("G1:G4,I1:I4").HorizontalAlignment = xlCenter
        MarginSheet.Sheets("Finance Margins").Range("G1:G4,I1:I4").VerticalAlignment = xlCenter
        MarginSheet.Sheets("Finance Margins").Range("A9:AV" & Lastrowfinancemargins).Borders.LineStyle = xlContinuous
        MarginSheet.Sheets("Finance Margins").Range("F1:I4").Borders.LineStyle = xlContinuous
        MarginSheet.Sheets("Finance Margins").Range("A1:B5").Borders.LineStyle = xlContinuous
        MarginSheet.Sheets("Finance Margins").Cells.Font.Name = "Arial"
        MarginSheet.Sheets("Finance Margins").Cells.Font.Size = "10"
        MarginSheet.Sheets("Finance Margins").Range("A1:A5,F1:F4,A9:AV10,H1:H4").Font.Bold = True
        MarginSheet.Sheets("Finance Margins").Range("I11:I" & Lastrowfinancemargins).Font.Bold = True
        MarginSheet.Sheets("Finance Margins").Range("J11:J" & Lastrowfinancemargins).Font.Bold = True
        MarginSheet.Sheets("Finance Margins").Range("K11:K" & Lastrowfinancemargins).Font.Bold = True
        MarginSheet.Sheets("Finance Margins").Range("U11:U" & Lastrowfinancemargins).Font.Bold = True
        MarginSheet.Sheets("Finance Margins").Range("X11:X" & Lastrowfinancemargins).Font.Bold = True
        MarginSheet.Sheets("Finance Margins").Range("AD11:AD" & Lastrowfinancemargins).Font.Bold = True
        MarginSheet.Sheets("Finance Margins").Range("AG11:AG" & Lastrowfinancemargins).Font.Bold = True
        MarginSheet.Sheets("Finance Margins").Range("AJ11:AJ" & Lastrowfinancemargins).Font.Bold = True
        MarginSheet.Sheets("Finance Margins").Range("AM11:AM" & Lastrowfinancemargins).Font.Bold = True
        MarginSheet.Sheets("Finance Margins").Range("AP11:AP" & Lastrowfinancemargins).Font.Bold = True
        MarginSheet.Sheets("Finance Margins").Range("A10:AV10").Font.Underline = True
        ActiveWindow.DisplayGridlines = False
        MarginSheet.Sheets("Finance Margins").Range("A10:AV10,A1:B5,F1:F4,H1:H4").Interior.Color = RGB(150, 194, 230)
        MarginSheet.Sheets("Finance Margins").Range("A9:H9,G1:G4,I1:I4").Interior.Color = RGB(237, 237, 237)
        MarginSheet.Sheets("Finance Margins").Range("A11:H" & Lastrowfinancemargins).Interior.Color = RGB(237, 237, 237)
        MarginSheet.Sheets("Finance Margins").Range("I9:O9").Interior.Color = RGB(198, 224, 180)
        MarginSheet.Sheets("Finance Margins").Range("I11:O" & Lastrowfinancemargins).Interior.Color = RGB(198, 224, 180)
        MarginSheet.Sheets("Finance Margins").Range("P9:R9").Interior.Color = RGB(146, 208, 80)
        MarginSheet.Sheets("Finance Margins").Range("P11:R" & Lastrowfinancemargins).Interior.Color = RGB(146, 208, 80)
        MarginSheet.Sheets("Finance Margins").Range("S9:U9").Interior.Color = RGB(255, 204, 204)
        MarginSheet.Sheets("Finance Margins").Range("S11:U" & Lastrowfinancemargins).Interior.Color = RGB(255, 204, 204)
        MarginSheet.Sheets("Finance Margins").Range("V9:X9").Interior.Color = RGB(248, 203, 173)
        MarginSheet.Sheets("Finance Margins").Range("V11:X" & Lastrowfinancemargins).Interior.Color = RGB(248, 203, 173)
        MarginSheet.Sheets("Finance Margins").Range("Y9:AV9").Interior.Color = RGB(242, 242, 242)
        MarginSheet.Sheets("Finance Margins").Range("Y11:AV" & Lastrowfinancemargins).Interior.Color = RGB(242, 242, 242)
        MarginSheet.Sheets("Finance Margins").Range("A10:AV" & Lastrowfinancemargins).AutoFilter
        MarginSheet.Sheets("Finance Margins").Columns("A:AV").AutoFit
        MarginSheet.Sheets("Finance Margins").Range("E10:E" & Lastrowfinancemargins).BorderAround ColorIndex:=31, Weight:=xlThick
        MarginSheet.Sheets("Finance Margins").Range("S10:S" & Lastrowfinancemargins).BorderAround ColorIndex:=31, Weight:=xlThick
        MarginSheet.Sheets("Finance Margins").Range("V10:V" & Lastrowfinancemargins).BorderAround ColorIndex:=31, Weight:=xlThick
        MarginSheet.Sheets("Finance Margins").Range("H10:H" & Lastrowfinancemargins).BorderAround ColorIndex:=31, Weight:=xlThick
        MarginSheet.Sheets("Finance Margins").Range("E11:E" & Lastrowfinancemargins).Interior.Color = RGB(255, 242, 204)
        MarginSheet.Sheets("Finance Margins").Range("S11:S" & Lastrowfinancemargins).Interior.Color = RGB(255, 242, 204)
        MarginSheet.Sheets("Finance Margins").Range("V11:V" & Lastrowfinancemargins).Interior.Color = RGB(255, 242, 204)
        MarginSheet.Sheets("Finance Margins").Range("H11:H" & Lastrowfinancemargins).Interior.Color = RGB(255, 242, 204)
        MarginSheet.Sheets("Finance Margins").Range("C11:C" & Lastrowfinancemargins).TextToColumns Destination:=Range("C11"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    
        MarginSheet.Sheets("Finance Margins").Range("A10:AV" & Lastrowfinancemargins).AutoFilter Field:=44, Criteria1:="Y"
        MarginSheet.Sheets("Finance Margins").Range("A10:AV" & Lastrowfinancemargins).AutoFilter Field:=13, Criteria1:="<>"
        MarginSheet.Sheets("Finance Margins").Range("A10:AV" & Lastrowfinancemargins).AutoFilter Field:=39, Criteria1:=">0"
        If MarginSheet.Sheets("Finance Margins").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            MarginSheet.Sheets("Finance Margins").Range("AW11:AW" & Lastrowfinancemargins).SpecialCells(xlCellTypeVisible).Formula = "=@AO:AO-@L:L"
            MarginSheet.Sheets("Finance Margins").AutoFilter.ShowAllData
            MarginSheet.Sheets("Finance Margins").Range("AW11:AW" & Lastrowfinancemargins).Copy
            MarginSheet.Sheets("Finance Margins").Range("AW11:AW" & Lastrowfinancemargins).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            MarginSheet.Sheets("Finance Margins").Range("A10:AT" & Lastrowfinancemargins).AutoFilter Field:=44, Criteria1:="Y"
            MarginSheet.Sheets("Finance Margins").Range("A10:AT" & Lastrowfinancemargins).AutoFilter Field:=13, Criteria1:="<>"
            MarginSheet.Sheets("Finance Margins").Range("A10:AT" & Lastrowfinancemargins).AutoFilter Field:=39, Criteria1:=">0"
            MarginSheet.Sheets("Finance Margins").Range("AO11:AO" & Lastrowfinancemargins).SpecialCells(xlCellTypeVisible).Formula = "=@AW:AW"
            MarginSheet.Sheets("Finance Margins").AutoFilter.ShowAllData
            MarginSheet.Sheets("Finance Margins").Range("AO11:AO" & Lastrowfinancemargins).Copy
            MarginSheet.Sheets("Finance Margins").Range("AO11:AO" & Lastrowfinancemargins).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            MarginSheet.Sheets("Finance Margins").Range("AW11:AW" & Lastrowfinancemargins).ClearContents
        Else
            MarginSheet.Sheets("Finance Margins").AutoFilter.ShowAllData
        End If
        
        MarginSheet.Sheets("Finance Margins").Range("A10:AV" & Lastrowfinancemargins).AutoFilter Field:=44, Criteria1:="Y"
        If MarginSheet.Sheets("Finance Margins").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            MarginSheet.Sheets("Finance Margins").Range("L11:L" & Lastrowfinancemargins).SpecialCells(xlCellTypeVisible).Interior.Color = RGB(255, 102, 0)
            MarginSheet.Sheets("Finance Margins").AutoFilter.ShowAllData
        Else
            MarginSheet.Sheets("Finance Margins").AutoFilter.ShowAllData
        End If
        
        If UsePriceFile = True Then
            Set ControlPanel = Workbooks.Open("\\WUKRLS00FP001\RLS_Data\Departments\NACET\Price File Maintenance\Estimators\Andrei Polin\Margin Sheet Database.xlsm", ReadOnly:=True)
            MarginSheet.Sheets("Finance Margins").Range("C1").Formula = "=IFERROR(VLOOKUP(IF(B1=""7958X28"",PricePoint!B10,PricePoint!B2),'[" & ControlPanel.Name & "]DATA'!$A:$C,3,FALSE),""No Price File"")"
            PriceFilePath = MarginSheet.Sheets("Finance Margins").Range("C1").Value
            If PriceFilePath <> "No Price File" Then
                Set PriceFile = Workbooks.Open(PriceFilePath, ReadOnly:=True)
                MarginSheet.Sheets("Finance Margins").Range("A11:A" & Lastrowfinancemargins).Formula = "=IFERROR(IF(VLOOKUP(C11,'[" & PriceFile.Name & "]Price File'!$A:$A,1,FALSE)=C11,""Price File"",""Non-Price File""),""Non-Price File"")"
                MarginSheet.Sheets("Finance Margins").Range("A11:A" & Lastrowfinancemargins).Copy
                MarginSheet.Sheets("Finance Margins").Range("A11:A" & Lastrowfinancemargins).PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
                MarginSheet.Sheets("Finance Margins").Range("H11:H" & Lastrowfinancemargins).Formula = "=IFERROR(IF(VLOOKUP(C11,'[" & PriceFile.Name & "]Price File'!$A:$AQ,43,FALSE)=""No Usage"",0,IF(VLOOKUP(C11,'[" & PriceFile.Name & "]Price File'!$A:$AQ,43,FALSE)=""No Data Available"",0,VLOOKUP(C11,'[" & PriceFile.Name & "]Price File'!$A:$AQ,43,FALSE))),0)"
                MarginSheet.Sheets("Finance Margins").Range("H11:H" & Lastrowfinancemargins).Copy
                MarginSheet.Sheets("Finance Margins").Range("H11:H" & Lastrowfinancemargins).PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
                MarginSheet.Sheets("Finance Margins").Range("C1").ClearContents
                PriceFile.Close False
            End If
            ControlPanel.Close False
        End If
        
        MarginSheet.Sheets("PricePoint").Delete
        
        
        AccountID = MarginSheet.Sheets("Finance Margins").Range("B1").Value
        QuoteID = MarginSheet.Sheets("Finance Margins").Range("C5").Value
        MarginSheet.Sheets("Finance Margins").Range("C5").ClearContents
        MarginSheet.Sheets("Finance Margins").Range("C1").ClearContents
        SavePath = "\\WUKRLS00FP001\RLS_Data\Departments\NACET\Price File Maintenance\Estimators\" & Estimator & "\Quotes\" & AccountID & " PP" & QuoteID & " - " & "Margin Sheet.xlsx"
        MarginSheet.SaveAs SavePath
        MarginSheet.Close
    Else
        MsgBox "Raw File does not appear to be PricePoint Download"
    End If
End Sub
