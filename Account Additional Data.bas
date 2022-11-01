Sub AccountAdditionalData()
Dim ControlPanel As ThisWorkbook
Dim Location As String
Dim ActiveAccount As Workbook
Set ControlPanel = ThisWorkbook
ControlPanel.Sheets("Control Panel").Range("B2").Select
    Do Until IsEmpty(ActiveCell)
        Location = Application.WorksheetFunction.Index(Sheets("Paths").Range("B2:B500"), Application.WorksheetFunction.Match(ActiveCell.Value, Sheets("Paths").Range("A2:A500"), 0), 1).Value & "\" & Application.WorksheetFunction.Index(Sheets("Paths").Range("C2:C500"), Application.WorksheetFunction.Match(ActiveCell.Value, Sheets("Paths").Range("A2:A500"), 0), 1).Value & ".xlsm"
        Set ActiveAccount = Workbooks.Open(Location)
        If ActiveAccount.Sheets("Price File").AutoFilterMode Then
            ActiveAccount.Sheets("Price File").AutoFilterMode = False
        End If
        LastRowActiveAccount = ActiveAccount.Sheets("Price File").Range("A" & Rows.Count).End(xlUp).Row
        ActiveAccount.Sheets("Price File").Range("A11:BV" & LastRowActiveAccount).AutoFilter Field:=59, Criteria1:="EOL"
        If ActiveAccount.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            OBSDATA = ActiveAccount.Sheets("Price File").Range("BG12:BG" & LastRowActiveAccount).SpecialCells(xlCellTypeVisible).Count
            ControlPanel.Activate
            OBSLocation = Application.WorksheetFunction.Index(Sheets("Control Panel").Range("B2:B500"), Application.WorksheetFunction.Match(ActiveCell.Value, Sheets("Control Panel").Range("B2:B500"), 0), 1).Offset(0, 4).Address
            ControlPanel.Sheets("Control Panel").Range(OBSLocation).Value = OBSDATA
        Else
            ControlPanel.Activate
            OBSLocation = Application.WorksheetFunction.Index(Sheets("Control Panel").Range("B2:B500"), Application.WorksheetFunction.Match(ActiveCell.Value, Sheets("Control Panel").Range("B2:B500"), 0), 1).Offset(0, 4).Address
            ControlPanel.Sheets("Control Panel").Range(OBSLocation).Value = "0"
        End If
        If ActiveAccount.Sheets("Price File").AutoFilterMode Then
            ActiveAccount.Sheets("Price File").AutoFilterMode = False
        End If
        ActiveAccount.Sheets("Price File").Range("A11:BV" & LastRowActiveAccount).AutoFilter Field:=52, Criteria1:="<>"
        ActiveAccount.Sheets("Price File").Range("A11:BV" & LastRowActiveAccount).AutoFilter Field:=53, Criteria1:="="
        If ActiveAccount.Sheets("Price File").AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            SUPPORTDATA = ActiveAccount.Sheets("Price File").Range("A12:A" & LastRowActiveAccount).SpecialCells(xlCellTypeVisible).Count
            ControlPanel.Activate
            OBSLocation = Application.WorksheetFunction.Index(Sheets("Control Panel").Range("B2:B500"), Application.WorksheetFunction.Match(ActiveCell.Value, Sheets("Control Panel").Range("B2:B500"), 0), 1).Offset(0, 3).Address
            ControlPanel.Sheets("Control Panel").Range(OBSLocation).Value = SUPPORTDATA
        Else
            ControlPanel.Activate
            OBSLocation = Application.WorksheetFunction.Index(Sheets("Control Panel").Range("B2:B500"), Application.WorksheetFunction.Match(ActiveCell.Value, Sheets("Control Panel").Range("B2:B500"), 0), 1).Offset(0, 3).Address
            ControlPanel.Sheets("Control Panel").Range(OBSLocation).Value = "0"
        End If
        ActiveAccount.Close False
        ActiveCell.Offset(1, 0).Select
    Loop
End Sub
