
Sub vld_clear_existing_positions()
    Dim positions_start As Range
    Set positions_start = SectionLim("Positions", "Positions", "start")
    
    ' Column Indices
    col_ix_pos_qt_ramid = SectionCol("Positions", "Positions", "QT RAM ID").Column - 1
    col_ix_pos_qt_qty = SectionCol("Positions", "Positions", "QT Quantity").Column - 1
    col_ix_pos_fm_ticker = SectionCol("Positions", "Positions", "FM Ticker").Column - 1
    col_ix_pos_fm_ramid = SectionCol("Positions", "Positions", "FM RAM ID").Column - 1
    col_ix_pos_fm_qty = SectionCol("Positions", "Positions", "FM Quantity").Column - 1
    
    ' Clean up existing rows
    second_row = positions_start.Offset(1, 0).Row
    With ActiveSheet
        .Rows(second_row & ":" & .Rows.Count).Delete
    End With
    positions_start.Value = "NUGT"
    positions_start.Offset(0, col_ix_pos_qt_ramid) = "NUGT_Q_TEMP"
    positions_start.Offset(0, col_ix_pos_qt_qty) = 100
    positions_start.Offset(0, col_ix_pos_fm_ticker) = "NUGT"
    positions_start.Offset(0, col_ix_pos_fm_ramid) = "NUGT_Q_TEMP"
    positions_start.Offset(0, col_ix_pos_fm_qty) = 100

End Sub

Sub vld_load_quant_tracker_positions()
    ' ##########################################
    ' QUANT TRACKER POSITIONS
    ' ##########################################
    
    Dim positions_end As Range
    Set positions_end = SectionLim("Positions", "Positions", "end")
    
    ' Column Indices
    col_ix_pos_qt_ramid = SectionCol("Positions", "Positions", "QT RAM ID").Column - 1
    col_ix_pos_qt_qty = SectionCol("Positions", "Positions", "QT Quantity").Column - 1
    col_ix_pos_fm_ticker = SectionCol("Positions", "Positions", "FM Ticker").Column - 1
    col_ix_pos_fm_ramid = SectionCol("Positions", "Positions", "FM RAM ID").Column - 1
    col_ix_pos_fm_qty = SectionCol("Positions", "Positions", "FM Quantity").Column - 1

    ' ERN Holding
    Dim data As Range
    Set data = SectionColData("ERN", "Holding", "Ticker")
	If data.Cells(1, 1).Value <> "NUGT" Then
		data.Copy
		positions_end.Offset(1, 0).PasteSpecial xlPasteValues
    
		Set data = SectionColData("ERN", "Holding", "RAM ID")
		data.Copy
		positions_end.Offset(1, col_ix_pos_qt_ramid).PasteSpecial xlPasteValues
		
		Set data = SectionColData("ERN", "Holding", "Position (Shares)")
		data.Copy
		positions_end.Offset(1, col_ix_pos_qt_qty).PasteSpecial xlPasteValues
	End If
    
    ' ERN Taking Off
    Set positions_end = SectionLim("Positions", "Positions", "end")
    Set data = SectionColData("ERN", "Taking Off", "Ticker")
	If data.Cells(1, 1).Value <> "NUGT" Then
		data.Copy
		positions_end.Offset(1, 0).PasteSpecial xlPasteValues
		
		Set data = SectionColData("ERN", "Taking Off", "RAM ID")
		data.Copy
		positions_end.Offset(1, col_ix_pos_qt_ramid).PasteSpecial xlPasteValues
		
		Set data = SectionColData("ERN", "Taking Off", "Position (Shares)")
		data.Copy
		positions_end.Offset(1, col_ix_pos_qt_qty).PasteSpecial xlPasteValues
	End If

    ' PEAD Holding
    Set positions_end = SectionLim("Positions", "Positions", "end")
    Set data = SectionColData("PEAD", "Holding", "Ticker")
	If data.Cells(1, 1).Value <> "NUGT" Then
		data.Copy
		positions_end.Offset(1, 0).PasteSpecial xlPasteValues
		
		Set data = SectionColData("PEAD", "Holding", "RAM ID")
		data.Copy
		positions_end.Offset(1, col_ix_pos_qt_ramid).PasteSpecial xlPasteValues
		
		Set data = SectionColData("PEAD", "Holding", "Position (Shares)")
		data.Copy
		positions_end.Offset(1, col_ix_pos_qt_qty).PasteSpecial xlPasteValues
	End If
    
    ' PEAD Taking Off
    Set positions_end = SectionLim("Positions", "Positions", "end")
    Set data = SectionColData("PEAD", "Taking Off", "Ticker")
	If data.Cells(1, 1).Value <> "NUGT" Then
		data.Copy
		positions_end.Offset(1, 0).PasteSpecial xlPasteValues
		
		Set data = SectionColData("PEAD", "Taking Off", "RAM ID")
		data.Copy
		positions_end.Offset(1, col_ix_pos_qt_ramid).PasteSpecial xlPasteValues
		
		Set data = SectionColData("PEAD", "Taking Off", "Position (Shares)")
		data.Copy
		positions_end.Offset(1, col_ix_pos_qt_qty).PasteSpecial xlPasteValues
	End If
    
End Sub

Sub vld_load_fund_manager_positions()
    ' ##########################################
    ' FUND MANAGER EXPORT POSITIONS
    ' ##########################################
    
    Dim fm_table_start As Range
    Dim fm_table_end As Range
    Dim fm_investments As Range
    Dim pos_table_start As Range
    Dim pos_table_next_row As Range
    
    ' Get data from Fund Manager Worksheet
    Sheets("FundManager").Visible = True
    Sheets("FundManager").Activate
    find_val = "Investment"
    Set fm_table_start = Cells.Find(What:=find_val, LookIn:=xlValues, LookAt:=xlWhole, _
                            SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False).Offset(1, 0)
    Set fm_table_end = fm_table_start.End(xlDown)
    Set fm_investments = Range(fm_table_start.Address, fm_table_end.Address)
    
    find_val = "Symbol"
    col_ix_tkr = Cells.Find(What:=find_val, LookIn:=xlValues, LookAt:=xlWhole, _
                            SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False).Column - 1
    find_val = "shares"
    col_ix_qty = Cells.Find(What:=find_val, LookIn:=xlValues, LookAt:=xlWhole, _
                            SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False).Column - 1
    Sheets("FundManager").Visible = False
    
    ' Change to Positions Sheet and populate data for ERN/PEAD non-zero shares
    Sheets("Positions").Activate
    col_ix_pos_fm_ticker = SectionCol("Positions", "Positions", "FM Ticker").Column - 1
    col_ix_pos_fm_ramid = SectionCol("Positions", "Positions", "FM RAM ID").Column - 1
    col_ix_pos_fm_qty = SectionCol("Positions", "Positions", "FM Quantity").Column - 1
    Set pos_table_start = SectionLim("Positions", "Positions", "start").Offset(0, col_ix_pos_fm_ticker)
    
    For Each investment In fm_investments
        
        inv_name = investment.Value
        inv_tkr = investment.Offset(0, col_ix_tkr)
        inv_shares = investment.Offset(0, col_ix_qty)
        
        If inv_shares = 0 Then
            GoTo NextInvestment
        End If
        
        pos_ern = InStr(inv_name, "_Q_ERN")
        pos_pead = InStr(inv_name, "_Q_PEAD")
        
        If pos_ern + pos_pead > 0 And inv_tkr <> "SPY" Then
            'HACK to get last row from below hard coding 1000 rows down
            Set pos_table_next_row = pos_table_start.Offset(1000, 0).End(xlUp)
            pos_table_next_row.Offset(1, 0).Value = inv_tkr
            pos_table_next_row.Offset(1, col_ix_pos_fm_ramid - col_ix_pos_fm_ticker).Value = inv_name
            pos_table_next_row.Offset(1, col_ix_pos_fm_qty - col_ix_pos_fm_ticker).Value = inv_shares
        End If
        
NextInvestment:
    Next investment
    
End Sub

Sub vld_fill_formulas()
    ' Copy formulas down
    Dim positions_start As Range
    Dim positions_end As Range
    Set positions_start = SectionLim("Positions", "Positions", "start")
    Set positions_end = SectionLim("Positions", "Positions", "end")

    col_ix_pos_chk_id = SectionCol("Positions", "Positions", "Check QT ID").Column - 1
    col_ix_pos_chk_qty = SectionCol("Positions", "Positions", "Check FM Qty").Column - 1
    
    formula_start = positions_start.Offset(0, col_ix_pos_chk_id).Address
    formula_end = positions_start.Offset(0, col_ix_pos_chk_qty).Address
    Range(formula_start, formula_end).Select

    fill_end = positions_end.Offset(0, col_ix_pos_chk_qty).Address
    Selection.AutoFill Destination:=Range(formula_start, fill_end), Type:=xlFillDefault
    
    ' Delete dummy row
    ActiveSheet.Rows(positions_start.Row).Delete
    
End Sub

Sub vld_sort_positions()

    Dim qt_positions As Range
    Dim fm_positions As Range

    ' Column Indices
    col_ix_pos_qt_ramid = SectionCol("Positions", "Positions", "QT RAM ID").Column - 1
    col_ix_pos_qt_qty = SectionCol("Positions", "Positions", "QT Quantity").Column - 1
    col_ix_pos_fm_ticker = SectionCol("Positions", "Positions", "FM Ticker").Column - 1
    col_ix_pos_fm_ramid = SectionCol("Positions", "Positions", "FM RAM ID").Column - 1
    col_ix_pos_fm_qty = SectionCol("Positions", "Positions", "FM Quantity").Column - 1
    
    ' Table Addresses and Sorting
    qt_table_start = SectionLim("Positions", "Positions", "start").Address
    qt_table_end = SectionLim("Positions", "Positions", "end").Offset(0, col_ix_pos_qt_qty).Address
    qt_table_sort_start = SectionLim("Positions", "Positions", "start").Offset(0, col_ix_pos_qt_ramid).Address
    qt_table_sort_end = SectionLim("Positions", "Positions", "end").Offset(0, col_ix_pos_qt_ramid).Address
    
    fm_table_start = SectionLim("Positions", "Positions", "start").Offset(0, col_ix_pos_fm_ticker).Address
    fm_table_end = SectionLim("Positions", "Positions", "end").Offset(0, col_ix_pos_fm_qty).Address
    fm_table_sort_start = SectionLim("Positions", "Positions", "start").Offset(0, col_ix_pos_fm_ramid).Address
    fm_table_sort_end = SectionLim("Positions", "Positions", "end").Offset(0, col_ix_pos_fm_ramid).Address
    
    Set qt_positions = Range(qt_table_start, qt_table_end)
    Set fm_positions = Range(fm_table_start, fm_table_end)

    qt_positions.Sort key1:=Range(qt_table_sort_start, qt_table_sort_end), order1:=xlAscending
    fm_positions.Sort key1:=Range(fm_table_sort_start, fm_table_sort_end), order1:=xlAscending
    
End Sub



Sub vld_validate_fund_manager()
    ' Check Sheet
    VerifySheet ("Positions")

    ' Clear Existing Rows
    vld_clear_existing_positions
    
    ' Load Quant Tracker Holding/Taking Offset
    vld_load_quant_tracker_positions
    
    ' Load Fund Manager Positions
    vld_load_fund_manager_positions
    
    ' Cleanup and Fill Formulas
    vld_fill_formulas
    
    ' Sort two groups of trades and fill formulas
    vld_sort_positions

End Sub


