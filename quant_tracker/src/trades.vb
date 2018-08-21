
Sub trd_GetDailyTrades()
'
' Move day trades to verify sheet
'

'   Check Sheet
    VerifySheet ("Trades")

    Dim TradeBaseStart As Range
    Dim TradeBaseEnd As Range
    Set TradeBaseStart = SectionLim("Trades", "Trades Base", "start")

    trd_CleanupBaseTrades

    ' Get ERN trades on
    newTkrOn = SectionLim("ERN", "Putting On", "start").Value
    If newTkrOn <> "NUGT" Then
        trd_GetErnTradesOn
    End If

    ' Get ERN trades off
    newTkrOff = SectionLim("ERN", "Taking Off", "start").Value
    If newTkrOff <> "NUGT" Then
        trd_GetErnTradesOff
    End If

    ' Get PEAD trades on
    newTkrOn = SectionLim("PEAD", "Putting On", "start").Value
    If newTkrOn <> "NUGT" Then
        trd_GetPeadTradesOn
    End If

    ' Get PEAD trades off
    newTkrOff = SectionLim("PEAD", "Taking Off", "start").Value
    If newTkrOff <> "NUGT" Then
        trd_GetPeadTradesOff
    End If

    ' Check if there are any new trades
    If IsEmpty(SectionLim("Trades", "Trades Base", "start").Offset(1, 0).Value) Then
		' Clean-up
		trd_CleanupNetTrades
		MsgBox "No New Trades Found"
		End
    Else
        ActiveSheet.Rows(TradeBaseStart.Row).Delete
		' Adjust quantity sign for trades taking off
		trd_adjust_quantity_for_side
        ' Sort tickers alphabetically and get Net Trades
        Set TradeBaseStart = SectionLim("Trades", "Trades Base", "start")
        Set TradeBaseEnd = SectionLim("Trades", "Trades Base", "end")
        baseRAMIDColIX = SectionCol("Trades", "Trades Base", "RAM ID").Column - 1
        Range(TradeBaseStart, TradeBaseEnd.Offset(0, baseRAMIDColIX)).Sort key1:=Range(TradeBaseStart, TradeBaseEnd), order1:=xlAscending
        trd_GetNetTrades
    End If
End Sub


Sub trd_adjust_quantity_for_side()
	' Trades Base Range
	Dim trades_base
	Set trades_base = SectionColData("Trades", "Trades Base", "Ticker")

    ' Column Indices
    col_ix_tb_qty = SectionCol("Trades", "Trades Base", "Quantity").Column - 1
    col_ix_tb_side = SectionCol("Trades", "Trades Base", "Side").Column - 1

    For Each trade In trades_base
		' Set order rows and check if has been processed
        If trade.Offset(0, col_ix_tb_side) = "OFF" Then
			trade.Offset(0, col_ix_tb_qty) = trade.Offset(0, col_ix_tb_qty) * -1
		End If
	Next trade
End Sub


Sub trd_CleanupBaseTrades()

    Dim TradeBaseStart As Range
    Set TradeBaseStart = SectionLim("Trades", "Trades Base", "start")

    remRow = TradeBaseStart.Offset(1, 0).Row
    With ActiveSheet
        .Rows(remRow & ":" & .Rows.Count).Delete
    End With

    TradeBaseStart.Value = "NUGT"
    QtyColIX = SectionCol("Trades", "Trades Base", "Quantity").Column - 1
    TradeBaseStart.Offset(0, QtyColIX) = 0

End Sub


Sub trd_GetErnTradesOn()
    Dim TradeBaseStart As Range
    Set TradeBaseStart = SectionLim("Trades", "Trades Base", "start")

'   Get Ern Tickers
    Dim data As Range
    Set data = SectionColData("ERN", "Putting On", "Ticker")
    data.Copy
    TradeBaseStart.Offset(1, 0).PasteSpecial xlPasteValues

'   Get Ern Quantities
    Set data = SectionColData("ERN", "Putting On", "RT Order")
    data.Copy
    QtyColIX = SectionCol("Trades", "Trades Base", "Quantity").Column - 1
    TradeBaseStart.Offset(1, QtyColIX).PasteSpecial xlPasteValues

'   Get Ern RAMIDs
    Set data = SectionColData("ERN", "Putting On", "RAM ID")
    data.Copy
    ram_id_colIX = SectionCol("Trades", "Trades Base", "RAM ID").Column - 1
    TradeBaseStart.Offset(1, ram_id_colIX).PasteSpecial xlPasteValues

'   Fill in Trade and Side Info
    Dim TradeBaseEnd As Range
    Set TradeBaseEnd = SectionLim("Trades", "Trades Base", "end")

    colTradeIX = SectionCol("Trades", "Trades Base", "Trade").Column - 1
    colSideIX = SectionCol("Trades", "Trades Base", "Side").Column - 1

    trdS = TradeBaseStart.Offset(1, colTradeIX).Address
    trde = TradeBaseEnd.Offset(0, colTradeIX).Address
    Range(trdS, trde).Value = "ERN"

    trdS = TradeBaseStart.Offset(1, colSideIX).Address
    trde = TradeBaseEnd.Offset(0, colSideIX).Address
    Range(trdS, trde).Value = "ON"

End Sub

Sub trd_GetErnTradesOff()
    Dim TradeBaseEndI As Range
    Set TradeBaseEndI = SectionLim("Trades", "Trades Base", "end")

'   Get Ern Tickers
    Dim data As Range
    Set data = SectionColData("ERN", "Taking Off", "Ticker")
    data.Copy
    TradeBaseEndI.Offset(1, 0).PasteSpecial xlPasteValues

'   Get Ern Quantities
    Set data = SectionColData("ERN", "Taking Off", "Position (Shares)")
    data.Copy
    QtyColIX = SectionCol("Trades", "Trades Base", "Quantity").Column - 1
    TradeBaseEndI.Offset(1, QtyColIX).PasteSpecial xlPasteValues

'   Get Ern RAMIDs
    Set data = SectionColData("ERN", "Taking Off", "RAM ID")
    data.Copy
    ram_id_colIX = SectionCol("Trades", "Trades Base", "RAM ID").Column - 1
    TradeBaseEndI.Offset(1, ram_id_colIX).PasteSpecial xlPasteValues

'   Fill in Trade and Side Info
    Dim TradeBaseEnd As Range
    Set TradeBaseEnd = SectionLim("Trades", "Trades Base", "end")

    colTradeIX = SectionCol("Trades", "Trades Base", "Trade").Column - 1
    colSideIX = SectionCol("Trades", "Trades Base", "Side").Column - 1

    trdS = TradeBaseEndI.Offset(1, colTradeIX).Address
    trde = TradeBaseEnd.Offset(0, colTradeIX).Address
    Range(trdS, trde).Value = "ERN"

    trdS = TradeBaseEndI.Offset(1, colSideIX).Address
    trde = TradeBaseEnd.Offset(0, colSideIX).Address
    Range(trdS, trde).Value = "OFF"

End Sub

Sub trd_GetPeadTradesOn()
    Dim TradeBaseEndI As Range
    Set TradeBaseEndI = SectionLim("Trades", "Trades Base", "end")

'   Get Pead Tickers
    Dim data As Range
    Set data = SectionColData("PEAD", "Putting On", "Ticker")
    data.Copy
    TradeBaseEndI.Offset(1, 0).PasteSpecial xlPasteValues

'   Get Pead Quantities
    Set data = SectionColData("PEAD", "Putting On", "RT Order")
    data.Copy
    QtyColIX = SectionCol("Trades", "Trades Base", "Quantity").Column - 1
    TradeBaseEndI.Offset(1, QtyColIX).PasteSpecial xlPasteValues

'   Get Pead RAMIDs
    Set data = SectionColData("PEAD", "Putting On", "RAM ID")
    data.Copy
    ram_id_colIX = SectionCol("Trades", "Trades Base", "RAM ID").Column - 1
    TradeBaseEndI.Offset(1, ram_id_colIX).PasteSpecial xlPasteValues

'   Fill in Trade and Side Info
    Dim TradeBaseEnd As Range
    Set TradeBaseEnd = SectionLim("Trades", "Trades Base", "end")

    colTradeIX = SectionCol("Trades", "Trades Base", "Trade").Column - 1
    colSideIX = SectionCol("Trades", "Trades Base", "Side").Column - 1

    trdS = TradeBaseEndI.Offset(1, colTradeIX).Address
    trde = TradeBaseEnd.Offset(0, colTradeIX).Address
    Range(trdS, trde).Value = "PEAD"

    trdS = TradeBaseEndI.Offset(1, colSideIX).Address
    trde = TradeBaseEnd.Offset(0, colSideIX).Address
    Range(trdS, trde).Value = "ON"
End Sub

Sub trd_GetPeadTradesOff()
    Dim TradeBaseEndI As Range
    Set TradeBaseEndI = SectionLim("Trades", "Trades Base", "end")

'   Get Pead Tickers
    Dim data As Range
    Set data = SectionColData("PEAD", "Taking Off", "Ticker")
    data.Copy
    TradeBaseEndI.Offset(1, 0).PasteSpecial xlPasteValues

'   Get Pead Quantities
    Set data = SectionColData("PEAD", "Taking Off", "Position (Shares)")
    data.Copy
    QtyColIX = SectionCol("Trades", "Trades Base", "Quantity").Column - 1
    TradeBaseEndI.Offset(1, QtyColIX).PasteSpecial xlPasteValues

'   Get Pead RAMIDs
    Set data = SectionColData("PEAD", "Taking Off", "RAM ID")
    data.Copy
    ram_id_colIX = SectionCol("Trades", "Trades Base", "RAM ID").Column - 1
    TradeBaseEndI.Offset(1, ram_id_colIX).PasteSpecial xlPasteValues

'   Fill in Trade and Side Info
    Dim TradeBaseEnd As Range
    Set TradeBaseEnd = SectionLim("Trades", "Trades Base", "end")

    colTradeIX = SectionCol("Trades", "Trades Base", "Trade").Column - 1
    colSideIX = SectionCol("Trades", "Trades Base", "Side").Column - 1

    trdS = TradeBaseEndI.Offset(1, colTradeIX).Address
    trde = TradeBaseEnd.Offset(0, colTradeIX).Address
    Range(trdS, trde).Value = "PEAD"

    trdS = TradeBaseEndI.Offset(1, colSideIX).Address
    trde = TradeBaseEnd.Offset(0, colSideIX).Address
    Range(trdS, trde).Value = "OFF"

End Sub


Sub trd_GetNetTrades()
'
'   Roll up all trades to a net position
'
    Dim TradeRange As Range
    Dim SearchRange As Range
    Dim FindRes As Range
    Dim CurrentEnd As Range

'   Cleanup
    trd_CleanupNetTrades

    Set TradeRange = SectionColData("Trades", "Trades Base", "Ticker")
    baseQtyColIX = SectionCol("Trades", "Trades Base", "Quantity").Column - 1
    baseSideColIX = SectionCol("Trades", "Trades Base", "Side").Column - 1
    netQtyColIX = SectionCol("Trades", "Trades Base", "Quantity").Column - 1

'   Iterate through the individual orders and create a net order
    For Each Cell In TradeRange

        trd_side = Cell.Offset(0, baseSideColIX).Value
        trd_qty = Cell.Offset(0, baseQtyColIX).Value

        Set SearchRange = SectionColData("Trades", "Trades Net", "Ticker")
        Set FindRes = SearchRange.Find(What:=Cell.Value, LookIn:=xlValues, LookAt:=xlWhole _
                        , SearchDirection:=xlNext, MatchCase:=True _
                        , SearchFormat:=False)

        If FindRes Is Nothing Then
            Set CurrentEnd = SectionLim("Trades", "Trades Net", "end")
            CurrentEnd.Offset(1, 0).EntireRow.Resize(1).Insert
            CurrentEnd.Offset(1, 0).Value = Cell.Value
            CurrentEnd.Offset(1, netQtyColIX).Value = trd_qty
        Else
            FindRes.Offset(0, netQtyColIX).Value = FindRes.Offset(0, netQtyColIX).Value + trd_qty
        End If

    Next Cell

' Copy formulas down
    Dim StartCell As Range
    Set StartCell = SectionLim("Trades", "Trades Net", "start")
    colKDPositionIX = SectionCol("Trades", "Trades Net", "RT Shares").Column - 1
    colFMPositionIX = SectionCol("Trades", "Trades Net", "FM Shares").Column - 1

    formulaS = StartCell.Offset(0, colKDPositionIX).Address
    formulaE = StartCell.Offset(0, colFMPositionIX).Address
    Range(formulaS, formulaE).Select

    Set CurrentEnd = SectionLim("Trades", "Trades Net", "end")
    fillE = CurrentEnd.Offset(0, colFMPositionIX).Address
    Selection.AutoFill Destination:=Range(formulaS, fillE), Type:=xlFillDefault

'   Delete First Row
    remRow = StartCell.Row
    ActiveSheet.Rows(remRow).Delete

End Sub

Sub trd_CleanupNetTrades()

'   Remove Rows from Trades Net Section except first
    Dim NetStart As Range
    Set NetStart = SectionLim("Trades", "Trades Net", "start")
    colNetQtyIX = SectionCol("Trades", "Trades Net", "Quantity").Column - 1

    Dim NetEnd As Range
    Set NetEnd = SectionLim("Trades", "Trades Net", "end")

    If Not IsEmpty(NetStart.Offset(1, 0).Value) Then
        With ActiveSheet
        .Rows(NetStart.Row + 1 & ":" & NetEnd.Row).Delete
        End With
    End If

    NetStart.Value = "NUGT"
    NetStart.Offset(0, colNetQtyIX).Value = 0

End Sub

Sub trd_GenerateRTBasket()
    '   Check Sheet
    VerifySheet ("Trades")
    '   Cleanup
    trd_CleanupFinalOrders

    ' Fill in Final Orders section
    trd_GetFinalOrders

    If SectionLim("Trades", "Final Orders", "start").Value = "NUGT" Then
        MsgBox "No new orders found"
        End
    End If

    'Make sheet Visible to work with
    Sheets("RT Basket").Visible = True
    Dim orderData As Range
    ' Copy orders into correct places in upload sheet
    With Sheets("RT Basket")
        .Rows(3 & ":" & .Rows.Count).Delete

        Set orderData = SectionColData("Trades", "Final Orders", "Ticker")
        orderData.Copy
        .Range("A3").PasteSpecial xlPasteValues

        Set orderData = SectionColData("Trades", "Final Orders", "Side")
        orderData.Copy
        .Range("B3").PasteSpecial xlPasteValues

        Set orderData = SectionColData("Trades", "Final Orders", "Quantity")
        orderData.Copy
        .Range("C3").PasteSpecial xlPasteValues

        Dim lastRow As Long
        lastRow = .Range("A" & .Rows.Count).End(xlUp).Row

        .Range("D2:D2").Copy
        .Range("D3:D" & lastRow).PasteSpecial xlPasteValues

        .Range("F2:F2").Copy
        .Range("F3:F" & lastRow).PasteSpecial xlPasteValues

        .Range("G2:G2").Select
        Selection.AutoFill Destination:=Range("G2:G" & lastRow), Type:=xlFillDefault

        .Range("H2:H2").Copy
        .Range("H3:H" & lastRow).PasteSpecial xlPasteValues

        .Rows(2).Delete

    End With

    ' Save to CSV
    res = SaveCSV("RT Basket", "C:\temp\Quant RT Basket.csv")
    Dim sDate As String
    sDate = Format(Now(), "yyyymmdd")
    outpath = Environ("DATA") & "\ramex\daily_diffs\" & sDate & "_ern_pead.csv"
    res = SaveCSV("RT Basket", outpath)

    ' Hide sheet again
    Sheets("RT Basket").Visible = False
    Sheets("Trades").Activate

End Sub

Sub trd_CleanupFinalOrders()

'   Remove Rows from Final Orders Section except first
    Dim FinalStart As Range
    Set FinalStart = SectionLim("Trades", "Final Orders", "start")
    colFinalQtyIX = SectionCol("Trades", "Final Orders", "Quantity").Column - 1

    Dim FinalEnd As Range
    Set FinalEnd = SectionLim("Trades", "Final Orders", "end")

    If Not IsEmpty(FinalStart.Offset(1, 0).Value) Then
        With ActiveSheet
        .Rows(FinalStart.Row + 1 & ":" & FinalEnd.Row).Delete
        End With
    End If

    FinalStart.Value = "NUGT"
    FinalStart.Offset(0, colFinalQtyIX).Value = 0

End Sub


Sub trd_GetFinalOrders()
'
'   Create final table of orders that can be moved to a KD upload sheet
'

    Dim NetRange As Range
    Dim CurrentEnd As Range

    Set NetRange = SectionColData("Trades", "Trades Net", "Ticker")
    colOrderQtyIX = SectionCol("Trades", "Final Orders", "Quantity").Column - 1
    colOrderSideIX = SectionCol("Trades", "Final Orders", "Side").Column - 1
    colNetQtyIX = SectionCol("Trades", "Trades Net", "Quantity").Column - 1

'   Use Opening Positions from KD or FM
    answer = MsgBox("Overide RT Shares with FM Shares for Opening Positions?", vbYesNo + vbQuestion, "Source for Opening Positions")
    Dim OpeningShares As Range

    If answer = vbYes Then
        Set OpeningShares = SectionColData("Trades", "Trades Net", "FM Shares")
    Else
        Set OpeningShares = SectionColData("Trades", "Trades Net", "RT Shares")
    End If
    OpeningShares.Copy
    colOpeningPositionIX = SectionCol("Trades", "Trades Net", "Opening Position").Column - 1
    SectionLim("Trades", "Trades Net", "start").Offset(0, colOpeningPositionIX).PasteSpecial xlPasteValues

'   Iterate through the idividual orders and create a net order
    For Each Cell In NetRange
        Set CurrentEnd = SectionLim("Trades", "Final Orders", "end")

        CurrentEnd.Offset(1, 0).EntireRow.Resize(1).Insert
        trd_qty = Cell.Offset(0, colNetQtyIX).Value
        trd_ex_pos = Cell.Offset(0, colOpeningPositionIX).Value
        CurrentEnd.Offset(1, 0).Value = Cell.Value

        ' Simple Orders
        If trd_ex_pos = 0 Then
        ' New position
            CurrentEnd.Offset(1, colOrderQtyIX).Value = Abs(Cell.Offset(0, colNetQtyIX).Value)
            If trd_qty > 0 Then
                CurrentEnd.Offset(1, colOrderSideIX).Value = "Buy"
            ElseIf trd_qty < 0 Then
                CurrentEnd.Offset(1, colOrderSideIX).Value = "SellShort"
            End If

        ElseIf (trd_qty + trd_ex_pos) = 0 Then
        ' Closing Position Out
            CurrentEnd.Offset(1, colOrderQtyIX).Value = Abs(Cell.Offset(0, colNetQtyIX).Value)
            If trd_qty < 0 Then
                CurrentEnd.Offset(1, colOrderSideIX).Value = "Sell"
            ElseIf trd_qty > 0 Then
                CurrentEnd.Offset(1, colOrderSideIX).Value = "BuyToCover"
            End If

        ' Complex Orders
        ElseIf (trd_qty + trd_ex_pos) > 0 And trd_ex_pos > 0 Then
        ' Long and Remaining Long
            CurrentEnd.Offset(1, colOrderQtyIX).Value = Abs(Cell.Offset(0, colNetQtyIX).Value)
            If trd_qty < 0 Then
                CurrentEnd.Offset(1, colOrderSideIX).Value = "Sell"
            ElseIf trd_qty > 0 Then
                CurrentEnd.Offset(1, colOrderSideIX).Value = "Buy"
            End If

        ElseIf (trd_qty + trd_ex_pos) < 0 And trd_ex_pos < 0 Then
        ' Short and Remaining Short
            CurrentEnd.Offset(1, colOrderQtyIX).Value = Abs(Cell.Offset(0, colNetQtyIX).Value)
            If trd_qty < 0 Then
                CurrentEnd.Offset(1, colOrderSideIX).Value = "SellShort"
            ElseIf trd_qty > 0 Then
                CurrentEnd.Offset(1, colOrderSideIX).Value = "BuyToCover"
            End If

        ElseIf (trd_qty + trd_ex_pos) < 0 And trd_ex_pos > 0 Then
        ' Long and Going Short
            CurrentEnd.Offset(2, 0).EntireRow.Resize(1).Insert
            CurrentEnd.Offset(2, 0).Value = Cell.Value
            CurrentEnd.Offset(1, colOrderQtyIX).Value = Abs(trd_ex_pos)
            CurrentEnd.Offset(1, colOrderSideIX).Value = "Sell"
            CurrentEnd.Offset(2, colOrderQtyIX).Value = Abs(trd_qty + trd_ex_pos)
            CurrentEnd.Offset(2, colOrderSideIX).Value = "SellShort"

        ElseIf (trd_qty + trd_ex_pos) > 0 And trd_ex_pos < 0 Then
        ' Short and Going Long
            CurrentEnd.Offset(2, 0).EntireRow.Resize(1).Insert
            CurrentEnd.Offset(2, 0).Value = Cell.Value
            CurrentEnd.Offset(1, colOrderQtyIX).Value = Abs(trd_ex_pos)
            CurrentEnd.Offset(1, colOrderSideIX).Value = "BuyToCover"
            CurrentEnd.Offset(2, colOrderQtyIX).Value = Abs(trd_qty + trd_ex_pos)
            CurrentEnd.Offset(2, colOrderSideIX).Value = "Buy"

        End If

    Next Cell

    'Copy Verification formulas down
    Dim OrderStart As Range
    Set OrderStart = SectionLim("Trades", "Final Orders", "start")
    colTrdsChkQtyIX = SectionCol("Trades", "Final Orders", "Check Quantity").Column - 1
    colTrdsChkSideIX = SectionCol("Trades", "Final Orders", "Check Side").Column - 1

    formulaS = OrderStart.Offset(0, colTrdsChkQtyIX).Address
    formulaE = OrderStart.Offset(0, colTrdsChkSideIX).Address
    Range(formulaS, formulaE).Select

    Dim OrderEnd As Range
    Set OrderEnd = SectionLim("Trades", "Final Orders", "end")
    fillE = OrderEnd.Offset(0, colTrdsChkSideIX).Address
    Selection.AutoFill Destination:=Range(formulaS, fillE), Type:=xlFillDefault

    ActiveSheet.Rows(SectionLim("Trades", "Final Orders", "start").Row).Delete

End Sub

Sub trd_VerifyTrades()

'   Check Sheet
    VerifySheet ("Trades")

'   Proceed
    answer = MsgBox("Have latest Trades from RT been pasted into RT Blotter?", _
                    vbYesNo + vbQuestion, "Trades from RT Order Blotter")

    If answer = vbYes Then
        Dim OrderRange As Range
        Set OrderRange = SectionColData("Trades", "Final Orders", "Ticker")
        colTrdsQtyIX = SectionCol("Trades", "Final Orders", "RT Quantity").Column - 1
        colTrdsSideIX = SectionCol("Trades", "Final Orders", "RT Side").Column - 1

        ' Check columns are in RT    Blotter and get indices
        Sheets("RT Blotter").Activate
        Dim BlotterCols As Range
        Set BlotterCols = Rows(1)

        Dim colA As Range
        Dim colB As Range
        Dim colC As Range
        Set colA = BlotterCols.Find(What:="Symbol", SearchFormat:=False)
        Set colB = BlotterCols.Find(What:="Side", SearchFormat:=False)
        Set colC = BlotterCols.Find(What:="Volume", SearchFormat:=False)

        If colA Is Nothing Or colB Is Nothing Or colC Is Nothing Then
            MsgBox "Required columns (Symbol, Side, Qty) not found in KD Blotter Sheet"
            End
        Else
            colBltrSymbIX = BlotterCols.Find(What:="Symbol", SearchFormat:=False).Column
            colBltrSideIX = BlotterCols.Find(What:="Side", SearchFormat:=False).Column
            colBltrQtyIX = BlotterCols.Find(What:="Volume", SearchFormat:=False).Column

            Dim SearchRange As Range
            lastRow = Range("A" & Rows.Count).End(xlUp).Row
            Set SearchRange = Range(Cells(2, colBltrSymbIX), Cells(lastRow, colBltrSymbIX))
        End If

        ' Pull Data from KD Blotter for each Symbol
        For Each Cell In OrderRange
            Dim FindRes As Range
            tkr = Cell.Value

            Set FindRes = SearchRange.Find(What:=tkr, LookIn:=xlValues, LookAt:=xlWhole _
                                        , MatchCase:=False, SearchFormat:=False)

            If Not FindRes Is Nothing Then
                Cell.Offset(0, colTrdsQtyIX).Value = Cells(FindRes.Row, colBltrQtyIX).Value
                Cell.Offset(0, colTrdsSideIX).Value = Cells(FindRes.Row, colBltrSideIX).Value
            End If

        Next Cell

        Sheets("Trades").Activate

    End If
End Sub

Sub trd_build_fm_upload()
    ' Check Sheet
    VerifySheet ("Trades")

	'If there are no rows go to end for SPY
	If SectionLim("Trades", "Trades Base", "start").Value = "NUGT" Then
        GoTo AddSPY
    End If

    Dim final_orders As Range
    Dim net_orders As Range
    Dim base_orders As Range

    Set final_orders = SectionColData("Trades", "Final Orders", "Ticker")
    Set net_orders = SectionColData("Trades", "Trades Net", "Ticker")
    Set base_orders = SectionColData("Trades", "Trades Base", "Ticker")

    ' Column Indices
    col_ix_tb_qty = SectionCol("Trades", "Trades Base", "Quantity").Column - 1
    col_ix_tb_trade = SectionCol("Trades", "Trades Base", "Trade").Column - 1
    col_ix_tb_side = SectionCol("Trades", "Trades Base", "Side").Column - 1
    col_ix_tb_ram_id = SectionCol("Trades", "Trades Base", "RAM ID").Column - 1
    col_ix_tn_open_qty = SectionCol("Trades", "Trades Net", "Opening Position").Column - 1
    col_ix_fo_ord_type = SectionCol("Trades", "Final Orders", "Side").Column - 1
    col_ix_fo_qty = SectionCol("Trades", "Final Orders", "Quantity").Column - 1

	' Make sheet Visible to work with and cleanup any existing records
    Sheets("FM Upload").Visible = True
    Sheets("FM Upload").Rows(2 & ":" & Sheets("FM Upload").Rows.Count).Delete

    Dim trade_base_a As Range
    Dim trade_base_b As Range
	Dim final_order_b As Range
	Dim trade_net As Range

    For Each final_order_a In final_orders
		' Set order rows and check if has been processed
		If final_order_a.Value = final_order_a.Offset(-1, 0).Value Then
			GoTo NextOrder
        ElseIf final_order_a.Value = final_order_a.Offset(1, 0).Value Then
			Set final_order_b = final_order_a.Offset(1, 0)
			order_b_ordtype = final_order_b.Offset(0, col_ix_fo_ord_type)
			order_b_qty = final_order_b.Offset(0, col_ix_fo_qty)
		Else
			Set final_order_b = Nothing
		End If

		' Get base trade rows
		Set trade_base_a = base_orders.Find(What:=final_order_a.Value, LookIn:=xlValues, LookAt:=xlWhole _
				, SearchDirection:=xlPrevious, MatchCase:=True, SearchFormat:=False)
		If trade_base_a.Value = trade_base_a.Offset(-1, 0).Value Then
			Set trade_base_b = trade_base_a
			Set trade_base_a = trade_base_b.Offset(-1, 0)
		Else
			Set trade_base_b = Nothing
		End If

		' Set variables with order A/trade A details
		trade_a_qty = trade_base_a.Offset(0, col_ix_tb_qty).Value
		trade_a_tradeid = trade_base_a.Offset(0, col_ix_tb_trade).Value
		trade_a_side = trade_base_a.Offset(0, col_ix_tb_side).Value
		trade_a_ramid = trade_base_a.Offset(0, col_ix_tb_ram_id).Value
		order_a_ordtype = final_order_a.Offset(0, col_ix_fo_ord_type)
		order_a_qty = final_order_a.Offset(0, col_ix_fo_qty)

        ' Trading in one or the other ERN/PEAD
		If trade_base_b is Nothing Then
			out_phantom = 0
			With Sheets("FM Upload")
				.Rows("2:3").Resize.Insert
				.Range("A2:A3").Value = final_order_a.Value
				.Range("B2:B3").Value = out_phantom
				.Range("C2:C3").Value = trade_a_tradeid
				.Range("D2:D3").Value = trade_a_side
				.Range("E2").Value = order_a_ordtype
				.Range("F2").Value = order_a_qty
				.Range("G2").Value = order_a_qty
				.Range("H2:H3").Value = trade_a_ramid
				' Handle situation where changing sides
				If Not(final_order_b is Nothing) Then
					.Range("E3").Value = order_b_ordtype
					.Range("F3").Value = order_b_qty
					.Range("G3").Value = order_b_qty
				Else
					.Rows(3).Delete
				End If
			End With
			GoTo NextOrder
		End If

		' Set variables with trade B details
		trade_b_qty = trade_base_b.Offset(0, col_ix_tb_qty).Value
		trade_b_tradeid = trade_base_b.Offset(0, col_ix_tb_trade).Value
		trade_b_side = trade_base_b.Offset(0, col_ix_tb_side).Value
		trade_b_ramid = trade_base_b.Offset(0, col_ix_tb_ram_id).Value

		' Get Opening/Closing Position
		Set trade_net = net_orders.Find(What:=final_order_a.Value, LookIn:=xlValues, LookAt:=xlWhole _
				, SearchDirection:=xlPrevious, MatchCase:=True, SearchFormat:=False)
		open_qty = trade_net.Offset(0, col_ix_tn_open_qty).Value
		close_qty = open_qty + trade_a_qty + trade_b_qty

		' No StatArb, Special Sit Position / Staying on same side
		If (open_qty = -trade_a_qty) Or (Sgn(open_qty) = Sgn(close_qty)) Then
			'Same side (no phantom)
			If Sgn(trade_a_qty) = Sgn(trade_b_qty) Then
				out_phantom = 0
				With Sheets("FM Upload")
					.Rows("2:3").Resize.Insert
					.Range("A2:A3").Value = final_order_a.Value
					.Range("B2:B3").Value = out_phantom
					.Range("C2").Value = trade_a_tradeid
					.Range("C3").Value = trade_b_tradeid
					.Range("D2").Value = trade_a_side
					.Range("D3").Value = trade_b_side
					.Range("F2:G2").Value = Abs(trade_a_qty)
					.Range("F3:G3").Value = Abs(trade_b_qty)
					.Range("H2").Value = trade_a_ramid
					.Range("H3").Value = trade_b_ramid
					' If not changing sides (big SP/STArb Position)
					If final_order_b is Nothing Then
						.Range("E2:E3").Value = order_a_ordtype
					' Changing sides with just PEAD/ERN positions
					ElseIf Abs(trade_a_qty) = order_a_qty Then
						.Range("E2").Value = order_a_ordtype
						.Range("E3").Value = order_b_ordtype
					Else
						.Range("E2").Value = order_b_ordtype
						.Range("E3").Value = order_a_ordtype
					End If
				End With
			'Different Sides (phantom)
			Else
				With Sheets("FM Upload")
					.Rows("2:4").Resize.Insert
					.Range("A2:A4").Value = final_order_a.Value
					.Range("B2").Value = 0
					.Range("B3:B4").Value = 1
					If Abs(trade_a_qty) > Abs(trade_b_qty) Then
						ph_qty = Abs(trade_b_qty)
						If trade_b_qty > 0 Then ph_ordtype_b = "Buy"  Else ph_ordtype_b = "SellShort"
						.Range("C2:C3").Value = trade_a_tradeid
						.Range("C4").Value = trade_b_tradeid
						.Range("D2:D3").Value = trade_a_side
						.Range("D4").Value = trade_b_side
						.Range("H2:H3").Value = trade_a_ramid
						.Range("H4").Value = trade_b_ramid
					Else
						ph_qty = Abs(trade_a_qty)
						If trade_a_qty > 0 Then ph_ordtype_b = "BuyToCover"  Else ph_ordtype_b = "Sell"
						.Range("C2:C3").Value = trade_b_tradeid
						.Range("C4").Value = trade_a_tradeid
						.Range("D2:D3").Value = trade_b_side
						.Range("D4").Value = trade_a_side
						.Range("H2:H3").Value = trade_b_ramid
						.Range("H4").Value = trade_a_ramid
					End If
					.Range("E2:E3").Value = order_a_ordtype
					.Range("E4").Value = ph_ordtype_b
					.Range("F2:G2").Value = order_a_qty
					.Range("F3:F4").Value = ph_qty
				End With
			End If
		' ERN/PEAD/SpSit or StatArb positions & Changing Sides
		Else
			ord_max_qty = WorksheetFunction.Max(order_a_qty, order_b_qty)
			ord_min_qty = WorksheetFunction.Min(order_a_qty, order_b_qty)
			trd_max_qty = WorksheetFunction.Max(Abs(trade_a_qty), Abs(trade_b_qty))
			trd_min_qty = WorksheetFunction.Min(Abs(trade_a_qty), Abs(trade_b_qty))

			If trd_max_qty = Abs(trade_a_qty) Then
				trd_max_tradeid = trade_a_tradeid
				trd_min_tradeid = trade_b_tradeid
				trd_max_ramid = trade_a_ramid
				trd_min_ramid = trade_b_ramid
				trd_max_side = trade_a_side
				trd_min_side = trade_b_side
			Else
				trd_max_tradeid = trade_b_tradeid
				trd_min_tradeid = trade_a_tradeid
				trd_max_ramid = trade_b_ramid
				trd_min_ramid = trade_a_ramid
				trd_max_side = trade_b_side
				trd_min_side = trade_a_side
			End If

			If ord_max_qty = order_a_qty Then
				ord_max_order_type = order_a_ordtype
				ord_min_order_type = order_b_ordtype
			Else
				ord_max_order_type = order_b_ordtype
				ord_min_order_type = order_a_ordtype
			End If

			' Same side ERN and PEAD
			If Sgn(trade_a_qty) = Sgn(trade_b_qty) Then
				out_phantom = 0
				With Sheets("FM Upload")
					.Rows("2:4").Resize.Insert
					.Range("A2:A4").Value = final_order_a.Value
					.Range("B2:B4").Value = out_phantom
					.Range("C2:C3").Value = trd_max_tradeid
					.Range("C4").Value = trd_min_tradeid
					.Range("D2:D3").Value = trd_max_side
					.Range("D4").Value = trd_min_side
					.Range("E2").Value = ord_min_order_type
					.Range("E3:E4").Value = ord_max_order_type
					.Range("F2:G2").Value = ord_min_qty
					.Range("F3:G3").Value = ord_max_qty - Abs(trd_min_qty)
					.Range("F4:G4").Value = Abs(trd_min_qty)
					.Range("H2:H3").Value = trd_max_ramid
					.Range("H4").Value = trd_min_ramid
				End With
			' Different sides (Phantom Trade)
			Else
				If order_a_ordtype = "Buy" Or order_a_ordtype = "BuyToCover" Then
					ph_ord_type = "Sell"
				Else
					ph_ord_type = "Buy"
				End If
				With Sheets("FM Upload")
					.Rows("2:5").Resize.Insert
					.Range("A2:A5").Value = final_order_a.Value
					.Range("B2:B3").Value = 0
					.Range("B4:B5").Value = 1
					.Range("C2:C4").Value = trd_max_tradeid
					.Range("C5").Value = trd_min_tradeid
					.Range("D2:D4").Value = trd_max_side
					.Range("D5").Value = trd_min_side
					.Range("E2").Value = ord_min_order_type
					.Range("E3:E4").Value = ord_max_order_type
					.Range("E5").Value = ph_ord_type
					.Range("F2:G2").Value = ord_min_qty
					.Range("F3:G3").Value = ord_max_qty
					.Range("F4:F5").Value = Abs(trd_min_qty)
					.Range("H2:H4").Value = trd_max_ramid
					.Range("H5").Value = trd_min_ramid
				End With
			End If
		End If
	NextOrder:
    Next final_order_a

	AddSPY:
	' Add SPY
	trd_AddSpyToFM

	' Save to CSV
	Dim sDate As String
	sDate = Format(Now(), "yyyymmdd")
	outpath = Environ("DATA") & "\bookkeeper\ern_pead_raw\" & sDate & "_ern_pead_trades.csv"
	res = SaveCSV("FM Upload", outpath)

	' Hide sheet again
	Sheets("FM Upload").Visible = False
	Sheets("Trades").Activate

	End Sub

Sub trd_AddSpyToFM()
    ' Confirm SPY Shares updated on Summary Sheet
    answer = MsgBox("Are SPY Shares Updated on Summary Sheet?", vbYesNo + vbQuestion, "Update SPY Shares")
    If answer = vbNo Then
        Sheets("Summary").Activate
        End
    End If

    ' Get shares to trade from Summary sheet
    Dim spy_quantities As Range
    Set spy_quantities = spyTradeQuantities()
    ern_qty = spy_quantities.Item(1, 1)
    pead_qty = spy_quantities.Item(2, 1)
    net_qty = ern_qty + pead_qty
    col_ix_open_shares = SectionCol("Summary", "SPY", "Open (Shares)").Column - 1
    open_qty = SectionLim("Summary", "SPY", "end").Offset(0, col_ix_open_shares).Value

    'If Not Trading then exit
    If ern_qty = 0 And pead_qty = 0 Then Exit Sub

    ' Set strategy_ids
    ern_strategy_id = "SPY_Q_" & get_strategy_id("ERN")
    pead_strategy_id = "SPY_Q_" & get_strategy_id("PEAD")

    ' Get SPY row in FM Upload
    Sheets("FM Upload").Activate
    lastRow = Sheets("FM Upload").Range("A" & Sheets("FM Upload").Rows.Count).End(xlUp).Row + 1

    ' Set Side
    If net_qty > 0 Then
        If open_qty < 0 Then ord_type_A = "BuyToCover" Else ord_type_A = "Buy"
		ph_ord_type = "Sell"
    ElseIf net_qty < 0 Then
        If open_qty >= 0 Then ord_type_A = "Sell" Else ord_type_A = "SellShort"
		ph_ord_type = "Buy"
    End If

    ' Only trading for ERN or PEAD
    If ern_qty = 0 Or pead_qty = 0 Then
        ' Set tradeID
        If ern_qty <> 0 Then
            tradeID = "ERN"
            ram_strategy_id = ern_strategy_id
        ElseIf pead_qty <> 0 Then
            tradeID = "PEAD"
            ram_strategy_id = pead_strategy_id
        End If
        ' Changing sides
        If Sgn(open_qty) <> Sgn(open_qty + net_qty) Then
			If open_qty < 0 Then ord_type_B = "Buy" Else ord_type_B = "SellShort"
            Sheets("FM Upload").Range(Cells(lastRow, 1), Cells(lastRow + 1, 1)).Value = "SPY"
            Sheets("FM Upload").Range(Cells(lastRow, 2), Cells(lastRow + 1, 2)).Value = 0
            Sheets("FM Upload").Range(Cells(lastRow, 3), Cells(lastRow + 1, 3)).Value = tradeID
            Sheets("FM Upload").Range(Cells(lastRow, 4), Cells(lastRow, 4)).Value = "OFF"
            Sheets("FM Upload").Range(Cells(lastRow + 1, 4), Cells(lastRow + 1, 4)).Value = "ON"
            Sheets("FM Upload").Range(Cells(lastRow, 5), Cells(lastRow, 5)).Value = ord_type_A
            Sheets("FM Upload").Range(Cells(lastRow + 1, 5), Cells(lastRow + 1, 5)).Value = ord_type_B
            Sheets("FM Upload").Range(Cells(lastRow, 6), Cells(lastRow, 7)).Value = Abs(open_qty)
            Sheets("FM Upload").Range(Cells(lastRow + 1, 6), Cells(lastRow + 1, 7)).Value = Abs(net_qty + open_qty)
            Sheets("FM Upload").Range(Cells(lastRow, 8), Cells(lastRow + 1, 8)).Value = ram_strategy_id
        ' Staying on the same side
        Else
            Sheets("FM Upload").Cells(lastRow, 1).Value = "SPY"
            Sheets("FM Upload").Cells(lastRow, 2).Value = 0
            Sheets("FM Upload").Cells(lastRow, 3).Value = tradeID
            Sheets("FM Upload").Cells(lastRow, 4).Value = "ON"
            Sheets("FM Upload").Cells(lastRow, 5).Value = ord_type_A
            Sheets("FM Upload").Cells(lastRow, 6).Value = Abs(net_qty)
            Sheets("FM Upload").Cells(lastRow, 7).Value = Abs(net_qty)
            Sheets("FM Upload").Cells(lastRow, 8).Value = ram_strategy_id
        End If

    ' Trading Shares for Both ERN and PEAD
	' Changing sides
    ElseIf Sgn(open_qty) <> Sgn(open_qty + net_qty) Then
		ord_max_qty = WorksheetFunction.Max(Abs(open_qty), Abs(open_qty + net_qty))
		ord_min_qty = WorksheetFunction.Min(Abs(open_qty), Abs(open_qty + net_qty))
		trd_max_qty = WorksheetFunction.Max(Abs(ern_qty), Abs(pead_qty))
		trd_min_qty = WorksheetFunction.Min(Abs(ern_qty), Abs(pead_qty))

		If trd_max_qty = Abs(pead_qty) Then trd_max_tradeid = "PEAD" Else trd_max_tradeid = "ERN"
		If trd_min_qty = Abs(ern_qty) Then trd_min_tradeid = "ERN" Else trd_min_tradeid = "PEAD"

		If trd_max_qty = Abs(pead_qty) Then trd_max_ramid = pead_strategy_id Else trd_max_ramid = ern_strategy_id
		If trd_min_qty = Abs(ern_qty) Then trd_min_ramid = ern_strategy_id Else trd_min_ramid = pead_strategy_id

		If Sgn(net_qty) = 1 And ord_max_qty = Abs(open_qty) Then
			ord_max_ord_type = "BuyToCover"
			ord_min_ord_type = "Buy"
			ph_ord_type = "Sell"
		ElseIf Sgn(net_qty) = 1 And ord_min_qty = Abs(open_qty) Then
			ord_max_ord_type = "Buy"
			ord_min_ord_type = "BuyToCover"
			ph_ord_type = "Sell"
		ElseIf Sgn(net_qty) = -1 And ord_max_qty = Abs(open_qty) Then
			ord_max_ord_type = "Sell"
			ord_min_ord_type = "SellShort"
			ph_ord_type = "Buy"
		ElseIf Sgn(net_qty) = -1 And ord_min_qty = Abs(open_qty) Then
			ord_max_ord_type = "SellShort"
			ord_min_ord_type = "Sell"
			ph_ord_type = "Buy"
		End If

		' Same side both ERN and PEAD
		If Sgn(ern_qty) = Sgn(pead_qty) Then
			Sheets("FM Upload").Range(Cells(lastRow, 1), Cells(lastRow + 2, 1)).Value = "SPY"
			Sheets("FM Upload").Range(Cells(lastRow, 2), Cells(lastRow + 2, 2)).Value = 0
			Sheets("FM Upload").Range(Cells(lastRow, 3), Cells(lastRow + 1, 3)).Value = trd_max_tradeid
			Sheets("FM Upload").Range(Cells(lastRow + 2, 3), Cells(lastRow + 2, 3)).Value = trd_min_tradeid
			Sheets("FM Upload").Range(Cells(lastRow, 4), Cells(lastRow + 2, 4)).Value = "ON"
			Sheets("FM Upload").Cells(lastRow, 5).Value = ord_min_ord_type
			Sheets("FM Upload").Range(Cells(lastRow + 1, 5), Cells(lastRow + 2, 5)).Value = ord_max_ord_type
			Sheets("FM Upload").Range(Cells(lastRow, 6), Cells(lastRow, 7)).Value = ord_min_qty
			Sheets("FM Upload").Range(Cells(lastRow + 1, 6), Cells(lastRow + 1, 7)).Value = ord_max_qty - trd_min_qty
			Sheets("FM Upload").Range(Cells(lastRow + 2, 6), Cells(lastRow + 2, 7)).Value = trd_min_qty
			Sheets("FM Upload").Range(Cells(lastRow, 8), Cells(lastRow + 1, 8)).Value = trd_max_ramid
			Sheets("FM Upload").Cells(lastRow + 2, 8).Value = trd_min_ramid
		' Different sides (Phantom Trade)
		Else
			Sheets("FM Upload").Range(Cells(lastRow, 1), Cells(lastRow + 3, 1)).Value = "SPY"
			Sheets("FM Upload").Range(Cells(lastRow, 2), Cells(lastRow + 1, 2)).Value = 0
			Sheets("FM Upload").Range(Cells(lastRow + 2, 2), Cells(lastRow + 3, 2)).Value = 1
			Sheets("FM Upload").Range(Cells(lastRow, 3), Cells(lastRow + 2, 3)).Value = trd_max_tradeid
			Sheets("FM Upload").Cells(lastRow + 3, 3).Value = trd_min_tradeid
			Sheets("FM Upload").Range(Cells(lastRow, 4), Cells(lastRow + 3, 4)).Value = "ON"
			Sheets("FM Upload").Cells(lastRow, 5).Value = ord_min_ord_type
			Sheets("FM Upload").Range(Cells(lastRow + 1, 5), Cells(lastRow + 2, 5)).Value = ord_max_ord_type
			Sheets("FM Upload").Cells(lastRow + 3, 5).Value = ph_ord_type
			Sheets("FM Upload").Range(Cells(lastRow, 6), Cells(lastRow, 7)).Value = ord_min_qty
			Sheets("FM Upload").Range(Cells(lastRow + 1, 6), Cells(lastRow + 1, 7)).Value = ord_max_qty
			Sheets("FM Upload").Range(Cells(lastRow + 2, 6), Cells(lastRow + 3, 6)).Value = trd_min_qty
			Sheets("FM Upload").Range(Cells(lastRow, 8), Cells(lastRow + 2, 8)).Value = trd_max_ramid
			Sheets("FM Upload").Cells(lastRow + 3, 8).Value = trd_min_ramid
		End If
	' Not changing sides but trading for both
	' Same side ERN and PEAD
	ElseIf Sgn(ern_qty) = Sgn(pead_qty) Then
		Sheets("FM Upload").Range(Cells(lastRow, 1), Cells(lastRow + 1, 1)).Value = "SPY"
		Sheets("FM Upload").Range(Cells(lastRow, 2), Cells(lastRow + 1, 2)).Value = 0
		Sheets("FM Upload").Cells(lastRow, 3).Value = "ERN"
		Sheets("FM Upload").Cells(lastRow + 1, 3).Value = "PEAD"
		Sheets("FM Upload").Range(Cells(lastRow, 4), Cells(lastRow + 1, 4)).Value = "ON"
		Sheets("FM Upload").Range(Cells(lastRow, 5), Cells(lastRow + 1, 5)).Value = ord_type_A
		Sheets("FM Upload").Range(Cells(lastRow, 6), Cells(lastRow, 7)).Value = Abs(ern_qty)
		Sheets("FM Upload").Range(Cells(lastRow + 1, 6), Cells(lastRow + 1, 7)).Value = Abs(pead_qty)
		Sheets("FM Upload").Cells(lastRow, 8).Value = ern_strategy_id
		Sheets("FM Upload").Cells(lastRow + 1, 8).Value = pead_strategy_id
	' Phantom trade required
	Else
		If Sgn(ern_qty) = Sgn(net_qty) Then
			trdA = "ERN"
			trdB = "PEAD"
			ph_qty = Abs(pead_qty)
			trdA_strat_id = ern_strategy_id
			trdB_strat_id = pead_strategy_id
		Else
			trdA = "PEAD"
			trdB = "ERN"
			ph_qty = Abs(ern_qty)
			trdA_strat_id = pead_strategy_id
			trdB_strat_id = ern_strategy_id
		End If
		Sheets("FM Upload").Range(Cells(lastRow, 1), Cells(lastRow + 2, 1)).Value = "SPY"
		Sheets("FM Upload").Cells(lastRow, 2).Value = 0
		Sheets("FM Upload").Range(Cells(lastRow + 1, 2), Cells(lastRow + 2, 2)).Value = 1
		Sheets("FM Upload").Range(Cells(lastRow, 3), Cells(lastRow + 1, 3)).Value = trdA
		Sheets("FM Upload").Cells(lastRow + 2, 3).Value = trdB
		Sheets("FM Upload").Range(Cells(lastRow, 4), Cells(lastRow + 1, 4)).Value = "ON"
		Sheets("FM Upload").Cells(lastRow + 2, 4).Value = "OFF"
		Sheets("FM Upload").Range(Cells(lastRow, 5), Cells(lastRow + 1, 5)).Value = ord_type_A
		Sheets("FM Upload").Cells(lastRow + 2, 5).Value = ph_ord_type
		Sheets("FM Upload").Range(Cells(lastRow, 6), Cells(lastRow, 7)).Value = Abs(net_qty)
		Sheets("FM Upload").Range(Cells(lastRow + 1, 6), Cells(lastRow + 2, 6)).Value = ph_qty
		Sheets("FM Upload").Range(Cells(lastRow, 8), Cells(lastRow + 1, 8)).Value = trdA_strat_id
		Sheets("FM Upload").Cells(lastRow + 2, 8).Value = trdB_strat_id
	End If
End Sub

Function spyTradeQuantities() As Range
    ' Get SPY Quantities
    Dim spy_shares As Range
    Sheets("Summary").Select
    col_ix_trade_shares = SectionCol("Summary", "SPY", "Trade Shares").Column - 1
    Set spy_shares = SectionLim("Summary", "SPY", "start").Offset(-1, col_ix_trade_shares)
    Set spyTradeQuantities = Range(spy_shares.Address, spy_shares.Offset(1, 0).Address)

End Function

Sub trd_Import_FM_Positions()

    Const strFileName = "J:\Common Folders\Roundabout\Operations\Roundabout Accounting\Fund Manager 2016\Fund Manager Export 2016.csv"

    answer = MsgBox("Import file from: " & strFileName, vbYesNo + vbQuestion, "Source for FM Positions")

    If answer = vbYes Then
        Dim wbkS As Workbook
        Dim wshS As Worksheet
        Dim wshT As Worksheet
        Set wshT = Worksheets("FundManager")
        Set wbkS = Workbooks.Open(Filename:=strFileName)
        Set wshS = wbkS.Worksheets(1)
        wshT.Rows(1 & ":" & wshT.Rows.Count).Value = ""
        wshS.UsedRange.Copy Destination:=wshT.Range("A1")
        wbkS.Close SaveChanges:=False
    End If

End Sub


Sub trdImportDailySummaryFiles()

    Const strErnSectorPath = "Q:\QUANT\DATA\earnings\implementation\processed\SectorInfo.csv"
    Const strPeadSectorPath = "Q:\QUANT\DATA\pead\implementation\processed\SectorInfo.csv"
    Const strErnSummaryPath = "Q:\QUANT\DATA\earnings\implementation\processed\SummaryInfo.csv"
    Const strPeadSummaryPath = "Q:\QUANT\DATA\pead\implementation\processed\SummaryInfo.csv"

    answer = MsgBox("Build trade summary sheet?", vbYesNo + vbQuestion, "Trade Summary")

    If answer = vbYes Then
        Dim wbkS As Workbook
        Dim wshS As Worksheet
        Dim wshT As Worksheet

        Set wshT = Worksheets("ErnSector")
        Set wbkS = Workbooks.Open(Filename:=strErnSectorPath)
        Set wshS = wbkS.Worksheets(1)
        wshT.Rows(1 & ":" & wshT.Rows.Count).Value = ""
        wshS.UsedRange.Copy Destination:=wshT.Range("A1")
        wbkS.Close SaveChanges:=False

        Set wshT = Worksheets("ErnSummary")
        Set wbkS = Workbooks.Open(Filename:=strErnSummaryPath)
        Set wshS = wbkS.Worksheets(1)
        wshT.Rows(1 & ":" & wshT.Rows.Count).Value = ""
        wshS.UsedRange.Copy Destination:=wshT.Range("A1")
        wbkS.Close SaveChanges:=False

        Set wshT = Worksheets("PeadSector")
        Set wbkS = Workbooks.Open(Filename:=strPeadSectorPath)
        Set wshS = wbkS.Worksheets(1)
        wshT.Rows(1 & ":" & wshT.Rows.Count).Value = ""
        wshS.UsedRange.Copy Destination:=wshT.Range("A1")
        wbkS.Close SaveChanges:=False

        Set wshT = Worksheets("PeadSummary")
        Set wbkS = Workbooks.Open(Filename:=strPeadSummaryPath)
        Set wshS = wbkS.Worksheets(1)
        wshT.Rows(1 & ":" & wshT.Rows.Count).Value = ""
        wshS.UsedRange.Copy Destination:=wshT.Range("A1")
        wbkS.Close SaveChanges:=False

        Sheets("Daily Trade Summary").Activate
    End If


End Sub

