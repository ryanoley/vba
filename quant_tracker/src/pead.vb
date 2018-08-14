
Sub pead_FillInPuttingOn()
'
' FillInPuttingOn Macro
'

'   Check Sheet
    VerifySheet ("PEAD")
    
    Dim StartCell As Range
    Set StartCell = SectionLim("PEAD", "Putting On", "start")
    
    If IsEmpty(StartCell.Offset(1, 0).Value) Then
        MsgBox "No new trades found"
    Else
        colEODEstIX = SectionCol("PEAD", "Putting On", "FM Shares").Column - 1
        
        formulaS = StartCell.Offset(0, 1).Address
        formulaE = StartCell.Offset(0, colEODEstIX).Address
        Range(formulaS, formulaE).Select
    
        Dim LastCell As Range
        Set LastCell = SectionLim("PEAD", "Putting On", "end")
        fillE = LastCell.Offset(0, colEODEstIX).Address
        Selection.AutoFill Destination:=Range(formulaS, fillE), Type:=xlFillDefault

        remRow = StartCell.Row
        ActiveSheet.Rows(remRow).Delete
        
        ActiveSheet.EnableCalculation = False
        ActiveSheet.EnableCalculation = True
        Application.Calculate
        Application.RTD.RefreshData
        Application.OnTime Now + TimeValue("0:00:5"), "pead_FillPutOn2"
        
    End If

End Sub

Public Sub pead_FillPutOn2()
    
    Dim StartCell As Range
    Set StartCell = SectionLim("PEAD", "Putting On", "start")
    Dim LastCell As Range
    Set LastCell = SectionLim("PEAD", "Putting On", "end")
    
    Dim tgtShares As Range
    Set tgtShares = SectionColData("PEAD", "Putting On", "Target (Shares)")
    tgtShares.Copy

    colKDOrderIX = SectionCol("PEAD", "Putting On", "RT Order").Column - 1
    orderS = StartCell.Offset(0, colKDOrderIX).Address
    Range(orderS).PasteSpecial xlPasteValues

    Dim tradedShrs As Range
    Set tradedShrs = SectionColData("PEAD", "Putting On", "Traded Shares")
    tradedShrs.Copy

    colPreExShrsIX = SectionCol("PEAD", "Putting On", "Pre-Existing Shares").Column - 1
    existingS = StartCell.Offset(0, colPreExShrsIX).Address
    Range(existingS).PasteSpecial xlPasteValues

End Sub


Sub pead_FillInHolding()
'
'   Copy Tickers and Positions from Putting On to Holding
'

'   Check Sheet
    VerifySheet ("PEAD")

    answer = MsgBox("Move Positions From Putting On To Holding?", vbYesNo + vbQuestion, _
                    "Move Trades To Holding")
    If answer = vbNo Then
        End
    End If
    
    Dim putonStart As Range
    Set putonStart = SectionLim("PEAD", "Putting On", "start")
    Dim putonEnd As Range
    Set putonEnd = SectionLim("PEAD", "Putting On", "end")
    Dim holdStart As Range
    Set holdStart = SectionLim("PEAD", "Holding", "start")
    Dim holdEnd As Range
    Set holdEnd = SectionLim("PEAD", "Holding", "end")
    
    ' Insert Rows into Holding section
    nNew = putonEnd.Row - putonStart.Row + 1
    If putonStart.Value = "NUGT" Then
        End
    End If
    holdEnd.Offset(1, 0).EntireRow.Resize(nNew).Insert

    ' Copy the tickers, shares and RAM IDs from Putting On
    Dim PutOnTickers As Range
    Set PutOnTickers = SectionColData("PEAD", "Putting On", "Ticker")
    PutOnTickers.Copy
    holdEnd.Offset(1, 0).PasteSpecial xlPasteValues
    
    Dim PutOnShares As Range
    Set PutOnShares = SectionColData("PEAD", "Putting On", "RT Order")
    PutOnShares.Copy
    colPositionShrsIX = SectionCol("PEAD", "Holding", "Position (Shares)").Column - 1
    holdEnd.Offset(1, colPositionShrsIX).PasteSpecial xlPasteValues
    
    Dim PutOnRAMID As Range
    Set PutOnRAMID = SectionColData("PEAD", "Putting On", "RAM ID")
    PutOnRAMID.Copy
    col_ramid_ix = SectionCol("PEAD", "Holding", "RAM ID").Column - 1
    holdEnd.Offset(1, col_ramid_ix).PasteSpecial xlPasteValues
    
    ' Fill In Formulas in Holding
    colFMSharesIX = SectionCol("PEAD", "Holding", "FM Shares").Column - 1
    formulaStart = holdStart.Offset(0, 1).Address
    formulaEnd = holdStart.Offset(0, colFMSharesIX).Address
    Range(formulaStart, formulaEnd).Select
    
    Dim holdEndNew As Range
    Set holdEndNew = SectionLim("PEAD", "Holding", "end")
    fillEnd = holdEndNew.Offset(0, colFMSharesIX).Address
    Selection.AutoFill Destination:=Range(formulaStart, fillEnd), Type:=xlFillDefault
    
    ' Initiate Hold Day Column to 1
    colHoldDayIX = SectionCol("PEAD", "Holding", "Hold Day").Column - 1
    hdayStart = holdEnd.Offset(1, colHoldDayIX).Address
    hdayEnd = holdEndNew.Offset(0, colHoldDayIX).Address
    Range(hdayStart, hdayEnd).Value = 1
    
    ' Only delete first row if it is NUGT
    If holdStart.Value = "NUGT" Then
        Rows(holdStart.Row).Delete
    End If
    

End Sub

Sub pead_FillInTakingOff()
'
' Move trades into Taking Off Section from Holding Section
'

'   Check Sheet
    VerifySheet ("PEAD")
    
    answer = MsgBox("Move Holding to Taking Off?", vbYesNo + vbQuestion, "Clear Out Taking Off")
    If answer = vbNo Then
        End
    End If

    'Delete AllRows in Taking Off Section besides the first
    Dim TakeOffStart As Range
    Set TakeOffStart = SectionLim("PEAD", "Taking Off", "start")
    remRow = TakeOffStart.Offset(1, 0).Row
    With ActiveSheet
        .Rows(remRow & ":" & .Rows.Count).Delete
    End With
       
    'Iterate through rows in Holding
    Dim HoldingStart As Range
    Set HoldingStart = SectionLim("PEAD", "Holding", "start")
    Dim HoldingEnd As Range
    Set HoldingEnd = SectionLim("PEAD", "Holding", "end")
    Dim TakeOffEnd As Range
    colHldTkrIX = SectionCol("PEAD", "Holding", "Ticker").Column
    colHldDayIX = SectionCol("PEAD", "Holding", "Hold Day").Column
    colHldQtyIX = SectionCol("PEAD", "Holding", "Position (Shares)").Column
    colRAMIDIX = SectionCol("PEAD", "Holding", "RAM ID").Column
    colTakeOffQtyIX = SectionCol("PEAD", "Taking Off", "Position (Shares)").Column
    colTakeOffRAMIDIX = SectionCol("PEAD", "Taking Off", "RAM ID").Column
    
    'If Nothing in Holding then End
    TakeOffStart.Value = "NUGT"
    TakeOffStart.Offset(0, colTakeOffQtyIX - 1).Value = 0
    If HoldingStart.Value = "NUGT" Then
        End
    End If
    
    'Add row to End of Holding in case all rows are moved
    Rows(HoldingEnd.Row).Select
    Selection.Copy
    Rows(HoldingEnd.Row + 1).Select
    Selection.Insert Shift:=xlDown
    HoldingEnd.Offset(1, 0).Value = "NUGT"
    HoldingEnd.Offset(1, colHldDayIX - 1).Value = 0
    HoldingEnd.Offset(1, colHldQtyIX - 1).Value = 0
    
    For i = HoldingEnd.Row To HoldingStart.Row Step -1
        Set TakeOffEnd = SectionLim("PEAD", "Taking Off", "end")
        tkr = Cells(i, colHldTkrIX).Value
        qty = Cells(i, colHldQtyIX).Value
        hday = Cells(i, colHldDayIX).Value
        ramid = Cells(i, colRAMIDIX).Value
        
        If qty < 0 Or hday = 2 Then
            'Move Shorts or Longs After 2 days
            TakeOffEnd.Offset(1, 0).Value = tkr
            TakeOffEnd.Offset(1, colTakeOffQtyIX - 1) = qty
            TakeOffEnd.Offset(1, colTakeOffRAMIDIX - 1) = ramid
            Rows(i).Delete
        ElseIf qty > 0 And hday = 1 Then
            'Longs remain in Hold after Day 1
            Cells(i, colHldDayIX).Value = 2
        End If
        
    Next i
    
    ' Fill In Formulas in Taking Off
    Set TakeOffEnd = SectionLim("PEAD", "Taking Off", "end")
    
    If TakeOffEnd.Value <> "NUGT" Then
        colFMSharesIX = SectionCol("PEAD", "Taking Off", "FM Shares").Column - 1
        formulaStart = TakeOffStart.Offset(0, 1).Address
        formulaEnd = TakeOffStart.Offset(0, colFMSharesIX).Address
        Range(formulaStart, formulaEnd).Select
        fillEnd = TakeOffEnd.Offset(0, colFMSharesIX).Address
        Selection.AutoFill Destination:=Range(formulaStart, fillEnd), Type:=xlFillDefault
        Rows(TakeOffStart.Row).Delete
    End If
    
    ' As long as all rows not removed, delete inserted row
    If SectionLim("PEAD", "Holding", "start").Value <> "NUGT" Then
        Rows(SectionLim("PEAD", "Holding", "end").Row).Delete
    End If

    
End Sub

Sub pead_CleanUpPuttingOn()
'
'   Clears out Trades from putting on section
'

'   Check Sheet
    VerifySheet ("PEAD")

'   Verify Position data is captured
    answer = MsgBox("All Positions are correct in Holding?", vbYesNo + vbQuestion, _
                    "Clear Out Putting On")
    If answer = vbYes Then
    '   Remove Rows from PuttingOn Section except for the first
        Dim putonStart As Range
        Set putonStart = SectionLim("PEAD", "Putting On", "start")
        Dim putonEnd As Range
        Set putonEnd = SectionLim("PEAD", "Putting On", "end")
        
        If Not IsEmpty(putonStart.Offset(1, 0).Value) Then
            Rows(putonStart.Offset(1, 0).Row & ":" & putonEnd.Row).Delete
        End If

        putonStart.Value = "NUGT"
        
        colAllocDolIX = SectionCol("PEAD", "Putting On", "Alloc ($)").Column - 1
        putonStart.Offset(0, colAllocDolIX).Value = 0
        
        colKDOrderIX = SectionCol("PEAD", "Putting On", "RT Order").Column - 1
        putonStart.Offset(0, colKDOrderIX).Value = 0
        
        colPreExtShrsIX = SectionCol("PEAD", "Putting On", "Pre-Existing Shares").Column - 1
        putonStart.Offset(0, colPreExtShrsIX).Value = 0

    End If
    
End Sub

