
Sub ern_FillInPuttingOn()
'
' FillInPuttingOn Macro
'

'   Check Sheet
    VerifySheet ("ERN")
    
'   Proceed
    Dim StartCell As Range
    Set StartCell = SectionLim("ERN", "Putting On", "start")
    
    If IsEmpty(StartCell.Offset(1, 0).Value) Then
        MsgBox "No new trades found"
        End
    End If
    
    colEODEstIX = SectionCol("ERN", "Putting On", "FM Shares").Column - 1
    
    formulaS = StartCell.Offset(0, 1).Address
    formulaE = StartCell.Offset(0, colEODEstIX).Address
    Range(formulaS, formulaE).Select

    Dim LastCell As Range
    Set LastCell = SectionLim("ERN", "Putting On", "end")
    fillE = LastCell.Offset(0, colEODEstIX).Address
    Selection.AutoFill Destination:=Range(formulaS, fillE), Type:=xlFillDefault

    remRow = StartCell.Row
    ActiveSheet.Rows(remRow).Delete
    
    ActiveSheet.EnableCalculation = False
    ActiveSheet.EnableCalculation = True
    Application.Calculate
    Application.RTD.RefreshData
    Application.OnTime Now + TimeValue("0:00:5"), "ern_FillPutOn2"
    

End Sub

Public Sub ern_FillPutOn2()
    
    Dim StartCell As Range
    Set StartCell = SectionLim("ERN", "Putting On", "start")
    Dim LastCell As Range
    Set LastCell = SectionLim("ERN", "Putting On", "end")
    
    Dim tgtShares As Range
    Set tgtShares = SectionColData("ERN", "Putting On", "Target (Shares)")
    tgtShares.Copy

    colKDOrderIX = SectionCol("ERN", "Putting On", "RT Order").Column - 1
    orderS = StartCell.Offset(0, colKDOrderIX).Address
    Range(orderS).PasteSpecial xlPasteValues

    Dim tradedShrs As Range
    Set tradedShrs = SectionColData("ERN", "Putting On", "Traded Shares")
    tradedShrs.Copy

    colPreExShrsIX = SectionCol("ERN", "Putting On", "Pre-Existing Shares").Column - 1
    existingS = StartCell.Offset(0, colPreExShrsIX).Address
    Range(existingS).PasteSpecial xlPasteValues

End Sub

Sub ern_FillInTakingOff()
'
' Move trades into Taking Off Section from Holding Section
'

'   Check Sheet
    VerifySheet ("ERN")

    answer = MsgBox("Move Holding to Taking Off?", vbYesNo + vbQuestion, "Clear Out Taking Off")
    If answer = vbYes Then
    '   Delete AllRows in Taking Off Section besides the first
        Dim tkOffStart As Range
        Set tkOffStart = SectionLim("ERN", "Taking Off", "start")
        
        remRow = tkOffStart.Offset(1, 0).Row
        With ActiveSheet
            .Rows(remRow & ":" & .Rows.Count).Delete
        End With
        
    '   Copy tickers and positions from Holding to Taking Off
        Dim HoldingTickers As Range
        Set HoldingTickers = SectionColData("ERN", "Holding", "Ticker")
        HoldingTickers.Copy
        tkOffStart.Offset(1, 0).PasteSpecial xlPasteValues

        Dim HoldingShares As Range
        Set HoldingShares = SectionColData("ERN", "Holding", "Position (Shares)")
        HoldingShares.Copy
        colPositionShrsIX = SectionCol("ERN", "Taking Off", "Position (Shares)").Column - 1
        tkOffStart.Offset(1, colPositionShrsIX).PasteSpecial xlPasteValues
        
        Dim HoldingRAMID As Range
        Set HoldingRAMID = SectionColData("ERN", "Holding", "RAM ID")
        HoldingRAMID.Copy
        colRAMIDIX = SectionCol("ERN", "Taking Off", "RAM ID").Column - 1
        tkOffStart.Offset(1, colRAMIDIX).PasteSpecial xlPasteValues
        
    '   Fill In Formulas in Taking Off
        Dim tkOffEnd As Range
        Set tkOffEnd = SectionLim("ERN", "Taking Off", "end")
        colFMSharesIX = SectionCol("ERN", "Taking Off", "FM Shares").Column - 1

        formulaStart = tkOffStart.Offset(0, 1).Address
        formulaEnd = tkOffStart.Offset(0, colFMSharesIX).Address
        Range(formulaStart, formulaEnd).Select
        
        fillEnd = tkOffEnd.Offset(0, colFMSharesIX).Address
        Selection.AutoFill Destination:=Range(formulaStart, fillEnd), Type:=xlFillDefault
        
        ActiveSheet.Rows(tkOffStart.Row).Delete
        
    '   Remove Rows from Holding Section except for the first
        Dim holdStart As Range
        Set holdStart = SectionLim("ERN", "Holding", "start")
        
        Dim holdEnd As Range
        Set holdEnd = SectionLim("ERN", "Holding", "end")
        
        If Not IsEmpty(holdStart.Offset(1, 0).Value) Then
            With ActiveSheet
            .Rows(holdStart.Row + 1 & ":" & holdEnd.Row).Delete
            End With
        End If
          
        holdStart.Value = "NUGT"
        holdStart.Offset(0, colPositionShrsIX).Value = 0
    End If
    
End Sub


Sub ern_FillInHolding()
'
'   Copy Tickers and Positions from Putting On to Holding
'

'   Check Sheet
    VerifySheet ("ERN")

    answer = MsgBox("Move Positions From Putting on To Holding?", vbYesNo + vbQuestion, "Move Trades To Holding")
    If answer = vbYes Then
        Dim putonStart As Range
        Set putonStart = SectionLim("ERN", "Putting On", "start")
        Dim putonEnd As Range
        Set putonEnd = SectionLim("ERN", "Putting On", "end")
        Dim holdStart As Range
        Set holdStart = SectionLim("ERN", "Holding", "start")
        
    '   Insert Rows into Holding section
        nNew = putonEnd.Row - putonStart.Row + 1
        holdStart.Offset(1, 0).EntireRow.Resize(nNew).Insert

    '   Copy the tickers, shares, and RAM IDs from Putting On
        Dim PutOnTickers As Range
        Set PutOnTickers = SectionColData("ERN", "Putting On", "Ticker")
        PutOnTickers.Copy
        holdStart.Offset(1, 0).PasteSpecial xlPasteValues
        
        Dim PutOnShares As Range
        Set PutOnShares = SectionColData("ERN", "Putting On", "RT Order")
        PutOnShares.Copy
        colPositionShrsIX = SectionCol("ERN", "Holding", "Position (Shares)").Column - 1
        holdStart.Offset(1, colPositionShrsIX).PasteSpecial xlPasteValues
        
        Dim PutOnID As Range
        Set PutOnID = SectionColData("ERN", "Putting On", "RAM ID")
        PutOnID.Copy
        colRAMID = SectionCol("ERN", "Holding", "RAM ID").Column - 1
        holdStart.Offset(1, colRAMID).PasteSpecial xlPasteValues
        
    '   Fill In Formulas in Holding
        Dim holdEnd As Range
        Set holdEnd = SectionLim("ERN", "Holding", "end")

        colFMSharesIX = SectionCol("ERN", "Holding", "FM Shares").Column - 1
        formulaStart = holdStart.Offset(0, 1).Address
        formulaEnd = holdStart.Offset(0, colFMSharesIX).Address
        Range(formulaStart, formulaEnd).Select
        fillEnd = holdEnd.Offset(0, colFMSharesIX).Address
        Selection.AutoFill Destination:=Range(formulaStart, fillEnd), Type:=xlFillDefault
        
        ActiveSheet.Rows(holdStart.Row).Delete

    End If

End Sub

Sub ern_CleanUpPuttingOn()
'
'   Clears out Trades from putting on section
'

'   Check Sheet
    VerifySheet ("ERN")

'   Verify Position data is captured
    answer = MsgBox("All Positions are correct in Holding?", vbYesNo + vbQuestion, _
                    "Clear Out Putting On")
    If answer = vbYes Then
    '   Remove Rows from PuttingOn Section except for the first
        Dim putonStart As Range
        Set putonStart = SectionLim("ERN", "Putting On", "start")
        Dim putonEnd As Range
        Set putonEnd = SectionLim("ERN", "Putting On", "end")
        
        If Not IsEmpty(putonStart.Offset(1, 0).Value) Then
            Rows(putonStart.Offset(1, 0).Row & ":" & putonEnd.Row).Delete
        End If
          
        putonStart.Value = "NUGT"
        
        colAllocDolIX = SectionCol("ERN", "Putting On", "Alloc ($)").Column - 1
        putonStart.Offset(0, colAllocDolIX).Value = 0
        
        colKDOrderIX = SectionCol("ERN", "Putting On", "RT Order").Column - 1
        putonStart.Offset(0, colKDOrderIX).Value = 0
        
        colPreExtShrsIX = SectionCol("ERN", "Putting On", "Pre-Existing Shares").Column - 1
        putonStart.Offset(0, colPreExtShrsIX).Value = 0

    End If
    
End Sub

