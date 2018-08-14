
Sub import_latest_files()
    Dim start_sheet
    Set start_sheet = Application.ActiveSheet
    import_realtick_raw
    import_ern_pead_raw
    import_statarb_raw
    import_ticker_mappings
    start_sheet.Activate
End Sub


Sub import_realtick_raw()
    Dim file_list As Variant
    Dim file_name As Variant
    Dim split_array() As String
    Dim pos As Integer
    Dim max_file As String
    Dim max_date As Long
    max_date = -1
        
    'Debug.Print "***************MARK****************"
    dir_name = Environ("RAMSHARE") & "\Roundabout\Operations\Roundabout Accounting\RealTick EOD Files"
    file_filter = "Roundabout_Grouped_User_"
    file_list = listfiles(dir_name)
    
    For Each file_name In file_list
        pos = InStr(file_name, file_filter)
        If pos > 0 Then
            split_array = Split(file_name, "_")
            date_string = Left(split_array(3), 8)
            date_long = CLng(date_string)
            If date_long > max_date Then
                max_date = CLng(date_string)
                max_file = file_name
            End If
        End If
    Next
    
    answer = MsgBox("Import RealTick File: " & max_file, vbYesNo + vbQuestion, "RealTick Raw")
    If answer = vbNo Then
        Exit Sub
    End If
    
    ' Copy File into Workbook
    Dim wbkS As Workbook
    Dim wshS As Worksheet
    Dim wshT As Worksheet
    Set wshT = Worksheets("RealTick RAW")
    Set wbkS = Workbooks.Open(Filename:=dir_name & "\" & max_file)
    Set wshS = wbkS.Worksheets(1)
    wshT.Rows(1 & ":" & wshT.Rows.Count).Value = ""
    wshS.UsedRange.Copy Destination:=wshT.Range("A1")
    wbkS.Close SaveChanges:=False
    
    ' Sort by UserID and then Side
    With wshT
    lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
    With .Range("A1", .Cells(1, .Columns.Count).End(xlToLeft))
        .Resize(lastRow).Sort _
                        key1:="UserID", order1:=xlAscending, DataOption1:=xlSortNormal, _
                        key2:="Side", order2:=xlAscending, DataOption2:=xlSortNormal, _
                        Header:=xlYes, MatchCase:=False, Orientation:=xlTopToBottom, SortMethod:=xlPinYin
        End With
    End With
    
End Sub


Sub import_ern_pead_raw()
    Dim file_list As Variant
    Dim file_name As Variant
    Dim split_array() As String
    Dim pos As Integer
    Dim max_file As String
    Dim max_date As Long
    max_date = -1

    ' Get the Max File From dir_name
    dir_name = Environ("DATA") & "\bookkeeper\ern_pead_raw"
    file_filter = "ern_pead_trades"
    file_list = listfiles(dir_name)

    For Each file_name In file_list
        pos = InStr(file_name, file_filter)
        If pos > 0 Then
            split_array = Split(file_name, "_")
            date_string = Left(split_array(0), 8)
            date_long = CLng(date_string)
            If date_long > max_date Then
                max_date = CLng(date_string)
                max_file = file_name
            End If
        End If
    Next
    
    answer = MsgBox("Import ERN PEAD File: " & max_file, vbYesNo + vbQuestion, "ERN PEAD Raw")
    If answer = vbNo Then
        Exit Sub
    End If
    
    ' Copy File into Workbook
    Dim wbkS As Workbook
    Dim wshS As Worksheet
    Dim wshT As Worksheet
    Set wshT = Worksheets("ERN PEAD RAW")
    Set wbkS = Workbooks.Open(Filename:=dir_name & "\" & max_file)
    Set wshS = wbkS.Worksheets(1)
    
    wshT.Range("A1:H" & wshT.Rows.Count).Value = ""
    wshS.UsedRange.Copy Destination:=wshT.Range("A1")
    wbkS.Close SaveChanges:=False
    
End Sub


Sub import_statarb_raw()
    Dim file_list As Variant
    Dim file_name As Variant
    Dim split_array() As String
    Dim pos As Integer
    Dim max_file As String
    Dim max_date As Long
    max_date = -1


    ' Get the Max File From dir_name
    dir_name = Environ("DATA") & "\ramex\processed"
    file_filter = "processed_aggregate"
    file_list = listfiles(dir_name)

    For Each file_name In file_list
        pos = InStr(file_name, file_filter)
        If pos > 0 Then
            split_array = Split(file_name, "_")
            date_string = Left(split_array(0), 8)
            date_long = CLng(date_string)
            If date_long > max_date Then
                max_date = CLng(date_string)
                max_file = file_name
            End If
        End If
    Next
    
    answer = MsgBox("Import StatArb File: " & max_file, vbYesNo + vbQuestion, "StatArb Raw")
    If answer = vbNo Then
        Exit Sub
    End If
    
    ' Copy File into Workbook
    Dim wbkS As Workbook
    Dim wshS As Worksheet
    Dim wshT As Worksheet
    Set wshT = Worksheets("StatArb RAW")
    Set wbkS = Workbooks.Open(Filename:=dir_name & "\" & max_file)
    Set wshS = wbkS.Worksheets(1)
    
    wshT.Rows(1 & ":" & wshT.Rows.Count).Value = ""
    wshS.UsedRange.Copy Destination:=wshT.Range("A1")
    wbkS.Close SaveChanges:=False
    
End Sub

Function listfiles(ByVal sPath As String) As Variant

    Dim vaArray     As Variant
    Dim i           As Integer
    Dim oFile       As Object
    Dim oFSO        As Object
    Dim oFolder     As Object
    Dim oFiles      As Object

    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(sPath)
    Set oFiles = oFolder.Files

    ReDim vaArray(1 To oFiles.Count)
    i = 1
    For Each oFile In oFiles
        vaArray(i) = oFile.Name
        i = i + 1
    Next

    listfiles = vaArray

End Function


Sub import_ticker_mappings()
    Dim ws As Worksheet

    Dim qad_tickers As Range
    Dim find_result As Range
    
    Dim jsonText As String
    Dim jsonObject As Object

    Set ws = Worksheets("TickerConversion")
    ws.Activate
    ws.Rows(2 & ":" & ws.Rows.Count).Delete

    ' QAD to EZE Ticker Mapping
    jsonText = GetFileContent("Q:\QUANT\data\ram\implementation\qad_to_eze_ticker_map.json")
    Set jsonObject = JsonConverter.ParseJson(jsonText)
    
    ws.Cells(1, 1) = "QAD"
    ws.Cells(1, 2) = "EZE"
    ws.Cells(1, 3) = "Bloomberg"
    
    i = 2
    For Each Item In jsonObject
        ws.Cells(i, 1) = Item
        ws.Cells(i, 2) = jsonObject(Item)
        i = i + 1
    Next

    qad_tickers_start = ws.Cells(2, 1).Address
    qad_tickers_end = ws.Cells(1, 1).End(xlDown).Address
    Set qad_tickers = Range(qad_tickers_start, qad_tickers_end)
     
    ' QAD to Bloomberg Ticker Mapping
    jsonText = GetFileContent("Q:\QUANT\data\ram\implementation\qad_to_bbrg_ticker_map.json")
    Set jsonObject = JsonConverter.ParseJson(jsonText)

    For Each Item In jsonObject
        Set find_result = qad_tickers.Find(What:=Item, LookIn:=xlValues, LookAt:=xlWhole, _
                                           MatchCase:=False, SearchFormat:=False)
                                           
        If Not find_result Is Nothing Then
            find_result.Offset(0, 2).Value = jsonObject(Item)
        Else
            ws.Cells(i, 1).Value = Item
            ws.Cells(i, 3).Value = jsonObject(Item)
            i = i + 1
        End If
    Next

End Sub


Function GetFileContent(Name As String) As String
    Dim intUnit As Integer
    
    On Error GoTo ErrGetFileContent
    intUnit = FreeFile
    Open Name For Input As intUnit
    GetFileContent = Input(LOF(intUnit), intUnit)
ErrGetFileContent:
    Close intUnit
    Exit Function
End Function









