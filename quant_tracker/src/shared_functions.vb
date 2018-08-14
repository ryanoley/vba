

Public Function SaveCSV(SheetName As String, SavePath)
    
    Application.DisplayAlerts = False
    
    Sheets(SheetName).Copy
    ActiveWorkbook.SaveAs SavePath, xlCSV
    ActiveWorkbook.Close SaveChanges:=False
    
    Application.DisplayAlerts = True
    
    MsgBox SheetName & " saved as csv to: " & SavePath
    
End Function

Public Function VerifySheet(sname As String)
    ' Check the active sheet is  what is expected
    SheetName = ActiveSheet.Name
    If SheetName <> sname Then
        MsgBox "Only Run this Procedure from " & sname & " Sheet"
        End
    End If

End Function


Public Function SectionColData(Sheet As String, Section As String, Column As String) As Range
    
    Dim StartSheet As Worksheet
    Set StartSheet = ActiveSheet
    
    Sheets(Sheet).Select
    Dim SecStart As Range
    Set SecStart = SectionLim(Sheet, Section, "start")
    Dim SecEnd As Range
    Set SecEnd = SectionLim(Sheet, Section, "end")
    
    SecColIX = SectionCol(Sheet, Section, Column).Column - 1
    
    SelRangeStart = SecStart.Offset(0, SecColIX).Address
    SelRangeEnd = SecEnd.Offset(0, SecColIX).Address

    Set SectionColData = Range(SelRangeStart, SelRangeEnd)
    SectionColData.Select
    
    StartSheet.Select

End Function


Public Function SectionLim(Sheet As String, Section As String, Position As String) As Range
    
    Dim StartSheet As Worksheet
    Set StartSheet = ActiveSheet

    Sheets(Sheet).Select
    With Application.FindFormat.Font
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .Underline = xlUnderlineStyleDouble
        .TintAndShade = 0
    End With
        
    C = Cells.Find(What:=Section, LookAt:=xlWhole, MatchCase:=True, SearchFormat:=True).Offset(3, 0).Address
    
    If LCase(Position) = "start" Then
        Set SectionLim = Range(C)
    ElseIf LCase(Position) = "end" Then
        If Range(C).Offset(1, 0).Value = "" Then
            Set SectionLim = Range(C)
        Else
            Set SectionLim = Range(C).End(xlDown)
        End If
    End If

    StartSheet.Select
End Function


Public Function SectionCol(Sheet As String, Section As String, Column As String) As Range
    
    Dim secHeader As Range
    Set secHeader = SectionLim(Sheet, Section, "start").Offset(-2, 0)
    
    With Application.FindFormat.Font
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
    End With
    
    Set res = secHeader.EntireRow.Find(What:=Column _
                                        , LookIn:=xlValues _
                                        , LookAt:=xlPart _
                                        , SearchOrder:=xlByColumns _
                                        , SearchDirection:=xlNext _
                                        , MatchCase:=False _
                                        , SearchFormat:=True)
    Set SectionCol = res

End Function


Public Function get_strategy_id(filter As String) As String
    Dim trade_ramids As Range
    Dim split_array() As String
    Set trade_ramids = SectionColData("Trades", "Trades Base", "RAM ID")

    For Each Cell In trade_ramids
        split_array = Split(Cell.Value, "_")

        pos = InStr(split_array(2), filter)

        If pos > 0 Then
            get_strategy_id = split_array(2)
            Exit Function
        End If
    Next Cell
    
    get_strategy_id = "Err Strategy " & filter & " Not Found"
    
End Function
