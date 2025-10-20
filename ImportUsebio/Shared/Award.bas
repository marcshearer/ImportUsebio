Attribute VB_Name = "ImportUsebio"
Function Award(MaxAward As Double, Rank As Integer, totalAwards As Integer, roundDP As Integer, ParamArray allRanksArray() As Variant) As Double
        Dim ties As Double
        Dim aggregateAwards As Double
        Dim AllRanks As Variant
        ties = 0
        aggregateAwards = 0
        AllRanks = allRanksArray(0)
        For Each OtherRank In AllRanks
            matchRank = OtherRank
            If matchRank = Rank Then
                If ties = 0 Or Rank + ties <= totalAwards Then
                    ties = ties + 1
                    aggregateAwards = aggregateAwards + SingleAward(MaxAward, matchRank + ties - 1, totalAwards)
                End If
            End If
        Next OtherRank
        Award = Application.WorksheetFunction.RoundUp(aggregateAwards / ties, roundDP)
End Function

Private Function SingleAward(MaxAward As Double, Rank As Variant, totalAwards As Integer) As Double
        If Rank <= totalAwards Then
            SingleAward = (MaxAward / totalAwards) * (totalAwards - Rank + 1)
        Else
            SingleAward = 0
        End If
End Function

Sub CreateCSV()
    Dim DateArray As Range
    Dim TitleRow As Range
    Dim CopyArea As Range
    Dim Height As Integer
    Dim Width As Integer
    Dim CurrentSheet As Worksheet
    Dim CurrentWorkBook As Workbook
    Dim NewWorkbook As Workbook
    Dim NewSheet As Worksheet
    Dim MovedSheet As Worksheet
    Dim Filename As String
    
    If Not CheckErrors Then
        Exit Sub
    End If
    
    Application.DisplayAlerts = False
    
    Filename = ActiveWorkbook.FullName
    Filename = Replace(Filename, ".xlsm", ".csv")
    
    Set DateArray = Range("ImportDateArray")
    Set TitleRow = Range("ImportTitleRow")
    Height = DateArray.Rows.Count
    Width = TitleRow.Columns.Count
    Set CopyArea = Range(TitleRow(1, Width), DateArray(Height, 1))
    CopyArea.Copy
    
    Set CurrentWorkBook = ActiveWorkbook
    Set CurrentSheet = ActiveSheet
    Set NewSheet = Sheets.Add(After:=CurrentSheet)
    NewSheet.Name = "CSV Export"
    NewSheet.Cells(1, 1).PasteSpecial xlPasteValuesAndNumberFormats
    NewSheet.Move
    
    Set NewWorkbook = ActiveWorkbook
    NewWorkbook.SaveAs Filename:=Filename, FileFormat:=xlCSV, Local:=True
    NewWorkbook.Close
    
    CurrentWorkBook.Activate
    CurrentSheet.Activate
    Application.DisplayAlerts = True
End Sub

Sub CopyRaceExport()
    Dim FirstArray As Range
    Dim TitleRow As Range
    Dim CopyArea As Range
    Dim Height As Integer
    Dim Width As Integer
    Dim NewSheet As Worksheet
    Dim CurrentSheet As Worksheet
 
    Application.DisplayAlerts = False
    
    Set CurrentSheet = ActiveSheet
    Set FirstArray = Range("RaceExportFirstArray")
    Set DataRow = Range("RaceExportDataRow")
    Height = FirstArray.Rows.Count
    Width = DataRow.Columns.Count
    Set CopyArea = Range(DataRow(1, Width), FirstArray(Height, 1))
    CopyArea.Copy
    
    Application.DisplayAlerts = True
End Sub

Function FormatNames(ParamArray paramNames() As Variant) As String()
        Dim Count() As Integer
        Dim thisName As Integer
        Dim FirstName As String
        Dim OtherNames As String
        Dim Result() As String
        Dim Names() As Variant
        On Error GoTo error
        
        namelist = paramNames(LBound(paramNames))
        ReDim Result(LBound(namelist) To UBound(namelist), 0)
        ReDim Count(LBound(namelist) To UBound(namelist))
        ReDim Names(LBound(namelist) To UBound(namelist), 1 To (UBound(paramNames) - LBound(paramNames) + 1) / 2)
        'Dedup names
        
        For Row = LBound(Count) To UBound(Count)
            Count(Row) = 0
            For i = LBound(paramNames) To UBound(paramNames) - 1 Step 2
                Duplicate = False
                If i > LBound(Count) Then
                    For j = LBound(paramNames) To i - 2 Step 2
                        If paramNames(i)(Row) = paramNames(j)(Row) And paramNames(i + 1)(Row) = paramNames(j + 1)(Row) Then
                            Duplicate = True
                        End If
                    Next j
                End If
                If Not Duplicate And (paramNames(i)(Row) <> "" Or paramNames(i + 1)(Row) <> "") Then
                        Count(Row) = Count(Row) + 1
                        Names(Row, Count(Row)) = paramNames(i)(Row) + " " + paramNames(i + 1)(Row)
                End If
            Next i
        Next Row
        
        ' Copy to results
        For Row = LBound(Count) To UBound(Count)
            For i = 1 To Count(Row)
                Separator = ""
                If i <> 1 Then
                    If i = Count(Row) Then
                        Separator = " & "
                    Else
                        Separator = ", "
                    End If
               End If
               Result(Row, 0) = Result(Row, 0) + Separator + Names(Row, i)
                 
            Next i
        Next Row
        
        FormatNames = Result
        Exit Function
error:
1        Exit Function
End Function

Function CombinedCategory(OtherNboGold As Boolean, ParamArray paramRanks() As Variant) As String()
        Dim Result() As String
        
        rankList = paramRanks(LBound(paramRanks))
        ReDim Result(LBound(rankList) To UBound(rankList), 0)
        
        For Row = LBound(rankList) To UBound(rankList)         ' 1 to rows
            Result(Row, 0) = ""
            HighestRank = -1
            For i = LBound(paramRanks) To UBound(paramRanks)    '0 to 1 for pairs or 0 to 3 teams
                Rank = paramRanks(i)(Row)
                If OtherNboGold And Rank = 1 Then
                    Rank = 999
                End If
                For lookupRow = 1 To Range("RanksFrom").Rows.Count   '1 to 5
                    If Rank >= CInt(Range("RanksFrom").Cells(lookupRow, 1).Value) Then
                        Value = Range("RanksCategory").Cells(lookupRow, 1).Value
                    End If
                Next
                RankCategory = Value
                
                If RankCategory = "" Then
                    ' Invalid rank - exclude
                    Result(Row, 0) = ""
                    HighestRank = 9999
                Else
                ' Use if higher than previous
                    If HighestRank <> 1 Then
                        If Rank = 1 Or Rank > HighestRank Then
                            Result(Row, 0) = RankCategory
                            HighestRank = Rank
                        End If
                    End If
                End If
            Next
        Next
        CombinedCategory = Result
End Function

Function Category(ParamArray paramRanks() As Variant) As String()
        Dim Result() As String
        
        rankList = paramRanks(LBound(paramRanks))
        ReDim Result(LBound(rankList) To UBound(rankList), 0)
        
        For Row = LBound(rankList) To UBound(rankList)         ' 1 to rows
            Rank = paramRanks(0)(Row)
            For lookupRow = 1 To Range("RanksFrom").Rows.Count   '1 to 5
                If Rank >= CInt(Range("RanksFrom").Cells(lookupRow, 1).Value) Then
                    Value = Range("RanksCategory").Cells(lookupRow, 1).Value
                End If
            Next
            Result(Row, 0) = Value
        Next
        Category = Result
End Function

Function RelativeTo(ToCell As Range, Row, Column) As Variant
        RelativeTo = ToCell.Cells(Row, Column)
End Function

Sub SelectFormatted()
    Dim ColumnArray As Range
    Dim TitleRow As Range
    Dim SaveArea As Range
    Dim Filename As String
    Dim MyWorksheet As Worksheet
    
    Set ColumnArray = Range("FormattedNameArray")
    Set TitleRow = Range("FormattedTitleRow")
    Filename = ActiveWorkbook.FullName
    Filename = Replace(Filename, ".xlsm", ".htm")
    Height = ColumnArray.Rows.Count
    RowWidth = TitleRow.Columns.Count
    Set SaveArea = Range(TitleRow(1, RowWidth), ColumnArray(Height, 1))
    ActiveWorkbook.Worksheets("Formatted").Select
    SaveArea.Select
End Sub
Sub PrintFormatted()
    Call PrintSheet("Formatted", "", " - Master Point Allocations", "&C&14", True)
End Sub

Sub PrintRaceFormatted()
    Call PrintSheet("Race Formatted", " - Race", "", "&C&20Race to Aviemore - ", False)
End Sub

Sub PrintSheet(SheetName As String, FileSuffix As String, TitleSuffix As String, TitlePrefix As String, AdjustSize As Boolean)
    Dim Filename As String
    
    Filename = ActiveWorkbook.FullName
    Filename = Replace(Filename, ".xlsm", FileSuffix + ".pdf")
    
    Title = TitlePrefix + Range("ImportEventDescriptionCell").Value + TitleSuffix
    RowsPerPage = Range("ImportLinesPerPageCell").Value

    ActiveWorkbook.Names("Printing").Value = "=True"
    
    ActiveWorkbook.Worksheets(SheetName).Select
    If AdjustSize Then
        ActiveSheet.ResetAllPageBreaks
        For Row = 1 To (Range("FormattedNameArray").Rows.Count / RowsPerPage)
                ActiveSheet.Rows((Row * RowsPerPage) + 2).PageBreak = xlPageBreakManual
        Next
        If UCase(Range("ImportPageOrientationCell").Value) = "LANDSCAPE" Then
            ActiveSheet.PageSetup.Orientation = xlLandscape
        Else
            ActiveSheet.PageSetup.Orientation = xlPortrait
        End If
    End If
    ActiveSheet.PageSetup.CenterHeader = Title
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=Filename, IncludeDocProperties:=True, OpenAfterPublish:=True
    If AdjustSize Then
        ActiveSheet.ResetAllPageBreaks
    End If
    ActiveWorkbook.Worksheets("Import").Select
    ActiveWorkbook.Names("Printing").Value = "=False"
End Sub

Function IntegerPart(ByVal Value As Variant) As Long
    Dim Result As String
    Result = ""
    For Char = 1 To Len(Value)
        If IsNumeric(Mid(Value, Char, 1)) Or (Result = "" And Mid(Value, Char, 1) = "-") Then
            Result = Result + Mid(Value, Char, 1)
        End If
    Next Char
    If Result = "" Then
        Result = "0"
    End If
    IntegerPart = CLng(Result)
End Function

Function CheckSum(MPs() As Variant, NationalIds() As Variant) As Double
    Dim Total As Double
    On Error GoTo label
    Total = 0
    If LBound(NationalIds) <> LBound(MPs) Or UBound(NationalIds) <> UBound(MPs) Then
        Err.Raise vbObjectError + 1000, , "Arrays must be of same size for checksum"
    End If
    For element = LBound(NationalIds) To UBound(NationalIds)
    If VarType(NationalIds(element, 1)) <> 0 Then
        Total = Total + IntegerPart(NationalIds(element, 1)) * MPs(element, 1)
    End If
    Next element
    CheckSum = Total
    Exit Function
label:
    CheckSum = CVErr(xlErrNA)
End Function

Sub Auto_Open()
    ' Automatically run when Workbook opened to set 'Save External Link Values' to false to keep file size down
    With ActiveWorkbook
        .SaveLinkValues = False
    End With
End Sub

Function SumMaxIf(Data() As Variant, Match() As Variant, MatchValue As String, Count As Integer) As Double
    Dim Indices() As Variant
    SumMaxIf = SumMaxIfInternal(Data, Indices, Match, MatchValue, Count)
End Function

Function SumMaxIfIncluded(Data() As Variant, Match() As Variant, MatchValue As String, Count As Integer, Offset As Integer) As Boolean
    Dim Indices() As Variant
    Dim Result As Boolean
    If Count = 0 Then
        Result = True
    Else
        Unused = SumMaxIfInternal(Data, Indices, Match, MatchValue, Count)
        Result = False
        For Index = LBound(Indices()) To LBound(Indices()) + Count - 1
            If Indices(Index) = Offset Then
                Result = True
            End If
        Next
    End If
    SumMaxIfIncluded = Result
End Function

Function SumMaxIfInternal(Data() As Variant, Indices() As Variant, Match() As Variant, MatchValue As String, Count As Integer) As Double
    Dim Result As Double
    Dim Included() As Variant
    Dim Sorted() As Variant
    Dim IncludeCount As Integer
    ReDim Included(LBound(Data()) To UBound(Data()))
    ReDim Indices(LBound(Data()) To UBound(Data()))
    ReDim IncludeIndices(LBound(Data()) To UBound(Data()))
    IncludeCount = 0
    Result = 0
    For Index = LBound(Data()) To UBound(Data())
        If Match(Index) = MatchValue Then
            IncludeCount = IncludeCount + 1
            Included(IncludeCount) = Data(Index)
            IncludeIndices(IncludeCount) = Index
        End If
    Next
    If IncludeCount > 0 Then
        If Count = 0 Then Count = IncludeCount
        Sorted = Included()
        Quicksort Sorted, IncludeIndices, False
        For Index = 1 To Application.WorksheetFunction.Min(Count, IncludeCount)
            Result = Result + Sorted(Index)
        Next
    End If
    Indices = IncludeIndices()
    SumMaxIfInternal = Result
End Function

Sub Quicksort(Data() As Variant, Indices() As Variant, Optional Ascending As Boolean = True, Optional MinIndex As Integer = -1, Optional MaxIndex As Integer = -1)
Dim midValue As Variant
Dim keepVariant As Variant
Dim keepIntegeer As Integer
Dim Low As Integer
Dim High As Integer
Dim Multiplier As Integer

If MinIndex = -1 Or MaxIndex = -1 Then
    MinIndex = LBound(Data)
    MaxIndex = UBound(Data)
End If
 
If Ascending Then Multiplier = 1 Else Multiplier = -1
 
Low = MinIndex
High = MaxIndex
midValue = Data(Round((MinIndex + MaxIndex) / 2, 0))
 
While (Low <= High)
   While ((Multiplier * Data(Low)) < (Multiplier * midValue) And Low < MaxIndex)
      Low = Low + 1
   Wend
  
   While ((Multiplier * midValue) < (Multiplier * Data(High)) And High > MinIndex)
      High = High - 1
   Wend
 
    If (Low <= High) Then
      keepVariant = Data(Low)
      Data(Low) = Data(High)
      Data(High) = keepVariant
      keepIntegeer = Indices(Low)
      Indices(Low) = Indices(High)
      Indices(High) = keepIntegeer
      Low = Low + 1
      High = High - 1
    End If
  Wend
  If (MinIndex < High) Then
    Quicksort Data, Indices, Ascending, MinIndex, High
  End If
  If (Low < MaxIndex) Then
    Quicksort Data, Indices, Ascending, Low, MaxIndex
  End If
End Sub

Function CheckErrors()
    Dim CurrentSheet As Worksheet
    Set CurrentSheet = ActiveSheet
    Dim Continue As Boolean
    Continue = True
    Continue = CheckSheetErrors("Import", True)
    If Continue Then
        Continue = CheckSheetErrors("Summary")
    End If
    CurrentSheet.Activate
    CheckErrors = Continue
End Function


Function CheckSheetErrors(SheetName As String, Optional RowSummary As Boolean = False) As Boolean
    Dim ImportRange As Range
    Dim Types
    Dim Colors
    Dim Errors() As Integer
    Dim RowErrors() As Integer
    Dim Message As String
    
    GetErrorParameters Types, Colors
    
    ReDim RowErrors(LBound(Types) To UBound(Types))
    ReDim Errors(LBound(Types) To UBound(Types))
    
    For Index = LBound(Types) To UBound(Types)
        Errors(Index) = 0
    Next
    
    Set ImportRange = Worksheets(SheetName).UsedRange
    For Row = 1 To ImportRange.Rows.Count
    
        For Index = LBound(Types) To UBound(Types)
            RowErrors(Index) = 0
        Next
    
        For Column = 1 To ImportRange.Columns.Count
            For Index = LBound(Types) To UBound(Types)
                If ImportRange.Cells(Row, Column).DisplayFormat.Interior.Color = Colors(Index) Then
                    RowErrors(Index) = RowErrors(Index) + 1
                End If
            Next
        Next

        For Index = LBound(Types) To UBound(Types)
            If RowErrors(Index) <> 0 Then
                If RowSummary Then
                    Errors(Index) = Errors(Index) + 1
                    Exit For
                Else
                    Errors(Index) = Errors(Index) + RowErrors(Index)
                End If
            End If
        Next
    Next
    Result = ""
    Started = False
    
    totalErrors = 0
    For Index = LBound(Errors) To UBound(Errors)
        totalErrors = totalErrors + Errors(Index)
    Next
    
    If totalErrors = 0 Then
        CheckSheetErrors = True
    Else
        Message = ""
        For Index = LBound(Types) To UBound(Types)
            AddMessage Message, Errors(Index), Types(Index), vbCr
        Next
        Result = SheetName + " worksheet has "
        If RowSummary Then Result = Result + "rows with " + vbCrLf
        Result = Result + Message
        Worksheets(SheetName).Activate
        If Errors(1) = 0 Then Default = vbDefaultButton2 Else Default = vbDefaultButton1
        Continue = MsgBox(Result + vbCrLf + "Do you want to continue?", vbYesNo + Default)
        CheckSheetErrors = (Continue = vbYes)
    End If
End Function

Sub GetErrorParameters(ByRef Types, ByRef Colors)
    Count = Range("ErrorTypes").Rows.Count
    ReDim Types(1 To Count)
    ReDim Colors(1 To Count)
    For Index = 1 To Count
        Types(Index) = Range("ErrorTypes").Cells(Index, 1)
        Colors(Index) = Range("ErrorTypes").Cells(Index, 2)
    Next
End Sub

Sub AddMessage(ByRef Message As String, ByVal Value As Integer, ByVal Text As String, Optional ByVal Separator As String = ", ")
    If Value <> 0 Then
        If Message <> "" Then Message = Message + Separator
        Message = Message + CStr(Value) + " " + Text
        If Value > 1 Then Message = Message + "s"
    End If
End Sub

