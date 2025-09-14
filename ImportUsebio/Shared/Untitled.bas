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
    Dim EventDateRange As Range
    Dim EventDate As Date
    Dim Message As String
    
    Application.DisplayAlerts = False
    
    Filename = ActiveWorkbook.FullName
    Filename = Replace(Filename, ".xlsm", ".csv")
    
    Set EventDateCell = Range("EventDateCell")
    EventDate = CDate(EventDateCell)
          
    If EventDate > Date Then
        Message = "Event Date is in the future." & vbNewLine & "File not created!"
        MsgBox Message
    Else
        If EventDate < Date - 8 Then
            Message = "Event Date is more than 1 week in the past." & vbNewLine & "Are you sure you want to continue?"
            Confirmed = MsgBox(Message, vbYesNo)
            If Confirmed <> vbYes Then
                Exit Sub
            End If
        End If
           
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
    End If
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

Function CombinedCategory(ParamArray paramRanks() As Variant) As String()
        Dim Result() As String
        
        rankList = paramRanks(LBound(paramRanks))
        ReDim Result(LBound(rankList) To UBound(rankList), 0)
        
        For Row = LBound(rankList) To UBound(rankList)         ' 1 to rows
            Result(Row, 0) = ""
            HighestRank = -1
            For i = LBound(paramRanks) To UBound(paramRanks)    '0 to 1 for pairs or 0 to 3 teams
                Rank = paramRanks(i)(Row)
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
    Dim Filename As String
    
    Filename = ActiveWorkbook.FullName
    Filename = Replace(Filename, ".xlsm", ".pdf")
    
    Title = Range("ImportEventDescriptionCell").Value + " - Master Point Allocations"
    RowsPerPage = Range("ImportLinesPerPageCell").Value

    ActiveWorkbook.Names("Printing").Value = "=True"
    
    ActiveWorkbook.Worksheets("Formatted").Select
    ActiveSheet.ResetAllPageBreaks
    For Row = 1 To (Range("FormattedNameArray").Rows.Count / RowsPerPage)
            ActiveSheet.Rows((Row * RowsPerPage) + 2).PageBreak = xlPageBreakManual
    Next
    ActiveSheet.PageSetup.CenterHeader = Title
    If UCase(Range("ImportPageOrientationCell").Value) = "LANDSCAPE" Then
        ActiveSheet.PageSetup.Orientation = xlLandscape
    Else
        ActiveSheet.PageSetup.Orientation = xlPortrait
    End If
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=Filename, IncludeDocProperties:=True, OpenAfterPublish:=True
    ActiveSheet.ResetAllPageBreaks
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


