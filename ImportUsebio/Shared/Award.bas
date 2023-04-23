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

Sub CopyImport()
    Dim DateArray As Range
    Dim TitleRow As Range
    Dim CopyArea As Range
    Dim Height As Integer
    Dim Width As Integer
    Set DateArray = Range("ImportDateArray")
    Set TitleRow = Range("ImportTitleRow")
    Height = DateArray.Rows.Count
    Width = TitleRow.Columns.Count
    Set CopyArea = Range(TitleRow(1, Width), DateArray(Height, 1))
    CopyArea.Copy
    
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
Sub SelectFormatted()
    Dim ColumnArray As Range
    Dim TitleRow As Range
    Dim SaveArea As Range
    Dim FileName As String
    Dim MyWorksheet As Worksheet
    
    Set ColumnArray = Range("FormattedNameArray")
    Set TitleRow = Range("FormattedTitleRow")
    FileName = ActiveWorkbook.FullName
    FileName = Replace(FileName, ".xlsm", ".htm")
    Height = ColumnArray.Rows.Count
    RowWidth = TitleRow.Columns.Count
    Set SaveArea = Range(TitleRow(1, RowWidth), ColumnArray(Height + 1, 1))
    ActiveWorkbook.Worksheets("Formatted").Select
    SaveArea.Select
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



