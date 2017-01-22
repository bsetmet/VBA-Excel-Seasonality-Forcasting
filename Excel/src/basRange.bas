Attribute VB_Name = "basRange"
Option Explicit
'Written by: Jeremy Dean Gerdes
'Norfolk Naval Shipyard
'C105 Health Physicist
'jeremy.gerdes@navy.mil
     'CC0 1.0 <https://creativecommons.org/publicdomain/zero/1.0/legalcode>
     'http://www.copyright.gov/title17/
     'In accrordance with 17 U.S.C. § 105 This work is 'noncopyright' or in the 'public domain'
         'Subject matter of copyright: United States Government works
         'protection under this title is not available for
         'any work of the United States Government, but the United States
         'Government is not precluded from receiving and holding copyrights
         'transferred to it by assignment, bequest, or otherwise.
     'as defined by 17 U.S.C § 101
         '...
         'A “work of the United States Government” is a work prepared by an
         'officer or employee of the United States Government as part of that
         'person’s official duties.
         '...
Private Function FindUniqueCellRange(strFindValue As String, wshActiveWorksheet As Worksheet) As Range
Dim c As Range
    Set c = wshActiveWorksheet.Cells.Find(strFindValue, , , xlPart)
    If IsSomething(c) Then
        'Get First Found
        Set FindUniqueCellRange = Cells(c.Areas(1).Row, c.Areas(1).Column)
    End If
                    
End Function

Private Function GetFirstRow(ByRef rngAddress As Range) As Long
Dim c As Range
Dim objArea As Variant
Dim lngLowestRow As Long
    For Each objArea In rngAddress.Areas
        For Each c In objArea.Cells
            If lngLowestRow <> 0 Then
                If c.Row < lngLowestRow Then
                    lngLowestRow = c.Row
                End If
            Else
                lngLowestRow = c.Row
            End If
        Next
    Next
    GetFirstRow = lngLowestRow
End Function

Private Function GetFirstColumn(ByRef rngAddress As Range) As Long
Dim c As Range
Dim objArea As Variant
Dim lngLowestColumn As Long
    For Each objArea In rngAddress.Areas
        For Each c In objArea.Cells
            If lngLowestColumn <> 0 Then
                If c.Column < lngLowestColumn Then
                    lngLowestColumn = c.Column
                End If
            Else
                lngLowestColumn = c.Column
            End If
        Next
    Next
    GetFirstColumn = lngLowestColumn
End Function

Public Function GetLastRow(ByRef rngAddress As Range) As Long
    If IsSomething(rngAddress) Then
        Dim c As Range
        Dim objArea As Variant
        Dim lngHighestRow As Long
        For Each objArea In rngAddress.Areas
            For Each c In objArea.Cells
                If lngHighestRow <> 0 Then
                    If c.Row > lngHighestRow Then
                        lngHighestRow = c.Row
                    End If
                Else
                    lngHighestRow = c.Row
                End If
            Next
        Next
        GetLastRow = lngHighestRow
    End If
End Function

Private Function GetLastColumn(ByRef rngAddress As Range) As Long
Dim c As Range
Dim objArea As Variant
Dim lngHighestColumn As Long
    For Each objArea In rngAddress.Areas
        For Each c In objArea.Cells
            If lngHighestColumn <> 0 Then
                If c.Column > lngHighestColumn Then
                    lngHighestColumn = c.Column
                End If
            Else
                lngHighestColumn = c.Column
            End If
        Next
    Next
    GetLastColumn = lngHighestColumn
End Function

Public Function GetBottomRightCellOfNonBlankCells(Optional ByRef wsh As Worksheet) As Range
Dim rng As Range
    If IsNothing(wsh) Then
        Set wsh = ThisWorkbook.ActiveSheet
    End If
    wsh.Activate
    'we only ever perfom a union of two ranges at a time to ensure that ranges
    'without cells do not cause nothing at all to be selected
    On Error Resume Next
    Set rng = GetUnionOfTwoRanges(rng, wsh.Cells.SpecialCells(xlCellTypeComments))
    Set rng = GetUnionOfTwoRanges(rng, wsh.Cells.SpecialCells(xlCellTypeConstants))
    Set rng = GetUnionOfTwoRanges(rng, wsh.Cells.SpecialCells(xlCellTypeFormulas))

    
    On Error GoTo 0
    'Stop selecting our ranges
    Set GetBottomRightCellOfNonBlankCells = GetBottomRightCellOfRanges(rng)
    
End Function


Public Function GetUnionOfTwoRanges(ByRef rng1 As Range, ByRef rng2 As Range) As Range
    If IsNothing(rng1) Then
        If IsNothing(rng2) Then
            Set GetUnionOfTwoRanges = Nothing
        Else
            Set GetUnionOfTwoRanges = rng2
        End If
    Else
        If IsNothing(rng2) Then
            Set GetUnionOfTwoRanges = rng1
        Else
            Set GetUnionOfTwoRanges = Union(rng1, rng2)
        End If
    End If
End Function

Public Function GetTopLeftCellOfRanges(ByRef rngAddress As Range) As Range
Dim area As Variant
Dim intArea As Integer: intArea = 1
Dim lngLowestRow As Long
Dim lngLowestColumn As Long
Dim lngCurrentRow
Dim lngCurrentColumn
    'Assign the lowest to the first area
    rngAddress.Activate
    lngLowestRow = GetAddressR1C1FirstRow(rngAddress, intArea)
    lngLowestColumn = GetAddressR1C1FirstColumn(rngAddress, intArea)
    intArea = intArea + 1
    
    'Verify none of the other areas have a lower row, or column. Lower the row/column as neccessary
    Do While intArea < rngAddress.Areas.Count
        lngCurrentRow = GetAddressR1C1FirstRow(rngAddress, intArea)
        lngLowestRow = IIf(lngCurrentRow < lngLowestRow, lngCurrentRow, lngLowestRow)
        lngCurrentColumn = GetAddressR1C1FirstColumn(rngAddress, intArea)
        lngLowestColumn = IIf(lngCurrentColumn < lngLowestColumn, lngCurrentColumn, lngLowestColumn)
        intArea = intArea + 1
    Loop

    Set GetTopLeftCellOfRanges = Cells(lngLowestRow, lngLowestColumn)
    
End Function


Public Function GetBottomRightCellOfRanges(ByRef rngAddress As Range) As Range
Dim area As Variant
Dim intArea As Integer: intArea = 1
Dim lngHighestRow As Long
Dim lngHighestColumn As Long
Dim lngCurrentRow
Dim lngCurrentColumn
    'Assign the Highest to the first area
    If IsNothing(rngAddress) Then
        Set GetBottomRightCellOfRanges = Nothing
        Exit Function
    Else
        rngAddress.Activate
    End If

    lngHighestRow = GetAddressR1C1LastRow(rngAddress, intArea)
    lngHighestColumn = GetAddressR1C1LastColumn(rngAddress, intArea)
    intArea = intArea + 1
    
    'Verify none of the other areas have a higher row, or column. Raise the row/column as neccessary
    Do While intArea <= rngAddress.Areas.Count
        lngCurrentRow = GetAddressR1C1LastRow(rngAddress, intArea)
        lngHighestRow = IIf(lngCurrentRow > lngHighestRow, lngCurrentRow, lngHighestRow)
        lngCurrentColumn = GetAddressR1C1LastColumn(rngAddress, intArea)
        lngHighestColumn = IIf(lngCurrentColumn > lngHighestColumn, lngCurrentColumn, lngHighestColumn)
        intArea = intArea + 1
    Loop
    Set GetBottomRightCellOfRanges = Cells(lngHighestRow, lngHighestColumn)
    
End Function

Private Function GetAddressR1C1LastRow(ByRef rngAddress As Range, ByRef intArea As Integer) As Double
    GetAddressR1C1LastRow = rngAddress.Areas(intArea).Row + rngAddress.Areas(intArea).Rows.Count - 1
End Function
 
Private Function GetAddressR1C1LastColumn(ByRef rngAddress As Range, ByRef intArea As Integer) As Double
    GetAddressR1C1LastColumn = rngAddress.Areas(intArea).Column + rngAddress.Areas(intArea).Columns.Count - 1
End Function


Private Function GetAddressR1C1FirstRow(ByRef rngAddress As Range, ByRef intArea As Integer) As Double
    GetAddressR1C1FirstRow = rngAddress.Areas(intArea).Row
End Function

Private Function GetAddressR1C1FirstColumn(ByRef rngAddress As Range, ByRef intArea As Integer) As Double
    GetAddressR1C1FirstColumn = rngAddress.Areas(intArea).Column
End Function

