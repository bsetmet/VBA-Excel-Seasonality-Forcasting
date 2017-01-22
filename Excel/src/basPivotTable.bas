Attribute VB_Name = "basPivotTable"
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
Public Function PivotTableExists(strPivotTableName As String, sht As Worksheet) As Boolean

Dim objPivot As PivotTable
Dim objPivots As PivotTables
Dim fPivotTableFound As Boolean

Set objPivots = sht.PivotTables
For Each objPivot In objPivots
    If LCase(objPivot.Name) = LCase(strPivotTableName) Then
        fPivotTableFound = True
        Exit For
    End If
Next

PivotTableExists = fPivotTableFound

End Function

Public Sub DeletePivotTable(strPivotTableName As String, sht As Worksheet)
    If PivotTableExists(strPivotTableName, ActiveSheet) Then
        With GetPivotTableRange(strPivotTableName, ActiveSheet)
            .ClearContents
            .Delete Shift:=xlToLeft
        End With
    End If
End Sub

Public Function GetPivotTableRange(strPivotTableName As String, sht As Worksheet) As Range
    
    sht.Activate
    sht.Select
    Dim objPivot As PivotTable: Set objPivot = sht.PivotTables(strPivotTableName)
    objPivot.PivotSelect "", xlDataAndLabel, True
    Set GetPivotTableRange = Selection
    
End Function

Public Function GetPivotTableColumnCount(strPivotTableName As String, sht As Worksheet) As Long

Dim rngPivotTable As Range: Set rngPivotTable = GetPivotTableRange(strPivotTableName, sht)
GetPivotTableColumnCount = UBound(rngPivotTable.Value2, 2)

End Function


Public Function GetPivotTableRowCount(strPivotTableName As String, sht As Worksheet) As Long

Dim rngPivotTable As Range: Set rngPivotTable = GetPivotTableRange(strPivotTableName, sht)
GetPivotTableRowCount = UBound(rngPivotTable.Value2, 1)

End Function

Public Sub ToolTestPivotTables()

Dim objPivot As PivotTable
Dim objPivots As PivotTables

Set objPivots = ActiveSheet.PivotTables
For Each objPivot In objPivots
    Debug.Print objPivot.Name
Next

End Sub

Public Function ToolTestPivotChaches() As Boolean

Dim objPivot As PivotCache
Dim objPivots As PivotCaches
    
    For Each objPivot In ThisWorkbook.PivotCaches
        Debug.Print objPivot.Index
        
    Next
    
End Function


