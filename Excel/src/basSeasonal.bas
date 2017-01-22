Attribute VB_Name = "basSeasonal"
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
Public Sub SetChartPeriodicity()
Dim shtCover As Worksheet: Set shtCover = ThisWorkbook.Worksheets(gcstrCoverSheetName)
Dim cForecast As Range: Set cForecast = shtCover.Range("F3")
Dim cHistorical As Range: Set cHistorical = shtCover.Range("F5")
    Select Case shtCover.Range("C2")
        Case 1 'Annually
            cForecast = 1
            cHistorical = 3
        Case 2 'Seasons
            cForecast = 4
            cHistorical = 12
        Case 3 'Fiscal Quarter
            cForecast = 4
            cHistorical = 12
        Case 4 'Monthly
            cForecast = 12
            cHistorical = 36
        Case 5 'Weekly
            cForecast = 26
            cHistorical = 104
        Case 6 'Daily
            cForecast = 31
            cHistorical = 365
    End Select
End Sub

Public Function GetSelectedSeasonFromDate(varDate As Variant) As Variant
Dim shtCover As Worksheet: Set shtCover = ThisWorkbook.Worksheets(gcstrCoverSheetName)
Dim dblDate As Double

    If varDate = "(blank)" Then
        GetSelectedSeasonFromDate = "NA()"
    Else
        dblDate = CDbl(CDate(varDate))
        
        Select Case shtCover.Range("C2")
            Case 1 'Annually
                GetSelectedSeasonFromDate = Year(dblDate)
            Case 2 'Geological Seasons
                GetSelectedSeasonFromDate = GetSeasonFromDate(dblDate)
            Case 3 'Fiscal Quarter
                GetSelectedSeasonFromDate = GetFiscalQuarterFromDate(dblDate)
            Case 4 'Monthly
                GetSelectedSeasonFromDate = Month(dblDate)
            Case 5 'Weekly
                GetSelectedSeasonFromDate = GetIsoWeekNumberFromDate(dblDate)
            Case 6 'Daily
                GetSelectedSeasonFromDate = GetOrdinalDayOfYearFromDate(dblDate)
        End Select
    End If
    
End Function

Public Function GetSelectedSeasonAddIncrementDate(ByVal dtStartDate) As Date
    Select Case ThisWorkbook.Worksheets(gcstrCoverSheetName).Range("C2")
        Case 1 'Annually
            GetSelectedSeasonAddIncrementDate = DateAdd("yyyy", 1, dtStartDate)
        Case 2
            GetSelectedSeasonAddIncrementDate = DateAdd("m", 3, dtStartDate)
        Case 3 'Quarter
            GetSelectedSeasonAddIncrementDate = DateAdd("m", 3, dtStartDate)
        Case 4 'Monthly
            GetSelectedSeasonAddIncrementDate = DateAdd("m", 1, dtStartDate)
        Case 5 'Weekly
            GetSelectedSeasonAddIncrementDate = DateAdd("ww", 1, dtStartDate)
        Case 6 'Daily
            GetSelectedSeasonAddIncrementDate = DateAdd("d", 1, dtStartDate)
    End Select

End Function

Public Function GetSelectedSeasonDeliniation() As String 'Passing a variant in as date
    Select Case ThisWorkbook.Worksheets(gcstrCoverSheetName).Range("C2")
        Case 1 'Annually
            GetSelectedSeasonDeliniation = "Annuall"
        Case 2 'Seasons
            GetSelectedSeasonDeliniation = "Seasonal"
        Case 3 'Fiscal Quarter
            GetSelectedSeasonDeliniation = "Fiscal Quarter"
        Case 4 'Monthly
            GetSelectedSeasonDeliniation = "Monthly"
        Case 5 'Weekly
            GetSelectedSeasonDeliniation = "Weekly"
        Case 6 'Daily
            GetSelectedSeasonDeliniation = "Daily"
    End Select
End Function

Public Function GetIsoWeekNumberFromDate(varDate As Variant) As Variant
'Per ISO 8601 a leap week occures 71 times every 400 year cycle, the last leap year was 2009
'Each Year is 52.1775 weeks long
'This method conforms to the ISO 8601 standard for fiscal weeks per year.

Dim intWeek As Integer
Dim dtDate As Date
dtDate = CDate(varDate)

intWeek = CInt(Truncate((GetOrdinalDayOfYearFromDate(dtDate) - Weekday(dtDate) + 10) / 7))
If intWeek = 0 Then 'Previous Year
    intWeek = 53
End If
GetIsoWeekNumberFromDate = intWeek

End Function

Public Function Truncate(dblNum As Double) As Long

    If InStr(1, CStr(dblNum), ".") > 1 Then
        Truncate = CLng(Split(CStr(dblNum), ".")(0))
    Else
        Truncate = CLng(dblNum)
    End If
    
End Function

Public Function GetOrdinalDayOfYearFromDate(varDate As Variant) As Long
    GetOrdinalDayOfYearFromDate = Truncate(CDbl(varDate - DateSerial(Year(varDate), 1, 1)))

End Function

Public Function GetFiscalQuarterFromDate(varDate As Variant) As String
    GetFiscalQuarterFromDate = GetFiscalQuarterFromMonth(DatePart("m", CDate(varDate), vbMonday, vbFirstFourDays))
End Function

Public Function GetSeasonFromDate(varDate As Variant) As String 'Passing a variant in as date
    GetSeasonFromDate = GetSeasonFromMonth(DatePart("m", CDate(varDate), vbMonday, vbFirstFourDays))
End Function

Public Function GetSeasonFromMonth(varMonth As Variant) As String 'Passing a variant in as date
Dim intMonth As Integer
If IsNumeric(varMonth) Then
    intMonth = Truncate(CDbl(varMonth))
Else
    intMonth = 0
    Dim intTestMonth As Integer
    For intTestMonth = 1 To 12
        If LCase(MonthName(intTestMonth, False)) = LCase(varMonth) Or _
            LCase(MonthName(intTestMonth, True)) = LCase(varMonth) _
        Then
            intMonth = intTestMonth
        End If
    Next
End If
Dim strResult As String
    Select Case True
      Case intMonth >= 3 And intMonth <= 5
        strResult = "Spring"
      Case intMonth >= 6 And intMonth <= 8
        strResult = "Summer"
      Case intMonth >= 9 And intMonth <= 11
        strResult = "Fall"
      Case intMonth = 1 Or intMonth = 2 Or intMonth = 12
        strResult = "Winter"
      Case Else
        strResult = "#N/A"
    End Select
    
    GetSeasonFromMonth = strResult
    
End Function


Public Function GetFiscalQuarterFromMonth(varMonth As Variant) As String 'Passing a variant in as date
Dim intMonth As Integer
If IsNumeric(varMonth) Then
    
    intMonth = Truncate(CDbl(varMonth))
    
Else
    intMonth = 0
    Dim intTestMonth As Integer
    For intTestMonth = 1 To 12
        If LCase(MonthName(intTestMonth, False)) = LCase(varMonth) Or _
            LCase(MonthName(intTestMonth, True)) = LCase(varMonth) _
        Then
            intMonth = intTestMonth
        End If
    Next
End If
Dim strResult As String
    Select Case True
      Case intMonth >= 10 And intMonth <= 12
        strResult = "1st Qtr"
      Case intMonth >= 1 And intMonth <= 3
        strResult = "2nd Qtr"
      Case intMonth >= 4 And intMonth <= 6
        strResult = "3rd Qtr"
      Case intMonth >= 7 And intMonth <= 9
        strResult = "4th Qtr"
      Case Else
        strResult = "#N/A"
    End Select
    
    GetFiscalQuarterFromMonth = strResult
    
End Function

'Private Sub ToolTestSeasons(dblYearsToTest As Double)
''ToolTestSeasons 7900 results in 7.378906 Seconds to Run Test, there are 718445 days of Fall in the next 7900 year(s)
'Dim dblStartTime
'dblStartTime = Timer()
'Dim dblTest As Double
'Dim dblDaysOfFall As Double
'For dblTest = 1 To 365 * dblYearsToTest
'    If (GetSeasonFromDate(Now() + dblTest)) = "Fall" Then
'        dblDaysOfFall = dblDaysOfFall + 1
'    End If
'Next
'
'Debug.Print Timer() - dblStartTime & " Seconds to Run Test, there are " & dblDaysOfFall & " days of Fall in the next " & dblYearsToTest & " year(s)"
'End Sub
