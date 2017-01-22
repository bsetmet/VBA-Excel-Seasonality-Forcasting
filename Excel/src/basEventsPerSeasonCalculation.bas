Attribute VB_Name = "basEventsPerSeasonCalculation"
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
         
'All of these constants could be replaced by a referance to a unified .ini file
Global Const gcstrCoverSheetName As String = "Cover"
Private Const mcstrPivotTableName As String = "PivotTableResults"
Private Const mcstrDataSheetName As String = "Data"
Private Const mcstrPivotSheetName As String = "Pivot"
Private Const mcstrCalcEventCorrectedSheetName As String = "Event Corrected"
Private Const mcstrCalcSeasonalIndexSheetName As String = "Seasonal Index"
Private Const mcstrCalcSeasonlallyCorrectedDataSheetName As String = "Seasonally Smoothed"
Private Const mcstrCalcForecastSheetName As String = "Forecast Model"
Private Const mcstrResultsChartSheetName As String = "Results"
Private Const mcintPivotDataStartRow As Integer = 3

'Modular Variables
Private mstrFieldNamePeriodicity As String
Private mstrFieldNameEvaluation As String
Private mstrFieldNameDate As String
Private mstrFieldNameOccurrancesPerPeriod As String
Private mrngDynamicField As Range
Private mstrPivotAtResolutionName As String
Private mintForecastPeriod As Integer

Public Sub BuildEventsPerSeasonReport()
On Error GoTo HandleError

' This annylasis uses the Multiplicative Decomposition Model to Deseasonalize the data to smooth.
' Yt=Trend(levels)*Season(t)*Irregular(t)
    
    Const cdblAxisPaddingPercent As Double = 0.01
    Const cintMaxNumberOfXAxisLables As Integer = 25
    
    'Initialize app options
    ThisWorkbook.GetOriginalAppOptions
    ThisWorkbook.SetCustomAppOptions

    'Get Initial Sheets
    Dim shts As Sheets: Set shts = ThisWorkbook.Sheets
    Dim sht As Worksheet: Set sht = shts(mcstrDataSheetName): sht.Activate
    Dim shtCover As Worksheet: Set shtCover = shts(gcstrCoverSheetName)
    
    'Cleanup BestFit line data on cover sheet
    shtCover.Range("H2:" & Cells(4, 126).Address).ClearContents '126 is the maximum number of categories allowed by this tool.

    'Cleanup Seasonality Column
    sht.Range("D:D").ClearContents
    
    'Get sheet defined values
    Dim lngCountOfDataSetRows As Long: lngCountOfDataSetRows = GetLastRow(GetBottomRightCellOfNonBlankCells(sht))
    Set mrngDynamicField = sht.Range(Cells(1, 1), sht.Cells(lngCountOfDataSetRows, 4))
    mstrFieldNameEvaluation = sht.Range("A1").Value
    mstrFieldNameDate = sht.Range("B1").Value
    mstrFieldNameOccurrancesPerPeriod = sht.Range("C1").Value
    
    SetPeriodicity
    
    'Sort original data on date
    sht.Sort.SortFields.Clear
    sht.Sort.SortFields.Add Key:=Range("B2:B" & lngCountOfDataSetRows), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Data").Sort
        .SetRange Range("A1:D" & lngCountOfDataSetRows)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Dim cData As Range
    'Zeros will be ignored later in graphing, so we will set all 0 values to =1*10^-80 or 1E-80
    Dim lngCurrentDataRow As Long
    For lngCurrentDataRow = 2 To lngCountOfDataSetRows
        Set cData = sht.Cells(lngCurrentDataRow, 3)
        If cData.Value = 0 And Not (IsEmpty(cData.Value)) Then
            cData.Formula = "= 1 * 10 ^ -80"
        End If
    Next
    
    'Update Periodicity of Seasons
    sht.Range("D1") = mstrFieldNamePeriodicity
    For lngCurrentDataRow = 2 To lngCountOfDataSetRows
        sht.Cells(lngCurrentDataRow, 4) = "=GetSelectedSeasonFromDate(B" & lngCurrentDataRow & ")"
    Next
    sht.Calculate
    
    '----------------------------------------------------------------------
    'Create First Pivot Table
    '----------------------------------------------------------------------
    Dim shtPivot As Worksheet: Set shtPivot = CreateWorksheet(mcstrPivotSheetName)
    mCreatePivotTable shtPivot, mcstrPivotTableName
    
    'Get pivot table ranges so we can perform our calculations
    Dim rngPivotTable As Range: Set rngPivotTable = GetPivotTableRange(mcstrPivotTableName, shtPivot)
    Dim lngCategoryColumn As Long
    Dim strCategories As String
    For lngCategoryColumn = 1 To UBound(rngPivotTable.Value2, 2)
        If Not IsEmpty(rngPivotTable.Value2(1, lngCategoryColumn)) Then
            strCategories = strCategories & "," & rngPivotTable.Value2(1, lngCategoryColumn)
        End If
    Next
    
    'Remove Leading Comma for array
    strCategories = Right(strCategories, Len(strCategories) - 1)
    Dim strArrayCategories() As String
    strArrayCategories = Split(strCategories, ",")
    Dim lngCategoriesCount: lngCategoriesCount = UBound(strArrayCategories)
    Dim lngPivotRowCount: lngPivotRowCount = GetPivotTableRowCount(mcstrPivotTableName, shtPivot)
    Dim lngPivotColumnCount: lngPivotColumnCount = GetPivotTableColumnCount(mcstrPivotTableName, shtPivot)
    
    'Setup Calculation Variables
    Dim rngResultsHeader As Range
    Dim lngSeasonality As Long
    Dim lngCurrentCategory As Long
'    Dim lngTopResultsRow As Long
'    Dim lngBotomResultsRow As Long
    Dim lngCountOfCategories As Long: lngCountOfCategories = UBound(strArrayCategories()) + 1
    Dim lngDataSetRowCount As Long
    lngDataSetRowCount = GetLastRow(GetBottomRightCellOfNonBlankCells(sht))
    
    '----------------------------------------------------------------------
    'Calculate Events Corrected
    '----------------------------------------------------------------------
    Dim shtCalcEventCorrections As Worksheet: Set shtCalcEventCorrections = CreateWorksheet(mcstrCalcEventCorrectedSheetName)
    Set rngResultsHeader = shtCalcEventCorrections.Cells(1, 1)
    rngResultsHeader.Formula = "Corrected for Events"
    
    For lngSeasonality = 1 To (lngPivotRowCount)
        'For each Category
        For lngCurrentCategory = 1 To lngCountOfCategories
            If lngSeasonality < mcintPivotDataStartRow Then
                shtCalcEventCorrections.Cells(lngSeasonality, lngCurrentCategory + 1) = _
                    shtPivot.Cells(lngSeasonality, 2 * lngCurrentCategory)
            Else
                shtCalcEventCorrections.Cells(lngSeasonality, 1) = _
                    shtPivot.Cells(lngSeasonality, 1)
                If shtPivot.Cells(lngSeasonality, 2 * lngCurrentCategory + 1) <> 0 Then
                    shtCalcEventCorrections.Cells(lngSeasonality, lngCurrentCategory + 1) = "=" & _
                        shtPivot.Cells(lngSeasonality, 2 * lngCurrentCategory) & "/" & _
                        shtPivot.Cells(lngSeasonality, 2 * lngCurrentCategory + 1)
                Else
                    shtCalcEventCorrections.Cells(lngSeasonality, lngCurrentCategory + 1) = "=NA()"
                End If
            End If
        Next
    Next
    SetSheetFormulasToValues shtCalcEventCorrections
    
    '----------------------------------------------------------------------
    'Calculate Seasonal Index
    '----------------------------------------------------------------------
    Dim shtCalcSeasonalIndex As Worksheet: Set shtCalcSeasonalIndex = CreateWorksheet(mcstrCalcSeasonalIndexSheetName)
    Set rngResultsHeader = shtCalcSeasonalIndex.Cells(1, 1)
    rngResultsHeader.Formula = "Seasonal Index"
    For lngSeasonality = 1 To (lngPivotRowCount) - 1
        'For each Category
        For lngCurrentCategory = 1 To lngCountOfCategories
            If lngSeasonality < mcintPivotDataStartRow Then
                If lngSeasonality = 2 Then
                    shtCalcSeasonalIndex.Cells(lngSeasonality, lngCurrentCategory + 1) = "Seasonal Index"
                Else
                    shtCalcSeasonalIndex.Cells(lngSeasonality, lngCurrentCategory + 1) = _
                        shtPivot.Cells(lngSeasonality, 2 * lngCurrentCategory)
                End If
            Else
                shtCalcSeasonalIndex.Cells(lngSeasonality, 1) = _
                    shtPivot.Cells(lngSeasonality, 1)
                On Error Resume Next
                shtCalcSeasonalIndex.Cells(lngSeasonality, lngCurrentCategory + 1) = "=" & _
                    shtCalcEventCorrections.Cells(lngSeasonality, lngCurrentCategory + 1) & "/Average('" & _
                    mcstrCalcEventCorrectedSheetName & "'!" & _
                    shtCalcEventCorrections.Cells(mcintPivotDataStartRow, lngCurrentCategory + 1).Address & ":" & _
                    shtCalcEventCorrections.Cells(lngPivotRowCount - 1, lngCurrentCategory + 1).Address & ")"
                If Err.Number = 13 Then
                    shtCalcSeasonalIndex.Cells(lngSeasonality, lngCurrentCategory + 1) = "=NA()"
                End If
                On Error GoTo HandleError
                
            End If
        Next
    Next
    SetSheetFormulasToValues shtCalcSeasonalIndex
    
    '----------------------------------------------------------------------
    'Create Second Pivot Table to transform the data into a usable data set
    '----------------------------------------------------------------------
    Dim shtPivotDataAtResolution As Worksheet
    mstrPivotAtResolutionName = mcstrPivotSheetName & "(2)"
    Set shtPivotDataAtResolution = CreateWorksheet(mstrPivotAtResolutionName, shtPivot)
    mCreatePivotTabeAtResolution shtPivotDataAtResolution, mstrPivotAtResolutionName
    
    '----------------------------------------------------------------------
    'Calculate Seasonally corrected data for use in building the trend line
    '----------------------------------------------------------------------
    'shtCalcSeasonlallyCorrectedData
    Dim shtCalcSeasonlallyCorrectedData As Worksheet: Set shtCalcSeasonlallyCorrectedData = CreateWorksheet(mcstrCalcSeasonlallyCorrectedDataSheetName)
    Set rngResultsHeader = shtCalcSeasonlallyCorrectedData.Cells(1, 1)
    rngResultsHeader.Formula = "Seasonally Smoothed Data"
    lngCountOfDataSetRows = GetPivotTableRowCount(mstrPivotAtResolutionName, shtPivotDataAtResolution) + 1
    For lngCurrentDataRow = 2 To lngCountOfDataSetRows - 1
        'For each Category
        For lngCurrentCategory = 1 To lngCountOfCategories
            If lngCurrentDataRow < mcintPivotDataStartRow Then
                shtCalcSeasonlallyCorrectedData.Cells(lngCurrentDataRow, lngCurrentCategory + 1) = _
                    shtPivotDataAtResolution.Cells(lngCurrentDataRow, lngCurrentCategory + 1)
            Else
                shtCalcSeasonlallyCorrectedData.Cells(lngCurrentDataRow, 1) = _
                    "=GetSelectedSeasonFromDate('" & mstrPivotAtResolutionName & "'!" & shtPivotDataAtResolution.Cells(lngCurrentDataRow, 1).Address & ")"
                On Error Resume Next
                'Should update to
                '=IF(VLOOKUP($A$6,'Seasonal Index'!$A$3:$B$14,2)=0,"",'Pivot(2)'!B3/VLOOKUP(A3,'Seasonal Index'!A3:C6,2))
                 shtCalcSeasonlallyCorrectedData.Cells(lngCurrentDataRow, lngCurrentCategory + 1) = "=IFERROR('" & mstrPivotAtResolutionName & "'!" & _
                    shtPivotDataAtResolution.Cells(lngCurrentDataRow, lngCurrentCategory + 1).Address & _
                    "/VLookup(" & _
                    Cells(lngCurrentDataRow, 1).Address & ",'" & _
                    mcstrCalcSeasonalIndexSheetName & "'!" & _
                    Cells(mcintPivotDataStartRow, 1).Address & ":" & _
                    Cells(lngCountOfDataSetRows - 1, lngCountOfCategories + 1).Address & "," & _
                    lngCurrentCategory + 1 & "),"""")"
                If Err.Number = 13 Then
                    shtCalcSeasonlallyCorrectedData.Cells(lngCurrentDataRow, lngCurrentCategory + 1) = "=NA()"
                End If
                On Error GoTo HandleError
            End If
        Next
    Next
    SetSheetFormulasToValues shtCalcSeasonlallyCorrectedData
    '----------------------------------------------------------------------
    'Calculate Slope and Intercept of our modified data set's simple linear regression trend line
    '----------------------------------------------------------------------
    Dim strKnownYs As String
    Dim strKnownXs As String
    For lngCurrentCategory = 1 To lngCountOfCategories
    
        shtCover.Cells(2, 8 + lngCurrentCategory).Formula = _
            shtPivotDataAtResolution.Cells(2, lngCurrentCategory + 1).Value

        strKnownYs = "'" & mstrPivotAtResolutionName & "'!" & _
                    Cells(mcintPivotDataStartRow, lngCurrentCategory + 1).Address & ":" & _
                    Cells(lngCountOfDataSetRows - 1, lngCurrentCategory + 1).Address
        strKnownXs = "'" & mstrPivotAtResolutionName & "'!" & _
                    Cells(mcintPivotDataStartRow, 1).Address & ":" & _
                    Cells(lngCountOfDataSetRows - 1, 1).Address
        
        'Calculate Intercepts
        shtCover.Cells(3, 8).Formula = "Intercept"
        '=INTERCEPT('Seasonally Corrected Data'!B3:B9,'Seasonally Corrected Data'!A3:A9
        shtCover.Cells(3, 8 + lngCurrentCategory).Formula = _
            "=INTERCEPT(" & strKnownYs & "," & strKnownXs & ")"
        
        'Calculate Slopes
            shtCover.Cells(4, 8).Formula = "Slope"
        '=INTERCEPT('Seasonally Corrected Data'!B3:B9,'Seasonally Corrected Data'!A3:A9
        shtCover.Cells(4, 8 + lngCurrentCategory).Formula = _
            "=SLOPE(" & strKnownYs & "," & strKnownXs & ")"
            
    Next
    SetSheetFormulasToValues shtCover
    
    '----------------------------------------------------------------------
    'Calculate Seasonal Forecast (aka ((Yt/St)*Tt) where Tt is the trend line calculated as mx+b and (Yt/St) is the seasonal index
    '----------------------------------------------------------------------
    Dim shtCalcForecast As Worksheet: Set shtCalcForecast = CreateWorksheet(mcstrCalcForecastSheetName)
    Set rngResultsHeader = shtCalcForecast.Range("A1")
    rngResultsHeader.Formula = "=Min(" & shtCalcForecast.Range("B3", Cells(lngCountOfDataSetRows, lngCountOfCategories * 2 + 1)).Address & ")"
    Set rngResultsHeader = shtCalcForecast.Range("A2")
    rngResultsHeader.Formula = "=Max(" & shtCalcForecast.Range("B3", Cells(lngCountOfDataSetRows, lngCountOfCategories * 2 + 1)).Address & ")"
    Dim dblForeCastModificationValue As Double
    dblForeCastModificationValue = 0
    Dim lngCountOfForecastRows
    lngCountOfForecastRows = shtCover.Range("F3")
    
    'Append the original source data columns in preperation of building chart
    For lngCurrentDataRow = 1 To lngCountOfDataSetRows - 1
       'For each Category
        For lngCurrentCategory = 1 + lngCountOfCategories To lngCountOfCategories * 2 + 1
            If lngCurrentDataRow < mcintPivotDataStartRow Then
                shtCalcForecast.Cells(lngCurrentDataRow, lngCurrentCategory + 1) = _
                    shtPivotDataAtResolution.Cells(lngCurrentDataRow, lngCurrentCategory - lngCountOfCategories + 1)
            Else
                On Error Resume Next
                shtCalcForecast.Cells(lngCurrentDataRow, lngCurrentCategory + 1) = _
                shtPivotDataAtResolution.Cells( _
                lngCurrentDataRow, lngCurrentCategory - lngCountOfCategories + 1)
                If shtCalcForecast.Cells(lngCurrentDataRow, lngCurrentCategory + 1) = 0 Then
                    shtCalcForecast.Cells(lngCurrentDataRow, lngCurrentCategory + 1) = "=NA()"
                End If
                If Err.Number = 13 Then
                    shtCalcForecast.Cells(lngCurrentDataRow, lngCurrentCategory + 1) = "=NA()"
                End If
                On Error GoTo HandleError
            End If
        Next
    Next
    
    For lngCurrentDataRow = 1 To (lngCountOfDataSetRows - 1) + lngCountOfForecastRows
        If lngCurrentDataRow >= mcintPivotDataStartRow Then
            shtCalcForecast.Cells(lngCurrentDataRow, 1) = _
                "='" & mstrPivotAtResolutionName & "'!" & _
                shtCalcSeasonlallyCorrectedData.Cells(lngCurrentDataRow, 1).Address
        End If
        'For each Category
        For lngCurrentCategory = 1 To lngCountOfCategories + mintForecastPeriod
            If lngCurrentDataRow < mcintPivotDataStartRow Then
                shtCalcForecast.Cells(lngCurrentDataRow, lngCurrentCategory + 1) = _
                    shtPivotDataAtResolution.Cells(lngCurrentDataRow, lngCurrentCategory + 1) & " Forecast"
            Else
                If lngCurrentDataRow > (lngCountOfDataSetRows - 1) Then
                    'Fill the Forecast Date rows
                    shtCalcForecast.Cells(lngCurrentDataRow, 1) = _
                        GetSelectedSeasonAddIncrementDate(shtCalcForecast.Cells(lngCurrentDataRow - 1, 1))
'                    If dblForeCastModificationValue = 0 And _
'                        shtCover.Range("D2") = True _
'                        And IsEmpty(shtCalcForecast.Cells(lngCurrentDataRow - 1, lngCurrentCategory * 2 + 1)) _
'                    Then
'                        dblForeCastModificationValue = _
'                            shtCalcForecast.Cells(lngCurrentDataRow - 1, lngCurrentCategory + 1) - _
'                            shtCalcForecast.Cells(lngCurrentDataRow - 1, lngCurrentCategory + 1 + lngCountOfCategories)
'                    End If
                End If
                On Error Resume Next
                '=VLOOKUP(GetSelectedSeasonFromDate(A3),'Seasonal Index'!$A$3:$C$20,2)*(Cover!$I$4*Forecast!A3+Cover!$I$3)
                shtCalcForecast.Cells(lngCurrentDataRow, lngCurrentCategory + 1) = "=" & _
                    "VLookup(GetSelectedSeasonFromDate(" & _
                    Cells(lngCurrentDataRow, 1).Address & "),'" & _
                    mcstrCalcSeasonalIndexSheetName & "'!" & _
                    Cells(mcintPivotDataStartRow, 1).Address & ":" & _
                    Cells(lngCountOfDataSetRows - 1, lngCountOfCategories + 1).Address & "," & _
                    lngCurrentCategory + 1 & ")*('" & _
                    gcstrCoverSheetName & "'!" & _
                    Cells(4, 8 + lngCurrentCategory).Address & _
                    "*" & Cells(lngCurrentDataRow, 1).Address & _
                    "+'" & gcstrCoverSheetName & "'!" & _
                    Cells(3, 8 + lngCurrentCategory).Address & _
                    ")+" & dblForeCastModificationValue
                If Err.Number = 13 Then
                    shtCalcForecast.Cells(lngCurrentDataRow, lngCurrentCategory + 1) = "=NA()"
                End If
                On Error GoTo HandleError

            End If
        Next
    Next
    SetSheetFormulasToValues shtCalcForecast
     
    '----------------------------------------------------------------------
    'Insert Results Chart
    '----------------------------------------------------------------------
    If lngCountOfCategories < (256 / 2) Then
        
        'Format Date field prior to generating chart
        shtCalcForecast.UsedRange.Columns("A:A").NumberFormat = "m/d/yyyy"
        shtCalcForecast.Range("A1", "A2").NumberFormat = "0"
        
        Dim rngResultsToChart As Range
        Set rngResultsToChart = shtCalcForecast.Range(shtCalcForecast.Cells(2, 1), _
            shtCalcForecast.Cells((lngCountOfDataSetRows - 1) + lngCountOfForecastRows, (lngCountOfCategories * 2) + 1))
        shtCalcForecast.Activate

        Dim chtResults As Chart
        Set chtResults = CreateChart(mcstrResultsChartSheetName, shtCover)
        shtCalcForecast.Calculate
        With chtResults
            'If we can, update the Values Scale
            'Update Max and Min Date Values
            'Calculate X and Y axis Major Units
            Dim dblMinValue As Double
            Dim dblMaxValue As Double
            Dim rngForecastAndData As Range
            Set rngForecastAndData = shtCalcForecast.Range("B3", shtCalcForecast.Cells(lngCountOfDataSetRows - 1 + lngCountOfForecastRows, lngCountOfCategories * 2 + 1))
            Dim c As Range
            dblMinValue = shtCalcForecast.Range("B3")
            For Each c In rngForecastAndData.Cells
                If Not IsEmpty(c) And Not IsError(c) Then
                    If dblMinValue > c.Value Then
                        dblMinValue = c.Value
                    End If
                    If dblMaxValue < c.Value Then
                        dblMaxValue = c.Value
                    End If
                End If
            Next
            
            .SetSourceData Source:=rngResultsToChart
            If dblMaxValue <= dblMinValue Then
                .Axes(xlValue).MajorUnitIsAuto = True
            Else
                'Don't set the max and min untill we get the AutoCaluculated Major unit,
                'then use that value and loop untill we hit the max and min.
                'Add Padding
                dblMinValue = dblMinValue - (cdblAxisPaddingPercent * (dblMaxValue - dblMinValue))
                dblMaxValue = dblMaxValue + (cdblAxisPaddingPercent * (dblMaxValue - dblMinValue))
                Dim dblMajorUnit As Double
                Dim dblScale As Double
                'Y Axis
                .Axes(xlValue).MajorUnitIsAuto = True
                dblMajorUnit = .Axes(xlValue).MajorUnit
                .Axes(xlValue).MajorUnitIsAuto = False
                .Axes(xlValue).MajorUnit = dblMajorUnit
                'Get Max scale
                dblScale = Truncate(dblMaxValue / dblMajorUnit) * dblMajorUnit + dblMajorUnit
                .Axes(xlValue).MaximumScale = dblScale
                'Get Min Scale
                Do While dblScale > dblMinValue
                    dblScale = dblScale - dblMajorUnit
                Loop
                .Axes(xlValue).MinimumScale = Truncate(dblScale / dblMajorUnit) * dblMajorUnit
            End If
            '.ChartTitle.Caption = "Seasonality with Trend Forecast"
            .SetElement (msoElementChartTitleAboveChart)
            .ChartTitle.Text = _
                "Annual Seasonality Forecast " & "with a " + GetSelectedSeasonDeliniation + " Deliniation"
                .ChartTitle.Format.TextFrame2.TextRange.Characters.Font.Size = 18
            .Refresh
        End With
    End If
    
    HideCalculationSheets

    'Reset Tab Visibility
    shtCover.Activate
    shtCover.Range("C2").Select
    'Display Results
    chtResults.Activate
    
ExitHere:
    'Clean up
    Set shts = Nothing
    Set sht = Nothing
    Set chtResults = Nothing
    Set shtCalcEventCorrections = Nothing
    Set shtCalcSeasonalIndex = Nothing
    Set shtCalcForecast = Nothing
    Set shtCalcSeasonlallyCorrectedData = Nothing
    Set shtCover = Nothing
    Set shtPivot = Nothing
    ThisWorkbook.SetOriginalAppOptions
Exit Sub
HandleError:
              MsgBox "Runtime error '" & Err.Number & "'" & vbCrLf & vbCrLf & Err.Description, vbOKOnly, Err.Source, Err.HelpFile, Err.HelpContext
Resume Next:
GoTo ExitHere
End Sub

Public Sub HideCalculationSheets()
On Error Resume Next
    Dim sht As Object
        For Each sht In Sheets
            If sht.Name <> mcstrDataSheetName _
                And sht.Name <> mcstrResultsChartSheetName _
                And sht.Name <> gcstrCoverSheetName _
                And sht.Name <> mcstrCalcForecastSheetName _
            Then
                sht.Visible = False
            End If
        Next
End Sub

Public Sub SetPeriodicity()
Dim shtCover As Worksheet: Set shtCover = ThisWorkbook.Worksheets(gcstrCoverSheetName)
    Select Case shtCover.Cells(2, 3)
        Case 1 'Annually
            mstrFieldNamePeriodicity = "Annualy"
        Case 2 'Geological Seasons
            mstrFieldNamePeriodicity = "Seasonally"
        Case 3 'Fiscal Quarter
            mstrFieldNamePeriodicity = "Fiscal Quarter"
        Case 4 'Monthly
            mstrFieldNamePeriodicity = "Monthly"
        Case 5 'Weekly
            mstrFieldNamePeriodicity = "Weekly"
        Case 6 'Daily
            mstrFieldNamePeriodicity = "Daily"
    End Select
    Set shtCover = Nothing
End Sub

Private Sub mCreatePivotTable(shtPivot As Worksheet, strPivotTableName As String)
Dim objPivots As PivotTables: Set objPivots = shtPivot.PivotTables
Dim PTCache As PivotCache
Dim PT As PivotTable
    Set PTCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, _
        SourceData:=mrngDynamicField)
    shtPivot.Activate
    

    Set PT = objPivots.Add(PTCache, _
        shtPivot.Range("A1"), TableName:=strPivotTableName)
           
    ActiveWorkbook.ShowPivotTableFieldList = True
    With PT
        With .PivotFields(mstrFieldNamePeriodicity)
            .Orientation = xlRowField
            .Position = 1
            .Caption = mstrFieldNamePeriodicity
        End With
        
        With .PivotFields(mstrFieldNameEvaluation)
            .Orientation = xlColumnField
            .Position = 1
            .Caption = mstrFieldNameEvaluation
        End With
        
        With .PivotFields(mstrFieldNameOccurrancesPerPeriod)
            .Orientation = xlDataField
            .Position = 1
            .Caption = "Sum of " & mstrFieldNameOccurrancesPerPeriod
            .Function = xlSum
        End With

        With .PivotFields(mstrFieldNameDate)
            .Orientation = xlDataField
            .Position = 2
            .Caption = "Count of " & mstrFieldNameDate
            .Function = xlCount
        End With
        
        'Adjusting some settings
        .RowGrand = False
        .ColumnGrand = True
        .DisplayFieldCaptions = False
        .HasAutoFormat = False
        .ShowTableStyleRowStripes = False
        .ShowTableStyleColumnStripes = False
        '.CommitChanges

    End With

End Sub



Private Sub mCreatePivotTabeAtResolution(shtPivot As Worksheet, strPivotTableName As String)
Dim objPivots As PivotTables: Set objPivots = shtPivot.PivotTables
Dim PTCache As PivotCache
Dim PT As PivotTable
    Set PTCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, _
        SourceData:=mrngDynamicField)
    Set PT = objPivots.Add(PTCache, _
        shtPivot.Range("A1"), TableName:=strPivotTableName)
           
    ActiveWorkbook.ShowPivotTableFieldList = True
    With PT
        With .PivotFields(mstrFieldNameDate)
            .Orientation = xlRowField
            .Position = 1
            .Caption = mstrFieldNameDate
        End With
        
        With .PivotFields(mstrFieldNameEvaluation)
            .Orientation = xlColumnField
            .Position = 1
            .Caption = mstrFieldNameEvaluation
        End With
        
        With .PivotFields(mstrFieldNameOccurrancesPerPeriod)
            .Orientation = xlDataField
            .Position = 1
            .Caption = "Average of " & mstrFieldNameOccurrancesPerPeriod
            .Function = xlAverage
        End With
        
        'Adjusting some settings
        .RowGrand = False
        .ColumnGrand = False
        .DisplayFieldCaptions = False
        .HasAutoFormat = False
        .ShowTableStyleRowStripes = False
        .ShowTableStyleColumnStripes = False
        '.CommitChanges
    End With
End Sub


Public Sub ToolBuildEventsPerSeasonReport()
Dim intLoopCount  As Integer
    Debug.Print Now
        For intLoopCount = 1 To 10
            BuildEventsPerSeasonReport
        Next
    Debug.Print Now
End Sub
