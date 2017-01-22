VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'Define App Option variables in this modual to restore later
Dim mfCalculateBeforeSave As Boolean
Dim mfEnableAutoRecover As Boolean
Dim mfEnableLargeOperationAlert As Boolean
Dim mxlCalculation As XlCalculation
Dim mrngLastChangedRange As Range
Dim mwshLastChangedSheet As Worksheet

Public Sub SetCustomAppOptions()
                
    'Set Custom App options
    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    Application.CalculateBeforeSave = False
    ThisWorkbook.EnableAutoRecover = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.EnableLargeOperationAlert = False

End Sub
      
Public Sub GetOriginalAppOptions()
    mfCalculateBeforeSave = Application.CalculateBeforeSave
    mfEnableAutoRecover = ThisWorkbook.EnableAutoRecover
    mxlCalculation = Application.Calculation
    mfEnableLargeOperationAlert = Application.EnableLargeOperationAlert

End Sub

Public Sub SetOriginalAppOptions()

    Application.CalculateBeforeSave = mfCalculateBeforeSave
    ThisWorkbook.EnableAutoRecover = mfEnableAutoRecover
    Application.Calculation = mxlCalculation
    Application.EnableLargeOperationAlert = mfEnableLargeOperationAlert
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True ' This must be set to true
    Application.Cursor = xlDefault
    
End Sub

Public Sub ToolResetAppOptions()
    
    'Set Custom App options
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.CalculateBeforeSave = True
    ThisWorkbook.EnableAutoRecover = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableLargeOperationAlert = True
    Application.Cursor = xlDefault

End Sub

