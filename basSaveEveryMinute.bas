Attribute VB_Name = "basSaveEveryMinute"
'Written by: Jeremy Dean Gerdes
'Norfolk Naval Shipyard
'C105 Health Physicist
'jeremy.gerdes@navy.mil EDIPI:1249388897
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
        
Option Explicit

'Used to call a mouse event, alternatively use SendInput (newer function in user32)
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwEctraInfo As Long)
Const MOUSEEVENTF_RIGHTDOWN = &H8
Const MOUSEEVENTF_RIGHTUP = &H10
Const MOUSEEVENTF_LEFTDOWN = &H2
Const MOUSEEVENTF_LEFTUP = &H4
Const MOUSEEVENTF_MIDDLEDOWN = &H20
Const MOUSEEVENTF_MIDDLEUP = &H40
Const MOUSEEVENTF_MOVE = &H1
Const MOUSEEVENTF_WHEEL = &H800
Const MOUSEEVENTF_ABSOLUTE = &H8000

'We make the dblScheduleTime global so that we can cancel it later
Global gdblScheduleTime As Double

Public Sub ScheduleEventsEveryMinute()
Dim strAndSeconds As String
    strAndSeconds = Round(Rnd() * 55, 0)
    gdblScheduleTime = Now + TimeValue("00:00:" & strAndSeconds) ' or Now + (1 / 24 / 60)
    Excel.Application.OnTime gdblScheduleTime, "SaveThisFile"
End Sub
        
Public Sub SaveThisFile()
    Application.Calculate
    Dim dblIdleTime As Double
    dblIdleTime = GetIdleTime
    If dblIdleTime > 1 Then
        'Only jiggle the mouse if we have been idle for more then a minute
        'The disadvantage of a mouse jiggle allways is that if we are typing it interupts and looses several keystrokes (that is why we check GetIdleTime)
        ThisWorkbook.Activate
        Dim sht As Worksheet
        MouseEventJiggleMouseOneSecond
        ThisWorkbook.Save
        Set sht = ThisWorkbook.ActiveSheet
        sht.Calculate
        Debug.Print "Saved:" & Now() & " after " & dblIdleTime & " idle time"
    Else
       Debug.Print dblIdleTime
    End If
    Wait 20
    'SendKeys "{Up}"
    'Wait 20
    'SendKeys "{Down}"
    ScheduleEventsEveryMinute
End Sub

Public Function GetLastSavedTime()
    Application.Volatile
    GetLastSavedTime = ThisWorkbook.BuiltinDocumentProperties("Last Save Time")
End Function

Public Sub MouseEventJiggleMouseOneSecond()
    SetCursorPos 500, 500
    Excel.ActiveWindow.Activate
    MouseEventsRandomMovement False, False, 2, 50, 2, 1.5
End Sub

Public Sub MouseEventJiggleMouseTenthSecond()
    MouseEventsRandomMovement False, False, 0.1, 4, 2, 1
End Sub

Public Sub MinorMouseJiggleForHours(dblHours As Double)
    MouseEventsRandomMovement False, False, dblSecondsToDraw:=(dblHours * 60 * 60), fSmallMove:=True
End Sub

Public Sub MouseEventsRandomMovement( _
    Optional fPromptAndDrawOnCanvas As Boolean = True, _
    Optional fIncrementStrokeLengthEachSecond As Boolean = True, _
    Optional dblSecondsToDraw As Double = 60, _
    Optional intMaxDistanceFromCenter As Integer = 300, _
    Optional intStrokeLength As Integer = 1, _
    Optional intStrokeLengthMultiplier As Integer = 5, _
    Optional fSmallMove As Boolean = False _
)
    Dim fMsgBoxResult As VbMsgBoxResult
    If fPromptAndDrawOnCanvas Then
        fMsgBoxResult = MsgBox("Center this Message Box over the paint canvas and Click OK " & vbCrLf & "(It should take up most of the screen, and a pen should be selected)", vbOKCancel)
    Else
        'Don't prompt just do as we may call this from automation code...
        fMsgBoxResult = vbOK
    End If
    If fMsgBoxResult = vbOK Then
        If fPromptAndDrawOnCanvas Then
            mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
        End If
        Dim xOffSet As Integer
        Dim yOffSet As Integer
        Dim xRelativeDistanceFromCenter As Integer
        Dim yRelativeDistanceFromCenter As Integer
        Dim intCurrentStrokeLengthMultiplier As Integer
        Dim dblTimeToStop As Double
        Dim intOriginalStrokeLength As Integer
        If intStrokeLength < 1 Then
            intOriginalStrokeLength = 1
        Else
            intOriginalStrokeLength = intStrokeLength
        End If
        dblTimeToStop = Timer + dblSecondsToDraw
        Do While Timer < dblTimeToStop
            If fSmallMove Then
                Select Case xOffSet
                    Case Is > 0
                        xOffSet = xOffSet - 1
                    Case Is < 0
                        xOffSet = xOffSet + 1
                    Case Else
                        If Rnd() < 0.0005 Then '10%
                            xOffSet = CInt(((Rnd()) * 2) - 1)
                        End If
                End Select
                Select Case yOffSet
                    Case Is > 0
                        yOffSet = yOffSet - 1
                    Case Is < 0
                        yOffSet = yOffSet + 1
                    Case Else
                        If Rnd() < 0.0005 Then '10%
                            yOffSet = CInt(((Rnd()) * 2) - 1)
                        End If
                End Select
            Else
                If fIncrementStrokeLengthEachSecond Then
                    intStrokeLength = intOriginalStrokeLength + dblSecondsToDraw - (dblTimeToStop - Timer)
                    If intStrokeLength > 60 Then
                         intStrokeLength = intStrokeLength - 1
                    End If
                End If
                intCurrentStrokeLengthMultiplier = Rnd() * intStrokeLengthMultiplier
                'Using the intStrokelengthMultiplier increases the possibility of no movement,
                'but also increases the stroke length when it does occur
                xOffSet = 0
                Do While intCurrentStrokeLengthMultiplier > 0
                    xOffSet = xOffSet + (((Rnd() * intStrokeLength) - (Rnd() * intStrokeLength)) * intCurrentStrokeLengthMultiplier)
                    intCurrentStrokeLengthMultiplier = intCurrentStrokeLengthMultiplier - 1
                    DoEvents
                Loop
                xRelativeDistanceFromCenter = xRelativeDistanceFromCenter + xOffSet
                Select Case True
                    Case xRelativeDistanceFromCenter > intMaxDistanceFromCenter And xOffSet > 0
                        xOffSet = -Abs(xOffSet)
                        xRelativeDistanceFromCenter = xRelativeDistanceFromCenter + xOffSet * 2
                    Case xRelativeDistanceFromCenter < -intMaxDistanceFromCenter And xOffSet < 0
                        xOffSet = Abs(xOffSet)
                        xRelativeDistanceFromCenter = xRelativeDistanceFromCenter + xOffSet * 2
                End Select
                intCurrentStrokeLengthMultiplier = Rnd() * intStrokeLengthMultiplier
                'Using the intStrokelengthMultiplier increases the possibility of no movement,
                'but also increases the stroke length when it does occur
                yOffSet = 0
                Do While intCurrentStrokeLengthMultiplier > 0
                    yOffSet = yOffSet + (((Rnd() * intStrokeLength) - (Rnd() * intStrokeLength)) * intCurrentStrokeLengthMultiplier)
                    intCurrentStrokeLengthMultiplier = intCurrentStrokeLengthMultiplier - 1
                    DoEvents
                Loop
                yRelativeDistanceFromCenter = yRelativeDistanceFromCenter + yOffSet
                Select Case True
                    Case yRelativeDistanceFromCenter > intMaxDistanceFromCenter And yOffSet > 0
                        yOffSet = -Abs(yOffSet)
                        yRelativeDistanceFromCenter = yRelativeDistanceFromCenter + yOffSet * 2
                    Case yRelativeDistanceFromCenter < -intMaxDistanceFromCenter And yOffSet < 0
                        yOffSet = Abs(yOffSet)
                        yRelativeDistanceFromCenter = yRelativeDistanceFromCenter + yOffSet * 2
                End Select
            End If
            mouse_event MOUSEEVENTF_MOVE, xOffSet, yOffSet, 0, 0
            'We need a delay here or we draw too fast
            DoEvents
        Loop
        If fPromptAndDrawOnCanvas Then
            'Recenter prior to LeftUp
            Do Until yRelativeDistanceFromCenter < 20 _
              And yRelativeDistanceFromCenter > -20 _
              And xRelativeDistanceFromCenter < 20 _
              And xRelativeDistanceFromCenter > -20
                If xRelativeDistanceFromCenter > 20 Then
                    xOffSet = -10
                End If
                If xRelativeDistanceFromCenter < -20 Then
                    xOffSet = 10
                End If
                If yRelativeDistanceFromCenter > 20 Then
                    yOffSet = -10
                End If
                If yRelativeDistanceFromCenter < -20 Then
                    yOffSet = 10
                End If
                xRelativeDistanceFromCenter = xRelativeDistanceFromCenter + xOffSet
                yRelativeDistanceFromCenter = yRelativeDistanceFromCenter + yOffSet
                mouse_event MOUSEEVENTF_MOVE, xOffSet, yOffSet, 0, 0
                DoEvents
            Loop
            mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
        End If
    End If
End Sub

'Example Drawings
Public Sub PaintARandomPicture()
    MouseEventsRandomMovement , , 10, , , 1
End Sub

Public Sub PaintAFakeConstilationMap()
    MouseEventsRandomMovement _
        fIncrementStrokeLengthEachSecond:=False, _
        dblSecondsToDraw:=15, _
        intStrokeLength:=500, _
        intStrokeLengthMultiplier:=1
End Sub

'Vibrating Pen
Public Sub MouseEventJiggleMouseTenSecondsFinishRightUp()
    MouseEventsRandomMovement False, False, 10, 10, 1, 1
    mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
End Sub
