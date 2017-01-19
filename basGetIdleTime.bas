Attribute VB_Name = "basGetIdleTime"
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
Private Declare Sub Sleep Lib "kernel32" (ByVal lngMilliseconds As Long)
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetLastInputInfo Lib "user32" (plii As Any) As Long

Private Type LastInputInformation
    cbSize As Long
    dwTime As Long
End Type

Public Function GetIdleTime() As Double
    Dim liiLastInputInfo As LastInputInformation
    liiLastInputInfo.cbSize = Len(liiLastInputInfo)
    Call GetLastInputInfo(liiLastInputInfo)
    GetIdleTime = (GetTickCount() - liiLastInputInfo.dwTime) / 1000
End Function

Sub Wait(lngMilliseconds As Long)
    Sleep lngMilliseconds
End Sub


Sub WaitSeconds(dblSeconds As Long)
    Sleep dblSeconds * 1000
End Sub
