Attribute VB_Name = "Mod_Main"
Rem -> ***|*********************************************************************|***|
Rem -> ***|                                                                     |***|
Rem -> ***|                      ______                __                       |***|
Rem -> ***|                     | ____ \              |  |                      |***|
Rem -> ***|                     | |   \ \             |  |                      |***|
Rem -> ***|                     | |    \ \            |  |                      |***|
Rem -> ***|                     | |    / /            |  |                      |***|
Rem -> ***|                     | |___/ /     __      |  |______                |***|
Rem -> ***|                     |______/     (__)     |_________|               |***|
Rem -> ***|                                                                     |***|
Rem -> ***|   _______________________________________________________________   |***|
Rem -> ***|                                                                     |***|
Rem -> ***|   Author       : John Fawzy (Dominator Legend)                      |***|
Rem -> ***|   Email        : Dominator_Legand@Yahoo.com                         |***|
Rem -> ***|   Date         : 21/3/2006                                          |***|
Rem -> ***|   Copyrights   : Some of these function not written by me,          |***|
Rem -> ***|                  However, Contents of code must be intact without   |***|
Rem -> ***|                  Change, If this work will used for commercial      |***|
Rem -> ***|                  Purpose please inform me, if you like this code    |***|
Rem -> ***|                  Please Rate It, Thanks                             |***|
Rem -> ***|                                                                     |***|
Rem -> ***|*********************************************************************|***|
Option Explicit
Rem -> *************************************************************************************************************************************************
Rem -> Sub Main Function To Intialize Application
Public Sub Main()
    If InitCommonControlsVB Then
        Frm_Main.Show
    Else
        Beep
        MsgBox "Environment Error [Int.001]" & vbCrLf & vbCrLf & "Error Initializing Windows XP Skin Manifest!", vbCritical + vbSystemModal, "Environment Error"
        End
    End If
End Sub
Rem -> *************************************************************************************************************************************************
Rem -> Function To Check If The Common Controls Library Loaded Or Not
Public Function InitCommonControlsVB() As Boolean
    On Error Resume Next
    Dim iccex As TagInitCommonControlsEx
    With iccex
        .LngSize = LenB(iccex)
        .LngICC = &H200
    End With
    InitCommonControlsEx iccex
    InitCommonControlsVB = (Err.Number = 0)
    On Error GoTo 0
End Function
Rem -> *************************************************************************************************************************************************
Rem -> Delay the execution to specified time
Sub Working(LenTimes As Long)
   Dim lngStar As Long
   lngStar = GetTickCount()
   Do Until GetTickCount() - lngStar > LenTimes
      DoEvents
   Loop
End Sub
