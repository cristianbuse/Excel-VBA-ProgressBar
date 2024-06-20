Attribute VB_Name = "Demo"
Option Explicit

Public Sub DemoMain()
    With New ProgressBar
        .Info1 = "Please wait..."
        .AllowCancel = True
        .BarColor = &H4D6A00
        .ShowTime = True
        .ShowType = vbModal
        Debug.Print .RunMacro(ThisWorkbook, "DoWork", .Self, 3000)
        .ShowType = vbModeless
        .ShowTime = False
        Debug.Print .RunObjMethod(New DemoClass, "DoWork", .Self, 2000)
    End With
End Sub

Public Function DoWork(ByVal progress As ProgressBar, ByRef stepCount As Long) As Boolean
    Dim i As Long
    For i = 1 To stepCount
        progress.Info2 = "Running " & i & " out of " & stepCount
        'Do stuff here
        progress.Value = i / stepCount
        If progress.WasCancelled Then
            'Clean-up code here
            Exit Function
        End If
    Next
    DoWork = True
End Function


