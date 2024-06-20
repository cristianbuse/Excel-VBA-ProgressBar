VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DemoClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function DoWork(ByVal progress As ProgressBar, ByRef stepCount As Variant) As Variant
    Dim i As Long
    For i = 1 To stepCount
        progress.Info2 = "Running " & i & " out of " & stepCount
        progress.Value = i / stepCount
        If progress.WasCancelled Then Exit Function
    Next
    DoWork = True
End Function
