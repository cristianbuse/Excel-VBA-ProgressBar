VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''=============================================================================
''' VBA ProgressBar for Excel
''' -----------------------------------------------------
''' https://github.com/cristianbuse/Excel-VBA-ProgressBar
''' -----------------------------------------------------
''' MIT License
'''
''' Copyright (c) 2022 Ion Cristian Buse
'''
''' Permission is hereby granted, free of charge, to any person obtaining a copy
''' of this software and associated documentation files (the "Software"), to
''' deal in the Software without restriction, including without limitation the
''' rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
''' sell copies of the Software, and to permit persons to whom the Software is
''' furnished to do so, subject to the following conditions:
'''
''' The above copyright notice and this permission notice shall be included in
''' all copies or substantial portions of the Software.
'''
''' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
''' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
''' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
''' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
''' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
''' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
''' IN THE SOFTWARE.
'''=============================================================================

Option Explicit

Private WithEvents m_form As ProgressForm
Attribute m_form.VB_VarHelpID = -1
Private m_allowCancel As Boolean
Private m_cancelled As Boolean
Private m_currentValue As Double
Private m_isAutoCentered As Boolean
Private m_isRunning As Boolean
Private m_procedure As String
Private m_result As Variant
Private m_showTime As Boolean
Private m_showType As FormShowConstants
Private m_startTime As Date
Private m_targetBook As Workbook
Private m_targetObj As Object
Private m_args() As Variant

'Controls
Private m_info1 As MSForms.Label
Private m_info2 As MSForms.Label
Private m_barFrame As MSForms.Frame
Private m_bar As MSForms.Label
Private m_elapsed As MSForms.Label
Private m_remaining As MSForms.Label
Private m_percent As MSForms.Label
Private WithEvents m_escButton As MSForms.CommandButton
Attribute m_escButton.VB_VarHelpID = -1

#If Mac Then
#ElseIf VBA7 Then
Private Declare PtrSafe _
Function rtcCallByName Lib "VBE7.DLL" (ByVal targetObj As Object _
                                     , ByVal procNamePtr As LongPtr _
                                     , ByVal vCallType As VbCallType _
                                     , ByRef args() As Any _
                                     , Optional ByVal lcid As Long) As Variant
#Else
Private Declare _
Function rtcCallByName Lib "msvbvm60" (ByVal targetObj As Object _
                                     , ByVal procNamePtr As Long _
                                     , ByVal vCallType As VbCallType _
                                     , ByRef args() As Any _
                                     , Optional ByVal lcid As Long) As Variant
#End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Class events
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub Class_Initialize()
    Set m_form = New ProgressForm
    BuildForm
    With Me
        .AllowCancel = False
        .Caption = "Progress..."
        .Info1 = "Please wait..."
        .Info2 = vbNullString
        .ShowTime = False
        .ShowType = vbModal
        .CenterOnApplication
    End With
End Sub
Private Sub Class_Terminate()
    TryHideForm
    Set m_form = Nothing
End Sub
Private Sub TryHideForm()
    On Error Resume Next 'Avoid error 402
    m_form.Hide
    On Error GoTo 0
End Sub

'*******************************************************************************
'Builds the necessary controls and alignment at runtime
'*******************************************************************************
Private Sub BuildForm()
    Const progIDLabel As String = "Forms.Label.1"
    Const progIDFrame As String = "Forms.Frame.1"
    Const progIDButton As String = "Forms.CommandButton.1"
    Const sideValue As Single = 6
    '
    m_form.Font.Name = "Tahoma"
    m_form.Font.Size = 8.25
    m_form.Width = 300
    '
    Set m_info1 = m_form.Controls.Add(progIDLabel)
    CastToControl(m_info1).Move sideValue, sideValue
    TextAlignLabel m_info1, False, True, fmTextAlignLeft
    '
    Set m_info2 = m_form.Controls.Add(progIDLabel)
    CastToControl(m_info2).Move sideValue, CastToControl(m_info1).Top + 12
    TextAlignLabel m_info2, False, True, fmTextAlignLeft
    '
    Set m_barFrame = m_form.Controls.Add(progIDFrame)
    CastToControl(m_barFrame).Move sideValue, CastToControl(m_info2).Top + 15 _
        , m_form.InsideWidth - sideValue * 2, 15
    m_barFrame.SpecialEffect = fmSpecialEffectSunken
    '
    Set m_bar = m_barFrame.Controls.Add(progIDLabel)
    CastToControl(m_bar).Move 0, 0, 15, 15
    m_bar.BackColor = &HC07000
    '
    Set m_elapsed = m_form.Controls.Add(progIDLabel)
    CastToControl(m_elapsed).Move sideValue, CastToControl(m_barFrame).Top + 18
    TextAlignLabel m_elapsed, False, True, fmTextAlignLeft
    '
    Set m_remaining = m_form.Controls.Add(progIDLabel)
    CastToControl(m_remaining).Move sideValue, CastToControl(m_elapsed).Top + 12
    TextAlignLabel m_remaining, False, True, fmTextAlignLeft
    m_form.Height = CastToControl(m_remaining).Top + sideValue
    With m_form
       .Height = .Height * 2 - .InsideHeight
    End With
    '
    Set m_percent = m_form.Controls.Add(progIDLabel)
    CastToControl(m_percent).Move CastToControl(m_barFrame).Width _
        + sideValue - 60, CastToControl(m_elapsed).Top, 60
    TextAlignLabel m_percent, False, False, fmTextAlignRight
    '
    Set m_escButton = m_form.Controls.Add(progIDButton)
    With CastToControl(m_escButton)
        .Cancel = True 'Allows for the form to be closed by pressing the Esc key
        .Move 0, 0, 0, 0
    End With
End Sub
Private Function CastToControl(ByVal c As MSForms.Control) As MSForms.Control
    Set CastToControl = c
End Function
Private Sub TextAlignLabel(ByVal labelControl As MSForms.Label _
                         , ByVal wordWrapValue As Boolean _
                         , ByVal autoSizeValue As Boolean _
                         , ByVal textAlignValue As fmTextAlign)
    With labelControl
        .WordWrap = wordWrapValue
        .AutoSize = autoSizeValue
        .TextAlign = textAlignValue
    End With
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Form/Control events
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub m_form_Activate()
    If m_showType = vbModal Then RunProcedure
End Sub
Private Sub m_form_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then 'User pressed X button
        Cancel = True
        OnCancel
    End If
End Sub
Private Sub m_escButton_Click()
    OnCancel
End Sub
Private Sub OnCancel()
    If Not m_allowCancel Then Exit Sub
    '
    If MsgBox(Prompt:="Are you sure you want to cancel?" _
            , Buttons:=vbQuestion + vbYesNo _
            , Title:="Please confirm" _
    ) = vbYes Then
        m_form.Hide
        m_cancelled = True
    End If
End Sub

'===============================================================================
'Caption text
'===============================================================================
Public Property Get Caption() As String
    Caption = m_form.Caption
End Property
Public Property Let Caption(ByVal formCaption As String)
    m_form.Caption = formCaption
    Refresh
End Property

'===============================================================================
'Info1 text
'===============================================================================
Public Property Get Info1() As String
    Info1 = m_info1.Caption
End Property
Public Property Let Info1(ByVal info1Label As String)
    m_info1.Caption = info1Label
    Refresh
End Property

'===============================================================================
'Info2 text
'===============================================================================
Public Property Get Info2() As String
    Info2 = m_info2.Caption
End Property
Public Property Let Info2(ByVal info2Label As String)
    m_info2.Caption = info2Label
    Refresh
End Property

'===============================================================================
'Color of the bar
'===============================================================================
Public Property Get BarColor() As Long
    BarColor = m_bar.BackColor
End Property
Public Property Let BarColor(ByVal colorCode As Long)
    m_bar.BackColor = colorCode
    Refresh
End Property

'===============================================================================
'Color of the frame (bar background)
'===============================================================================
Public Property Get BarBackColor() As Long
    BarBackColor = m_barFrame.BackColor
End Property
Public Property Let BarBackColor(ByVal colorCode As Long)
     m_barFrame.BackColor = colorCode
     Refresh
End Property

'===============================================================================
'Enables/disables the X button on the progress form
'===============================================================================
Public Property Get AllowCancel() As Boolean
    AllowCancel = m_allowCancel
End Property
Public Property Let AllowCancel(ByVal canCancel As Boolean)
     m_allowCancel = canCancel
End Property

'===============================================================================
'Can be modal or modeless
'===============================================================================
Public Property Get ShowType() As FormShowConstants
    ShowType = m_showType
End Property
Public Property Let ShowType(ByVal formShowType As FormShowConstants)
    If formShowType <> vbModal Then formShowType = vbModeless 'Restrict value
    m_showType = formShowType
End Property

'===============================================================================
'Enables/disables the time labels
'===============================================================================
Public Property Get ShowTime() As Boolean
    ShowTime = m_showTime
End Property
Public Property Let ShowTime(ByVal displayTime As Boolean)
    If m_showTime Xor displayTime Then
        Dim adjustment As Single: adjustment = CastToControl(m_remaining).Height
        If Not displayTime Then adjustment = -adjustment
        m_form.Height = m_form.Height + adjustment
    End If
    m_showTime = displayTime
    m_elapsed.Visible = m_showTime
    m_remaining.Visible = m_showTime
    Refresh
End Property

'===============================================================================
'Indicates if the X button on the progress form was pressed
'===============================================================================
Public Property Get WasCancelled() As Boolean
    WasCancelled = m_cancelled
End Property

'===============================================================================
'Vertical position
'===============================================================================
Public Property Get Top() As Single
    Top = m_form.Top
End Property
Public Property Let Top(ByVal topValue As Single)
    m_form.Top = topValue
    m_isAutoCentered = False
End Property

'===============================================================================
'Horizontal position
'===============================================================================
Public Property Get Left() As Single
    Left = m_form.Left
End Property
Public Property Let Left(ByVal leftValue As Single)
    m_form.Left = leftValue
    m_isAutoCentered = False
End Property

'*******************************************************************************
'Utility for positioning
'*******************************************************************************
Public Sub CenterOnApplication()
    If Application.WindowState = xlMinimized Then Exit Sub
    If ThisWorkbook.Windows.Count > 0 Then
        If ThisWorkbook.Windows(1).WindowState = xlMinimized Then Exit Sub
    End If
    '
    Dim leftPosition As Single
    Dim topPosition As Single
    '
    With Application
        leftPosition = .Left + (.Width - m_form.Width) / 2
        If leftPosition < .Left Then leftPosition = .Left
        '
        topPosition = .Top + (.Height - m_form.Height) / 2
        If topPosition < .Top Then topPosition = .Top
    End With
    With m_form
        .StartUpPosition = 0
        .Left = leftPosition
        .Top = topPosition
    End With
    m_isAutoCentered = True
End Sub

'===============================================================================
'Size
'===============================================================================
Public Property Get Height() As Single
    Height = m_form.Height
End Property
Public Property Get Width() As Single
    Width = m_form.Width
End Property
Public Property Let Width(ByVal widthValue As Single)
    Const minWidth As Single = 180
    Const maxWidth As Single = 450
    Dim finalWidth As Single: finalWidth = widthValue
    Dim offsetValue As Single
    '
    If finalWidth < minWidth Then finalWidth = minWidth
    If finalWidth > maxWidth Then finalWidth = maxWidth
    If finalWidth = m_form.Width Then Exit Property
    offsetValue = finalWidth - m_form.Width
    '
    m_form.Width = finalWidth
    m_barFrame.Width = m_barFrame.Width + offsetValue
    m_percent.Left = m_percent.Left + offsetValue
    If m_isAutoCentered Then m_form.Left = m_form.Left - offsetValue / 2
End Property

'*******************************************************************************
'Self-instance
'*******************************************************************************
Public Function Self() As ProgressBar
    Set Self = Me
End Function

'===============================================================================
'Current progress value
'===============================================================================
Public Property Get Value() As Double
    Value = m_currentValue
End Property
Public Property Let Value(ByVal percentValue As Double)
    If percentValue < 0 Or percentValue > 1 Then Exit Property
    m_currentValue = percentValue
    '
    m_bar.Width = m_currentValue * m_barFrame.InsideWidth
    m_percent.Caption = "Done: " & Format$(m_currentValue, "0%")
    '
    Refresh
End Property

'*******************************************************************************
'Updates the time and allows for events so that the form is updated visually
'*******************************************************************************
Private Sub Refresh()
    If m_isRunning Then
        UpdateTime
        DoEvents
    End If
End Sub
Private Sub UpdateTime()
    If Not m_showTime Then Exit Sub
    If m_currentValue = 0 Then
        m_elapsed.Caption = vbNullString
        m_remaining.Caption = vbNullString
        Exit Sub
    End If
    '
    Dim elapsedTime As Date
    Dim remainingTime As Date
    '
    elapsedTime = VBA.Now - m_startTime
    remainingTime = elapsedTime / m_currentValue * (1 - m_currentValue)
    '
    UpdateTimeLabel m_elapsed, elapsedTime, "Elapsed time: "
    UpdateTimeLabel m_remaining, remainingTime, "Remaining time: "
End Sub
Private Sub UpdateTimeLabel(ByVal labelControl As MSForms.Label _
                          , ByVal timeValue As Date _
                          , ByVal prefix As String)
    Dim labelValue As String: labelValue = prefix
    If timeValue > 1 Then labelValue = labelValue & Int(CDbl(timeValue)) & "d "
    labelControl.Caption = labelValue & Format$(timeValue, "hh:mm:ss")
End Sub

'*******************************************************************************
'Runs a macro in a standard module
'*******************************************************************************
Public Function RunMacro(ByVal targetBook As Workbook _
                       , ByVal procedure As String _
                       , ParamArray args() As Variant) As Variant
    If m_isRunning Then Exit Function
    Dim methodName As String: methodName = TypeName(Me) & ".RunMacro"
    '
    If procedure = vbNullString Then
        Err.Raise 5, methodName, "Invalid procedure name"
    ElseIf targetBook Is Nothing Then
        Err.Raise 91, methodName, "Workbook not set"
    ElseIf UBound(args) >= LBound(args) Then 'Save arguments for async use
        CloneParamArray args(0), UBound(args) + 1, m_args 'ByRef is preserved!
    Else
        m_args = Array()
    End If
    '
    LetSet(RunMacro) = Run(procedure, targetBook, Nothing)
End Function

'*******************************************************************************
'Runs a method of a given object
'*******************************************************************************
Public Function RunObjMethod(ByVal targetObject As Object _
                           , ByVal procedure As String _
                           , ParamArray args() As Variant) As Variant
    If m_isRunning Then Exit Function
    Dim methodName As String: methodName = TypeName(Me) & ".RunObjMethod"
    '
    If procedure = vbNullString Then
        Err.Raise 5, methodName, "Invalid procedure name"
    ElseIf targetObject Is Nothing Then
        Err.Raise 91, methodName, "Object not set"
    ElseIf UBound(args) >= LBound(args) Then 'Save arguments for async use
        CloneParamArray args(0), UBound(args) + 1, m_args 'ByRef is preserved!
    Else
        m_args = Array()
    End If
    '
    LetSet(RunObjMethod) = Run(procedure, Nothing, targetObject)
End Function

'*******************************************************************************
'Runs a method:
'   - in a standard module if 'targetBook' is provided
'   - on an object if 'targetObject' is provided
'*******************************************************************************
Private Function Run(ByVal procedure As String _
                   , ByVal targetBook As Workbook _
                   , ByVal targetObject As Object) As Variant
    m_procedure = procedure
    Set m_targetBook = targetBook
    Set m_targetObj = targetObject
    '
    m_isRunning = True
    m_cancelled = False
    Value = 0
    '
    m_form.Show m_showType
    If m_showType = vbModeless Then
        RunProcedure
    Else 'vbModal. RunProcedure was already executed via Form_Activate event
    End If
    LetSet(Run) = m_result
End Function

'*******************************************************************************
'Utility - assigns a variant to another variant
'*******************************************************************************
Private Property Let LetSet(ByRef result As Variant, ByRef v As Variant)
    If IsObject(v) Then Set result = v Else result = v
End Property

'*******************************************************************************
'Runs the actual method
'*******************************************************************************
Private Sub RunProcedure()
    m_startTime = Now()
    '
    Dim cKey As XlEnableCancelKey: cKey = Application.EnableCancelKey
    If cKey <> xlDisabled Then Application.EnableCancelKey = xlDisabled
    '
    On Error GoTo Clean
    If m_targetObj Is Nothing Then
        RunOnBook
    Else
        #If Mac Then
            RunOnObject m_args
        #Else
            LetSet(m_result) = rtcCallByName(targetObj:=m_targetObj _
                                           , procNamePtr:=StrPtr(m_procedure) _
                                           , vCallType:=VbMethod _
                                           , args:=m_args)
        #End If
    End If
Clean:
    If cKey <> xlDisabled Then Application.EnableCancelKey = cKey
    m_isRunning = False
    If Err.Number = 0 Then
        TryHideForm 'Protection if multiple progress bars are displayed
    Else
        m_form.Hide
        Err.Raise Err.Number, TypeName(Me) & ".RunProcedure"
    End If
End Sub
Private Sub RunOnBook(Optional ByVal Missing As Variant)
    Const maxRunArgs As Long = 30
    Dim argsCount As Long: argsCount = UBound(m_args) + 1
    Dim i As Long
    '
    ReDim Preserve m_args(0 To maxRunArgs - 1)
    For i = argsCount To UBound(m_args)
        m_args(i) = Missing
    Next i
    '
    LetSet(m_result) = Application.Run(FullProcedureName() _
        , m_args(0), m_args(1), m_args(2), m_args(3), m_args(4) _
        , m_args(5), m_args(6), m_args(7), m_args(8), m_args(9) _
        , m_args(10), m_args(11), m_args(12), m_args(13), m_args(14) _
        , m_args(15), m_args(16), m_args(17), m_args(18), m_args(19) _
        , m_args(20), m_args(21), m_args(22), m_args(23), m_args(24) _
        , m_args(25), m_args(26), m_args(27), m_args(28), m_args(29))
End Sub
Private Function FullProcedureName() As String
    If m_targetBook Is ThisWorkbook Then 'No need to qualify book. Default is always ThisWB
        FullProcedureName = m_procedure
    Else
        FullProcedureName = "'" & Replace(m_targetBook.Name, "'", "''") & "'!" & m_procedure
    End If
End Function
#If Mac Then
Private Sub RunOnObject(ByRef args() As Variant)
    Dim o As Object: Set o = m_targetObj
    Dim p As String: p = m_procedure
    Dim v As VbCallType: v = VbMethod
    '
    Select Case UBound(args) - LBound(args) + 1
    Case 0: LetSet(m_result) = CallByName(o, p, v)
    Case 1: LetSet(m_result) = CallByName(o, p, v, args(0))
    Case 2: LetSet(m_result) = CallByName(o, p, v, args(0), args(1))
    Case 3: LetSet(m_result) = CallByName(o, p, v, args(0), args(1), args(2))
    Case 4: LetSet(m_result) = CallByName(o, p, v, args(0), args(1), args(2), args(3))
    Case 5: LetSet(m_result) = CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4))
    Case 6: LetSet(m_result) = CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5))
    Case 7: LetSet(m_result) = CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6))
    Case 8: LetSet(m_result) = CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7))
    Case 9: LetSet(m_result) = CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8))
    Case 10: LetSet(m_result) = CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9))
    Case 11: LetSet(m_result) = CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10))
    Case 12: LetSet(m_result) = CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11))
    Case 13: LetSet(m_result) = CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12))
    Case 14: LetSet(m_result) = CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13))
    Case 15: LetSet(m_result) = CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14))
    Case 16: LetSet(m_result) = CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15))
    Case 17: LetSet(m_result) = CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16))
    Case 18: LetSet(m_result) = CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17))
    Case 19: LetSet(m_result) = CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18))
    Case 20: LetSet(m_result) = CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19))
    Case 21: LetSet(m_result) = CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20))
    Case 22: LetSet(m_result) = CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21))
    Case 23: LetSet(m_result) = CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22))
    Case 24: LetSet(m_result) = CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23))
    Case 25: LetSet(m_result) = CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24))
    Case 26: LetSet(m_result) = CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24), args(25))
    Case 27: LetSet(m_result) = CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24), args(25), args(26))
    Case 28: LetSet(m_result) = CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24), args(25), args(26), args(27))
    Case 29: LetSet(m_result) = CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24), args(25), args(26), args(27), args(28))
    Case Else: LetSet(m_result) = CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24), args(25), args(26), args(27), args(28), args(29))
    End Select
End Sub
#End If
