Attribute VB_Name = "AccountingSolver"
'==========================================================================
'  Solver - accounting problem solver for Excel (Windows)
'
'  SolveStepByStep : fills the answers into the open workbook ONE BOX AT A
'                    TIME. Each box is selected and explained; click OK for
'                    the next box (Cancel fills the rest at once).
'  SolveAutoPlay   : same, but plays by itself (one box every ~0.7 s).
'  SolveAllAtOnce  : fills every box immediately.
'
'  Setup (once):  see README.md, "Use it inside Excel".
'  Works offline. Needs the Solver folder (Solver.exe or Python + run_solver.py).
'==========================================================================
Option Explicit

' >>> Change this to wherever you put the Solver folder <<<
Public Const SOLVER_FOLDER As String = "C:\Solver"

Private Const AUTO_DELAY_SECONDS As Double = 0.7

Public Sub SolveStepByStep()
    RunSolver 1
End Sub

Public Sub SolveAutoPlay()
    RunSolver 2
End Sub

Public Sub SolveAllAtOnce()
    RunSolver 0
End Sub

Private Function Launcher() As String
    Dim folder As String
    folder = SOLVER_FOLDER
    If Right(folder, 1) <> "\" Then folder = folder & "\"
    If Dir(folder & "Solver.exe") <> "" Then
        Launcher = """" & folder & "Solver.exe"""
    ElseIf Dir(folder & "run_solver.py") <> "" Then
        Launcher = "pythonw """ & folder & "run_solver.py"""
    Else
        Launcher = ""
    End If
End Function

Private Function ReadUtf8(path As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    ReadUtf8 = st.ReadText
    st.Close
End Function

Private Sub RunSolver(mode As Integer)
    Dim wb As Workbook, solved As Workbook
    Dim exe As String, tmp As String, outFile As String, stepsFile As String, cmd As String
    Dim lines() As String, parts() As String
    Dim i As Long, total As Long, n As Long, answer As VbMsgBoxResult

    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    If wb.Path = "" Then
        MsgBox "Save this workbook first (as .xlsx or .xlsm), then run Solver again.", _
               vbExclamation, "Solver"
        Exit Sub
    End If
    If LCase(Right(wb.Name, 4)) = ".xls" Then
        MsgBox "Solver needs an .xlsx or .xlsm file. Use File > Save As > Excel Workbook.", _
               vbExclamation, "Solver"
        Exit Sub
    End If
    exe = Launcher()
    If exe = "" Then
        MsgBox "Solver was not found in " & SOLVER_FOLDER & "." & vbCrLf & _
               "Open the macro (Alt+F11) and change SOLVER_FOLDER to your Solver folder.", _
               vbExclamation, "Solver"
        Exit Sub
    End If

    wb.Save
    tmp = Environ("TEMP") & "\"
    outFile = tmp & "solver_result" & IIf(LCase(Right(wb.Name, 5)) = ".xlsm", ".xlsm", ".xlsx")
    stepsFile = tmp & "solver_steps.txt"
    On Error Resume Next
    Kill outFile
    Kill stepsFile
    On Error GoTo 0

    Application.StatusBar = "Solver is working..."
    cmd = exe & " """ & wb.FullName & """ --out """ & outFile & """ --steps-file """ & _
          stepsFile & """"
    CreateObject("WScript.Shell").Run cmd, 0, True
    Application.StatusBar = False

    If Dir(stepsFile) = "" Then
        MsgBox "Solver did not run. Check that Python or Solver.exe is in " & SOLVER_FOLDER, _
               vbCritical, "Solver"
        Exit Sub
    End If
    lines = Split(ReadUtf8(stepsFile), vbCrLf)
    If UBound(lines) >= 0 Then
        parts = Split(lines(0), vbTab)
        If parts(0) = "error" And UBound(parts) >= 3 Then
            MsgBox "Solver could not read this workbook:" & vbCrLf & parts(3), _
                   vbCritical, "Solver"
            Exit Sub
        End If
    End If
    For i = 0 To UBound(lines)
        If Len(lines(i)) > 0 Then total = total + 1
    Next i
    If total = 0 Then
        MsgBox "Solver did not find anything to fill. Each number needs a label next to it, " & _
               "and answer boxes should be empty or contain ?", vbInformation, "Solver"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Set solved = Workbooks.Open(outFile, UpdateLinks:=0, ReadOnly:=True)
    solved.Windows(1).Visible = False
    wb.Activate
    ' Create every new sheet up front so formulas never point at a missing sheet.
    For i = 0 To UBound(lines)
        parts = Split(lines(i), vbTab)
        If UBound(parts) >= 1 Then
            If parts(0) = "sheet" Then AddSheet wb, parts(1)
        End If
    Next i
    Application.ScreenUpdating = True

    n = 0
    For i = 0 To UBound(lines)
        parts = Split(lines(i), vbTab)
        If UBound(parts) >= 3 Then
            n = n + 1
            If parts(0) = "sheet" Then
                Application.ScreenUpdating = False
                CopySheet solved.Worksheets(parts(1)), wb.Worksheets(parts(1)), False
                Application.ScreenUpdating = True
                wb.Worksheets(parts(1)).Activate
            ElseIf parts(0) = "fill" Then
                FillBox solved.Worksheets(parts(1)).Range(parts(2)), _
                        wb.Worksheets(parts(1)).Range(parts(2))
            End If
            Application.StatusBar = "Solver box " & n & " of " & total & ": " & parts(3)
            If mode = 1 Then
                answer = MsgBox(parts(3) & IIf(parts(0) = "fill", vbCrLf & vbCrLf & _
                                parts(1) & "!" & parts(2) & "   " & _
                                wb.Worksheets(parts(1)).Range(parts(2)).Formula, ""), _
                                vbOKCancel + vbInformation, _
                                "Solver - box " & n & " of " & total & _
                                "   (OK = next box, Cancel = fill the rest)")
                If answer = vbCancel Then mode = 0: Application.ScreenUpdating = False
            ElseIf mode = 2 Then
                Pause AUTO_DELAY_SECONDS
            End If
        End If
    Next i

    ' The notes sheet lists every answer and any warnings.
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets("Solver Notes").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    AddSheet wb, "Solver Notes"
    CopySheet solved.Worksheets("Solver Notes"), wb.Worksheets("Solver Notes")
    solved.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Application.CalculateFull
    MsgBox "Done - " & total & " step(s). Green boxes were filled by Solver." & vbCrLf & _
           "Click any green box to see its formula. See the 'Solver Notes' sheet for details.", _
           vbInformation, "Solver"
End Sub

Private Sub AddSheet(wb As Workbook, sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
    End If
End Sub

Private Sub FillBox(src As Range, dst As Range)
    dst.Worksheet.Activate
    dst.Select
    If src.HasFormula Then
        dst.Formula = src.Formula
    Else
        dst.Value = src.Value
    End If
    dst.NumberFormat = src.NumberFormat
    If src.Interior.ColorIndex <> xlColorIndexNone Then dst.Interior.Color = src.Interior.Color
End Sub

Private Sub CopySheet(src As Worksheet, dst As Worksheet, Optional withFormulas As Boolean = True)
    ' Labels and formatting; formulas only when withFormulas (otherwise they are
    ' filled afterwards, one box at a time, by their own steps).
    Dim c As Range, d As Range, col As Long
    For Each c In src.UsedRange.Cells
        If Not IsEmpty(c.Value) Or c.HasFormula Then
            Set d = dst.Range(c.Address)
            If c.HasFormula Then
                If withFormulas Then d.Formula = c.Formula
            Else
                d.Value = c.Value
            End If
            d.NumberFormat = c.NumberFormat
            d.Font.Bold = c.Font.Bold
            d.Font.Italic = c.Font.Italic
            d.Font.Size = c.Font.Size
            If c.Interior.ColorIndex <> xlColorIndexNone Then d.Interior.Color = c.Interior.Color
            d.Borders(xlEdgeTop).LineStyle = c.Borders(xlEdgeTop).LineStyle
            d.Borders(xlEdgeBottom).LineStyle = c.Borders(xlEdgeBottom).LineStyle
        End If
    Next c
    For col = 1 To src.UsedRange.Columns.Count + src.UsedRange.Column
        dst.Columns(col).ColumnWidth = src.Columns(col).ColumnWidth
    Next col
End Sub

Private Sub Pause(seconds As Double)
    Dim t As Double
    t = Timer
    Do While Timer < t + seconds
        DoEvents
    Loop
End Sub
