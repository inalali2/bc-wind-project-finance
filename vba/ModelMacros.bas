Attribute VB_Name = "ModelMacros"
'==============================================================================
' Coastal Ridge Wind project finance model: macros
' Author: Inal Ali
'
' Import: save the workbook as .xlsm, press Alt+F11, File > Import File,
'         pick this .bas file. Run macros with Alt+F8.
'
' Named ranges this module relies on (all created by the workbook):
'   TaxPaste, TaxCalc, TaxGap      tax circularity rows on Model
'   EquityCF, DSCRRow, EquityRate  outputs used for results
'   Flex_*                         flex cells on Inputs (column H)
'   EPAPrice, CF                   live input cells
'   TornadoTop, TwoWayCorner, BreakevenOut   anchors on Sensitivity
'==============================================================================
Option Explicit

Private Const MAX_ITER As Long = 100
Private Const TOL As Double = 0.000001
Private Const FLEX_STEP As Double = 0.1

' ---------------------------------------------------------------------------
' Core: break the tax / interest / debt size circularity by copy and paste.
' Returns the number of iterations used, or -1 if it failed to converge.
' ---------------------------------------------------------------------------
Private Function ConvergeTax() As Long
    Dim i As Long
    Dim gap As Double
    Application.Calculate
    For i = 1 To MAX_ITER
        Range("TaxPaste").Value = Range("TaxCalc").Value
        Application.Calculate
        gap = Application.WorksheetFunction.Max(Range("TaxGap"))
        If gap < TOL Then
            ConvergeTax = i
            Exit Function
        End If
    Next i
    ConvergeTax = -1
End Function

Private Function EquityIRR() As Double
    EquityIRR = Application.Evaluate("IRR(EquityCF)")
End Function

Private Function EquityNPV() As Double
    EquityNPV = Application.Evaluate("NPV(EquityRate,EquityCF)")
End Function

Private Function MinDSCR() As Double
    MinDSCR = Application.WorksheetFunction.Min(Range("DSCRRow"))
End Function

Private Sub BeginRun(ByRef calcMode As Long)
    calcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
End Sub

Private Sub EndRun(ByVal calcMode As Long)
    Application.Calculation = calcMode
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

' ---------------------------------------------------------------------------
' Run after changing any input or the scenario selector.
' ---------------------------------------------------------------------------
Public Sub SculptDebt()
    Dim calcMode As Long, n As Long
    BeginRun calcMode
    On Error GoTo Fail
    n = ConvergeTax()
    EndRun calcMode
    If n < 0 Then
        MsgBox "Tax circularity did not converge in " & MAX_ITER & " iterations. Check inputs.", vbExclamation
    Else
        MsgBox "Converged in " & n & " iteration(s)." & vbCrLf & _
               "Equity IRR: " & Format(EquityIRR(), "0.00%") & vbCrLf & _
               "Min DSCR: " & Format(MinDSCR(), "0.00") & "x", vbInformation
    End If
    Exit Sub
Fail:
    EndRun calcMode
    MsgBox "SculptDebt failed: " & Err.Description, vbCritical
End Sub

' ---------------------------------------------------------------------------
' Tornado: flex each driver -10% and +10%, re-sculpt, record IRR and min DSCR.
' Driver list lives on the Sensitivity sheet (column C holds the flex name).
' ---------------------------------------------------------------------------
Public Sub RunSensitivities()
    Dim calcMode As Long, r As Range, flexName As String
    Dim original As Double, baseIRR As Double
    BeginRun calcMode
    On Error GoTo Fail

    If ConvergeTax() < 0 Then Err.Raise vbObjectError + 1, , "Base case did not converge"
    baseIRR = EquityIRR()

    Set r = Range("TornadoTop")
    Do While Len(r.Value) > 0
        flexName = "Flex_" & r.Offset(0, 1).Value
        original = Range(flexName).Value
        Application.StatusBar = "Sensitivity: " & r.Value

        Range(flexName).Value = original - FLEX_STEP
        If ConvergeTax() < 0 Then Err.Raise vbObjectError + 2, , "No convergence on " & r.Value
        r.Offset(0, 2).Value = EquityIRR()
        r.Offset(0, 5).Value = MinDSCR()

        Range(flexName).Value = original + FLEX_STEP
        If ConvergeTax() < 0 Then Err.Raise vbObjectError + 2, , "No convergence on " & r.Value
        r.Offset(0, 4).Value = EquityIRR()
        r.Offset(0, 6).Value = MinDSCR()

        r.Offset(0, 3).Value = baseIRR
        Range(flexName).Value = original
        Set r = r.Offset(1, 0)
    Loop

    ConvergeTax
    EndRun calcMode
    MsgBox "Tornado refreshed.", vbInformation
    Exit Sub
Fail:
    If Len(flexName) > 0 Then Range(flexName).Value = original
    ConvergeTax
    EndRun calcMode
    MsgBox "RunSensitivities failed: " & Err.Description, vbCritical
End Sub

' ---------------------------------------------------------------------------
' Two way grid: EPA price (rows) by capacity factor (columns) -> equity IRR.
' Prices and capacity factors are read from the grid headers, so you can
' edit them on the sheet.
' ---------------------------------------------------------------------------
Public Sub RunTwoWay()
    Dim calcMode As Long, corner As Range
    Dim origPriceFlex As Double, origCFFlex As Double
    Dim basePrice As Double, baseCF As Double
    Dim i As Long, j As Long
    BeginRun calcMode
    On Error GoTo Fail

    Application.Calculate
    origPriceFlex = Range("Flex_EPAPrice").Value
    origCFFlex = Range("Flex_CF").Value
    basePrice = Range("EPAPrice").Value / (1 + origPriceFlex)   ' scenario value before flex
    baseCF = Range("CF").Value / (1 + origCFFlex)

    Set corner = Range("TwoWayCorner")
    i = 1
    Do While Len(corner.Offset(i, 0).Value) > 0
        j = 1
        Do While Len(corner.Offset(0, j).Value) > 0
            Application.StatusBar = "Two way grid: row " & i & ", column " & j
            Range("Flex_EPAPrice").Value = corner.Offset(i, 0).Value / basePrice - 1
            Range("Flex_CF").Value = corner.Offset(0, j).Value / baseCF - 1
            If ConvergeTax() < 0 Then
                corner.Offset(i, j).Value = CVErr(xlErrNA)
            Else
                corner.Offset(i, j).Value = EquityIRR()
            End If
            j = j + 1
        Loop
        i = i + 1
    Loop

    Range("Flex_EPAPrice").Value = origPriceFlex
    Range("Flex_CF").Value = origCFFlex
    ConvergeTax
    EndRun calcMode
    MsgBox "Two way grid refreshed.", vbInformation
    Exit Sub
Fail:
    Range("Flex_EPAPrice").Value = origPriceFlex
    Range("Flex_CF").Value = origCFFlex
    ConvergeTax
    EndRun calcMode
    MsgBox "RunTwoWay failed: " & Err.Description, vbCritical
End Sub

' ---------------------------------------------------------------------------
' Breakeven EPA price: bisection on the price flex until equity NPV at the
' hurdle rate is zero. Goal Seek cannot be used here because every trial
' needs the circularity loop to run.
' ---------------------------------------------------------------------------
Public Sub BreakevenPrice()
    Dim calcMode As Long, origFlex As Double, basePrice As Double
    Dim lo As Double, hi As Double, midPt As Double, k As Long
    BeginRun calcMode
    On Error GoTo Fail

    Application.Calculate
    origFlex = Range("Flex_EPAPrice").Value
    basePrice = Range("EPAPrice").Value / (1 + origFlex)
    lo = -0.5
    hi = 1#
    For k = 1 To 60
        midPt = (lo + hi) / 2
        Range("Flex_EPAPrice").Value = midPt
        If ConvergeTax() < 0 Then Err.Raise vbObjectError + 3, , "No convergence during bisection"
        If EquityNPV() > 0 Then hi = midPt Else lo = midPt
        If (hi - lo) * basePrice < 0.0001 Then Exit For
    Next k

    Range("BreakevenOut").Value = basePrice * (1 + (lo + hi) / 2)
    Range("Flex_EPAPrice").Value = origFlex
    ConvergeTax
    EndRun calcMode
    MsgBox "Breakeven EPA price: $" & Format(Range("BreakevenOut").Value, "0.00") & "/MWh (2024$)", vbInformation
    Exit Sub
Fail:
    Range("Flex_EPAPrice").Value = origFlex
    ConvergeTax
    EndRun calcMode
    MsgBox "BreakevenPrice failed: " & Err.Description, vbCritical
End Sub

' ---------------------------------------------------------------------------
' Convenience: run everything in order.
' ---------------------------------------------------------------------------
Public Sub RunAll()
    SculptDebt
    RunSensitivities
    RunTwoWay
    BreakevenPrice
End Sub
