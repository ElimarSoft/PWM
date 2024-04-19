Attribute VB_Name = "Calculation"
Option Explicit
Dim ws3 As Worksheet
Const steps As Long = 36000
Const index As Long = 360 '90
'Const index As Long = 180 '45
'Const index As Long = 120 '30
'Const index As Long = 60 '15

Dim ptr1 As Double
Dim modIndex As Double
Dim ramp As Boolean
Public Sub Main()

    modIndex = 0.34
    Update (modIndex)

End Sub

Public Sub Update(modIndex As Double)

Set ws3 = Worksheets("Graphics"): ws3.Cells.Clear
Dim n As Long
ptr1 = 0

    Dim values(steps, 6) As Double
    ws3.Range("L68") = modIndex * 1000

    ramp = True
    For n = 1 To steps
        
        Dim Angle As Double: Angle = (n - 1) * 2 * 3.1416 / steps
        Dim Angle2 As Double: Angle2 = (n) * 2 * 3.1416 / steps
        
        Dim SinVal As Double: SinVal = Sin(Angle)
        
        If (ptr1 >= (steps / index)) Then ramp = False
        If (ptr1 <= -(steps / index)) Then ramp = True
                 
        If ramp Then
            ptr1 = ptr1 + 1
        Else
            ptr1 = ptr1 - 1
        End If
    
    
        Dim rampval As Double: rampval = ptr1 / (steps / index)
    
        Dim outval As Integer
        If (SinVal * modIndex > rampval) Then
            outval = 1
        Else
            outval = -1
        End If
        
        values(n - 1, 0) = n - 1
        values(n - 1, 1) = SinVal * modIndex
        values(n - 1, 2) = Sin(Angle * 5) * modIndex
        values(n - 1, 3) = Sin(Angle * 7) * modIndex
        values(n - 1, 4) = rampval 'Ramp signal
        values(n - 1, 5) = outval 'Binary signal
    
        Const maxVals As Integer = 60
    
        'Calculate harmonic values
        Dim m As Integer
        Dim outSin(maxVals) As Double
        Dim outCos(maxVals) As Double
        
        For m = 1 To maxVals Step 2
            outSin(m) = outSin(m) + Sin(Angle * m) * outval
            outCos(m) = outCos(m) + Cos(Angle * m) * outval
        Next m
    
    Next n

        Dim Har1 As Double: Har1 = Sqr(outSin(1) ^ 2 + outCos(1) ^ 2) / steps * 2
        Debug.Print modIndex, Har1

        Range(ws3.Cells(1, 1), ws3.Cells(steps, 6)) = values
        
        'Plot Harmonics
        
        Dim ptr2 As Integer: ptr2 = 1

        For m = 1 To maxVals Step 2

            If Not (m Mod 3) = 0 Then
                Dim outAbs As Double: outAbs = Sqr(outSin(m) ^ 2 + outCos(m) ^ 2)
                ws3.Cells(ptr2, 8) = m
                ws3.Cells(ptr2, 9) = Format(outAbs / steps, "#0.000000")
                ptr2 = ptr2 + 1
            End If

        Next m

End Sub
