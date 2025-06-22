Attribute VB_Name = "Module1"
Option Explicit

' Block size (X x Y cells)
Const blockLength = 20
Const blockHeight = 50
Const initTemp = 25
Const fluidTemp = 200
Const totalSteps = 100

Dim temperature(blockLength, blockHeight) As Double
Dim holeY() As Integer

Sub InitializeBlock()
    Dim i As Integer, j As Integer
    ReDim holeY(1 To 3)
    holeY(1) = 10
    holeY(2) = 25
    holeY(3) = 40

    ' Set initial temperature
    For i = 1 To blockLength
        For j = 1 To blockHeight
            temperature(i, j) = initTemp
            Cells(j + 1, i + 1).Value = ""
        Next j
    Next i

    ' Set holes with fluid temp
    For i = 1 To blockLength
        For j = 1 To UBound(holeY)
            temperature(i, holeY(j)) = fluidTemp
        Next j
    Next i

    ApplyColorMap
End Sub

Sub SimulateHeatTransfer()
    Dim stepNum As Integer
    Dim i As Integer, j As Integer
    Dim tempNew(blockLength, blockHeight) As Double
    Dim k As Double: k = 0.15 ' Thermal diffusivity constant

    For stepNum = 1 To totalSteps
        For i = 2 To blockLength - 1
            For j = 2 To blockHeight - 1
                If IsHole(i, j) Then
                    tempNew(i, j) = fluidTemp
                Else
                    tempNew(i, j) = temperature(i, j) + _
                        k * (temperature(i + 1, j) + temperature(i - 1, j) + _
                        temperature(i, j + 1) + temperature(i, j - 1) - _
                        4 * temperature(i, j))
                End If
            Next j
        Next i

        For i = 1 To blockLength
            For j = 1 To blockHeight
                temperature(i, j) = tempNew(i, j)
            Next j
        Next i

        ApplyColorMap
        DoEvents
       ' Application.Wait (Now + TimeValue("0:00:01"))
    Next stepNum
End Sub

Function IsHole(i As Integer, j As Integer) As Boolean
    Dim idx As Integer
    For idx = 1 To UBound(holeY)
        If j = holeY(idx) Then
            IsHole = True
            Exit Function
        End If
    Next idx
    IsHole = False
End Function

Sub ApplyColorMap()
    Dim i As Integer, j As Integer
    Dim temp As Double
    Dim colorValue As Long

    For i = 1 To blockLength
        For j = 1 To blockHeight
            temp = temperature(i, j)
            colorValue = TemperatureToColor(temp)
          '  With Cells(j + 1, i + 1).Interior
          With Cells(j + 3, i + 3).Interior
                .Color = colorValue
            End With
        Next j
    Next i
End Sub

Function TemperatureToColor(temp As Double) As Long
    Dim r As Integer, g As Integer, b As Integer
    Dim ratio As Double
    ratio = (temp - initTemp) / (fluidTemp - initTemp)
    'If ratio > 1 Then ratio = 1
    If ratio > 1 Then ratio = 10
    If ratio < 0 Then ratio = 0

    r = 255 * ratio
    g = 0
    b = 255 * (1 - ratio)
    TemperatureToColor = RGB(r, g, b)
End Function

Sub StartSimulation()
    Call InitializeBlock
    Call SimulateHeatTransfer
End Sub


