Attribute VB_Name = "Module1"
Function AddNumbersInCells() As Double
    Dim ws As Worksheet
    Dim num1 As Double
    Dim num2 As Double
    Dim result As Double

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Calculator")

    ' Read numbers from cells A1 and B1
    num1 = ws.Range("A2").Value
    num2 = ws.Range("B2").Value

    ' Calculate the sum
    AddNumbersInCells = num1 + num2

    ' Write the result in cell C1
    ws.Range("D2").Value = AddNumbersInCells
End Function

Function MulNumbersInCells() As Double
    Dim ws As Worksheet
    Dim num1 As Double
    Dim num2 As Double
    Dim result As Double

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Calculator")

    ' Read numbers from cells A1 and B1
    num1 = ws.Range("A3").Value
    num2 = ws.Range("B3").Value

    ' Calculate the sum
    MulNumbersInCells = num1 * num2

    ' Write the result in cell C1
    ws.Range("D3").Value = MulNumbersInCells
End Function

Sub DivideNumbersInCells()
    Dim ws As Worksheet
    Dim num1 As Double
    Dim num2 As Double
    Dim result As Double

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Calculator")

    ' Read numbers from cells A1 and B1
    num1 = ws.Range("A4").Value
    num2 = ws.Range("B4").Value

    ' Calculate the sum
    result = num1 / num2

    ' Write the result in cell C1
    ws.Range("D4").Value = result
End Sub

Sub SubNumbersInCells()
    Dim ws As Worksheet
    Dim num1 As Double
    Dim num2 As Double
    Dim result As Double

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Calculator")

    ' Read numbers from cells A1 and B1
    num1 = ws.Range("A5").Value
    num2 = ws.Range("B5").Value

    ' Calculate the sum
    result = num1 - num2

    ' Write the result in cell C1
    ws.Range("D5").Value = result
End Sub

Sub AllTogether()
    Dim ws As Worksheet
    Dim sumResult As Double
    Dim mulResult As Double
    Dim result As Double

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Calculator")

    ' Read numbers from cells A1 and B1
    sumResult = AddNumbersInCells
    mulResult = MulNumbersInCells

    ' Calculate the sum
    result = sumResult + mulResult

    ' Write the result in cell C1
    ws.Range("D11").Value = result
End Sub


