Attribute VB_Name = "Module5"
Sub AdviceOutput()

    Dim ws As Worksheet
    Dim actualExp As Double
    Dim recommendedExp As Double

    Set ws = ThisWorkbook.Worksheets("Home Page")

    ' get the actual and recommended expenses
    actualExp = ws.Range("E10").Value
    recommendedExp = ws.Range("W11").Value

    ' check for valid entries
    If Not IsNumeric(actualExp) Or Not IsNumeric(recommendedExp) Then
        MsgBox "E10 or W11 must contain valid numbers."
        Exit Sub
    End If

    ' clears
    ws.Range("W14").Value = ""
    ws.Range("W18").Value = ""

    ' compares and gives the correct advice
    If actualExp > recommendedExp Then
        ws.Range("W14").Value = "Your expenses are above the recommended level, tighten spending."
        ws.Range("W18").Value = "Every dollar you save now multiplies your future options."
    ElseIf actualExp < recommendedExp Then
        ws.Range("W14").Value = "Great job, your expenses are below the recommended level."
        ws.Range("W18").Value = "Strong discipline compounds. Keep going."
    Else
        ws.Range("W14").Value = "Your expenses match the recommended target."
        ws.Range("W18").Value = "Consistency is how people get financially ahead."
    End If

End Sub

