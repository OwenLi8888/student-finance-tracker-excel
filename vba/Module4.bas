Attribute VB_Name = "Module4"
Sub Goals_Output()

    Dim wsGoals As Worksheet
    Dim wsHome As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim selectedDate As Date
    Dim goalDate As Date
    Dim goalName As String
    Dim targetAmount As Double
    Dim amountAllocated As Double
    Dim progressPercent As Double
    Dim outputRow As Long

    ' sheet references
    Set wsGoals = ThisWorkbook.Worksheets("Goals")
    Set wsHome = ThisWorkbook.Worksheets("Home Page")

    ' selected month/year from L1 in the goals sheet
    selectedDate = wsGoals.Range("L1").Value

    ' Clear previous Home Page results (Q10:S1000)but not R
    wsHome.Range("Q10:Q1000").ClearContents
    wsHome.Range("S10:S1000").ClearContents


    ' start writing on Home Page at row 10
    outputRow = 10

    ' last row with data on Goals sheet
    lastRow = wsGoals.Cells(wsGoals.Rows.Count, "A").End(xlUp).Row

    ' loop through *all* goals
    For i = 2 To lastRow
        
        goalDate = wsGoals.Cells(i, "A").Value
        
        ' match by month + year only
        If Month(goalDate) = Month(selectedDate) And Year(goalDate) = Year(selectedDate) Then
            
            ' pull fields
            goalName = wsGoals.Cells(i, "B").Value
            targetAmount = wsGoals.Cells(i, "D").Value
            amountAllocated = wsGoals.Cells(i, "E").Value

            ' calculate progress safely
            If targetAmount > 0 Then
                progressPercent = amountAllocated / targetAmount
            Else
                progressPercent = 0
            End If

            ' output to Home Page
            wsHome.Cells(outputRow, "Q").Value = goalName
            wsHome.Cells(outputRow, "S").Value = progressPercent
            wsHome.Cells(outputRow, "S").NumberFormat = "0.0%"

            ' move down a row for next result
            outputRow = outputRow + 1
        End If

    Next i

    MsgBox "All goals for the selected month loaded into Home Page (Q10:S)."

End Sub

