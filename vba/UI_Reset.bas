Attribute VB_Name = "Module1"
Sub ClearExpensesIncomes()

    'Select the range
    Range("A2:D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    
    'Clear the selected range
    Selection.ClearContents
    
End Sub

Sub ClearOutput()

    'Clear start date
    Range("A2").Select
    Selection.ClearContents
    
    'Clear end date
    Range("A4").Select
    Selection.ClearContents
    
    'Clear output data
    Range("E2:H2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

End Sub

