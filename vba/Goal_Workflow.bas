Attribute VB_Name = "Module3"
Sub ClearGoals()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Goals")
    
    'Clear all rows below the header (row 1)
    ws.Rows("2:" & ws.Rows.Count).ClearContents
    
    MsgBox "All goals have been cleared.", vbInformation
End Sub
Sub Generate_Financial_Data()
    UserMonthandDate.Show
End Sub

