VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OutputForm 
   Caption         =   "Output"
   ClientHeight    =   4172
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8477.001
   OleObjectBlob   =   "OutputForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "OutputForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Global declarations

'Declare an integer to track reading row
Dim intReadRow As Integer

'Declare an integer to track writing row
Dim intWriteRow As Integer

Private Sub SubmitBtn_Click()
    
    'Set workbook and sheet
    Set WB = ThisWorkbook
    Set ws = WB.Worksheets("Output")

    'initialize row counters
    intReadRow = 2
    intWriteRow = 2

    'Test value of start dates textBoxes
    If (txtDay1.Value <> "" And txtMonth1.Value <> "" And txtYear1.Value <> "") Then
    
        'Test value of end dates textBoxes
        If (txtDay2.Value <> "" And txtMonth2.Value <> "" And txtYear2.Value <> "") Then
 
            'Write date into start cell
            ws.Cells(2, "A") = txtYear1.Value + "-" + txtMonth1.Value + "-" + txtDay1.Value
        
            'Format cell so Excel recognizes a date
            ws.Cells(2, "A").NumberFormat = "yyyy-mm-dd;@"
            
            'Write date into end cell
            ws.Cells(4, "A") = txtYear2.Value + "-" + txtMonth2.Value + "-" + txtDay2.Value
        
            'Format cell so Excel recognizes a date
            ws.Cells(4, "A").NumberFormat = "yyyy-mm-dd;@"
            
            'Loop down established expenses & incomes searching for matches
            Do While (Worksheets("Expenses&Incomes").Cells(intReadRow, "A") <> "")
    
                'Compare given dates to dates in spreadsheet
                If (Worksheets("Expenses&Incomes").Cells(intReadRow, "A") >= ws.Cells(2, "A") And Worksheets("Expenses&Incomes").Cells(intReadRow, "A") <= ws.Cells(4, "A")) Then
                
                    'Print to the output sheet
                    ws.Cells(intWriteRow, "E") = Worksheets("Expenses&Incomes").Cells(intReadRow, "A")
                    ws.Cells(intWriteRow, "F") = Worksheets("Expenses&Incomes").Cells(intReadRow, "B")
                    ws.Cells(intWriteRow, "G") = Worksheets("Expenses&Incomes").Cells(intReadRow, "C")
                    ws.Cells(intWriteRow, "H") = Worksheets("Expenses&Incomes").Cells(intReadRow, "D")
                    
                    'Format cell so Excel recognizes a date
                    ws.Cells(intWriteRow, "E").NumberFormat = "yyyy-mm-dd;@"
                    
                    'Increment write row
                    intWriteRow = intWriteRow + 1
                
                End If
            
                'Increment read row counter
                intReadRow = intReadRow + 1
    
            Loop
        
        Else

            'Give error message for no end date
            MsgBox ("Please Enter a valid end date")
    
        End If
        
    Else

        'Give error message for no start date
        MsgBox ("Please Enter a valid start date")
    
    End If
    
End Sub
