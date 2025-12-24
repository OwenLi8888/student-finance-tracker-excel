VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserMonthandDate 
   Caption         =   "UserForm1"
   ClientHeight    =   4774
   ClientLeft      =   -238
   ClientTop       =   -1022
   ClientWidth     =   7945
   OleObjectBlob   =   "UserMonthandDate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserMonthandDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SubmitButton_Click()
    
    'the goal of this code is to grab the data from expenses and income and seperates and displays it in the other
    'assign variables
    
    Dim wsData As Worksheet
    Dim wsHome As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim nextIncRow As Long
    Dim nextExpRow As Long
    Dim transType As String
    Dim description As String
    Dim amount As Double
    Dim transDate As Date
    Dim selectedMonth As Integer
    Dim selectedYear As Integer
    Dim incomeStart As Long
    Dim expenseStart As Long

    'Set references
    Set wsData = ThisWorkbook.Worksheets("Expenses&Incomes")
    Set wsHome = ThisWorkbook.Worksheets("Home Page")

    'check if user inputted
    If Trim(MonthSet.Value) = "" Or Trim(YearSet.Value) = "" Then
        MsgBox "Please enter both month and year.", vbExclamation
        Exit Sub
    End If

    'convert user input into proper month and year
    On Error Resume Next
    selectedMonth = Month(DateValue("1 " & MonthSet.Value & " " & YearSet.Value))
    selectedYear = CInt(YearSet.Value)
    On Error GoTo 0

    If selectedMonth = 0 Or selectedYear = 0 Then
        MsgBox "Invalid month or year. Example: Month = Novemeber, Year = 2025.", vbCritical
        Exit Sub
    End If

    'Set starting rows
    incomeStart = 12
    expenseStart = 12
    nextIncRow = incomeStart
    nextExpRow = expenseStart

    'Clear everythign below the headers of the income and expense table
    wsHome.Range("B" & incomeStart & ":C1000").ClearContents
    wsHome.Range("D" & expenseStart & ":E1000").ClearContents

    'Find last row of data on source sheet
    lastRow = wsData.Cells(wsData.Rows.Count, "B").End(xlUp).Row

    'Loop through the data
    For i = 2 To lastRow
        If IsDate(wsData.Cells(i, "A").Value) Then
            transDate = wsData.Cells(i, "A").Value
            If Month(transDate) = selectedMonth And Year(transDate) = selectedYear Then
                description = wsData.Cells(i, "B").Value
                transType = LCase(Trim(wsData.Cells(i, "C").Value))
                amount = wsData.Cells(i, "D").Value

                If transType = "income" Then
                    wsHome.Cells(nextIncRow, "B").Value = description
                    wsHome.Cells(nextIncRow, "C").Value = amount
                    nextIncRow = nextIncRow + 1
                Else
                    wsHome.Cells(nextExpRow, "D").Value = description
                    wsHome.Cells(nextExpRow, "E").Value = Abs(amount)
                    nextExpRow = nextExpRow + 1
                End If
            End If
        End If
    Next i

    'Totals on row 10
    wsHome.Range("C10").Formula = "=SUM(C" & incomeStart & ":C" & nextIncRow - 1 & ")"
    wsHome.Range("E10").Formula = "=SUM(E" & expenseStart & ":E" & nextExpRow - 1 & ")"

    MsgBox "Financial data generated for " & MonthSet.Value & " " & YearSet.Value & "."
    
    'Write the user month amd year to Home Page cells
    wsHome.Range("C2").Value = MonthSet.Value
    wsHome.Range("E2").Value = YearSet.Value

    'Close form
    Unload Me
End Sub

