VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddGoal 
   Caption         =   "UserForm1"
   ClientHeight    =   4858
   ClientLeft      =   -616
   ClientTop       =   -2639
   ClientWidth     =   7560
   OleObjectBlob   =   "AddGoal.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddGoal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Goal_Category_Change()

End Sub

Private Sub Save_Goal_Click()
    Dim WB As Workbook
    Dim ws As Worksheet
    Dim intRow As Long

    Set WB = ThisWorkbook
    Set ws = WB.Worksheets("Goals")

    ' making sure category is not empty
    If Goal_Category.Value = "" Then
        MsgBox "Please enter a category"
        Exit Sub
    End If

    ' making sure category is not empty
    If Money_Allocation.Value = "" Or Not IsNumeric(Money_Allocation.Value) Then
        MsgBox "Please enter a valid money allocation amount"
        Exit Sub
    End If

    ' checking for a valid date
    If Not (IsNumeric(Year.Value) And IsNumeric(Month.Value) And IsNumeric(Day.Value)) Then
        MsgBox "Day, Month, and Year must be numbers"
        Exit Sub
    End If

    On Error GoTo InvalidDate
    Dim goalDate As Date
    goalDate = DateSerial(Year.Value, Month.Value, Day.Value)
    On Error GoTo 0

    ' add amount allocated
    If Amount_Allocated.Value = "" Then
        Amount_Allocated.Value = 0
    ElseIf Not IsNumeric(Amount_Allocated.Value) Then
        MsgBox "Amount allocated must be a number"
        Exit Sub
    End If

    ' next empty row
    intRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1

    ' write data
    ws.Cells(intRow, "A").Value = goalDate
    ws.Cells(intRow, "B").Value = Goal_Category.Value
    ws.Cells(intRow, "C").Value = ""                'empty (removed column)
    ws.Cells(intRow, "D").Value = Money_Allocation.Value
    ws.Cells(intRow, "E").Value = Amount_Allocated.Value

    MsgBox "Goal saved successfully!"
    Exit Sub

InvalidDate:
    MsgBox "Please enter a valid date."
    On Error GoTo 0
End Sub

Private Sub Cancel_Btn_Click()
    Unload Me
End Sub
