Attribute VB_Name = "Module1"

Public Sub WorkingWithIFStatements()

    ' Checking the value of the currently active cell
    ' and displaying a message according to its value
    If ActiveCell.Value >= 90 Then
        MsgBox ("User is 90 or older")
    ElseIf ActiveCell.Value >= 21 Then
        MsgBox ("User is 21 or older")
    Else
        MsgBox ("User is younger than 21 years old")
    End If
    
           
     
End Sub


Public Sub WorkingWithSelectStatement()

    'Using Select statements to check the currently active cell value
    Select Case ActiveCell.Value
            Case Is > 90
                MsgBox ("Greater than 90")
            Case 21 To 89
                MsgBox ("Greater than 21")
            Case Else
                MsgBox ("Smaller than 21")
    End Select
    
End Sub


Public Sub WorkingWithWhileLoop()
    Dim i As Integer
    i = 1
    ActiveSheet.Range("B2").Select
    Do While i <= 10
        WorkingWithIFStatements
        ActiveCell.Offset(1, 0).Select
        i = i + 1
    Loop
End Sub


Public Sub DynamicLoop()
' Same thing as the previous procedure but more dynamic
' Dynamic in the sense that the user may delete or add records
    ActiveSheet.Range("B2").Select
    Do While ActiveCell.Value <> ""
        WorkingWithIFStatements
        ActiveCell.Offset(1, 0).Select
    Loop
'We took advantage of the empty cell after the end of the AGE data
'But this method creates problems when the AGE column has missing data
End Sub


Public Sub WorkingWithForEachLoop()
' Using for each statement to solve the previously mentioned problem
    Dim user As Range
    
    For Each user In Selection
        WorkingWithIFStatements
        ActiveCell.Offset(1, 0).Select
    Next user
End Sub

Public Sub WorkingWithNextLoop()
' Optimal way to solve all the above mentioned problems
    Dim i As Integer
    ActiveSheet.Range("B2").Select
    For i = 1 To (ActiveSheet.UsedRange.Rows.Count - 1)
        WorkingWithIFStatements
        ActiveCell.Offset(1, 0).Select
    Next i
End Sub
