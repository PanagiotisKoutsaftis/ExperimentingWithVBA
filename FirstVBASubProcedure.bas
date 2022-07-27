Attribute VB_Name = "Module1"

Public Sub FirsrtVBASubProcedure()
    ' Changing the value of a cell
    ActiveSheet.Range("A1").Value = "New Value"
    ' Changing the name of the active sheet
    ActiveSheet.Name = "Name with VBA"
    
    ' Message Box -> function that provides a promt message to the user for eg an error
    
    ' Showing the value of the A1 cell to the user through the msgbox
    MsgBox ("Cell Value = " & ActiveSheet.Range("A1").Value)
     
    ' Working with variables
    ' First the variables and their data type must be declared
    Dim name_variable As String
    Dim age As Integer
    
    ' Assigning values to the variables
    name_variable = "Panagiotis"
    age = 26
    MsgBox ("Hello " & name_variable & "! You are " & age & " years old")
    MsgBox ("You were born in " & (Year(Now()) - age) & "!")
    
    
    ' If we reach here then everything is fine
    MsgBox ("FirstVBASubProcedure executed successfully.")
    
       
        
    
End Sub
