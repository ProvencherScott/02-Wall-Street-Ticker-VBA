Attribute VB_Name = "Module2"
Sub Ticker()

        'Retrieve each ticker symbol from Column 1.

    ' Set variables
    'Originally, was dim Ticker As string. Ticker = Cells(i + 1).value
    Dim Ticker As String
    Dim row As Integer
    row = 2
    Lastrow = Cells(Rows.Count, 1).End(xlUp).row
    
    
    'Loop through the rows in column 1
    For i = 2 To Lastrow
    
        'you need an If statement here'
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        'If so, Print the Ticker symbol from column 1 into column 9, begin at row 2.
        Ticker = Cells(i, 1).Value
        Cells((row), 9).Value = Ticker
   
    
    
    'need to add a row to sumary table
    row = row + 1


            End If
         
    Next i


End Sub




    
    
    
    
    'Needs to be added in
    'OpenPrice = Cell(3, 1).Value

