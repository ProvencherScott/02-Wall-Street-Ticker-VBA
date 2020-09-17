Attribute VB_Name = "Module3"
Sub Volume()

    ' Inserting Data Via Cells
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

    ' Set variables
    'Originally, was dim Ticker As string. Ticker = Cells(i + 1).value
    Dim Ticker As String
    Dim row As Integer
    Dim Volume As Double
    row = 2
    Lastrow = Cells(Rows.Count, 1).End(xlUp).row
    
    'TOTAL the VOLUME for each assigned ticker by looping through Column 7.
    
    'Loop through the rows in column 1
    For i = 2 To Lastrow
    
        'you need an If statement here'
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        'If so, Print the Ticker symbol from column 1 into column 9, begin at row 2.
        Ticker = Cells(i, 1).Value
        Cells((row), 9).Value = Ticker
        
       
    'Loop through rows in column 7
    
       Volume = Volume + Cells(i, 7).Value
        Cells((row), 12).Value = Volume
    
    
    ' Set Volume to 0
    Volume = 0
    
    row = row + 1
    
    
    Else: Volume = Volume + Cells(i, 7).Value
    
        
    End If
         
    Next i


End Sub

