Attribute VB_Name = "Module1"
Sub Multipleyearstockdata():
    
    'This has been tested and is running on all the worksheets
    Dim ws As Worksheet
    For Each ws In Worksheets
    ws.Activate
    
     'Inserting Data Via Cells
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest Totlal Volume"
    
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"

    

        'RETRIEVE each TICKER symbol from Column 1.

    ' Set variables
    'Originally, was dim Ticker As string and was here Ticker = Cells(i + 1).value
    Dim Ticker As String
    Dim row As Integer
    Dim Volume As Double
    'Initially set the Volume to be 0 for each row
    Volume = 0
    Dim opening_price As Double
    Dim closing_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    row = 2
    'Used to loop through columns. Formula will start from the last row and work up.
    Lastrow = Cells(Rows.Count, 1).End(xlUp).row
    
    
    
    
    'Loop through each row in column 1
    For i = 2 To Lastrow
    
        'you need an If statement here'
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        'If so, Print the Ticker symbol from column 1 into column 9, begin at row 2.
        Ticker = Cells(i, 1).Value
        Cells((row), 9).Value = Ticker
        
       
       
       
        'TOTAL the VOLUME for each assigned ticker by looping through Column 7.

    'Loop through rows in column 7
    
    Volume = Volume + Cells(i, 7).Value
    Cells((row), 12).Value = Volume
    

        'Set the Volume back to 0
    Volume = 0
    
     
     'set opening price
        opening_price = Cells((row), 3).Value
        
        'set closing price
        closing_price = Cells(i, 6).Value
        
        'write the equation yearly change, closing price minus opening price.
        yearly_change = (closing_price - opening_price)
        Cells((row), 10).Value = yearly_change
        
        'write the equation percent change, (yearly change / opening price)
            If (opening_price = 0 And closing_price = 0) Then
                percent_change = 0
                ElseIf (opening_price = 0 And closing_price <> 0) Then
                percent_change = 1
                Else: percent_change = yearly_change / opening_price
                Cells((row), 11).Value = percent_change
                
            End If
    
    'set the next opening price
    open_price = Cells(i + 1, 3).Value
    
    'need to format column to Percent, found format reference online.
        Cells((row), 11).NumberFormat = "0.00%"
        
    
    'need to do conditional formatting to highlight cells with ColorIndex
    'Check if yearly change is greater than 0.
            ' Set the Cell Colors to Green
            If Cells((row), 10).Value > 0 Then
            Cells((row), 10).Interior.ColorIndex = 4
            ' Set the Cell Colors to Red
            ElseIf Cells((row), 10).Value < 0 Then
            Cells((row), 10).Interior.ColorIndex = 3
            
            End If

        
    
    
    
    
    'you need to add a 1 to row here. This must included in the If statement or it will only show data from last row.
    row = row + 1
        
     'Else statement must be after row = row + 1 or it shows Error
    Else: Volume = Volume + Cells(i, 7).Value
    
    
        End If
         
    
    Next i

    'Retrieve maximum volume
    volume_max = WorksheetFunction.Max(Range("L2:L8000"))
    Cells(4, 16).Value = volume_max
    'Retrieve the Ticker symbol.attempted to use the Match function and could not get it to work.
   ' max_volumeticker = WorskeetFunction.Match(volume_max, Range("L2:L8000"), 0)
    'Cells(4, 15).Value = Cells(max_volumeticker + 1, 1)
    
    'Retreive Minimum Percent Change
    percent_min = WorksheetFunction.Min(Range("K2:K8000"))
    Cells(3, 16).Value = percent_min
    Cells(3, 16).NumberFormat = "0.00%"

    'Retreive Maximum Percent Change
    percent_max = WorksheetFunction.Max(Range("K2:K8000"))
    Cells(2, 16).Value = percent_max
    Cells(2, 16).NumberFormat = "0.00%"




    Next ws


End Sub





