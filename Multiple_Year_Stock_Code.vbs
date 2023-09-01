Sub Multiple_Year_Stock()
    

    For Each Ws In ThisWorkbook.Worksheets ''''''''''''''' To Loop through each WorkSheet
        Ws.Activate
        
        'Set Dimensions
        Dim Tickor_Symbol As String
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        Dim Total_Volume As Double
        Dim Summary_Table_Row As Integer 'Table where the results will be located
        Dim Opening_Price As Double
        Dim Greatest_Increase_Tickor As String
        Dim Greatest_Decrease_Tickor As String
        Dim Greatest_TotalVolume_Tickor As String
        Dim Greatest_Total_Volume As Double
        
        
        'Inserting Headers
        Ws.Cells(1, 9).Value = "Ticker" '''''''''''''''''''' This could also be done by typing
        Ws.Cells(1, 10).Value = "Yearly Change"            ' Range("I1 or J1, .............")
        Ws.Cells(1, 11).Value = "Percent Change"           ' Instead of Cells!
        Ws.Cells(1, 12).Value = "Total Stock Volume"
        
        Ws.Cells(2, 14).Value = "Greatest % Increase"
        Ws.Cells(3, 14).Value = "Greatest % Decrease"
        Ws.Cells(4, 14).Value = "Greatest Total Volume"
        Ws.Cells(1, 15).Value = "Ticker"
        Ws.Cells(1, 16).Value = "Value"
        
        'Set Initials
        Yearly_Change = 0
        First_OpenPrice = 2             'Starting at first Opening Price
        Summary_Table_Row = 2    'To tell us to start inserting the values starting the 2nd row
        Total_Volume = 0
        Greatest_Increase = 0
        Greatest_Decrease = 0
        Greatest_Total_Volume = 0
        Percent_Change = 0
        
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        '''''we use this code to find the last row in the sheet,
        '''''instead of counting how many rows are there
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To lastrow
            
            ''Calculating the Total_Volume
            Total_Volume = Total_Volume + Cells(i, 7).Value
            
            'When Tickor Changes print the results
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

               
                                ''''''''''' To Print Each Tickor Once '''''''''''''
               Tickor_Symbol = Ws.Cells(i, 1).Value
               
                              
                                ''''''''' To Calculate the Yearly Change '''''''''''
               Closing_Price = Ws.Cells(i, 6).Value
               Opening_Price = Ws.Cells(First_OpenPrice, 3).Value
               Yearly_Change = Closing_Price - Opening_Price
               Percent_Change = (Yearly_Change / Opening_Price)
               
              

            '''''Finding the Greatest Total Vol., % of increase and decrease and their corresponding tickors'''''
               If Percent_Change > Greatest_Increase Then
                    Greatest_Increase = Percent_Change
                    Greatest_Increase_Tickor = Tickor_Symbol
               End If
               
               If Percent_Change < Greatest_Decrease Then
                    Greatest_Decrease = Percent_Change
                    Greatest_Decrease_Tickor = Tickor_Symbol
               End If
               
               If Total_Volume > Greatest_Total_Volume Then
                    Greatest_Total_Volume = Total_Volume
                    Greatest_TotalVolume_Tickor = Tickor_Symbol
               End If

             
               
               First_OpenPrice = i + 1 'Increment the rows by 1
               
               
                              ''''''''''''''Inserting the values in the Summary Table'''''''''''''''
               Ws.Range("I" & Summary_Table_Row).Value = Tickor_Symbol               '''' This telling us to insert in col. I
               Ws.Range("J" & Summary_Table_Row).Value = Yearly_Change               '''' starting 2nd row
               Ws.Range("J" & Summary_Table_Row).NumberFormat = "0.00"               '''''''''''''''
               Ws.Range("K" & Summary_Table_Row).Value = Percent_Change
               Ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
               Ws.Range("L" & Summary_Table_Row).Value = Total_Volume
               
               Total_Volume = 0 '''Reset the Total_Volume to 0, so it calculates the total volume for the next tickor
                           
               Summary_Table_Row = Summary_Table_Row + 1 '''''Incrementing the Summary Table By 1
                
               
            End If
        Next i
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                 '''''Printing the maximum and minimum'''''
        
        Ws.Range("P2").Value = Greatest_Increase
        Ws.Range("P2").NumberFormat = "0.00%"
        Ws.Range("O2").Value = Greatest_Increase_Tickor
           
        Ws.Range("P3").Value = Greatest_Decrease
        Ws.Range("P3").NumberFormat = "0.00%"
        Ws.Range("O3").Value = Greatest_Decrease_Tickor
        
        Ws.Range("P4").Value = Greatest_Total_Volume
        Ws.Range("O4").Value = Greatest_TotalVolume_Tickor
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                  ''''''''' Formating the Sheets ''''''''''
                                  
        Ws.Range("A2:A" & lastrow).Interior.Color = RGB(217, 225, 242)
        Ws.Range("I2:I" & Summary_Table_Row - 1).Interior.Color = RGB(217, 225, 242)
        Ws.Range("A1,B1,C1,D1,E1,F1,G1,I1,J1,K1,L1,O1,P1").Interior.Color = RGB(217, 225, 242)
        Ws.Range("N2:N4").Interior.Color = RGB(217, 225, 242)
                
                  ''''''''' Changing the Color of Yearly_Changed Column Based on Value ''''''''''
        
        For j = 2 To Summary_Table_Row
            If Ws.Cells(j, 10) < 0 Then
                 Ws.Cells(j, 10).Interior.ColorIndex = 3
            ElseIf Ws.Cells(j, 10) > 0 Then
                 Ws.Cells(j, 10).Interior.ColorIndex = 4
            Else
                Ws.Cells(j, 10).Interior.ColorIndex = 0
             End If
        Next j
    
    Next Ws
        

End Sub



