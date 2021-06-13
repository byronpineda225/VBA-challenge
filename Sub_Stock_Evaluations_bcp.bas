<<<<<<< HEAD
Attribute VB_Name = "Module1"
Sub Stock_Evaluations()

'============================================
'
' Author: Byron Pineda
' Date: 6/12/2021
'
'============================================
'
' A VBA script was created that loops through all stock worksheets by year and
' generates key information relating to ticker, yearly change, percentage
' change, and total stock volume. In addition, Bonus items were implemented
' for obtaining greatest total volume by ticker; greatest percentage increase;
' and greatest percentage decrease.
'
' The yearly change is measured as the change from the stock's opening
' price at the beginning of a given year to the closing price at the end of
' that year.
'
' The percentage change is the differential from the opening price at the
' beginning of a given year to its closing price at the end of that year.
'
' Also the total volume of the stock is measured by ticker for a given year.
'
' The yearly change is colored to indicate losses, gains, or zero changes.
' A green Yearly Change cell indicates a positive change; a red Yearly
' Change cell indicates a negative change; and a yellow Yearly Change
' indicates a zero change.
'
' All of the Bonus section was completed successfully.  The greatest percentage
' increase/decrease and the greated total volume were added to the secondary
' summary table.  As stated earlier, the VBA script will run on all worksheets, every
' year, just by running the script once.  A message box pops up after completion
' indicating that all worksheets have been processed to alert the user.
'
' I need to pay credit for VBA Session 3 class activities notably #6 and #7 that
' played a key role in allowing this homework to be successfully completed. Those
' activities provided basic code and structures that were  implemented for this homework.
' Those were carefully curated enabling such key concepts as checking the next row
' against the current row and processing a batch of worksheets with one run command.
' Those takeaways saved countless hours!  In addition those activities showed the importance
' of commenting of code and making it easier to follow the logic.
'
' Finally credit must be given to our study group that collaborated on concepts for this
' challenging assignment.
'
'============================================
 
     ' Loop through all the stock worksheets by year
    
    Dim ws As Worksheet
    
    For Each ws In Worksheets

        Dim WorksheetName As String
    
  ' Set an initial variable for holding the ticker name
        Dim Ticker_Name As String

  ' Set an initial variable for holding the opening/closing prices, total volume,
  ' yearly change, and percentage change in the opening/closing prices
  
        Dim Opening_Price As Double
        Dim Closing_Price As Double
        Dim Total_Stock_Volume As Double
        Dim Percentage_Change As Double
        Dim Yearly_Change As Double
        Dim PC_as_Percentage As String
    
  'Initialize the variables
        Opening_Price = 0
        Closing_Price = 0
        Total_Stock_Volume = 0
        Percentage_Change = 0
        Yearly_Change = 0
        PC_as_Percentage = ""
    
  'Set the column headers for the new summary table
  
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
  ' Bonus column headers for another summary table
  ' with metrics for Greatest % increase/decrease, and
  ' Greatest total volume.
  
       ws.Cells(2, 15).Value = "Greatest % Increase"
       ws.Cells(3, 15).Value = "Greatest % Decrease"
       ws.Cells(4, 15).Value = "Greatest Total Volume"
       ws.Cells(1, 16).Value = "Ticker"
       ws.Cells(1, 17).Value = "Value"
  
  'Set the column widths so numbers are not squished!
  ' Right align the Yearly Change, Percentage Change,
  ' and Total Stock Volume headers in the summary table.
  
        ws.Columns("I").ColumnWidth = 12
        ws.Columns("J").ColumnWidth = 15
        ws.Columns("J").Cells.HorizontalAlignment = xlHAlignRight
        ws.Columns("K").ColumnWidth = 15
        ws.Columns("K").Cells.HorizontalAlignment = xlHAlignRight
        ws.Columns("L").ColumnWidth = 20
        ws.Columns("L").Cells.HorizontalAlignment = xlHAlignRight
        
     '  Bonus columns Ticker/Value etc.
       ws.Columns("O").ColumnWidth = 20
       ws.Columns("P").ColumnWidth = 12
       ws.Columns("Q").ColumnWidth = 15
       ws.Columns("Q").Cells.HorizontalAlignment = xlHAlignRight

  ' Keep track of the location for each ticker in the summary table
  ' Set this initially at 2 as the column headers are in the first row
        
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
  
   ' Get the WorksheetName
   ' The MsgBox showing worksheet name will be commented out
   ' later.  It was used just to make sure that each was being
   ' processed one at a time.
   
         WorksheetName = ws.Name
   '      MsgBox (WorksheetName + " is being processed")
   
  ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

  ' Loop through all the ticker transactions
    
 ' Set the first ticker's opening price
         Opening_Price = ws.Cells(2, 3).Value
 
  For i = 2 To LastRow
      
    ' Check if we are still within the same Ticker if not then
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the Ticker name
            Ticker_Name = ws.Cells(i, 1).Value
      
      'Set the Closing Price
            Closing_Price = ws.Cells(i, 6).Value

      ' Add to the Stock Volume Total
             Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
      
        ' Print the Ticker in the Summary Table
            ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
      
       'Print the Yearly Change in the Summary Table
            Yearly_Change = (Closing_Price - Opening_Price)
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
      
      ' Print the Percentage Change in the Summary Table.
      ' You need to account for division by zero.
      ' If opening price is zero reset the percentage change to zero.
      
             If Opening_Price <> 0 Then
                     Percentage_Change = ((Yearly_Change / Opening_Price))
                     ws.Range("K" & Summary_Table_Row).Value = Percentage_Change
             Else
                    ws.Range("K" & Summary_Table_Row).Value = 0
            End If
                    
        ' Set the Yearly Change cell to green if it is positive;
        ' set to red if negative; and to yellow if equal to zero.
        
                If Yearly_Change > 0 Then
                
                    'Color the Yearly Change cell green
                          ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                    
                ElseIf Yearly_Change < 0 Then
                     'Color the Yearly Change cell red
                           ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                Else
                     ' Color the Yearly Change cell yellow
                           ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 6
                     
                End If
             
             ' Convert Percentage_Change to a string;
             ' change to %; and limit it to 2 characters
             ' after the decimal.  Put this conversion
             ' here to ensure calculations are not adversely affected.
             
             PC_as_Percentage = CStr(Percentage_Change)
             ws.Range("K" & Summary_Table_Row).Value = FormatPercent(PC_as_Percentage, 2)

      ' Print the Total Stock Volume to the Summary Table
              ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
      
      ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
            
      ' Reset the Volume Totals and Ticker name
                Total_Stock_Volume = 0
                Closing_Price = 0
                Yearly_Change = 0
                Ticker_Name = ""
            
       'Reset the Opening Price for the the next new ticker
       ' i.e. the ticker is changing from "A" to "AA" etc.
       
                Opening_Price = ws.Cells(i + 1, 3).Value
      
         ' If the cell immediately below is the same brand continue
         ' accumulating total stock volume.
                
                Else

      ' Add to the Stock Volume Total
             Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
     
     End If
     
  Next i
  
 '===============================================
 ' This section is mainly devoted to getting the data elements
 ' for the Bonus section.
 '===============================================
 
 ' Determine the Last Row of the Summary Table's Column J
 ' This is needed to know the number of rows to scan against.
 
  LastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
 
'**********************************************************************
' Now fetch the greatest % increase for the Bonus Summary Table
' but first get the needed variables and initialize them.

 Dim Tckr As String
Dim GreatestPerIncr As Double
Tckr = ""
GreatestPerIncr = 0

For r = 2 To LastRow2
        If ws.Cells(r, 11) > GreatestPerIncr Then
            Tckr = ws.Cells(r, 9)
            GreatestPerIncr = ws.Cells(r, 11)
        End If
Next r

' Output the greatest % Increase and the associated ticker
' Remember to change to FormatPercent for greatest % increase
ws.Cells(2, 16).Value = Tckr
ws.Cells(2, 17).Value = FormatPercent(GreatestPerIncr, 2)

'**********************************************************************
' Now fetch the greatest % decrease for the Bonus Summary Table
' but first set and initialize the variables.

Dim GreatestPerDecr As Double
GreatestPerDecr = 0

For p = 2 To LastRow2
        If ws.Cells(p, 11) < GreatestPerDecr Then
            Tckr = ws.Cells(p, 9)
            GreatestPerDecr = ws.Cells(p, 11)
        End If
Next p

' Output the greatest % decrease and the associated ticker
' Remember to change to FormatPercent for greatest % decrease
ws.Cells(3, 16).Value = Tckr
ws.Cells(3, 17).Value = FormatPercent(GreatestPerDecr, 2)

'**********************************************************************
 ' As part of the Bonus fetch the Greatest Total Volume for the second summary table.
 ' Run a bubble sort against column J. This is not very efficient but the data
 ' set in that column is relatively small.
 
Dim MaxTotVolume As Double
MaxTotVolume = 0
  
For q = 2 To LastRow2
        If ws.Cells(q, 12) > MaxTotVolume Then
            Tckr = ws.Cells(q, 9)
            MaxTotVolume = ws.Cells(q, 12)
        End If
Next q

' Output the greatest total volume and the associated ticker
ws.Cells(4, 16).Value = Tckr
ws.Cells(4, 17).Value = MaxTotVolume

'**********************************************************************

 ' Start processing the next worksheet
 
Next ws

MsgBox ("The stock data has finished processing")

End Sub

=======
Attribute VB_Name = "Module1"
Sub Stock_Evaluations()

'============================================
'
' Author: Byron Pineda
' Date: 6/12/2021
'
'============================================
'
' A VBA script was created that loops through all stock worksheets by year and
' generates key information relating to ticker, yearly change, percentage
' change, and total stock volume. In addition, Bonus items were implemented
' for obtaining greatest total volume by ticker; greatest percentage increase;
' and greatest percentage decrease.
'
' The yearly change is measured as the change from the stock's opening
' price at the beginning of a given year to the closing price at the end of
' that year.
'
' The percentage change is the differential from the opening price at the
' beginning of a given year to its closing price at the end of that year.
'
' Also the total volume of the stock is measured by ticker for a given year.
'
' The yearly change is colored to indicate losses, gains, or zero changes.
' A green Yearly Change cell indicates a positive change; a red Yearly
' Change cell indicates a negative change; and a yellow Yearly Change
' indicates a zero change.
'
' All of the Bonus section was completed successfully.  The greatest percentage
' increase/decrease and the greated total volume were added to the secondary
' summary table.  As stated earlier, the VBA script will run on all worksheets, every
' year, just by running the script once.  A message box pops up after completion
' indicating that all worksheets have been processed to alert the user.
'
' I need to pay credit for VBA Session 3 class activities notably #6 and #7 that
' played a key role in allowing this homework to be successfully completed. Those
' activities provided basic code and structures that were  implemented for this homework.
' Those were carefully curated enabling such key concepts as checking the next row
' against the current row and processing a batch of worksheets with one run command.
' Those takeaways saved countless hours!  In addition those activities showed the importance
' of commenting of code and making it easier to follow the logic.
'
' Finally credit must be given to our study group that collaborated on concepts for this
' challenging assignment.
'
'============================================
 
     ' Loop through all the stock worksheets by year
    
    Dim ws As Worksheet
    
    For Each ws In Worksheets

        Dim WorksheetName As String
    
  ' Set an initial variable for holding the ticker name
        Dim Ticker_Name As String

  ' Set an initial variable for holding the opening/closing prices, total volume,
  ' yearly change, and percentage change in the opening/closing prices
  
        Dim Opening_Price As Double
        Dim Closing_Price As Double
        Dim Total_Stock_Volume As Double
        Dim Percentage_Change As Double
        Dim Yearly_Change As Double
        Dim PC_as_Percentage As String
    
  'Initialize the variables
        Opening_Price = 0
        Closing_Price = 0
        Total_Stock_Volume = 0
        Percentage_Change = 0
        Yearly_Change = 0
        PC_as_Percentage = ""
    
  'Set the column headers for the new summary table
  
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
  ' Bonus column headers for another summary table
  ' with metrics for Greatest % increase/decrease, and
  ' Greatest total volume.
  
       ws.Cells(2, 15).Value = "Greatest % Increase"
       ws.Cells(3, 15).Value = "Greatest % Decrease"
       ws.Cells(4, 15).Value = "Greatest Total Volume"
       ws.Cells(1, 16).Value = "Ticker"
       ws.Cells(1, 17).Value = "Value"
  
  'Set the column widths so numbers are not squished!
  ' Right align the Yearly Change, Percentage Change,
  ' and Total Stock Volume headers in the summary table.
  
        ws.Columns("I").ColumnWidth = 12
        ws.Columns("J").ColumnWidth = 15
        ws.Columns("J").Cells.HorizontalAlignment = xlHAlignRight
        ws.Columns("K").ColumnWidth = 15
        ws.Columns("K").Cells.HorizontalAlignment = xlHAlignRight
        ws.Columns("L").ColumnWidth = 20
        ws.Columns("L").Cells.HorizontalAlignment = xlHAlignRight
        
     '  Bonus columns Ticker/Value etc.
       ws.Columns("O").ColumnWidth = 20
       ws.Columns("P").ColumnWidth = 12
       ws.Columns("Q").ColumnWidth = 15
       ws.Columns("Q").Cells.HorizontalAlignment = xlHAlignRight

  ' Keep track of the location for each ticker in the summary table
  ' Set this initially at 2 as the column headers are in the first row
        
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
  
   ' Get the WorksheetName
   ' The MsgBox showing worksheet name will be commented out
   ' later.  It was used just to make sure that each was being
   ' processed one at a time.
   
         WorksheetName = ws.Name
   '      MsgBox (WorksheetName + " is being processed")
   
  ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

  ' Loop through all the ticker transactions
    
 ' Set the first ticker's opening price
         Opening_Price = ws.Cells(2, 3).Value
 
  For i = 2 To LastRow
      
    ' Check if we are still within the same Ticker if not then
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the Ticker name
            Ticker_Name = ws.Cells(i, 1).Value
      
      'Set the Closing Price
            Closing_Price = ws.Cells(i, 6).Value

      ' Add to the Stock Volume Total
             Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
      
        ' Print the Ticker in the Summary Table
            ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
      
       'Print the Yearly Change in the Summary Table
            Yearly_Change = (Closing_Price - Opening_Price)
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
      
      ' Print the Percentage Change in the Summary Table.
      ' You need to account for division by zero.
      ' If opening price is zero reset the percentage change to zero.
      
             If Opening_Price <> 0 Then
                     Percentage_Change = ((Yearly_Change / Opening_Price))
                     ws.Range("K" & Summary_Table_Row).Value = Percentage_Change
             Else
                    ws.Range("K" & Summary_Table_Row).Value = 0
            End If
                    
        ' Set the Yearly Change cell to green if it is positive;
        ' set to red if negative; and to yellow if equal to zero.
        
                If Yearly_Change > 0 Then
                
                    'Color the Yearly Change cell green
                          ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                    
                ElseIf Yearly_Change < 0 Then
                     'Color the Yearly Change cell red
                           ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                Else
                     ' Color the Yearly Change cell yellow
                           ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 6
                     
                End If
             
             ' Convert Percentage_Change to a string;
             ' change to %; and limit it to 2 characters
             ' after the decimal.  Put this conversion
             ' here to ensure calculations are not adversely affected.
             
             PC_as_Percentage = CStr(Percentage_Change)
             ws.Range("K" & Summary_Table_Row).Value = FormatPercent(PC_as_Percentage, 2)

      ' Print the Total Stock Volume to the Summary Table
              ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
      
      ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
            
      ' Reset the Volume Totals and Ticker name
                Total_Stock_Volume = 0
                Closing_Price = 0
                Yearly_Change = 0
                Ticker_Name = ""
            
       'Reset the Opening Price for the the next new ticker
       ' i.e. the ticker is changing from "A" to "AA" etc.
       
                Opening_Price = ws.Cells(i + 1, 3).Value
      
         ' If the cell immediately below is the same brand continue
         ' accumulating total stock volume.
                
                Else

      ' Add to the Stock Volume Total
             Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
     
     End If
     
  Next i
  
 '===============================================
 ' This section is mainly devoted to getting the data elements
 ' for the Bonus section.
 '===============================================
 
 ' Determine the Last Row of the Summary Table's Column J
 ' This is needed to know the number of rows to scan against.
 
  LastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
 
'**********************************************************************
' Now fetch the greatest % increase for the Bonus Summary Table
' but first get the needed variables and initialize them.

 Dim Tckr As String
Dim GreatestPerIncr As Double
Tckr = ""
GreatestPerIncr = 0

For r = 2 To LastRow2
        If ws.Cells(r, 11) > GreatestPerIncr Then
            Tckr = ws.Cells(r, 9)
            GreatestPerIncr = ws.Cells(r, 11)
        End If
Next r

' Output the greatest % Increase and the associated ticker
' Remember to change to FormatPercent for greatest % increase
ws.Cells(2, 16).Value = Tckr
ws.Cells(2, 17).Value = FormatPercent(GreatestPerIncr, 2)

'**********************************************************************
' Now fetch the greatest % decrease for the Bonus Summary Table
' but first set and initialize the variables.

Dim GreatestPerDecr As Double
GreatestPerDecr = 0

For p = 2 To LastRow2
        If ws.Cells(p, 11) < GreatestPerDecr Then
            Tckr = ws.Cells(p, 9)
            GreatestPerDecr = ws.Cells(p, 11)
        End If
Next p

' Output the greatest % decrease and the associated ticker
' Remember to change to FormatPercent for greatest % decrease
ws.Cells(3, 16).Value = Tckr
ws.Cells(3, 17).Value = FormatPercent(GreatestPerDecr, 2)

'**********************************************************************
 ' As part of the Bonus fetch the Greatest Total Volume for the second summary table.
 ' Run a bubble sort against column J. This is not very efficient but the data
 ' set in that column is relatively small.
 
Dim MaxTotVolume As Double
MaxTotVolume = 0
  
For q = 2 To LastRow2
        If ws.Cells(q, 12) > MaxTotVolume Then
            Tckr = ws.Cells(q, 9)
            MaxTotVolume = ws.Cells(q, 12)
        End If
Next q

' Output the greatest total volume and the associated ticker
ws.Cells(4, 16).Value = Tckr
ws.Cells(4, 17).Value = MaxTotVolume

'**********************************************************************

 ' Start processing the next worksheet
 
Next ws

MsgBox ("The stock data has finished processing")

End Sub

>>>>>>> acfc758c544cb68682b5e6f18de48688cc076f41
