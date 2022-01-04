

 'Loop through each Worksheets in Spreadsheet and Calculate Grand Totals
     Sub WorksheetLoop1()

         Dim xSH As Worksheet
         Application.ScreenUpdating = False

         ' Begin the loop.
         For Each xSH In Worksheets
            xSH.Select
            Call stock_assess
         Next
         Application.ScreenUpdating = True
         
         Call GrandTotal
     
          
                 
      End Sub

'Calculate all the Totals in each Worksheet

Sub stock_assess()

  ' Set an initial variable for holding the stock name
  Dim Stock_Name As String

  ' Set an initial variable for holding the total stock volume
  Dim Stock_Total As Double
  Stock_Total = 0
  
' Set initial variables for calculating and holding yearly stock change
  Dim Yearly_Stock_Start As Double
  Yearly_Stock_Start = Cells(2, 6).Value
  
  Dim Yearly_Stock_End As Double
  Yearly_Stock_End = 0
  
  Dim Yearly_Stock_Change As Double
  Yearly_Stock_Change = 0
  
  Dim Yearly_Stock_Percent_Change As Double
  Yearly_Stock_Percent_Change = 0
  


  ' Keep track of the location for each stock name in the summary table
  Dim Summary_Table_Row As Double
  Summary_Table_Row = 2
  Dim lastrow As Long
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row

  ' Loop through all stock prices
  For i = 2 To lastrow
    ' Check if we are still within the same stock, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Stock name
      Stock_Name = Cells(i, 1).Value
    
      
      ' Calculate the Yearly Stock Price Change
      Yearly_Stock_Change = Cells(i, 6).Value - Yearly_Stock_Start
    
          
      ' Calculate the Yearly Stock Price Percent Change
      If Yearly_Stock_Start <> 0 Then
       Yearly_Stock_Percent_Change = (Yearly_Stock_Change / Yearly_Stock_Start)
      
    
      
      Else: Yearly_Stock_Percent_Change = 0
      End If
      

      ' Add to the Stock Total
      Stock_Total = Stock_Total + Cells(i, 7).Value
   
              
      ' Print the Stock Name in the Summary Table
      Range("J" & Summary_Table_Row).Value = Stock_Name
      
      
      ' Print the Yearly Change in the Summary Table
      Range("K" & Summary_Table_Row).Value = Yearly_Stock_Change
      
      ' Print the Yearly Change Percent in the Summary Table
       Range("L" & Summary_Table_Row).Value = Yearly_Stock_Percent_Change
       
                 
      ' Print the Brand Amount to the Summary Table
      Range("M" & Summary_Table_Row).Value = Stock_Total

         
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Stock Total
      Stock_Total = 0
      
      ' Reset the Yearly Stock Start Amount
      Yearly_Stock_Start = Cells(i + 1, 6).Value

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Stock Total
      Stock_Total = Stock_Total + Cells(i, 7).Value
      
    End If
        
  Next i
  
    'Call Routine to format Percent Change Column
    
    MakeColumnPercent
    
    'Call Routine to Format Yearly Change Column
    
    Formatredgreen
    
    'Call Routine to Format Column Headers
           
    AddColumnHeaders
    
    'Call Routine to Format Max/Min Percent and Highest Volume
           
    FindMaxPercent
    
    FindMinPercent
    
    FindMaxVolume
    
    
    
End Sub
    
'Format Percent Columns
Sub MakeColumnPercent()
'
    Columns("L:L").Select
    Selection.NumberFormat = "0.00%"
    Columns("M:M").Select
    Selection.NumberFormat = "0,000"
End Sub


'Format the Yearly Change Column Red and Green
Sub Formatredgreen()
 
    Range("K2:K5000").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub
' Add Column Headers
'
Sub AddColumnHeaders()

    Range("J1") = "Ticker"
    Range("K1") = "Yearly Change"
    Range("L1") = "Percent Change"
    Range("M1") = "Total Stock Volume"
        
End Sub

' Find Maximum Percent Increase for Each Sheet
Sub FindMaxPercent()
    
     
    
    Dim Max_Percent_Ticker As String
    Dim CellsArray As Variant
    Dim Max_Percent As Double
    Dim x As Integer
    CellsArray = Range("J2").CurrentRegion.Value
    upperbound = UBound(CellsArray)
    
        For x = 1 To upperbound
            If x = 1 Then
                Max_Percent = CellsArray(2, 3)
                Max_Percent_Ticker = CellsArray(2, 1)
               
            Else
                If CellsArray(x, 3) > Max_Percent Then
                Max_Percent = CellsArray(x, 3)
                Max_Percent_Ticker = CellsArray(x, 1)
                End If
            End If
        Next x
        
    Range("P1") = "Ticker"
    Range("Q1") = "Value"
    Range("O2") = "Greatest % Increase in Year"
    Range("O3") = "Greatest % Decrease in Year"
    Range("O4") = "Greatest Total Volume in Year"
    Range("P2") = Max_Percent_Ticker
    Range("Q2") = Max_Percent
    Range("Q2").NumberFormat = "0.00%"
     
     
End Sub

'Find Minimum % Decrease for Each Sheet
Sub FindMinPercent()
    
    Dim min_percent_ticker As String
    Dim CellsArray As Variant
    Dim Min_Percent As Double
    Dim x As Integer
    CellsArray = Range("J2").CurrentRegion.Value
    upperbound = UBound(CellsArray)
   
        For x = 1 To upperbound
            If x = 1 Then
                Min_Percent = CellsArray(2, 3)
                min_percent_ticker = CellsArray(2, 1)
               
            Else
                If CellsArray(x, 3) < Min_Percent Then
                Min_Percent = CellsArray(x, 3)
                min_percent_ticker = CellsArray(x, 1)
                End If
            End If
        Next x
        
    Range("P3") = min_percent_ticker
    Range("Q3") = Min_Percent
    Range("Q3").NumberFormat = "0.00%"
    
End Sub

'Find Maximum Volume Increase for each Sheet
Sub FindMaxVolume()
    
     
    Dim Max_Volume_Ticker As String
    Dim CellsArray As Variant
    Dim Max_Volume As Double
    Dim x As Integer
    CellsArray = Range("J2").CurrentRegion.Value
    upperbound = UBound(CellsArray)
    
        For x = 1 To upperbound
            If x = 1 Then
                Max_Volume = CellsArray(2, 4)
                Max_Volume_Ticker = CellsArray(2, 1)
               
            Else
                If CellsArray(x, 4) > Max_Volume Then
                Max_Volume = CellsArray(x, 4)
                Max_Volume_Ticker = CellsArray(x, 1)
                End If
                                
            End If
        Next x
        
    Range("P4") = Max_Volume_Ticker
    Range("Q4") = Max_Volume
    Range("Q4").NumberFormat = "0,000"
End Sub

'Find Grand Totals

Sub GrandTotal()
    
      Dim grand_percent_increase As Double
      grand_percent_increase = 0
         
      Dim grand_percent_ticker As String
      grand_percent_ticker = "False"
         
      Dim grand_min_percent_decrease As Double
      min_percent_decrease = 0
         
      Dim grand_min_percent_ticker As String
      grand_min_percent_ticker = "False"
      
      Dim grand_volume_increase As Double
      grand_volume_increase = 0
      
      Dim grand_volume_ticker As String
      grand_volume_ticker = "False"
         
      For Each xSH In Worksheets
         xSH.Select
            
              If Range("Q2") > grand_percent_increase Then
                 grand_percent_increase = Range("Q2")
                 grand_percent_ticker = Range("P2")
              Else
              End If
              
              If Range("Q3") < grand_min_percent_decrease Then
                 grand_min_percent_decrease = Range("Q3")
                 grand_min_percent_ticker = Range("P3")
              Else
              End If
              
              
              If Range("Q4") > grand_volume_increase Then
                 grand_volume_increase = Range("Q4")
                 grand_volume_ticker = Range("P4")
              Else
              End If
              
              
              
         Next
          'Print out the Grand Totals
          
              Range("P6") = "Ticker"
              Range("Q6") = "Value"
              Range("O7") = "Greatest % Increase Over All Years"
              Range("O8") = "Greatest % Decrease Over All Years"
              Range("O9") = "Greatest Total Volume Over All Years"
              
              Range("P7") = grand_percent_ticker
              Range("Q7") = grand_percent_increase
              Range("Q7").NumberFormat = "0.00%"
              
              Range("P8") = grand_min_percent_ticker
              Range("Q8") = grand_min_percent_decrease
              Range("Q8").NumberFormat = "0.00%"
              
              Range("P9") = grand_volume_ticker
              Range("Q9") = grand_volume_increase
              Range("Q9").NumberFormat = "0,000"
    
         End Sub
