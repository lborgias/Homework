Attribute VB_Name = "VBA_HW_LB"
Sub TickerTape()
        
    Application.ScreenUpdating = False

    
        Dim Ticker As String
        Dim CurrentTickerDate As Long
        Dim CurrentOpen As Double
        Dim CurrentCose As Double
        Dim MinTickerDate As Long
        Dim MinOpen As Double
        Dim MaxTickerDate As Long
        Dim MaxClose As Double
        Dim Volume As Double
        Dim RowCount As Long
        
  For Each ws In ActiveWorkbook.Worksheets
  
        ws.Activate
        
        ' Set RowCount for sheet
        RowCount = Cells(Rows.Count, 1).End(xlUp).Row
                    
        Dim SummaryColumn As Integer
        ' Count number of columns in first row and add 2 for start column of summary table
        SummaryColumn = Cells(1, Columns.Count).End(xlToLeft).Column + 2
        
        'Declare Summary Table last row
        Dim SummaryRow As Integer
        
        'Set initial values
        Volume = 0
        MinTickerDate = Cells(2, 2).Value
        MaxTickerDate = Cells(2, 2).Value
        
        For i = 2 To RowCount
        
        Ticker = Cells(i, 1).Value
        CurrentTickerDate = Cells(i, 2).Value
           
            
        If Ticker <> Cells(i + 1, 1).Value Then
            
               'Compare Min Values
                If CurrentTickerDate <= MinTickerDate Then
                    MinTickerDate = CurrentTickerDate
                    MinOpen = Cells(i, 3).Value
                End If
            
            ' Compare Max Values
                If CurrentTickerDate >= MaxTickerDate Then
                    MaxTickerDate = CurrentTickerDate
                    MaxClose = Cells(i, 6).Value
            End If
            
            Volume = Volume + Cells(i, 7).Value
            
            'Find last row of summary table
            SummaryRow = Cells(Rows.Count, SummaryColumn).End(xlUp).Row + 1
            
                YearlyChange = MaxClose - MinOpen
                    'MsgBox (MinTickerDate & "," & MinOpen)
                   ' MsgBox (MaxTickerDate & "," & MaxClose)
                If MinOpen = 0 Then
                PercentChange = 1 ' Not sure of this value
                Else
                
                PercentChange = (YearlyChange / MinOpen)
                End If
            'Set Summary Table Values
             ' Set values for header row
                Cells(1, SummaryColumn).Value = "Ticker"
                Cells(1, SummaryColumn + 1).Value = "Yearly Change"
                Cells(1, SummaryColumn + 2).Value = "Percent Change"
                Cells(1, SummaryColumn + 3).Value = "Total Stock Volume"
                ' Set Values
                Cells(SummaryRow, SummaryColumn) = Ticker
                Cells(SummaryRow, SummaryColumn + 1).Value = YearlyChange
                Cells(SummaryRow, SummaryColumn + 1).NumberFormat = "0.000000000" 'Tried a million different ways to format/Round , unable to get exact value on instructions
                Cells(SummaryRow, SummaryColumn + 2).Value = Format(PercentChange, "Percent")
                Cells(SummaryRow, SummaryColumn + 3).Value = Volume
                
                'Format Cell color
                If YearlyChange < 0 Then
                 Cells(SummaryRow, SummaryColumn + 1).Interior.ColorIndex = 3
                Else: Cells(SummaryRow, SummaryColumn + 1).Interior.ColorIndex = 50
                    End If
                
                ' Reset Values
                Volume = 0
                MinTickerDate = Cells(i + 1, 2).Value
                MinOpen = Cells(i + 1, 3).Value
                MaxClose = Cells(i + 1, 6).Value
                'MsgBox (Volume)
        Else
            
            'CurrentTickerDate = Cells(i, 2).Value
            Volume = Volume + Cells(i, 7).Value
            CurrentOpen = Cells(i, 3).Value
            CurrentClose = Cells(i, 6).Value
            
           'Compare Min Values
                If CurrentTickerDate <= MinTickerDate Then
                    MinTickerDate = CurrentTickerDate
                    MinOpen = Cells(i, 3).Value
                End If
            
            ' Compare Max Values
                If CurrentTickerDate >= MaxTickerDate Then
                    MaxTickerDate = CurrentTickerDate
                    MaxClose = Cells(i, 6).Value
                End If
         End If
                
      Next i
        
        
   '-------------------------- Increate/Decrease/Volme--------------------------
   
   'Update RowCount & Column
        
    SummaryColumn = Cells(1, Columns.Count).End(xlToLeft).Column + 2
               ' MsgBox (SummaryColumn)
    RowCount = Cells(Rows.Count, SummaryColumn - 2).End(xlUp).Row
                'MsgBox (RowCount)
                
    
    'Declare Variables
    Dim IncreaseTicker As String
    Dim Increase As Double
    Dim DecreaseTicker As String
    Dim Decrease As Double
    Dim VolumeTicker As String
    Dim MaxVolume As Double
    
    
    Dim CurrentTicker As String
    Dim CurrentPercent As Double
    Dim CurrentVolume As Double
    
    CurrentTicker = Cells(2, SummaryColumn - 5).Value
    CurrentPercent = Cells(2, SummaryColumn - 3).Value
    CurrentVolume = Cells(2, SummaryColumn - 2).Value
    
    
    IncreaseTicker = CurrentTicker
    Increase = CurrentPercent
    
    DecreaseTicker = CurrentTicker
    Decrease = CurrentPercent
    
    VolumeTicker = CurrentTicker
    MaxVolume = CurrentVolume
   
    For j = 2 To RowCount
        
        CurrentTicker = Cells(j, SummaryColumn - 5).Value
        CurrentPercent = Cells(j, SummaryColumn - 3).Value
        CurrentVolume = Cells(j, SummaryColumn - 2).Value
    
    ' Compare Increase
        If CurrentPercent >= Increase Then
            Increase = CurrentPercent
            IncreaseTicker = CurrentTicker
        End If
    
    ' Compare Decrease
        If CurrentPercent <= Decrease Then
            Decrease = CurrentPercent
            DecreaseTicker = CurrentTicker
        End If
    
    ' Compare Volume
        If CurrentVolume > MaxVolume Then
            MaxVolume = CurrentVolume
            VolumeTicker = CurrentTicker
        End If
        
    Next j
    
    Cells(1, SummaryColumn + 1).Value = "Ticker"
    Cells(1, SummaryColumn + 2).Value = "Value"
    
    Cells(2, SummaryColumn).Value = "Greatest % Increase"
    Cells(2, SummaryColumn + 1).Value = IncreaseTicker
    Cells(2, SummaryColumn + 2).Value = Format(Increase, "Percent")
    
    
    
    Cells(3, SummaryColumn).Value = "Greatest % Decrease"
    Cells(3, SummaryColumn + 1).Value = DecreaseTicker
    Cells(3, SummaryColumn + 2).Value = Format(Decrease, "Percent")
    
    Cells(4, SummaryColumn).Value = "Greatest Total Volume"
    Cells(4, SummaryColumn + 1).Value = VolumeTicker
    Cells(4, SummaryColumn + 2).Value = MaxVolume
    
   Next
    
   Application.ScreenUpdating = True
    
    MsgBox ("Complete!")
End Sub
