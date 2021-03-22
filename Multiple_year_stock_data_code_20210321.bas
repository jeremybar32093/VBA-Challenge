Attribute VB_Name = "Module1"
Sub StockReport()

'Turn of screen updating for performance purposes
Application.ScreenUpdating = False

'Declare variables
Dim CurrentSheetName As String
Dim NumIterations As Long
Dim TickerCopyRange As Range
Dim TickerPasteRange As Range
Dim DateRange As Range
Dim CellLoopRange As Range
Dim MaxDate As Long
Dim MinDate As Long
Dim InitialOpenPrice As Double
Dim EndingClosePrice As Double
Dim YearlyChange As Double
Dim PercentChange As Double
Dim TotalVolume As Double
Dim WS_Count As Integer
Dim MaxPercentChange As Double
Dim MinPercentChange As Double
Dim MaxTotalVolume As Double
Dim FoundMaxPercentChange As Boolean
Dim FoundMinPercentChange As Boolean
Dim FoundMaxTotalVolume As Boolean
Dim CurrentPercentChange As Double
Dim CurrentTotalVolume As Double

' Bonus - Loop through all worksheets - populate output on each sheet by running macro once
WS_Count = ActiveWorkbook.Worksheets.Count

For j = 1 To WS_Count

    'Activate the current sheet iterated against and perform output tasks on active sheet
    Worksheets(j).Activate
    
    'Step 1 - Populate output column headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    'Step 2 - create unique list of individual tickers to loop through and perform calculations
    '         Use copy/paste and remove duplicates functionality
    '         Paste result in column I
    Set TickerCopyRange = Range("A2")
    Set TickerCopyRange = Range(TickerCopyRange, TickerCopyRange.End(xlDown))
    TickerCopyRange.Copy
    Range("I2").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    'Remove duplicates from pasted range
    Set TickerPasteRange = Range("I2")
    Set TickerPasteRange = Range(TickerPasteRange, TickerPasteRange.End(xlDown))
    TickerPasteRange.RemoveDuplicates Columns:=1
    
    'Step 3 - Loop through individual tickers and populate output information
    '         Loop range is ticker paste range from step 2 but after duplicates were removed
    Set CellLoopRange = Range("I2")
    Set CellLoopRange = Range(CellLoopRange, CellLoopRange.End(xlDown))
    NumIterations = CellLoopRange.Rows.Count
    
    For i = 1 To NumIterations
        'Determine current ticker value
        CurrentTicker = Cells(i + 1, 9).Value
        '----- Step 3a - Calculate min and max date for each ticker
        MaxDate = Application.WorksheetFunction.MaxIfs(Range("B:B"), Range("A:A"), CurrentTicker)
        MinDate = Application.WorksheetFunction.MinIfs(Range("B:B"), Range("A:A"), CurrentTicker)
        '-----------------------------------
        
        '----- Step 3b - Calculate Yearly change -----
        'For current ticker value, look up using sumifs relevant open/close prices based off of Min/Max Dates calculated in step 3
        'Sumifs will work in this case because there is only 1 open/close price for each ticker/date combination
        InitialOpenPrice = Application.WorksheetFunction.SumIfs(Range("C:C"), Range("B:B"), MinDate, Range("A:A"), CurrentTicker)
        EndingClosePrice = Application.WorksheetFunction.SumIfs(Range("F:F"), Range("B:B"), MaxDate, Range("A:A"), CurrentTicker)
        'Calculate YearlyChange and write value to column J
        YearlyChange = EndingClosePrice - InitialOpenPrice
        Cells(i + 1, 10).Value = YearlyChange
        'Check if the value is positive or negative and update background color
        'Positive change -> change background color to green
        'Negative change -> change background color to red
        'Zero change -> don't update formatting
        If YearlyChange > 0 Then
            Cells(i + 1, 10).Interior.Color = vbGreen
        ElseIf YearlyChange < 0 Then
            Cells(i + 1, 10).Interior.Color = vbRed
        End If
        '-----------------------------------
        
        '----- Step 3c - Calculate  % Change
        'Account for 0/0 scenario - in this case, just set percent change equal to 0
        If InitialOpenPrice = 0 Then
            PercentChange = 0
        Else
            PercentChange = YearlyChange / InitialOpenPrice
        End If
        
        Cells(i + 1, 11).Value = PercentChange
        'Change values to % format
        Cells(i + 1, 11).NumberFormat = "0.00%"
        '-----------------------------------
        
        '----- Step 3d - Calculate  Total Volume
        TotalVolume = Application.WorksheetFunction.SumIfs(Range("G:G"), Range("A:A"), CurrentTicker)
        Cells(i + 1, 12).Value = TotalVolume
        '-----------------------------------
    'iterate to next ticker
    Next i
    
    '----- Bonus -----
    'Declare variables
    
    
    'Step 1 - Populate output static values
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
    'Step 2 - calculat the max % change, min % change, and max total volume using worksheet functions
    'Populate in column Q
    MaxPercentChange = Application.WorksheetFunction.Max(Range("K:K"))
    Range("Q2").Value = MaxPercentChange
    Range("Q2").NumberFormat = "0.00%"
    MinPercentChange = Application.WorksheetFunction.Min(Range("K:K"))
    Range("Q3") = MinPercentChange
    Range("Q3").NumberFormat = "0.00%"
    MaxTotalVolume = Application.WorksheetFunction.Max(Range("L:L"))
    Range("Q4") = MaxTotalVolume
    
    'Step 3 - Loop through tickers and compare individual values to calculated values
    'Reuse CellLoopRange and NumIterations variables from above
    Set CellLoopRange = Range("I2")
    Set CellLoopRange = Range(CellLoopRange, CellLoopRange.End(xlDown))
    NumIterations = CellLoopRange.Rows.Count
    'Step 3a - Set "found" variables to False to start - will use to exit loop if all tickers are found
    FoundMaxPercentChange = False
    FoundMinPercentChange = False
    FoundMaxTotalVolume = False
    'Step 3b - Begin loop
    For i = 1 To NumIterations
        CurrentPercentChange = Cells(i + 1, 11).Value
        CurrentTotalVolume = Cells(i + 1, 12).Value
        'Step 3c - Check for equivalence on summary values calculated in step 2
        If CurrentPercentChange = MaxPercentChange Then
            CurrentTicker = Cells(i + 1, 9).Value
            Range("P2").Value = CurrentTicker
            FoundMaxPercentChange = True
        End If
        
        If CurrentPercentChange = MinPercentChange Then
            CurrentTicker = Cells(i + 1, 9).Value
            Range("P3").Value = CurrentTicker
            FoundMinPercentChange = True
        End If
        
        If CurrentTotalVolume = MaxTotalVolume Then
            CurrentTicker = Cells(i + 1, 9).Value
            Range("P4").Value = CurrentTicker
            FoundMaxTotalVolume = True
        End If
        
        'Step 3d - check if all tickers are found, and if so, then stop processing
        If FoundMaxPercentChange = True And FoundMinPercentChange = True And FoundMaxTotalVolume = True Then
            Exit For
        End If
    'iterate to next cell for loop to populate summary table
    Next i
'iterate to next worksheet
Next j


'--------------------------------------------



'Turn back on screen updating upon macro completion
Application.ScreenUpdating = True

'Messagebox to alert that macro is complete
MsgBox ("Macro Complete!")

End Sub

