Attribute VB_Name = "Module1"
Sub StockAnalysis()
Application.ScreenUpdating = False

Dim Ticker As String
Dim iRow As Long
Dim iStart As Double
Dim TickerTotal As Double
Dim ws As Worksheet
Dim IncreaseTicker As String, IncreaseValue As Double
Dim DecreaseTicker As String, DecreaseValue As Double
Dim VolumeTicker As String, VolumeValue As Double


For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
    Range("J:M").Clear
    
'Sort all of the Data for the year
'This helps considerable to give peace of mind
'when making assumptions later

    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:=Range("A:A"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveSheet.Sort.SortFields.Add Key:=Range("B:B"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range("A:G")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

'Copy the first column into column J
    Range("A:A").Copy Range("J:J")
'Remove Duplicates in column J
    ActiveSheet.Range("J:J").RemoveDuplicates Columns:=1, Header:=xlYes
    
'Create headers for remaining columns
    Range("K1").Value = "Yearly Change"
    Range("L1").Value = "Percent Change"
    Range("M1").Value = "Total Stock Volume"
'and for the other table
    Range("O1").Value = "Metric"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"

'Setting initial values for bonus table
    IncreaseValue = 0
    DecreaseValue = 0
    VolumeValue = 0
    
'The macro will loop through the stock and collect data to fill the table.
    Range("J2").Activate
    iRow = 2
    Do While ActiveCell.Value <> ""
        Ticker = ActiveCell.Value
        iStart = Cells(iRow, 3).Value
        TickerTotal = 0
'This loop goes through the rows of data
'where the ticker is the same as the active one in table
        Do While Cells(iRow, 1).Value = Ticker
            TickerTotal = TickerTotal + Cells(iRow, 7)
            iRow = iRow + 1
        Loop
        
        With ActiveCell.Offset(0, 1)
            .Value = Cells(iRow - 1, 6).Value - iStart
            .Font.Bold = True
        End With
        
        If ActiveCell.Offset(0, 1).Value <= 0 Then
            ActiveCell.Offset(0, 1).Interior.ColorIndex = 3
        Else
            ActiveCell.Offset(0, 1).Interior.ColorIndex = 4
        End If
        
        If iStart = 0 Then
            With ActiveCell.Offset(0, 2)
                .Value = 0
                .Style = "Percent"
            End With
        Else
            With ActiveCell.Offset(0, 2)
                .Value = ActiveCell.Offset(0, 1).Value / iStart
                .Style = "Percent"
            End With
        End If
        
        ActiveCell.Offset(0, 3).Value = TickerTotal
    
'Check to see if the current ticker satisfies any of the bonus data
        If ActiveCell.Offset(0, 2).Value < DecreaseValue Then
            DecreaseValue = ActiveCell.Offset(0, 2).Value
            DecreaseTicker = Ticker
        End If
        If ActiveCell.Offset(0, 2).Value > IncreaseValue Then
            IncreaseValue = ActiveCell.Offset(0, 2).Value
            IncreaseTicker = Ticker
        End If
        If ActiveCell.Offset(0, 3).Value > VolumeValue Then
            VolumeValue = ActiveCell.Offset(0, 3).Value
            VolumeTicker = Ticker
        End If
    
        ActiveCell.Offset(1, 0).Activate
    Loop


'Assigning final values and formats to bonus table
    Range("P2").Value = IncreaseTicker
    With Range("Q2")
        .Value = IncreaseValue
        .Style = "Percent"
    End With
    
    Range("P3").Value = DecreaseTicker
    With Range("Q3")
        .Value = DecreaseValue
        .Style = "Percent"
    End With
    
    Range("P4").Value = VolumeTicker
    Range("Q4").Value = VolumeValue
    
    
'Fitting these columns after data is added
    Columns("J:M").AutoFit
    Columns("O:Q").AutoFit
    
    Range("A1").Activate

Next '''Ends the loop through the sheets

Application.ScreenUpdating = True
End Sub


Sub Resetti()

Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets
ws.Activate
Range("J:Z").Clear
Next
End Sub
