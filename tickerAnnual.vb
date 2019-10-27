' 1. Loop through all worksheets
' 2. For each ticker find yearly change, percent change, and total volume
' 3. Conditional format yearly change - positive: green, negative: red
' 4. Track "Greatest % increase", "Greatest % Decrease" and "Greatest total volume

Sub tickerAnnual()

    ' ticker ID's as string
    Dim tickerID, percentMaxID, percentMinID, volumeMaxID As String
    ' last row in worksheet counter
    Dim lastRow As Double

    ' set year as tag
    Dim yearTag, yearTagMax, yearTagMin, yearTagVol As Double
    
    ' summary table row counter
    Dim summaryRow, yearRow As Integer
    ' value of stock on open and close day of year
    Dim openBell, closeBell As Double
    Dim yearChange As Double
    Dim percentChange, percentChangeMax, percentChangeMin As Double
    Dim totalVolume, totalVolumeMax As Double
    
    Dim WorksheetName As String

    ' initialize variables
    yearRow = 1
    summaryRow = 2
    percentChange = 0
    totalVolumeMax = 0
    percentChangeMax = 0
    percentChangeMin = 0
    
    ' ---------------------------------------------------------------
    ' Code to add a sheet at the beginning of the workbook from 
    ' in class assignment solution to Wells Fargo Part 2
    ' --------------------------------------------------------------- 
    ' Add a sheet named "Combined Data"
    Sheets.Add.Name = "Combined_Data"
    'move created sheet to be first sheet
    Sheets("Combined_Data").Move Before:=Sheets(1)
    ' Specify the location of the combined sheet
    Set combined_sheet = Worksheets("Combined_Data")


    
    ' add ticker summary headers to first worksheet
    Sheets(1).Cells(1, 1).Value = "Year"
    Sheets(1).Cells(1, 2).Value = "Ticker"
    Sheets(1).Cells(1, 3).Value = "Yearly Change"
    Sheets(1).Cells(1, 4).Value = "% Change"
    Sheets(1).Cells(1, 5).Value = "Total Volume"

    ' add header for greatest increase, decrease, and volume by year statistics
    Sheets(1).Cells(yearRow, 7).Value = "Year"
    Sheets(1).Cells(yearRow, 8).Value = "Category"
    Sheets(1).Cells(yearRow, 9).Value = "Ticker"
    Sheets(1).Cells(yearRow, 10).Value = "Value"
    

    ' Loop through worksheets
    ' ------------------------------
    For Each ws In Worksheets
        ' Grabbed the WorksheetName
        WorksheetName = ws.Name
        If WorksheetName <> "Combined_Data" Then
        
            ' Initialize open for first ticker in worksheet
            tickerID = ws.Cells(2, 1).Value
            openBell = ws.Cells(2, 3).Value
            yearTag = Left(ws.Cells(2, 2).Value, 4)
            ' MsgBox ("Year is " & yearTag)

            ' Find the last row in the worksheet
            lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            ' MsgBox ("Last Row is " & lastRow)
        
            ' Loop through each line on the sheet
            For i = 2 To lastRow
                ' calculate total volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value
 
                ' Find where ticker changes
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    closeBell = ws.Cells(i, 6).Value
                
                    ' calculate the percent change if non-zero
                    yearChange = closeBell - openBell
                
                    If openBell <> 0 Then
                      percentChange = yearChange / openBell
                    End If
                
                    ' assign data to summary table
                    Sheets(1).Cells(summaryRow, 1).Value = yearTag
                    Sheets(1).Cells(summaryRow, 2).Value = tickerID
                    Sheets(1).Cells(summaryRow, 3).Value = yearChange
                
                    ' Color the year change red if at or below zero, green = positive growth
                    If (yearChange <= 0) Then
                        Sheets(1).Cells(summaryRow, 3).Interior.ColorIndex = 3
                    Else
                        Sheets(1).Cells(summaryRow, 3).Interior.ColorIndex = 4
                    End If
                    ' assign percent change and format in the next column
                    
                    Sheets(1).Cells(summaryRow, 4) = Format(percentChange, "Percent")
        
                    ' assign total volume in the next column
                    Sheets(1).Cells(summaryRow, 5).Value = totalVolume

                    ' Check for max or min values
                    If (totalVolume > totalVolumeMax) Then
                        volumeMaxID = tickerID
                        totalVolumeMax = totalVolume
                        yearTagVol = yearTag
                        'MsgBox ("Total Volume Maximum is " & totalVolumeMax)
                    End If

                    If (percentChange > percentChangeMax) Then
                        percentMaxID = tickerID
                        percentChangeMax = percentChange
                        yearTagMax = yearTag
                    ElseIf (percentChange < percentChangeMin) Then
                        percentMinID = tickerID
                        percentChangeMin = percentChange
                        yearTagMin = yearTag
                    End If

                    summaryRow = summaryRow + 1
                    tickerID = ws.Cells(i + 1, 1).Value
                    openBell = ws.Cells(i + 1, 3).Value
                    percentChange = 0
                    totalVolume = 0

                End If
            Next i

            ' Next ws  ' comment out for annual ticker data

            ' report greatest increase, decrease, and greatest total volume statistics
            Sheets(1).Cells(yearRow + 1, 8).Value = "Greatest % Increase"
            Sheets(1).Cells(yearRow + 1, 7).Value = yearTagMax
            Sheets(1).Cells(yearRow + 1, 9).Value = percentMaxID
            Sheets(1).Cells(yearRow + 1, 10).Value = Format(percentChangeMax, "Percent")
    
            Sheets(1).Cells(yearRow + 2, 8).Value = "Greatest % Decrease"
            Sheets(1).Cells(yearRow + 2, 7).Value = yearTagMin
            Sheets(1).Cells(yearRow + 2, 9).Value = percentMinID
            Sheets(1).Cells(yearRow + 2, 10).Value = Format(percentChangeMin, "Percent")
    
            Sheets(1).Cells(yearRow + 3, 8).Value = "Greatest Total Volume"
            Sheets(1).Cells(yearRow + 3, 7).Value = yearTagVol
            Sheets(1).Cells(yearRow + 3, 9).Value = volumeMaxID
            Sheets(1).Cells(yearRow + 3, 10).Value = totalVolumeMax

' comment out from here down to *** for alpha data
            ' increment yearRow by 3 for next ws
            yearRow = yearRow + 3
            ' re-set statistics for the next year of data
            percentChange = 0
            totalVolumeMax = 0
            percentChangeMax = 0
            percentChangeMin = 0

        End If
    Next ws
' *** -> comment out to here for alpha data

    ' auto format column width for the combined data
    Sheets(1).Columns("A:J").Select
    Selection.EntireColumn.AutoFit
 
End Sub