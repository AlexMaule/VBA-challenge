Attribute VB_Name = "StockMarketAnalysis"
Sub StockMarketAnalysis5()
    
    For Each ws In Worksheets
    
        Dim summaryTableRow As Integer
        Dim openPrice As Double
        Dim closePrice As Double
        Dim totalStock As LongLong
        Dim lastRow As Long
        Dim percentChange As Double
        Dim keepGoing As Boolean
        Dim priceDif As Double
        Dim ticker As String
    
        keepGoing = True
        lastRow = 1
        totalStock = 0
    
        Do While keepGoing

            If ws.Cells(lastRow, 1).Value = "" Then
                keepGoing = False
            Else
                lastRow = lastRow + 1
            End If
        
        Loop
        
        summaryTableRow = 2
        
        openPrice = ws.Cells(2, 3).Value
    
        ws.Range("I1:L1").Value = [{"Ticker","Yearly Change","Percent Change","Total Stock Volume"}]
    
        For i = 2 To lastRow
    
            totalStock = totalStock + ws.Cells(i, 7).Value
            
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        
                closePrice = ws.Cells(i, 6).Value
                priceDif = closePrice - openPrice
                percentChange = priceDif / openPrice
                ticker = ws.Cells(i, 1).Value
                
                ws.Range("I" & summaryTableRow).Value = ticker
                ws.Range("J" & summaryTableRow).Value = priceDif
                ws.Range("K" & summaryTableRow).Value = percentChange
                ws.Range("L" & summaryTableRow).Value = totalStock
                
                ws.Range("K" & summaryTableRow).NumberFormat = "0.00%"

                If priceDif < 0 Then
                    ws.Range("J" & summaryTableRow).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & summaryTableRow).Interior.ColorIndex = 4
                End If
                
                totalStock = 0
                summaryTableRow = summaryTableRow + 1
                openPrice = ws.Cells(i + 1, 3).Value
            
            End If
            
        Next i
        
        ws.Columns("I:L").AutoFit
        
    Next ws
    
End Sub

Sub Summary2()
    
    For Each ws In Worksheets
    
        Dim greaterIncrease As Double
        Dim greaterDecrease As Double
        Dim greaterTotalStock As LongLong
        Dim tempValue As Double
        Dim tempTotalStock As LongLong
        Dim lastRow As Long
        Dim keepGoing As Boolean
        Dim greaterIncreaseTicker As String
        Dim greaterDecreaseTicker As String
        Dim totalStockTicker As String
        
        greaterIncrease = 0
        greaterDecrease = 0
        greaterTotalStock = 0
    
        keepGoing = True
        lastRow = 1
        
        Do While keepGoing
            If ws.Cells(lastRow + 1, 11) = "" Then
                keepGoing = False
            Else
                lastRow = lastRow + 1
            End If
        Loop
        

        
        For i = 2 To lastRow
        
            tempValue = CDbl(ws.Cells(i, 11).Value)
            
            tempTotalStock = ws.Cells(i, 12).Value
            
            
            If tempValue > greaterIncrease Then
                greaterIncrease = tempValue
                greaterIncreaseTicker = ws.Range("I" & i).Value
            ElseIf tempValue < greaterDecrease Then
                greaterDecrease = tempValue
                greaterDecreaseTicker = ws.Range("I" & i).Value
            End If
            
            If tempTotalStock > greaterTotalStock Then
                greaterTotalStock = tempTotalStock
                totalStockTicker = ws.Range("I" & i).Value
            End If
            
        Next i
        
        ws.Range("P1:Q1").Value = [{"Ticker","Value"}]
        ws.Range("O2:O4").Value = Application.Transpose([{"Greatest % Increase","Greatest % Decrease","Greatest Total Volume"}])

        ws.Range("P2").Value = greaterIncreaseTicker
        ws.Range("Q2").Value = greaterIncrease
        ws.Range("Q2").NumberFormat = "0.00%"
        
        ws.Range("P3").Value = greaterDecreaseTicker
        ws.Range("Q3").Value = greaterDecrease
        ws.Range("Q3").NumberFormat = "0.00%"
        
        ws.Range("P4").Value = totalStockTicker
        ws.Range("Q4").Value = greaterTotalStock
        
        ws.Columns("O:Q").AutoFit
    
    Next ws
    
End Sub

