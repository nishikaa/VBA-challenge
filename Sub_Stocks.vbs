Attribute VB_Name = "Module1"
Sub Sub_Stocks():
    For Each ws In Worksheets
        ws.Cells(1, 10).Value = "Ticker"
        ws.Cells(1, 11).Value = "Yearly Change"
        ws.Cells(1, 12).Value = "Percent Change"
        ws.Cells(1, 13).Value = "Total Stcok Volume"
        
        Dim Name As String
        Dim ChangeO As Double
        Dim CahngeC As Double
        Dim PChange As Double
        Dim Total As Double
        Dim Summary_Table_Row As Integer
        Dim Zero_open As String
        Zero_open = "NA"
        Summary_Table_Row = 2
       
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Total = 0
        ChangeO = ws.Cells(2, 3).Value
        For i = 2 To LastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Name = ws.Cells(i, 1).Value
                ws.Range("J" & Summary_Table_Row).Value = Name
                ChangeC = ws.Cells(i, 6).Value
                Total = Total + ws.Cells(i, 7).Value
                ws.Range("K" & Summary_Table_Row).Value = (ChangeC - ChangeO)
                ws.Range("M" & Summary_Table_Row).Value = Total
                
             'since some Tickers have zero opening value I will update the % increase as NA for them
                If ChangeO <> 0 Then
                    ws.Range("L" & Summary_Table_Row).Value = FormatPercent(((ChangeC - ChangeO) / ChangeO), 0)
                Else
                    
                    ws.Range("L" & Summary_Table_Row).Value = Zero_open
                End If
                
                Summary_Table_Row = Summary_Table_Row + 1
                ChangeO = ws.Cells(i + 1, 3).Value
                Total = 0
            Else
                Total = Total + ws.Cells(i, 7).Value
            End If
        Next i
        
        LastRowK = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
        For j = 2 To LastRowK
            If ws.Cells(j, 11).Value >= 0 Then
                ws.Cells(j, 11).Interior.ColorIndex = 4
            Else
                ws.Cells(j, 11).Interior.ColorIndex = 3
            End If
        Next j
        
    'Bonus code :)
    
     Dim Inname As String
     Dim Dname As String
     Dim Vname As String
     Dim Incr As Double
     Dim Decr As Double
     Dim Vol As Double
    
     LastRowK = ws.Cells(Rows.Count, 12).End(xlUp).Row
    
     ws.Cells(1, 16).Value = "Ticker"
     ws.Cells(1, 17).Value = "Value"
     ws.Cells(2, 15).Value = "Greatest % Increase"
     ws.Cells(3, 15).Value = "Greatest % Decrease"
     ws.Cells(4, 15).Value = "Greatest Total Volume"
    
     Inname = ""
     Dname = ""
     Vname = ""
    
     Incr = 0
     Decr = 0
     Vol = 0
        
     For k = 2 To LastRowK
        If ws.Cells(k, 12).Value = "NA" Then
            Incr = Incr
            Inname = Inname
        Else
            If ws.Cells(k, 12).Value > Incr Then
                Incr = ws.Cells(k, 12).Value
                Inname = ws.Cells(k, 10)
            End If
        End If
     Next k
            
     For l = 2 To LastRowK
        If ws.Cells(l, 12).Value = "NA" Then
            Decr = Decr
            Dname = Dname
            
        Else
            If ws.Cells(l, 12).Value < Decr Then
                Decr = ws.Cells(l, 12).Value
                Dname = ws.Cells(l, 10)
            End If
        End If
        
     Next l
     
     For m = 2 To LastRowK
        
        If ws.Cells(m, 13).Value > Vol Then
            Vol = ws.Cells(m, 13).Value
            Vname = ws.Cells(m, 10)
        End If
     Next m
                
            
         ws.Cells(2, 16).Value = Inname
         ws.Cells(3, 16).Value = Dname
         ws.Cells(4, 16).Value = Vname
        
         ws.Cells(2, 17).Value = Format(Incr, "0.00%")
         ws.Cells(3, 17).Value = Format(Decr, "0.00%")
         ws.Cells(4, 17).Value = Vol
     
       
        
    Next ws
End Sub


