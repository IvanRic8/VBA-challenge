# 02-VBA Homework Script

Sub stock():

    For Each ws In Worksheets
    
        Dim ticker As String
        Dim i As LongLong
        Dim i_sum As Integer
        Dim last_row As Long
        Dim cum_sum As LongLong
        Dim yearly_change As Double
        Dim open_value As Double
        Dim close_value As Double
        
        
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        i_sum = 2
        open_value = ws.Cells(2, 3).Value
    

        For i = 2 To last_row
            On Error Resume Next
            ticker = ws.Cells(i, 1).Value
            cum_sum = cum_sum + ws.Cells(i, 7).Value
        
            If ws.Cells(i + 1, 1).Value <> ticker Then
                close_value = ws.Cells(i, 6).Value
                ws.Cells(i_sum, 9).Value = ticker
                ws.Cells(i_sum, 10).Value = close_value - open_value
                yearly_change = ws.Cells(i_sum, 10).Value
                    
                    If yearly_change >= 0 Then
                        ws.Cells(i_sum, 10).Interior.ColorIndex = 4
                        Else
                        ws.Cells(i_sum, 10).Interior.ColorIndex = 3
                    End If
                    
                ws.Cells(i_sum, 11).Value = yearly_change / open_value
                open_value = ws.Cells(i + 1, 3).Value
                ws.Cells(i_sum, 12).Value = cum_sum
                i_sum = i_sum + 1
                cum_sum = 0
                
            End If
        
        Next i
            
    Next ws
        
End Sub