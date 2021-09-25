Attribute VB_Name = "Module1"
Option Explicit

Sub Run()

    'Vars
    Dim curTick As String
    Dim nextTick As String
    Dim totalVol As Variant
    Dim curVol
    Dim groupNum As Long
    Dim i As Long
    Dim lr As Long
    Dim curSh As String
    Dim curD As Long
    Dim nextD As Long
    Dim maxD As Long
    Dim mOpen As Double
    Dim mClose As Double
    Dim yChange As Double
    Dim greatInc As Double
    Dim greatIncT As String
    Dim greatDec As Double
    Dim greatDecT As String
    Dim greatVol As Variant
    Dim greatVolT As String
    
    'Default Vals
    curSh = "2014"
    totalVol = 0
    groupNum = 1
    curD = 0
    nextD = 0
    mOpen = Sheets(curSh).Cells(2, 3).Value
    lr = Sheets(curSh).Cells(Rows.Count, 1).End(xlUp).Row
    greatInc = 0
    greatDec = 0
    greatVol = 0
    
    'New Headers
    Sheets(curSh).Range("I1").Value = "Tickers"
    Sheets(curSh).Range("J1").Value = "Yearly Change"
    Sheets(curSh).Range("K1").Value = "Percent Change"
    Sheets(curSh).Range("L1").Value = "Total Stock Volume"
    
    Sheets(curSh).Range("N2").Value = "Greatest % Increase"
    Sheets(curSh).Range("N3").Value = "Greatest % Decrease"
    Sheets(curSh).Range("N4").Value = "Greatest Total Volume"
    Sheets(curSh).Range("O1").Value = "Ticker"
    Sheets(curSh).Range("P1").Value = "Value"
    
    'Iterate Through Sheet
    For i = 2 To lr
    
        'Assign Dynamic Values
        curTick = Sheets(curSh).Cells(i, 1).Value
        nextTick = Sheets(curSh).Cells(i + 1, 1).Value
        curD = Sheets(curSh).Cells(i, 2).Value
        nextD = Sheets(curSh).Cells(i + 1, 2).Value
        maxD = Sheets(curSh).Cells(i, 2).Value
        curVol = CLng(Sheets(curSh).Cells(i, 7).Value)
        
        'If Next Ticker and Current Ticker are Same
        If (nextTick = curTick) Then
            
            totalVol = totalVol + curVol
            
        Else
            
            'If Next Ticker is Different
            totalVol = totalVol + curVol
            
            'If Current Date is Greater than or Equal to Max, Change Max Date to Current
            If (curD >= maxD) Then
                maxD = curD
                mClose = Sheets(curSh).Cells(i, 6).Value
            End If
            
            'Preliminary Calculations
            groupNum = groupNum + 1
            yChange = WorksheetFunction.Sum(mClose, -(mOpen))
            
            'Assign Values to Headers
            Sheets(curSh).Cells(groupNum, 9).Value = curTick
            Sheets(curSh).Cells(groupNum, 10).Value = yChange
            Sheets(curSh).Cells(groupNum, 12).Value = totalVol
            
            'If Opening Value is Not 0 Then
            If (mOpen <> 0) Then
                Sheets(curSh).Cells(groupNum, 11).Value = Sheets(curSh).Cells(groupNum, 10).Value / mOpen
            Else
                Sheets(curSh).Cells(groupNum, 11).Value = 0
            End If
            
            'Check for Greatest % Increase
            If (Sheets(curSh).Cells(groupNum, 11).Value > greatInc) Then
                greatInc = Sheets(curSh).Cells(groupNum, 11).Value
                greatIncT = Sheets(curSh).Cells(groupNum, 9).Value
            End If
            
            'Check for Greatest % Decrease
            If (Sheets(curSh).Cells(groupNum, 11).Value < greatDec) Then
                greatDec = Sheets(curSh).Cells(groupNum, 11).Value
                greatDecT = Sheets(curSh).Cells(groupNum, 9).Value
            End If
            
            'Check for Greatest Total Volume
            If (totalVol > greatVol) Then
                greatVol = totalVol
                greatVolT = Sheets(curSh).Cells(groupNum, 9).Value
            End If
            
            'Formatting Cells
            If (yChange > 0) Then
                Sheets(curSh).Cells(groupNum, 10).Interior.Color = RGB(0, 255, 0)
            ElseIf (yChange <= 0) Then
                Sheets(curSh).Cells(groupNum, 10).Interior.Color = RGB(255, 0, 0)
            End If
            
            Sheets(curSh).Cells(groupNum, 11).NumberFormat = "0.00%"
            
            'Assign Next Ticker Opening Price
            mOpen = Sheets(curSh).Cells(i + 1, 3).Value
    
            'Reset Total Stock Volume
            totalVol = 0
    
        End If
    
    Next i
    
    'Assign Stats Results
    Sheets(curSh).Range("O2").Value = greatIncT
    Sheets(curSh).Range("P2").Value = greatInc
    Sheets(curSh).Range("O3").Value = greatDecT
    Sheets(curSh).Range("P3").Value = greatDec
    Sheets(curSh).Range("O4").Value = greatVolT
    Sheets(curSh).Range("P4").Value = greatVol
    
    'Format Percentages
    Sheets(curSh).Range("P2:P3").NumberFormat = "0.00%"
    
    'Complete Message
    MsgBox ("Macro Completed Successfully!")

End Sub
Sub Clear()
    Sheets("2014").Range("I1:P" & Sheets("2014").Cells(Rows.Count, 9).End(xlUp).Row).Clear
    MsgBox ("Cleared")
End Sub
Sub ClearAll()
    Dim ws As Integer
    For ws = 1 To ActiveWorkbook.Sheets.Count
        Sheets(ws).Range("I1:P" & Sheets(ws).Cells(Rows.Count, 9).End(xlUp).Row).Clear
    Next ws
    MsgBox ("All Sheets Cleared!")
End Sub
