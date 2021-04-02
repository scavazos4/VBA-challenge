VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub tickers()


Dim ticker1 As String
Dim tickerSave1 As String
Dim tickerSave2 As String
Dim tickerSave3 As String
Dim ticker2 As String
Dim marketOpen As Double
Dim marketClose As Double
Dim volume As Double
Dim volumeSum As Double
Dim iSave As Double
Dim iSave2 As Double

Dim maxVolume As Double
Dim increase1 As Double
Dim increase2 As Double
Dim increaseSave As Double
Dim decrease1 As Double
Dim decrease2 As Double
Dim decreaseSave As Double
Dim volume1 As Double
Dim volume2 As Double

Dim ws As Worksheet

For Each ws In Sheets
    ws.Cells(1, 9) = "ticker"
    ws.Cells(1, 10) = "yearly_change"
    ws.Cells(1, 11) = "percent_change"
    ws.Cells(1, 12) = "Total_volume"
    
    ws.Cells(1, 16) = "ticker"
    ws.Cells(1, 17) = "Value"
    ws.Cells(2, 15) = "greatest_percent_increase"
    ws.Cells(3, 15) = "greatest_percent_decrease"
    ws.Cells(4, 15) = "greates_total_Volume"
    
    For n = 800000 To 2 Step -1
        If ws.Cells(n, 3) = "0" Then
        ws.Rows(n).EntireRow.Delete
        End If
    Next n
         
    iSave = 2

    For j = 2 To 5000
    
    marketOpen = ws.Cells(iSave, 3).Value
    volumeSum = 0
    
    
        For I = iSave To iSave + 365
            ticker1 = ws.Cells(I, 1)
            ticker2 = ws.Cells(I + 1, 1)
            volume = ws.Cells(I, 7).Value
            volumeSum = volumeSum + volume

            ' if new ticker then submit compiled data
            If ticker1 <> ticker2 Then
            marketClose = ws.Cells(I, 6).Value
    
            ' output displaying values in spreadsheet
            
            ws.Cells(j, 9).Value = ticker1
            ws.Cells(j, 10).Value = marketClose - marketOpen
            ws.Cells(j, 11).Value = ((marketClose - marketOpen) / marketOpen)
            
            ws.Cells(j, 12).Value = volumeSum
        
            ' reset values for new set
            ticker1 = ws.Cells(I + 1, 1).Value
            marketOpen = ws.Cells(I + 1, 3).Value
            iSave = I + 1
            Exit For
            End If

    Next I
    
   If ws.Cells(j, 11).Value > 0 Then
        ws.Cells(j, 11).Interior.ColorIndex = 4
    ElseIf ws.Cells(j, 11).Value = 0 Then
        ws.Cells(j, 11).Interior.ColorIndex = 2
    Else
        ws.Cells(j, 11).Interior.ColorIndex = 3
    End If

Next j

maxVolume = 0
For m = 2 To 3500
    
    
    volume1 = ws.Cells(m, 11).Value
    volume2 = ws.Cells(m + 1, 12).Value
    ticker2 = ws.Cells(m + 1, 9).Value
        
    If volume2 > maxVolume Then
        maxVolume = volume2
        tickerSave3 = ticker2
    End If
             
Next m

increaseSave = 0
For k = 2 To 3500
    
   ticker1 = ws.Cells(k, 9).Value
   increase1 = ws.Cells(k, 11).Value
   
        
    If increase1 > increaseSave Then
        increaseSave = increase1
        tickerSave1 = ticker1
    End If
    
Next k

decreaseSave = 0
For l = 2 To 3500
    
    ticker2 = ws.Cells(l, 9).Value
    decrease1 = ws.Cells(l, 11).Value
    
        
    If decrease1 < decreaseSave Then
        decreaseSave = decrease1
        tickerSave2 = ticker2
    End If
    
               
Next l


    ws.Cells(2, 16).Value = tickerSave1
    ws.Cells(2, 17).Value = increaseSave
    ws.Cells(3, 16).Value = tickerSave2
    ws.Cells(3, 17).Value = decreaseSave
    ws.Cells(4, 16).Value = tickerSave3
    ws.Cells(4, 17).Value = maxVolume
    

Next ws


     
End Sub



