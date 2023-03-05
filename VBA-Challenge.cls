VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub vbahomework()

Dim volume As LongLong
Dim i As Long
Dim j As Integer
Dim openy As Double
Dim closey As Double
Dim ws As Worksheet
Dim ticker As String
Dim LastRow As Long
Dim gincrease As Double
Dim gdecrease As Double
Dim gvolume As LongLong





For Each ws In Worksheets
    ws.Activate
    'initialize variables and table headers
    volume = 0
    j = 2
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    openy = Cells(2, 3).Value
    ticker = Cells(2, 1).Value
    gincrease = 0
    gdecrease = 0
    gvolume = 0
            
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    
    
    
   
    
    For i = 2 To LastRow
    volume = (volume + Cells(i, 7).Value)
    
    If ticker = Cells(i + 1, 1).Value Then
    
    
    Else
        'print data for ticker
        Cells(j, 9).Value = Cells(i, 1).Value
        Cells(j, 10).Value = Cells(i, 6).Value - openy
        Cells(j, 11).Value = (Cells(i, 6).Value / openy) - 1
        Cells(j, 11).NumberFormat = "0.00%"
        Cells(j, 12).Value = volume
        
        'fill color into % if positive or negative
        If Cells(j, 11).Value < 0 Then
            Cells(j, 11).Interior.ColorIndex = 3
            Else
            Cells(j, 11).Interior.ColorIndex = 4
        End If
        
            
        'check for largest increase/decrease/volume
        If Cells(j, 11).Value > gincrease Then
            gincrease = Cells(j, 11).Value
            Cells(2, 16).Value = Cells(i, 1).Value
            Cells(2, 17).Value = Cells(j, 11).Value
            Cells(2, 17).NumberFormat = "0.00%"
        End If
        If Cells(j, 11).Value < gdecrease Then
            gdecrease = Cells(j, 11).Value
            Cells(3, 16).Value = Cells(i, 1).Value
            Cells(3, 17).Value = Cells(j, 11).Value
            Cells(3, 17).NumberFormat = "0.00%"
        End If
        If gvolume < volume Then
            gvolume = volume
            Cells(4, 17).Value = gvolume
            Cells(4, 16).Value = Cells(i, 1).Value
        End If
        
        'reset variables
        volume = 0
        j = j + 1
        openy = Cells(i + 1, 3).Value
        ticker = Cells(i + 1, 1)
    End If
    


    Next i

Next ws

End Sub
