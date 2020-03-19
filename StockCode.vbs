'Step 1: Loop through each spreadsheet and add Headers - Done

'Step 2: Add the headers to each spreadsheet - Done

'Step 3: Summarize the Tickers on each page - Done

'Step 4: Add totals by Ticket on each page - Done

'Last Close Date $ - First Open Date $

'(Last Close Date $ - First Open Date $ / First Open Date $) * 100

Sub ABC_test()
'Variables
Dim Ticker As String
Dim Volume As Double
Dim Open_Cost As Double
Dim Close_Cost As Double
Dim YearChange As Double
Dim Percentage_Change As Double
Dim check As Double
Dim date_value As Double
Dim year As String
Dim max_date As Double
Dim x As Double


'Loop through each workbook

For Each ws In Worksheets

    Dim WorksheetName As String

    WorksheetName = ws.Name
    'MsgBox (WorksheetName)
    'MsgBox (WorksheetName & "0101")
    
    'Add header to each workbook
    'ws.Cells(1, 9).Value = "Ticker"
    Worksheets(WorksheetName).Range("I1").Value = "Ticker"
    Worksheets(WorksheetName).Range("J1").Value = "Yearly Change"
    Worksheets(WorksheetName).Range("K1").Value = "Percentage Change"
    Worksheets(WorksheetName).Range("L1").Value = "Total Stock Volume"
    
    'Get the year from the date column
  'date_value = 20161231
   year = Left(Worksheets(WorksheetName).Range("B2").Value, 4)
 ' MsgBox (year)
  date_value = Int(year & 1231)
  ' MsgBox (date_value)
    
    
    'Last Field in the row
    lastrow1 = Worksheets(WorksheetName).Cells(Rows.Count, 1).End(xlUp).Row
    ' MsgBox (lastrow1)
    
        'Loop through the data to find the total amount of Volume & other caculations
        For i = 2 To lastrow1
    
        
        If Worksheets(WorksheetName).Cells(i + 1, 1).Value <> Worksheets(WorksheetName).Cells(i, 1).Value Then
        
            
            'Ticker name captured
             Ticker = Worksheets(WorksheetName).Cells(i, 1).Value
        
             'Find last row in the ticker name field
             lastrow2 = Worksheets(WorksheetName).Cells(Rows.Count, 9).End(xlUp).Row
             
             'Checking the last row field and display on screen
            ' Worksheets(WorksheetName).Range("H" & lastrow2 + 1).Value = lastrow2
             
             'Append Ticker to sheet
             Worksheets(WorksheetName).Range("I" & lastrow2 + 1).Value = Ticker
           
            'Add Last volume amount to the summarized totals
             Volume = Worksheets(WorksheetName).Cells(i, 7).Value + Volume
             
             'Append Volume Sum to sheet
             Worksheets(WorksheetName).Range("L" & lastrow2 + 1).Value = Volume
             
           ' MsgBox ("B" & (i - x) & ":B" & i)
             'max_date = Application.WorksheetFunction.Max(Columns("B"))
             max_date = Application.WorksheetFunction.Max(Range("B" & (i - x) & ":B" & i))
               ' MsgBox max_date
             
             'Finding the end date
             If Worksheets(WorksheetName).Cells(i, 2).Value >= max_date Then
               ' MsgBox (Worksheets(WorksheetName).Cells(i, 2).Value)
                Close_Cost = Worksheets(WorksheetName).Cells(i, 6).Value
                'MsgBox (Close_Cost)
             End If
            
             
             'Caculation for Yearly Change
             
             YearChange = Close_Cost - Open_Cost
             Worksheets(WorksheetName).Range("J" & lastrow2 + 1).Value = YearChange
             
             'Worksheets(WorksheetName).Range("M" & lastrow2 + 1).Value = Close_Cost
             'Worksheets(WorksheetName).Range("N" & lastrow2 + 1).Value = Open_Cost
             
             'Caculation for Percentage
             If Open_Cost <> 0 Then
                Percentage_Change = YearChange / Open_Cost
                Worksheets(WorksheetName).Range("K" & lastrow2 + 1).Value = FormatPercent(Percentage_Change, , , vbTrue)
             Else
                 Worksheets(WorksheetName).Range("K" & lastrow2 + 1).Value = 0
             End If
             
             'Put Color if negative = Red(3) and Positive = Green (4)
             If YearChange < 0 Then
                Worksheets(WorksheetName).Range("J" & lastrow2 + 1).Interior.ColorIndex = 3
             Else
                Worksheets(WorksheetName).Range("J" & lastrow2 + 1).Interior.ColorIndex = 4
             End If
            
                
             
             'Zero out the volume for next Ticker
             Volume = 0
             YearChange = 0
             Close_Cost = 0
             Open_Cost = 0
             Percentage_Change = 0
             date_value = 20161231
             x = 0
        
        Else
            Volume = Volume + Worksheets(WorksheetName).Cells(i, 7).Value
            'format(date, "")
            If Volume = 113200 Then
           ' MsgBox (Volume)
            'MsgBox (Worksheets(WorksheetName).Cells(i, 1).Value)
            
            End If
            
            
            'date_value = Worksheets(WorksheetName).Cells(i, 2).Value
            
            If Worksheets(WorksheetName).Cells(i, 2).Value < date_value Then
               ' MsgBox (Worksheets(WorksheetName).Cells(i, 2).Value & " " & date_value)
              Open_Cost = Worksheets(WorksheetName).Cells(i, 3).Value
              'MsgBox (Open_Cost)
                
                
           ' ElseIf Worksheets(WorksheetName).Cells(i, 2).Value = 0 + Open_Cost Then
            '    check = Worksheets(WorksheetName).Cells(i, 3).Value
            End If
            
            date_value = Worksheets(WorksheetName).Cells(i, 2).Value
           ' MsgBox (date_value)
           x = x + 1
        
        End If

    Next i
    


Next ws



End Sub

