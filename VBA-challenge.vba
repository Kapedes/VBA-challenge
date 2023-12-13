Attribute VB_Name = "Module1"

Sub Ticker()

Dim WS_Count As Integer
Dim j As Integer
Dim i As Long
Dim k As Single

Dim Lookup_Value As Double
Dim Lookup_Value2 As Double
Dim Lookup_Array As Range
Dim Return_Array As Range
Dim If_Not_Found As String
Dim Result As Variant
Dim Result2 As Variant
Dim Result3 As Variant
         
         ' The following was taken from this link
         
        'https://support.microsoft.com/en-au/topic/macro-to-loop-through-all-worksheets-in-a-workbook-feef14e3-97cf-00e2-538b-5da40186e2b0
         ' Set WS_Count equal to the number of worksheets in the active workbook.
         WS_Count = ActiveWorkbook.Worksheets.Count
             MsgBox (WS_Count)
         ' Begin the loop.
         For j = 1 To WS_Count
        
        'End of code from link
        
'Go to next worksheet
 Worksheets(j).Activate

Dim TickerName As String

' Set a variable for holding the stock code total

Dim Ticker_total As Double
Ticker_total = 0

' Create a dim to record the stock code in the adjacent table.
Dim Ticker_Code_Row As Integer
Ticker_Code_Row = 2


' Define the headers
    Range("I1") = "Ticker"
    Range("J1") = "Yearly change"
    Range("K1") = "Percentage change"
    Range("L1") = "Total Stock Volume"
    
    Range("P1") = "Ticker"
    Range("Q1") = "Value"
    Range("O2") = "Greatest % increase"
    Range("O3") = "Greatest % decrease"
    Range("O4") = "Greatest total volume"

' Determine last row (lr)

   lr = Cells(Rows.Count, 1).End(xlUp).Row
    
'Loop through all stock codes

For i = 2 To lr

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        ' Set the ticker name
        TickerName = Cells(i, 1).Value
    
        'Add to the volume total
        Volume_total = Volume_total + Cells(i, 7).Value
    
        'Print the ticker in the summary table
        Range("I" & Ticker_Code_Row).Value = TickerName
    
        'Print the Volume total in the summary table
        Range("L" & Ticker_Code_Row).Value = Volume_total
    
    
    
        'Calculate the closing value
    
         Closing_value = Cells(i, 6).Value
    
        'Calculate yearly change and percentage change
    
        Yearly_change = Closing_value - Opening_Value
    
        Percentage_change = (Closing_value - Opening_Value) / (Opening_Value)
    
    
        'Print the yearly change
        Range("J" & Ticker_Code_Row).Value = Yearly_change
    
    
        'Print the Percentage change
        Range("K" & Ticker_Code_Row).Value = FormatPercent((Percentage_change), , , , vbFalse)
    
        ' Conditional formatting

                If Cells(Ticker_Code_Row, 10).Value > 0 Then
                 Cells(Ticker_Code_Row, 10).Interior.ColorIndex = 4
                ElseIf Cells(Ticker_Code_Row, 10).Value < 0 Then
                  Cells(Ticker_Code_Row, 10).Interior.ColorIndex = 3
        
                End If
        
        'Add one to the Ticker_code_row
        Ticker_Code_Row = Ticker_Code_Row + 1
    
        'Reset Volume total
        Volume_total = 0
    
        'Reset both opening value and closing value
        Opening_Value = 0
        Closing_value = 0
    
    Else

    
    'Add to the Volume Total
    Volume_total = Volume_total + Cells(i, 7).Value
    
        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
        'Take the Opening value
        Opening_Value = Cells(i, 3).Value

        End If
    End If
'Go to next i
Next i






' Define yearly change header

Range("J1") = "Yearly change"
    
'Calculate max value in range
maxPcnt = WorksheetFunction.Max(Range("K:K"))
minPcnt = WorksheetFunction.Min(Range("K:K"))
maxVolume = WorksheetFunction.Max(Range("L:L"))

'Print the values in their
Range("Q2").Value = FormatPercent((maxPcnt), , , , vbFalse)
Range("Q3").Value = FormatPercent((minPcnt), , , , vbFalse)
Range("Q4").Value = maxVolume

'Xlookup for MaxPct
  Lookup_Value = Range("Q2").Value
  Set Lookup_Array1 = Range("K:K")
  Set Return_Array2 = Range("I:I")
  If_Not_Found = "N/A"
  
  
'Xlookup for MaxPct
  Lookup_Value2 = Range("Q3").Value
  If_Not_Found = "N/A"
  
  
'Xlookup for MaxPct
  Lookup_Value3 = maxVolume
  Set Lookup_Array3 = Range("L:L")
  Set Return_Array4 = Range("I:I")
  If_Not_Found = "N/A"

'Perform XLOOKUP and Store To Variable
    Result = Application.WorksheetFunction.Xlookup( _
      Lookup_Value, _
      Lookup_Array1, _
      Return_Array2, _
      If_Not_Found)
  On Error GoTo 0

'Perform XLOOKUP and Store To Variable
    Result2 = Application.WorksheetFunction.Xlookup( _
      Lookup_Value2, _
      Lookup_Array1, _
      Return_Array2, _
      If_Not_Found)
  On Error GoTo 0

'Perform XLOOKUP and Store To Variable
    Result3 = Application.WorksheetFunction.Xlookup( _
      Lookup_Value3, _
      Lookup_Array3, _
      Return_Array4, _
      If_Not_Found)
  On Error GoTo 0
  
'Print the values in their
Range("P2").Value = Result
Range("P3").Value = Result2
Range("P4").Value = Result3

'Autofit the column width
Columns("A:L").Autofit  
  
  
Next j
  
  
 
MsgBox ("Complete")
  

End Sub

