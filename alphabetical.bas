Attribute VB_Name = "Module1"
Sub alphabetical()

'Create the dimension and set a property. First Part
    Dim ticker As String
    Dim vol As Double
    Dim ws As Worksheet
    
    
' This code will perform the same loop in all the sheets
 For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
        
'Create the dimensions and sets a property. Second Part

    Dim Summary_Table_Row As Integer
    Dim year_open As Double
    Dim year_close As Double
    
'Creates the label in the cell referente. Example (in Cells(1, 9) which is ("I1") set the word or .Value = "ticker"

    Cells(1, 9).Value = "ticker"
    Cells(1, 10).Value = "Yearly_change"
    Cells(1, 12).Value = "Total Stock Vol"
    Cells(1, 11).Value = "Yearly_percentage"

    
'declare the start number for this variable. The counter will begin in 2. Basically we have set a variable 2 for Summary_Table_Row
    
    Summary_Table_Row = 2
    
'for loop from 2 to the last row of the collumn. This will help if the other ws do not have the same number of rows. This is the code for last row Cells(Rows.Count, 1).End(xlUp).Row
  
      For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
    
'the if statement says that is the cell above the
      If Cells(i - 1, 1) = Cells(i, 1) And Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
          year_close = Cells(i, 6).Value
          yearly_change = year_close - year_open

          ticker = Cells(i, 1).Value

          vol = vol + Cells(i, 7).Value

'Displays the values in excel for ticker, vol, etc...

          Range("j" & Summary_Table_Row).Value = yearly_change

          Range("I" & Summary_Table_Row).Value = ticker

          Range("K" & Summary_Table_Row).Value = year_percent

          Range("L" & Summary_Table_Row).Value = vol

          Summary_Table_Row = Summary_Table_Row + 1

          vol = 0

      Else

          vol = vol + Cells(i, 7).Value

      End If

    Next i
    
'When kicks out of the loop above go to the next worksheet
Next ws
    

End Sub



