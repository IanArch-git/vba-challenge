Attribute VB_Name = "Module1"
Sub challenge():

'Define initial variables
Dim Ticker As String
Dim tOpen As Double
Dim tClose As Double
Dim ws As Worksheet
Dim tVol As Long
Dim Summary_Table_Row As Integer

'Found on Google to resolve Overflow error
On Error Resume Next


For Each ws In ThisWorkbook.Worksheets

'Set variables
Summary_Table_Row = 2
ws.Range("I1") = "Ticker"
ws.Range("J1") = "Annual Change"
ws.Range("K1") = "Annual Percentage Change"
ws.Range("L1") = "Total Volume"

Ticker = ws.Range("A2")
tOpen = ws.Range("C2")
tVol = ws.Range("G2")

'Loop through all data
For i = 2 To ws.Cells(Rows.Count, 1).End(xlDown).Row


    ' Check if we are still within the same Ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value And ws.Cells(i + 1, 1).Value <> "" Then

      'Set the variable
      tClose = ws.Cells(i, 6).Value
      tVol = tVol + ws.Cells(i, 7).Value
      
      'Print the variables in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Ticker
      ws.Range("J" & Summary_Table_Row).Value = tClose - tOpen
            
            If ws.Range("J" & Summary_Table_Row).Value > 0 Then
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            ElseIf ws.Range("J" & Summary_Table_Row).Value < 0 Then
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            ElseIf ws.Range("J" & Summary_Table_Row).Value = 0 Then
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 5
            End If
           
      ws.Range("K" & Summary_Table_Row).Value = (tClose - tOpen) / tOpen
      ws.Range("L" & Summary_Table_Row).Value = CLng(tVol)
      
      'Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      'Set variables
      Ticker = ws.Cells(i + 1, 1).Value
      tOpen = ws.Cells(i + 1, 3).Value
      tClose = ws.Cells(i + 1, 6).Value
      
      'Reset Volume
      tVol = 0
      
    Else
    tVol = tVol + ws.Cells(i + 1, 7).Value
    
    End If
    
Next i

Next ws

End Sub



