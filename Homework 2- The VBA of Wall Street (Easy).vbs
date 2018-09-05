Attribute VB_Name = "Module1"
Sub MultipleWorksheets()

    'Define variables
    Dim ws As Worksheet
    
    'Make the code run faster
    Application.ScreenUpdating = False
    
    'Loop through each worksheet
    For Each ws In Worksheets
        
        'Select each worksheet
        ws.Select
        
        'Run the VBAHomeworkEasy sub procedure
        Call VBAHomeworkEasy
    
    Next
    
    'Set Application.ScreenUpdating back to normal
    Application.ScreenUpdating = True

End Sub
Sub VBAHomeworkEasy()

'Define variables
Dim Ticker As String
Dim Volume As Double
Dim Summary_Table_Row As Integer

'Initialize variables
Volume = 0
Summary_Table_Row = 2
    
'Loop through stock data
For i = 2 To 797711

'Check for the same stock market ticker
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    'Define where the ticker name is on the worksheet
        Ticker = Cells(i, 1).Value
        
        'Add to Total Volume
        Volume = Volume + Cells(i, 7).Value
        
        'Print the ticker name in the summary table
        Range("I" & Summary_Table_Row).Value = Ticker
        
        'Print the total volume amount in the summary table
        Range("J" & Summary_Table_Row).Value = Volume
        
        'Add one to the summary table
        Summary_Table_Row = Summary_Table_Row + 1
        
        'Reset the total volume
        Volume = 0
          
        'If the stock market ticker is the same as the previous one
    Else
    
    'Add to Total Volume
    Volume = Volume + Cells(i, 7).Value
    
    End If

Next i

End Sub
