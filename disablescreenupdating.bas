Sub SwitchOff(bSwitchOff As Boolean)
'Disable Manual Calculations, Screen Updates, Animations to improve performance
'Taken from https://techcommunity.microsoft.com/t5/excel/9-quick-tips-to-improve-your-vba-macro-performance/m-p/173687

  Dim ws As Worksheet
    
  With Application
    If bSwitchOff Then

      ' OFF
    bScreenUpdate = .ScreenUpdating
      .ScreenUpdating = False
      .EnableAnimations = False
      
      '
      ' switch off display pagebreaks for all worksheets
      '
      For Each ws In ActiveWorkbook.Worksheets
        ws.DisplayPageBreaks = False
      Next ws
    Else
 
      ' ON
      If .Calculation <> lCalcSave And lCalcSave <> 0 Then .Calculation = lCalcSave
      .ScreenUpdating = bScreenUpdate
      .EnableAnimations = True
      
    End If
  End With
End Sub
