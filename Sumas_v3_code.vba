Sub Limpiar()
'
' Limpiar Macro
'
    Sheets("Hoja1").UsedRange.ClearContents
    Sheets("Hoja1").Columns("A").Font.Bold = False
    Sheets("Hoja1").Range("A1").Select
'
End Sub
Sub Cerrar()
'
' Cerrar Macro
'
    ActiveWorkbook.Close savechanges:=False
    Application.Quit
End Sub
Sub SelectPrinter()
    Application.Dialogs(xlDialogPrinterSetup).Show
End Sub
Sub SumAndPrint()
'Suma e imprime
    Dim LastRow As Long
    With ActiveSheet
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
    
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
        .LeftMargin = Application.CentimetersToPoints(0.3)
        .RightMargin = Application.CentimetersToPoints(0.3)
        .TopMargin = Application.CentimetersToPoints(0.3)
        .BottomMargin = Application.CentimetersToPoints(0.3)
        .HeaderMargin = Application.CentimetersToPoints(0.2)
        .FooterMargin = Application.CentimetersToPoints(0.2)
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    
    Total = Application.Sum(Range(Cells(1, 1), Cells(LastRow, 1)))
    Worksheets("Hoja1").Cells(LastRow + 2, 1).Value = "Total: "
    Worksheets("Hoja1").Cells(LastRow + 3, 1).Value = Total
    Worksheets("Hoja1").Cells(LastRow + 3, 1).Font.Bold = True
    ActualPrinter = Application.ActivePrinter
    Application.ActivePrinter = "POS-58 en Ne00:"
    ActiveSheet.PrintOut
    Application.ActivePrinter = ActualPrinter
End Sub

