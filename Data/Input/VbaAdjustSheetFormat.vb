Sub AdjustSheetToExportAsPDF()
    With ActiveSheet

        .Cells.EntireColumn.AutoFit

        .Columns("B").ColumnWidth = 48
        .Columns("D").ColumnWidth = 18
        .Columns("E").ColumnWidth = 6
        .Columns("G").ColumnWidth = 20

        With .PageSetup
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1
        End With
    End With
End Sub