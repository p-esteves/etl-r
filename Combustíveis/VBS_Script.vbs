Set objexcel = CreateObject("Excel.Application")
objexcel.Visible = False
objexcel.DisplayAlerts = False

Set wb = objexcel.Workbooks.Open("V:\Bot\Energia\Combustíveis\Producao-de-Biodiesel-m3.xls")
With objexcel
    With .Sheets("Plan1").PivotTables("Tabela dinâmica4")
        .ColumnGrand = True
        .RowGrand = True
    End With
    .Sheets("Plan1").Activate
    r = objexcel.Cells.Find("Total Geral").Row
    c = objexcel.Cells.Find("Total Geral").Column
    objexcel.Cells(r + 1, c).ShowDetail = True
End With
wb.Save
wb.Close

Set wb = objexcel.Workbooks.Open("V:\Bot\Energia\Combustíveis\Producao-de-Etanol-m3.xls")
With objexcel
    With .Sheets("Plan1").PivotTables("Tabela dinâmica3")
        .ColumnGrand = True
        .RowGrand = True
    End With
    .Sheets("Plan1").Activate
    r = objexcel.Cells.Find("Total Geral").Row
    c = objexcel.Cells.Find("Total Geral").Column
    objexcel.Cells(r + 1, c).ShowDetail = True
End With
wb.Save
wb.Close

Set wb = objexcel.Workbooks.Open("V:\Bot\Energia\Combustíveis\Producao_de_Gas_Natural_m3.xls")
With objexcel
    With .Sheets("Plan1").PivotTables("Tabela dinâmica1")
        .ColumnGrand = True
        .RowGrand = True
    End With
    .Sheets("Plan1").Activate
    r = objexcel.Cells.Find("Total Geral").Row
    c = objexcel.Cells.Find("Total Geral").Column
    objexcel.Cells(r + 1, c).ShowDetail = True
End With
wb.Save
wb.Close

Set wb = objexcel.Workbooks.Open("V:\Bot\Energia\Combustíveis\Producao_de_Petroleo_m3.xls")
With objexcel
    With .Sheets("Plan1").PivotTables("Tabela dinâmica2")
        .ColumnGrand = True
        .RowGrand = True
    End With
    .Sheets("Plan1").Activate
    r = objexcel.Cells.Find("Total Geral").Row
    c = objexcel.Cells.Find("Total Geral").Column
    objexcel.Cells(r + 1, c).ShowDetail = True
End With
wb.Save
wb.Close

Set wb = objexcel.Workbooks.Open("V:\Bot\Energia\Combustíveis\Vendas_de_Combustiveis_m3.xls")
With objexcel
    With .Sheets("Plan1").PivotTables("Tabela dinâmica1")
        .ColumnGrand = True
        .RowGrand = True
    End With
    .Sheets("Plan1").Activate
    r = objexcel.Cells.Find("Total Geral").Row
    c = objexcel.Cells.Find("Total Geral").Column
    objexcel.Cells(r + 1, c).ShowDetail = True
End With
wb.Save
wb.Close

Set wb = objexcel.Workbooks.Open("V:\Bot\Energia\Combustíveis\Processamento-de-Petroleo-m3.xls")
With objexcel
    With .Sheets("Plan1").PivotTables("Tabela dinâmica5")
        .ColumnGrand = True
        .RowGrand = True
    End With
    .Sheets("Plan1").Activate
    r = objexcel.Cells.Find("Total Geral").Row
    c = objexcel.Cells.Find("Total Geral").Column
    objexcel.Cells(r + 1, c).ShowDetail = True
End With
wb.Save
wb.Close

objexcel.Quit