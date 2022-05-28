Set objexcel = CreateObject("Excel.Application")
objexcel.Visible = False
objexcel.DisplayAlerts = False

Set wb = objexcel.Workbooks.Open("V:\Bot\Energia\Combust�veis\Producao-de-Biodiesel-m3.xls")
With objexcel
    With .Sheets("Plan1").PivotTables("Tabela din�mica4")
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

Set wb = objexcel.Workbooks.Open("V:\Bot\Energia\Combust�veis\Producao-de-Etanol-m3.xls")
With objexcel
    With .Sheets("Plan1").PivotTables("Tabela din�mica3")
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

Set wb = objexcel.Workbooks.Open("V:\Bot\Energia\Combust�veis\Producao_de_Gas_Natural_m3.xls")
With objexcel
    With .Sheets("Plan1").PivotTables("Tabela din�mica1")
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

Set wb = objexcel.Workbooks.Open("V:\Bot\Energia\Combust�veis\Producao_de_Petroleo_m3.xls")
With objexcel
    With .Sheets("Plan1").PivotTables("Tabela din�mica2")
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

Set wb = objexcel.Workbooks.Open("V:\Bot\Energia\Combust�veis\Vendas_de_Combustiveis_m3.xls")
With objexcel
    With .Sheets("Plan1").PivotTables("Tabela din�mica1")
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

Set wb = objexcel.Workbooks.Open("V:\Bot\Energia\Combust�veis\Processamento-de-Petroleo-m3.xls")
With objexcel
    With .Sheets("Plan1").PivotTables("Tabela din�mica5")
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