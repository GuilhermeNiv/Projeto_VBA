Option Explicit

Dim ws As Worksheet
Dim Sheet As String
Sheet = "Cotação"

Sub CriarOuLimparAba()
    Application.DisplayAlerts = False
    On Error Resume Next
    Worksheets(Sheet).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    Set ws = Worksheets.Add
    ws.Name = Sheet
    ws.Cells.Clear
End Sub

Sub ImportarCotacao(titulo As String, url As String, celulaInicio As Range)
    Dim qt As QueryTable
    celulaInicio.Value = titulo
    Set qt = ws.QueryTables.Add(Connection:="URL;" & url, Destination:=celulaInicio.Offset(1, 0))
    With qt
        .TablesOnlyFromHTML = True
        .Refresh BackgroundQuery:=False
        .SaveData = True
        .MaintainConnection = False
    End With
    With celulaInicio
        .Font.Bold = True
        .Font.Size = 13
        .Interior.Color = RGB(0, 102, 204)
        .Font.Color = RGB(255, 255, 255)
    End With
End Sub

Sub AtualizarTudo()
    CriarOuLimparAba
    Dim linhaAtual As Long
    linhaAtual = 1
    
    ImportarCotacao "Cotação do Dólar", "https://www.infomoney.com.br/cotacoes/dolar/", ws.Range("A" & linhaAtual)
    linhaAtual = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 3
    
    ImportarCotacao "Cotação do Euro", "https://www.infomoney.com.br/cotacoes/euro/", ws.Range("A" & linhaAtual)
    linhaAtual = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 3
    
    ImportarCotacao "Cotação do Bitcoin", "https://www.infomoney.com.br/cotacoes/bitcoin/", ws.Range("A" & linhaAtual)
    
    ws.Range("D1").Value = "Atualizado em:"
    ws.Range("E1").Value = Format(Now, "dd/mm/yyyy hh:mm:ss")
    ws.Range("D1:E1").Font.Bold = True
    ws.Range("D1:E1").Font.Color = RGB(0, 102, 204)
    
    Call CriarGraficos
End Sub

Sub AtualizarDolar()
    Set ws = Worksheets(Sheet)
    Dim cel As Range
    Set cel = ws.Range("A1")
    
    ws.Range("A1:A20").EntireRow.Delete
    
    ImportarCotacao "Cotação do Dólar", "https://www.infomoney.com.br/cotacoes/dolar/", cel
    
    ws.Range("D1").Value = "Atualizado em:"
    ws.Range("E1").Value = Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    Call CriarGraficos
End Sub

Sub AtualizarEuro()
    Set ws = Worksheets(Sheet)
    Dim cel As Range
    Set cel = ws.Columns("A").Find(What:="Cotação do Euro", LookAt:=xlWhole)
    If cel Is Nothing Then
        MsgBox "Cotação do Euro não encontrada. Atualize tudo primeiro.", vbExclamation
        Exit Sub
    End If
    
    ws.Range(cel, cel.Offset(19, 5)).EntireRow.Delete
    
    ImportarCotacao "Cotação do Euro", "https://www.infomoney.com.br/cotacoes/euro/", cel
    
    ws.Range("D1").Value = "Atualizado em:"
    ws.Range("E1").Value = Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    Call CriarGraficos
End Sub

Sub AtualizarBitcoin()
    Set ws = Worksheets(Sheet)
    Dim cel As Range
    Set cel = ws.Columns("A").Find(What:="Cotação do Bitcoin", LookAt:=xlWhole)
    If cel Is Nothing Then
        MsgBox "Cotação do Bitcoin não encontrada. Atualize tudo primeiro.", vbExclamation
        Exit Sub
    End If
    
    ws.Range(cel, cel.Offset(19, 5)).EntireRow.Delete
    
    ImportarCotacao "Cotação do Bitcoin", "https://www.infomoney.com.br/cotacoes/bitcoin/", cel
    
    ws.Range("D1").Value = "Atualizado em:"
    ws.Range("E1").Value = Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    Call CriarGraficos
End Sub

Sub CriarGraficos()
    Dim chartObj As ChartObject
    Dim grafico As Chart
    Dim ultimaLinha As Long
    
    For Each chartObj In ws.ChartObjects
        chartObj.Delete
    Next chartObj
    
    Dim rngDolar As Range
    Set rngDolar = ws.Range("A2:B6")
    
    Set chartObj = ws.ChartObjects.Add(Left:=400, Width:=400, Top:=10, Height:=200)
    Set grafico = chartObj.Chart
    grafico.SetSourceData Source:=rngDolar
    grafico.ChartType = xlColumnClustered
    grafico.HasTitle = True
    grafico.ChartTitle.Text = "Cotação do Dólar"
    
    Dim rngEuro As Range
    Set rngEuro = ws.Range("A22:B26")
    
    Set chartObj = ws.ChartObjects.Add(Left:=400, Width:=400, Top:=220, Height:=200)
    Set grafico = chartObj.Chart
    grafico.SetSourceData Source:=rngEuro
    grafico.ChartType = xlColumnClustered
    grafico.HasTitle = True
    grafico.ChartTitle.Text = "Cotação do Euro"
    
    Dim rngBitcoin As Range
    Set rngBitcoin = ws.Range("A42:B46")
    
    Set chartObj = ws.ChartObjects.Add(Left:=400, Width:=400, Top:=430, Height:=200)
    Set grafico = chartObj.Chart
    grafico.SetSourceData Source:=rngBitcoin
    grafico.ChartType = xlColumnClustered
    grafico.HasTitle = True
    grafico.ChartTitle.Text = "Cotação do Bitcoin"
End Sub

Sub CriarBotoes()
    Dim btnAtualizarTudo As Button
    Dim btnAtualizarDolar As Button
    Dim btnAtualizarEuro As Button
    Dim btnAtualizarBitcoin As Button
    Dim wsBotao As Worksheet
    
    Set wsBotao = Worksheets(Sheet)
    
    Dim shp As Shape
    For Each shp In wsBotao.Shapes
        If shp.Type = msoFormControl Then shp.Delete
    Next shp
    
    Set btnAtualizarTudo = wsBotao.Buttons.Add(10, 10, 120, 30)
    btnAtualizarTudo.Caption = "Atualizar Tudo"
    btnAtualizarTudo.OnAction = "AtualizarTudo"
    
    Set btnAtualizarDolar = wsBotao.Buttons.Add(10, 50, 120, 30)
    btnAtualizarDolar.Caption = "Atualizar Dólar"
    btnAtualizarDolar.OnAction = "AtualizarDolar"
    
    Set btnAtualizarEuro = wsBotao.Buttons.Add(10, 90, 120, 30)
    btnAtualizarEuro.Caption = "Atualizar Euro"
    btnAtualizarEuro.OnAction = "AtualizarEuro"
    
    Set btnAtualizarBitcoin = wsBotao.Buttons.Add(10, 130, 120, 30)
    btnAtualizarBitcoin.Caption = "Atualizar Bitcoin"
    btnAtualizarBitcoin.OnAction = "AtualizarBitcoin"
    
    Dim b As Button
    For Each b In wsBotao.Buttons
        b.Font.Size = 10
        b.Font.Bold = True
        b.ShapeRange.Fill.ForeColor.RGB = RGB(0, 102, 204)
        b.ShapeRange.Fill.Transparency = 0#
        b.ShapeRange.Line.Visible = msoFalse
        b.Font.Color = RGB(255, 255, 255)
    Next b
End Sub

Sub ConfigurarDashboard()
    Application.ScreenUpdating = False
    Call AtualizarTudo
    Call CriarBotoes
    Application.ScreenUpdating = True
    MsgBox "Dashboard configurado com sucesso! Use os botões para atualizar.", vbInformation
End Sub