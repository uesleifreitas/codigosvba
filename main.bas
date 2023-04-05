Attribute VB_Name = "main"
Sub run()


Call ExportJPG("GROSS R2", 400, 670)

End Sub

Sub UpdateQueries()
'https://pt.stackoverflow.com/questions/521779/atualizar-dados-power-query-e-copiar-valores-pelo-vba
Dim QueryTable As WorkbookConnection
Dim bRfresh As Boolean
For Each QueryTable In ThisWorkbook.Connections
QueryTable_name = QueryTable
If QueryTable_name = "ThisWorkbookDataModel" Or Name = "ThisWorkbookDataModel" Then GoTo finish
    With ThisWorkbook.Connections(QueryTable_name).OLEDBConnection
        bRfresh = .BackgroundQuery
        .BackgroundQuery = False
        .Refresh
        .BackgroundQuery = bRfresh
    End With
Next QueryTable
finish:
End Sub

Sub ExportJPG(ByVal NamePlan As String, Optional ByVal Altura As Integer = 350, Optional ByVal Largura As Integer = 1000)
    Dim abaTemporaria As Worksheet
    Dim graficoTemporario As Chart
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    utmlinha = ActiveSheet.Cells(Rows.Count, 5).End(xlUp).Row
    utmColuna = ActiveSheet.Cells(utmlinha, Columns.Count).End(xlToLeft).Column
    
    Range(Cells(5, 3), Cells(25, 9)).Select
    Application.Wait (Now + TimeValue("0:00:02"))
    Selection.CopyPicture

    Set abaTemporaria = Worksheets.Add
    Charts.Add
    ActiveChart.Location where:=xlLocationAsObject, Name:=abaTemporaria.Name
  
    Set graficoTemporario = ActiveChart
    Application.Wait (Now + TimeValue("0:00:02"))
    graficoTemporario.Paste
        
    With Selection
        .Height = Altura
        .Width = Largura
    End With
  
    abaTemporaria.ChartObjects(1).Select
    With Selection
        .Height = Altura
        .Width = Largura
    End With
  
    CaminhoDoExcel = ThisWorkbook.Path
    NomeDaImagem = "\" & NamePlan & ".jpg" '& Format(Now, "yymmdd_hhmmss")
    CaminhoDaImagem = Replace(CaminhoDoExcel, "Reports", "imgs") & NomeDaImagem
    Application.Wait (Now + TimeValue("0:00:01"))
    graficoTemporario.Export Filename:=CaminhoDaImagem
    
    Application.DisplayAlerts = False
    abaTemporaria.Delete
    Application.DisplayAlerts = True
             
    Set abaTemporaria = Nothing
    Set graficoTemporario = Nothing
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub

