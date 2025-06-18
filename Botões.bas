Attribute VB_Name = "Botões"
Option Explicit

Sub Atualizar(Optional ShowOnMacroList As Boolean = False)
Attribute Atualizar.VB_ProcData.VB_Invoke_Func = " \n14"
   ActiveWorkbook.RefreshAll
End Sub

Sub LimparFiltros(Optional ShowOnMacroList As Boolean = False)
    ThisWorkbook.Sheets("PROJETOS").TextBoxProjetoGlobal.Value = ""
    ThisWorkbook.Sheets("PROJETOS").ComboBoxStatus.Value = ""
    ThisWorkbook.Sheets("PROJETOS").ComboBoxAno.Value = ""
    ThisWorkbook.Sheets("PROJETOS").TextBoxOV.Value = ""
    ThisWorkbook.Sheets("PROJETOS").TextBoxPEP.Value = ""
    ThisWorkbook.Sheets("PROJETOS").TextBoxPM.Value = ""
    ThisWorkbook.Sheets("PROJETOS").TextBoxCliente.Value = ""
End Sub

Sub CriarCopia(Optional ShowOnMacroList As Boolean = False)
    Dim wsSrc As Worksheet
    Dim wsDest As Worksheet
    Dim rngData As Range
    Dim rngFiltered As Range
    Dim shtName As String

    ' Origem: planilha ativa
    Set wsSrc = ThisWorkbook.Sheets("PROJETOS")

    ' Intervalo completo do Autofiltro (inclui cabeçalho)
    Set rngData = wsSrc.ListObjects("TABELA_FILTRO").Range

    ' Tenta capturar apenas as células visíveis
    On Error Resume Next
    Set rngFiltered = rngData.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    ' Cria nova planilha ao final com nome único
    shtName = "Cópia_" & Format(Now, "yyyymmdd_hhmmss")
    Set wsDest = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Sheets("PROJETOS"))
    On Error Resume Next
    wsDest.Name = shtName
    On Error GoTo 0

    ' Copia somente as linhas visíveis para A1 da nova planilha
    rngFiltered.Copy Destination:=wsDest.Range("A1")

    ' Ajusta largura de colunas
    wsDest.Columns.AutoFit
End Sub
