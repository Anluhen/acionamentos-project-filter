VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub ComboBoxAno_DropButtonClick()
    Dim ws As Worksheet, tbl As ListObject
    Dim vUnique As Variant
    Dim i As Long, j As Long
    Dim Match As Boolean
    Dim minYear As Long, maxYear As Long
    Dim countYears As Long
    
    Set ws = ThisWorkbook.Sheets("PROJETOS")
    Set tbl = ws.ListObjects("TABELA_FILTRO")
    
    minYear = Application.WorksheetFunction.Min(tbl.ListColumns(30).DataBodyRange)
    maxYear = Application.WorksheetFunction.Max(tbl.ListColumns(30).DataBodyRange)
    countYears = maxYear - minYear + 1
    
    ReDim vUnique(1 To countYears)
    For i = 1 To countYears
        vUnique(i) = minYear + i - 1
    Next i

    ' Se o ComboBox já tiver o mesmo número de itens, checa se coincidem
    For i = LBound(vUnique, 1) To UBound(vUnique, 1)
        Match = False
        For j = 0 To Me.ComboBoxAno.ListCount - 1
            If Me.ComboBoxAno.List(j) * 1 = vUnique(i) Then
                Match = True
                Exit For
            End If
        Next j
        
        If Not Match Then
            Exit For
        End If
    Next i
        
    If Match Then
        Exit Sub
    End If
    
    ' Preenche o ComboBox
    With Me.ComboBoxAno
        .Clear
        For i = LBound(vUnique, 1) To UBound(vUnique, 1)
            If Len(Trim(vUnique(i))) > 0 Then
                .AddItem vUnique(i)
            End If
        Next i
    End With
End Sub

Private Sub ComboBoxStatus_DropButtonClick()
    If IsNull(Me.ComboBoxAno.List(1)) Then
        
        Me.ComboBoxStatus.Clear
        
        Me.ComboBoxStatus.AddItem "FATURADO"
        Me.ComboBoxStatus.AddItem "FABRICAÇÃO"
        Me.ComboBoxStatus.AddItem "PEND. REC. RECEITA"
    End If
End Sub

Private Sub ProjetoGlobal_Click()
    TextBoxProjetoGlobal.Value = ""
End Sub

Private Sub Status_Click()
    ComboBoxStatus.Value = ""
End Sub

Private Sub Ano_Click()
    ComboBoxAno.Value = ""
End Sub

Private Sub OV_Click()
    TextBoxOV.Value = ""
End Sub

Private Sub PEP_Click()
    TextBoxPEP.Value = ""
End Sub

Private Sub PM_Click()
    TextBoxPM.Value = ""
End Sub

Private Sub Cliente_Click()
    TextBoxCliente.Value = ""
End Sub

Private Sub ComboBoxStatus_Change()
    Dim ws As Worksheet, tbl As ListObject, rng As Range, criteria As String
    Dim lastRow As Long

    ' Modify Sheet1 and Table1 to match your sheet and table names
    Set ws = ThisWorkbook.Sheets("PROJETOS")
    Set tbl = ws.ListObjects("TABELA_FILTRO")
    
    ' Get the last row with data in the table
    lastRow = tbl.Range.Rows.Count

    ' Get the criteria from the textbox
    criteria = ComboBoxStatus.Value

    With tbl.Range
        If Len(criteria) = 0 Then
            ' Remove filter on field 32 only
            .AutoFilter Field:=32
        Else
            ' (Re-)apply filter on field 32
            .AutoFilter Field:=32, Criteria1:="*" & criteria & "*"
        End If
    End With

End Sub

Private Sub ComboBoxAno_Change()
    Dim ws As Worksheet, tbl As ListObject, rng As Range, criteria As String
    Dim lastRow As Long

    ' Modify Sheet1 and Table1 to match your sheet and table names
    Set ws = ThisWorkbook.Sheets("PROJETOS")
    Set tbl = ws.ListObjects("TABELA_FILTRO")
    
    ' Get the last row with data in the table
    lastRow = tbl.Range.Rows.Count

    ' Get the criteria from the textbox
    criteria = ComboBoxAno.Value

    With tbl.Range
        If Len(criteria) = 0 Then
            ' Remove filter on field 30 only
            .AutoFilter Field:=30
        Else
            ' (Re-)apply filter on field 30
            .AutoFilter Field:=30, Criteria1:="*" & criteria & "*"
        End If
    End With

End Sub

Private Sub TextBoxProjetoGlobal_Change()
    Dim ws As Worksheet, tbl As ListObject, rng As Range, criteria As String
    Dim lastRow As Long

    ' Modify Sheet1 and Table1 to match your sheet and table names
    Set ws = ThisWorkbook.Sheets("PROJETOS")
    Set tbl = ws.ListObjects("TABELA_FILTRO")
    
    ' Get the last row with data in the table
    lastRow = tbl.Range.Rows.Count

    ' Get the criteria from the textbox
    criteria = TextBoxProjetoGlobal.Value

    With tbl.Range
        If Len(criteria) = 0 Then
            ' Remove filter on field 6 only
            .AutoFilter Field:=6
        Else
            ' (Re-)apply filter on field 6
            .AutoFilter Field:=6, Criteria1:="*" & criteria & "*"
        End If
    End With

End Sub

Private Sub TextBoxCliente_Change()
    Dim ws As Worksheet, tbl As ListObject, rng As Range, criteria As String
    Dim lastRow As Long

    ' Modify Sheet1 and Table1 to match your sheet and table names
    Set ws = ThisWorkbook.Sheets("PROJETOS")
    Set tbl = ws.ListObjects("TABELA_FILTRO")
    
    ' Get the last row with data in the table
    lastRow = tbl.Range.Rows.Count

    ' Get the criteria from the textbox
    criteria = TextBoxCliente.Value

    With tbl.Range
        If Len(criteria) = 0 Then
            ' Remove filter on field 4 only
            .AutoFilter Field:=4
        Else
            ' (Re-)apply filter on field 4
            .AutoFilter Field:=4, Criteria1:="*" & criteria & "*"
        End If
    End With
End Sub

Private Sub TextBoxPM_Change()
    Dim ws As Worksheet, tbl As ListObject, rng As Range, criteria As String
    Dim lastRow As Long

    ' Modify Sheet1 and Table1 to match your sheet and table names
    Set ws = ThisWorkbook.Sheets("PROJETOS")
    Set tbl = ws.ListObjects("TABELA_FILTRO")
    
    ' Get the last row with data in the table
    lastRow = tbl.Range.Rows.Count

    ' Get the criteria from the textbox
    criteria = TextBoxPM.Value

    With tbl.Range
        If Len(criteria) = 0 Then
            ' Remove filter on field 8 only
            .AutoFilter Field:=8
        Else
            ' (Re-)apply filter on field 8
            .AutoFilter Field:=8, Criteria1:="*" & criteria & "*"
        End If
    End With
    
End Sub

Private Sub TextBoxOV_Change()
    Dim ws As Worksheet, tbl As ListObject, rng As Range, criteria As String
    Dim lastRow As Long

    ' Modify Sheet1 and Table1 to match your sheet and table names
    Set ws = ThisWorkbook.Sheets("PROJETOS")
    Set tbl = ws.ListObjects("TABELA_FILTRO")
    
    ' Get the last row with data in the table
    lastRow = tbl.Range.Rows.Count

    ' Get the criteria from the textbox
    criteria = TextBoxOV.Value

    With tbl.Range
        If Len(criteria) = 0 Then
            ' Remove filter on field 10 only
            .AutoFilter Field:=10
        Else
            ' (Re-)apply filter on field 10
            .AutoFilter Field:=10, Criteria1:="*" & criteria & "*"
        End If
    End With

End Sub

Private Sub TextBoxPEP_Change()
    Dim ws As Worksheet, tbl As ListObject, rng As Range, criteria As String
    Dim lastRow As Long

    ' Modify Sheet1 and Table1 to match your sheet and table names
    Set ws = ThisWorkbook.Sheets("PROJETOS")
    Set tbl = ws.ListObjects("TABELA_FILTRO")
    
    ' Get the last row with data in the table
    lastRow = tbl.Range.Rows.Count

    ' Get the criteria from the textbox
    criteria = TextBoxPEP.Value

    With tbl.Range
        If Len(criteria) = 0 Then
            ' Remove filter on field 9 only
            .AutoFilter Field:=9
        Else
            ' (Re-)apply filter on field 9
            .AutoFilter Field:=9, Criteria1:="*" & criteria & "*"
        End If
    End With

End Sub

