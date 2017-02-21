
Sub showAllButtons()
  Dim bt As Button
  For Each bt In ActiveSheet.Buttons
    bt.Enabled = True
    bt.Visible = True
    Debug.Print bt.name & " - "; bt.Text
  Next
End Sub

Function isCommaTheDecimalSeparator()

  isCommaTheDecimalSeparator = (Application.International(xlDecimalSeparator) = K_COMMA)

End Function

'Efetua a leitura do estoque e dos dados de material
Sub StockRead()
  Dim sapConn As Object

  Set ws = ThisWorkbook.Worksheets("Stock")
  ws.Cells(1, 1).Select
  
  unprotectThisFile
  If MsgBox("Confirma nova leitura de estoque?" & vbCrLf & "Esta ação limpará todo histórico da última execução", vbYesNo, "ATENÇÃO!!!") = vbNo Then
    protectThisFile
    Exit Sub
  End If
  
  setHeader
  buttonsCheck
  
  
  Set sapConn = Logon
        
  If sapConn Is Nothing Then
    protectThisFile
    OptimizeCode_End
    Exit Sub
  End If
  
  fillParameters sapConn, K_EVENT_STOCKREAD
  setProgressBarTitle "Logon feito com sucesso"
  
  Dim lin As Integer
  lin = 13
    
  Set ws = ThisWorkbook.Sheets("Stock")
  ws.Activate

  loadMARDStock sapConn
  
  If sapConn Is Nothing Then
    OptimizeCode_End
    Exit Sub
  End If
  
  loadMARCFields sapConn
  loadMARAFields sapConn
  loadMBEWFields sapConn
  loadMAKTFields sapConn
  loadMCHBFields sapConn 'EStoque por Lote
  loadMSLBFields sapConn 'Estoque em poder de Terceiros (por Lote quando ha)
  'loadAddressLines sapConn 'Busca Enderecamentos de Estoque e acrescenta linhas
  
  setupFormulas (False)
  loadABCCurve sapConn
  
  protectThisFile
  
  With tbStock.Range.Cells.Font
    .name = "Verdana"
    .Size = 8
  End With
  
  MsgBox "Processamento Finalizado!", vbApplicationModal, "Leitura de Estoque"

  Dim bt As Button
  For Each bt In ActiveSheet.Buttons
    If bt.name = "btnOpenInventory" Then
      bt.Visible = True
      bt.Enabled = True
    End If
  Next
  Set cellDesc = Nothing
  
End Sub

Sub unprotectThisFile()
  Dim pd As String
  Dim wsSrc As Worksheet
  pd = "!tC0rnIngiT"
  
  Set ws = ThisWorkbook.Worksheets("Stock")
  Set wsCard = ThisWorkbook.Worksheets("Cartões")
  Set wsSrc = ThisWorkbook.Worksheets("Source Data")
  
  ws.Unprotect pd
  wsCard.Unprotect pd
  wsSrc.Unprotect pd
  
End Sub
Sub protectThisFile()
  Dim pd As String
  Dim wsSrc As Worksheet
  pd = "!tC0rnIngiT"
  
  Set ws = ThisWorkbook.Worksheets("Stock")
  Set wsCard = ThisWorkbook.Worksheets("Cartões")
  Set wsSrc = ThisWorkbook.Worksheets("Source Data")
  
  ' Only for ws (Stock tab) Allow Filtering withot allowing Sorting
  ws.Protect Password:=pd, DrawingObjects:=True, _
             Scenarios:=True, contents:=True, _
             AllowFormattingCells:=False, _
             AllowFormattingColumns:=False, _
             AllowFormattingRows:=False, _
             AllowInsertingColumns:=False, _
             AllowInsertingRows:=False, _
             AllowInsertingHyperlinks:=False, _
             AllowSorting:=False, _
             AllowFiltering:=True, _
             AllowDeletingRows:=False, _
             AllowDeletingColumns:=False, _
             UserInterfaceOnly:=True
             
  wsCard.Protect Password:=pd, DrawingObjects:=True, _
             Scenarios:=True, _
             contents:=True, _
             AllowFormattingCells:=False, _
             AllowFormattingColumns:=False, _
             AllowFormattingRows:=False, _
             AllowInsertingColumns:=False, _
             AllowInsertingRows:=False, _
             AllowInsertingHyperlinks:=False, _
             AllowSorting:=False, _
             AllowFiltering:=False, _
             AllowDeletingRows:=False, _
             AllowDeletingColumns:=False, _
             UserInterfaceOnly:=True

  wsSrc.Protect Password:=pd, DrawingObjects:=True, _
             Scenarios:=True, _
             contents:=True, _
             AllowFormattingCells:=False, _
             AllowFormattingColumns:=False, _
             AllowFormattingRows:=False, _
             AllowInsertingColumns:=False, _
             AllowInsertingRows:=False, _
             AllowInsertingHyperlinks:=False, _
             AllowSorting:=False, _
             AllowFiltering:=False, _
             AllowDeletingRows:=False, _
             AllowDeletingColumns:=False, _
             UserInterfaceOnly:=True

End Sub

Sub a_test()
  setupFormulas (True)
  setupFormulas (False)
End Sub

Sub setupFormulas(clearFormulas As Boolean)

  Dim rg1stCount As Range
  Dim rg1stCheck As Range
  Dim rg2ndCheck As Range
  Dim rg3rdCheck As Range
  Dim rg1stSum As Range
  Dim rg2ndSum As Range
  Dim rg3rdSum As Range
  Dim rgHas1stCount As Range
  Dim rgHas2ndCount As Range
  Dim rgHas3rdCount As Range
  Dim rgKey As Range
  Dim rgDiff As Range
  Dim rgFinalBalance As Range
  Dim rgInitialBalance As Range
  Dim rgInitialStdBalance As Range
  Dim rgApprove As Range
  Dim rgNewBalance As Range
  Dim rgQtyMB52 As Range
  Dim rgQtyMB52Final As Range
  
  Set ws = ThisWorkbook.Worksheets("Stock")
  Set tbStock = ws.ListObjects("tbStock")
  Set tbRange = tbStock.DataBodyRange
  
  Set rg1stCount = tbStock.ListColumns(K_COL_MENGE_1ST).DataBodyRange
  Set rg1stCheck = tbStock.ListColumns(K_COL_1ST_CHECK).DataBodyRange
  Set rg2ndCheck = tbStock.ListColumns(K_COL_2ND_CHECK).DataBodyRange
  Set rg3rdCheck = tbStock.ListColumns(K_COL_3RD_CHECK).DataBodyRange
  Set rg1stSum = tbStock.ListColumns(K_COL_MENGE_1ST).DataBodyRange
  Set rg2ndSum = tbStock.ListColumns(K_COL_MENGE_2ND).DataBodyRange
  Set rg3rdSum = tbStock.ListColumns(K_COL_MENGE_3RD).DataBodyRange
  Set rgInitialStdBalance = tbStock.ListColumns(K_COL_SDSTD).DataBodyRange
  'Set rgInitialBalance = tbStock.ListColumns(K_COL_SDINI).DataBodyRange
  Set rgFinalBalance = tbStock.ListColumns(K_COL_SDFIM).DataBodyRange
  Set rgDiff = tbStock.ListColumns(K_COL_DPRIC).DataBodyRange
  Set rgKey = tbStock.ListColumns(K_COL_XCARD).DataBodyRange
  Set rgHas1stCount = tbStock.ListColumns(K_COL_HAS_1ST).DataBodyRange
  Set rgHas2ndCount = tbStock.ListColumns(K_COL_HAS_2ND).DataBodyRange
  Set rgHas3rdCount = tbStock.ListColumns(K_COL_HAS_3RD).DataBodyRange
  Set rgApprove = tbStock.ListColumns(K_COL_APROV).DataBodyRange
  Set rgNewBalance = tbStock.ListColumns(K_COL_NEWQT).DataBodyRange
  Set rgQtyMB52 = tbStock.ListColumns(K_COL_SMB52).DataBodyRange
  Set rgQtyMB52Final = tbStock.ListColumns(K_COL_AFINV).DataBodyRange
  
  rg1stCount.Locked = True
  
  If (clearFormulas) Then
    rg1stCheck.FormulaR1C1 = ""
    rg2ndCheck.FormulaR1C1 = ""
    rg3rdCheck.FormulaR1C1 = ""
    rg1stSum.FormulaR1C1 = ""
    rg2ndSum.FormulaR1C1 = ""
    rg3rdSum.FormulaR1C1 = ""
    rgInitialStdBalance.FormulaR1C1 = ""
    'rgInitialBalance.FormulaR1C1 = ""
    rgFinalBalance.FormulaR1C1 = ""
    rgDiff.FormulaR1C1 = ""
    rgKey.FormulaR1C1 = ""
    rgHas1stCount.FormulaR1C1 = ""
    rgHas2ndCount.FormulaR1C1 = ""
    rgHas3rdCount.FormulaR1C1 = ""
    rgApprove.Formula = ""
    rgNewBalance.Formula = ""
    rgQtyMB52.Formula = ""
    rgQtyMB52Final.Formula = ""
  Else
    rg1stCheck.FormulaR1C1 = "=IF([@[1a.Contagem?]]<>""X"","""",IF([@[1a.Contagem]]=[Est.Livre],""Okay"",""Falha""))" '"=IF(OR(VALUE([@[1a.Contagem]])=0,[@[1a.Contagem?]]<>""X"",[@[Fornecedor]]<>""""),"""",IF([@[1a.Contagem]]=@[Est.Livre],""Okay"",""Falha""))"
    rg2ndCheck.FormulaR1C1 = "=IF([@[2a.Contagem?]]<>""X"","""",IF([@[2a.Contagem]]=[Est.Livre],""Okay"",""Falha""))" '"=IF(OR(VALUE([@[2a.Contagem]])=0,[@[2a.Contagem?]]<>""X"",[@[Fornecedor]]<>""""),"""",IF([@[2a.Contagem]]=@[Est.Livre],""Okay"",""Falha""))"
    rg3rdCheck.FormulaR1C1 = "=IF([@[3a.Contagem?]]<>""X"","""",IF([@[3a.Contagem]]=[Est.Livre],""Okay"",""Falha""))" '"=IF(OR(VALUE([@[3a.Contagem]])=0,[@[3a.Contagem?]]<>""X"",[@[Fornecedor]]<>""""),"""",IF([@[3a.Contagem]]=@[Est.Livre],""Okay"",""Falha""))"
    rg1stSum.Formula = "=SUMIF(Cartões!A:N,[@[Chave?]],Cartões!I:I)"
    rg2ndSum.Formula = "=SUMIF(Cartões!A:N,[@[Chave?]],Cartões!J:J)"
    rg3rdSum.Formula = "=SUMIF(Cartões!A:N,[@[Chave?]],Cartões!K:K)"
    rgInitialStdBalance.FormulaR1C1 = "=IF(OR([@Fornecedor]<>"""",AND([@[Batch?]]=""X"",[@Batch]="""")),0, IF(ISNUMBER([@[Preço Std]]),[@[Preço Std]],0)*[@[Est.Livre]])"
    'rgInitialBalance.FormulaR1C1 = "=IF(OR([@Fornecedor]<>"""",AND([@[Batch?]]=""X"",[@Batch]="""")),0, IF(ISNUMBER([@[Preço Med]]),[@[Preço Med]],0)*[@[Est.Livre]])"
    rgFinalBalance.FormulaR1C1 = "=IF(OR([@Fornecedor]<>"""",AND([@[Batch?]]=""X"",[@Batch]="""")),0,IF(AND([@[Novo Saldo]]<>"""",ISNUMBER([@[Novo Saldo]])),[@[Preço Std]]*[@[Novo Saldo]],IF(ISNUMBER([@[Preço Std]]),[@[Preço Std]],0)*[@[Est.Livre]]))" ' "=IF(OR([@Fornecedor]<>"""",AND([@[Batch?]]=""X"",[@Batch]="""")),0,IF([@[3a.Okay?]]<>"""",[@[3a.Contagem]]*[@[Preço Std]],IF([@[2a.Okay?]]<>"""",[@[2a.Contagem]]*[@[Preço Std]],IF([@[1a.Okay?]]<>"""",[@[1a.Contagem]]*[@[Preço Std]],IF(ISNUMBER([@[Preço Std]]),[@[Preço Std]],0)*[@[Est.Livre]]))))"
    rgDiff.FormulaR1C1 = "=IF(OR([@Fornecedor]<>"""",AND([@[Batch?]]=""X"",[@Batch]="""")),0,ROUND([@[Sld.Final]]-[@[Sld.Ini.Std]], 2))"
    rgKey.FormulaR1C1 = "=TEXT([@Material],""000000"")&""|""&[@Batch]"
    rgHas1stCount.Formula = "=IF(ISNA(VLOOKUP([Chave?],tbCard,tbCard[1a. Contagem],FALSE)),"""",IF(OR(NOT(ISNA(VLOOKUP([Chave?]&""|X"",Cartões!B:N,tbCard[Key_1a],FALSE))),SUMIF(tbCard,[@[Chave?]],tbCard[1a. Contagem])>0),""X"",""""))"
    rgHas2ndCount.Formula = "=IF(ISNA(VLOOKUP([Chave?],tbCard,tbCard[2a. Contagem],FALSE)),"""",IF(OR(NOT(ISNA(VLOOKUP([Chave?]&""|X"",Cartões!C:N,tbCard[Key_2a],FALSE))),SUMIF(tbCard,[@[Chave?]],tbCard[2a. Contagem])>0),""X"",""""))"
    rgHas3rdCount.Formula = "=IF(ISNA(VLOOKUP([Chave?],tbCard,tbCard[3a. Contagem],FALSE)),"""",IF(OR(NOT(ISNA(VLOOKUP([Chave?]&""|X"",Cartões!D:N,tbCard[Key_3a],FALSE))),SUMIF(tbCard,[@[Chave?]],tbCard[3a. Contagem])>0),""X"",""""))"
    rgApprove.FormulaR1C1 = "=IF([@[3a.Okay?]]=""Okay"",""Sim"",IF([@[2a.Okay?]]=""Okay"",""Sim"",IF([@[1a.Okay?]]=""Okay"",""Sim"",IF(AND([@Crítico]=""B"",[@[1a.Contagem?]]=""X""),""Sim"",IF(AND([@[3a.Contagem?]]=""X"",[@[3a.Contagem]]=[@[2a.Contagem]]),""Contagens Iguais"",IF(AND([@[3a.Contagem?]]=""X"",[@[3a.Contagem]]=[@[1a.Contagem]]),""Contagens Iguais"",IF([@[3a.Contagem?]]=""X"",""Sim"",IF(AND([@[2a.Contagem?]]=""X"",[@[2a.Contagem]]=[@[1a.Contagem]]),""Contagens Iguais"",IF([@[1a.Okay?]]<>"""",""Não"","""")))))))))"
    rgNewBalance.FormulaR1C1 = "=IF([@[Aprovado?]]="""","""",IF([@Crítico]=""B"",[@[1a.Contagem]],IF([@[3a.Okay?]]<>"""",[@[3a.Contagem]],IF(AND([@[3a.Contagem?]]=""X"",OR([@[3a.Contagem]]=[@[2a.Contagem]],[@[3a.Contagem]]=[@[1a.Okay?]],[@[3a.Okay?]]=[@[Est.Livre]])),[@[3a.Contagem]],IF(AND([@[2a.Contagem?]]=""X"",OR([@[2a.Contagem]]=[@[1a.Contagem]],[@[2a.Contagem]]=[@[Est.Livre]])),[@[2a.Contagem]],IF(AND([@[2a.Contagem?]]=""X"",[@[2a.Contagem]]<>[@[1a.Contagem]]),""Fazer 3a.Contagem"",[@[1a.Contagem]]))))))"
    rgQtyMB52.Formula = "=IF(OR([@Fornecedor]<>"""",AND([@[Batch?]]=""X"",[@Batch]="""")),0,[@[Est.Livre]])"
    rgQtyMB52Final.Formula = "=IF([@[Novo Saldo]]="""",[@[Est.MB52]],[@[Novo Saldo]])"
    
    FormatConditions rg1stCheck, convertNumberToLetter(rg1stCheck.Cells(1, 1).Column) & "14"
    FormatConditions rg2ndCheck, convertNumberToLetter(rg2ndCheck.Cells(1, 1).Column) & "14"
    FormatConditions rg3rdCheck, convertNumberToLetter(rg3rdCheck.Cells(1, 1).Column) & "14"
    FormatConditionsForApproval rgApprove, convertNumberToLetter(rgApprove.Column) & "14"
    FormatConditionsForPercentCheck rg1stSum, convertNumberToLetter(rg1stSum.Column) & "14"
    FormatConditionsForPercentCheck rg2ndSum, convertNumberToLetter(rg2ndSum.Column) & "14"
    FormatConditionsForPercentCheck rg3rdSum, convertNumberToLetter(rg3rdSum.Column) & "14"
    Application.CalculateFullRebuild
  End If
  
End Sub

Sub FormatConditions(lRange As Range, cellAddress As String)

  On Error Resume Next
  lRange.FormatConditions.Delete
  lRange.FormatConditions.Add type:=xlExpression, Formula1:="=$" & cellAddress & "=""Falha"""
  lRange.FormatConditions.Add type:=xlExpression, Formula1:="=$" & cellAddress & "=""Okay"""
  With lRange.FormatConditions(1).Interior 'Falha
      .Pattern = xlSolid
      .PatternColorIndex = xlAutomatic
      .Color = 192
      .TintAndShade = 0
      .PatternTintAndShade = 0
  End With
  With lRange.FormatConditions(1).Font
    .ThemeColor = xlThemeColorDark1
    .TintAndShade = 0
    .Bold = True
  End With
  
  With lRange.FormatConditions(2).Interior 'Okay
      .Pattern = xlSolid
      .PatternColorIndex = xlAutomatic
      .ThemeColor = xlThemeColorAccent3
      .TintAndShade = 0
      .PatternTintAndShade = 0
  End With
  With lRange.FormatConditions(1).Font
    '.ColorIndex = xlAutomatic
    .TintAndShade = 0
  End With
  On Error GoTo 0
End Sub

Sub FormatConditionsForPercentCheck(lRange As Range, cellAddress As String)

  On Error Resume Next
  lRange.FormatConditions.Delete
  '=$H13>($P13*1,1)
  lRange.FormatConditions.Add type:=xlExpression, Formula1:="=$" & cellAddress & ">($" & convertNumberToLetter(K_COL_LABST) & "14*1" & Application.International(xlDecimalSeparator) & "1)"
  With lRange.FormatConditions(1).Interior 'Não
      .Pattern = xlSolid
      .PatternColorIndex = xlAutomatic
      .Color = 192
      .TintAndShade = 0
      .PatternTintAndShade = 0
  End With
  With lRange.FormatConditions(1).Font
    .ThemeColor = xlThemeColorDark1
    .TintAndShade = 0
    .Bold = True
  End With
  On Error GoTo 0
End Sub


Sub FormatConditionsForApproval(lRange As Range, cellAddress As String)

  On Error Resume Next
  lRange.FormatConditions.Delete
  lRange.FormatConditions.Add type:=xlExpression, Formula1:="=$" & cellAddress & "=""Não"""
  lRange.FormatConditions.Add type:=xlExpression, Formula1:="=$" & cellAddress & "=""Sim"""
  lRange.FormatConditions.Add type:=xlExpression, Formula1:="=$" & cellAddress & "=""Contagens Iguais"""
  With lRange.FormatConditions(1).Interior 'Não
      .Pattern = xlSolid
      .PatternColorIndex = xlAutomatic
      .Color = 192
      .TintAndShade = 0
      .PatternTintAndShade = 0
  End With
  With lRange.FormatConditions(1).Font
    .ThemeColor = xlThemeColorDark1
    .TintAndShade = 0
    .Bold = True
  End With
  
  With lRange.FormatConditions(2).Interior 'Sim
      .Pattern = xlSolid
      .PatternColorIndex = xlAutomatic
      .ThemeColor = xlThemeColorAccent3
      .TintAndShade = 0
      .PatternTintAndShade = 0
  End With
  With lRange.FormatConditions(2).Font
    '.ColorIndex = xlAutomatic
    .TintAndShade = 0
  End With
  
  With lRange.FormatConditions(3).Interior 'Qtds. Iguais
      .Pattern = xlSolid
      .PatternColorIndex = xlAutomatic
      .ThemeColor = xlThemeColorAccent6
      .TintAndShade = 0.799981688894314
      .PatternTintAndShade = 0
  End With
  With lRange.FormatConditions(3).Font
    .TintAndShade = 0
    .Bold = True
  End With
  
  On Error GoTo 0
End Sub

Sub firstCount()
  Dim sapConn As Object
'
'  If Not isFirstCountClosingPossible Then
'    MsgBox "Não é possível Fechar a primeira contagem do inventário." & vbCrLf & _
'           "Verificar o botão 'Ajuda' para entender o processo", _
'           vbCritical, "Erro!! Não é possível fechar Primeira Contagem"
'    Exit Sub
'  End If
'

  ' "NÃO será possível adicionar novos cartões",
  If MsgBox("ATENÇÃO!!" & vbCrLf & _
            "Confirma FECHAR a primeira contagem?" & vbCrLf & vbCrLf & _
            "Somente itens contados na primeira contagem estarão disponíveis na segunda." & vbCrLf & _
            "Itens contados com sucesso também estarão indisponíveis." & vbCrLf & _
            "Itens classificados como Crítico:'B' também estarão indisponíveis", vbYesNo, "") = vbNo Then
    Exit Sub
  End If
  
  Set sapConn = Logon
        
  If sapConn Is Nothing Then
    Exit Sub
  End If
  
  Dim previousSystem As String
  previousSystem = Mid(Range(ThisWorkbook.Names("running_environment")).Value, 1, 3)
  
  If sapConn.Connection.system <> previousSystem Then
    MsgBox "Sistema em que o usuario se logou (" & sapConn.Connection.system & ") difere do ambiente de Leitura de Estoque (" & previousSystem & ")", vbCritical, "ERRO!!!"
    Exit Sub
  End If
  
  fillParameters sapConn, K_EVENT_1ST_COUNT
  
  If Not isSheetReadyForSecondCount Then
    If MsgBox("Ainda há Documentos de Inventário abertos que não receberam contagem" & vbCrLf & _
           "Confirma o fechamento da contagem?", vbYesNo, "Erro! Contagem Pendente") = vbNo Then
      'LéoIncluiu (motivo: mesmo com o cancelamento do fechamento da 1ª contagem o campo first_count continuava preenchido
      Range("first_count").Value = ""
      'End LéoIncluiu
      Exit Sub
    End If
  End If
  
  'populateMaterialTable (True) 'Only documents not previously locked
  'generatePhysInventoryDocument sapConn
  
  
  setupEditableRowsAndColumns (K_COL_MENGE_2ND)
  
  For Each bt In ws.Buttons
    If bt.name = "btn2ndCount" Then
      bt.Visible = True
    End If
  Next

  MsgBox "Primeira contagem encerrada com sucesso"

  
End Sub

Sub secondCount()
  Dim sapConn As Object
'
'  If Not isFirstCountClosingPossible Then
'    MsgBox "Não é possível Fechar a primeira contagem do inventário." & vbCrLf & _
'           "Verificar o botão 'Ajuda' para entender o processo", _
'           vbCritical, "Erro!! Não é possível fechar Primeira Contagem"
'    Exit Sub
'  End If
'

  ' "NÃO será possível adicionar novos cartões",
  If MsgBox("ATENÇÃO!!" & vbCrLf & _
            "Confirma FECHAR a segunda contagem?" & vbCrLf & vbCrLf & _
            "Somente itens contados na segunda contagem estarão disponíveis na terceira." & vbCrLf & _
            "Itens contados com sucesso também estarão indisponíveis.", vbYesNo, "") = vbNo Then
    Exit Sub
  End If
  
  Set sapConn = Logon
        
  If sapConn Is Nothing Then
    Exit Sub
  End If
  
  Dim previousSystem As String
  previousSystem = Mid(Range(ThisWorkbook.Names("running_environment")).Value, 1, 3)
  
  If sapConn.Connection.system <> previousSystem Then
    MsgBox "Sistema em que o usuario se logou (" & sapConn.Connection.system & ") difere do ambiente de Leitura de Estoque (" & previousSystem & ")", vbCritical, "ERRO!!!"
    Exit Sub
  End If
  
  fillParameters sapConn, K_EVENT_2ND_COUNT
  
  If Not isSheetReadyForThirdCount Then
    If MsgBox("Ainda há Documentos de Inventário abertos que não receberam contagem" & vbCrLf & _
           "Confirma o fechamento da contagem?", vbYesNo, "Erro! Contagem Pendente") = vbNo Then
      'LéoIncluiu (motivo: mesmo com o cancelamento do fechamento da 1ª contagem o campo first_count continuava preenchido
      Range("third_count").Value = ""
      'End LéoIncluiu
      Exit Sub
    End If
  End If
  
  setupEditableRowsAndColumns (K_COL_MENGE_3RD)
  
  For Each bt In ws.Buttons
    If bt.name = "btn3rdCount" Then
      bt.Visible = True
    End If
  Next

  MsgBox "Segunda contagem encerrada com sucesso"

End Sub
Sub thirdCount()

  Dim sapConn As Object

  ' "NÃO será possível adicionar novos cartões",
  If MsgBox("ATENÇÃO!!" & vbCrLf & _
            "Confirma FECHAR a terceira contagem?" & vbCrLf & vbCrLf & _
            "O Inventário será submetido a Aprovação e não permitirá novas contagens.", vbYesNo, "") = vbNo Then
    Exit Sub
  End If
  
  Set sapConn = Logon
        
  If sapConn Is Nothing Then
    Exit Sub
  End If
  
  Dim previousSystem As String
  previousSystem = Mid(Range(ThisWorkbook.Names("running_environment")).Value, 1, 3)
  
  If sapConn.Connection.system <> previousSystem Then
    MsgBox "Sistema em que o usuario se logou (" & sapConn.Connection.system & ") difere do ambiente de Leitura de Estoque (" & previousSystem & ")", vbCritical, "ERRO!!!"
    Exit Sub
  End If
  
  fillParameters sapConn, K_EVENT_3RD_COUNT
  
  If Not isSheetReadyForAproval Then
    If MsgBox("Ainda há Documentos de Inventário abertos que não receberam contagem" & vbCrLf & _
           "Confirma o fechamento da contagem?", vbYesNo, "Erro! Contagem Pendente") = vbNo Then
      'LéoIncluiu (motivo: mesmo com o cancelamento do fechamento da 1ª contagem o campo first_count continuava preenchido
      Range("third_count").Value = ""
      'End LéoIncluiu
      Exit Sub
    End If
  End If
  
  setupEditableRowsAndColumns (K_COL_APROV)
  
  For Each bt In ws.Buttons
    If bt.name = "btnApprove" Then
      bt.Visible = True
    End If
  Next

  MsgBox "Terceira contagem encerrada com sucesso"



End Sub
Sub approveCount()

  Dim sapConn As Object
  Dim dDiff As Double
  Dim sDiff As String
  ' "NÃO será possível adicionar novos cartões",
  
  dDiff = Range("difference").Value
  sDiff = Format(dDiff, "###,###,##0.00")
  If MsgBox("ATENÇÃO!!" & vbCrLf & _
            "Confirma a APROVAÇÃO do inventário?" & vbCrLf & vbCrLf & _
            "Caso confirme, serão lançadas a Contagem de inventário e o as Diferenças Contábeis." & vbCrLf & _
            "Total estimado de Diferença: BRL " & sDiff, vbYesNo, "") = vbNo Then
    Exit Sub
  End If
  
  Set sapConn = Logon
        
  If sapConn Is Nothing Then
    Exit Sub
  End If
  
  Dim previousSystem As String
  previousSystem = Mid(Range(ThisWorkbook.Names("running_environment")).Value, 1, 3)
  
  If sapConn.Connection.system <> previousSystem Then
    MsgBox "Sistema em que o usuario se logou (" & sapConn.Connection.system & ") difere do ambiente de Leitura de Estoque (" & previousSystem & ")", vbCritical, "ERRO!!!"
    Exit Sub
  End If
  
  fillParameters sapConn, K_EVENT_APPROVAL
  
  ReDim tbMatnr(0) 'Limpa a tabela de materiais
  populateMaterialTable (True)
  
  
  Set ws = ThisWorkbook.Worksheets("Stock")
  Set tbStock = ws.ListObjects("tbStock")
  ReDim tbPhysInvDoc(tbStock.DataBodyRange.Rows.Count)
  Dim oldDoc As String
  Dim pos As Integer
  oldDoc = ""
  pos = 0
  'Obtem a lista de documentos de Inventário
  For Each Row In tbStock.DataBodyRange.Rows
    If ws.Cells(Row.Row, K_COL_IVDHD) <> "" And _
       ws.Cells(Row.Row, K_COL_IVDHD) <> oldDoc Then
       oldDoc = ws.Cells(Row.Row, K_COL_IVDHD)
       pos = pos + 1
       tbPhysInvDoc(pos).physInv_doc = ws.Cells(Row.Row, K_COL_IVDHD)
       tbPhysInvDoc(pos).goods_doc = ""
    End If
  Next
  ReDim Preserve tbPhysInvDoc(pos)
  
  For lin = 1 To UBound(tbPhysInvDoc)
    If PhysicalInventoryCount(sapConn, tbPhysInvDoc(lin).physInv_doc) Then
       BAPICommit sapConn
       If PhysicalInventoryPostDiff(sapConn, tbPhysInvDoc(lin).physInv_doc) Then
          BAPICommit sapConn
       End If
    End If
  Next
  
  setupEditableRowsAndColumns (K_COL_APROV)
  
  For Each bt In ws.Buttons
    If bt.name = "btnNotaFiscal" Then
      bt.Visible = True
    End If
  Next

  MsgBox "Diferenças de inventário Lançadas"


End Sub
Sub postNotaFiscal()

End Sub
Sub setHeader()
    
  Set ws = ThisWorkbook.Sheets("Stock")
  ws.Activate
  
  Set tbStock = ws.ListObjects("tbStock")
  Set tbRange = tbStock.DataBodyRange
  
  tbStock.AutoFilter.ShowAllData
  With tbStock.Range.Cells.Font
    .name = "Verdana"
    .Size = 8
  End With
  
  On Error Resume Next
  tbRange.Delete
  On Error GoTo 0
  
  Range("B4:D12").Value = 0                'Percentuais e quantidades
  Range("B2:E3").Value = vbNullString      'Sob os botoes
  Range("E2:E12").Value = vbNullString     'MEnsagem de status das leituras
  Range("J5:K12").Value = vbNullString     'Campos de controle de processamento

  
  If tbRange Is Nothing Then
    ws.Cells(tbStock.Range(tbStock.Range.Rows.Count, 1).Row, 1) = "1"
    Set tbRange = tbStock.DataBodyRange
  End If
  
  Dim rg As Range
  Set rg = ws.Range("A" & tbRange.Cells(1, 1).Row).EntireRow
  paintRangeNoColor rg
  
  Range("A" & tbRange.Cells(0, 1).Row & ":AR" & tbRange.Cells(0, 1).Row).Select
  Selection = ""
  ws.Cells(1, 1).Select
  
  tbRange.Cells(0, K_COL_MATNR) = "Material"
  tbRange.Cells(0, K_COL_MAKTX) = "Description"
  tbRange.Cells(0, K_COL_ABCIN) = "Crítico"
  tbRange.Cells(0, K_COL_XCHPF) = "Batch?"
  tbRange.Cells(0, K_COL_CHARG) = "Batch"
  tbRange.Cells(0, K_COL_LGORT) = "Local"
  tbRange.Cells(0, K_COL_MENGE_1ST) = "1a.Contagem"
  tbRange.Cells(0, K_COL_1ST_CHECK) = "1a.Okay?"
  tbRange.Cells(0, K_COL_MENGE_2ND) = "2a.Contagem"
  tbRange.Cells(0, K_COL_2ND_CHECK) = "2a.Okay?"
  tbRange.Cells(0, K_COL_MENGE_3RD) = "3a.Contagem"
  tbRange.Cells(0, K_COL_3RD_CHECK) = "3a.Okay?"
  tbRange.Cells(0, K_COL_APROV) = "Aprovado?"
  tbRange.Cells(0, K_COL_NEWQT) = "Novo Saldo"
  tbRange.Cells(0, K_COL_SPERR) = "Bloqueio Inv."
  tbRange.Cells(0, K_COL_LIFNR) = "Fornecedor"
  tbRange.Cells(0, K_COL_LABST) = "Est.Livre"
  tbRange.Cells(0, K_COL_MEINS) = "UoM"
  tbRange.Cells(0, K_COL_SMB52) = "Est.MB52"
  tbRange.Cells(0, K_COL_INSME) = "Est.Qual."
  tbRange.Cells(0, K_COL_SPEME) = "Est.Bloq."
  tbRange.Cells(0, K_COL_IVDHD) = "Inv.Document"
  tbRange.Cells(0, K_COL_IVDIT) = "Inv.Doc.Item"
  tbRange.Cells(0, K_COL_MTDHD) = "Mat.Document"
  tbRange.Cells(0, K_COL_MTDIT) = "Mat.Doc.Item"
  tbRange.Cells(0, K_COL_DOCNM) = "Nota Fiscal"
  tbRange.Cells(0, K_COL_ZZNSR) = "Catalog#"
  tbRange.Cells(0, K_COL_STEUC) = "NCM"
  tbRange.Cells(0, K_COL_MCLAS) = "Mat.Class"
  tbRange.Cells(0, K_COL_MTUSE) = "Mat.Use"
  tbRange.Cells(0, K_COL_MTORG) = "Mat.Origin"
  tbRange.Cells(0, K_COL_VPRSV) = "Contr.Preço"
  tbRange.Cells(0, K_COL_UPRIC_S) = "Preço Std"
  'tbRange.Cells(0, K_COL_UPRIC_V) = "Preço Med"
  tbRange.Cells(0, K_COL_SDSTD) = "Sld.Ini.Std"
  'tbRange.Cells(0, K_COL_SDINI) = "Sld.Inicial"
  tbRange.Cells(0, K_COL_SDFIM) = "Sld.Final"
  tbRange.Cells(0, K_COL_DPRIC) = "$Diferença"
  tbRange.Cells(0, K_COL_MTART) = "Tipo Mat."
  tbRange.Cells(0, K_COL_XCARD) = "Chave?"
  tbRange.Cells(0, K_COL_HAS_1ST) = "1a.Contagem?"
  tbRange.Cells(0, K_COL_HAS_2ND) = "2a.Contagem?"
  tbRange.Cells(0, K_COL_HAS_3RD) = "3a.Contagem?"
  tbRange.Cells(0, K_COL_SALK3) = "Vl.Estq.Planta"
  tbRange.Cells(0, K_COL_LBKUM) = "Qt.Estq.Planta"
  tbRange.Cells(0, K_COL_AFINV) = "Est.MB52 Final"
  tbRange.Cells(0, K_COL_PSTAT) = "MBEW Mat.Status"
  
End Sub

Sub populateMaterialTable(onlyMissingDocs As Boolean)

  Set tbStock = ws.ListObjects("tbStock")

  setProgressCells (K_PROGRESS_SELE)
  setProgressBarTotal (tbStock.DataBodyRange.Rows.Count)
  setProgressBarTitle "Monta lista de materiais a inventariar"
  ReDim tbMatnr(tbStock.DataBodyRange.Rows.Count)
  
  lin = tbStock.DataBodyRange.Cells(1, 1).Row
  pos = 0
  
  While ws.Cells(lin, 1) <> vbNullString
    setProgressBarUpByOne
    'If (Not onlyMissingDocs And isLineSelectableForCounting(ws, lin)) Or _
    '   (onlyMissingDocs And Not isLineSelectableForCounting(ws, lin) And ws.Cells(lin, K_COL_MENGE_1ST) > 0 And ws.Cells(lin, K_COL_IVDHD) <> "") Then
    If (isLineSelectableForCounting(ws, lin)) Then
      pos = pos + 1
      tbMatnr(pos).matnr = ws.Cells(lin, K_COL_MATNR)
      tbMatnr(pos).matnr_edit = Format(ws.Cells(lin, K_COL_MATNR), "000000000000000000")
      If Trim(ws.Cells(lin, K_COL_XCHPF)) = vbNullString Then
        tbMatnr(pos).batchManaged = False
        tbMatnr(pos).batch = vbNullString
      Else
        tbMatnr(pos).batchManaged = True
        tbMatnr(pos).batch = ws.Cells(lin, K_COL_CHARG)
      End If
      tbMatnr(pos).materialType = ws.Cells(lin, K_COL_MTART)
      tbMatnr(pos).description = ws.Cells(lin, K_COL_MAKTX)
      tbMatnr(pos).meins = ws.Cells(lin, K_COL_MEINS)
      tbMatnr(pos).firstCount = Val(ws.Cells(lin, K_COL_MENGE_1ST))
      tbMatnr(pos).firstCountCheck = ws.Cells(lin, K_COL_1ST_CHECK)
      tbMatnr(pos).secondCount = Val(ws.Cells(lin, K_COL_MENGE_2ND))
      tbMatnr(pos).secondCountCheck = ws.Cells(lin, K_COL_2ND_CHECK)
      tbMatnr(pos).thirdCount = Val(ws.Cells(lin, K_COL_MENGE_3RD))
      tbMatnr(pos).thirdCountCheck = ws.Cells(lin, K_COL_3RD_CHECK)
      If ws.Cells(lin, K_COL_NEWQT) <> "" Then
        If Not IsNumeric(ws.Cells(lin, K_COL_NEWQT)) Then
          tbMatnr(pos).postingQuantity = ws.Cells(lin, K_COL_LABST)
        Else
          tbMatnr(pos).postingQuantity = ws.Cells(lin, K_COL_NEWQT)
        End If
      Else
        tbMatnr(pos).postingQuantity = ws.Cells(lin, K_COL_LABST)
      End If
      If tbMatnr(pos).postingQuantity = 0 Then
        tbMatnr(pos).zero_count = "X"
      Else
        tbMatnr(pos).zero_count = vbNullString
      End If
      tbMatnr(pos).physInv_doc = ws.Cells(lin, K_COL_IVDHD)
      tbMatnr(pos).physInv_doc_item = ws.Cells(lin, K_COL_IVDIT)
      tbMatnr(pos).goods_doc = ws.Cells(lin, K_COL_MTDHD)
      tbMatnr(pos).goods_doc_item = ws.Cells(lin, K_COL_MTDIT)
      tbMatnr(pos).NCM = ws.Cells(lin, K_COL_STEUC)
      tbMatnr(pos).matkl = ws.Cells(lin, K_COL_MCLAS)
      tbMatnr(pos).mtuse = ws.Cells(lin, K_COL_MTUSE)
      tbMatnr(pos).mtorg = ws.Cells(lin, K_COL_MTORG)
      tbMatnr(pos).priceControl = ws.Cells(lin, K_COL_VPRSV)
      tbMatnr(pos).unitPrice = Val(ws.Cells(lin, K_COL_UPRIC_S))
      tbMatnr(pos).adjustmentQty = 0
    End If
    lin = lin + 1
  Wend

  ReDim Preserve tbMatnr(pos)

End Sub

Function isLineSelectableForCounting(ByRef ws As Worksheet, ByVal lin As Integer) As Boolean
  'Assume por padrão que não pode contar
  isLineSelectableForCounting = False
  
  
  ' Estoque e poder de 3os nao pode ser contado no inventario
  If ws.Cells(lin, K_COL_LIFNR) <> vbNullString Then
    Exit Function
  End If
    
  ' Somente permitir contagem de linhas com controle de lote
  ' onde o lote estiver preenchido
  If ws.Cells(lin, K_COL_XCHPF) = "X" And _
      ws.Cells(lin, K_COL_CHARG) = vbNullString Then
    Exit Function
  End If
  
  ' Se o material estiver bloqueado para inventário, nao pode ser contado
  If ws.Cells(lin, K_COL_SPERR) <> vbNullString Then
    Exit Function
  End If
  
  ' Se tipo de material não permitir contagem, sair com Falso
  If ws.Cells(lin, K_COL_MTART) = "ZLAG" Or _
     ws.Cells(lin, K_COL_MTART) = "NLAG" Then
    Exit Function
  End If
  
  ' Se o material não foi expandido para a planta (MARC), sair com Falso
  If ws.Cells(lin, K_COL_MTUSE) = "N.Ext." Or _
     ws.Cells(lin, K_COL_XCHPF) = "N.Ext" Then
    Exit Function
  End If
  
  'Se o material não tiver visão de contabilidade cadastrada, não selecionar
  Dim pos As Integer
  pos = InStr(ws.Cells(lin, K_COL_PSTAT), "B") 'B = Visão de Contabilidade
  If pos <= 0 Then
    Exit Function
  End If
    
  ' Se o estoque estiver zerado, inicialmente não abrir inventário
  'If ws.Cells(lin, K_COL_LABST) = 0 Then
  '  Exit Function
  'End If
  
  'Se chegou até aqui, é porque pode contar no inventário
  isLineSelectableForCounting = True
  
End Function

Sub InventoryPost()
  Dim sapConn As Object
  Dim sQty As String
  Set sapConn = Logon
  If sapConn Is Nothing Then
    Exit Sub
  End If
      
  Set ws = ThisWorkbook.Sheets("Stock")
  ws.Activate
  
  ReDim tbMatnr(0)
  
  findMaterialLineByMatnr vbNullString
  setProgressCells (K_PROGRESS_SELE)
  
  cellCurrent.Value = 0
  cellTotal.Value = UBound(tbCells)
  cellPercent.Value = 0

  
  lin = 14
  pos = 0
  While ws.Cells(lin, 1) <> vbNullString
    sQty = ws.Cells(lin, 3)
    If sQty <> vbNullString And _
       ws.Cells(lin, 3) < ws.Cells(lin, 5) Then   'Apenas ajustes "para menos"
       pos = pos + 1
       ReDim Preserve tbMatnr(pos)
       tbMatnr(pos).matnr = ws.Cells(lin, 1)
       tbMatnr(pos).matnr_edit = Format(ws.Cells(lin, 1), "000000000000000000")
       sQty = ws.Cells(lin, 3)
       If sQty = 0 Then
          tbMatnr(pos).firstCount = 0
          tbMatnr(pos).zero_count = "X"
       Else
          tbMatnr(pos).firstCount = ws.Cells(lin, 3)
          tbMatnr(pos).zero_count = vbNullString
       End If
       tbMatnr(pos).description = ws.Cells(lin, 2)
       tbMatnr(pos).meins = ws.Cells(lin, 6)
       tbMatnr(pos).physInv_doc = ws.Cells(lin, 9)
       tbMatnr(pos).physInv_doc_item = ws.Cells(lin, 10)
       tbMatnr(pos).goods_doc = ws.Cells(lin, 11)
       tbMatnr(pos).goods_doc_item = ws.Cells(lin, 12)
       tbMatnr(pos).NCM = ws.Cells(lin, 15)
       tbMatnr(pos).matkl = ws.Cells(lin, 16)
       tbMatnr(pos).mtuse = ws.Cells(lin, 17)
       tbMatnr(pos).mtorg = ws.Cells(lin, 18)
       tbMatnr(pos).priceControl = ws.Cells(lin, 19)
       tbMatnr(pos).unitPrice = ws.Cells(lin, 20)
       tbMatnr(pos).adjustmentQty = ws.Cells(lin, 5) - ws.Cells(lin, 3)
       cellCurrent.Value = pos
       cellPercent.Value = pos / cellTotal
    End If
    lin = lin + 1
  Wend
  
  If UBound(tbMatnr) = 0 Then
    MsgBox "Nenhuma nova quantidade foi informada", vbInformation
    Exit Sub
  End If
  
      
  generatePhysInventoryDocument sapConn
  
' Erro de processamento
  If UBound(tbMatnr) = 0 Then
    MsgBox "Nenhum material teve 'Nova Quantidade' informada'"
    Exit Sub
  End If
  
  BAPICommit sapConn
  
  Dim oldPhysInvDoc As String
  oldPhysInvDoc = vbNullString
  pos = 0
  ReDim tbPhysInvDoc(0)
  For lin = 1 To UBound(tbMatnr)
    If tbMatnr(lin).physInv_doc <> oldPhysInvDoc Then
      pos = pos + 1
      ReDim Preserve tbPhysInvDoc(pos)
      tbPhysInvDoc(pos).physInv_doc = tbMatnr(lin).physInv_doc
      tbPhysInvDoc(pos).goods_doc = tbMatnr(lin).goods_doc
      oldPhysInvDoc = tbMatnr(lin).physInv_doc
    End If
  Next
  
  For lin = 1 To UBound(tbPhysInvDoc)
    PhysicalInventoryCount sapConn, tbPhysInvDoc(lin).physInv_doc
    BAPICommit sapConn
    PhysicalInventoryPostDiff sapConn, tbPhysInvDoc(lin).physInv_doc
    BAPICommit sapConn
    IssueNotaFiscal sapConn, tbPhysInvDoc(lin).physInv_doc, tbPhysInvDoc(lin).goods_doc
    BAPICommit sapConn
  Next
  
  MsgBox "Processamento Finalizado!", vbApplicationModal, "Leitura de Estoque"

End Sub



Function Logon() As Object

  Set ws = ThisWorkbook.Worksheets("Stock")
  setProgressCells (K_PROGRESS_LOGO)
  setProgressBarTotal (5)
  setProgressBarTitle ("(não responsivo) Iniciando Logon")
  setProgressBarUpByOne
  
  
  Set sapConn = CreateObject("SAP.Functions") 'Create ActiveX object
  
  sapConn.LogLevel = 0 '9
  sapConn.Connection.traceLevel = 0 '1

  Set ws = ThisWorkbook.Sheets("Stock")
  ws.Activate


  If sapConn.Connection.Logon(0, False) <> True Then 'Try Logon
    MsgBox "Cannot Log on to SAP"
    Exit Function
  End If
  
  Dim rg As Range
  Set rg = Range(ThisWorkbook.Names("no_zero_stock"))
  If rg.Cells(1, 1) <> "Incluir Estoque Zero" Then
    NoZeroStock = True
  Else
    NoZeroStock = False
  End If
  
  Set Logon = sapConn
End Function

Sub BAPICommit(sapConn As Object)
  Set rfcObj = sapConn.Add("BAPI_TRANSACTION_COMMIT")
  Dim stWait As Object
  Set stWait = rfcObj.Exports("WAIT")
  
  stWait.Value = "X"
  
  If rfcObj.Call = False Then
    MsgBox rfcObj.Exception
  End If
  
End Sub

Sub generatePhysInventoryDocument(sapConn As Object)
  
  Dim lin As Integer
  Dim resultDoc() As String
  
  setProgressCells (K_PROGRESS_PHYS)
  setProgressBarTotal (UBound(tbMatnr))
  setProgressBarTitle "Prepara criação do documento de inventário"
  

  Set rfcObj = sapConn.Add("BAPI_MATPHYSINV_CREATE_MULT")
  Dim stHead, stMaxLines As Object
  Set stHead = rfcObj.Exports("HEAD")
  Set stMaxLines = rfcObj.Exports("MAXITEMS")


  Dim tbItems, tbReturn As Object
  Set tbItems = rfcObj.Tables("ITEMS")
  Set tbReturn = rfcObj.Tables("RETURN")

  'First we set the condition
  'Refresh table
  tbItems.freetable
  tbReturn.freetable
  
  stHead.Value("PLANT") = Plant
  stHead.Value("STGE_LOC") = StLoc
  stHead.Value("POST_BLOCK") = "X"
  
  stMaxLines.Value = "00200"
  
  For lin = 1 To UBound(tbMatnr)
      setProgressBarUpByOne
      tbItems.Rows.Add
      tbItems(tbItems.RowCount, "MATERIAL") = tbMatnr(lin).matnr_edit
      tbItems(tbItems.RowCount, "BATCH") = tbMatnr(lin).batch
  Next
  
  If tbItems.Rows.Count = 0 Then
    ReDim tbMatnr(0)
    Range("open_inventory").Value = ""
    Exit Sub
  End If
    
  If UBound(tbMatnr) = 0 Then
    MsgBox "Nenhum material teve 'Nova quantidade' informada'", vbCritical
    Range("open_inventory").Value = ""
    Exit Sub
  End If
    
  setProgressBarTitle "(não responsivo)Cria documentos de inventário"
  If rfcObj.Call = False Then
     MsgBox rfcObj.Exception
     Range("open_inventory").Value = ""
  End If
  
  Dim sMsg As String
  sMsg = vbNullString
  If tbReturn(1, "TYPE") <> "S" Then
    ReDim tbMatnr(0)
    Range("open_inventory").Value = ""
    For lin = 1 To tbReturn.Rows.Count
      sMsg = sMsg & tbReturn(lin, "MESSAGE") & ": " & tbReturn(lin, "MESSAGE_V1") & vbCrLf
      Debug.Print sMsg
    Next
    For lin = 1 To tbReturn.Rows.Count
      sMsg = sMsg & tbReturn(lin, "MESSAGE") & ": " & tbReturn(lin, "MESSAGE_V1") & vbCrLf
      If lin > 10 Then
        sMsg = sMsg & "(Há mais " & (tbReturn.Rows.Count - 10) & " mensagens não exibidas)"
        lin = tbReturn.Rows.Count
      End If
    Next
    MsgBox sMsg, vbCritical
    Exit Sub
  End If

  'Efetiva dados na tabela
  BAPICommit sapConn

  setProgressBarTotal (tbReturn.Rows.Count)
  setProgressBarTitle "Mapeamento de Material por Doc.Inventário"
  ReDim field(4)
  field(1) = "IBLNR"
  field(2) = "ZEILI"
  field(3) = "MATNR"
  field(4) = "CHARG"
  
  Dim oldDoc As String
  ReDim filtro(tbReturn.Rows.Count)
  lin = 0
  pos = 0
  oldDoc = vbNullString
  For Each lrow In tbReturn.Rows
    setProgressBarUpByOne
    lin = lin + 1
    If tbReturn(lin, "MESSAGE_V1") <> oldDoc Then
      oldDoc = tbReturn(lin, "MESSAGE_V1")
      pos = pos + 1
      filtro(pos) = "OR IBLNR = '" & Format(tbReturn(lin, "MESSAGE_V1"), "0000000000") & "'"
    End If
  Next
  ReDim Preserve filtro(pos)
  filtro(1) = Mid(filtro(1), 4)

  setProgressBarTitle "Busca materiais por doc.inventario (tabela ISEG)"
  resultDoc = readTable(sapConn, "ISEG", field, filtro)
  
  setProgressBarTotal (UBound(resultDoc))
  setProgressBarTitle "Popula campos com Documento de Inventário"
  
  Dim whichMatnr As String
  
  ReDim tbCells(0)
  For lin = 1 To UBound(resultDoc)
    setProgressBarUpByOne
    whichMatnr = Val(resultDoc(lin, 3))
    If resultDoc(lin, K_COL_XCHPF) <> vbNullString Then  'Tem Lote
      pos = findMaterialLineByMatnrAndBatch(whichMatnr, resultDoc(lin, K_COL_XCHPF))
    Else
      pos = findMaterialLineByMatnr(whichMatnr)
    End If
    While ws.Cells(pos, K_COL_MATNR) = Val(whichMatnr) And _
          ws.Cells(pos, K_COL_LIFNR) <> vbNullString                 'Pula linhas com estoque em fornecedor
      pos = pos + 1
    Wend
    If ws.Cells(pos, K_COL_MATNR) <> Val(whichMatnr) Then ' Garante que o material na célula ainda é o mesmo
      pos = 0
    End If
    If pos > 0 Then
      ws.Cells(pos, K_COL_IVDHD) = "'" & resultDoc(lin, 1)
      ws.Cells(pos, K_COL_IVDIT) = "'" & resultDoc(lin, 2)
    End If
  Next
  
  For Each bt In ws.Buttons
    If bt.name = "btn1stCount" Then
      bt.Visible = True
    End If
  Next
  
  setupEditableRowsAndColumns (K_COL_MENGE_1ST)

  setProgressBarTitle "Criação do documento de Inventário Finalizado"
End Sub

Sub setupEditableRowsAndColumns(countNumber As Integer)
  
  Dim lineRange As Range
  Dim cellRange As Range
  
  Set tbStock = ws.ListObjects("tbStock")
  Set tbRange = tbStock.DataBodyRange
  Set ws = ThisWorkbook.Worksheets("Stock")
  
  setProgressCells (K_PROGRESS_PHYS)
  setProgressBarTotal (tbRange.Rows.Count)
  setProgressBarTitle "Configurando campos editáveis"
  
  For Each lrow In tbRange.Rows
    setProgressBarUpByOne
    If mustGrayLineOut(ws, lrow.Row) Then
      Set lineRange = ws.Range("A" & lrow.Row).EntireRow
      lineRange.Locked = True
      With lineRange.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
      End With
      With lineRange.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.0499893185216834
      End With
    Else
      ws.Cells(lrow.Row, K_COL_MENGE_1ST).Locked = True
      ws.Cells(lrow.Row, K_COL_MENGE_2ND).Locked = True
      ws.Cells(lrow.Row, K_COL_MENGE_3RD).Locked = True

      Select Case countNumber
        Case K_COL_MENGE_1ST
          If ws.Cells(lrow.Row, K_COL_XCARD) <> "H" Then
            paintRangeYellow (ws.Cells(lrow.Row, K_COL_MENGE_1ST))
            paintRangeNoColor (ws.Cells(lrow.Row, K_COL_MENGE_2ND))
            paintRangeNoColor (ws.Cells(lrow.Row, K_COL_MENGE_3RD))
          End If
        Case K_COL_MENGE_2ND
          If ws.Cells(lrow.Row, K_COL_XCARD) <> "H" Then
            paintRangeNoColor (ws.Cells(lrow.Row, K_COL_MENGE_1ST))
            If ws.Cells(lrow.Row, K_COL_HAS_1ST) = "X" And _
               ws.Cells(lrow.Row, K_COL_1ST_CHECK) = "Falha" And _
               ws.Cells(lrow.Row, K_COL_ABCIN) = "A" Then
              paintRangeYellow (ws.Cells(lrow.Row, K_COL_MENGE_2ND))
            Else
              paintRangeNoColor (ws.Cells(lrow.Row, K_COL_MENGE_2ND))
            End If
            paintRangeNoColor (ws.Cells(lrow.Row, K_COL_MENGE_3RD))
          End If
        Case K_COL_MENGE_3RD
          If ws.Cells(lrow.Row, K_COL_XCARD) <> "H" Then
            paintRangeNoColor (ws.Cells(lrow.Row, K_COL_MENGE_1ST))
            paintRangeNoColor (ws.Cells(lrow.Row, K_COL_MENGE_2ND))
            If ws.Cells(lrow.Row, K_COL_HAS_2ND) = "X" And _
               ws.Cells(lrow.Row, K_COL_2ND_CHECK) = "Falha" And _
               ws.Cells(lrow.Row, K_COL_MENGE_1ST) <> ws.Cells(lrow.Row, K_COL_MENGE_2ND) Then
              paintRangeYellow (ws.Cells(lrow.Row, K_COL_MENGE_3RD))
            Else
              paintRangeNoColor (ws.Cells(lrow.Row, K_COL_MENGE_3RD))
            End If
          End If
        Case K_COL_APROV
          paintRangeNoColor (ws.Cells(lrow.Row, K_COL_MENGE_1ST))
          paintRangeNoColor (ws.Cells(lrow.Row, K_COL_MENGE_2ND))
          paintRangeNoColor (ws.Cells(lrow.Row, K_COL_MENGE_3RD))
      End Select
      
    End If
  Next
  
  protectThisFile
  
End Sub

Function mustGrayLineOut(ByRef ws As Worksheet, ByVal lin As Integer) As Boolean
  'configura 'cinza' por padrão
  mustGrayLineOut = True
  
  ' Se o material não foi expandido para a planta, acinzenta a linha
  If ws.Cells(lin, K_COL_XCHPF) = "N.Ext" Or _
     ws.Cells(lin, K_COL_STEUC) = "N.Ext" Or _
     ws.Cells(lin, K_COL_MTUSE) = "N.Ext." Then
     Exit Function
  End If

  ' Se o material é administrado por lote (mas não é a linha com estoque do lote), acinzenta
  If ws.Cells(lin, K_COL_XCHPF) = "X" And _
      ws.Cells(lin, K_COL_CHARG) = vbNullString Then
    Exit Function
  End If
  
  ' Se for saldo em poder de terceiros (fornecedor preenchido), acinzenta
  If ws.Cells(lin, K_COL_LIFNR) <> vbNullString Then
    Exit Function
  End If
    
  ' Não contar tipos de material que não geram estoque
  If ws.Cells(lin, K_COL_MTART) = "ZLAG" Or _
     ws.Cells(lin, K_COL_MTART) = "NLAG" Then
     Exit Function
  End If
  
  'Se o material não tiver visão de contabilidade cadastrada, não selecionar
  Dim pos As Integer
  pos = InStr(ws.Cells(lin, K_COL_PSTAT), "B") 'B = Visão de Contabilidade
  If pos <= 0 Then
    Exit Function
  End If
  
  'Se chegou até aqui é porque a linha pode ser usada
  mustGrayLineOut = False
 
End Function

Sub paintRangeNoColor(rg As Range)
    
    With rg.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With rg.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With

End Sub

Sub paintRangeYellow(rg As Range)
    With rg.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With rg.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    rg.Font.Bold = True
End Sub

' REtorna TRUE se conseguiu efetuar o lancamento com sucesso
Function PhysicalInventoryCount(sapConn As Object, physInvDoc As String) As Boolean

  Dim lin As Integer
  
  PhysicalInventoryCount = False
  setProgressCells (K_PROGRESS_PHYS)
  setProgressBarTotal (UBound(tbMatnr))
  setProgressBarTitle physInvDoc & " - Informando contagem (lotes de 200)"
    
  Set rfcObj = sapConn.Add("BAPI_MATPHYSINV_COUNT")
  Dim stPhysInv, stFiscalYear, stCountDate As Object
  Set stPhysInv = rfcObj.Exports("PHYSINVENTORY")
  Set stFiscalYear = rfcObj.Exports("FISCALYEAR")
  'Set stCountDate = rfcObj.Exports("COUNT_DATE")


  Dim tbItems, tbReturn As Object
  Set tbItems = rfcObj.Tables("ITEMS")
  Set tbReturn = rfcObj.Tables("RETURN")

  'First we set the condition
  'Refresh table
  tbItems.freetable
  tbReturn.freetable
  
  stPhysInv.Value = Format(physInvDoc, "0000000000")
  stFiscalYear.Value = Year(Date)
  'stCountDate.Value = Day(Date) & Month(Date) & Year(Date)
  
  Dim item As Integer
  For lin = 1 To UBound(tbMatnr)
    setProgressBarUpByOne
    If tbMatnr(lin).physInv_doc = physInvDoc Then
      tbItems.Rows.Add
      
      item = tbMatnr(lin).physInv_doc_item
      tbItems(tbItems.RowCount, "ITEM") = item
      tbItems(tbItems.RowCount, "MATERIAL") = tbMatnr(lin).matnr_edit
      tbItems(tbItems.RowCount, "BATCH") = tbMatnr(lin).batch
      tbItems(tbItems.RowCount, "ENTRY_QNT") = tbMatnr(lin).postingQuantity
      tbItems(tbItems.RowCount, "ENTRY_UOM") = tbMatnr(lin).meins
      tbItems(tbItems.RowCount, "ZERO_COUNT") = tbMatnr(lin).zero_count
    End If
  Next
  
  If tbItems.RowCount = 0 Then
    Exit Function
  End If
  
  If rfcObj.Call = False Then
    MsgBox rfcObj.Exception
  End If
  
  If tbReturn(1, "TYPE") <> "S" Then
    ReDim tbMatnr(0)
    MsgBox tbReturn(1, "MESSAGE"), vbCritical
    Exit Function
  End If
  
  PhysicalInventoryCount = True
  
End Function

' Retorna TRUE se conseguiu processar com sucesso
Function PhysicalInventoryPostDiff(sapConn As Object, physInvDoc As String) As Boolean

  PhysicalInventoryPostDiff = False
  Dim lin As Integer
  setProgressCells (K_PROGRESS_PHYS)
  setProgressBarTotal (UBound(tbMatnr))
  setProgressBarTitle physInvDoc & " - Lançamento de Diferenças (lotes de 200)"
  
  Set ws = ThisWorkbook.Sheets("Stock")
  ws.Activate
  

  Set rfcObj = sapConn.Add("BAPI_MATPHYSINV_POSTDIFF")
  Dim stPhysInv, stFiscalYear, stCountDate As Object
  Set stPhysInv = rfcObj.Exports("PHYSINVENTORY")
  Set stFiscalYear = rfcObj.Exports("FISCALYEAR")

  Dim tbItems, tbReturn As Object
  Set tbItems = rfcObj.Tables("ITEMS")
  Set tbReturn = rfcObj.Tables("RETURN")

  'First we set the condition
  'Refresh table
  tbItems.freetable
  tbReturn.freetable
  
  stPhysInv.Value = Format(physInvDoc, "0000000000")
  stFiscalYear.Value = Year(Date)
  
  Dim item As Integer
  For lin = 1 To UBound(tbMatnr)
    setProgressBarUpByOne
    If tbMatnr(lin).physInv_doc = physInvDoc Then
      tbItems.Rows.Add
'      cellCurrent.Value = tbItems.RowCount
'      cellPercent.Value = cellCurrent.Value / cellTotal.Value
      
      item = tbMatnr(lin).physInv_doc_item
      tbItems(tbItems.RowCount, "ITEM") = item
      tbItems(tbItems.RowCount, "MATERIAL") = tbMatnr(lin).matnr_edit
      tbItems(tbItems.RowCount, "BATCH") = tbMatnr(lin).batch
    End If
  Next
  
  If tbItems.RowCount = 0 Then
    Exit Function
  End If
  
  If rfcObj.Call = False Then
    MsgBox rfcObj.Exception
  End If
  
  If tbReturn(1, "TYPE") <> "S" Then
    ReDim tbMatnr(0)
    MsgBox tbReturn(1, "MESSAGE"), vbCritical
    Exit Function
  End If

  BAPICommit sapConn

  ReDim field(5)
  field(1) = "MATNR"
  field(2) = "CHARG"
  field(3) = "MBLNR"
  field(4) = "MJAHR"
  field(5) = "ZEILE"
  
  ReDim filtro(tbReturn.Rows.Count)
  Dim sOldDoc As String
  sOldDoc = vbNullString
  pos = 0
  For lin = 1 To tbReturn.Rows.Count
    If tbReturn(lin, "MESSAGE_V2") <> sOldDoc Then
      sOldDoc = tbReturn(lin, "MESSAGE_V2")
      pos = pos + 1
      filtro(pos) = "OR MBLNR = '" & Format(sOldDoc, "0000000000") & "'"
    End If
  Next
  
  If pos = 0 Then
    Exit Function
  End If
  
  ' Remove o "OR" da primeira ocorrencia
  filtro(1) = Mid(filtro(1), 4)
  
  setProgressBarTitle "Leitura do documento de Material"
  resultDoc = readTable(sapConn, "MSEG", field, filtro)

  If IsArrayEmpty(resultDoc) Then
    Exit Function
  End If
  
  setProgressCells (K_PROGRESS_PHYS)
  setProgressBarTotal (UBound(resultDoc))
  setProgressBarTitle "Atribui Documento Material"
  
  Dim whichMatnr As String
  Dim whichBatch As String
  ReDim tbCells(0)
  For lin = 1 To UBound(resultDoc)
    setProgressBarUpByOne
    whichMatnr = Val(resultDoc(lin, 1))
    whichBatch = Trim(resultDoc(lin, 2))
    If whichBatch <> vbNullString Then
      pos = findMaterialLineByMatnrAndBatch(whichMatnr, whichBatch)
    Else
      pos = findMaterialLineByMatnr(whichMatnr)
    End If
    
    If pos > 0 Then
      ws.Cells(pos, K_COL_MTDHD) = "'" & resultDoc(lin, 3) '& "-" & resultDoc(lin, 4)
      ws.Cells(pos, K_COL_MTDIT) = "'" & resultDoc(lin, 5)
    End If
  Next

'
'  'TODO - VOLTAR AQUI
'  Dim pos As Integer
'  Dim pos2 As Integer
'  pos2 = 1
'  For lin = 1 To UBound(tbMatnr)
'    cellCurrent.Value = lin
'    cellPercent.Value = cellCurrent.Value / cellTotal.Value
'    If tbMatnr(lin).physInv_doc = physInvDoc Then
'      If tbMatnr(lin).batchManaged Then
'        pos = findMaterialLineByMatnrAndBatch(tbMatnr(lin).matnr, tbMatnr(lin).batch)
'      Else
'        pos = findMaterialLineByMatnr(tbMatnr(lin).matnr)
'      End If
'      'pos = findMaterialLine(lin)
'      pos2 = pos2 + 1
'      ws.Cells(pos, K_COL_MTDHD) = tbReturn(pos2, "MESSAGE_V2")
'      ws.Cells(pos, K_COL_MTDIT) = tbReturn(pos2, "ROW")
'      tbMatnr(lin).goods_doc = ws.Cells(pos, K_COL_MTDHD)
'      tbMatnr(lin).goods_doc_item = ws.Cells(pos, K_COL_MTDIT)
'    End If
'  Next


  PhysicalInventoryPostDiff = True
End Function


Sub IssueNotaFiscal(sapConn As Object, physInvDoc As String, goodsDoc As String)

  Dim lin As Integer
  Dim sPartnerID As String

  Set ws = ThisWorkbook.Sheets("Stock")
  ws.Activate
  
  Set rfcObj = sapConn.Add("BAPI_J_1B_NF_CREATEFROMDATA")
  Dim stHeader, stHeaderAdd As Object
  Set stHeader = rfcObj.Exports("OBJ_HEADER")
  Set stHeaderAdd = rfcObj.Exports("OBJ_HEADER_ADD")
  
  Dim stDocnum As Object
  Set stDocnum = rfcObj.Imports("E_DOCNUM")

  Dim tbPartner, tbItem, tbItemAdd, tbItemTax As Object
  Dim tbHeaderMsg, tbAddInfo, tbHeaderText, tbItemText, tbReturn As Object
  Set tbPartner = rfcObj.Tables("OBJ_PARTNER")
  Set tbItem = rfcObj.Tables("OBJ_ITEM")
  Set tbItemAdd = rfcObj.Tables("OBJ_ITEM_ADD")
  Set tbItemTax = rfcObj.Tables("OBJ_ITEM_TAX")
  Set tbHeaderMsg = rfcObj.Tables("OBJ_HEADER_MSG")
  Set tbAddInfo = rfcObj.Tables("OBJ_ADD_INFO")
  Set tbHeaderText = rfcObj.Tables("OBJ_HEADER_TEXT")
  Set tbItemText = rfcObj.Tables("OBJ_ITEM_TEXT")
  Set tbReturn = rfcObj.Tables("RETURN")
  
  'First we set the condition
  'Refresh table
  tbPartner.freetable
  tbItem.freetable
  tbItemAdd.freetable
  tbItemTax.freetable
  tbHeaderMsg.freetable
  tbAddInfo.freetable
  tbHeaderText.freetable
  tbItemText.freetable
  tbReturn.freetable
  
  'sDate = Year(Date) & "." & Month(Date) & "." & Day(Date)
  sDate = Year(Date) & Format(Month(Date), "00") & Format(Day(Date), "00")
  sTime = Hour(Date) & Minute(Date) & Second(Time)
  
  If Plant = "8014" Then
    sPartnerID = "0000073616"         'Camorim '"0000074555" BR14
  Else
    sPartnerID = "0000073600"         'Lagoas
  End If
  
  stHeader.Value("MANDT") = "010"
  stHeader.Value("NFTYPE") = "Z1"
  stHeader.Value("DOCTYP") = "1" 'Nota Fiscal
  stHeader.Value("DIRECT") = "2" 'Saida
  stHeader.Value("DOCDAT") = sDate
  stHeader.Value("PSTDAT") = sDate
  stHeader.Value("CREDAT") = sDate
  stHeader.Value("CRETIM") = sTime
  stHeader.Value("FORM") = "NFE1"
  stHeader.Value("MODEL") = "55"
  stHeader.Value("SERIES") = "1"
  stHeader.Value("FATURA") = "X"
  stHeader.Value("MANUAL") = "X"
  'stHeader.Value("BELNR") = Format(goodsDoc, "0000000000")
  'stHeader.Value("GJAHR") = Year(Date)
  stHeader.Value("BUKRS") = CompanyCode
  stHeader.Value("BRANCH") = Branch
  stHeader.Value("PARVW") = "AG"
  stHeader.Value("PARID") = sPartnerID
  stHeader.Value("PARTYP") = "C"
  stHeader.Value("NFE") = "X" 'Eletronica
  stHeader.Value("WAERK") = "BRL"
  
  tbPartner.Rows.Add
  tbPartner(tbPartner.RowCount, "PARVW") = "AG"
  If Plant = "8014" Then
    tbPartner(tbPartner.RowCount, "PARID") = sPartnerID
  Else
    tbPartner(tbPartner.RowCount, "PARID") = sPartnerID
  End If
  tbPartner(tbPartner.RowCount, "PARTYP") = "C"
  
  tbHeaderMsg.Rows.Add
  tbHeaderMsg(tbHeaderMsg.RowCount, "SEQNUM") = 1
  tbHeaderMsg(tbHeaderMsg.RowCount, "LINNUM") = 1
  tbHeaderMsg(tbHeaderMsg.RowCount, "MESSAGE") = "Nota fiscal emitida de acordo com a resolução SEFAZ 720/2014" '"Emissão nota fiscal para perdimento de mercadoria conforme"
  tbHeaderMsg(tbHeaderMsg.RowCount, "MANUAL") = "X"
  
  tbHeaderMsg.Rows.Add
  tbHeaderMsg(tbHeaderMsg.RowCount, "SEQNUM") = 2
  tbHeaderMsg(tbHeaderMsg.RowCount, "LINNUM") = 1
  tbHeaderMsg(tbHeaderMsg.RowCount, "MESSAGE") = "Parte II - Anexo XIII"
  tbHeaderMsg(tbHeaderMsg.RowCount, "MANUAL") = "X"
  
  tbHeaderMsg.Rows.Add
  tbHeaderMsg(tbHeaderMsg.RowCount, "SEQNUM") = 3
  tbHeaderMsg(tbHeaderMsg.RowCount, "LINNUM") = 1
  tbHeaderMsg(tbHeaderMsg.RowCount, "MESSAGE") = "Centro " & Range("Plant").Value & " / Deposito: " & Range("StorageLocation").Value
  tbHeaderMsg(tbHeaderMsg.RowCount, "MANUAL") = "X"
  
  
  Dim item As Integer
  Dim iAmount As Double
  Dim iQty As Double
  ReDim tbCells(0)
  For lin = 1 To UBound(tbMatnr)
    If physInvDoc = tbMatnr(lin).physInv_doc Then
      If tbMatnr(lin).batchManaged Then
        pos = findMaterialLineByMatnrAndBatch(tbMatnr(lin).matnr, tbMatnr(lin).batch)
      Else
        pos = findMaterialLineByMatnr(tbMatnr(lin).matnr)
      End If
      If ws.Cells(pos, K_COL_NEWQT) = "" Then
        iQty = 0
      Else
        iQty = ws.Cells(pos, K_COL_NEWQT) - ws.Cells(pos, K_COL_LABST)
      End If
      If iQty < 0 Then
        iQty = iQty * -1
        iAmount = iQty * ws.Cells(pos, K_COL_UPRIC_S) 'tbMatnr(lin).adjustmentQty * tbMatnr(lin).unitPrice
        tbItem.Rows.Add
        tbItem(tbItem.RowCount, "MANDT") = "010"
        tbItem(tbItem.RowCount, "ITMNUM") = lin
        tbItem(tbItem.RowCount, "ITMTYP") = "01"
        tbItem(tbItem.RowCount, "MATNR") = tbMatnr(lin).matnr_edit
        tbItem(tbItem.RowCount, "MAKTX") = tbMatnr(lin).description
        tbItem(tbItem.RowCount, "MEINS") = tbMatnr(lin).meins
        tbItem(tbItem.RowCount, "MATKL") = tbMatnr(lin).matkl
        tbItem(tbItem.RowCount, "BWKEY") = Plant
        tbItem(tbItem.RowCount, "WERKS") = Plant
        tbItem(tbItem.RowCount, "REFTYP") = "MD"
        tbItem(tbItem.RowCount, "REFKEY") = tbMatnr(lin).goods_doc
        tbItem(tbItem.RowCount, "REFITM") = tbMatnr(lin).goods_doc_item
        tbItem(tbItem.RowCount, "CFOP_10") = "5927AA"
        tbItem(tbItem.RowCount, "NBM") = tbMatnr(lin).NCM
        tbItem(tbItem.RowCount, "TAXLW1") = "SL0"
        tbItem(tbItem.RowCount, "TAXLW2") = "S30"
        tbItem(tbItem.RowCount, "TAXLW4") = "C08"
        tbItem(tbItem.RowCount, "TAXLW5") = "P08"
        tbItem(tbItem.RowCount, "MENGE") = iQty
        tbItem(tbItem.RowCount, "NETPR") = iAmount / iQty
        tbItem(tbItem.RowCount, "NETWR") = iAmount
        tbItem(tbItem.RowCount, "MATUSE") = Mid(tbMatnr(lin).mtuse, 1, 1)
        tbItem(tbItem.RowCount, "MATORG") = Mid(tbMatnr(lin).mtorg, 1, 1)
        
        tbItemTax.Rows.Add
        tbItemTax(tbItemTax.RowCount, "MANDT") = "010"
        tbItemTax(tbItemTax.RowCount, "ITMNUM") = lin
        tbItemTax(tbItemTax.RowCount, "TAXTYP") = "ICM3"
        tbItemTax(tbItemTax.RowCount, "OTHBAS") = iAmount
        
        tbItemTax.Rows.Add
        tbItemTax(tbItemTax.RowCount, "MANDT") = "010"
        tbItemTax(tbItemTax.RowCount, "ITMNUM") = lin
        tbItemTax(tbItemTax.RowCount, "TAXTYP") = "IPI3"
        tbItemTax(tbItemTax.RowCount, "OTHBAS") = iAmount
      End If
    End If
  Next
  
  If tbItem.RowCount = 0 Then
    Exit Sub
  End If
  
  If rfcObj.Call = False Then
    MsgBox rfcObj.Exception
  End If
  
  ' Se deu erro, exibe as primeiras 10 linhas de erro
  Dim sMsg As String
  sMsg = vbNullString
  If tbReturn(1, "TYPE") <> "S" Then
    ReDim tbMatnr(0)
    For lin = 1 To tbReturn.Count
      sMsg = sMsg & tbReturn(1, "MESSAGE")
      If lin > 10 Then
        sMsg = sMsg & vbCrLf & "mais de 10 linhas de mensagem..."
      End If
    Next
    MsgBox sMsg, vbCritical
    Exit Sub
  End If

  BAPICommit sapConn
  
  Dim sDocnum As String
  sDocnum = Format(Val(tbReturn(1, "MESSAGE_V1")), "0000000000")
  
  ReDim field(2)
  field(1) = "MATNR"
  field(2) = "CHARG"
  
  ReDim filtro(1)
  filtro(1) = "DOCNUM = '" & sDocnum & "'"
  
  setProgressBarTitle "Leitura da Nota Fiscal"
  resultDoc = readTable(sapConn, "J_1BNFLIN", field, filtro)

  If IsArrayEmpty(resultDoc) Then
    Exit Sub
  End If
  
  setProgressCells (K_PROGRESS_PHYS)
  setProgressBarTotal (UBound(resultDoc))
  setProgressBarTitle "Atribui DOCNUM"
  
  Dim whichMatnr As String
  Dim whichBatch As String
  ReDim tbCells(0)
  For lin = 1 To UBound(resultDoc)
    setProgressBarUpByOne
    whichMatnr = Val(resultDoc(lin, 1))
    whichBatch = resultDoc(lin, 2)
    If whichBatch <> vbNullString Then
      pos = findMaterialLineByMatnrAndBatch(whichMatnr, whichBatch)
    Else
      pos = findMaterialLineByMatnr(whichMatnr)
    End If
    
    If pos > 0 Then
      ws.Cells(pos, K_COL_DOCNM) = "'" & sDocnum
    End If
  Next

End Sub


Function findMaterialLine(line As Integer)
  Set ws = ThisWorkbook.Sheets("Stock")
  ws.Activate
  
  findMaterialLine = 0
    
  lin = 14
  While ws.Cells(lin, 1) <> vbNullString
    If ws.Cells(lin, 1) = tbMatnr(line).matnr Then
      findMaterialLine = lin
      Exit Function
    End If
    lin = lin + 1
  Wend
End Function


' look for one cell value using Excel Find
Private Function ExcelSearch(ByVal strWorksheet As String _
  , ByVal strSearchArg As String) As Boolean
    
    On Error GoTo Err_Exit
    Worksheets(strWorksheet).Activate
    Worksheets(strWorksheet).Range("A:A").Find(What:=strSearchArg, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Activate
    ExcelSearch = True
    Exit Function
Err_Exit:
    ExcelSearch = False
End Function


Function findMaterialLineByMatnr(matnr As String) As Integer

  Set ws = ThisWorkbook.Sheets("Stock")
  ws.Activate
  
  On Error GoTo ArrayHell
  If UBound(tbCells) = 0 Then
    GoTo ArrayHell
  End If
  On Error Resume Next
  
  GoTo Heaven

ArrayHell:
  Dim lin As Integer
  Dim pos As Integer
  lin = 14
  While ws.Cells(lin, 1) <> vbNullString
    lin = lin + 1
  Wend
  ReDim tbCells(lin - 14)

  For pos = 1 To UBound(tbCells)
    tbCells(pos).matnr = ws.Cells(pos + 13, K_COL_MATNR)
    tbCells(pos).key = Format(ws.Cells(pos + 13, K_COL_MATNR), "000000000000000000") & ws.Cells(pos + 13, K_COL_CHARG)
  Next

Heaven:
  '** Binary Search (once the Material list is already sorted)
  Dim nFirst As Long, nLast As Long
  nFirst = 1
  nLast = UBound(tbCells)
  Do While True
      Dim nMiddle As Long
      Dim strValue As String
      If nFirst > nLast Then
          findMaterialLineByMatnr = 0
          Exit Function
          'Exit Do     ' Failed to find search arg
      End If
      nMiddle = Round((nLast - nFirst) / 2 + nFirst)
      'SheetNameAndRowFromIdx nMiddle, strSheetName, nRow
      strValue = ws.Cells(nMiddle + 13, 1)
      matnr = matnr
      If Val(matnr) < Val(strValue) Then
          nLast = nMiddle - 1
      ElseIf Val(matnr) > Val(strValue) Then
          nFirst = nMiddle + 1
      Else
          findMaterialLineByMatnr = nMiddle + 13
          Exit Do
      End If
  Loop

Hell:
  'findMaterialLineByMatnr = 0
End Function


Function findInDatabase(matnr As String) As ty_db
  
  '** Binary Search (once the Material list is already sorted)
  Dim nFirst As Long, nLast As Long
  nFirst = 1
  nLast = UBound(tbDB)
  Do While True
      Dim nMiddle As Long
      Dim strValue As String
      If nFirst > nLast Then
          Exit Function
          'Exit Do     ' Failed to find search arg
      End If
      nMiddle = Round((nLast - nFirst) / 2 + nFirst)
      'SheetNameAndRowFromIdx nMiddle, strSheetName, nRow
      strValue = tbDB(nMiddle).matnr
      matnr = matnr
      If Val(matnr) < Val(strValue) Then
          nLast = nMiddle - 1
      ElseIf Val(matnr) > Val(strValue) Then
          nFirst = nMiddle + 1
      Else
          findInDatabase = tbDB(nMiddle)
          Exit Do
      End If
  Loop

Hell:
  'findInDatabase = 0
End Function


Function findMaterialLineByMatnrAndBatch(matnr As String, batch As String) As Integer

  Set ws = ThisWorkbook.Sheets("Stock")
  ws.Activate
  
  On Error GoTo ArrayHell
  If UBound(tbCells) = 0 Then
    GoTo ArrayHell
  End If
  On Error Resume Next
  
  GoTo Heaven

ArrayHell:
  Dim lin As Integer
  Dim pos As Integer
  lin = 14
  While ws.Cells(lin, 1) <> vbNullString
    lin = lin + 1
  Wend
  ReDim tbCells(lin - 14)

  For pos = 1 To UBound(tbCells)
    tbCells(pos).matnr = ws.Cells(pos + 13, K_COL_MATNR)
    tbCells(pos).key = Format(ws.Cells(pos + 13, K_COL_MATNR), "000000000000000000") & ws.Cells(pos + 13, K_COL_CHARG)
  Next

Heaven:
  '** Binary Search (once the Material list is already sorted)
  Dim nFirst As Long, nLast As Long
  nFirst = 1
  nLast = UBound(tbCells)
  key = Format(matnr, "000000000000000000") & batch
  Do While True
      Dim nMiddle As Long
      Dim strValue As String
      If nFirst > nLast Then
          findMaterialLineByMatnrAndBatch = 0
          Exit Function
          'Exit Do     ' Failed to find search arg
      End If
      nMiddle = Round((nLast - nFirst) / 2 + nFirst)
      'SheetNameAndRowFromIdx nMiddle, strSheetName, nRow
      If ws.Cells(nMiddle + 13, K_COL_CHARG) = vbNullString Then
        strValue = Format(ws.Cells(nMiddle + 13, K_COL_MATNR), "000000000000000000") & "9999999999"
      Else
        strValue = Format(ws.Cells(nMiddle + 13, K_COL_MATNR), "000000000000000000") & Format(Val(ws.Cells(nMiddle + 13, K_COL_CHARG)), "0000000000")
      End If
      
      
      If key < strValue Then
          nLast = nMiddle - 1
      ElseIf key > strValue Then
          nFirst = nMiddle + 1
      Else
          findMaterialLineByMatnrAndBatch = nMiddle + 13
          Exit Do
      End If
  Loop
  
Hell:
  'findMaterialLineByMatnrAndBatch = 0
End Function



Sub loadMARDStock(sapConn As Object)

  setProgressCells (K_PROGRESS_MARD)
  
  ReDim field(5)
  field(1) = "MATNR"
  field(2) = "SPERR"  'Phys. Inv. Block
  field(3) = "LABST"  'Unrestricted
  field(4) = "INSME"  'In Quality Insp.
  field(5) = "SPEME"  'Blocked
  
    
  If NoZeroStock Then
    ReDim filtro(3)
  Else
    ReDim filtro(2)
  End If
  
  filtro(1) = "WERKS = '" & Plant & "' AND"
  filtro(2) = "LGORT = '" & StLoc & "' "
  
  If NoZeroStock Then
    filtro(3) = "AND (LABST > 0 OR INSME > 0 OR SPEME > 0)"
  End If
     
  setProgressBarTitle "Verifica estoque para planta " & Plant & " / Depósito " & StLoc
  resultDoc = readTable(sapConn, "MARD", field, filtro)
  
  If IsArrayEmpty(resultDoc) Then
    MsgBox "Nenhum registro encontrado", vbInformation
    Set sapConn = Nothing
    Exit Sub
  End If
  
  unprotectThisFile
  Set tbStock = Nothing
  Set tbStock = ws.ListObjects("tbStock")
  Set tbRange = tbStock.Range
  Set tbRange = tbRange.Resize(UBound(resultDoc) + 1, tbRange.Columns.Count)
  On Error Resume Next
  ws.AutoFilterMode = False
  On Error GoTo 0
  
  'Set tbRange = tbStock.DataBodyRange
  If tbRange Is Nothing Then
    ws.Cells(tbStock.Range(tbStock.Range.Rows.Count, 1).Row, 1) = "1"
    Set tbRange = tbStock.Range
    Set tbRange = tbRange.Resize(UBound(resultDoc) + 1, tbRange.Columns.Count)
    On Error Resume Next
    tbStock.Resize tbRange
    On Error GoTo 0
    Set tbRange = tbStock.DataBodyRange
  Else
    On Error GoTo 0
    tbStock.Resize tbRange
    On Error Resume Next
  End If
  
  Dim lValue As String
  setProgressBarTotal (UBound(resultDoc))
  setProgressBarTitle "Populando dados do SAP na planilha"
  
  Set tbRange = tbStock.DataBodyRange
  For doc = 1 To UBound(resultDoc)
    'Application.Calculation = xlCalculationManual
    setProgressBarUpByOne
    tbRange.Cells(doc, K_COL_MATNR) = resultDoc(doc, 1)
    tbRange.Cells(doc, K_COL_LGORT) = Plant & " / " & StLoc 'Range("StorageLocation")
    tbRange.Cells(doc, K_COL_LABST) = CDbl(resultDoc(doc, 3)) / IIf(isCommaTheDecimalSeparator, 1000, 1)
    tbRange.Cells(doc, K_COL_INSME) = CDbl(resultDoc(doc, 4)) / IIf(isCommaTheDecimalSeparator, 1000, 1)
    tbRange.Cells(doc, K_COL_SPEME) = CDbl(resultDoc(doc, 5)) / IIf(isCommaTheDecimalSeparator, 1000, 1)
          
    lValue = Trim(resultDoc(doc, 2))
    Select Case lValue
     Case vbNullString
       tbRange.Cells(doc, K_COL_SPERR) = lValue
     Case "A"
       tbRange.Cells(doc, K_COL_SPERR) = "A - Inventário Físico incompleto"
     Case Else
       tbRange.Cells(doc, K_COL_SPERR) = lValue & " - Material bloqueado p/ Mvto."
    End Select
    
    If doc = 1 Then
      setupFormulas (True)
    End If
  Next
  Application.Calculation = xlCalculationAutomatic
  
  setProgressBarTitle "MARD - Finalizado"
  
End Sub

Sub loadMARAFields(sapConn As Object)

  ReDim field(5)
  field(1) = "MATNR"
  field(2) = "MEINS"
  field(3) = "ZZSNR"
  field(4) = "MATKL"
  field(5) = "MTART"
  
  ReDim filtro(tbStock.DataBodyRange.Rows.Count)
  setProgressCells (K_PROGRESS_MARA)
  setProgressBarTotal (tbStock.DataBodyRange.Rows.Count)
  setProgressBarTitle "Montagem do Filtro de Pesquisa"
  
  lin = tbStock.Range.Cells(1, 1).Row + 1
  filterline = 0
  For Each lrow In tbStock.DataBodyRange.Rows
    setProgressBarUpByOne
    filterline = filterline + 1
    If lrow.Row = lin Then
      filtro(filterline) = "MATNR = '" & Format(ws.Cells(lrow.Row, K_COL_MATNR), "000000000000000000") & "'"
    Else
      filtro(filterline) = "OR MATNR = '" & Format(ws.Cells(lrow.Row, K_COL_MATNR), "000000000000000000") & "'"
    End If
  Next

  setProgressBarTitle "Leitura da Unid.de Medida"
  resultDoc = readTable(sapConn, "MARA", field, filtro)

  setProgressBarTotal (tbStock.DataBodyRange.Rows.Count)
  setProgressBarTitle "Populando Unid. de Medida"
  
  
  
  Dim pos As Integer
  Dim whichMatnr As String
  
  lin = tbStock.Range.Cells(1, 1).Row + 1
  filterline = 0
  For Each Row In tbStock.DataBodyRange.Rows
    setProgressBarUpByOne
    filterline = filterline + 1
    whichMatnr = Val(resultDoc(filterline, 1))
    pos = findMaterialLineByMatnr(whichMatnr)
    If pos > 0 Then
      ws.Cells(pos, K_COL_MEINS) = resultDoc(filterline, 2)
      ws.Cells(pos, K_COL_ZZNSR) = "'" & resultDoc(filterline, 3)
      ws.Cells(pos, K_COL_MCLAS) = "'" & resultDoc(filterline, 4)
      ws.Cells(pos, K_COL_MTART) = resultDoc(filterline, 5)
    End If
  Next
  
  setProgressBarTitle "MARA - Finalizado"

End Sub


Sub loadMARCFields(sapConn As Object)

  Dim lin As Integer

  ReDim field(4)
  field(1) = "MATNR"
  field(2) = "STEUC"
  field(3) = "XCHPF" 'Material administrado por lote na Planta
  field(4) = "ABCIN" 'Inventory ABC Indicator
  
  Set ws = ThisWorkbook.Worksheets("Stock")
  Set tbStock = ws.ListObjects("tbStock")
  
  ReDim filtro(tbStock.DataBodyRange.Rows.Count)
  setProgressCells (K_PROGRESS_MARC)
  setProgressBarTotal (tbStock.DataBodyRange.Rows.Count)
  setProgressBarTitle "Montagem do Filtro de Pesquisa"
  

  lin = tbStock.Range.Cells(1, 1).Row + 1
  filterline = 0
  For Each lrow In tbStock.DataBodyRange.Rows
    setProgressBarUpByOne
    filterline = filterline + 1
    If lrow.Row = lin Then
      filtro(filterline) = "WERKS = '" & Plant & "' AND (MATNR = '" & Format(ws.Cells(lrow.Row, K_COL_MATNR), "000000000000000000") & "'"
    Else
      filtro(filterline) = "OR MATNR = '" & Format(ws.Cells(lrow.Row, K_COL_MATNR), "000000000000000000") & "'"
    End If
  Next
  filtro(filterline) = filtro(filterline) & ")"

  setProgressBarTitle "Buscando informações contábeis do material"
  resultDoc = readTable(sapConn, "MARC", field, filtro)

  setProgressBarTotal (tbStock.DataBodyRange.Rows.Count)
  setProgressBarTitle "Populando informações contábeis do material"
  ReDim tbCells(0)
  Dim whichMatnr As String
  pos = 0
  whichMatnr = vbNullString
  For lin = 1 To UBound(resultDoc)
    setProgressBarUpByOne
    whichMatnr = Val(resultDoc(lin, 1))
    pos = findMaterialLineByMatnr(whichMatnr)
    If pos > 0 Then
      ws.Cells(pos, K_COL_STEUC) = "'" & resultDoc(lin, 2)
      ws.Cells(pos, K_COL_XCHPF) = resultDoc(lin, 3)
      ws.Cells(pos, K_COL_ABCIN) = resultDoc(lin, 4)
    End If
  Next

  setProgressBarTotal (tbStock.DataBodyRange.Rows.Count)
  setProgressBarTitle "Identifica Materiais não extendidos na planta " & Plant
  For lin = tbRange.Cells(1, 1).Row To tbStock.Range.Cells(tbStock.DataBodyRange.Rows.Count, 1).Row + 1
    setProgressBarUpByOne
    If ws.Cells(lin, K_COL_STEUC) = vbNullString Then
      ws.Cells(lin, K_COL_STEUC) = "N.Ext"
      ws.Cells(lin, K_COL_XCHPF) = "N.Ext"
      ws.Cells(lin, K_COL_ABCIN) = "N.Ext"
    End If
  Next
  
  setProgressBarTitle "MARC - Finalizado"
End Sub

Sub loadMCHBFields(sapConn As Object)

  Dim lin As Integer

  ReDim field(5)
  field(1) = "MATNR"
  field(2) = "CHARG"
  field(3) = "CLABS" 'Unrestricted stock
  field(4) = "CINSM" 'Quality Inspection stock
  field(5) = "CSPEM" 'Blocked stock
  
  ReDim filtro(tbStock.DataBodyRange.Rows.Count + 2) 'Linhas adicionais para planta e deposito e para remover estoque zero
  setProgressCells (K_PROGRESS_MCHB)
  setProgressBarTotal (tbStock.DataBodyRange.Rows.Count)
  setProgressBarTitle "Montagem do Filtro de Pesquisa"
  
  
  lin = tbStock.Range.Cells(1, 1).Row + 1
  filterline = 0
  For Each lrow In tbStock.DataBodyRange.Rows
    setProgressBarUpByOne
    If ws.Cells(lrow.Row, K_COL_XCHPF) = "X" Then
      filterline = filterline + 1
      If filterline = 1 Then
        filtro(filterline) = "(MATNR = '" & Format(ws.Cells(lrow.Row, K_COL_MATNR), "000000000000000000") & "'"
      Else
        filtro(filterline) = "OR MATNR = '" & Format(ws.Cells(lrow.Row, K_COL_MATNR), "000000000000000000") & "'"
      End If
    End If
  Next
  
  If filterline = 0 Then
    setProgressBarTitle "MCHB - Finalizado"
    Exit Sub
  End If
  
  filterline = filterline + 1
  filtro(filterline) = ") AND WERKS = '" & Plant & "' AND LGORT = '" & StLoc & "'"
  If NoZeroStock Then
    filterline = filterline + 1
    filtro(filterline) = "AND (CLABS > 0 OR CINSM > 0 OR CSPEM > 0)"
  End If
  ReDim Preserve filtro(filterline)

  setProgressBarTitle "Buscando lotes para o  material"
  resultDoc = readTable(sapConn, "MCHB", field, filtro)

  If IsArrayEmpty(resultDoc) Then
    Exit Sub
  End If

  setProgressBarTotal (UBound(resultDoc))
  setProgressBarTitle "Acrescentando estoque em lote"
  
  pos2 = tbStock.Range.Cells(tbStock.Range.Rows.Count, 1).Row 'Grava a última linha da planilha
  
  'Amplia a tabela, acrescentando o numero de linhas retornadas da tabela MCHB
  Set tbRange = tbRange.Resize(tbStock.Range.Rows.Count + UBound(resultDoc), tbRange.Columns.Count)
  On Error Resume Next
  tbStock.Resize tbRange
  On Error GoTo 0
  Set tbRange = tbStock.DataBodyRange
  
  
  Dim whichMatnr As String
  Dim dUnrestricted As Double
  Dim dQuality As Double
  Dim dBlocked As Double
  pos = 0
  whichMatnr = vbNullString
  For lin = 1 To UBound(resultDoc)
    setProgressBarUpByOne
    whichMatnr = Val(resultDoc(lin, 1))
    pos = findMaterialLineByMatnr(whichMatnr)
    If pos > 0 Then
      pos2 = pos2 + 1
      dUnrestricted = CDbl(resultDoc(lin, 3)) / IIf(isCommaTheDecimalSeparator, 1000, 1)
      dQuality = CDbl(resultDoc(lin, 4)) / IIf(isCommaTheDecimalSeparator, 1000, 1)
      dBlocked = CDbl(resultDoc(lin, 5)) / IIf(isCommaTheDecimalSeparator, 1000, 1)
      ws.Cells(pos2, K_COL_MATNR) = ws.Cells(pos, K_COL_MATNR)
      ws.Cells(pos2, K_COL_ABCIN) = ws.Cells(pos, K_COL_ABCIN)
      ws.Cells(pos2, K_COL_MAKTX) = ws.Cells(pos, K_COL_MAKTX)
      ws.Cells(pos2, K_COL_MTART) = ws.Cells(pos, K_COL_MTART)
      ws.Cells(pos2, K_COL_XCHPF) = ws.Cells(pos, K_COL_XCHPF)
      ws.Cells(pos2, K_COL_LGORT) = ws.Cells(pos, K_COL_LGORT)
      ws.Cells(pos2, K_COL_SPERR) = ws.Cells(pos, K_COL_SPERR)
      ws.Cells(pos2, K_COL_MEINS) = ws.Cells(pos, K_COL_MEINS)
      ws.Cells(pos2, K_COL_ZZNSR) = ws.Cells(pos, K_COL_ZZNSR)
      ws.Cells(pos2, K_COL_STEUC) = ws.Cells(pos, K_COL_STEUC)
      ws.Cells(pos2, K_COL_MCLAS) = ws.Cells(pos, K_COL_MCLAS)
      ws.Cells(pos2, K_COL_MTUSE) = ws.Cells(pos, K_COL_MTUSE) 'getMaterialUseDescription(ws.Cells(pos, K_COL_MTUSE))
      ws.Cells(pos2, K_COL_MTORG) = ws.Cells(pos, K_COL_MTORG) 'getMaterialOriginDescription(ws.Cells(pos, K_COL_MTORG))
      ws.Cells(pos2, K_COL_VPRSV) = ws.Cells(pos, K_COL_VPRSV)
      ws.Cells(pos2, K_COL_UPRIC_S) = ws.Cells(pos, K_COL_UPRIC_S)
      'ws.Cells(pos2, K_COL_UPRIC_V) = ws.Cells(pos, K_COL_UPRIC_V)
      ws.Cells(pos2, K_COL_CHARG) = "'" & resultDoc(lin, 2)
      ws.Cells(pos2, K_COL_LABST) = dUnrestricted
      ws.Cells(pos2, K_COL_INSME) = dQuality
      ws.Cells(pos2, K_COL_SPEME) = dBlocked
      
    End If
  Next
  
  
  tbStock.Sort.SortFields.Clear
  tbStock.Sort.SortFields.Add _
      key:=Range("tbStock[Material]"), SortOn:=xlSortOnValues, Order:=xlAscending _
      , DataOption:=xlSortNormal
  tbStock.Sort.SortFields.Add _
      key:=Range("tbStock[Batch]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
      DataOption:=xlSortNormal
  With tbStock.Sort
      .Header = xlYes
      .MatchCase = False
      .Orientation = xlTopToBottom
      .SortMethod = xlPinYin
      .Apply
  End With
  
  
  setProgressBarTitle "MCHB - Finalizado"
End Sub

Function getMaterialUseDescription(use As String) As String
    
  Select Case use
    Case "0"
      getMaterialUseDescription = "0 - Revenda"
    Case "1"
      getMaterialUseDescription = "1 - Industr."
    Case "2"
      getMaterialUseDescription = "2 - Consumo"
    Case Else
      getMaterialUseDescription = use & " - Imobiliz."
  End Select
  
End Function

Function getMaterialOriginDescription(origen As String) As String
  
  Select Case origen
    Case "1"
      getMaterialOriginDescription = "1 - Import.Diretamente"
    Case "2"
      getMaterialOriginDescription = "2 - Import.Adiqu.nacionalmente"
    Case "0"
      getMaterialOriginDescription = "0 - Nac.Exceto cód. 3, 4, 5, 8"
    Case "3"
      getMaterialOriginDescription = "3 - Nac.%Import.40% a 70%"
    Case "4"
      getMaterialOriginDescription = "4 - Nac.Incentivo tributário"
    Case "5"
      getMaterialOriginDescription = "5 - Nac.%Import inf.a 40%"
    Case "6"
      getMaterialOriginDescription = "6 - Import.Direta,sem similar nac"
    Case "7"
      getMaterialOriginDescription = "7 - Import.Adiqu.Nac.,sem similar"
    Case Else
      getMaterialOriginDescription = origen & " - Nac.Conteúdo > 70%"
  End Select
End Function
Sub loadMSLBFields(sapConn As Object)

  Dim lin As Integer

  ReDim field(6)
  field(1) = "MATNR"
  field(2) = "CHARG"
  field(3) = "LIFNR"
  field(4) = "LBSPR" 'Phys. Inv. Block
  field(5) = "LBLAB" 'Unrestricted stock
  field(6) = "LBINS" 'Quality Inspection stock
  
  ReDim filtro(tbStock.DataBodyRange.Rows.Count + 2) 'Linhas adicionais para planta e deposito e somente itens com saldo
  setProgressCells (K_PROGRESS_MCHB)
  setProgressBarTotal (tbStock.DataBodyRange.Rows.Count)
  setProgressBarTitle "Montagem do Filtro de Pesquisa"
  
  
  lin = tbStock.Range.Cells(1, 1).Row + 1
  filterline = 0
  For Each lrow In tbStock.DataBodyRange.Rows
    setProgressBarUpByOne
    filterline = filterline + 1
    If filterline = 1 Then
      filtro(filterline) = "(MATNR = '" & Format(ws.Cells(lrow.Row, K_COL_MATNR), "000000000000000000") & "'"
    Else
      filtro(filterline) = "OR MATNR = '" & Format(ws.Cells(lrow.Row, K_COL_MATNR), "000000000000000000") & "'"
    End If
  Next
  filterline = filterline + 1
  filtro(filterline) = ") AND WERKS = '" & Plant & "'"
  'If NoZeroStock Then
    filterline = filterline + 1
    filtro(filterline) = " AND (LBLAB > 0 OR LBINS > 0)"
  'End If
  ReDim Preserve filtro(filterline)

  setProgressBarTitle "Buscando estoque em poder de Terceiros"
  resultDoc = readTable(sapConn, "MSLB", field, filtro)

  If IsArrayEmpty(resultDoc) Then
    Exit Sub
  End If
  
  setProgressBarTotal (UBound(resultDoc))
  setProgressBarTitle "Acrescentando estoque em Terceiros"
  
  pos2 = tbStock.Range.Cells(tbStock.Range.Rows.Count, 1).Row 'Grava a última linha da planilha
  
  'Amplia a tabela, acrescentando o numero de linhas retornadas da tabela MSLB
  Set tbRange = tbRange.Resize(tbStock.Range.Rows.Count + UBound(resultDoc), tbRange.Columns.Count)
  On Error Resume Next
  tbStock.Resize tbRange
  On Error GoTo 0
  Set tbRange = tbStock.DataBodyRange
  
  
  Dim whichMatnr As String
  Dim dUnrestricted As Double
  Dim dQuality As Double
  Dim dBlocked As Double
  pos = 0
  whichMatnr = vbNullString
  For lin = 1 To UBound(resultDoc)
    setProgressBarUpByOne
    whichMatnr = Val(resultDoc(lin, 1))
    pos = findMaterialLineByMatnr(whichMatnr)
    If pos > 0 Then
      pos2 = pos2 + 1
      dUnrestricted = CDbl(resultDoc(lin, 5)) / IIf(isCommaTheDecimalSeparator, 1000, 1)
      dQuality = CDbl(resultDoc(lin, 6)) / IIf(isCommaTheDecimalSeparator, 1000, 1)
      dBlocked = 0
      ws.Cells(pos2, K_COL_MATNR) = ws.Cells(pos, K_COL_MATNR)
      ws.Cells(pos2, K_COL_XCHPF) = ws.Cells(pos, K_COL_XCHPF)
      ws.Cells(pos2, K_COL_CHARG) = ws.Cells(pos, K_COL_CHARG)
      ws.Cells(pos2, K_COL_LGORT) = ws.Cells(pos, K_COL_LGORT)
      ws.Cells(pos2, K_COL_ABCIN) = ws.Cells(pos, K_COL_ABCIN)
      ws.Cells(pos2, K_COL_MTART) = ws.Cells(pos, K_COL_MTART)
      ws.Cells(pos2, K_COL_MAKTX) = ws.Cells(pos, K_COL_MAKTX)
      ws.Cells(pos2, K_COL_LIFNR) = "'" & resultDoc(lin, 3)
      ws.Cells(pos2, K_COL_SPERR) = ws.Cells(pos, K_COL_SPERR)
      ws.Cells(pos2, K_COL_MEINS) = ws.Cells(pos, K_COL_MEINS)
      ws.Cells(pos2, K_COL_ZZNSR) = ws.Cells(pos, K_COL_ZZNSR)
      ws.Cells(pos2, K_COL_STEUC) = ws.Cells(pos, K_COL_STEUC)
      ws.Cells(pos2, K_COL_MCLAS) = ws.Cells(pos, K_COL_MCLAS)
      ws.Cells(pos2, K_COL_MTUSE) = ws.Cells(pos, K_COL_MTUSE)
      ws.Cells(pos2, K_COL_MTORG) = ws.Cells(pos, K_COL_MTORG)
      ws.Cells(pos2, K_COL_VPRSV) = ws.Cells(pos, K_COL_VPRSV)
      ws.Cells(pos2, K_COL_UPRIC_S) = ws.Cells(pos, K_COL_UPRIC_S)
      'ws.Cells(pos2, K_COL_UPRIC_V) = ws.Cells(pos, K_COL_UPRIC_V)
      ws.Cells(pos2, K_COL_CHARG) = "'" & resultDoc(lin, 2)
      ws.Cells(pos2, K_COL_LABST) = dUnrestricted
      ws.Cells(pos2, K_COL_INSME) = dQuality
      ws.Cells(pos2, K_COL_SPEME) = dBlocked
      
    End If
  Next
  
  
  tbStock.Sort.SortFields.Clear
  tbStock.Sort.SortFields.Add key:=Range("tbStock[Material]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  tbStock.Sort.SortFields.Add key:=Range("tbStock[Batch]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  With tbStock.Sort
      .Header = xlYes
      .MatchCase = False
      .Orientation = xlTopToBottom
      .SortMethod = xlPinYin
      .Apply
  End With
  
  
  setProgressBarTitle "MCHB/MSLB - Finalizado"
End Sub

Sub loadABCCurve(sapConn As Object)

  Dim lin As Integer
  Dim currentSum As Double
  Dim ABC_80 As Double

  setProgressCells (K_PROGRESS_MCHB)
  setProgressBarTotal (tbStock.DataBodyRange.Rows.Count)
  setProgressBarTitle "Ordenação para atribuir curva ABC"
  
  
  tbStock.Sort.SortFields.Clear
  'tbStock.Sort.SortFields.Add key:=Range("tbStock[Sld.Inicial]"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
  tbStock.Sort.SortFields.Add key:=Range("tbStock[Sld.Ini.Std]"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
  With tbStock.Sort
      .Header = xlYes
      .MatchCase = False
      .Orientation = xlTopToBottom
      .SortMethod = xlPinYin
      .Apply
  End With
  
  setProgressBarTitle "Assinala curva ABC"
  currentSum = 0
  ABC_80 = Range("=[" & ThisWorkbook.name & "]" & Mid(ThisWorkbook.Names("ABC_80"), 2))
  Dim hasChangedAssignement As Boolean
  hasChangedAssignement = False
  For Each Row In tbStock.DataBodyRange.Rows
    setProgressBarUpByOne
    If ws.Cells(Row.Row, K_COL_LIFNR) = vbNullString Then
      currentSum = currentSum + ws.Cells(Row.Row, K_COL_SDSTD)
      If currentSum <= ABC_80 Then
        ws.Cells(Row.Row, K_COL_ABCIN) = "A"
      Else
        If Not hasChangedAssignement Then  'Garante que o item que está no limite entre A e B seja A
          hasChangedAssignement = True
          ws.Cells(Row.Row, K_COL_ABCIN) = "A"
        Else
          ws.Cells(Row.Row, K_COL_ABCIN) = "B"
        End If
      End If
    Else
      ws.Cells(Row.Row, K_COL_ABCIN) = "B"
    End If
  Next
  
  setProgressBarTitle "Atribuir Contagem a lotes"
  tbStock.Sort.SortFields.Clear
  tbStock.Sort.SortFields.Add key:=Range("tbStock[Material]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  tbStock.Sort.SortFields.Add key:=Range("tbStock[Crítico]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  With tbStock.Sort
      .Header = xlYes
      .MatchCase = False
      .Orientation = xlTopToBottom
      .SortMethod = xlPinYin
      .Apply
  End With
  
  Dim oldMatnr As String
  Dim abc_ind As String
  oldMatnr = ""
  abc_ind = ""
  setProgressBarTotal (tbStock.DataBodyRange.Rows.Count)
  setProgressBarTitle ("ABC - Atribui curva a lote inteiro")
  For Each Row In tbStock.DataBodyRange.Rows
    setProgressBarUpByOne
    If ws.Cells(Row.Row, K_COL_MATNR) <> ws.Cells(Row.Row - 1, K_COL_MATNR) Or _
       ws.Cells(Row.Row, K_COL_LIFNR) <> vbNullString Then
      abc_ind = ws.Cells(Row.Row, K_COL_ABCIN)
    End If
    ws.Cells(Row.Row, K_COL_ABCIN) = abc_ind
  Next
  
  
  setProgressBarTitle "Ordenação por Material / Lote"
  tbStock.Sort.SortFields.Clear
  tbStock.Sort.SortFields.Add key:=Range("tbStock[Material]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  tbStock.Sort.SortFields.Add key:=Range("tbStock[Batch]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  With tbStock.Sort
      .Header = xlYes
      .MatchCase = False
      .Orientation = xlTopToBottom
      .SortMethod = xlPinYin
      .Apply
  End With
  
  setProgressBarTitle "ABC atribuido - Processamento Finalizado"
End Sub

Sub loadAddressLines(sapConn As Object)
  Dim wbDB As Workbook
  Dim wsDB As Worksheet
  Dim lin As Integer
  Dim pos As Integer
  Dim qtLinhas As Integer
  
  Set wb = ThisWorkbook
  Set wbDB = Workbooks.Open("C:\Users\nascimenr\Desktop\Endereçamento " & Plant & ".xlsm")
  ws.Activate
  Set wsDB = wbDB.Sheets("Banco de Dados")
  
  wsDB.Activate
'  wsDB.ListObjects (1)
'  wsDB.Range("A1").Select
'  Set mylastcell = wsDB.Cells(1, 1).SpecialCells(xlLastCell)
'  mylastcelladd = wsDB.Cells(mylastcell.Row, mylastcell.Column).Address
'  myrange = "A2:" & mylastcelladd
'  wsDB.Range(myrange).Select
  
  wsDB.ListObjects(1).Sort.SortFields.Clear
  wsDB.ListObjects(1).Sort.SortFields.Add key:=Range(wsDB.ListObjects(1).name & "[Código]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  With wsDB.ListObjects(1).Sort
      .Header = xlYes
      .MatchCase = False
      .Orientation = xlTopToBottom
      .SortMethod = xlPinYin
      .Apply
  End With
  ws.Activate
  
  
  setProgressCells (K_PROGRESS_MCHB)
  setProgressBarTotal (1)
  setProgressBarTitle "DB - Identificando dados"
  
  lin = 3
  While wsDB.Cells(lin, 1) <> vbNullString
    lin = lin + 1
  Wend
  qtLinhas = lin - 3
  
  setProgressBarTotal (qtLinhas)
  setProgressBarTitle "Copiando dados do DB Endereços para memória"
  
  ReDim tbDB(qtLinhas)
  pos = 0
  For lin = 3 To qtLinhas + 2
    setProgressBarUpByOne
    If wsDB.Cells(lin, K_ADR_LOCAL_1) <> vbNullString Then
      pos = pos + 1
      tbDB(pos).matnr = Val(wsDB.Cells(lin, K_ADR_MATNR))
      tbDB(pos).matnr_edit = Format(tbDB(pos).matnr, "000000000000000000")
'      For i = 1 To 5
'        fillAddressStructure wsDB, pos, lin, i
'      Next
      If wsDB.Cells(lin, K_ADR_LOCAL_1) <> vbNullString Then
        tbDB(pos).endereco(1).local = "P" & Mid(wsDB.Cells(lin, K_ADR_LOCAL_1), 1, 1)
        tbDB(pos).endereco(1).corredor = IIf(wsDB.Cells(lin, K_ADR_CORRE_1) = vbNullString, "X", wsDB.Cells(lin, K_ADR_CORRE_1))
        tbDB(pos).endereco(1).altura = IIf(wsDB.Cells(lin, K_ADR_ALTUR_1) = vbNullString, "00", Format(wsDB.Cells(lin, K_ADR_ALTUR_1), "00"))
        tbDB(pos).endereco(1).prateleira = IIf(wsDB.Cells(lin, K_ADR_PRATE_1) = vbNullString, "00", Format(wsDB.Cells(lin, K_ADR_PRATE_1), "00"))
        tbDB(pos).endereco(1).armario = IIf(wsDB.Cells(lin, K_ADR_ARMAR_1) = vbNullString, "00", Format(wsDB.Cells(lin, K_ADR_ARMAR_1), "00"))
        tbDB(pos).endereco(1).cartao = IIf(wsDB.Cells(lin, K_ADR_CARTA_1) = vbNullString, "00", wsDB.Cells(lin, K_ADR_CARTA_1))
        tbDB(pos).endereco(1).posicao_edit = tbDB(pos).endereco(1).local & "-" & _
                                             tbDB(pos).endereco(1).corredor & _
                                             tbDB(pos).endereco(1).altura & "." & _
                                             tbDB(pos).endereco(1).prateleira & "-" & _
                                             tbDB(pos).endereco(1).armario & " - Chave: " & _
                                             tbDB(pos).endereco(1).cartao
      End If
      If wsDB.Cells(lin, K_ADR_LOCAL_2) <> vbNullString Then
        tbDB(pos).endereco(2).local = "P" & Mid(wsDB.Cells(lin, K_ADR_LOCAL_2), 1, 1)
        tbDB(pos).endereco(2).corredor = IIf(wsDB.Cells(lin, K_ADR_CORRE_2) = vbNullString, "X", wsDB.Cells(lin, K_ADR_CORRE_2))
        tbDB(pos).endereco(2).altura = IIf(wsDB.Cells(lin, K_ADR_ALTUR_2) = vbNullString, "00", Format(wsDB.Cells(lin, K_ADR_ALTUR_2), "00"))
        tbDB(pos).endereco(2).prateleira = IIf(wsDB.Cells(lin, K_ADR_PRATE_2) = vbNullString, "00", Format(wsDB.Cells(lin, K_ADR_PRATE_2), "00"))
        tbDB(pos).endereco(2).armario = IIf(wsDB.Cells(lin, K_ADR_ARMAR_2) = vbNullString, "00", Format(wsDB.Cells(lin, K_ADR_ARMAR_2), "00"))
        tbDB(pos).endereco(2).cartao = IIf(wsDB.Cells(lin, K_ADR_CARTA_2) = vbNullString, "00", wsDB.Cells(lin, K_ADR_CARTA_2))
        tbDB(pos).endereco(2).posicao_edit = tbDB(pos).endereco(2).local & "-" & _
                                             tbDB(pos).endereco(2).corredor & _
                                             tbDB(pos).endereco(2).altura & "." & _
                                             tbDB(pos).endereco(2).prateleira & "-" & _
                                             tbDB(pos).endereco(2).armario & " - Chave: " & _
                                             tbDB(pos).endereco(2).cartao
      End If
      If wsDB.Cells(lin, K_ADR_LOCAL_3) <> vbNullString Then
        tbDB(pos).endereco(3).local = "P" & Mid(wsDB.Cells(lin, K_ADR_LOCAL_3), 1, 1)
        tbDB(pos).endereco(3).corredor = IIf(wsDB.Cells(lin, K_ADR_CORRE_3) = vbNullString, "X", wsDB.Cells(lin, K_ADR_CORRE_3))
        tbDB(pos).endereco(3).altura = IIf(wsDB.Cells(lin, K_ADR_ALTUR_3) = vbNullString, "00", Format(wsDB.Cells(lin, K_ADR_ALTUR_3), "00"))
        tbDB(pos).endereco(3).prateleira = IIf(wsDB.Cells(lin, K_ADR_PRATE_3) = vbNullString, "00", Format(wsDB.Cells(lin, K_ADR_PRATE_3), "00"))
        tbDB(pos).endereco(3).armario = IIf(wsDB.Cells(lin, K_ADR_ARMAR_3) = vbNullString, "00", Format(wsDB.Cells(lin, K_ADR_ARMAR_3), "00"))
        tbDB(pos).endereco(3).cartao = IIf(wsDB.Cells(lin, K_ADR_CARTA_3) = vbNullString, "00", wsDB.Cells(lin, K_ADR_CARTA_3))
        tbDB(pos).endereco(3).posicao_edit = tbDB(pos).endereco(3).local & "-" & _
                                             tbDB(pos).endereco(3).corredor & _
                                             tbDB(pos).endereco(3).altura & "." & _
                                             tbDB(pos).endereco(3).prateleira & "-" & _
                                             tbDB(pos).endereco(3).armario & " - Chave: " & _
                                             tbDB(pos).endereco(3).cartao
      End If
      If wsDB.Cells(lin, K_ADR_LOCAL_4) <> vbNullString Then
        tbDB(pos).endereco(4).local = "P" & Mid(wsDB.Cells(lin, K_ADR_LOCAL_4), 1, 1)
        tbDB(pos).endereco(4).corredor = IIf(wsDB.Cells(lin, K_ADR_CORRE_4) = vbNullString, "X", wsDB.Cells(lin, K_ADR_CORRE_4))
        tbDB(pos).endereco(4).altura = IIf(wsDB.Cells(lin, K_ADR_ALTUR_4) = vbNullString, "00", Format(wsDB.Cells(lin, K_ADR_ALTUR_4), "00"))
        tbDB(pos).endereco(4).prateleira = IIf(wsDB.Cells(lin, K_ADR_PRATE_4) = vbNullString, "00", Format(wsDB.Cells(lin, K_ADR_PRATE_4), "00"))
        tbDB(pos).endereco(4).armario = IIf(wsDB.Cells(lin, K_ADR_ARMAR_4) = vbNullString, "00", Format(wsDB.Cells(lin, K_ADR_ARMAR_4), "00"))
        tbDB(pos).endereco(4).cartao = IIf(wsDB.Cells(lin, K_ADR_CARTA_4) = vbNullString, "00", wsDB.Cells(lin, K_ADR_CARTA_4))
        tbDB(pos).endereco(4).posicao_edit = tbDB(pos).endereco(4).local & "-" & _
                                             tbDB(pos).endereco(4).corredor & _
                                             tbDB(pos).endereco(4).altura & "." & _
                                             tbDB(pos).endereco(4).prateleira & "-" & _
                                             tbDB(pos).endereco(4).armario & " - Chave: " & _
                                             tbDB(pos).endereco(4).cartao
      End If
      If wsDB.Cells(lin, K_ADR_LOCAL_5) <> vbNullString Then
        tbDB(pos).endereco(5).local = "P" & Mid(wsDB.Cells(lin, K_ADR_LOCAL_5), 1, 1)
        tbDB(pos).endereco(5).corredor = IIf(wsDB.Cells(lin, K_ADR_CORRE_5) = vbNullString, "X", wsDB.Cells(lin, K_ADR_CORRE_5))
        tbDB(pos).endereco(5).altura = IIf(wsDB.Cells(lin, K_ADR_ALTUR_5) = vbNullString, "00", Format(wsDB.Cells(lin, K_ADR_ALTUR_5), "00"))
        tbDB(pos).endereco(5).prateleira = IIf(wsDB.Cells(lin, K_ADR_PRATE_5) = vbNullString, "00", Format(wsDB.Cells(lin, K_ADR_PRATE_5), "00"))
        tbDB(pos).endereco(5).armario = IIf(wsDB.Cells(lin, K_ADR_ARMAR_5) = vbNullString, "00", Format(wsDB.Cells(lin, K_ADR_ARMAR_5), "00"))
        tbDB(pos).endereco(5).cartao = IIf(wsDB.Cells(lin, K_ADR_CARTA_5) = vbNullString, "00", wsDB.Cells(lin, K_ADR_CARTA_5))
        tbDB(pos).endereco(5).posicao_edit = tbDB(pos).endereco(5).local & "-" & _
                                             tbDB(pos).endereco(5).corredor & _
                                             tbDB(pos).endereco(5).altura & "." & _
                                             tbDB(pos).endereco(5).prateleira & "-" & _
                                             tbDB(pos).endereco(5).armario & " - Chave: " & _
                                             tbDB(pos).endereco(5).cartao
      End If
    End If
  Next
  
  ReDim Preserve tbDB(pos)
  
  
  Set tbRange = tbStock.DataBodyRange
  setProgressCells (K_PROGRESS_MCHB)
  setProgressBarTotal (tbRange.Rows.Count)
  setProgressBarTitle "DB - Adicionando cartões de inventário"
  
  Dim stDB As ty_db
  Dim whichMatnr As String
  ReDim tbCells(0)
  pos2 = ws.Cells(tbRange.Rows.Count + 13, 1).Row
  For Each Row In tbRange.Rows
    setProgressBarUpByOne
    If Val(ws.Cells(Row.Row, K_COL_LIFNR)) = 0 Then
      whichMatnr = ws.Cells(Row.Row, K_COL_MATNR)
      If ws.Cells(Row.Row, K_COL_XCHPF) = "X" And _
         ws.Cells(Row.Row, K_COL_CHARG) <> vbNullString Then
        pos = findMaterialLineByMatnrAndBatch(whichMatnr, ws.Cells(Row.Row, K_COL_CHARG))
      Else
        pos = findMaterialLineByMatnr(whichMatnr)
      End If
      idx = 0
      Do While ws.Cells(pos + 1, K_COL_MATNR) = whichMatnr
        idx = idx + 1
        If idx >= 10 Then
          Exit Do
        End If
        If ws.Cells(pos, K_COL_LIFNR) <> vbNullString Then
          pos = pos + 1
        End If
      Loop
      
      stDB = findInDatabase(whichMatnr)
      
      If stDB.endereco(1).local <> vbNullString Then
        'Colorir diferente as linhas totalizadoras
        Set lineRange = ws.Range("A" & pos).EntireRow
        lineRange.Locked = True
        With lineRange.Interior
          .Pattern = xlSolid
          .PatternColorIndex = xlAutomatic
          .ThemeColor = xlThemeColorAccent4
          .TintAndShade = 0.399975585192419
          .PatternTintAndShade = 0
        End With
      End If
      
      If pos > 0 Then
        For idx = 1 To UBound(stDB.endereco)
          If stDB.endereco(idx).local <> vbNullString Then
            pos2 = pos2 + 1
            
            ws.Cells(pos, K_COL_XCARD) = "H" 'Head
            ws.Cells(pos2, K_COL_XCARD) = "I" 'Item
            ws.Cells(pos2, K_COL_MATNR) = ws.Cells(pos, K_COL_MATNR)
            ws.Cells(pos2, K_COL_ABCIN) = ws.Cells(pos, K_COL_ABCIN)
            ws.Cells(pos2, K_COL_MTART) = ws.Cells(pos, K_COL_MTART)
            ws.Cells(pos2, K_COL_MAKTX) = ws.Cells(pos, K_COL_MAKTX)
            ws.Cells(pos2, K_COL_XCHPF) = ws.Cells(pos, K_COL_XCHPF)
            ws.Cells(pos2, K_COL_LIFNR) = ws.Cells(pos, K_COL_LIFNR)
            ws.Cells(pos2, K_COL_SPERR) = ws.Cells(pos, K_COL_SPERR)
            ws.Cells(pos2, K_COL_MEINS) = ws.Cells(pos, K_COL_MEINS)
            ws.Cells(pos2, K_COL_ZZNSR) = ws.Cells(pos, K_COL_ZZNSR)
            ws.Cells(pos2, K_COL_STEUC) = ws.Cells(pos, K_COL_STEUC)
            ws.Cells(pos2, K_COL_MCLAS) = ws.Cells(pos, K_COL_MCLAS)
            ws.Cells(pos2, K_COL_MTUSE) = ws.Cells(pos, K_COL_MTUSE)
            ws.Cells(pos2, K_COL_MTORG) = ws.Cells(pos, K_COL_MTORG)
            ws.Cells(pos2, K_COL_VPRSV) = ws.Cells(pos, K_COL_VPRSV)
            ws.Cells(pos2, K_COL_UPRIC_S) = ws.Cells(pos, K_COL_UPRIC_S)
            ws.Cells(pos2, K_COL_UPRIC_V) = ws.Cells(pos, K_COL_UPRIC_V)
            ws.Cells(pos2, K_COL_CHARG) = ws.Cells(pos, K_COL_CHARG)
            ws.Cells(pos2, K_COL_LABST) = ws.Cells(pos, K_COL_LABST)
            ws.Cells(pos2, K_COL_INSME) = ws.Cells(pos, K_COL_INSME)
            ws.Cells(pos2, K_COL_SPEME) = ws.Cells(pos, K_COL_SPEME)
            ws.Cells(pos2, K_COL_LGORT) = stDB.endereco(idx).posicao_edit
          End If
        Next
      End If
    End If
  Next
  
  tbStock.Sort.SortFields.Clear
  tbStock.Sort.SortFields.Add key:=Range("tbStock[Material]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  tbStock.Sort.SortFields.Add key:=Range("tbStock[Batch]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  tbStock.Sort.SortFields.Add key:=Range("tbStock[Chave]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  With tbStock.Sort
      .Header = xlYes
      .MatchCase = False
      .Orientation = xlTopToBottom
      .SortMethod = xlPinYin
      .Apply
  End With

  wbDB.Close False
  
  Set tbRange = tbStock.DataBodyRange
  
  setProgressBarTitle ("Gera fórmulas de soma")
  setProgressBarTotal (tbRange.Rows.Count)
  
  For Each Row In tbRange.Rows
    setProgressBarUpByOne
    If ws.Cells(Row.Row, K_COL_LGORT) = "H" Then
      whichMatnr = ws.Cells(Row.Row, K_COL_MATNR)
      stDB = findInDatabase(whichMatnr)
      If ws.Cells(Row.Row, K_COL_XCHPF) = "X" And _
         ws.Cells(Row.Row, K_COL_CHARG) <> vbNullString Then
        pos = findMaterialLineByMatnrAndBatch(whichMatnr, ws.Cells(Row.Row, K_COL_CHARG))
      Else
        pos = findMaterialLineByMatnr(whichMatnr)
      End If
      idx = 0
      Do While ws.Cells(pos + 1, K_COL_MATNR) = whichMatnr
        idx = idx + 1
        If idx >= 10 Then
          Exit Do
        End If
        If ws.Cells(pos, K_COL_LGORT) <> vbNullString Or _
          (ws.Cells(pos, K_COL_LGORT) = vbNullString And ws.Cells(pos, K_COL_LIFNR) <> vbNullString) Then
          pos = pos + 1
        End If
      Loop
        
      If stDB.endereco(1).local <> vbNullString And pos > 0 Then
        setupTotalLine pos, stDB
      End If
    End If
  Next

  
  setProgressBarTitle "MCHB/MSLB/DB Estoque - Finalizado"
End Sub

Sub fillAddressStructure(ByRef wsDB As Worksheet, ByVal pos As Integer, ByVal lin As Integer, ByVal cardNumber As Integer)

'  Dim commandLine(7) As String
'
'  If Eval("wsDB.Cells(lin, K_ADR_LOCAL_" & cardNumber & ")") <>  "" Then
'  commandLine(2) = "tbDB(pos).endereco(" & cardNumber & ").local = ""P"" & Mid(wsDB.Cells(lin, K_ADR_LOCAL_" & cardNumber & "), 1, 1)"
'  commandLine(3) = "tbDB(pos).endereco(" & cardNumber & ").corredor = IIf(wsDB.Cells(lin, K_ADR_CORRE_" & cardNumber & ") = """", ""X"", wsDB.Cells(lin, K_ADR_CORRE_" & cardNumber & "))"
'  commandLine(4) = "tbDB(pos).endereco(" & cardNumber & ").altura = IIf(wsDB.Cells(lin, K_ADR_ALTUR_" & cardNumber & ") = """", ""00"", Format(wsDB.Cells(lin, K_ADR_ALTUR_" & cardNumber & "), ""00""))"
'  commandLine(5) = "tbDB(pos).endereco(" & cardNumber & ").prateleira = IIf(wsDB.Cells(lin, K_ADR_PRATE_" & cardNumber & ") = """", ""00"", Format(wsDB.Cells(lin, K_ADR_PRATE_" & cardNumber & "), ""00""))"
'  commandLine(6) = "tbDB(pos).endereco(" & cardNumber & ").armario = IIf(wsDB.Cells(lin, K_ADR_ARMAR_" & cardNumber & ") = """", ""00"", Format(wsDB.Cells(lin, K_ADR_ARMAR_" & cardNumber & "), ""00""))"
'  commandLine(7) = "tbDB(pos).endereco(" & cardNumber & ").cartao = IIf(wsDB.Cells(lin, K_ADR_CARTA_" & cardNumber & ") = """", ""00"", wsDB.Cells(lin, K_ADR_CARTA_" & cardNumber & "))"
'  commandLine(8) = "tbDB(pos).endereco(" & cardNumber & ").posicao_edit = tbDB(pos).endereco(" & cardNumber & ").local & ""-"" " & _
'                                             "tbDB(pos).endereco(" & cardNumber & ").corredor " & _
'                                             "tbDB(pos).endereco(" & cardNumber & ").altura & ""."" " & _
'                                             "tbDB(pos).endereco(" & cardNumber & ").prateleira & ""-"" " & _
'                                             "tbDB(pos).endereco(" & cardNumber & ").armario &  "" - Chave:  "" " & _
'                                             "tbDB(pos).endereco(" & cardNumber & ").cartao"
'  End If
  
End Sub


Sub setupTotalLine(pos As Integer, stDB As ty_db)
  Dim numberOfCards As Integer
  Dim rg1stCount As Range
  Dim rg2ndCount As Range
  Dim rg3rdCount As Range
  Dim firstCountLetter As String
  Dim secondCountLetter As String
  Dim thirdCountLetter As String
  Dim firstRangeString As String
  Dim secondRangeString As String
  Dim thirdRangeString As String
  
  numberOfCards = 0
  For i = 1 To UBound(stDB.endereco)
    If stDB.endereco(i).local <> vbNullString Then
      numberOfCards = numberOfCards + 1
    End If
  Next
  
  firstCountLetter = convertNumberToLetter(K_COL_MENGE_1ST)
  secondCountLetter = convertNumberToLetter(K_COL_MENGE_2ND)
  thirdCountLetter = convertNumberToLetter(K_COL_MENGE_3RD)
  
  Set rg1stCount = Range(firstCountLetter & pos)
  Set rg2ndCount = Range(secondCountLetter & pos)
  Set rg3rdCount = Range(thirdCountLetter & pos)
  
  firstRangeString = "R[-" & numberOfCards & "]C" & K_COL_MENGE_1ST & ":R[-1]C" & K_COL_MENGE_1ST
  secondRangeString = "R[-" & numberOfCards & "]C" & K_COL_MENGE_2ND & ":R[-1]C" & K_COL_MENGE_2ND
  thirdRangeString = "R[-" & numberOfCards & "]C" & K_COL_MENGE_3RD & ":R[-1]C" & K_COL_MENGE_3RD
  
  rg1stCount.FormulaR1C1Local = "=IF(SUM(" & firstRangeString & ")=0,"""",SUM(" & firstRangeString & "))"
  rg2ndCount.FormulaR1C1Local = "=IF(SUM(" & secondRangeString & ")=0,"""",SUM(" & secondRangeString & "))"
  rg3rdCount.FormulaR1C1Local = "=IF(SUM(" & thirdRangeString & ")=0,"""",SUM(" & thirdRangeString & "))"

  Dim localRange As Range
  Set localRange = Range(firstCountLetter & (pos - numberOfCards) & ":" & firstCountLetter & (pos - 1))
  localRange.FormulaR1C1 = vbNullString
  Set localRange = Range(secondCountLetter & (pos - numberOfCards) & ":" & secondCountLetter & (pos - 1))
  localRange.FormulaR1C1 = vbNullString
  Set localRange = Range(thirdCountLetter & (pos - numberOfCards) & ":" & thirdCountLetter & (pos - 1))
  localRange.FormulaR1C1 = vbNullString
  

End Sub

Sub loadMBEWFields(sapConn As Object)

  Dim lin As Integer

  ReDim field(10)
  field(1) = "MATNR"
  field(2) = "MTUSE"
  field(3) = "MTORG"
  field(4) = "VPRSV"
  field(5) = "VERPR"
  field(6) = "STPRS"
  field(7) = "PEINH"
  field(8) = "SALK3"
  field(9) = "LBKUM"
  field(10) = "PSTAT"
  
  ReDim filtro(tbStock.DataBodyRange.Rows.Count)
  setProgressCells (K_PROGRESS_MBEW)
  setProgressBarTotal (tbStock.DataBodyRange.Rows.Count)
  setProgressBarTitle "Montagem do Filtro de Pesquisa"

  lin = tbStock.Range.Cells(1, 1).Row + 1
  filterline = 0
  For Each lrow In tbStock.DataBodyRange.Rows
    setProgressBarUpByOne
    filterline = filterline + 1
    If lrow.Row = lin Then
      filtro(filterline) = "BWKEY = '" & Plant & "' AND (MATNR = '" & Format(ws.Cells(lrow.Row, K_COL_MATNR), "000000000000000000") & "'"
    Else
      filtro(filterline) = "OR MATNR = '" & Format(ws.Cells(lrow.Row, K_COL_MATNR), "000000000000000000") & "'"
    End If
  Next
  filtro(filterline) = filtro(filterline) & ")"

  setProgressBarTitle "Buscando informações de preço do material"
  resultDoc = readTable(sapConn, "MBEW", field, filtro)

  setProgressBarTotal (tbStock.DataBodyRange.Rows.Count)
  setProgressBarTitle "Populando informações de preço do material"


  Dim pos As Integer
  Dim whichMatnr As String
  Dim dPriceStd As Double  'Unitary Standard Price
  Dim dPriceMap As Double  'Unitary Moving Average Price
  Dim dStockCost As Double
  Dim dSalk3 As Double
  Dim dLbkum As Double
  Dim dPer As Double
  Dim sUse As String
  Dim sOrigin As String
  pos = 0
  whichMatnr = vbNullString
  For lin = 1 To UBound(resultDoc)
    setProgressBarUpByOne
    whichMatnr = Val(resultDoc(lin, 1))
    
    If whichMatnr = "759100" Or _
       whichMatnr = "760303" Or _
       whichMatnr = "760306" Then
       x = 1
    End If
    
    pos = findMaterialLineByMatnr(whichMatnr)
    If pos > 0 Then
'      If ws.Cells(pos, K_COL_VPRSV) = "V" Then
'        dPrice = resultDoc(lin, 5) / 100
'      Else
'        dPrice = resultDoc(lin, 6) / 100
'      End If
      sUse = resultDoc(lin, 2)
      sOrigin = resultDoc(lin, 3)
      ws.Cells(pos, K_COL_MTUSE) = getMaterialUseDescription(sUse)
      ws.Cells(pos, K_COL_MTORG) = getMaterialOriginDescription(sOrigin)
      ws.Cells(pos, K_COL_VPRSV) = resultDoc(lin, 4)
      dPriceMap = CDbl(resultDoc(lin, 5)) / IIf(isCommaTheDecimalSeparator, 100, 1)
      dPriceStd = CDbl(resultDoc(lin, 6)) / IIf(isCommaTheDecimalSeparator, 100, 1)
      dPer = resultDoc(lin, 7)
      dSalk3 = CDbl(resultDoc(lin, 8)) / IIf(isCommaTheDecimalSeparator, 100, 1)
      dLbkum = CDbl(resultDoc(lin, 9)) / IIf(isCommaTheDecimalSeparator, 1000, 1)
      If dLbkum = 0 Then
        'dStockCost = 0
        dStockCost = dPriceStd / dPer 'Caso nao exista estoque de planta, assume custo unitario do mestre de materiais
      Else
        dStockCost = dSalk3 / dLbkum
      End If
      'ws.Cells(pos, K_COL_SDSTD) = dStockCost
      'ws.Cells(pos, K_COL_UPRIC_S) = dPriceStd / dPer
      If ws.Cells(pos, K_COL_VPRSV) = "S" Then
        ws.Cells(pos, K_COL_UPRIC_S) = dStockCost
      Else
        ws.Cells(pos, K_COL_UPRIC_S) = dPriceMap / dPer
      End If
      
      ws.Cells(pos, K_COL_SALK3) = dSalk3
      ws.Cells(pos, K_COL_LBKUM) = dLbkum
      ws.Cells(pos, K_COL_PSTAT) = resultDoc(lin, 10)
    End If
  Next
  
  setProgressBarTotal (tbStock.DataBodyRange.Rows.Count)
  setProgressBarTitle "Identifica Materiais não extendidos na planta " & Plant
  For lin = tbRange.Cells(1, 1).Row To tbStock.Range(tbStock.DataBodyRange.Rows.Count, 1).Row + 1
    setProgressBarUpByOne
    If ws.Cells(lin, K_COL_MTUSE) = vbNullString Then
      ws.Cells(lin, K_COL_MTUSE) = "N.Ext."
      ws.Cells(lin, K_COL_MTORG) = "N.Ext."
      ws.Cells(lin, K_COL_VPRSV) = "N.Ext."
      ws.Cells(lin, K_COL_UPRIC_S) = "N.Ext."
      'ws.Cells(lin, K_COL_UPRIC_V) = "N.Ext."
    End If
  Next
  
  
  setProgressBarTitle "MBEW - Finalizado"

End Sub


Sub loadMAKTFields(sapConn As Object)

  Dim lin As Integer


  ReDim field(2)
  field(1) = "MATNR"
  field(2) = "MAKTX"
  

  ReDim filtro(tbStock.DataBodyRange.Rows.Count)
  setProgressCells (K_PROGRESS_MAKT)
  setProgressBarTotal (tbStock.DataBodyRange.Rows.Count)
  setProgressBarTitle "Montagem do Filtro de Pesquisa"
  
  Set tbStock = ws.ListObjects("tbStock")
  Set tbRange = tbStock.DataBodyRange
  lin = tbRange.Cells(1, 1).Row
  filterline = 0
  For Each lrow In tbStock.DataBodyRange.Rows
    setProgressBarUpByOne
    filterline = filterline + 1
    If lrow.Row = lin Then
      filtro(filterline) = "SPRAS = 'P' AND (MATNR = '" & Format(ws.Cells(lrow.Row, K_COL_MATNR), "000000000000000000") & "'"
    Else
      filtro(filterline) = "OR MATNR = '" & Format(ws.Cells(lrow.Row, K_COL_MATNR), "000000000000000000") & "'"
    End If
  Next
  filtro(filterline) = filtro(filterline) & ")"

  setProgressBarTitle "Buscando descrição em Português do material"
  resultDoc = readTable(sapConn, "MAKT", field, filtro)
  
  setProgressBarTotal (tbStock.DataBodyRange.Rows.Count)
  setProgressBarTitle "Populando descrição em PT do material"
  Dim pos As Integer
  Dim whichMatnr As String
  pos = 0
  whichMatnr = vbNullString
  For lin = 1 To UBound(resultDoc)
    setProgressBarUpByOne
    whichMatnr = Val(resultDoc(lin, 1))
    pos = findMaterialLineByMatnr(whichMatnr)
    If pos > 0 Then
      ws.Cells(pos, K_COL_MAKTX) = UCase(resultDoc(lin, 2))
    End If
  Next
  
  setProgressBarTotal (tbStock.DataBodyRange.Rows.Count)
  setProgressBarTitle "Identifica Materiais sem descrição em Portugês"
  For Each Row In tbStock.DataBodyRange.Rows
  'For lin = tbRange.Cells(1, 1).row To tbStock.Range(tbStock.DataBodyRange.Rows.Count, 1).row + 1
    setProgressBarUpByOne
    If ws.Cells(Row.Row, K_COL_MAKTX) = vbNullString Then
      ws.Cells(Row.Row, K_COL_MAKTX) = "<Sem Descrição em Português>"
    End If
  Next
  
  setProgressBarTitle "MAKT - Finalizado"

End Sub


Sub getShipmentDetails(sapConn As Object)

  Dim resultShipment() As String
  Dim resultHistory() As String
  Dim field() As String
  Dim filtro() As String
  Dim shipment As Integer
  Dim history As Integer
  Dim lin As Integer

  Set ws = ThisWorkbook.Sheets("NFWriter")
  ws.Activate


  onlyManual = Range(ThisWorkbook.Names("manualOnly")).Value
  StartDate = Range(ThisWorkbook.Names("start_date")).Value
  EndDate = Range(ThisWorkbook.Names("end_date")).Value

  ReDim field(48)
  Dim fieldQty As Integer
  fieldQty = UBound(field)

  field(1) = "DOCNUM"
  field(2) = "NFTYPE"
  field(3) = "DOCTYP"
  field(4) = "DIRECT"
  field(5) = "DOCDAT"
  field(6) = "PSTDAT"
  field(7) = "CREDAT"
  field(8) = "CRETIM"
  field(9) = "CRENAM"
  field(10) = "CHADAT"
  field(11) = "CHATIM"
  field(12) = "CHANAM"
  field(13) = "FORM"
  field(14) = "MODEL"
  field(15) = "SERIES"
  field(16) = "NFNUM"
  field(17) = "ENTRAD"
  field(18) = "FATURA"
  field(19) = "ZTERM"
  field(20) = "PRINTD"
  field(21) = "MANUAL"
  field(22) = "WAERK"
  field(23) = "BELNR"
  field(24) = "GJAHR"
  field(25) = "BUKRS"
  field(26) = "BRANCH"
  field(27) = "PARVW"
  field(28) = "PARID"
  field(29) = "PARTYP"
  field(30) = "CANCEL"
  field(31) = "DOCREF"
  field(32) = "BRGEW"
  field(33) = "NTGEW"
  field(34) = "NFTOT"
  field(35) = "NFENUM"
  field(36) = "AUTHCOD"
  field(37) = "DOCSTAT"
  field(38) = "XMLVERS"
  field(39) = "CODE"
  field(40) = "NAME1"
  field(41) = "STRAS"
  field(42) = "ORT01"
  field(43) = "ORT02"
  field(44) = "REGIO"
  field(45) = "LAND1"
  field(46) = "CGC"
  field(47) = "CPF"
  field(48) = "NATOP"
  
  
  sStartDate = Year(StartDate) & Format(Month(StartDate), "00") & Format(Day(StartDate), "00")
  sEndDate = Year(EndDate) & Format(Month(EndDate), "00") & Format(Day(EndDate), "00")
  
  ReDim filtro(2)
  filtro(1) = "BUKRS = '" & CompanyCode & "' AND"
  filtro(2) = "CREDAT BETWEEN '" & sStartDate & "' AND '" & sEndDate & "' "
  
  If (onlyManual = "Sim") Then
    ReDim Preserve filtro(3)
    filtro(3) = "AND MANUAL = 'X'"
  End If
   
  resultDoc = readTable(sapConn, "J_1BNFDOC", field, filtro)
  
  ReDim field(12)
  ReDim filtro(2)
  lin = 13
  ws.Cells(lin, 1) = "DocNum"
  ws.Cells(lin, 2) = "NF Type"
  ws.Cells(lin, 3) = "Doc.Type"
  ws.Cells(lin, 4) = "Direction"
  ws.Cells(lin, 5) = "Doc.Date"
  ws.Cells(lin, 6) = "Posting Date"
  ws.Cells(lin, 7) = "Create Date"
  ws.Cells(lin, 8) = "Create Time"
  ws.Cells(lin, 9) = "Create Name"
  ws.Cells(lin, 10) = "Change Date"
  ws.Cells(lin, 11) = "Change Time"
  ws.Cells(lin, 12) = "Change Name"
  ws.Cells(lin, 13) = "Form"
  ws.Cells(lin, 14) = "Model"
  ws.Cells(lin, 15) = "Series"
  ws.Cells(lin, 16) = "NF Num"
  ws.Cells(lin, 17) = "Entrada"
  ws.Cells(lin, 18) = "Fatura"
  ws.Cells(lin, 19) = "Pmt.Term"
  ws.Cells(lin, 20) = "Printed"
  ws.Cells(lin, 21) = "Manual"
  ws.Cells(lin, 22) = "Moeda"
  ws.Cells(lin, 23) = "Acct.Document"
  ws.Cells(lin, 24) = "Acct.Doc.Year"
  ws.Cells(lin, 25) = "Company Code"
  ws.Cells(lin, 26) = "Branch"
  ws.Cells(lin, 27) = "Partner Function"
  ws.Cells(lin, 28) = "Partner ID"
  ws.Cells(lin, 29) = "Partner Type"
  ws.Cells(lin, 30) = "Cancel"
  ws.Cells(lin, 31) = "Reference Document"
  ws.Cells(lin, 32) = "Gross Weight"
  ws.Cells(lin, 33) = "Net Weight"
  ws.Cells(lin, 34) = "Total Amount (incl Tax)"
  ws.Cells(lin, 35) = "NFe Number"
  ws.Cells(lin, 36) = "Auth.Code"
  ws.Cells(lin, 37) = "Doc.Status"
  ws.Cells(lin, 38) = "XML Version"
  ws.Cells(lin, 39) = "Code"
  ws.Cells(lin, 40) = "Vendor Name"
  ws.Cells(lin, 41) = "Address"
  ws.Cells(lin, 42) = "City"
  ws.Cells(lin, 43) = "District"
  ws.Cells(lin, 44) = "State"
  ws.Cells(lin, 45) = "Country"
  ws.Cells(lin, 46) = "CNPJ"
  ws.Cells(lin, 47) = "CPF"
  ws.Cells(lin, 48) = "Operation Nature"
  
  
  If IsArrayEmpty(resultDoc) Then
    Exit Sub
  End If

  
  Set cellCurrent = Range(ThisWorkbook.Names("auxdata_currentQty"))
  Set cellTotal = Range(ThisWorkbook.Names("auxdata_totalQty"))
  Set cellPercent = Range(ThisWorkbook.Names("auxdata_percentCompleted"))
  
  cellCurrent.Value = 0
  cellTotal.Value = UBound(resultDoc)
  cellPercent.Value = 0


  For doc = 1 To UBound(resultDoc)
    
    addUser (resultDoc(doc, 9))
    addUser (resultDoc(doc, 12))
    addNFType (resultDoc(doc, 2))
    cellCurrent.Value = doc
    cellPercent.Value = doc / UBound(resultDoc)
  Next

        
  Set cellCurrent = Range(ThisWorkbook.Names("process_currentQty"))
  Set cellTotal = Range(ThisWorkbook.Names("process_totalQty"))
  Set cellPercent = Range(ThisWorkbook.Names("process_percentCompleted"))
    
  cellCurrent.Value = 0
  cellTotal.Value = UBound(resultDoc)
  cellPercent.Value = 0
  
    Dim lValue As String
    For doc = 1 To UBound(resultDoc)
      cellCurrent.Value = doc
      cellPercent.Value = doc / UBound(resultDoc)
      
  
        lin = lin + 1
        For i = 1 To fieldQty
            Select Case i
                Case 2
                    lValue = resultDoc(doc, i)
                    ws.Cells(lin, i) = Trim(lValue) & " - " & getNFType(lValue)
                Case 9, 12
                    lValue = resultDoc(doc, i)
                    If (Trim(lValue) = vbNullString) Then
                        ws.Cells(lin, i) = vbNullString
                    Else
                        ws.Cells(lin, i) = Trim(lValue) & " - " & getUser(lValue)
                    End If
                Case 5, 6, 7, 10
                    lValue = resultDoc(doc, i)
                    lValue = Mid(lValue, 7, 2) & "/" & Mid(lValue, 5, 2) & "/" & Mid(lValue, 1, 4)
                    If lValue = "00/00/0000" Then
                        ws.Cells(lin, i) = vbNullString
                    Else
                        ws.Cells(lin, i) = CDate(lValue)
                    End If
                Case 8, 11
                    lValue = resultDoc(doc, i)
                    lValue = Mid(lValue, 1, 2) & ":" & Mid(lValue, 3, 2) & ":" & Mid(lValue, 5, 2)
                    If lValue = "00:00:00" Then
                        ws.Cells(lin, i) = vbNullString
                    Else
                        ws.Cells(lin, i) = TimeValue(lValue)
                    End If
                Case Else
                    ws.Cells(lin, i) = resultDoc(doc, i)
            End Select
        Next
        'DoEvents
    Next
  
End Sub

Function getNFType(iNFType As String) As String

    For i = 1 To UBound(tbNFType)
        If tbNFType(i).type = iNFType Then
            getNFType = tbNFType(i).description
            Exit Function
        End If
    Next

End Function

Function getUser(iUser As String) As String

    For i = 1 To UBound(tbUser)
        If tbUser(i).user = Trim(iUser) Then
            getUser = tbUser(i).name
            Exit Function
        End If
    Next

End Function
Sub addUser(iUser As String)
    
    Dim ub As Integer
    Dim lEmptyArray As Boolean
    
    On Error Resume Next
    ub = UBound(tbUser)
    If Err.number > 0 Then
        lEmptyArray = True
        lSize = 0
    Else
        lEmptyArray = False
    End If
    On Error GoTo 0
    
    If Not lEmptyArray Then
        For i = 1 To UBound(tbUser)
            If tbUser(i).user = iUser Then
                Exit Sub
            End If
        Next
    End If
    
    'Dim lSize As Integer
    
    On Error Resume Next
    lSize = UBound(tbUser) + 1
    If Err.number > 0 Then
        lSize = 1
    End If
    On Error GoTo 0
    
    ReDim Preserve tbUser(lSize)
    
    tbUser(lSize).user = Trim(iUser)
    
    If tbUser(lSize).user = vbNullString Then
        tbUser(lSize).name = vbNullString
        Exit Sub
    End If
    
    Dim field() As String
    Dim filtro() As String
    
    ReDim field(2)
    ReDim filtro(1)

    field(1) = "ADDRNUMBER"
    field(2) = "PERSNUMBER"
    filtro(1) = "BNAME = '" & tbUser(lSize).user & "'"
    result = readTable(sapConn, "USR21", field, filtro)
    
    ReDim field(1)
    ReDim filtro(2)
    field(1) = "NAME_TEXT"
    filtro(1) = "ADDRNUMBER = '" & result(1, 1) & "' AND "
    filtro(2) = "PERSNUMBER = '" & result(1, 2) & "' "
    result = readTable(sapConn, "V_ADDR_USR", field, filtro)
    
    tbUser(lSize).name = result(1, 1)
    
End Sub

Sub addNFType(iNFType As String)
    
    Dim ub As Integer
    Dim lEmptyArray As Boolean
    
    On Error Resume Next
    ub = UBound(tbNFType)
    If Err.number > 0 Then
        lEmptyArray = True
    Else
        lEmptyArray = False
    End If
    On Error GoTo 0
    
    If Not lEmptyArray Then
        For i = 1 To UBound(tbNFType)
            If tbNFType(i).type = iNFType Then
                Exit Sub
            End If
        Next
    End If
    
    Dim lSize As Integer
    
    On Error Resume Next
    lSize = UBound(tbNFType) + 1
    If Err.number > 0 Then
        lSize = 1
    End If
    On Error GoTo 0
    
    ReDim Preserve tbNFType(lSize)
    
    tbNFType(lSize).type = iNFType
    
    Dim field() As String
    Dim filtro() As String
    
    ReDim field(1)
    ReDim filtro(2)

    field(1) = "NFTTXT"
    filtro(1) = "SPRAS = 'P' AND "
    filtro(2) = "NFTYPE = '" & tbNFType(lSize).type & "'"
    result = readTable(sapConn, "J_1BAAT", field, filtro)
    
    tbNFType(lSize).description = result(1, 1)
    
End Sub


Function searchResultShipment(ByRef shipments() As String, po As String, item As String) As Integer

    Dim lin As Integer
    searchResultShipment = 0
    For lin = 1 To UBound(shipments)
        If shipments(lin, 4) = po And shipments(lin, 5) = item Then
           searchResultShipment = lin
           Exit Function
        End If
    Next

End Function

Function getDebCred(shkzg As String) As String

    If shkzg = "S" Then
        getDebCred = "Deb"
    Else
        getDebCred = "Cred"
    End If

End Function

Function getHistType(tp As String) As String

    Select Case tp
        Case "E"
            getHistType = "GR"
        Case "F"
            getHistType = "DCGR"
        Case "G"
            getHistType = "DCIR"
        Case "L"
            getHistType = "DlNt"
        Case "N"
            getHistType = "SD-L"
        Case "Q"
            getHistType = "IR-L"
        Case Else
            getHistType = tp
    End Select
End Function

Function getRightAmount(amt, debCred As String) As Double

    getRightAmount = amt
    If debCred = "Cred" Then
        getRightAmount = getRightAmount * -1
    End If
End Function


Function getDate(dt As Date) As String
  getDate = Year(dt) & Format(Month(dt), "00") & Format(Day(dt), "00")
End Function

Function getCompanyCode(comp As String) As String
  getCompanyCode = Mid(comp, 1, 4)
End Function

Function getDocumentType(tp As String) As String
  getDocumentType = Mid(tp, 1, 2)
End Function

Function getAccountType(tp As String) As String
  getAccountType = Mid(tp, 1, 1)
End Function

Function getAccount(acct As String) As String
  getAccount = Mid(acct, 1, 10)
End Function

Function getCustomer(cust As String) As String
  getCustomer = Mid(cust, 1, 10)
End Function

Function getVendor(vendor As String) As String
  getVendor = Mid(vendor, 1, 10)
End Function

Function getCostCenter(ccenter As String) As String
  getCostCenter = Mid(ccenter, 8, 10)
End Function

Sub Help()
  frmHelp.Show
End Sub

Sub fillParameters(sapConn As Object, whichEvent As String)

  Dim result() As String
  Dim field() As String
  Dim filtro() As String

  ReDim field(2)
  ReDim filtro(1)

  setProgressBarUpByOne
  setProgressBarTitle "Busca dados do usuário (1/2)"
  field(1) = "ADDRNUMBER"
  field(2) = "PERSNUMBER"
  filtro(1) = "BNAME = '" & sapConn.Connection.user & "'"
  result = readTable(sapConn, "USR21", field, filtro)

  setProgressBarUpByOne
  setProgressBarTitle "Busca nome do usuário (2/2)"
  ReDim field(1)
  ReDim filtro(2)
  field(1) = "NAME_TEXT"
  filtro(1) = "ADDRNUMBER = '" & result(1, 1) & "' AND "
  filtro(2) = "PERSNUMBER = '" & result(1, 2) & "' "
  result = readTable(sapConn, "V_ADDR_USR", field, filtro)

  Dim rParameters As Range
  Select Case (whichEvent)
    Case K_EVENT_STOCKREAD
      Set rParameters = Range(ThisWorkbook.Names("stock_read"))
    Case K_EVENT_INVENTORY
      Set rParameters = Range(ThisWorkbook.Names("open_inventory"))
    Case K_EVENT_1ST_COUNT
      Set rParameters = Range(ThisWorkbook.Names("first_count"))
    Case K_EVENT_2ND_COUNT
      Set rParameters = Range(ThisWorkbook.Names("second_count"))
    Case K_EVENT_3RD_COUNT
      Set rParameters = Range(ThisWorkbook.Names("third_count"))
    Case K_EVENT_APPROVAL
      Set rParameters = Range(ThisWorkbook.Names("count_approver"))
    Case K_EVENT_NOTAFISCAL
      Set rParameters = Range(ThisWorkbook.Names("nf_posting"))
  End Select
  
  Dim sDate As String
  Dim sTime As String
  
  sDate = Format(Date, "dd.MM.yyyy")
  sTime = Format(Time, "hh:mm:ss")
  
  rParameters.Value = sapConn.Connection.user & " - " & result(1, 1) & " - " & sDate & " " & sTime

  setProgressBarUpByOne
  setProgressBarTitle "Busca dados da Filial a partir da Planta"
  
  loadKeyFields

  ReDim field(1)
  ReDim filtro(1)
  Dim sBusinessPlace As String
  field(1) = "J_1BBRANCH"
  filtro(1) = "WERKS = '" & Plant & "'"
  result = readTable(sapConn, "T001W", field, filtro)

  
  setProgressBarUpByOne
  setProgressBarTitle "Busca nome da Filial"
  
  sBusinessPlace = Format(result(1, 1), "0000")
  Branch = sBusinessPlace
  
  ReDim field(1)
  ReDim filtro(2)
  field(1) = "NAME"
  filtro(1) = "BUKRS = '" & CompanyCode & "' AND "
  filtro(2) = "BRANCH = '" & sBusinessPlace & "'"
  result = readTable(sapConn, "J_1BBRANCH", field, filtro)
  
  Range(ThisWorkbook.Names("businessPlace")).Value = sBusinessPlace & " - " & result(1, 1)
  Range(ThisWorkbook.Names("running_environment")).Value = sapConn.Connection.system & " - " & sapConn.Connection.client & " - " & sapConn.Connection.Language
  

End Sub

Sub loadKeyFields()
  
  CompanyCode = Mid(Range(ThisWorkbook.Names("company_code")).Value, 1, 4)
  Plant = Mid(Range(ThisWorkbook.Names("Plant")).Value, 1, 4)
  StLoc = Mid(Range(ThisWorkbook.Names("StorageLocation")).Value, 1, 4)

End Sub

Function readTable(sapConn As Object, table As String, fields As Variant, filter() As String) As String()

  If Not IsArray(fields) Then
    MsgBox "Erro: Nao foi passada array para leitura da tabela: " & table
    Exit Function
  End If

  On Error Resume Next
  cellDesc.Value = "(não resposivo)" & cellDesc.Value
  On Error GoTo 0

  Set rfcObj = sapConn.Add("RFC_READ_TABLE")
  Dim objQueryTab, objRowCount As Object
  Set objQueryTab = rfcObj.Exports("QUERY_TABLE")
  Set objRowCount = rfcObj.Exports("ROWCOUNT")

  objQueryTab.Value = table
  ' objRowCount.Value = "10"

  Dim objOptTab, objFldTab, objDatTab As Object
  Set objOptTab = rfcObj.Tables("OPTIONS")
  Set objFldTab = rfcObj.Tables("FIELDS")
  Set objDatTab = rfcObj.Tables("DATA")
  'First we set the condition
  'Refresh table
  objOptTab.freetable
  objFldTab.freetable
  objDatTab.freetable
  'Then set values

  If Not IsArrayEmpty(filter) Then
    For lin = 1 To UBound(filter)
      objOptTab.Rows.Add
      objOptTab(objOptTab.RowCount, "TEXT") = filter(lin)
    Next
  End If
  'objOptTab.Rows.Add
  'objOptTab(objOptTab.RowCount, "TEXT") = "MESTYP = 'ORDERS' and "
  'objOptTab.Rows.Add
  'objOptTab(objOptTab.RowCount, "TEXT") = "STATUS = '53'"

  'Next we set fields to obtain
  'Refresh table
  objFldTab.freetable
  'Then set values
  For fld = 1 To UBound(fields)
    objFldTab.Rows.Add
    objFldTab(objFldTab.RowCount, "FIELDNAME") = fields(fld)
  Next

  If rfcObj.Call = False Then
     MsgBox rfcObj.Exception
  End If


  Dim s As String
  s = vbNullString

  Dim objDatRec As Object
  Dim objFldRec As Object

  Dim retTable() As String
  
  If objDatTab.Rows.Count = 0 Then 'Não retornou conteudo
    readTable = retTable
    Exit Function
  End If

  ReDim retTable(1 To objDatTab.Rows.Count, 1 To UBound(fields))
  
  'On Error Resume Next
  'Set cellCurrent = Range(ThisWorkbook.Names(table & "_currentQty"))
  'If cellCurrent Is Nothing Then
  '  Set cellCurrent = Range(ThisWorkbook.Names("currentQty"))
  'End If
  'Set cellTotal = Range(ThisWorkbook.Names(table & "_totalQty"))
  'If cellTotal Is Nothing Then
  '  Set cellTotal = Range(ThisWorkbook.Names("totalQty"))
  'End If
  'Set cellPercent = Range(ThisWorkbook.Names(table & "_percentCompleted"))
  'If cellPercent Is Nothing Then
  '  Set cellPercent = Range(ThisWorkbook.Names("percentCompleted"))
  'End If
  'On Error GoTo 0
  
  'cellCurrent.Value = 0
  'cellTotal.Value = objDatTab.Rows.Count
  'cellPercent.Value = 0
  
  lin = 0
  For Each objDatRec In objDatTab.Rows
     lin = lin + 1
     For fld = 1 To UBound(fields)
       Set objFldRec = objFldTab.Rows(fld)
       campo = Mid(objDatRec("WA"), objFldRec("OFFSET") + 1, objFldRec("LENGTH")) 'objFldTab.Rows(fld)
       retTable(lin, fld) = campo
     Next
     'cellCurrent.Value = lin
     'cellPercent.Value = (lin / objDatTab.Rows.Count)
     DoEvents
  Next

  Set objOptTab = Nothing
  Set objFldTab = Nothing
  Set objDatTab = Nothing


  readTable = retTable
End Function


Public Sub buttonsCheck()
  Dim myButton As Button
  
  unprotectThisFile
  For Each myButton In ActiveSheet.Buttons
    Select Case (Mid(myButton.Text, 1, 8))
      Case "Ajuda"
        myButton.name = "btnHelp"
      Case "Ler Esto"
        myButton.name = "btnReadStock"
      Case "Abrir Co"
        myButton.name = "btnOpenInventory"
      Case "Fecha 1a"
        myButton.name = "btn1stCount"
      Case "Fecha 2a"
        myButton.name = "btn2ndCount"
      Case "Fecha 3a"
        myButton.name = "btn3rdCount"
      Case "Aprov"
        myButton.name = "btnApprove"
      Case "Nota "
        myButton.name = "btnNotaFiscal"
    End Select
    DoEvents
  Next
  
  Dim bt As Button
  Dim rangetest As Range
  Dim bStatus As Boolean
  
  For Each bt In ActiveSheet.Buttons
    Select Case bt.name
      Case "btnHelp"
        bStatus = True
      Case "btnReadStock"
        bStatus = True
      Case "btnOpenInventory"
        bStatus = isStockReadButtonActive()
      Case "btn1stCount"
        bStatus = isOpenInventoryButtonActive
      Case "btn2ndCount"
        bStatus = isFirstReadButtonActive
      Case "btn3rdCount"
        bStatus = isSecondCountButtonActive
      Case "btnApprove"
        bStatus = isThirdCountButtonActive
      Case "btnNotaFiscal"
        bStatus = isCountingApprovalButtonActive
      Case Else
        bStatus = False
    End Select
    
    bt.Enabled = bStatus
    bt.Visible = bStatus
  Next
  
  protectThisFile
End Sub

Function isStockReadButtonActive() As Boolean
  Dim rg As Range
  
  Set rg = Range(ThisWorkbook.Names("stock_read"))
  If rg.Cells(1, 1) = vbNullString Then
    isStockReadButtonActive = False
  Else
    isStockReadButtonActive = True
  End If

End Function


Function isOpenInventoryButtonActive() As Boolean
  Dim rg As Range
  
  Set rg = Range(ThisWorkbook.Names("open_inventory"))
  If rg.Cells(1, 1) = vbNullString Then
    isOpenInventoryButtonActive = False
  Else
    isOpenInventoryButtonActive = True
  End If

End Function

Function isFirstReadButtonActive() As Boolean
  Dim rg As Range
  
  Set rg = Range(ThisWorkbook.Names("first_count"))
  If rg.Cells(1, 1) = vbNullString Then
    isFirstReadButtonActive = False
  Else
    isFirstReadButtonActive = True
  End If

End Function

Function isSecondCountButtonActive() As Boolean
  Dim rg As Range
  
  Set rg = Range(ThisWorkbook.Names("second_count"))
  If rg.Cells(1, 1) = vbNullString Then
    isSecondCountButtonActive = False
  Else
    isSecondCountButtonActive = True
  End If

End Function


Function isThirdCountButtonActive() As Boolean
  Dim rg As Range
  
  Set rg = Range(ThisWorkbook.Names("third_count"))
  If rg.Cells(1, 1) = vbNullString Then
    isThirdCountButtonActive = False
  Else
    isThirdCountButtonActive = True
  End If

End Function

Function isApprovalButtonActive() As Boolean
  Dim rg As Range
  
  Set rg = Range(ThisWorkbook.Names("count_approver"))
  If rg.Cells(1, 1) = vbNullString Then
    isApprovalButtonActive = False
  Else
    isApprovalButtonActive = True
  End If


End Function

Function isCountingApprovalButtonActive() As Boolean
  Dim rg As Range
  
  Set rg = Range(ThisWorkbook.Names("count_approver"))
  If rg.Cells(1, 1) = vbNullString Then
    isCountingApprovalButtonActive = False
  Else
    isCountingApprovalButtonActive = True
  End If

End Function


Sub setProgressCells(progressName As String)
  Dim current As String
  'current = "=[" & ThisWorkbook.name & "]" & Mid(ThisWorkbook.Names(progressName & "_currentQty"), 2)
  Set cellCurrent = Range("=[" & ThisWorkbook.name & "]" & Mid(ThisWorkbook.Names(progressName & "_currentQty"), 2))
  Set cellTotal = Range("=[" & ThisWorkbook.name & "]" & Mid(ThisWorkbook.Names(progressName & "_totalQty"), 2))
  Set cellPercent = Range("=[" & ThisWorkbook.name & "]" & Mid(ThisWorkbook.Names(progressName & "_percentCompleted"), 2)) 'Range(ThisWorkbook.Names(progressName & "_percentCompleted"))
  Set cellDesc = Range("=[" & ThisWorkbook.name & "]" & Mid(ThisWorkbook.Names(progressName & "_process_desc"), 2)) 'Range(wb.Names(progressName & "_process_desc"))
  
End Sub


Sub setProgressBarTotal(totalQty As Long)
  cellTotal.Value = totalQty
  cellPercent.Value = 0
  cellCurrent.Value = 0
  cellDesc.Value = vbNullString
  OptimizeCode_End
  DoEvents
  OptimizeCode_Begin
End Sub

Sub setProgressBarTitle(title As String)
  ws.Activate
  cellDesc.Value = title
  OptimizeCode_End
  DoEvents
  OptimizeCode_Begin
End Sub

Sub setProgressBarUpByOne()
  cellCurrent.Value = cellCurrent.Value + 1
  cellPercent.Value = cellCurrent.Value / cellTotal.Value
  ws.Cells(1, 1).Select
  If (cellCurrent.Value Mod 100) = 0 Then
    OptimizeCode_End
    DoEvents
    OptimizeCode_Begin
  End If
End Sub



Sub openInventoryDocument()
  Dim sapConn As Object
  
  If Not isOpenInventoryPossible Then
    MsgBox "Não é possível abrir o inventário." & vbCrLf & _
           "Verificar o botão 'Ajuda' para entender o processo", _
           vbCritical, "Erro!! Abertura de inventário não é possível"
    Exit Sub
  End If
  
  Set sapConn = Logon
  
  If sapConn Is Nothing Then
    OptimizeCode_End
    Exit Sub
  End If
  setProgressBarTitle "Login feito com sucesso"
  
  Dim previousSystem As String
  previousSystem = Mid(Range(ThisWorkbook.Names("running_environment")).Value, 1, 3)
  
  If sapConn.Connection.system <> previousSystem Then
    MsgBox "Sistema em que o usuario se logou (" & sapConn.Connection.system & ") difere do ambiente de Leitura de Estoque (" & previousSystem & ")", vbCritical, "ERRO!!!"
    OptimizeCode_End
    Exit Sub
  End If
  
  fillParameters sapConn, K_EVENT_INVENTORY
  
  populateMaterialTable (True) 'All documents
  generatePhysInventoryDocument sapConn
  OptimizeCode_End
  
  MsgBox ("Processamento Finalizado!")
End Sub

Public Function convertNumberToLetter(number As Integer) As String
  If number < 1 Or number > 27 Then
    convertNumberToLetter = vbNullString
  Else
    convertNumberToLetter = Chr(64 + number)
  End If
End Function



Sub OptimizeCode_Begin()
  ActiveSheet.DisplayPageBreaks = False
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  'Application.DisplayStatusBar = False
  Application.Calculation = xlCalculationManual
End Sub

Sub OptimizeCode_End()
  ActiveSheet.DisplayPageBreaks = False
  Application.Calculation = xlCalculationAutomatic
  Application.EnableEvents = True
  Application.ScreenUpdating = True
  'Application.DisplayStatusBar = True
  Application.Calculate
End Sub

Sub OpenCardEditForm()

  Dim wsCard As Worksheet
  Dim tbCard As ListObject
  
  Set ws = ThisWorkbook.Worksheets("Stock")
  Set tbStock = ws.ListObjects("tbStock")
  
  'Se clicar no Ctrl fora da tabela (testa somente no topo) não faz nada
  On Error Resume Next
  If ActiveCell.Row < tbStock.DataBodyRange.Cells(1, 1).Row Then
    On Error GoTo 0
    Exit Sub
  End If
  On Error GoTo 0
  
'Léoincluiu (motivo: Resolver o problema de selecionar a contagem errada)

  If Range("second_count").Value <> "" Then
    Cells(ActiveCell.Row, K_COL_MENGE_3RD).Select
  ElseIf Range("first_count").Value <> "" Then
    Cells(ActiveCell.Row, K_COL_MENGE_2ND).Select
  ElseIf Range("open_inventory").Value <> "" Then
    Cells(ActiveCell.Row, K_COL_MENGE_1ST).Select
  Else
    MsgBox "Contagem não está ativa", vbExclamation, "Não é possível contar"
    Macroloc
    Exit Sub
  End If
  
'verificar por cor
  'If ActiveCell.Interior.Color <> 65535 And ActiveCell.Interior.Color <> 192 Then
  '  MsgBox "Contagem não está ativa", vbExclamation, "Não é possível contar"
  '  Exit Sub
  'End If
  
  
  'verificar por função
  If podCont(ActiveCell.Column) = False Then
    MsgBox "Contagem não está ativa", vbExclamation, "Não é possível contar"
    Macroloc
    Exit Sub
  End If

  'End Léoincluiu
  
 
  frmCard.Caption = "Informe os Cartões e Quantidades para o Material"
  frmCard.Show
  
  ws.Calculate
  Application.Calculate
  
End Sub

Function tryActivingCellForCounting() As Boolean
  
  tryActivingCellForCounting = True
  Set ws = ThisWorkbook.Worksheets("Stock")
  
  If ws.name <> ThisWorkbook.ActiveSheet.name Then
    tryActivingCellForCounting = False
    Exit Function
  End If
  
  
  If isOpenInventoryButtonActive And Not isSecondCountButtonActive And Not isThirdCountButtonActive Then
    ws.Cells(ActiveCell.Row, K_COL_MENGE_1ST).Select
    Exit Function
  End If

  If isSecondCountClosingPossible Then
    ws.Cells(ActiveCell.Row, K_COL_MENGE_2ND).Select
    Exit Function
  End If

  If isThirdCountClosingPossible Then
    ws.Cells(ActiveCell.Row, K_COL_MENGE_3RD).Select
    Exit Function
  End If
  
  tryActivingCellForCounting = False
End Function

Function canAddCardCount() As Boolean
  Dim tbStock As ListObject
  Dim tbRange As Range
  Dim lin As Integer
  
  canAddCardCount = False
  
  Set ws = ThisWorkbook.Sheets("Stock")
  Set tbStock = ws.ListObjects("tbStock")
  
  If ActiveCell.Row < tbStock.DataBodyRange.Cells(1, 1).Row Then
    Exit Function
  End If
  
  If isFirstCountClosingPossible And Not isSecondCountClosingPossible And Not isThirdCountClosingPossible And _
     ActiveCell.Column <> K_COL_MENGE_1ST Then
     Exit Function
  End If
  
  If Not isFirstCountClosingPossible And isSecondCountClosingPossible And _
     ActiveCell.Column <> K_COL_MENGE_2ND Then
     Exit Function
  End If
  
  If Not isSecondCountClosingPossible And isThirdCountClosingPossible And _
     ActiveCell.Column <> K_COL_MENGE_3RD Then
     Exit Function
  End If

  If (isSecondCountClosingPossible Or isThirdCountClosingPossible) And _
     ws.Cells(ActiveCell.Row, K_COL_ABCIN) <> "A" Then
     Exit Function
  End If
  
  If ws.Cells(ActiveCell.Row, K_COL_APROV) = "Sim" Then
     Exit Function
  End If
  
  If (Not isSecondCountClosingPossible And isThirdCountClosingPossible) And _
     ws.Cells(ActiveCell.Row, K_COL_MENGE_1ST) = ws.Cells(ActiveCell.Row, K_COL_MENGE_2ND) Then
     Exit Function
  End If
  
  
  'Passar a permitir acesso a edição de cartões de qualquer coluna
'  ' Se nao for uma das colunas de contagem, sai
'  If Not (ActiveCell.Column = K_COL_MENGE_1ST Or _
'        ActiveCell.Column = K_COL_MENGE_2ND Or _
'        ActiveCell.Column = K_COL_MENGE_3RD) Then
'     Exit Function
'  End If
'
'  ' Se ainda não foi aberta contagem, não permite digitar
'  If ActiveCell.Column = K_COL_MENGE_1ST And _
'    Not isOpenInventoryButtonActive Then
'    MsgBox "Contagem ainda não foi aberta." & vbCrLf & _
'           "Abrir contagem para que sejam criados os documentos de inventário e seja efetuado o bloqueio de estoque", _
'           vbInformation, title:="Contagem ainda não foi aberta"
'    Exit Function
'  End If
'
'
'  ' Se foi pedida primeira contagem, certificar de que nao tem outra contagem aberta
'  If ActiveCell.Column = K_COL_MENGE_1ST And _
'     (isSecondCountButtonActive Or isThirdCountButtonActive) Then
'    MsgBox "Primeira Contagem já foi encerrada." & vbCrLf & _
'           "Não é possível mais efetuar a primeira contagem pois ela já foi encerrada", _
'           vbCritical, title:="Primeira Contagem encerrada"
'     Exit Function
'  End If
'
'
'  ' se foi pedida segunda contagem, certificar de que ela esteja aberta
'  If ActiveCell.Column = K_COL_MENGE_2ND And _
'     Not isFirstReadButtonActive Then
'    MsgBox "Segunda Contagem ainda não foi iniciada." & vbCrLf & _
'           "Não é possível efetuar a segunda contagem pois ela ainda não foi aberta", _
'           vbInformation, title:="Segunda Contagem ainda não iniciada"
'     Exit Function
'  End If
'
'
'  ' se foi pedida segunda contagem, certificar de que nao tem terceira contagem aberta
'  If ActiveCell.Column = K_COL_MENGE_2ND And _
'     isSecondCountButtonActive Then
'    MsgBox "Segunda Contagem já foi encerrada." & vbCrLf & _
'           "Não é possível mais efetuar a segunda contagem pois ela já foi encerrada", _
'           vbCritical, title:="Segunda Contagem encerrada"
'     Exit Function
'  End If
'
'  ' se foi pedida terceira contagem, certificar de que ela esteja aberta
'  If ActiveCell.Column = K_COL_MENGE_3RD And _
'     Not isThirdCountButtonActive Then
'    MsgBox "Terceira Contagem ainda não foi iniciada." & vbCrLf & _
'           "Não é possível efetuar a terceira contagem pois ela ainda não foi aberta", _
'           vbInformation, title:="Terceira Contagem ainda não iniciada"
'     Exit Function
'  End If
'
'  If (ActiveCell.Column = K_COL_MENGE_2ND And _
'      ws.Cells(ActiveCell.Row, K_COL_HAS_1ST) <> "X") Then
'    MsgBox "Segunda contagem só possível para itens que receberam a primeira contagem.", _
'           vbInformation, title:="Segunda contagem não possível para o item"
'     Exit Function
'  End If
'
'  If (ActiveCell.Column = K_COL_MENGE_2ND And _
'      ws.Cells(ActiveCell.Row, K_COL_1ST_CHECK) <> "Falha") Then
'    MsgBox "Segunda contagem só possível para itens que falharam na primeira contagem.", _
'           vbInformation, title:="Segunda contagem não possível para o item"
'     Exit Function
'  End If
'
'  If ActiveCell.Column = K_COL_MENGE_2ND And _
'     ws.Cells(ActiveCell.Row, K_COL_ABCIN) = "A" And _
'     isLineSelectableForCounting(ws, ActiveCell.Row) And _
'     ws.Cells(ActiveCell.Row, K_COL_HAS_1ST) <> "X" Then
'    MsgBox "Segunda contagem - Item não está ativo para segunda contagem", _
'           vbInformation, title:="Segunda contagem não possível para o item"
'     Exit Function
'  End If
'
'
'  If (ActiveCell.Column = K_COL_MENGE_3RD And _
'      ws.Cells(ActiveCell.Row, K_COL_HAS_3RD) <> "X") Then
'    MsgBox "Terceira contagem só possível para itens que receberam a segunda contagem.", _
'           vbInformation, title:="Terceira contagem não possível para o item"
'     Exit Function
'  End If
'
'  If (ActiveCell.Column = K_COL_MENGE_3RD And _
'      ws.Cells(ActiveCell.Row, K_COL_2ND_CHECK) <> "Falha") Then
'    MsgBox "Terceira contagem só possível para itens que falharam na segunda contagem.", _
'           vbInformation, title:="Terceira contagem não possível para o item"
'     Exit Function
'  End If
'
'  If ActiveCell.Column = K_COL_MENGE_3RD And _
'     ws.Cells(ActiveCell.Row, K_COL_ABCIN) = "A" And _
'     isLineSelectableForCounting(ws, ActiveCell.Row) And _
'     ws.Cells(ActiveCell.Row, K_COL_HAS_2ND) = "X" Then
'    MsgBox "Terceira contagem - Item não está ativo para segunda contagem", _
'           vbInformation, title:="Terceira contagem não possível para o item"
'     Exit Function
'  End If
  
  ' Se o material não foi expandido para a planta, acinzenta a linha
  ' Se o material é administrado por lote (mas não é a linha com estoque do lote), acinzenta
  ' Se for saldo em poder de terceiros (fornecedor preenchido), acinzenta
  ' Não contar tipos de material que não geram estoque
  
  If mustGrayLineOut(ws, ActiveCell.Row) Then
    MsgBox "Item bloqueado para contagem." & vbCrLf & _
           "A linha selecionada tem cor 'cinza' para indicar que está bloqueada." & vbCrLf & _
           "Motivos de bloqueio da linha:" & vbCrLf & _
           "- Material não extendido para a Planta " & Plant & vbCrLf & _
           "- Linha em nível de planta para materiais administrados por lote (há outras linhas com os lotes)" & vbCrLf & _
           "- Material em poder de terceiros (impossibilidade legal de ajuste direto de inventário)" & vbCrLf & _
           "- Tipo de Material não permite controle de estoque", _
           vbInformation, title:="Item bloqueado para contagem"
     Exit Function
  End If
  
  canAddCardCount = True
End Function

Function isOpenInventoryPossible() As Boolean
  isOpenInventoryPossible = Not isFirstReadButtonActive
End Function

Function isFirstCountClosingPossible() As Boolean
  isFirstCountClosingPossible = (isStockReadButtonActive And Not isFirstReadButtonActive)
End Function

Function isSecondCountClosingPossible() As Boolean
  isSecondCountClosingPossible = (isFirstReadButtonActive And isSecondCountButtonActive And Not isThirdCountButtonActive)
End Function

Function isThirdCountClosingPossible() As Boolean
  isThirdCountClosingPossible = Not isThirdCountButtonActive
End Function

Function isSheetReadyForSecondCount() As Boolean
  Dim tbStock As ListObject
  Dim tbRange As Range
  
  isSheetReadyForSecondCount = False
  
  Set ws = ThisWorkbook.Sheets("Stock")
  Set tbStock = ws.ListObjects("tbStock")
  Set tbRange = tbStock.DataBodyRange
  
  If tbRange Is Nothing Then
    Exit Function
  End If
  
  setProgressCells (K_PROGRESS_SELE)
  setProgressBarTotal (tbRange.Rows.Count)
  setProgressBarTitle ("Verificando se contagens foram entradas")
  
  For Each Row In tbRange.Rows
    setProgressBarUpByOne
    If isLineSelectableForCounting(ws, Row.Row) And _
       ws.Cells(Row.Row, K_COL_HAS_1ST) <> "X" Then
      If ws.Cells(Row.Row, K_COL_LABST) <> 0 Then
        setProgressBarTitle ("Fechamento da 1a. Contagem - ERRO")
        Exit Function
      End If
    End If
  Next
  
  setProgressBarTitle ("Fechamento da 1a. Contagem - OKAY")
  isSheetReadyForSecondCount = True
End Function

Function isSheetReadyForThirdCount() As Boolean
  Dim tbStock As ListObject
  Dim tbRange As Range
  
  isSheetReadyForThirdCount = False
  
  Set ws = ThisWorkbook.Sheets("Stock")
  Set tbStock = ws.ListObjects("tbStock")
  Set tbRange = tbStock.DataBodyRange
  
  If tbRange Is Nothing Then
    Exit Function
  End If
  
  setProgressCells (K_PROGRESS_SELE)
  setProgressBarTotal (tbRange.Rows.Count)
  setProgressBarTitle ("Verificando se contagens foram entradas")
  
  For Each Row In tbRange.Rows
    setProgressBarUpByOne
    If isLineSelectableForCounting(ws, Row.Row) And _
       ws.Cells(Row.Row, K_COL_HAS_2ND) <> "X" And _
       (ws.Cells(Row.Row, K_COL_APROV) <> "Sim" And _
        ws.Cells(Row.Row, K_COL_APROV) <> "") Then
       If ws.Cells(Row.Row, K_COL_LABST) > 0 Then
          setProgressBarTitle ("Fechamento da 1a. Contagem - ERRO")
          Exit Function
       End If
    End If
  Next
  
  setProgressBarTitle ("Fechamento da 2a. Contagem - OKAY")
  isSheetReadyForThirdCount = True
End Function

Function isSheetReadyForAproval() As Boolean
  Dim tbStock As ListObject
  Dim tbRange As Range
  
  isSheetReadyForAproval = False
  
  
  Set ws = ThisWorkbook.Sheets("Stock")
  Set tbStock = ws.ListObjects("tbStock")
  Set tbRange = tbStock.DataBodyRange
  
  If tbRange Is Nothing Then
    Exit Function
  End If
  
  setProgressCells (K_PROGRESS_SELE)
  setProgressBarTotal (tbRange.Rows.Count)
  setProgressBarTitle ("Verificando se contagens foram entradas")

  
  For Each Row In tbRange.Rows
    setProgressBarUpByOne
    If isLineSelectableForCounting(ws, Row.Row) And _
       ws.Cells(Row.Row, K_COL_HAS_3RD) <> "X" And _
       (ws.Cells(Row.Row, K_COL_APROV) <> "Sim" And _
        ws.Cells(Row.Row, K_COL_APROV) <> "Contagens Iguais") And _
       ws.Cells(Row.Row, K_COL_LABST) > 0 Then
       '(ws.Cells(Row.Row, K_COL_APROV) <> "Sim" And ws.Cells(Row.Row, K_COL_APROV) <> "Contagens Iguais") Then
       setProgressBarTitle ("Fechamento da 1a. Contagem - ERRO")
       Exit Function
    End If
  Next
  
  setProgressBarTitle ("Fechamento da 2a. Contagem - OKAY")
  isSheetReadyForAproval = True
End Function

Sub NotaFiscal()

  Dim sapConn As Object
  If MsgBox("ATENÇÃO!!" & vbCrLf & _
            "Confirma a Emissão de Notas Fiscais sobre o ajuste de inventário?" & vbCrLf & vbCrLf & _
            "Somente serão escriturados os itens que tiveram BAIXA de estoque.", vbYesNo, "") = vbNo Then
    Exit Sub
  End If
  
  Set sapConn = Logon
        
  If sapConn Is Nothing Then
    Exit Sub
  End If
  
  Dim previousSystem As String
  previousSystem = Mid(Range(ThisWorkbook.Names("running_environment")).Value, 1, 3)
  
  If sapConn.Connection.system <> previousSystem Then
    MsgBox "Sistema em que o usuario se logou (" & sapConn.Connection.system & ") difere do ambiente de Leitura de Estoque (" & previousSystem & ")", vbCritical, "ERRO!!!"
    Exit Sub
  End If
  
  fillParameters sapConn, K_EVENT_NOTAFISCAL
  
  ReDim tbMatnr(0) 'Limpa a tabela de materiais
  populateMaterialTable (True)
  
  
  Set ws = ThisWorkbook.Worksheets("Stock")
  Set tbStock = ws.ListObjects("tbStock")
  ReDim tbPhysInvDoc(tbStock.DataBodyRange.Rows.Count)
  Dim oldDoc As String
  Dim pos As Integer
  oldDoc = ""
  pos = 0
  'Obtem a lista de documentos de Inventário
  For Each Row In tbStock.DataBodyRange.Rows
    If ws.Cells(Row.Row, K_COL_MTDHD) <> "" And _
       ws.Cells(Row.Row, K_COL_MTDHD) <> oldDoc Then
       oldDoc = ws.Cells(Row.Row, K_COL_MTDHD)
       pos = pos + 1
       tbPhysInvDoc(pos).physInv_doc = ws.Cells(Row.Row, K_COL_IVDHD)
       tbPhysInvDoc(pos).goods_doc = ws.Cells(Row.Row, K_COL_MTDHD)
    End If
  Next
  ReDim Preserve tbPhysInvDoc(pos)
  
  For lin = 1 To UBound(tbPhysInvDoc)
    IssueNotaFiscal sapConn, tbPhysInvDoc(lin).physInv_doc, tbPhysInvDoc(lin).goods_doc
  Next
  
  setupEditableRowsAndColumns (K_COL_DOCNM)
  
  MsgBox "Notas Fiscais com Ajustes de Inventário Lançadas"

End Sub
