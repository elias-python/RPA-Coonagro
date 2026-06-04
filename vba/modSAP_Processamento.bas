Option Explicit

' ==============================
' Controle global
' ==============================
Public ContadorErros As Long
Public ContadorProcessos As Long
Public Executar As Boolean
Public Pausado As Boolean
Public PararMacro As Boolean

' ==============================
' Configuracao da planilha
' ==============================
Public Const PLAN_SHEET As String = "Base"
Public Const NOME_DADOS As String = "Base_Dados_Coonagro.xlsm"

' Campos usados da planilha (ajuste apenas aqui)
' Layout gerado pelo XML Monitoring.py:
'   A=Nº Nota Fiscal  B=Seq.  C=Cod.Material  D=Descricao  E=Ordem Producao
'   F=Qtd.  G=UN.  H=Vlr.Unit.  I=Vlr.Total Item  J=CFOP
'   K=NF Vinculada(5124)  L=NF Origem  M=Serie(Origem)  N=Vlr.Total NF
'   O=Emissao  P=Chave de Acesso
Public Const COL_LINHA_VALIDA As String = "A"     ' Nº Nota Fiscal (linha valida se preenchida)
Public Const COL_CFOP As String = "J"              ' CFOP (5902=materiais / 5124=linha mae)
Public Const COL_NF_ORIGEM As String = "L"        ' NF origem do material (NFR)
Public Const COL_SERIE_ORIGEM As String = "M"     ' Serie origem (NFR)
Public Const COL_QTD As String = "F"              ' Quantidade (NFR)
Public Const COL_EMISSAO As String = "O"          ' Data de emissao (DT_LANC)
' NF Alta = Col A das linhas 5902 (StartRow)
' NF Baixa = Col A da linha 5124 (EndRow) - mesma coluna, linha diferente
Public Const COL_NF_HIGH As String = "A"          ' Nº Nota Fiscal (NF Alta e NF Baixa, linhas distintas)
Public Const COL_XPROD As String = "D"            ' Descricao do material (regra de pedido, lida da linha 5124)
Public Const COL_OPERACAO As String = "E"         ' Ordem de producao (apoio)

Private Type GrupoRange
    StartRow As Long
    EndRow As Long
    NfLow As String
    NfHigh As String
    XProd As String
    Operacao As String
    Emissao As String
    StatusAtual As String
    PedidoSAP As String
End Type

' Cores de status na coluna A (conforme regra operacional)
' AUTCARR_OK = AZUL | ERRO = VERMELHO | CONCLUIDO = VERDE | SEM PAR = LARANJA

Private Function LinhaPendentePorCor(ByVal cel As Range) As Boolean
    ' Pendente = sem preenchimento na coluna A.
    LinhaPendentePorCor = (cel.Interior.Pattern = xlPatternNone Or cel.Interior.ColorIndex = xlColorIndexNone)
End Function

Private Function StatusPorCorLinha(ByVal cel As Range) As String
    If LinhaPendentePorCor(cel) Then
        StatusPorCorLinha = ""
        Exit Function
    End If

    Select Case cel.Interior.Color
        Case RGB(66, 165, 245)
            StatusPorCorLinha = "AUTCARR_OK"
        Case RGB(239, 83, 80)
            StatusPorCorLinha = "ERRO"
        Case RGB(102, 187, 106)
            StatusPorCorLinha = "CONCLUIDO"
        Case RGB(255, 152, 0)
            StatusPorCorLinha = "SEM_PAR"
        Case Else
            StatusPorCorLinha = "OUTRO"
    End Select
End Function

Public Sub IniciarComControle()
    ResetEstadoExecucao
    AtualizarDashboard       ' nao recria shapes — apenas atualiza textos
    On Error Resume Next
    With ThisWorkbook.Worksheets(PLAN_SHEET)
        .Activate
        ActiveWindow.ScrollRow = 1   ' sempre rola para o topo ao iniciar
        ActiveWindow.ScrollColumn = 1
    End With
    On Error GoTo 0
    AtualizarStatus "Iniciando..."
    IniciarProcessamentoSAP
End Sub

Public Sub ResetEstadoExecucao()
    ContadorErros = 0
    ContadorProcessos = 0
    Executar = True
    Pausado = False
    PararMacro = False
End Sub

Public Sub IniciarProcessamentoSAP()
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim abriuAgora As Boolean
    Dim lastRow As Long
    Dim cursor As Long
    Dim grp As GrupoRange
    Dim caminhoArquivo As String
    Dim resultado As String

    On Error GoTo TratarErroFatal

    ' Abre o arquivo de dados externo (gerado pelo XML Monitoring.py)
    caminhoArquivo = ThisWorkbook.Path & "\" & NOME_DADOS
    abriuAgora = False

    If StrComp(ThisWorkbook.Name, NOME_DADOS, vbTextCompare) = 0 Then
        Set wb = ThisWorkbook
    Else
        On Error Resume Next
        Set wb = Workbooks(NOME_DADOS)
        On Error GoTo TratarErroFatal
    End If

    If wb Is Nothing Then
        Set wb = Workbooks.Open(caminhoArquivo, ReadOnly:=False)
        abriuAgora = Not wb Is Nothing
    End If
    If wb Is Nothing Then
        MsgBox "Nao foi possivel abrir: " & caminhoArquivo, vbCritical
        Exit Sub
    End If
    On Error Resume Next
    Set ws = wb.Worksheets(PLAN_SHEET)
    On Error GoTo TratarErroFatal
    If ws Is Nothing Then
        Set ws = wb.Worksheets(1)   ' fallback: primeira aba
    End If
    If ws Is Nothing Then
        MsgBox "Nao foi possivel encontrar a aba '" & PLAN_SHEET & "' em:" & vbCrLf & caminhoArquivo, vbCritical
        If abriuAgora Then wb.Close False
        Exit Sub
    End If

    ' Verifica se o SAP esta aberto antes de iniciar
    If Not SAPEstaAberto() Then
        MsgBox "O SAP GUI nao esta aberto." & vbCrLf & _
               "Abra o SAP e faca login antes de iniciar a automacao.", _
               vbExclamation, "SAP nao encontrado"
        If abriuAgora Then wb.Close False
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, COL_LINHA_VALIDA).End(xlUp).Row
    cursor = 3  ' Dados começam na linha 3 (linha 1=título, linha 2=cabeçalho)

    AtualizarPainelControle "Rodando"

    Do While Executar And cursor <= lastRow
        DoEvents

        If PararMacro Then Exit Do

        Do While Pausado
            DoEvents
            AtualizarPainelControle "Pausado"
            Application.Wait Now + TimeValue("0:00:01")
            If PararMacro Then Exit Do
        Loop

        If PararMacro Then Exit Do

        If Not ProximoGrupo(ws, cursor, lastRow, grp) Then Exit Do

        If grp.StartRow = 0 Then
            cursor = cursor + 1
            GoTo ContinueLoop
        End If

        ' Passagem 2 (AUTCARR_OK) nao precisa de PedidoSAP — valida apenas para Passagem 1
        If Len(grp.PedidoSAP) = 0 And UCase$(grp.StatusAtual) <> "AUTCARR_OK" Then
            AtualizarStatusFaixa ws, grp.StartRow, grp.EndRow, "Pedido SAP nao mapeado"
            ContadorErros = ContadorErros + 1
            AtualizarPainelControle "Erro de mapeamento"
            cursor = grp.EndRow + 1
            GoTo ContinueLoop
        End If

        AtualizarPainelControle "Processando " & grp.NfLow & ":" & grp.NfHigh
        resultado = ProcessarGrupoSAP(ws, grp)
        wb.Save   ' salva o status escrito no arquivo externo

        If resultado = "RETRY" Then
            AtualizarPainelControle "Aguardando impressao de remessa"
            ' cursor nao avanca: mesmo grupo sera tentado novamente
        Else
            cursor = grp.EndRow + 1
        End If
        If resultado = "PARAR" Then PararMacro = True

ContinueLoop:
        AtualizarContadoresUI
        DoEvents
    Loop

    If PararMacro Then
        AtualizarPainelControle "Parado"
        MsgBox "Processo interrompido pelo usuario.", vbInformation
    Else
        AtualizarPainelControle "Finalizado"
        MsgBox "Processo finalizado.", vbInformation
    End If

    If Not wb Is Nothing Then
        wb.Save
        ' Arquivo permanece aberto para consulta apos execucao
    End If
    Exit Sub

TratarErroFatal:
    Dim msgErro As String
    Dim numErro As Long
    msgErro = Err.Description   ' salva ANTES do cleanup (On Error GoTo 0 zera o Err)
    numErro = Err.Number
    If msgErro = "" Then msgErro = "Erro desconhecido (sem descricao)"
    If Not wb Is Nothing Then
        On Error Resume Next
        wb.Save
        If abriuAgora Then wb.Close False
        On Error GoTo 0
    End If
    AtualizarPainelControle "Erro fatal"
    MsgBox "Erro fatal (#" & numErro & "): " & msgErro, vbCritical
End Sub

Private Function ProximoGrupo(ByVal ws As Worksheet, ByVal startRow As Long, ByVal lastRow As Long, ByRef grp As GrupoRange) As Boolean
    Dim r As Long
    Dim linhaValida As String
    Dim cfopNext As String

    grp.StartRow = 0
    grp.EndRow = 0
    ProximoGrupo = False

    r = startRow

    ' Processa linhas sem cor (pendentes), azuis (AUTCARR_OK — segunda passagem pendente)
    ' ou laranja (SEM_PAR/NAO_ENCONTRADA — reprocessar apos cockpit ou impressao).
    ' Ignora VERDE, VERMELHO e outras cores finais.
    Dim stLinha As String
    Do While r <= lastRow
        linhaValida = Trim$(CStr(ws.Cells(r, COL_LINHA_VALIDA).Value))
        If Len(linhaValida) > 0 Then
            stLinha = StatusPorCorLinha(ws.Cells(r, COL_LINHA_VALIDA))
            If stLinha = "" Or stLinha = "AUTCARR_OK" Or stLinha = "SEM_PAR" Then Exit Do
        End If
        r = r + 1
    Loop

    If r > lastRow Then
        ProximoGrupo = False
        Exit Function
    End If

    grp.StartRow = r
    grp.NfHigh = Trim$(CStr(ws.Cells(r, COL_NF_HIGH).Value))  ' Col A da 1a linha (5902)
    grp.Emissao = CStr(ws.Cells(r, COL_EMISSAO).Value)
    grp.StatusAtual = StatusPorCorLinha(ws.Cells(r, COL_LINHA_VALIDA))

    ' Estende EndRow enquanto Col A = NfHigh (todas as linhas 5902 do grupo)
    grp.EndRow = r
    Do While grp.EndRow + 1 <= lastRow
        If Trim$(CStr(ws.Cells(grp.EndRow + 1, COL_NF_HIGH).Value)) = grp.NfHigh Then
            grp.EndRow = grp.EndRow + 1
        Else
            Exit Do
        End If
    Loop

    ' Inclui a linha 5124 (ultima do grupo): Col A diferente, CFOP = "5124"
    If grp.EndRow + 1 <= lastRow Then
        cfopNext = Trim$(CStr(ws.Cells(grp.EndRow + 1, COL_CFOP).Value))
        Dim nfNext As String
        nfNext = Trim$(CStr(ws.Cells(grp.EndRow + 1, COL_NF_HIGH).Value))
        If cfopNext = "5124" And Len(nfNext) > 0 Then
            grp.EndRow = grp.EndRow + 1
        End If
    End If

    ' NF Baixa = Col A da ultima linha do grupo (linha 5124)
    grp.NfLow = Trim$(CStr(ws.Cells(grp.EndRow, COL_NF_HIGH).Value))

    ' XProd e Operacao lidos da linha 5124 (onde esta a descricao da operacao)
    grp.XProd = NormalizeText(CStr(ws.Cells(grp.EndRow, COL_XPROD).Value))
    grp.Operacao = NormalizeText(CStr(ws.Cells(grp.EndRow, COL_OPERACAO).Value))
    grp.PedidoSAP = PedidoPorXProd(grp.XProd)

    ProximoGrupo = True
End Function

Private Function ProcessarGrupoSAP(ByVal ws As Worksheet, ByRef grp As GrupoRange) As String
    ProcessarGrupoSAP = "OK"   ' valor de retorno padrao
    Dim sapGuiAuto As Object
    Dim SAPApp As Object
    Dim SAPCon As Object
    Dim session As Object
    Dim sapGrid As Object
    Dim i As Long
    Dim rowCount As Long
    Dim tentativa As Long
    Dim itensParaProcessar As Long
    Dim linhaPlanilha As Long
    Dim resposta As VbMsgBoxResult
    Dim msgRemessa As String
    msgRemessa = "Va ate a transacao J1B3N para efetuar a impressao da Nota de Remessa." & vbCrLf & vbCrLf & _
                 "Apos imprimir, clique em 'OK' (Feito) para o robo tentar este grupo novamente." & vbCrLf & _
                 "Clique em 'Cancelar' para PARAR a macro."

    On Error GoTo TratarErroGrupo

    ' Reconexao forcada por grupo
    On Error Resume Next
    Set sapGuiAuto = GetObject("SAPGUI")
    On Error GoTo TratarErroGrupo
    If sapGuiAuto Is Nothing Then
        MsgBox "SAP GUI nao esta aberto. Abra o SAP e faca login antes de iniciar a automacao.", vbExclamation, "SAP nao encontrado"
        PararMacro = True
        ProcessarGrupoSAP = "PARAR"
        Exit Function
    End If

    Set SAPApp = sapGuiAuto.GetScriptingEngine

    ' Verifica se ha conexoes abertas (Logon Pad aberto sem login nao tem Children)
    If SAPApp.Children.Count = 0 Then
        MsgBox "SAP GUI esta aberto mas nenhuma conexao foi encontrada." & vbCrLf & _
               "Faca login no sistema PA1 e tente novamente.", vbExclamation, "Sem conexao SAP"
        PararMacro = True
        ProcessarGrupoSAP = "PARAR"
        Exit Function
    End If

    Set SAPCon = SAPApp.Children(0)

    ' Verifica se a conexao tem sessoes (pode estar apenas na tela de logon do sistema)
    If SAPCon.Children.Count = 0 Then
        MsgBox "Conexao SAP encontrada mas sem sessao ativa." & vbCrLf & _
               "Faca login no sistema PA1 e tente novamente.", vbExclamation, "Sessao SAP invalida"
        PararMacro = True
        ProcessarGrupoSAP = "PARAR"
        Exit Function
    End If

    Set session = SAPCon.Children(0)

    ' Confirma que a sessao esta logada verificando o tipo da janela principal
    On Error Resume Next
    Dim tipoJanela As String
    tipoJanela = session.findById("wnd[0]").Type
    If Err.Number <> 0 Or Len(tipoJanela) = 0 Then
        MsgBox "Sessao SAP encontrada mas nao esta respondendo." & vbCrLf & _
               "Verifique se o login no PA1 foi concluido.", vbExclamation, "Sessao SAP invalida"
        PararMacro = True
        ProcessarGrupoSAP = "PARAR"
        Exit Function
    End If
    On Error GoTo TratarErroGrupo

    With session
        Application.Wait Now + TimeValue("0:00:02")  ' aguarda breve delay antes de navegar
        .findById("wnd[0]").maximize
        .findById("wnd[0]/tbar[0]/okcd").Text = "/nzt_mm_94n"
        .findById("wnd[0]").sendVKey 0

        ' Carrega a variante de selecao pelo nome (gravado do SAP)
        .findById("wnd[0]/tbar[1]/btn[17]").press
        .findById("wnd[1]/usr/txtENAME-LOW").Text = "NMELO4"
        .findById("wnd[1]/usr/txtENAME-LOW").setFocus
        .findById("wnd[1]/tbar[0]/btn[8]").press
        .findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellColumn = "TEXT"
        .findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
        .findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell

        ' Abre o popup de selecao de NFs
        .findById("wnd[0]/usr/btn%_S_NFNUM_%_APP_%-VALU_PUSH").press

        .findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,0]").Text = grp.NfLow
        .findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,1]").Text = grp.NfHigh
        .findById("wnd[0]/tbar[1]/btn[8]").press
        .findById("wnd[0]/tbar[1]/btn[8]").press

        If .Children.Count > 1 Then
            If InStr(1, CStr(.Children(1).Text), "Informacao", vbTextCompare) > 0 _
                Or InStr(1, CStr(.Children(1).Text), "Informação", vbTextCompare) > 0 Then
                .Children(1).sendVKey 0
                ' Nao encontrada: pode ser que o Cockpit ainda nao foi executado para este grupo.
                ' Marca LARANJA (atencao manual) em vez de VERDE.
                AtualizarStatusFaixa ws, grp.StartRow, grp.EndRow, "NAO ENCONTRADA NO SAP"
                ContadorErros = ContadorErros + 1
                GoTo FinalizaGrupo
            End If
        End If

        Set sapGrid = .findById("wnd[0]/usr/shell/shellcont[1]/shell")

        sapGrid.setCurrentCell -1, "CFOP"
        sapGrid.selectColumn "CFOP"
        sapGrid.pressToolbarButton "&SORT_ASC"

        If Trim$(CStr(sapGrid.GetCellValue(0, "CFOP"))) <> "1124AA" Then
            AtualizarStatusFaixa ws, grp.StartRow, grp.EndRow, "CFOP 1124AA nao encontrado"
            ContadorErros = ContadorErros + 1
            GoTo FinalizaGrupo
        End If

        rowCount = CLng(sapGrid.rowCount)

        ' Passagem 1
        If UCase$(grp.StatusAtual) <> "AUTCARR_OK" Then
            For i = 0 To rowCount - 1
                If Trim$(CStr(sapGrid.GetCellValue(i, "BSTNR"))) <> "" Then sapGrid.modifyCell i, "BSTNR", ""
                If Trim$(CStr(sapGrid.GetCellValue(i, "EBELP"))) <> "" Then sapGrid.modifyCell i, "EBELP", ""
            Next i

            sapGrid.modifyCell 0, "BSTNR", grp.PedidoSAP
            sapGrid.modifyCell 0, "EBELP", "10"

            For i = 0 To rowCount - 1
                sapGrid.modifyCell i, "PESO_BAL", sapGrid.GetCellValue(i, "MENGE")
            Next i

            sapGrid.modifyCell 0, "DT_LANC", ToDDMMYYYY(grp.Emissao)
            sapGrid.triggerModified
            sapGrid.currentCellColumn = ""
            sapGrid.selectedRows = "0"

            If Not RunAUTCARR(session, sapGrid, 3) Then
                AtualizarStatusFaixa ws, grp.StartRow, grp.EndRow, "Erro AUTCARR"
                ContadorErros = ContadorErros + 1
                GoTo FinalizaGrupo
            End If

            ' AUTCARR_OK concluido — re-navega imediatamente para executar NFR no mesmo grupo
            AtualizarStatusFaixa ws, grp.StartRow, grp.EndRow, "AUTCARR_OK"
            grp.StatusAtual = "AUTCARR_OK"

            Application.Wait Now + TimeValue("0:00:03")
            .findById("wnd[0]").maximize
            .findById("wnd[0]/tbar[0]/okcd").Text = "/nzt_mm_94n"
            .findById("wnd[0]").sendVKey 0

            .findById("wnd[0]/tbar[1]/btn[17]").press
            .findById("wnd[1]/usr/txtENAME-LOW").Text = "NMELO4"
            .findById("wnd[1]/usr/txtENAME-LOW").setFocus
            .findById("wnd[1]/tbar[0]/btn[8]").press
            .findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellColumn = "TEXT"
            .findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
            .findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell

            .findById("wnd[0]/usr/btn%_S_NFNUM_%_APP_%-VALU_PUSH").press
            .findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,0]").Text = grp.NfLow
            .findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,1]").Text = grp.NfHigh
            .findById("wnd[0]/tbar[1]/btn[8]").press
            .findById("wnd[0]/tbar[1]/btn[8]").press

            ' Verifica popup de NF nao encontrada apos re-navegacao
            If .Children.Count > 1 Then
                If InStr(1, CStr(.Children(1).Text), "Informacao", vbTextCompare) > 0 _
                    Or InStr(1, CStr(.Children(1).Text), "Informação", vbTextCompare) > 0 Then
                    .Children(1).sendVKey 0
                    ' AUTCARR_OK ja gravado — NFR nao realizado pois NF sumiu apos re-navegacao
                    ' Mantém AZUL (nao sobrescreve) para reprocessar depois
                    GoTo FinalizaGrupo
                End If
            End If

            ' Atualiza grid e rowCount para a Passagem 2
            Set sapGrid = .findById("wnd[0]/usr/shell/shellcont[1]/shell")
            sapGrid.setCurrentCell -1, "CFOP"
            sapGrid.selectColumn "CFOP"
            sapGrid.pressToolbarButton "&SORT_ASC"
            rowCount = CLng(sapGrid.rowCount)
        End If

        ' Passagem 2
        sapGrid.SelectAll
        sapGrid.pressToolbarButton "NFR"

        tentativa = 0
        Do While .findById("wnd[1]/usr/tblZMM_RE033_IND_PROCESSTC_ZTMM_INSUMOS", False) Is Nothing And tentativa < 15
            Application.Wait Now + TimeValue("0:00:01")
            tentativa = tentativa + 1
        Loop

        If tentativa >= 15 Then
            AtualizarStatusFaixa ws, grp.StartRow, grp.EndRow, "Erro: janela NFR nao abriu"
            ContadorErros = ContadorErros + 1
            GoTo FinalizaGrupo
        End If

        ' Exclui a linha 5124 (mae) da contagem — apenas linhas 5902 vao para o NFR
        itensParaProcessar = WorksheetFunction.Min(grp.EndRow - grp.StartRow, rowCount)

        ' 1. Preenche todas as NFs origens
        For i = 0 To itensParaProcessar - 1
            .findById("wnd[1]/usr/tblZMM_RE033_IND_PROCESSTC_ZTMM_INSUMOS/txtTI_ZTMM_INSUMOS-Z_NFINS[2," & i & "]").Text = CStr(ws.Cells(grp.StartRow + i, COL_NF_ORIGEM).Value)
        Next i

        ' 2. Preenche todas as series
        For i = 0 To itensParaProcessar - 1
            .findById("wnd[1]/usr/tblZMM_RE033_IND_PROCESSTC_ZTMM_INSUMOS/txtTI_ZTMM_INSUMOS-Z_SERIES[3," & i & "]").Text = CStr(ws.Cells(grp.StartRow + i, COL_SERIE_ORIGEM).Value)
        Next i

        ' 3. Enter para validar e carregar os dados da NF origem
        .findById("wnd[1]").sendVKey 0

        On Error Resume Next
        If Not .findById("wnd[2]", False) Is Nothing Then
            If InStr(1, CStr(.findById("wnd[2]").Text), "Nota Fiscal de Remessa", vbTextCompare) > 0 Then
                ' Imprime automaticamente via J1BNFE; RETRY para reprocessar o grupo
                If ImprimirNFRemessa(session) Then
                    ProcessarGrupoSAP = "RETRY"
                Else
                    PararMacro = True
                    ProcessarGrupoSAP = "PARAR"
                End If
                GoTo FinalizaGrupo
            End If
        End If
        On Error GoTo TratarErroGrupo

        ' 4. Preenche todas as quantidades (formato SAP brasileiro: virgula decimal)
        For i = 0 To itensParaProcessar - 1
            .findById("wnd[1]/usr/tblZMM_RE033_IND_PROCESSTC_ZTMM_INSUMOS/txtTI_ZTMM_INSUMOS-Z_QDADE[5," & i & "]").Text = FormatQtdSAP(ws.Cells(grp.StartRow + i, COL_QTD).Value)
        Next i

        ' 5. Foco na primeira quantidade e Enter para confirmar
        .findById("wnd[1]/usr/tblZMM_RE033_IND_PROCESSTC_ZTMM_INSUMOS/txtTI_ZTMM_INSUMOS-Z_QDADE[5,0]").setFocus
        Application.Wait Now + TimeValue("0:00:01")
        .findById("wnd[1]").sendVKey 0

        On Error Resume Next
        If Not .findById("wnd[2]", False) Is Nothing Then
            If InStr(1, CStr(.findById("wnd[2]").Text), "Nota Fiscal de Remessa", vbTextCompare) > 0 Then
                ' Imprime automaticamente via J1BNFE; RETRY para reprocessar o grupo
                If ImprimirNFRemessa(session) Then
                    ProcessarGrupoSAP = "RETRY"
                Else
                    PararMacro = True
                    ProcessarGrupoSAP = "PARAR"
                End If
                GoTo FinalizaGrupo
            End If
        End If
        On Error GoTo TratarErroGrupo

        .findById("wnd[1]/tbar[0]/btn[8]").press
        .findById("wnd[2]/usr/btnSPOP-OPTION1").press

        sapGrid.SelectAll
        sapGrid.pressToolbarButton "EINFE2"
        .findById("wnd[1]/usr/btnSPOP-OPTION1").press
        .findById("wnd[1]/usr/btnSPOP-OPTION1").press
        .findById("wnd[1]/usr/btnSPOP-OPTION1").press

        AtualizarStatusFaixa ws, grp.StartRow, grp.EndRow, "OK"
        ContadorProcessos = ContadorProcessos + 1
    End With

FinalizaGrupo:
    Set sapGrid = Nothing
    Set session = Nothing
    Set SAPCon = Nothing
    Set SAPApp = Nothing
    Set sapGuiAuto = Nothing
    Exit Function

TratarErroGrupo:
    ProcessarGrupoSAP = "ERRO"
    AtualizarStatusFaixa ws, grp.StartRow, grp.EndRow, "Verificar erro: " & Err.Description
    ContadorErros = ContadorErros + 1
    Resume FinalizaGrupo
End Function

Private Function RunAUTCARR(ByVal session As Object, ByVal sapGrid As Object, ByVal maxTentativas As Integer) As Boolean
    Dim tentativa As Integer

    RunAUTCARR = False

    For tentativa = 1 To maxTentativas
        On Error Resume Next
        session.findById("wnd[0]").sendVKey 0
        Application.Wait Now + TimeValue("0:00:01")
        sapGrid.pressToolbarButton "AUTCARR"
        Application.Wait Now + TimeValue("0:00:02")

        Err.Clear
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        If Err.Number = 0 Then
            RunAUTCARR = True
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0
    Next tentativa
End Function

Private Function PopupRemessaAberto(ByVal session As Object) As Boolean
    On Error Resume Next
    If session.findById("wnd[2]", False) Is Nothing Then
        PopupRemessaAberto = False
    Else
        PopupRemessaAberto = (InStr(1, CStr(session.findById("wnd[2]").Text), "Nota Fiscal de Remessa", vbTextCompare) > 0)
    End If
    On Error GoTo 0
End Function

Private Sub AtualizarStatusFaixa(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal statusText As String)
    Dim i As Long
    Dim corFundo As Long
    Dim corFonte As Long
    Dim usarCor As Boolean
    Dim su As String
    usarCor = True
    su = UCase$(Trim$(statusText))

    Select Case su
        Case "AUTCARR_OK"
            corFundo = RGB(66, 165, 245)    ' azul
            corFonte = RGB(255, 255, 255)
        Case "OK", "CONCLUIDO", "CONCLUÍDO"
            corFundo = RGB(102, 187, 106)   ' verde
            corFonte = RGB(255, 255, 255)
        Case Else
            If InStr(1, su, "NAO MAPEADO") > 0 Or InStr(1, su, "SEM PAR") > 0 _
                Or InStr(1, su, "NAO ENCONTRADO") > 0 Or InStr(1, su, "NAO ENCONTRADA") > 0 Then
                corFundo = RGB(255, 152, 0)  ' laranja
                corFonte = RGB(255, 255, 255)
            ElseIf InStr(1, su, "ERRO") > 0 Then
                corFundo = RGB(239, 83, 80)  ' vermelho
                corFonte = RGB(255, 255, 255)
            Else
                usarCor = False              ' neutro: sem preenchimento
            End If
    End Select

    For i = startRow To endRow
        With ws.Cells(i, COL_LINHA_VALIDA)
            If usarCor Then
                .Interior.Color = corFundo
                .Font.Color     = corFonte
            Else
                .Interior.ColorIndex = xlColorIndexNone
                .Font.ColorIndex     = xlColorIndexAutomatic
            End If
        End With
    Next i
End Sub

Public Sub PausarProcessamento()
    Pausado = True
    AtualizarPainelControle "Pausado"
End Sub

Public Sub ContinuarProcessamento()
    Pausado = False
    AtualizarPainelControle "Rodando"
End Sub

Public Sub PararProcessamento()
    PararMacro = True
    Executar = False
    AtualizarPainelControle "Parando..."
End Sub

Public Sub AtualizarPainelControle(ByVal statusTxt As String)
    AtualizarStatus statusTxt
    AtualizarContadoresUI
End Sub

Public Sub AtualizarContadoresUI()
    AtualizarContadores ContadorErros, ContadorProcessos
End Sub

Private Function NormalizeText(ByVal value As String) As String
    Dim txt As String
    txt = UCase$(Trim$(CStr(value)))

    txt = Replace(txt, "Á", "A")
    txt = Replace(txt, "À", "A")
    txt = Replace(txt, "Â", "A")
    txt = Replace(txt, "Ã", "A")
    txt = Replace(txt, "Ä", "A")

    txt = Replace(txt, "É", "E")
    txt = Replace(txt, "È", "E")
    txt = Replace(txt, "Ê", "E")
    txt = Replace(txt, "Ë", "E")

    txt = Replace(txt, "Í", "I")
    txt = Replace(txt, "Ì", "I")
    txt = Replace(txt, "Î", "I")
    txt = Replace(txt, "Ï", "I")

    txt = Replace(txt, "Ó", "O")
    txt = Replace(txt, "Ò", "O")
    txt = Replace(txt, "Ô", "O")
    txt = Replace(txt, "Õ", "O")
    txt = Replace(txt, "Ö", "O")

    txt = Replace(txt, "Ú", "U")
    txt = Replace(txt, "Ù", "U")
    txt = Replace(txt, "Û", "U")
    txt = Replace(txt, "Ü", "U")

    txt = Replace(txt, "Ç", "C")

    NormalizeText = txt
End Function

Private Function PedidoPorXProd(ByVal xprodNormalizado As String) As String
    ' Pedidos NOVOS (obrigatorio usar estes)
    If InStr(1, xprodNormalizado, "INDUSTRIALIZACAO MISTURA SACARIA", vbTextCompare) > 0 Then
        PedidoPorXProd = "4510000397"
        Exit Function
    End If

    If InStr(1, xprodNormalizado, "INDUSTRIALIZACAO MISTURA FOSFATADO", vbTextCompare) > 0 Then
        PedidoPorXProd = "4510000395"
        Exit Function
    End If

    If InStr(1, xprodNormalizado, "INDUSTRIALIZACAO MISTURA NITROGENADO", vbTextCompare) > 0 Then
        PedidoPorXProd = "4510000396"
        Exit Function
    End If

    ' Sem fallback para pedidos antigos.
    PedidoPorXProd = ""
End Function

Private Function SAPEstaAberto() As Boolean
    Dim obj As Object
    On Error Resume Next
    Set obj = GetObject("SAPGUI")
    On Error GoTo 0
    SAPEstaAberto = Not (obj Is Nothing)
End Function

Private Function FormatQtdSAP(ByVal valor As Variant) As String
    ' Converte numero para formato SAP brasileiro (virgula como separador decimal)
    ' Ex: 26.968 -> "26,968"  |  720 -> "720,000"
    Dim d As Double
    On Error Resume Next
    d = CDbl(valor)
    If Err.Number <> 0 Then
        Err.Clear
        FormatQtdSAP = CStr(valor)
        Exit Function
    End If
    On Error GoTo 0
    FormatQtdSAP = Replace(Format$(d, "0.000"), ".", ",")
End Function

Private Function ToDDMMYYYY(ByVal value As String) As String
    Dim dt As Date
    On Error Resume Next
    dt = CDate(value)
    If Err.Number = 0 Then
        ToDDMMYYYY = Format$(dt, "ddmmyyyy")
    Else
        Err.Clear
        ToDDMMYYYY = Format$(Date, "ddmmyyyy")
    End If
    On Error GoTo 0
End Function

' ==============================
' Impressao automatica via J1BNFE
' ==============================

' Tenta extrair NF e Serie a partir do texto do popup wnd[2].
' Retorna True se conseguiu; nfNum sem zeros a esquerda, nfSerie com 3 digitos.
Private Function ExtrairNFDoPopup(ByVal session As Object, ByRef nfNum As String, ByRef nfSerie As String) As Boolean
    Dim textoErro As String
    Dim pos1 As Long, pos2 As Long

    textoErro = ""
    nfNum = ""
    nfSerie = ""

    On Error Resume Next
    ' Tentativa 1: campo de texto padrao de mensagem SAP
    textoErro = CStr(session.findById("wnd[2]/usr/txtMESSAGE").Text)
    ' Tentativa 2: propriedade MessageText (GuiMessageWindow)
    If Err.Number <> 0 Or Len(Trim(textoErro)) = 0 Then
        Err.Clear
        textoErro = CStr(session.findById("wnd[2]").MessageText)
    End If
    ' Tentativa 3: primeiro filho de wnd[2]/usr
    If Err.Number <> 0 Or Len(Trim(textoErro)) = 0 Then
        Err.Clear
        textoErro = CStr(session.findById("wnd[2]/usr").Children(0).Text)
    End If
    ' Tentativa 4: segundo filho
    If Err.Number <> 0 Or Len(Trim(textoErro)) = 0 Then
        Err.Clear
        textoErro = CStr(session.findById("wnd[2]/usr").Children(1).Text)
    End If
    On Error GoTo 0

    If Len(Trim(textoErro)) = 0 Or InStr(1, textoErro, "NF", vbTextCompare) = 0 Then
        ExtrairNFDoPopup = False
        Exit Function
    End If

    ' Parse NF: procura "NF." ou "NF " seguido pelo numero
    pos1 = InStr(1, textoErro, "NF.", vbTextCompare)
    If pos1 > 0 Then
        pos1 = pos1 + 3
    Else
        pos1 = InStr(1, textoErro, "NF ", vbTextCompare)
        If pos1 > 0 Then pos1 = pos1 + 3 Else GoTo FalhaExtracao
    End If
    Do While pos1 <= Len(textoErro) And Mid(textoErro, pos1, 1) = " " : pos1 = pos1 + 1 : Loop
    pos2 = InStr(pos1, textoErro, " ")
    If pos2 = 0 Then pos2 = Len(textoErro) + 1
    nfNum = Trim(Mid(textoErro, pos1, pos2 - pos1))
    On Error Resume Next
    nfNum = CStr(CLng(nfNum))  ' remove zeros a esquerda
    If Err.Number <> 0 Then Err.Clear : GoTo FalhaExtracao
    On Error GoTo 0

    ' Parse Serie: procura "rie:" — cobre tanto "Série:" quanto "Serie:"
    pos1 = InStr(1, textoErro, "rie:", vbTextCompare)
    If pos1 = 0 Then GoTo FalhaExtracao
    pos1 = pos1 + 4
    Do While pos1 <= Len(textoErro) And Mid(textoErro, pos1, 1) = " " : pos1 = pos1 + 1 : Loop
    pos2 = InStr(pos1, textoErro, " ")
    If pos2 = 0 Then pos2 = Len(textoErro) + 1
    nfSerie = Trim(Mid(textoErro, pos1, pos2 - pos1))
    On Error Resume Next
    nfSerie = Format(CLng(nfSerie), "000")  ' garante 3 digitos: 5 -> "005"
    If Err.Number <> 0 Then Err.Clear : GoTo FalhaExtracao
    On Error GoTo 0

    ExtrairNFDoPopup = (Len(nfNum) > 0 And Len(nfSerie) > 0)
    Exit Function

FalhaExtracao:
    ExtrairNFDoPopup = False
End Function

' Imprime a NF de remessa via transacao J1BNFE.
' Extrai NF e Serie automaticamente do popup wnd[2]; usa InputBox como fallback.
' Retorna True se a impressao foi concluida com sucesso.
Private Function ImprimirNFRemessa(ByVal session As Object) As Boolean
    Dim nfNum As String, nfSerie As String, sSerie As String

    On Error GoTo ErroImpressaoNF

    ' Tenta extrair NF e Serie automaticamente do texto do popup
    If Not ExtrairNFDoPopup(session, nfNum, nfSerie) Then
        ' Fallback: solicita ao usuario (caso o SAP use um layout de popup desconhecido)
        nfNum = Trim(InputBox("Nao foi possivel ler a NF automaticamente." & vbCrLf & _
                              "Informe o numero da NF:", "Impressao de Remessa Necessaria"))
        If Len(nfNum) = 0 Then GoTo ImpressaoFalhou
        sSerie = Trim(InputBox("Informe a Serie da NF " & nfNum & ":", "Impressao de Remessa Necessaria"))
        If Len(sSerie) = 0 Then GoTo ImpressaoFalhou
        On Error Resume Next
        nfSerie = Format(CLng(sSerie), "000")
        If Err.Number <> 0 Then nfSerie = sSerie : Err.Clear
        On Error GoTo ErroImpressaoNF
    End If

    ' Navega para J1BNFE — o prefixo /n fecha wnd[1] e wnd[2] automaticamente
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n j1bnfe"
    session.findById("wnd[0]").sendVKey 0
    Application.Wait Now + TimeValue("0:00:02")

    ' Preenche os campos de selecao
    session.findById("wnd[0]/usr/ctxtDATE0-LOW").Text = ""
    session.findById("wnd[0]/usr/txtNFNUM9-LOW").Text = nfNum
    session.findById("wnd[0]/usr/txtSERIE-LOW").Text = nfSerie
    session.findById("wnd[0]/usr/ctxtBUKRS-LOW").Text = "4000"
    session.findById("wnd[0]/usr/ctxtBUKRS-LOW").setFocus
    session.findById("wnd[0]/tbar[1]/btn[8]").press   ' Executar
    Application.Wait Now + TimeValue("0:00:02")

    ' Seleciona a linha e aciona a impressao
    session.findById("wnd[0]/usr/cntlNFE_CONTAINER/shellcont/shell").currentCellColumn = ""
    session.findById("wnd[0]/usr/cntlNFE_CONTAINER/shellcont/shell").selectedRows = "0"
    session.findById("wnd[0]/tbar[0]/btn[86]").press  ' Imprimir
    Application.Wait Now + TimeValue("0:00:02")

    ' Confirma os popups de impressao
    On Error Resume Next
    session.findById("wnd[1]/usr/btnCONTINUE").press
    Application.Wait Now + TimeValue("0:00:01")
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    On Error GoTo ErroImpressaoNF

    Application.Wait Now + TimeValue("0:00:02")
    ImprimirNFRemessa = True
    Exit Function

ImpressaoFalhou:
    ImprimirNFRemessa = False
    Exit Function

ErroImpressaoNF:
    ImprimirNFRemessa = False
End Function
