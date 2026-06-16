Option Explicit
Option Private Module

' ==============================
' PAINEL DE AUTOMACAO — Cockpit
' ==============================
' Botoes disponiveis na aba "Painel":
'   [Processar Pendentes] — le Base_Operacional_Toll_Coonagro.xlsm e exporta para o Python
'                           somente grupos sem cor na coluna A.
'   [Controle Base]       — prepara os controles no topo da aba Base.
' ======================================================

Private Const ARQUIVO_XLSX As String = "Base_Operacional_Toll_Coonagro.xlsm"
Private Const ARQUIVO_XLSX_LEGADO As String = "Base_Dados_Coonagro.xlsm"

Public Const SHEET_PAINEL As String = "Painel"

' Retorna o caminho do Python lendo config.json na pasta do projeto.
' Fallback: tenta "python" no PATH do sistema.
Private Function GetPythonExe() As String
    Dim cfg As String
    Dim fso As Object
    Dim ts As Object
    Dim linha As String
    Dim pasta As String

    pasta = GetPastaRPA()
    cfg = pasta & "\config.json"

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(cfg) Then
        Set ts = fso.OpenTextFile(cfg, 1, False)
        Do While Not ts.AtEndOfStream
            linha = Trim(ts.ReadLine)
            ' Procura linha com "python_exe"
            If InStr(1, linha, """python_exe""", vbTextCompare) > 0 Then
                Dim p1 As Long, p2 As Long
                p1 = InStr(linha, ": """) + 3
                p2 = InStr(p1, linha, """")
                If p1 > 3 And p2 > p1 Then
                    GetPythonExe = Mid(linha, p1, p2 - p1)
                    ts.Close
                    Set ts = Nothing
                    Set fso = Nothing
                    Exit Function
                End If
            End If
        Loop
        ts.Close
        Set ts = Nothing
    End If
    Set fso = Nothing

    ' Fallback: "python" no PATH
    GetPythonExe = "python"
End Function

' Retorna a pasta do projeto RPA (OneDrive).
' Usa Chr(193)="A" com acento para evitar problema de encoding no .bas.
Private Function GetPastaRPA() As String
    Dim aTrabalho As String
    aTrabalho = Chr(193) & "rea de Trabalho"
    GetPastaRPA = Environ("USERPROFILE") & "\OneDrive - The Mosaic Company\" & _
                  aTrabalho & "\Projetos\RPA - Coonagro"
End Function

' Retorna a pasta do SharePoint onde o xlsm esta salvo.
Private Function GetPastaSP() As String
    GetPastaSP = ThisWorkbook.Path
End Function

Private Function ResolverArquivoXlsx(ByVal pastaSharePoint As String) As String
    If LenB(Dir$(pastaSharePoint & "\" & ARQUIVO_XLSX)) > 0 Then
        ResolverArquivoXlsx = ARQUIVO_XLSX
        Exit Function
    End If

    If LenB(Dir$(pastaSharePoint & "\" & ARQUIVO_XLSX_LEGADO)) > 0 Then
        ResolverArquivoXlsx = ARQUIVO_XLSX_LEGADO
        Exit Function
    End If

    ResolverArquivoXlsx = ARQUIVO_XLSX
End Function

' ──────────────────────────────────────────────────────────
' INICIALIZAR PAINEL
' Cria a aba "Painel" e formata o layout. Execute uma unica vez.
' ──────────────────────────────────────────────────────────
Public Sub InicializarPainel()
    Dim wb As Workbook
    Dim ws As Worksheet

    Set wb = ThisWorkbook

    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_PAINEL)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = SHEET_PAINEL
    End If

    With ws
        .Cells.ClearContents
        .Cells.ClearFormats

        ' Cabecalho principal
        With .Cells(1, 1)
            .Value = "Painel de Execucao"
            .Font.Bold = True
            .Font.Size = 12
            .Interior.Color = RGB(30, 90, 160)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With
        .Columns("A:A").ColumnWidth = 28

        ' Instrucao
        With .Cells(1, 2)
            .Value = "Clique em 'Processar Pendentes' para enviar ao cockpit apenas grupos sem cor na coluna A."
            .Font.Italic = True
            .Font.Color = RGB(80, 80, 80)
        End With
        .Columns("B:B").ColumnWidth = 70

        ' Area informativa
        With .Range("A2:B8")
            .Interior.Color = RGB(235, 244, 255)
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        End With

        .Range("A2").Value = "Regra"
        .Range("B2").Value = "O robo processa somente quando a coluna A estiver sem preenchimento."
        .Range("A3").Value = "Ignorados"
        .Range("B3").Value = "Linhas com coluna A colorida (azul/verde/vermelho/laranja) sao ignoradas."
        .Range("A4").Value = "Cores"
        .Range("B4").Value = "Atualizadas automaticamente durante o processamento."
        .Range("A2:A4").Font.Bold = True

        ' Remove botoes existentes (evita duplicatas ao reexecutar)
        Dim shp As Shape
        For Each shp In .Shapes
            shp.Delete
        Next shp

        ' Botao principal — le base automaticamente e executa cockpit
        Dim btn0 As Shape
        Set btn0 = .Shapes.AddShape(msoShapeRoundedRectangle, 10, 10, 220, 28)
        btn0.TextFrame.Characters.Text = "Processar Pendentes"
        btn0.TextFrame.Characters.Font.Bold = True
        btn0.TextFrame.Characters.Font.Size = 10
        btn0.Fill.ForeColor.RGB = RGB(30, 90, 160)
        btn0.Line.Visible = msoFalse
        btn0.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        btn0.OnAction = "BuscarNFsParaCockpit"

        ' Botao "Controle Base" — prepara controles no topo da aba Base
        Dim btn4 As Shape
        Set btn4 = .Shapes.AddShape(msoShapeRoundedRectangle, 240, 10, 120, 28)
        btn4.TextFrame.Characters.Text = ChrW(9632) & "  Controle Base"
        btn4.TextFrame.Characters.Font.Bold = True
        btn4.TextFrame.Characters.Font.Size = 10
        btn4.Fill.ForeColor.RGB = RGB(0, 118, 130)
        btn4.Line.Visible = msoFalse
        btn4.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        btn4.OnAction = "InicializarDashboard"
    End With

    MsgBox "Aba 'Painel' criada com sucesso!" & vbCrLf & vbCrLf & _
            "Use 'Processar Pendentes' para enviar apenas grupos sem cor na coluna A." & vbCrLf & _
           "As cores de status sao atualizadas automaticamente durante a execucao.", vbInformation, "Painel criado"
End Sub


' ──────────────────────────────────────────────────────────
' BUSCAR NFS PENDENTES NA BASE XLSM
' Le Base_Operacional_Toll_Coonagro.xlsm, agrupa por col K (NF Vinculada 5124)
' e coleta grupos que possuem QUALQUER linha sem cor na coluna A.
' Calcula NF baixa (menor col A) e NF alta (maior col A) de cada grupo
' e grava "baixa|alta" por linha no txt para o Python.
' ──────────────────────────────────────────────────────────
Public Sub BuscarNFsParaCockpit()
    Dim caminhoXlsx As String
    Dim caminhoTxt As String
    Dim caminhoScript As String
    Dim nomeArquivoXlsx As String

    nomeArquivoXlsx = ResolverArquivoXlsx(GetPastaSP())
    caminhoXlsx = GetPastaSP() & "\" & nomeArquivoXlsx
    caminhoTxt = GetPastaRPA() & "\nfs_para_processar.txt"
    caminhoScript = GetPastaRPA() & "\force_cockpit.py"

    ' Reutiliza workbook se ja estiver aberto
    Dim wbBase As Workbook
    Dim wsBase As Worksheet
    Dim jaAberto As Boolean
    jaAberto = False

    Dim wbTmp As Workbook
    For Each wbTmp In Application.Workbooks
        If wbTmp.Name = nomeArquivoXlsx Then
            Set wbBase = wbTmp
            jaAberto = True
            Exit For
        End If
    Next wbTmp

    If Not jaAberto Then
        If Dir(caminhoXlsx) = "" Then
            MsgBox "Arquivo nao encontrado:" & vbCrLf & caminhoXlsx, vbExclamation, "Arquivo nao encontrado"
            Exit Sub
        End If
        Application.ScreenUpdating = False
        Set wbBase = Workbooks.Open(caminhoXlsx, ReadOnly:=True, UpdateLinks:=False)
        Application.ScreenUpdating = True
    End If

    On Error Resume Next
    Set wsBase = wbBase.Worksheets("Base")
    On Error GoTo 0

    If wsBase Is Nothing Then
        If Not jaAberto Then wbBase.Close SaveChanges:=False
        MsgBox "Aba 'Base' nao encontrada em " & nomeArquivoXlsx, vbExclamation, "Aba nao encontrada"
        Exit Sub
    End If

    Dim ultimaLinha As Long
    ultimaLinha = wsBase.Cells(wsBase.Rows.Count, 1).End(xlUp).Row

    Dim dicGrupos As Object
    Set dicGrupos = CreateObject("Scripting.Dictionary")
    dicGrupos.CompareMode = vbTextCompare

    Dim r As Long
    Dim nfMae As String
    Dim nfA As String
    Dim semCor As Boolean

    For r = 3 To ultimaLinha
        nfMae = Trim$(CStr(wsBase.Cells(r, 11).Value))
        nfA = Trim$(CStr(wsBase.Cells(r, 1).Value))
        semCor = (wsBase.Cells(r, 1).Interior.Pattern = xlPatternNone Or wsBase.Cells(r, 1).Interior.ColorIndex = xlColorIndexNone)

        If Len(nfMae) = 0 Or Len(nfA) = 0 Then GoTo ProximaLinha

        If semCor Then
            If Not dicGrupos.Exists(nfMae) Then
                dicGrupos.Add nfMae, CreateObject("Scripting.Dictionary")
            End If
            If Not dicGrupos(nfMae).Exists(nfA) Then
                Dim nfANum As Long
                nfANum = 0
                On Error Resume Next
                nfANum = CLng(nfA)
                On Error GoTo 0
                dicGrupos(nfMae).Add nfA, nfANum
            End If
        End If
ProximaLinha:
    Next r

    If Not jaAberto Then wbBase.Close SaveChanges:=False

    If dicGrupos.Count = 0 Then
         MsgBox "Nenhum grupo sem cor na coluna A encontrado na base." & vbCrLf & _
             "Todas as NFs estao com marcacao de cor.", vbInformation, "Sem pendencias"
        Exit Sub
    End If

    Dim fileNum As Integer
    fileNum = FreeFile
    Open caminhoTxt For Output As #fileNum

    Dim previa As String
    Dim total As Long
    total = 0
    Dim chave As Variant

    For Each chave In dicGrupos.Keys
        Dim subDic As Object
        Set subDic = dicGrupos(chave)
        Dim vals As Variant
        vals = subDic.Items

        Dim baixaNum As Long
        Dim altaNum As Long
        baixaNum = vals(0)
        altaNum = vals(0)

        Dim j As Long
        For j = 1 To UBound(vals)
            If vals(j) < baixaNum Then baixaNum = vals(j)
            If vals(j) > altaNum Then altaNum = vals(j)
        Next j

        Print #fileNum, CStr(baixaNum) & "|" & CStr(altaNum)
        If total < 15 Then
            previa = previa & CStr(baixaNum) & " -> " & CStr(altaNum) & vbCrLf
        End If
        total = total + 1
    Next chave
    Close #fileNum

    If total > 15 Then previa = previa & "... (" & total - 15 & " oculto(s))"

    Dim resp As VbMsgBoxResult
    resp = MsgBox(total & " grupo(s) sem cor na coluna A:" & vbCrLf & vbCrLf & _
                  previa & vbCrLf & _
                  "Deseja lancar o force_cockpit.py agora?", _
                  vbYesNo + vbQuestion, "NFs Pendentes - Base xlsm")

    If resp = vbYes Then
        ' Shell.Application.ShellExecute: suporte nativo a caminhos Unicode.
        Dim oSA As Object
        Set oSA = CreateObject("Shell.Application")
        Dim exeCockpit As String
        exeCockpit = GetPastaRPA() & "\force_cockpit.exe"
        If Dir(exeCockpit) <> "" Then
            oSA.ShellExecute exeCockpit, "", GetPastaRPA(), "open", 1
        Else
            ' Fallback: usa Python se o exe nao existir
            oSA.ShellExecute GetPythonExe(), Chr(34) & caminhoScript & Chr(34), GetPastaRPA(), "open", 1
        End If
        Set oSA = Nothing
        MsgBox "force_cockpit iniciado!" & vbCrLf & "Acompanhe o terminal Python.", _
               vbInformation, "Script lancado"
    End If
End Sub

' ──────────────────────────────────────────────────────────
' FORÇAR REATUALIZAÇÃO DO EXCEL
' Executa forcar_atualizacao.py para reler o banco SQLite
' e reescrever o xlsm sem precisar reiniciar o monitoramento.
' ──────────────────────────────────────────────────────────
Public Sub ForcarAtualizacaoExcel()
    Dim caminhoScript As String
    caminhoScript = GetPastaRPA() & "\forcar_atualizacao.py"

    If Dir(caminhoScript) = "" Then
        MsgBox "Script nao encontrado:" & vbCrLf & caminhoScript, vbExclamation
        Exit Sub
    End If

    Dim oSA As Object
    Set oSA = CreateObject("Shell.Application")
    Dim exeRecarregar As String
    exeRecarregar = GetPastaRPA() & "\forcar_atualizacao.exe"
    If Dir(exeRecarregar) <> "" Then
        oSA.ShellExecute exeRecarregar, "", GetPastaRPA(), "open", 1
    Else
        ' Fallback: usa Python se o exe nao existir
        oSA.ShellExecute GetPythonExe(), Chr(34) & caminhoScript & Chr(34), GetPastaRPA(), "open", 1
    End If
    Set oSA = Nothing

    MsgBox "Reatualização iniciada!" & vbCrLf & _
           "O Excel será atualizado em instantes.", _
           vbInformation, "Recarregar Base"
End Sub

' ──────────────────────────────────────────────────────────
' MARCAR AUTCARR_OK
' Selecione as linhas na aba Base (qualquer celula da linha serve)
' e execute esta macro. Pinta a coluna A de AZUL para que o robo
' saiba que a Passagem 1 ja foi feita e processe apenas a Passagem 2.
' ──────────────────────────────────────────────────────────
Public Sub MarcarComoAutcarrOK()
    Dim ws As Worksheet
    Dim sel As Range
    Dim cel As Range
    Dim linhas As Collection
    Dim r As Long
    Dim count As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Base")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Aba 'Base' nao encontrada.", vbExclamation
        Exit Sub
    End If

    If TypeName(Selection) <> "Range" Then
        MsgBox "Selecione as linhas desejadas na aba Base antes de executar.", vbInformation
        Exit Sub
    End If

    Set sel = Selection
    count = 0

    ' Coleta linhas unicas da selecao
    Set linhas = New Collection
    For Each cel In sel.Rows
        r = cel.Row
        On Error Resume Next
        linhas.Add r, CStr(r)   ' chave unica evita duplicatas
        On Error GoTo 0
    Next cel

    Dim item As Variant
    For Each item In linhas
        With ws.Cells(CLng(item), "A")
            If Len(Trim$(CStr(.Value))) > 0 Then
                .Interior.Color = RGB(66, 165, 245)   ' AZUL = AUTCARR_OK
                .Font.Color = RGB(255, 255, 255)
                count = count + 1
            End If
        End With
    Next item

    If count = 0 Then
        MsgBox "Nenhuma linha com Nota Fiscal encontrada na selecao.", vbInformation
    Else
        ws.Parent.Save
        MsgBox count & " linha(s) marcadas como AUTCARR_OK (azul)." & vbCrLf & _
               "Clique em 'Iniciar' para o robo executar a Passagem 2 nessas notas.", _
               vbInformation, "Marcacao concluida"
    End If
End Sub


' ──────────────────────────────────────────────────────────
' EXPORTAR DADOS LANCADOS
' Gera arquivos com os dados da aba Concluidos (cores de status preservadas).
' Arquivos nomeados com timestamp e salvos em "Exportados\dd.mm.yy\".
' ──────────────────────────────────────────────────────────
Public Sub ExportarDadosLancados()
    Dim resp As VbMsgBoxResult

    resp = MsgBox("Escolha o tipo de exportacao:" & vbNewLine & vbNewLine & _
                  "Sim = Exportar Excel (.xlsx)" & vbNewLine & _
                  "Nao = Exportar PDF (.pdf)" & vbNewLine & _
                  "Cancelar = Sair", _
                  vbYesNoCancel + vbQuestion, "Exportar Dados Lancados")

    Select Case resp
        Case vbYes
            ExportarExcel
        Case vbNo
            ExportarPDF
    End Select
End Sub

Public Sub ExportarExcel()
    Dim ws As Worksheet
    Dim nomeBase As String
    Dim caminhoXlsx As String
    Dim pastaExport As String

    If Not PrepararContextoExportacao(ws, nomeBase) Then Exit Sub

    caminhoXlsx = EscolherCaminhoExportacao( _
        nomeBase, _
        "xlsx", _
        "Salvar exportacao Excel", _
        "Arquivo Excel (*.xlsx), *.xlsx" _
    )
    If Len(caminhoXlsx) = 0 Then Exit Sub

    pastaExport = ObterPastaArquivo(caminhoXlsx)

    Application.ScreenUpdating = False
    If ExportarBaseComoXlsx(ws, caminhoXlsx) Then
        MostrarResumoExportacao "Exportacao Excel concluida!", _
            "XLSX: " & caminhoXlsx, _
            pastaExport, "Exportar Excel"
    End If
    Application.ScreenUpdating = True
End Sub

Public Sub ExportarPDF()
    Dim ws As Worksheet
    Dim nomeBase As String
    Dim caminhoPdf As String
    Dim pastaExport As String

    If Not PrepararContextoExportacao(ws, nomeBase) Then Exit Sub

    caminhoPdf = EscolherCaminhoExportacao( _
        nomeBase, _
        "pdf", _
        "Salvar exportacao PDF", _
        "Arquivo PDF (*.pdf), *.pdf" _
    )
    If Len(caminhoPdf) = 0 Then Exit Sub

    pastaExport = ObterPastaArquivo(caminhoPdf)

    Application.ScreenUpdating = False
    If ExportarBaseComoPDF(ws, caminhoPdf) Then
        MostrarResumoExportacao "Exportacao PDF concluida!", _
            "PDF: " & caminhoPdf, _
            pastaExport, "Exportar PDF"
    End If
    Application.ScreenUpdating = True
End Sub

Private Function PrepararContextoExportacao(ByRef ws As Worksheet, ByRef nomeBase As String) As Boolean
    Set ws = GetWorksheetParaExportacao()

    If ws Is Nothing Then
        MsgBox "Use a aba Concluidos para exportar os grupos finalizados.", vbExclamation
        Exit Function
    End If

    nomeBase = "Exportacao_Coonagro_" & NomeAbaParaArquivo(ws.Name) & "_" & Format(Now, "yyyymmdd_hhmmss")
    PrepararContextoExportacao = True
End Function

Private Function GetWorksheetParaExportacao() As Worksheet
    Dim wsAtiva As Worksheet

    On Error Resume Next
    Set wsAtiva = ActiveSheet
    On Error GoTo 0

    If Not wsAtiva Is Nothing Then
        If EhAbaExportavel(wsAtiva.Name) Then
            Set GetWorksheetParaExportacao = wsAtiva
            Exit Function
        End If
    End If

    On Error Resume Next
    Set GetWorksheetParaExportacao = ThisWorkbook.Worksheets(NomeAbaConcluidosExport())
    On Error GoTo 0
End Function

Private Function EhAbaExportavel(ByVal nomeAba As String) As Boolean
    EhAbaExportavel = (StrComp(nomeAba, NomeAbaConcluidosExport(), vbTextCompare) = 0)
End Function

Private Function NomeAbaConcluidosExport() As String
    NomeAbaConcluidosExport = "Conclu" & ChrW(237) & "dos"
End Function

Private Function NomeAbaParaArquivo(ByVal nomeAba As String) As String
    If StrComp(nomeAba, NomeAbaConcluidosExport(), vbTextCompare) = 0 Then
        NomeAbaParaArquivo = "Concluidos"
    Else
        NomeAbaParaArquivo = Replace(nomeAba, " ", "_")
    End If
End Function

Private Function EscolherCaminhoExportacao(ByVal nomeBase As String, ByVal extensao As String, ByVal titulo As String, ByVal filtro As String) As String
    Dim caminhoSelecionado As Variant
    Dim sugestao As String

    sugestao = GetPastaSP() & "\" & nomeBase & "." & extensao
    caminhoSelecionado = Application.GetSaveAsFilename( _
        InitialFileName:=sugestao, _
        FileFilter:=filtro, _
        Title:=titulo _
    )

    If VarType(caminhoSelecionado) = vbBoolean Then Exit Function

    EscolherCaminhoExportacao = CStr(caminhoSelecionado)
    If LCase$(Right$(EscolherCaminhoExportacao, Len(extensao) + 1)) <> "." & LCase$(extensao) Then
        EscolherCaminhoExportacao = EscolherCaminhoExportacao & "." & extensao
    End If
End Function

Private Function ObterPastaArquivo(ByVal caminhoArquivo As String) As String
    Dim posBarra As Long

    posBarra = InStrRev(caminhoArquivo, "\")
    If posBarra > 0 Then
        ObterPastaArquivo = Left$(caminhoArquivo, posBarra - 1)
    Else
        ObterPastaArquivo = GetPastaSP()
    End If
End Function

Private Function ExportarBaseComoXlsx(ByVal ws As Worksheet, ByVal caminhoXlsx As String) As Boolean
    Dim wbNovo As Workbook
    Dim wsNova As Worksheet
    Dim shp As Shape

    On Error GoTo TrataErro

    ws.Copy
    Set wbNovo = ActiveWorkbook
    Set wsNova = wbNovo.Worksheets(1)

    For Each shp In wsNova.Shapes
        shp.Delete
    Next shp

    With wsNova.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .Zoom = False
    End With

    Application.DisplayAlerts = False
    wbNovo.SaveAs Filename:=caminhoXlsx, FileFormat:=xlOpenXMLWorkbook
    wbNovo.Close SaveChanges:=False
    Application.DisplayAlerts = True

    ExportarBaseComoXlsx = True
    Exit Function

TrataErro:
    Application.DisplayAlerts = True
    On Error Resume Next
    If Not wbNovo Is Nothing Then wbNovo.Close SaveChanges:=False
    On Error GoTo 0
    MsgBox "Falha ao exportar Excel: " & Err.Description, vbExclamation, "Exportar Excel"
End Function

Private Function ExportarBaseComoPDF(ByVal ws As Worksheet, ByVal caminhoPdf As String) As Boolean
    Dim shp As Shape
    Dim ultimaLinha As Long
    Dim ultimaCol As Long

    On Error GoTo TrataErro

    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ultimaCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1

    For Each shp In ws.Shapes
        If Left$(shp.Name, 4) = "ctl_" Then shp.Visible = msoFalse
    Next shp

    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(ultimaLinha, ultimaCol)).Address
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .Zoom = False
    End With

    ws.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=caminhoPdf, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    ExportarBaseComoPDF = True

TrataSaida:
    On Error Resume Next
    For Each shp In ws.Shapes
        If Left$(shp.Name, 4) = "ctl_" Then shp.Visible = msoTrue
    Next shp
    ws.PageSetup.PrintArea = ""
    On Error GoTo 0
    Exit Function

TrataErro:
    MsgBox "Falha ao exportar PDF: " & Err.Description, vbExclamation, "Exportar PDF"
    GoTo TrataSaida
End Function

Private Sub MostrarResumoExportacao(ByVal titulo As String, ByVal detalhe As String, ByVal pastaExport As String, ByVal tituloMsg As String)
    Dim resp As VbMsgBoxResult

    resp = MsgBox(titulo & vbNewLine & vbNewLine & _
                  detalhe & vbNewLine & vbNewLine & _
                  "Deseja abrir a pasta do arquivo?", _
                  vbYesNo + vbInformation, tituloMsg)

    If resp = vbYes Then
        Shell "explorer.exe " & Chr(34) & pastaExport & Chr(34), vbNormalFocus
    End If
End Sub

' ──────────────────────────────────────────────────────────
' MARCAR LINHAS SELECIONADAS COMO JA LANCADAS
' Grava as NFs das linhas selecionadas em nfs_lancadas.txt
' e dispara o Recarregar para remove-las da planilha.
' ──────────────────────────────────────────────────────────
Public Sub MarcarSelecionadasComoLancadas()
    Dim ws As Worksheet
    Dim sel As Range
    Dim cel As Range
    Dim nfVal As String
    Dim dic As Object
    Dim chave As Variant
    Dim caminhoLancadas As String
    Dim fn As Integer
    Dim totalMarcadas As Long

    Set ws = ThisWorkbook.Worksheets(PLAN_SHEET)
    Set sel = Selection

    ' Verifica se ha selecao valida na planilha Base
    If sel Is Nothing Or sel.Worksheet.Name <> ws.Name Then
        MsgBox "Selecione as linhas a marcar como Lancadas na aba Base.", vbExclamation, "Selecao invalida"
        Exit Sub
    End If

    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = vbTextCompare

    ' Coleta NFs unicas da coluna A nas linhas selecionadas (a partir da linha 3)
    For Each cel In sel.Cells
        If cel.Row >= 3 Then
            nfVal = Trim$(CStr(ws.Cells(cel.Row, COL_LINHA_VALIDA).Value))
            If Len(nfVal) > 0 And Not dic.Exists(nfVal) Then
                dic.Add nfVal, 1
            End If
        End If
    Next cel

    If dic.Count = 0 Then
        MsgBox "Nenhuma NF valida encontrada nas linhas selecionadas.", vbExclamation, "Nada a marcar"
        Exit Sub
    End If

    ' Confirma com o usuario
    totalMarcadas = dic.Count
    Dim resp As VbMsgBoxResult
    resp = MsgBox(totalMarcadas & " NF(s) serao marcadas como Lancadas e removidas da planilha:" & vbNewLine & vbNewLine & _
                  Join(dic.Keys, ", ") & vbNewLine & vbNewLine & _
                  "Deseja continuar?", vbYesNo + vbQuestion, "Confirmar Marcacao")
    If resp <> vbYes Then Exit Sub

    ' Grava no nfs_lancadas.txt
    caminhoLancadas = GetPastaRPA() & "\nfs_lancadas.txt"
    fn = FreeFile
    Open caminhoLancadas For Append As #fn
    For Each chave In dic.Keys
        Print #fn, chave
    Next chave
    Close #fn

    Set dic = Nothing

    MsgBox totalMarcadas & " NF(s) registradas em nfs_lancadas.txt." & vbNewLine & _
           "Clique OK e em seguida use o botao Recarregar para atualizar a planilha.", _
           vbInformation, "Marcacao concluida"
End Sub
