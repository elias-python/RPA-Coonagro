Option Explicit
Option Private Module

' ======================================================
' modDashboard (modo sem aba Dashboard)
' Mantem controles e informacoes no topo da aba Base.
' ======================================================

Private Const BASE_SHEET_NAME As String = "Base"
Private Const LEGACY_DASHBOARD_SHEET As String = "Dashboard"
Private Const STATUS_DIA_PREFIX As String = "status_dia_"

Private mAtualizandoDashboard As Boolean

Private Const SHP_BG As String = "ctl_bg"
Private Const SHP_INFO_UPDATED As String = "ctl_info_updated"
Private Const SHP_INFO_STATUS As String = "ctl_info_status"
Private Const SHP_INFO_SUMMARY As String = "ctl_info_summary"
Private Const SHP_BTN_START As String = "ctl_btn_start"
Private Const SHP_BTN_PAUSE As String = "ctl_btn_pause"
Private Const SHP_BTN_STOP As String = "ctl_btn_stop"
Private Const SHP_BTN_REFRESH As String = "ctl_btn_refresh"
Private Const SHP_BTN_COCKPIT As String = "ctl_btn_cockpit"
Private Const SHP_BTN_EXPORT_XLSX As String = "ctl_btn_export_xlsx"
Private Const SHP_BTN_EXPORT_PDF  As String = "ctl_btn_export_pdf"
Private Const SHP_BTN_CLEAR_CONCLUIDOS As String = "ctl_btn_clear_concluidos"
Private Const SHP_BTN_RELOAD   As String = "ctl_btn_reload"
Private Const UI_PROTECTION_PASSWORD As String = "RPA_Coonagro_UI"
Private Const MAINTENANCE_SHORTCUT_OPEN As String = "^+m"
Private Const MAINTENANCE_SHORTCUT_OPEN_ALT As String = "^+M"
Private Const MAINTENANCE_SHORTCUT_CLOSE As String = "^+l"
Private Const MAINTENANCE_SHORTCUT_CLOSE_ALT As String = "^+L"
Private Const MAINTENANCE_SHORTCUT_CLOSE_BRACED As String = "^+{L}"

' Executado automaticamente ao abrir o arquivo.
' Preserva o layout existente e apenas recompõe controles operacionais quando
' algum shape obrigatório estiver ausente.
Public Sub Auto_Open()
    On Error GoTo SaidaSegura
    Application.ScreenUpdating = False
    RegistrarAtalhosManutencao
    PrepararModoEdicaoInterno
    RemoverAbaDashboardLegada
    Dim ws As Worksheet
    Dim wsConcluidos As Worksheet
    Set ws = GetBaseSheet()
    If Not ws Is Nothing Then
        EnsureTopo ws
        ws.Activate
    End If
    Set wsConcluidos = GetConcluidosSheet(True)
    If Not wsConcluidos Is Nothing Then EnsureTopoConcluidos wsConcluidos
    AtualizarDashboard

SaidaSegura:
    AplicarModoOperacao
    Application.ScreenUpdating = True
End Sub

Public Sub InicializarDashboard()
    Dim ws As Worksheet
    Dim wsConcluidos As Worksheet

    On Error GoTo SaidaSegura
    Application.ScreenUpdating = False
    RegistrarAtalhosManutencao
    PrepararModoEdicaoInterno

    Set ws = GetBaseSheet()
    If ws Is Nothing Then
        MsgBox "Aba 'Base' nao encontrada.", vbExclamation
        GoTo SaidaSegura
    End If

    RemoverAbaDashboardLegada
    Set wsConcluidos = GetConcluidosSheet(True)
    If Not wsConcluidos Is Nothing Then EnsureTopoConcluidos wsConcluidos
    EnsureTopo ws
    AtualizarDashboard
    ws.Activate

SaidaSegura:
    AplicarModoOperacao
    Application.ScreenUpdating = True
End Sub

Public Sub Auto_Close()
    LimparAtalhosManutencao
    RestaurarInterfaceExcel
End Sub

Public Sub AtivarModoOperacional()
    Application.ScreenUpdating = False
    RegistrarAtalhosManutencao
    AplicarModoOperacao
    Application.ScreenUpdating = True
End Sub

Public Sub LiberarModoEdicao()
    Application.ScreenUpdating = False
    RegistrarAtalhosManutencao
    PrepararModoEdicaoInterno
    Application.ScreenUpdating = True
End Sub

Public Sub EntrarModoManutencao()
    Dim loginInformado As String
    Dim senhaInformada As String

    loginInformado = Trim$(InputBox$( _
        "Login de manutencao:", _
        "Acesso de manutencao", _
        Environ$("Username") _
    ))
    If Len(loginInformado) = 0 Then Exit Sub

    senhaInformada = InputBox$("Senha de manutencao:", "Acesso de manutencao")
    If Len(senhaInformada) = 0 Then Exit Sub

    If Not CredenciaisManutencaoValidas(loginInformado, senhaInformada) Then
        MsgBox "Login ou senha invalidos.", vbCritical, "Acesso de manutencao"
        Exit Sub
    End If

    LiberarModoEdicao
    MsgBox _
        "Modo de manutencao liberado." & vbNewLine & _
        "Use Ctrl+Shift+L para voltar ao modo operacional.", _
        vbInformation, _
        "Acesso de manutencao"
End Sub

Public Sub SairModoManutencao()
    RegistrarAtalhosManutencao
    AtivarModoOperacional
End Sub

Public Sub AtualizarDashboard()
    Dim ws As Worksheet
    Dim wsConcluidos As Worksheet

    If mAtualizandoDashboard Then Exit Sub
    mAtualizandoDashboard = True

    On Error GoTo SaidaSegura

    Set ws = GetBaseSheet()
    If ws Is Nothing Then GoTo SaidaSegura

    EnsureTopo ws
    SetCardText ws, SHP_INFO_UPDATED, "Atualizado: " & Format$(Now, "dd/mm/yyyy hh:mm:ss"), RGB(209, 232, 255)
    AtualizarCardsBase ws

    On Error Resume Next
    ArquivarGruposConcluidos ws
    On Error GoTo SaidaSegura

    Set wsConcluidos = GetConcluidosSheet(False)
    If Not wsConcluidos Is Nothing Then
        On Error Resume Next
        EnsureTopoConcluidos wsConcluidos
        On Error GoTo SaidaSegura
    End If

    AtualizarCardsBase ws

SaidaSegura:
    mAtualizandoDashboard = False
End Sub

Public Sub LimparListaConcluidos()
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim resposta As VbMsgBoxResult

    Set ws = GetConcluidosSheet(False)
    If ws Is Nothing Then
        MsgBox "Aba 'Concluidos' nao encontrada.", vbExclamation, "Limpar Concluidos"
        Exit Sub
    End If

    ultimaLinha = UltimaLinhaUsada(ws)
    If ultimaLinha < 3 Then
        MsgBox "Nao ha itens para limpar na aba Concluidos.", vbInformation, "Limpar Concluidos"
        Exit Sub
    End If

    resposta = MsgBox( _
        "Isso vai limpar a lista atual da aba Concluidos." & vbNewLine & _
        "Os arquivos ja exportados nao serao alterados." & vbNewLine & vbNewLine & _
        "Deseja continuar?", _
        vbYesNo + vbQuestion, _
        "Limpar Concluidos" _
    )
    If resposta <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    ws.Rows("3:" & ultimaLinha).Delete
    EnsureTopoConcluidos ws
    ThisWorkbook.Save
    Application.ScreenUpdating = True
End Sub

Private Sub AtualizarCardsBase(ByVal ws As Worksheet)
    Dim totalNf As Long
    Dim pend As Long
    Dim ok As Long
    Dim errCor As Long

    SetCardText ws, SHP_INFO_UPDATED, "Atualizado: " & Format$(Now, "dd/mm/yyyy hh:mm:ss"), RGB(209, 232, 255)
    AtualizarTextoStatusPainel ws, TextoStatusSapAtual(ws)
    On Error Resume Next
    ColetarIndicadoresDia ws, totalNf, pend, ok, errCor
    If Err.Number <> 0 Then
        totalNf = 0
        pend = 0
        ok = 0
        errCor = 0
        Err.Clear
    End If
    On Error GoTo 0
    SetCardText ws, SHP_INFO_SUMMARY, "Total NF: " & totalNf & " | Pend: " & pend & " | OK: " & ok & " | Erros: " & errCor, RGB(232, 239, 247)
    DoEvents
End Sub

Private Sub ArquivarGruposConcluidos(ByVal wsBase As Worksheet)
    Dim ultimaLinha As Long
    Dim ultimaCol As Long
    Dim colChave As Long
    Dim blocos As Collection
    Dim bloco As Variant
    Dim r As Long
    Dim inicioBloco As Long
    Dim fimBloco As Long
    Dim wsDestino As Worksheet
    Dim linhaDestino As Long
    Dim dicNFLanc As Object
    Dim caminhoLancadas As String
    Dim fnLanc As Integer
    Dim chaveLanc As Variant
    Dim houveMovimento As Boolean

    If Not PodeArquivarConcluidos(wsBase) Then Exit Sub

    ultimaLinha = wsBase.Cells(wsBase.Rows.Count, 1).End(xlUp).Row
    If ultimaLinha < 3 Then Exit Sub

    ultimaCol = wsBase.UsedRange.Column + wsBase.UsedRange.Columns.Count - 1
    If ultimaCol < 1 Then ultimaCol = 1
    colChave = ColunaCabecalho(wsBase, "Chave de Acesso", ultimaCol)

    Set blocos = New Collection
    Set dicNFLanc = CreateObject("Scripting.Dictionary")
    dicNFLanc.CompareMode = vbTextCompare

    r = 3
    Do While r <= ultimaLinha
        If LinhaConcluida(wsBase, r) Then
            inicioBloco = r
            fimBloco = r
            Do While fimBloco + 1 <= ultimaLinha And LinhaConcluida(wsBase, fimBloco + 1)
                fimBloco = fimBloco + 1
            Loop

            bloco = Array(inicioBloco, fimBloco)
            blocos.Add bloco

            ColetarNFsDoBloco wsBase, inicioBloco, fimBloco, dicNFLanc
            r = fimBloco + 1
        Else
            r = r + 1
        End If
    Loop

    If blocos.Count = 0 Then Exit Sub

    Set wsDestino = GetConcluidosSheet(True)
    If wsDestino Is Nothing Then Exit Sub

    PrepararCabecalhoConcluidos wsBase, wsDestino, ultimaCol
    linhaDestino = ProximaLinhaConcluidos(wsDestino)

    Application.ScreenUpdating = False

    For Each bloco In blocos
        inicioBloco = CLng(bloco(0))
        fimBloco = CLng(bloco(1))

        wsDestino.Range( _
            wsDestino.Cells(linhaDestino, 1), _
            wsDestino.Cells(linhaDestino + (fimBloco - inicioBloco), ultimaCol) _
        ).Value = wsBase.Range( _
            wsBase.Cells(inicioBloco, 1), _
            wsBase.Cells(fimBloco, ultimaCol) _
        ).Value

        If colChave > 0 Then
            ReescreverColunaChaveComoTexto wsBase, wsDestino, inicioBloco, fimBloco, linhaDestino, colChave
        End If

        linhaDestino = linhaDestino + (fimBloco - inicioBloco + 1)
        linhaDestino = linhaDestino + 1
        houveMovimento = True
    Next bloco

    If dicNFLanc.Count > 0 Then
        RegistrarNFsConcluidasHoje dicNFLanc
        caminhoLancadas = GetPastaRPA() & "\nfs_lancadas.txt"
        fnLanc = FreeFile
        Open caminhoLancadas For Append As #fnLanc
        For Each chaveLanc In dicNFLanc.Keys
            Print #fnLanc, chaveLanc
        Next chaveLanc
        Close #fnLanc
    End If

    For r = blocos.Count To 1 Step -1
        bloco = blocos(r)
        inicioBloco = CLng(bloco(0))
        fimBloco = CLng(bloco(1))
        wsBase.Rows(inicioBloco & ":" & fimBloco).Delete

        If inicioBloco <= wsBase.Cells(wsBase.Rows.Count, 1).End(xlUp).Row Then
            If Application.WorksheetFunction.CountA(wsBase.Rows(inicioBloco)) = 0 Then
                wsBase.Rows(inicioBloco).Delete
            End If
        End If
    Next r

    CompactarTopoBase wsBase

    If houveMovimento Then
        ThisWorkbook.Save
    End If

    Application.CutCopyMode = False
    Application.ScreenUpdating = True
End Sub

Private Sub CompactarTopoBase(ByVal wsBase As Worksheet)
    Dim primeiraLinhaComDados As Long
    Dim ultimaLinha As Long

    ultimaLinha = wsBase.Cells(wsBase.Rows.Count, 1).End(xlUp).Row
    If ultimaLinha < 3 Then Exit Sub

    primeiraLinhaComDados = 3
    Do While primeiraLinhaComDados <= ultimaLinha
        If Application.WorksheetFunction.CountA(wsBase.Rows(primeiraLinhaComDados)) > 0 Then Exit Do
        primeiraLinhaComDados = primeiraLinhaComDados + 1
    Loop

    If primeiraLinhaComDados > 3 And primeiraLinhaComDados <= ultimaLinha Then
        wsBase.Rows("3:" & primeiraLinhaComDados - 1).Delete
    End If
End Sub

Private Function LinhaConcluida(ByVal ws As Worksheet, ByVal rowNum As Long) As Boolean
    With ws.Cells(rowNum, 1)
        LinhaConcluida = (.Interior.Pattern <> xlPatternNone And _
            .Interior.Color = RGB(102, 187, 106) And _
            Len(Trim$(CStr(.Value))) > 0)
    End With
End Function

Private Sub ColetarNFsDoBloco(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal dic As Object)
    Dim r As Long
    Dim nfVal As String

    For r = startRow To endRow
        nfVal = Trim$(CStr(ws.Cells(r, 1).Value))
        If Len(nfVal) > 0 And Not dic.Exists(nfVal) Then
            dic.Add nfVal, 1
        End If
    Next r
End Sub

Private Function GetConcluidosSheet(Optional ByVal createIfMissing As Boolean = False) As Worksheet
    Dim nomeAba As String

    nomeAba = NomeAbaConcluidos()

    On Error Resume Next
    Set GetConcluidosSheet = ThisWorkbook.Worksheets(nomeAba)
    On Error GoTo 0

    If GetConcluidosSheet Is Nothing And createIfMissing Then
        On Error Resume Next
        Set GetConcluidosSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        If Not GetConcluidosSheet Is Nothing Then
            GetConcluidosSheet.Name = nomeAba
        End If
        On Error GoTo 0
    End If
End Function

Private Function NomeAbaConcluidos() As String
    NomeAbaConcluidos = "Conclu" & ChrW(237) & "dos"
End Function

Private Sub PrepararCabecalhoConcluidos(ByVal wsBase As Worksheet, ByVal wsDestino As Worksheet, ByVal ultimaCol As Long)
    wsDestino.Range(wsDestino.Cells(2, 1), wsDestino.Cells(2, ultimaCol)).Value = _
        wsBase.Range(wsBase.Cells(2, 1), wsBase.Cells(2, ultimaCol)).Value
End Sub

Private Function ColunaCabecalho(ByVal ws As Worksheet, ByVal headerName As String, ByVal ultimaCol As Long) As Long
    Dim col As Long

    For col = 1 To ultimaCol
        If StrComp(Trim$(CStr(ws.Cells(2, col).Text)), headerName, vbTextCompare) = 0 Then
            ColunaCabecalho = col
            Exit Function
        End If
    Next col
End Function

Private Sub ReescreverColunaChaveComoTexto(ByVal wsBase As Worksheet, ByVal wsDestino As Worksheet, ByVal inicioBloco As Long, ByVal fimBloco As Long, ByVal linhaDestino As Long, ByVal colChave As Long)
    Dim r As Long
    Dim valorChave As String

    wsDestino.Range( _
        wsDestino.Cells(linhaDestino, colChave), _
        wsDestino.Cells(linhaDestino + (fimBloco - inicioBloco), colChave) _
    ).NumberFormat = "@"

    For r = 0 To (fimBloco - inicioBloco)
        valorChave = Trim$(CStr(wsBase.Cells(inicioBloco + r, colChave).Text))
        wsDestino.Cells(linhaDestino + r, colChave).Value = valorChave
    Next r
End Sub

Private Function ProximaLinhaConcluidos(ByVal ws As Worksheet) As Long
    Dim ultima As Long

    ultima = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ultima < 3 Then
        ProximaLinhaConcluidos = 3
    ElseIf Application.WorksheetFunction.CountA(ws.Rows(ultima)) = 0 Then
        ProximaLinhaConcluidos = ultima
    Else
        ProximaLinhaConcluidos = ultima + 1
    End If
End Function

Private Function UltimaLinhaUsada(ByVal ws As Worksheet) As Long
    Dim ultimaCelula As Range

    On Error Resume Next
    Set ultimaCelula = ws.Cells.Find(What:="*", _
                                     After:=ws.Cells(1, 1), _
                                     LookIn:=xlFormulas, _
                                     LookAt:=xlPart, _
                                     SearchOrder:=xlByRows, _
                                     SearchDirection:=xlPrevious, _
                                     MatchCase:=False)
    On Error GoTo 0

    If Not ultimaCelula Is Nothing Then
        UltimaLinhaUsada = ultimaCelula.Row
    End If
End Function

Public Sub AtualizarStatus(ByVal texto As String)
    Dim ws As Worksheet
    Set ws = GetBaseSheet()
    If ws Is Nothing Then Exit Sub

    EnsureTopo ws

    AtualizarTextoStatusPainel ws, texto
    DoEvents
End Sub

Public Sub AtualizarContadores(ByVal erros As Long, ByVal gruposLancados As Long, ByVal notasLancadas As Long)
    Dim ws As Worksheet
    Dim totalDia As Long
    Set ws = GetBaseSheet()
    If ws Is Nothing Then Exit Sub

    EnsureTopo ws

    Dim totalNf As Long, pend As Long, ok As Long, err As Long
    ColetarIndicadoresDia ws, totalNf, pend, ok, err
    totalDia = TotalNotasLancadasDia(ws)

    SetCardText ws, SHP_INFO_SUMMARY, _
        "Grupos: " & gruposLancados & " | Notas: " & notasLancadas & " | Pend: " & pend & " | Erros: " & err & " | Dia: " & totalDia, _
        RGB(232, 239, 247)
    DoEvents
End Sub

Private Function GetBaseSheet() As Worksheet
    On Error Resume Next
    Set GetBaseSheet = ThisWorkbook.Worksheets(BASE_SHEET_NAME)
    On Error GoTo 0
End Function

Private Function GetPastaRPA() As String
    Dim aTrabalho As String
    aTrabalho = Chr(193) & "rea de Trabalho"
    GetPastaRPA = Environ("USERPROFILE") & "\OneDrive - The Mosaic Company\" & _
                  aTrabalho & "\Projetos\RPA - Coonagro"
End Function

Private Function CaminhoXmlsPendentesRecarregar() As String
    CaminhoXmlsPendentesRecarregar = GetPastaRPA() & "\xmls_pendentes_recarregar.txt"
End Function

Private Function LerXmlsPendentesRecarregar() As Long
    Dim caminho As String
    Dim fn As Integer
    Dim linha As String

    caminho = CaminhoXmlsPendentesRecarregar()
    If Dir(caminho) = "" Then Exit Function

    On Error GoTo SaidaSegura
    fn = FreeFile
    Open caminho For Input As #fn
    If Not EOF(fn) Then Line Input #fn, linha
    Close #fn

    If IsNumeric(Trim$(linha)) Then
        LerXmlsPendentesRecarregar = CLng(Val(Trim$(linha)))
    End If
    Exit Function

SaidaSegura:
    On Error Resume Next
    If fn <> 0 Then Close #fn
    On Error GoTo 0
End Function

Private Function TextoStatusSapAtual(ByVal ws As Worksheet) As String
    Dim txt As String
    Dim posExtra As Long

    txt = ""
    On Error Resume Next
    txt = CStr(ws.Shapes(SHP_INFO_STATUS).TextFrame.Characters.Text)
    If Len(Trim$(txt)) = 0 Then txt = CStr(ws.Shapes(SHP_INFO_STATUS).TextFrame2.TextRange.Text)
    On Error GoTo 0

    txt = Trim$(txt)
    If Len(txt) = 0 Then
        TextoStatusSapAtual = "aguardando"
        Exit Function
    End If

    If UCase$(Left$(txt, 11)) = "STATUS SAP:" Then txt = Trim$(Mid$(txt, 12))
    posExtra = InStr(1, txt, " | ", vbTextCompare)
    If posExtra > 0 Then txt = Trim$(Left$(txt, posExtra - 1))
    If Len(txt) = 0 Then txt = "aguardando"

    TextoStatusSapAtual = txt
End Function

Private Function CorStatusPainel(ByVal texto As String) As Long
    Dim u As String

    u = UCase$(texto)

    If InStr(u, "ERRO") > 0 Or InStr(u, "FATAL") > 0 Then
        CorStatusPainel = RGB(255, 110, 110)
    ElseIf InStr(u, "RODANDO") > 0 Or InStr(u, "PROCESSANDO") > 0 Then
        CorStatusPainel = RGB(120, 235, 140)
    ElseIf InStr(u, "FINALIZADO") > 0 Then
        CorStatusPainel = RGB(130, 200, 255)
    ElseIf InStr(u, "PAUSADO") > 0 Then
        CorStatusPainel = RGB(255, 220, 110)
    ElseIf InStr(u, "PARAND") > 0 Then
        CorStatusPainel = RGB(255, 180, 100)
    Else
        CorStatusPainel = RGB(205, 218, 235)
    End If
End Function

Private Function MontarTextoStatusPainel(ByVal textoSap As String) As String
    Dim xmlPendentes As Long

    textoSap = Trim$(textoSap)
    If Len(textoSap) = 0 Then textoSap = "aguardando"

    MontarTextoStatusPainel = "Status SAP: " & textoSap

    xmlPendentes = LerXmlsPendentesRecarregar()
    If xmlPendentes > 0 Then
        MontarTextoStatusPainel = MontarTextoStatusPainel & _
            " | Notas novas: " & xmlPendentes & _
            " | Recarregue a planilha"
    End If
End Function

Private Sub AtualizarTextoStatusPainel(ByVal ws As Worksheet, ByVal textoSap As String)
    SetCardText ws, SHP_INFO_STATUS, MontarTextoStatusPainel(textoSap), CorStatusPainel(textoSap)
End Sub

Private Function PodeArquivarConcluidos(ByVal wsBase As Worksheet) As Boolean
    Dim txtStatus As String

    txtStatus = ""
    On Error Resume Next
    txtStatus = UCase$(CStr(wsBase.Shapes(SHP_INFO_STATUS).TextFrame.Characters.Text))
    On Error GoTo 0

    If InStr(txtStatus, "PROCESSANDO") > 0 Then Exit Function
    If InStr(txtStatus, "RODANDO") > 0 Then Exit Function
    If InStr(txtStatus, "PAUSADO") > 0 Then Exit Function
    If InStr(txtStatus, "PARANDO") > 0 Then Exit Function
    If InStr(txtStatus, "INICIANDO") > 0 Then Exit Function
    If InStr(txtStatus, "AGUARDANDO") > 0 Then Exit Function

    PodeArquivarConcluidos = True
End Function

Private Sub RemoverAbaDashboardLegada()
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(LEGACY_DASHBOARD_SHEET).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub

Private Sub PrepararModoEdicaoInterno()
    RestaurarInterfaceExcel
    DesprotegerEstruturaOperacional
End Sub

Private Sub RegistrarAtalhosManutencao()
    On Error Resume Next
    Application.OnKey MAINTENANCE_SHORTCUT_OPEN, "'" & ThisWorkbook.Name & "'!EntrarModoManutencao"
    Application.OnKey MAINTENANCE_SHORTCUT_OPEN_ALT, "'" & ThisWorkbook.Name & "'!EntrarModoManutencao"
    Application.OnKey MAINTENANCE_SHORTCUT_CLOSE, "'" & ThisWorkbook.Name & "'!SairModoManutencao"
    Application.OnKey MAINTENANCE_SHORTCUT_CLOSE_ALT, "'" & ThisWorkbook.Name & "'!SairModoManutencao"
    Application.OnKey MAINTENANCE_SHORTCUT_CLOSE_BRACED, "'" & ThisWorkbook.Name & "'!SairModoManutencao"
    On Error GoTo 0
End Sub

Private Sub LimparAtalhosManutencao()
    On Error Resume Next
    Application.OnKey MAINTENANCE_SHORTCUT_OPEN
    Application.OnKey MAINTENANCE_SHORTCUT_OPEN_ALT
    Application.OnKey MAINTENANCE_SHORTCUT_CLOSE
    Application.OnKey MAINTENANCE_SHORTCUT_CLOSE_ALT
    Application.OnKey MAINTENANCE_SHORTCUT_CLOSE_BRACED
    On Error GoTo 0
End Sub

Private Function CredenciaisManutencaoValidas(ByVal loginInformado As String, ByVal senhaInformada As String) As Boolean
    CredenciaisManutencaoValidas = _
        (LCase$(Trim$(loginInformado)) = LoginManutencaoEsperado()) And _
        (senhaInformada = UI_PROTECTION_PASSWORD)
End Function

Private Function LoginManutencaoEsperado() As String
    LoginManutencaoEsperado = LCase$(Trim$(Environ$("Username")))
End Function

Private Sub AplicarModoOperacao()
    ProtegerPlanilhasOperacionais
    ProtegerEstruturaWorkbook
    OcultarInterfaceExcel
End Sub

Private Sub DesprotegerEstruturaOperacional()
    Dim ws As Worksheet

    On Error Resume Next
    ThisWorkbook.Unprotect Password:=UI_PROTECTION_PASSWORD
    On Error GoTo 0

    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        ws.Unprotect Password:=UI_PROTECTION_PASSWORD
        ws.EnableSelection = xlNoRestrictions
        On Error GoTo 0
    Next ws
End Sub

Private Sub ProtegerPlanilhasOperacionais()
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        ws.Unprotect Password:=UI_PROTECTION_PASSWORD
        ws.Protect Password:=UI_PROTECTION_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
        ws.EnableSelection = xlNoRestrictions
        On Error GoTo 0
    Next ws
End Sub

Private Sub ProtegerEstruturaWorkbook()
    On Error Resume Next
    ThisWorkbook.Unprotect Password:=UI_PROTECTION_PASSWORD
    ThisWorkbook.Protect Password:=UI_PROTECTION_PASSWORD, Structure:=True, Windows:=False
    On Error GoTo 0
End Sub

Private Sub OcultarInterfaceExcel()
    Dim wnd As Window

    On Error Resume Next
    If Not Application.ActiveWindow Is Nothing Then
        Application.ActiveWindow.DisplayHeadings = False
    End If
    For Each wnd In ThisWorkbook.Windows
        wnd.DisplayHeadings = False
    Next wnd
    On Error GoTo 0
End Sub

Private Sub RestaurarInterfaceExcel()
    Dim wnd As Window

    On Error Resume Next
    If Not Application.ActiveWindow Is Nothing Then
        Application.ActiveWindow.DisplayHeadings = True
    End If
    For Each wnd In ThisWorkbook.Windows
        wnd.DisplayHeadings = True
    Next wnd
    On Error GoTo 0
End Sub

Private Sub EnsureTopo(ByVal ws As Worksheet)
    AjustarControleLegadoExportacaoBase ws

    If Not PainelBaseIncompleto(ws) Then Exit Sub

    GarantirCardBase ws, SHP_BG, 4, 4, 1508, 34, "", RGB(14, 26, 44), RGB(230, 240, 255), ""
    GarantirCardBase ws, SHP_INFO_UPDATED, 10, 8, 220, 26, "Atualizado: -", RGB(28, 50, 84), RGB(209, 232, 255), ""
    GarantirCardBase ws, SHP_INFO_STATUS, 240, 8, 300, 26, "Status SAP: aguardando", RGB(28, 50, 84), RGB(205, 218, 235), ""
    GarantirCardBase ws, SHP_INFO_SUMMARY, 550, 8, 310, 26, "Total NF: - | Pend: - | OK: - | Erros: -", RGB(28, 50, 84), RGB(232, 239, 247), ""

    GarantirCardBase ws, SHP_BTN_START, 870, 8, 90, 26, ChrW(9654) & " Iniciar", RGB(33, 150, 243), RGB(255, 255, 255), "IniciarComControle"
    GarantirCardBase ws, SHP_BTN_PAUSE, 965, 8, 80, 26, ChrW(9646) & ChrW(9646) & " Pausar", RGB(255, 152, 0), RGB(255, 255, 255), "PausarProcessamento"
    GarantirCardBase ws, SHP_BTN_STOP, 1050, 8, 76, 26, ChrW(9632) & " Parar", RGB(220, 60, 60), RGB(255, 255, 255), "PararProcessamento"
    GarantirCardBase ws, SHP_BTN_REFRESH, 1130, 8, 92, 26, ChrW(8635) & " Atualizar", RGB(56, 84, 130), RGB(255, 255, 255), "AtualizarDashboard"
    GarantirCardBase ws, SHP_BTN_COCKPIT, 1230, 8, 176, 26, ChrW(9670) & " Cockpit", RGB(0, 118, 130), RGB(255, 255, 255), "BuscarNFsParaCockpit"
    GarantirCardBase ws, SHP_BTN_RELOAD, 1412, 8, 84, 26, ChrW(8635) & " Recarregar", RGB(100, 80, 180), RGB(255, 255, 255), "ForcarAtualizacaoExcel"
End Sub

Private Sub EnsureTopoConcluidos(ByVal ws As Worksheet)
    OcultarInterfaceExcel

    RemoverShapeSeExistir ws, SHP_BG
    RemoverShapeSeExistir ws, SHP_INFO_UPDATED
    RemoverShapeSeExistir ws, SHP_INFO_STATUS
    RemoverShapeSeExistir ws, SHP_INFO_SUMMARY
    RemoverShapeSeExistir ws, SHP_BTN_START
    RemoverShapeSeExistir ws, SHP_BTN_PAUSE
    RemoverShapeSeExistir ws, SHP_BTN_STOP
    RemoverShapeSeExistir ws, SHP_BTN_REFRESH
    RemoverShapeSeExistir ws, SHP_BTN_COCKPIT
    RemoverShapeSeExistir ws, SHP_BTN_RELOAD

    GarantirCardConcluidos ws, SHP_BTN_EXPORT_XLSX, 1352, 8, 84, 26, ChrW(8595) & " Excel", RGB(67, 160, 71), RGB(255, 255, 255), "ExportarExcel"
    GarantirCardConcluidos ws, SHP_BTN_EXPORT_PDF, 1442, 8, 54, 26, "PDF", RGB(56, 142, 60), RGB(255, 255, 255), "ExportarPDF"
    GarantirCardConcluidos ws, SHP_BTN_CLEAR_CONCLUIDOS, 1502, 8, 74, 26, "Limpar", RGB(211, 47, 47), RGB(255, 255, 255), "LimparListaConcluidos"
End Sub

Private Function PainelBaseIncompleto(ByVal ws As Worksheet) As Boolean
    PainelBaseIncompleto = _
        (Not ShapeExiste(ws, SHP_BG)) Or _
        (Not ShapeExiste(ws, SHP_INFO_UPDATED)) Or _
        (Not ShapeExiste(ws, SHP_INFO_STATUS)) Or _
        (Not ShapeExiste(ws, SHP_INFO_SUMMARY)) Or _
        (Not ShapeExiste(ws, SHP_BTN_START)) Or _
        (Not ShapeExiste(ws, SHP_BTN_PAUSE)) Or _
        (Not ShapeExiste(ws, SHP_BTN_STOP)) Or _
        (Not ShapeExiste(ws, SHP_BTN_REFRESH)) Or _
        (Not ShapeExiste(ws, SHP_BTN_COCKPIT)) Or _
        (Not ShapeExiste(ws, SHP_BTN_RELOAD))
End Function

Private Sub GarantirCardConcluidos(ByVal ws As Worksheet, ByVal nm As String, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, _
                                   ByVal txt As String, ByVal corBg As Long, ByVal corFg As Long, ByVal macro As String)
    Dim shp As Shape

    Set shp = GetShapeControle(ws, nm)
    If Not shp Is Nothing Then Exit Sub

    CriarCard ws, nm, x, y, w, h, txt, corBg, corFg, macro
End Sub

Private Sub GarantirCardBase(ByVal ws As Worksheet, ByVal nm As String, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, _
                             ByVal txt As String, ByVal corBg As Long, ByVal corFg As Long, ByVal macro As String)
    Dim shp As Shape

    Set shp = GetShapeControle(ws, nm)
    If Not shp Is Nothing Then Exit Sub

    CriarCard ws, nm, x, y, w, h, txt, corBg, corFg, macro
End Sub

Private Sub CriarCard(ByVal ws As Worksheet, ByVal nm As String, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, _
                      ByVal txt As String, ByVal corBg As Long, ByVal corFg As Long, ByVal macro As String)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    With shp
        .Placement = xlFreeFloating   ' nao move nem redimensiona com celulas
        .Name = nm
        .TextFrame.Characters.Text = txt
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 9
        .TextFrame.Characters.Font.Color = corFg
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .Fill.ForeColor.RGB = corBg
        .Line.Visible = msoFalse
        If Len(macro) > 0 Then
            .OnAction = "'" & ThisWorkbook.Name & "'!" & macro
        End If
    End With
End Sub

Private Sub SetCardText(ByVal ws As Worksheet, ByVal nm As String, ByVal txt As String, Optional ByVal corFg As Variant)
    Dim shp As Shape

    Set shp = GetShapeControle(ws, nm)
    If shp Is Nothing Then Exit Sub

    On Error Resume Next
    shp.TextFrame.Characters.Text = txt
    shp.TextFrame2.TextRange.Text = txt
    On Error GoTo 0
End Sub

Private Function GetShapeControle(ByVal ws As Worksheet, ByVal nm As String) As Shape
    Dim shp As Shape
    Dim itens As GroupShapes
    Dim item As Shape

    On Error Resume Next
    Set GetShapeControle = ws.Shapes(nm)
    On Error GoTo 0
    If Not GetShapeControle Is Nothing Then Exit Function

    For Each shp In ws.Shapes
        On Error Resume Next
        Set itens = shp.GroupItems
        If Err.Number <> 0 Then
            Err.Clear
            Set itens = Nothing
        End If
        On Error GoTo 0

        If Not itens Is Nothing Then
            For Each item In itens
                If StrComp(item.Name, nm, vbTextCompare) = 0 Then
                    Set GetShapeControle = item
                    Exit Function
                End If
            Next item
        End If
    Next shp
End Function

Private Function ShapeExiste(ByVal ws As Worksheet, ByVal nm As String) As Boolean
    ShapeExiste = Not GetShapeControle(ws, nm) Is Nothing
End Function

Private Sub RemoverShapeSeExistir(ByVal ws As Worksheet, ByVal nm As String)
    On Error Resume Next
    ws.Shapes(nm).Delete
    On Error GoTo 0
End Sub

Private Sub AjustarControleLegadoExportacaoBase(ByVal ws As Worksheet)
    Dim shpExport As Shape

    Set shpExport = GetShapeControle(ws, "ctl_btn_export")
    If shpExport Is Nothing Then Exit Sub

    On Error Resume Next
    If shpExport.Visible <> msoFalse Then shpExport.Visible = msoFalse
    If shpExport.Width <> 0 Then shpExport.Width = 0
    On Error GoTo 0
End Sub

Private Sub RemoverShapesControle(ByVal ws As Worksheet)
    Dim s As Shape
    For Each s In ws.Shapes
        If Left$(s.Name, 4) = "ctl_" Then s.Delete
    Next s
End Sub

Private Function StatusPorCorA(ByVal cel As Range) As String
    If cel.Interior.Pattern = xlPatternNone Or cel.Interior.ColorIndex = xlColorIndexNone Then
        StatusPorCorA = ""
        Exit Function
    End If

    Select Case cel.Interior.Color
        Case RGB(66, 165, 245)
            StatusPorCorA = "AUTCARR_OK"
        Case RGB(102, 187, 106)
            StatusPorCorA = "CONCLUIDO"
        Case RGB(239, 83, 80)
            StatusPorCorA = "ERRO"
        Case RGB(255, 152, 0)
            StatusPorCorA = "SEM_PAR"
        Case Else
            StatusPorCorA = "OUTRO"
    End Select
End Function

Private Sub ColetarIndicadores(ByVal ws As Worksheet, ByRef totalNf As Long, ByRef pend As Long, ByRef ok As Long, ByRef err As Long)
    Dim dTot As Object
    Dim dPen As Object
    Dim dOk As Object
    Dim dErr As Object

    Set dTot = CreateObject("Scripting.Dictionary")
    Set dPen = CreateObject("Scripting.Dictionary")
    Set dOk = CreateObject("Scripting.Dictionary")
    Set dErr = CreateObject("Scripting.Dictionary")

    ColetarIndicadoresEmDicionarios ws, dTot, dPen, dOk, dErr

    totalNf = dTot.Count
    pend = dPen.Count
    ok = dOk.Count
    err = dErr.Count
End Sub

Private Sub ColetarIndicadoresDia(ByVal ws As Worksheet, ByRef totalNf As Long, ByRef pend As Long, ByRef ok As Long, ByRef err As Long)
    Dim dTot As Object
    Dim dPen As Object
    Dim dOk As Object
    Dim dErr As Object
    Dim dConcluidasHoje As Object
    Dim chave As Variant

    Set dTot = CreateObject("Scripting.Dictionary")
    Set dPen = CreateObject("Scripting.Dictionary")
    Set dOk = CreateObject("Scripting.Dictionary")
    Set dErr = CreateObject("Scripting.Dictionary")

    ColetarIndicadoresEmDicionarios ws, dTot, dPen, dOk, dErr

    Set dConcluidasHoje = CarregarNFsConcluidasHoje()
    For Each chave In dConcluidasHoje.Keys
        If Not dTot.Exists(CStr(chave)) Then dTot.Add CStr(chave), 1
        If Not dOk.Exists(CStr(chave)) Then dOk.Add CStr(chave), 1
    Next chave

    totalNf = dTot.Count
    pend = dPen.Count
    ok = dOk.Count
    err = dErr.Count
End Sub

Private Sub ColetarIndicadoresEmDicionarios(ByVal ws As Worksheet, ByVal dTot As Object, ByVal dPen As Object, ByVal dOk As Object, ByVal dErr As Object)
    Dim ultima As Long
    Dim r As Long
    Dim nf As String
    Dim st As String

    ultima = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ultima < 2 Then Exit Sub

    For r = 3 To ultima
        nf = Trim$(CStr(ws.Cells(r, 1).Value))
        If Len(nf) = 0 Then GoTo Prox

        If Not dTot.Exists(nf) Then dTot.Add nf, 1
        st = StatusPorCorA(ws.Cells(r, 1))

        Select Case UCase$(st)
            Case "", "SEM_PAR"
                If Not dPen.Exists(nf) Then dPen.Add nf, 1
            Case "AUTCARR_OK", "CONCLUIDO"
                If Not dOk.Exists(nf) Then dOk.Add nf, 1
            Case Else
                If Not dErr.Exists(nf) Then dErr.Add nf, 1
        End Select
Prox:
    Next r

End Sub

Public Function TotalNotasLancadasDia(Optional ByVal wsBase As Worksheet = Nothing) As Long
    Dim dicHoje As Object
    Dim chave As Variant
    Dim ultima As Long
    Dim r As Long
    Dim nf As String

    Set dicHoje = CarregarNFsConcluidasHoje()

    If wsBase Is Nothing Then Set wsBase = GetBaseSheet()
    If Not wsBase Is Nothing Then
        ultima = wsBase.Cells(wsBase.Rows.Count, 1).End(xlUp).Row
        For r = 3 To ultima
            nf = Trim$(CStr(wsBase.Cells(r, 1).Value))
            If Len(nf) > 0 Then
                If UCase$(StatusPorCorA(wsBase.Cells(r, 1))) = "CONCLUIDO" Then
                    If Not dicHoje.Exists(nf) Then dicHoje.Add nf, 1
                End If
            End If
        Next r
    End If

    TotalNotasLancadasDia = dicHoje.Count
End Function

Private Sub RegistrarNFsConcluidasHoje(ByVal dicNovas As Object)
    Dim dicHoje As Object
    Dim chave As Variant
    Dim caminho As String
    Dim fn As Integer

    Set dicHoje = CarregarNFsConcluidasHoje()
    For Each chave In dicNovas.Keys
        If Not dicHoje.Exists(CStr(chave)) Then dicHoje.Add CStr(chave), 1
    Next chave

    caminho = CaminhoStatusDia()
    fn = FreeFile
    Open caminho For Output As #fn
    For Each chave In dicHoje.Keys
        Print #fn, CStr(chave)
    Next chave
    Close #fn
End Sub

Private Function CarregarNFsConcluidasHoje() As Object
    Dim dic As Object
    Dim caminho As String
    Dim fn As Integer
    Dim linha As String

    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = vbTextCompare

    caminho = CaminhoStatusDia()
    If Dir(caminho) = "" Then
        Set CarregarNFsConcluidasHoje = dic
        Exit Function
    End If

    fn = FreeFile
    Open caminho For Input As #fn
    Do While Not EOF(fn)
        Line Input #fn, linha
        linha = Trim$(linha)
        If Len(linha) > 0 Then
            If Not dic.Exists(linha) Then dic.Add linha, 1
        End If
    Loop
    Close #fn

    Set CarregarNFsConcluidasHoje = dic
End Function

Private Function CaminhoStatusDia() As String
    CaminhoStatusDia = GetPastaRPA() & "\" & STATUS_DIA_PREFIX & Format$(Date, "yyyymmdd") & ".txt"
End Function
