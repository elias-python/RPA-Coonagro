Option Explicit

' ======================================================
' modDashboard (modo sem aba Dashboard)
' Mantem controles e informacoes no topo da aba Base.
' ======================================================

Private Const BASE_SHEET_NAME As String = "Base"
Private Const LEGACY_DASHBOARD_SHEET As String = "Dashboard"

Private Const SHP_BG As String = "ctl_bg"
Private Const SHP_INFO_UPDATED As String = "ctl_info_updated"
Private Const SHP_INFO_STATUS As String = "ctl_info_status"
Private Const SHP_INFO_SUMMARY As String = "ctl_info_summary"
Private Const SHP_BTN_START As String = "ctl_btn_start"
Private Const SHP_BTN_PAUSE As String = "ctl_btn_pause"
Private Const SHP_BTN_STOP As String = "ctl_btn_stop"
Private Const SHP_BTN_REFRESH As String = "ctl_btn_refresh"
Private Const SHP_BTN_COCKPIT As String = "ctl_btn_cockpit"
Private Const SHP_BTN_EXPORT   As String = "ctl_btn_export"
Private Const SHP_BTN_RELOAD   As String = "ctl_btn_reload"

' Executado automaticamente ao abrir o arquivo.
' Reconstroi o painel caso os shapes tenham sido removidos pela atualizacao Python (openpyxl).
Public Sub Auto_Open()
    Application.ScreenUpdating = False
    RemoverAbaDashboardLegada
    Dim ws As Worksheet
    Set ws = GetBaseSheet()
    If Not ws Is Nothing Then
        If Not ShapeExiste(ws, SHP_BG) Then
            RemoverShapesControle ws
            EnsureTopo ws
        End If
        ws.Activate
    End If
    AtualizarDashboard
    Application.ScreenUpdating = True
End Sub

Public Sub InicializarDashboard()
    Dim ws As Worksheet
    Set ws = GetBaseSheet()
    If ws Is Nothing Then
        MsgBox "Aba 'Base' nao encontrada.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    RemoverAbaDashboardLegada
    RemoverShapesControle ws   ' Forca recriacao completa ao inicializar
    EnsureTopo ws
    AtualizarDashboard
    ws.Activate
    Application.ScreenUpdating = True
End Sub

Public Sub AtualizarDashboard()
    Dim ws As Worksheet
    Set ws = GetBaseSheet()
    If ws Is Nothing Then Exit Sub

    EnsureTopo ws

    Dim totalNf As Long, pend As Long, ok As Long, err As Long
    ColetarIndicadores ws, totalNf, pend, ok, err

    SetCardText ws, SHP_INFO_UPDATED, "Atualizado: " & Format$(Now, "dd/mm/yyyy hh:mm:ss"), RGB(209, 232, 255)
    SetCardText ws, SHP_INFO_SUMMARY, "Total NF: " & totalNf & " | Pend: " & pend & " | OK: " & ok & " | Erros: " & err, RGB(232, 239, 247)
End Sub

Public Sub AtualizarStatus(ByVal texto As String)
    Dim ws As Worksheet
    Set ws = GetBaseSheet()
    If ws Is Nothing Then Exit Sub

    EnsureTopo ws

    Dim cor As Long
    Dim u As String
    u = UCase$(texto)

    If InStr(u, "ERRO") > 0 Or InStr(u, "FATAL") > 0 Then
        cor = RGB(255, 110, 110)
    ElseIf InStr(u, "RODANDO") > 0 Or InStr(u, "PROCESSANDO") > 0 Then
        cor = RGB(120, 235, 140)
    ElseIf InStr(u, "FINALIZADO") > 0 Then
        cor = RGB(130, 200, 255)
    ElseIf InStr(u, "PAUSADO") > 0 Then
        cor = RGB(255, 220, 110)
    ElseIf InStr(u, "PARAND") > 0 Then
        cor = RGB(255, 180, 100)
    Else
        cor = RGB(205, 218, 235)
    End If

    SetCardText ws, SHP_INFO_STATUS, "Status SAP: " & texto, cor
    DoEvents
End Sub

Public Sub AtualizarContadores(ByVal erros As Long, ByVal concluidos As Long)
    Dim ws As Worksheet
    Set ws = GetBaseSheet()
    If ws Is Nothing Then Exit Sub

    EnsureTopo ws

    Dim totalNf As Long, pend As Long, ok As Long, err As Long
    ColetarIndicadores ws, totalNf, pend, ok, err

    SetCardText ws, SHP_INFO_SUMMARY, _
        "Total NF: " & totalNf & " | Pend: " & pend & " | OK: " & ok & " | Erros cor: " & err & " | Proc: " & concluidos & " | Erros SAP: " & erros, _
        RGB(232, 239, 247)
    DoEvents
End Sub

Private Function GetBaseSheet() As Worksheet
    On Error Resume Next
    Set GetBaseSheet = ThisWorkbook.Worksheets(BASE_SHEET_NAME)
    On Error GoTo 0
End Function

Private Sub RemoverAbaDashboardLegada()
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(LEGACY_DASHBOARD_SHEET).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub

Private Sub EnsureTopo(ByVal ws As Worksheet)
    If ShapeExiste(ws, SHP_BG) Then Exit Sub

    RemoverShapesControle ws

    CriarCard ws, SHP_BG, 4, 4, 1576, 34, "", RGB(14, 26, 44), RGB(230, 240, 255), ""
    CriarCard ws, SHP_INFO_UPDATED, 10, 8, 220, 26, "Atualizado: -", RGB(28, 50, 84), RGB(209, 232, 255), ""
    CriarCard ws, SHP_INFO_STATUS, 240, 8, 300, 26, "Status SAP: aguardando", RGB(28, 50, 84), RGB(205, 218, 235), ""
    CriarCard ws, SHP_INFO_SUMMARY, 550, 8, 310, 26, "Total NF: - | Pend: - | OK: - | Erros: -", RGB(28, 50, 84), RGB(232, 239, 247), ""

    CriarCard ws, SHP_BTN_START, 870, 8, 90, 26, ChrW(9654) & " Iniciar", RGB(33, 150, 243), RGB(255, 255, 255), "IniciarComControle"
    CriarCard ws, SHP_BTN_PAUSE, 965, 8, 80, 26, ChrW(9646) & ChrW(9646) & " Pausar", RGB(255, 152, 0), RGB(255, 255, 255), "PausarProcessamento"
    CriarCard ws, SHP_BTN_STOP, 1050, 8, 76, 26, ChrW(9632) & " Parar", RGB(220, 60, 60), RGB(255, 255, 255), "PararProcessamento"
    CriarCard ws, SHP_BTN_REFRESH, 1130, 8, 92, 26, ChrW(8635) & " Atualizar", RGB(56, 84, 130), RGB(255, 255, 255), "AtualizarDashboard"
    CriarCard ws, SHP_BTN_COCKPIT, 1230, 8, 136, 26, ChrW(9670) & " Cockpit", RGB(0, 118, 130), RGB(255, 255, 255), "BuscarNFsParaCockpit"
    CriarCard ws, SHP_BTN_EXPORT,  1372, 8,  96, 26, ChrW(8595) & " Exportar", RGB(67, 160, 71), RGB(255, 255, 255), "ExportarDadosLancados"
    CriarCard ws, SHP_BTN_RELOAD,  1474, 8, 100, 26, ChrW(8635) & " Recarregar", RGB(100, 80, 180), RGB(255, 255, 255), "ForcarAtualizacaoExcel"
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
    On Error Resume Next
    ws.Shapes(nm).TextFrame.Characters.Text = txt
    If Not IsMissing(corFg) Then
        ws.Shapes(nm).TextFrame.Characters.Font.Color = CLng(corFg)
    End If
    On Error GoTo 0
End Sub

Private Function ShapeExiste(ByVal ws As Worksheet, ByVal nm As String) As Boolean
    On Error Resume Next
    ShapeExiste = Not ws.Shapes(nm) Is Nothing
    On Error GoTo 0
End Function

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
    Dim ultima As Long
    Dim r As Long
    Dim nf As String
    Dim st As String
    Dim dTot As Object, dPen As Object, dOk As Object, dErr As Object

    Set dTot = CreateObject("Scripting.Dictionary")
    Set dPen = CreateObject("Scripting.Dictionary")
    Set dOk = CreateObject("Scripting.Dictionary")
    Set dErr = CreateObject("Scripting.Dictionary")

    ultima = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ultima < 2 Then Exit Sub

    For r = 3 To ultima
        nf = Trim$(CStr(ws.Cells(r, 1).Value))
        If Len(nf) = 0 Then GoTo Prox

        If Not dTot.Exists(nf) Then dTot.Add nf, 1
        st = StatusPorCorA(ws.Cells(r, 1))

        Select Case UCase$(st)
            Case ""
                If Not dPen.Exists(nf) Then dPen.Add nf, 1
            Case "AUTCARR_OK", "CONCLUIDO"
                If Not dOk.Exists(nf) Then dOk.Add nf, 1
            Case Else
                If Not dErr.Exists(nf) Then dErr.Add nf, 1
        End Select
Prox:
    Next r

    totalNf = dTot.Count
    pend = dPen.Count
    ok = dOk.Count
    err = dErr.Count
End Sub
