# -*- coding: utf-8 -*-
"""
Injeta VBA na base .xlsx via COM (win32com) e salva como .xlsm.
- Módulos padrão de vba_modules.txt
- frmConsulta (UserForm) recriado com controles + code-behind
- ThisWorkbook e Sheet(Menu) mínimos (não sobrescrevem fórmulas nativas)
- mod_Acoes (stubs Remarcar/Finalizar)
- Botões de navegação/ação com OnAction
"""
import re, os, sys, time
import win32com.client as win32

BUILD = r"D:\Vittor\Projetos\Saluti\PsiPlanner\build"
SRC_XLSX = BUILD + r"\PsiPlanner_base.xlsx"
FINAL = r"D:\Vittor\Projetos\Saluti\PsiPlanner\Psi Planner.xlsm"

# Constantes Excel/VBE
xlOpenXMLWorkbookMacroEnabled = 52
msoShapeRoundedRectangle = 5
vbext_ct_StdModule = 1
vbext_ct_MSForm = 3

# ---- parse dos módulos ----
with open(BUILD + r"\vba_modules.txt", encoding="utf-8") as f:
    raw = f.read()
parts = re.split(r'(?m)^==== MODULE: (.+?) ====\s*$', raw)
# parts: [preamble, name1, code1, name2, code2, ...]
modules = {}
for i in range(1, len(parts), 2):
    modules[parts[i].strip()] = parts[i + 1]

# mod_Consulta é substituído por versão sem UserForm (InputBox) — ver MOD_CONSULTA.
STD_MODULES = ["mod_Navegacao","mod_Calculos","mod_AlertaMEI","mod_Idade","mod_Avisos",
               "mod_Recibo","mod_Filtros","mod_Export","mod_Hover"]

# UserForms criados via COM Designer não recompilam de forma confiável (erro fm20
# "Forms.Form.1"). Substituímos o frmConsulta por um fluxo de InputBox sequencial.
MOD_CONSULTA = (
    "Option Explicit\r\n"
    "'------------------------------------------------------------------------------\r\n"
    "' mod_Consulta - cadastro rápido de consulta via InputBox (sem UserForm)\r\n"
    "'------------------------------------------------------------------------------\r\n\r\n"
    "' Abre o fluxo de nova consulta (paciente, data, hora, tipo, status).\r\n"
    "Sub AbrirFormConsulta()\r\n"
    "    On Error Resume Next\r\n"
    "    Dim pac As String, dataStr As String, hora As String, tipo As String, status As String\r\n"
    "    pac = InputBox(\"Paciente:\", \"Nova Consulta\")\r\n"
    "    If Len(Trim$(pac)) = 0 Then Exit Sub\r\n"
    "    dataStr = InputBox(\"Data (dd/mm/aaaa):\", \"Nova Consulta\", Format(Date, \"dd/mm/yyyy\"))\r\n"
    "    If Not IsDate(dataStr) Then MsgBox \"Data inválida.\", vbExclamation, \"Nova Consulta\": Exit Sub\r\n"
    "    hora = InputBox(\"Hora (hh:mm):\", \"Nova Consulta\", \"08:00\")\r\n"
    "    tipo = InputBox(\"Tipo (Presencial / Online):\", \"Nova Consulta\", \"Presencial\")\r\n"
    "    status = InputBox(\"Status (Confirmado / Remarcado / Faltou / Cancelado / Realizado):\", \"Nova Consulta\", \"Confirmado\")\r\n"
    "    RegistrarConsulta pac, dataStr, hora, tipo, status\r\n"
    "    MsgBox \"Consulta registrada no registro de agenda (aba Calendário).\", vbInformation, \"Nova Consulta\"\r\n"
    "End Sub\r\n\r\n"
    "' Acrescenta uma linha de agenda no Calendário (area de log a partir da linha 100).\r\n"
    "Public Sub RegistrarConsulta(ByVal paciente As String, ByVal dataStr As String, _\r\n"
    "                             ByVal hora As String, ByVal tipo As String, ByVal status As String)\r\n"
    "    On Error Resume Next\r\n"
    "    Dim ws As Worksheet, lin As Long\r\n"
    "    Set ws = Sheets(\"Calendário\")\r\n"
    "    If ws Is Nothing Then Exit Sub\r\n"
    "    lin = 100\r\n"
    "    Do While Not IsEmpty(ws.Cells(lin, 1).Value)\r\n"
    "        lin = lin + 1\r\n"
    "    Loop\r\n"
    "    ws.Cells(lin, 1).Value = IIf(IsDate(dataStr), CDate(dataStr), dataStr)\r\n"
    "    ws.Cells(lin, 1).NumberFormat = \"dd/mm/yyyy\"\r\n"
    "    ws.Cells(lin, 2).Value = hora\r\n"
    "    ws.Cells(lin, 3).Value = paciente\r\n"
    "    ws.Cells(lin, 4).Value = tipo\r\n"
    "    ws.Cells(lin, 5).Value = status\r\n"
    "End Sub\r\n"
)

THISWORKBOOK_CODE = (
    "Option Explicit\r\n\r\n"
    "' Ao abrir: recalcula tudo (formulas nativas), aplica privacidade, tema e vai ao Menu.\r\n"
    "Private Sub Workbook_Open()\r\n"
    "    On Error Resume Next\r\n"
    "    Application.CalculateFull\r\n"
    "    AplicarExibicaoConvenio\r\n"
    "    AplicarTema\r\n"
    "    Sheets(\"Menu Inicial\").Activate\r\n"
    "End Sub\r\n\r\n"
    "' Reage em tempo real aos toggles em Configurações (C18 convenio / C21 tema).\r\n"
    "Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)\r\n"
    "    On Error Resume Next\r\n"
    "    If Sh.Name = \"Configurações\" Then\r\n"
    "        If Not Intersect(Target, Sh.Range(\"C18\")) Is Nothing Then AplicarExibicaoConvenio\r\n"
    "        If Not Intersect(Target, Sh.Range(\"C21\")) Is Nothing Then AplicarTema\r\n"
    "    End If\r\n"
    "End Sub\r\n"
)

MENU_SHEET_CODE = (
    "Option Explicit\r\n\r\n"
    "' Ao ativar o Menu, recalcula KPIs/avisos (formulas nativas).\r\n"
    "Private Sub Worksheet_Activate()\r\n"
    "    On Error Resume Next\r\n"
    "    Me.Calculate\r\n"
    "End Sub\r\n"
)

MOD_CONVENIO = (
    "Option Explicit\r\n"
    "'------------------------------------------------------------------------------\r\n"
    "' mod_Convenio - exibe/oculta informacoes de convenio conforme toggle\r\n"
    "'                em Configurações!C18 (CFG_ExibirConvenio = Sim/Não)\r\n"
    "'------------------------------------------------------------------------------\r\n\r\n"
    "Public Sub AplicarExibicaoConvenio()\r\n"
    "    On Error Resume Next\r\n"
    "    Dim mostrar As Boolean, v As String\r\n"
    "    v = CStr(ThisWorkbook.Sheets(\"Configurações\").Range(\"C18\").Value)\r\n"
    "    mostrar = (LCase$(Trim$(v)) <> \"não\" And LCase$(Trim$(v)) <> \"nao\")\r\n"
    "    ' Pacientes: coluna J = Convênio\r\n"
    "    ThisWorkbook.Sheets(\"Pacientes\").Columns(\"J\").Hidden = Not mostrar\r\n"
    "    ' Relatórios: linha 6 colunas H:I (rótulo + valor Receita por convênio)\r\n"
    "    Dim rel As Worksheet\r\n"
    "    Set rel = ThisWorkbook.Sheets(\"Relatórios\")\r\n"
    "    rel.Range(\"H6:I6\").EntireRow.Hidden = False  ' nunca esconde a linha toda (afeta outros KPIs)\r\n"
    "    rel.Range(\"H6\").Value = IIf(mostrar, \"Receita por convênio\", \"\")\r\n"
    "    rel.Range(\"I6\").NumberFormat = IIf(mostrar, \"\"\"R$\"\" #,##0.00\", \";;;\")  ' \"R$\" entre aspas: evita duplicar o cifrao (RR$)\r\n"
    "    ' Gráfico Receita por convênio: 3º ChartObject da aba (ordem: receita / atend. / convênio)\r\n"
    "    If rel.ChartObjects.Count >= 3 Then\r\n"
    "        rel.ChartObjects(3).Visible = mostrar\r\n"
    "    End If\r\n"
    "End Sub\r\n"
)

MOD_TEMA = r'''Option Explicit
'------------------------------------------------------------------------------
' mod_Tema - Alternancia de tema Claro/Escuro ("Vinho Noturno").
' Reescreve APENAS formatacao (.Interior/.Font/.Borders/.Tab/graficos).
' NUNCA toca .Value/.Formula. Remap de cor bijetivo por contexto (fill/fonte/linha).
' Paleta dark (Material dark theme + WCAG AA): carvao quente (sem preto puro),
' texto off-white, acento vinho dessaturado, pasteis com luminosidade reduzida e
' texto claro legivel. Todos os pares texto/fundo >= 4.5:1.
'------------------------------------------------------------------------------

' HEX(6 sem #) -> Long de cor para .Color (RGB(r,g,b)).
Private Function HX(ByVal h As String) As Long
    HX = RGB(CLng("&H" & Mid$(h, 1, 2)), CLng("&H" & Mid$(h, 3, 2)), CLng("&H" & Mid$(h, 5, 2)))
End Function

' Abas tematizadas (nao toca "Recibo" gerado em runtime).
Private Function Abas() As Variant
    Abas = Array("Menu Inicial", "Financeiro", "Pacientes", "Calendário", _
                 "Aniversariantes", "Relatórios", "Configurações")
End Function

' Retangulo visivel/print por aba (cobre screenshots e graficos).
Private Function RectDaAba(ByVal nome As String) As String
    ' começa na linha 2: a linha 1 (faixa de cabeçalho vinho + logo dourado)
    ' é CONSTANTE nos dois temas, então fica fora do remap.
    Select Case nome
        Case "Menu Inicial":     RectDaAba = "A2:L24"
        Case "Financeiro":       RectDaAba = "A2:N118"
        Case "Pacientes":        RectDaAba = "A2:T58"
        Case "Calendário":       RectDaAba = "A2:K25"
        Case "Aniversariantes":  RectDaAba = "A2:Q24"
        Case "Relatórios":       RectDaAba = "A2:O34"
        Case "Configurações":    RectDaAba = "A2:D26"
        Case Else:               RectDaAba = "A2:A2"
    End Select
End Function

' Monta src()/dst() de um mapa de pares (claro,escuro) conforme a direcao.
Private Sub Montar(ByVal escuro As Boolean, ByVal claro As Variant, ByVal esc As Variant, _
                   ByRef src() As Long, ByRef dst() As Long)
    Dim i As Long, n As Long
    n = UBound(claro)
    ReDim src(0 To n)
    ReDim dst(0 To n)
    For i = 0 To n
        If escuro Then
            src(i) = HX(CStr(claro(i))): dst(i) = HX(CStr(esc(i)))
        Else
            src(i) = HX(CStr(esc(i))):  dst(i) = HX(CStr(claro(i)))
        End If
    Next i
End Sub

' Mapa de PREENCHIMENTO (Interior + fills de CF).
Private Sub MapaFill(ByVal escuro As Boolean, ByRef src() As Long, ByRef dst() As Long)
    Montar escuro, _
        Array("FAFAF8", "F4F4F4", "EAEAEA", "D9EAD3", "F4CCCC", "FFF2CC", "5A1E2A"), _
        Array("1A1719", "2A2528", "322B2E", "2E4A34", "5C2A31", "4D421F", "8E3B4C"), _
        src, dst
End Sub

' Mapa de FONTE (Font.Color + fontes de CF).
Private Sub MapaFonte(ByVal escuro As Boolean, ByRef src() As Long, ByRef dst() As Long)
    ' inclui acento DOURADO: claro=ouro_escuro 8A6A33 (AA s/ off-white) <-> escuro=champanhe C9A86A (AAA s/ carvao)
    Montar escuro, _
        Array("000000", "3B3B3B", "5A1E2A", "1F3B1F", "5A4A1E", "3A1018", "8A6A33"), _
        Array("E8E4E6", "A89FA4", "D08395", "CDE9D2", "F6E6B0", "F3CDD2", "C9A86A"), _
        src, dst
End Sub

' Mapa de LINHA/DIVISORIA (Borders).
Private Sub MapaLinha(ByVal escuro As Boolean, ByRef src() As Long, ByRef dst() As Long)
    ' filete DOURADO sob a faixa: claro=8A6A33 (bronze, visivel no claro) <-> escuro=C9A86A (champanhe)
    Montar escuro, _
        Array("EDEDED", "5A1E2A", "3B3B3B", "8A6A33"), _
        Array("3C353A", "D08395", "A89FA4", "C9A86A"), _
        src, dst
End Sub

' Substitui a cor se houver match; senao devolve a propria cor.
Private Function Remap(ByVal cor As Long, ByRef src() As Long, ByRef dst() As Long) As Long
    Dim i As Long
    Remap = cor
    For i = LBound(src) To UBound(src)
        If cor = src(i) Then Remap = dst(i): Exit Function
    Next i
End Function

' Remapeia a cor de uma aresta de borda (so escreve se mudou).
Private Sub RemapBorda(ByVal rg As Range, ByVal edge As Long, ByRef sl() As Long, ByRef dl() As Long)
    On Error Resume Next
    Dim nova As Long
    With rg.Borders(edge)
        If .LineStyle <> xlLineStyleNone Then
            nova = Remap(.Color, sl, dl)
            If nova <> .Color Then .Color = nova
        End If
    End With
End Sub

' Cor de aba por tema (Menu = vinho; demais = mute).
Private Function TabCor(ByVal nome As String, ByVal escuro As Boolean) As Long
    If nome = "Menu Inicial" Then
        If escuro Then TabCor = HX("8E3B4C") Else TabCor = HX("5A1E2A")
    Else
        If escuro Then TabCor = HX("6E656A") Else TabCor = HX("3B3B3B")
    End If
End Function

' Remapeia formatacao condicional (pasteis de status) por cor.
Private Sub RemapCF(ByVal ws As Worksheet, ByRef sf() As Long, ByRef df() As Long, _
                    ByRef st() As Long, ByRef dt() As Long)
    On Error Resume Next
    Dim fc As Variant, nova As Long
    For Each fc In ws.Cells.FormatConditions
        If fc.Interior.Pattern <> xlNone Then
            nova = Remap(fc.Interior.Color, sf, df)
            If nova <> fc.Interior.Color Then fc.Interior.Color = nova
        End If
        nova = Remap(fc.Font.Color, st, dt)
        If nova <> fc.Font.Color Then fc.Font.Color = nova
        If fc.Borders(xlEdgeLeft).LineStyle <> xlLineStyleNone Then
            nova = Remap(fc.Borders(xlEdgeLeft).Color, st, dt)
            If nova <> fc.Borders(xlEdgeLeft).Color Then fc.Borders(xlEdgeLeft).Color = nova
        End If
    Next fc
End Sub

' Tematiza graficos: area, texto, eixos e serie/fatias vinho (best-effort).
Private Sub RemapGraficos(ByVal ws As Worksheet, ByVal escuro As Boolean)
    On Error Resume Next
    Dim co As ChartObject, ch As Chart, s As Variant, j As Long
    Dim corSurf As Long, corTxt As Long, corVinho As Long
    If escuro Then
        corSurf = HX("1A1719"): corTxt = HX("E8E4E6"): corVinho = HX("8E3B4C")
    Else
        corSurf = HX("FFFFFF"): corTxt = HX("3B3B3B"): corVinho = HX("5A1E2A")
    End If
    For Each co In ws.ChartObjects
        Set ch = co.Chart
        ch.ChartArea.Format.Fill.ForeColor.RGB = corSurf
        ch.ChartArea.Format.Fill.Solid
        ch.ChartArea.Format.Line.Visible = msoFalse
        ch.PlotArea.Format.Fill.Visible = msoFalse
        ch.ChartArea.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = corTxt
        If ch.HasTitle Then ch.ChartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = corTxt
        ch.Axes(xlCategory).TickLabels.Font.Color = corTxt
        ch.Axes(xlValue).TickLabels.Font.Color = corTxt
        For Each s In ch.SeriesCollection
            If s.Format.Fill.ForeColor.RGB = HX("5A1E2A") Or s.Format.Fill.ForeColor.RGB = HX("8E3B4C") Then _
                s.Format.Fill.ForeColor.RGB = corVinho
            If s.Format.Line.ForeColor.RGB = HX("5A1E2A") Or s.Format.Line.ForeColor.RGB = HX("8E3B4C") Then _
                s.Format.Line.ForeColor.RGB = corVinho
            For j = 1 To s.Points.Count
                If s.Points(j).Format.Fill.ForeColor.RGB = HX("5A1E2A") Or _
                   s.Points(j).Format.Fill.ForeColor.RGB = HX("8E3B4C") Then _
                    s.Points(j).Format.Fill.ForeColor.RGB = corVinho
            Next j
        Next s
    Next co
End Sub

' === API publica ===

' Aplica o tema corrente (le CFG_Tema em Configurações). Idempotente.
Public Sub AplicarTema()
    On Error Resume Next
    Dim modo As String, escuro As Boolean
    modo = LCase$(Trim$(CStr(ThisWorkbook.Names("CFG_Tema").RefersToRange.Value)))
    escuro = (modo = "escuro")

    Dim sf() As Long, df() As Long, st() As Long, dt() As Long, sl() As Long, dl() As Long
    MapaFill escuro, sf, df
    MapaFonte escuro, st, dt
    MapaLinha escuro, sl, dl

    Dim corBase As Long
    If escuro Then corBase = HX("1A1719") Else corBase = HX("FAFAF8")

    Dim oldUpd As Boolean, oldEvt As Boolean
    oldUpd = Application.ScreenUpdating
    oldEvt = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim nome As Variant, ws As Worksheet, rng As Range, c As Range, nova As Long
    For Each nome In Abas()
        Set ws = Nothing
        Set ws = ThisWorkbook.Worksheets(CStr(nome))
        If Not ws Is Nothing Then
            Set rng = ws.Range(RectDaAba(CStr(nome)))
            For Each c In rng.Cells
                If c.Interior.ColorIndex = xlColorIndexNone Then
                    c.Interior.Color = corBase
                Else
                    nova = Remap(c.Interior.Color, sf, df)
                    If nova <> c.Interior.Color Then c.Interior.Color = nova
                End If
                nova = Remap(c.Font.Color, st, dt)
                If nova <> c.Font.Color Then c.Font.Color = nova
                RemapBorda c, xlEdgeBottom, sl, dl
                RemapBorda c, xlEdgeTop, sl, dl
                RemapBorda c, xlEdgeLeft, sl, dl
                RemapBorda c, xlEdgeRight, sl, dl
            Next c
            RemapCF ws, sf, df, st, dt
            ws.Tab.Color = TabCor(CStr(nome), escuro)
            RemapGraficos ws, escuro
        End If
    Next nome

    Application.EnableEvents = oldEvt
    Application.ScreenUpdating = oldUpd
End Sub

' Alterna Claro<->Escuro e repinta (atribuido ao botao "Alternar tema").
Public Sub AlternarTema()
    On Error Resume Next
    Dim r As Range, novo As String
    Set r = ThisWorkbook.Names("CFG_Tema").RefersToRange
    If r Is Nothing Then Exit Sub
    If LCase$(Trim$(CStr(r.Value))) = "escuro" Then novo = "Claro" Else novo = "Escuro"
    Application.EnableEvents = False
    r.Value = novo
    Application.EnableEvents = True
    AplicarTema
End Sub
'''.replace("\n", "\r\n")

MOD_ACOES = (
    "Option Explicit\r\n"
    "'------------------------------------------------------------------------------\r\n"
    "' mod_Acoes - acoes auxiliares da agenda (remarcar / finalizar)\r\n"
    "'------------------------------------------------------------------------------\r\n\r\n"
    "' Remarcar: abre o form de consulta para registrar nova data.\r\n"
    "Sub RemarcarConsulta()\r\n"
    "    On Error Resume Next\r\n"
    "    MsgBox \"Informe os dados da nova data/horario na sequencia.\", vbInformation, \"Remarcar\"\r\n"
    "    AbrirFormConsulta\r\n"
    "End Sub\r\n\r\n"
    "' Finalizar: marca a consulta selecionada como Realizado no registro de agenda.\r\n"
    "Sub FinalizarConsulta()\r\n"
    "    On Error Resume Next\r\n"
    "    Dim ws As Worksheet, r As Long\r\n"
    "    Set ws = Sheets(\"Calendário\")\r\n"
    "    If ws Is Nothing Then Exit Sub\r\n"
    "    r = ActiveCell.Row\r\n"
    "    If r >= 100 And ws.Cells(r, 3).Value <> \"\" Then\r\n"
    "        ws.Cells(r, 5).Value = \"Realizado\"\r\n"
    "        MsgBox \"Atendimento finalizado.\", vbInformation, \"Finalizar\"\r\n"
    "    Else\r\n"
    "        MsgBox \"Selecione uma linha no registro de agenda (a partir da linha 100).\", vbExclamation, \"Finalizar\"\r\n"
    "    End If\r\n"
    "End Sub\r\n\r\n"
    "' Busca rapida por paciente: localiza no cadastro e salta para a linha.\r\n"
    "Sub BuscarPaciente()\r\n"
    "    On Error Resume Next\r\n"
    "    Dim termo As String, rNome As Range, c As Range\r\n"
    "    termo = Trim$(InputBox(\"Nome do paciente (ou parte):\", \"Buscar paciente\"))\r\n"
    "    If Len(termo) = 0 Then Exit Sub\r\n"
    "    Set rNome = ThisWorkbook.Names(\"PAC_Nome\").RefersToRange\r\n"
    "    If rNome Is Nothing Then Exit Sub\r\n"
    "    For Each c In rNome.Cells\r\n"
    "        If Len(CStr(c.Value)) > 0 Then\r\n"
    "            If InStr(1, CStr(c.Value), termo, vbTextCompare) > 0 Then\r\n"
    "                c.Worksheet.Activate\r\n"
    "                c.Select\r\n"
    "                MsgBox \"Paciente encontrado: \" & c.Value, vbInformation, \"Buscar paciente\"\r\n"
    "                Exit Sub\r\n"
    "            End If\r\n"
    "        End If\r\n"
    "    Next c\r\n"
    "    MsgBox \"Nenhum paciente encontrado para '\" & termo & \"'.\", vbExclamation, \"Buscar paciente\"\r\n"
    "End Sub\r\n"
)

# controles do form: (progid, name, left, top, width, height, caption)
FORM_CONTROLS = [
    ("Forms.Label.1","lblPaciente",12,12,80,18,"Paciente:"),
    ("Forms.Label.1","lblData",12,42,80,18,"Data:"),
    ("Forms.Label.1","lblHora",12,72,80,18,"Hora:"),
    ("Forms.Label.1","lblTipo",12,102,80,18,"Tipo:"),
    ("Forms.Label.1","lblStatus",12,132,80,18,"Status:"),
    ("Forms.ComboBox.1","cboPaciente",100,10,170,18,None),
    ("Forms.TextBox.1","txtData",100,40,170,18,None),
    ("Forms.TextBox.1","txtHora",100,70,170,18,None),
    ("Forms.ComboBox.1","cboTipo",100,100,170,18,None),
    ("Forms.ComboBox.1","cboStatus",100,130,170,18,None),
    ("Forms.CommandButton.1","btnSalvar",100,170,80,26,"Salvar"),
    ("Forms.CommandButton.1","btnCancelar",190,170,80,26,"Cancelar"),
]

# botões: (sheet, label, sub, anchor, w_px, h_px, style, icon_hex)
# ícones Segoe MDL2 Assets (codepoints verificados)
BUTTONS = [
    ("Menu Inicial","Financeiro","IrFinanceiro","B11",150,56,"nav","E1D0"),
    ("Menu Inicial","Pacientes","IrPacientes","C11",150,56,"nav","E716"),
    ("Menu Inicial","Calendário","IrCalendario","D11",150,56,"nav","E787"),
    ("Menu Inicial","Aniversariantes","IrAniversariantes","B13",150,56,"nav","EC92"),
    ("Menu Inicial","Relatórios","IrRelatorios","C13",150,56,"nav","E9D9"),
    ("Menu Inicial","Configurações","IrConfiguracoes","D13",150,56,"nav","E713"),
    ("Financeiro","Menu","IrMenu","B1",110,24,"back","E80F"),
    ("Financeiro","Gerar Recibo","GerarRecibo","L13",150,30,"primary","E749"),
    ("Pacientes","Menu","IrMenu","B1",110,24,"back","E80F"),
    ("Calendário","Menu","IrMenu","B1",110,24,"back","E80F"),
    ("Calendário","Adicionar","AbrirFormConsulta","C23",130,30,"primary","E710"),
    ("Calendário","Remarcar","RemarcarConsulta","E23",130,30,"primary","E895"),
    ("Calendário","Finalizar","FinalizarConsulta","G23",130,30,"primary","E73E"),
    ("Calendário","Buscar","BuscarPaciente","I23",130,30,"primary","E721"),
    ("Aniversariantes","Menu","IrMenu","B1",110,24,"back","E80F"),
    ("Relatórios","Menu","IrMenu","B1",110,24,"back","E80F"),
    ("Relatórios","Exportar PDF","ExportarRelatorioPDF","H4",150,30,"primary","E749"),
    ("Configurações","Menu","IrMenu","B1",110,24,"back","E80F"),
    ("Configurações","Alternar tema","AlternarTema","C22",170,30,"primary","E706"),
]
STYLES = {
    "nav":     dict(fill=(90,30,42),  line=None, lw=0, txt=(255,255,255), font="Inter", size=11, bold=True),
    "primary": dict(fill=(58,16,24),  line=None, lw=0, txt=(255,255,255), font="Inter", size=11, bold=False),
    # back sobre a faixa vinho: pílula vinho-escuro + glifo dourado + rótulo off-white
    "back":    dict(fill=(58,16,24),  line=None, lw=0, txt=(232,228,230), font="Inter", size=10, bold=False),
}
PX = 0.75  # px -> pt

def main():
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        _build(excel)
    finally:
        try:
            excel.Quit()
        except Exception:
            pass

def _build(excel):
    wb = excel.Workbooks.Open(SRC_XLSX)
    vbp = wb.VBProject

    # 1) módulos padrão
    for name in STD_MODULES:
        comp = vbp.VBComponents.Add(vbext_ct_StdModule)
        comp.Name = name
        comp.CodeModule.AddFromString(modules[name])
    # mod_Acoes
    comp = vbp.VBComponents.Add(vbext_ct_StdModule)
    comp.Name = "mod_Acoes"
    comp.CodeModule.AddFromString(MOD_ACOES)
    # mod_Consulta (versão InputBox, sem UserForm)
    comp = vbp.VBComponents.Add(vbext_ct_StdModule)
    comp.Name = "mod_Consulta"
    comp.CodeModule.AddFromString(MOD_CONSULTA)
    # mod_Convenio (toggle de privacidade via Configurações!C18)
    comp = vbp.VBComponents.Add(vbext_ct_StdModule)
    comp.Name = "mod_Convenio"
    comp.CodeModule.AddFromString(MOD_CONVENIO)
    # mod_Tema (tema Claro/Escuro via Configurações!C21 = CFG_Tema)
    comp = vbp.VBComponents.Add(vbext_ct_StdModule)
    comp.Name = "mod_Tema"
    comp.CodeModule.AddFromString(MOD_TEMA)

    # 3) ThisWorkbook + Sheet(Menu) mínimos (usar CodeName, pode ser localizado)
    vbp.VBComponents(wb.CodeName).CodeModule.AddFromString(THISWORKBOOK_CODE)
    menu_cn = wb.Worksheets("Menu Inicial").CodeName
    vbp.VBComponents(menu_cn).CodeModule.AddFromString(MENU_SHEET_CODE)

    # 4) botões (com ícone monocromático Segoe MDL2 Assets)
    def rgb(t):
        return t[0] + t[1]*256 + t[2]*65536
    VINHO_RGB = rgb((90, 30, 42))
    WHITE_RGB = rgb((255, 255, 255))
    GOLD_RGB = rgb((201, 168, 106))  # ouro champanhe C9A86A (acento dos glifos)
    for sheet, label, sub, anchor, wpx, hpx, style, icon_hex in BUTTONS:
        ws = wb.Worksheets(sheet)
        rng = ws.Range(anchor)
        left, top = rng.Left, rng.Top
        shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, left, top, wpx*PX, hpx*PX)
        st = STYLES[style]
        shp.Fill.ForeColor.RGB = rgb(st["fill"])
        if st["line"] is None:
            shp.Line.Visible = False
        else:
            shp.Line.ForeColor.RGB = rgb(st["line"])
            shp.Line.Weight = st["lw"]
        glyph = chr(int(icon_hex, 16))
        tf = shp.TextFrame  # API legada: confiável p/ texto + cor por caractere
        tf.Characters().Text = glyph + "  " + label   # ícone + rótulo, uma linha
        tf.Characters().Font.Name = st["font"]
        tf.Characters().Font.Size = st["size"]
        tf.Characters().Font.Color = rgb(st["txt"])
        tf.Characters().Font.Bold = bool(st.get("bold", False))
        tf.HorizontalAlignment = -4108  # xlHAlignCenter
        try:
            shp.TextFrame2.WordWrap = 0       # msoFalse: nunca quebra/oculta o rótulo
            shp.TextFrame2.AutoSize = 0       # msoAutoSizeNone
        except Exception:
            pass
        # glifo (1º caractere): fonte de ícone + cor de destaque
        ico = tf.Characters(1, 1)
        ico.Font.Name = "Segoe MDL2 Assets"
        ico.Font.Size = st["size"] + 3
        # glifo SEMPRE em dourado (acento da identidade), rótulo permanece claro
        ico.Font.Color = GOLD_RGB
        shp.TextFrame2.VerticalAnchor = 3  # middle
        shp.OnAction = sub
        shp.Name = "btn_" + sub + "_" + sheet.replace(" ", "")[:6]

    # gráficos: plotar dados mesmo em colunas/linhas ocultas (aux escondido)
    for i in range(1, wb.Worksheets.Count + 1):
        ws = wb.Worksheets(i)
        for co in ws.ChartObjects():
            try:
                co.Chart.PlotVisibleOnly = False
            except Exception as e:
                print("chart warn:", e)

    # NÃO definir áreas de impressão (_xlnm.Print_Area). Em .xlsm, a mera
    # presença desse nome faz o Excel paginar contra o driver de impressora ao
    # ABRIR, o que pode travar a abertura (confirmado neste ambiente, mesmo com
    # impressora padrão e macros desabilitadas). Orientação/fit-to-width vêm do
    # openpyxl (base) e sobrevivem ao SaveAs; a impressão usa o range visível
    # (colunas/linhas ocultas não são impressas), que é o conteúdo desejado.

    # propriedades do documento (identidade Psi Planner) — autoritativo no .xlsm
    def _setprop(name, value):
        try:
            wb.BuiltinDocumentProperties(name).Value = value
        except Exception as e:
            print("prop warn (%s): %r" % (name, e))
    _setprop("Title", "Psi Planner")
    _setprop("Subject", "Planejador financeiro-clínico para psicólogos e terapeutas")
    _setprop("Author", "Psi Planner")
    _setprop("Company", "Psi Planner")
    _setprop("Keywords", "psicologia, financeiro, MEI, agenda, clínica")

    # recalc e salvar como .xlsm
    excel.Calculate()
    wb.Worksheets("Menu Inicial").Activate()
    if os.path.exists(FINAL):
        os.remove(FINAL)
    wb.SaveAs(FINAL, FileFormat=xlOpenXMLWorkbookMacroEnabled)

    # ---- validação ----
    print("=== VALIDAÇÃO ===")
    print("Abas:", [wb.Worksheets(i+1).Name for i in range(wb.Worksheets.Count)])
    excel.Run("IrFinanceiro")
    print("Após IrFinanceiro, ativa =", excel.ActiveSheet.Name)
    excel.Run("IrMenu")
    print("Após IrMenu, ativa =", excel.ActiveSheet.Name)
    menu = wb.Worksheets("Menu Inicial")
    print("KPI PacientesAtivos (B7) =", menu.Range("B7").Value)
    print("KPI ReceitaMes   (C7) =", menu.Range("C7").Value)
    print("KPI ValorPendente(D7) =", menu.Range("D7").Value)
    print("KPI Aniversariantes(H7)=", menu.Range("H7").Value)
    fin = wb.Worksheets("Financeiro")
    print("FIN TotalRecebido(M6) =", fin.Range("M6").Value)
    print("FIN MEI_Faturamento(M11)=", fin.Range("M11").Value)
    print("FIN MEI_Indicador(M10)=", fin.Range("M10").Value)
    print("Nº de Shapes Menu =", menu.Shapes.Count)
    print("Nº componentes VBA =", vbp.VBComponents.Count)

    wb.Close(SaveChanges=False)
    print("OK ->", FINAL)

if __name__ == "__main__":
    main()
