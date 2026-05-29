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

BUILD = r"D:\Vittor\Projetos\Saluti\PsiPlanner\build2"
SRC_XLSX = BUILD + r"\PsiPlanner2_base.xlsx"
FINAL = r"D:\Vittor\Projetos\Saluti\PsiPlanner\modelo2\Psi Planner 2.xlsm"

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
    "' Ao abrir: recalcula tudo (formulas nativas) e vai ao Menu Inicial.\r\n"
    "Private Sub Workbook_Open()\r\n"
    "    On Error Resume Next\r\n"
    "    Application.CalculateFull\r\n"
    "    Sheets(\"Menu Inicial\").Activate\r\n"
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
]
STYLES = {
    "nav":     dict(fill=(255,255,255), line=(59,59,59),   lw=0.75, txt=(0,0,0),     font="Poppins Light", size=11),
    "primary": dict(fill=(58,16,24),    line=None,         lw=0,    txt=(255,255,255), font="Poppins Light", size=11),
    "back":    dict(fill=(244,244,244), line=(166,166,166),lw=0.75, txt=(0,0,0),     font="Poppins Light",      size=10),
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

    # 3) ThisWorkbook + Sheet(Menu) mínimos (usar CodeName, pode ser localizado)
    vbp.VBComponents(wb.CodeName).CodeModule.AddFromString(THISWORKBOOK_CODE)
    menu_cn = wb.Worksheets("Menu Inicial").CodeName
    vbp.VBComponents(menu_cn).CodeModule.AddFromString(MENU_SHEET_CODE)

    # 4) botões (com ícone monocromático Segoe MDL2 Assets)
    def rgb(t):
        return t[0] + t[1]*256 + t[2]*65536
    VINHO_RGB = rgb((90, 30, 42))
    WHITE_RGB = rgb((255, 255, 255))
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
        ico.Font.Color = (VINHO_RGB if style in ("nav", "back") else WHITE_RGB)
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
    _setprop("Title", "Psi Planner 2")
    _setprop("Subject", "Planejador financeiro-clínico para psicólogos e terapeutas")
    _setprop("Author", "Psi Planner 2")
    _setprop("Company", "Psi Planner 2")
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
