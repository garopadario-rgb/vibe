from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter

wb = Workbook()
ws = wb.active
ws.title = "Tutoriais IA Websites"

# ── Cores ──────────────────────────────────────────────────────────────────
DARK_BG    = "1A1A2E"   # cabeçalho principal
CLAUDE_BG  = "16213E"   # linhas Claude
CLAUDE_ACC = "0F3460"   # acento Claude
GENERAL_BG = "1B2A4A"   # linhas gerais
GENERAL_ACC= "E94560"   # acento geral
HEADER_BG  = "0F3460"   # cabeçalho das colunas
WHITE      = "FFFFFF"
GOLD       = "FFD700"
LIGHT_GRAY = "F0F4FF"
CLAUDE_ROW = "D6E4FF"   # fundo linhas Claude (claro)
GENERAL_ROW= "FFE4EC"   # fundo linhas gerais (claro)
ALT_CLAUDE = "EBF3FF"
ALT_GENERAL= "FFF0F5"

# ── Bordas ─────────────────────────────────────────────────────────────────
thin = Side(style="thin", color="CCCCCC")
medium = Side(style="medium", color="888888")
border_thin = Border(left=thin, right=thin, top=thin, bottom=thin)
border_medium = Border(left=medium, right=medium, top=medium, bottom=medium)

# ── Título principal ───────────────────────────────────────────────────────
ws.merge_cells("A1:J1")
title_cell = ws["A1"]
title_cell.value = "🎓 20 MELHORES TUTORIAIS: CONSTRUIR WEBSITES COM IA"
title_cell.font = Font(name="Calibri", size=16, bold=True, color=WHITE)
title_cell.fill = PatternFill("solid", fgColor=DARK_BG)
title_cell.alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[1].height = 38

# ── Subtítulo ──────────────────────────────────────────────────────────────
ws.merge_cells("A2:J2")
sub = ws["A2"]
sub.value = "10 vídeos com Claude AI  •  10 vídeos com IA Geral  |  Tutoriais completos passo a passo • Fonte de renda extra"
sub.font = Font(name="Calibri", size=10, italic=True, color="AAAACC")
sub.fill = PatternFill("solid", fgColor=DARK_BG)
sub.alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[2].height = 20

# ── Linha vazia separadora ─────────────────────────────────────────────────
ws.row_dimensions[3].height = 6

# ── Cabeçalhos das colunas ─────────────────────────────────────────────────
headers = ["#", "Título do Vídeo", "Link YouTube", "Categoria", "IA Utilizada",
           "Canal / Criador", "Língua", "Data", "Visualizações", "Notas"]

for col, h in enumerate(headers, 1):
    cell = ws.cell(row=4, column=col, value=h)
    cell.font = Font(name="Calibri", size=10, bold=True, color=WHITE)
    cell.fill = PatternFill("solid", fgColor=HEADER_BG)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = border_medium
ws.row_dimensions[4].height = 24

# ── Dados ──────────────────────────────────────────────────────────────────
data = [
    # Claude videos (1–12)
    (1,  "How To Use Claude AI To Build A Website 2025",
     "https://www.youtube.com/watch?v=GY8opNDcJY4",
     "Claude AI", "Claude AI", "Verificar no YT", "Inglês", "Jun 2025",
     "Verificar no YT", "Tutorial completo passo a passo para iniciantes"),

    (2,  "How to Build a Website Using Claude AI (Updated 2025)",
     "https://www.youtube.com/watch?v=kkfn3ZiKnNg",
     "Claude AI", "Claude AI", "Verificar no YT", "Inglês", "Nov 2025",
     "Verificar no YT", "Versão actualizada — mostra todo o processo"),

    (3,  "Claude Code for Beginners — Build Your First App with AI (No Coding Required!)",
     "https://www.youtube.com/watch?v=s-Mc26Ytz10",
     "Claude AI", "Claude Code", "Verificar no YT", "Inglês", "Fev 2026",
     "Verificar no YT", "App web completa sem escrever uma linha de código"),

    (4,  "How To Use Claude AI To Build A Website: No-Code Site Tutorial",
     "https://www.youtube.com/watch?v=lLp-jC_ezA4",
     "Claude AI", "Claude AI", "Verificar no YT", "Inglês", "Mar 2026",
     "Verificar no YT", "Abordagem no-code — resultado profissional garantido"),

    (5,  "I Built an Entire Website with Claude Code… Here's How",
     "https://www.youtube.com/watch?v=SKBDC3QugZw",
     "Claude AI", "Claude Code", "Verificar no YT", "Inglês", "Ago 2025",
     "Verificar no YT", "Landing page completa do zero com Claude Code"),

    (6,  "Complete guide to building out a website with Claude for Beginners",
     "https://www.youtube.com/watch?v=YVzNA1XlVbQ",
     "Claude AI", "Claude AI", "Verificar no YT", "Inglês", "Jan 2026",
     "Verificar no YT", "Guia completo para iniciantes — nada oculto"),

    (7,  "How To Build A Website With Claude: (2026) Easy Tutorial",
     "https://www.youtube.com/watch?v=wUD9lWKVMu8",
     "Claude AI", "Claude AI", "Verificar no YT", "Inglês", "Mar 2026",
     "Verificar no YT", "Tutorial 2026 detalhado e fácil de seguir"),

    (8,  "Build & Host FREE Websites in 5 Minutes (Claude Code Full Tutorial!)",
     "https://www.youtube.com/watch?v=HSgxNuGBfPk",
     "Claude AI", "Claude Code", "Brendan's AI Launchpad", "Inglês", "Mai 2026",
     "Verificar no YT", "Construção + hospedagem gratuita com Claude Code"),

    (9,  "Build Gorgeous Websites With Claude Code (FULL COURSE)",
     "https://www.youtube.com/watch?v=MCr0ej5c-QM",
     "Claude AI", "Claude Code", "Verificar no YT", "Inglês", "Abr 2026",
     "Verificar no YT", "Curso completo — design profissional com Claude Code"),

    (10, "Build $10,000 Websites using Claude Code (Ultimate Guide)",
     "https://www.youtube.com/watch?v=VMvZuhcDdnw",
     "Claude AI", "Claude Code", "Verificar no YT", "Inglês", "Mai 2026",
     "Verificar no YT", "Guia definitivo para websites de valor elevado"),

    (11, "Claude Code + 3D Animations: Build Client Ready Websites in Minutes",
     "https://www.youtube.com/watch?v=AGHqBAVyrQs",
     "Claude AI", "Claude Code", "Kirk M Design", "Inglês", "Jun 2026",
     "Verificar no YT", "Websites com animações 3D — pronto para clientes"),

    (12, "Claude Design Tutorial | Build Sites Faster Than Ever",
     "https://www.youtube.com/watch?v=6OYgqv1Wxkg",
     "Claude AI", "Claude Design", "Designing for Uncertainty", "Inglês", "Jun 2026",
     "Verificar no YT", "Claude Design para designers e agências — produtividade máxima"),

    # General AI videos (13–20)
    (13, "How To Build A $10,000 Website In 30 Minutes (AI + No Code)",
     "https://www.youtube.com/watch?v=CGEjwCUcMHo",
     "IA Geral", "ChatGPT + 7 IAs", "Wes McDowell", "Inglês", "Nov 2024",
     "500.000 +", "7 ferramentas IA gratuitas — website de 10 mil dólares em 30 min"),

    (14, "Build a Full Stack Website with ChatGPT + AI | No Code | Complete Step-by-Step",
     "https://www.youtube.com/watch?v=j3D1IOpqGwk",
     "IA Geral", "ChatGPT", "Verificar no YT", "Inglês", "Set 2025",
     "Verificar no YT", "Website full-stack completo sem código com ChatGPT"),

    (15, "How to Build a Website With ChatGPT In 2024",
     "https://www.youtube.com/watch?v=Kry6548MBFY",
     "IA Geral", "ChatGPT", "Verificar no YT", "Inglês", "Jan 2024",
     "Verificar no YT", "Guia directo passo a passo com ChatGPT"),

    (16, "Build Your First AI Chatbot | Complete Guide 2024",
     "https://www.youtube.com/watch?v=YzRwq3wSu6w",
     "IA Geral", "ChatGPT API", "Verificar no YT", "Inglês", "Mar 2024",
     "Verificar no YT", "Chatbot interactivo completo — tutorial sem segredos"),

    (17, "Como criar um SITE com INTELIGÊNCIA ARTIFICIAL | Guia Completo",
     "https://www.youtube.com/watch?v=pY6AncNPekk",
     "IA Geral", "IA múltipla", "Verificar no YT", "Português (BR)", "Jul 2024",
     "Verificar no YT", "Guia completo em Português — processo de A a Z"),

    (18, "COMO CRIAR UM SITE COMPLETO COM INTELIGÊNCIA ARTIFICIAL | Passo a Passo - 2024",
     "https://www.youtube.com/watch?v=FciuDpPxKnQ",
     "IA Geral", "IA múltipla", "Verificar no YT", "Português (BR)", "Mar 2024",
     "Verificar no YT", "Tutorial PT completo passo a passo 2024"),

    (19, "Aprenda a CRIAR SITE COM IA (Inteligência Artificial) em apenas 15 minutos",
     "https://www.youtube.com/watch?v=51YnYmdxxHk",
     "IA Geral", "Hostinger AI + ChatGPT", "Leonardo (Hostinger)", "Português (BR)", "Mai 2023",
     "Verificar no YT", "Site profissional em 15 min — excelente para começar a monetizar"),

    (20, "Tutorial FRAMER: Como CRIAR um SITE com IA?",
     "https://www.youtube.com/watch?v=IXCejzBC21Q",
     "IA Geral", "Framer AI", "Verificar no YT", "Português (BR)", "Dez 2023",
     "Verificar no YT", "Framer + IA em Português — processo completo revelado"),
]

# ── Inserir dados ──────────────────────────────────────────────────────────
for i, row in enumerate(data):
    excel_row = i + 5
    is_claude = row[3] == "Claude AI"
    is_alt = i % 2 == 1

    if is_claude:
        bg = ALT_CLAUDE if is_alt else CLAUDE_ROW
        cat_color = "0F3460"
    else:
        bg = ALT_GENERAL if is_alt else GENERAL_ROW
        cat_color = "C0392B"

    fill = PatternFill("solid", fgColor=bg)

    for col, val in enumerate(row, 1):
        cell = ws.cell(row=excel_row, column=col, value=val)
        cell.fill = fill
        cell.border = border_thin
        cell.alignment = Alignment(vertical="center", wrap_text=True,
                                   horizontal="center" if col in (1, 4, 6, 7, 8, 9) else "left")

        # Número — destaque
        if col == 1:
            cell.font = Font(name="Calibri", size=11, bold=True, color=DARK_BG)
        # Título
        elif col == 2:
            cell.font = Font(name="Calibri", size=9, bold=True, color="1A1A2E")
        # Link — azul clicável
        elif col == 3:
            cell.font = Font(name="Calibri", size=8, color="1155CC", underline="single")
            cell.hyperlink = val
            cell.value = val
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        # Categoria — cor por tipo
        elif col == 4:
            cell.font = Font(name="Calibri", size=9, bold=True, color=cat_color)
        else:
            cell.font = Font(name="Calibri", size=9, color="2C2C2C")

    ws.row_dimensions[excel_row].height = 36

# ── Secção separadora Claude / Geral ──────────────────────────────────────
# Inserir linha de grupo antes da linha 13 (índice 17 no excel = linha 5+12)
# Já estão ordenados, mas vamos adicionar cabeçalhos de grupo
def insert_group_header(row_num, text, color):
    ws.merge_cells(f"A{row_num}:J{row_num}")
    c = ws[f"A{row_num}"]
    c.value = text
    c.font = Font(name="Calibri", size=10, bold=True, color=WHITE)
    c.fill = PatternFill("solid", fgColor=color)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row_num].height = 20

# Não há linhas de grupo inseridas (dados contíguos), mas colorimos as células de categoria

# ── Larguras das colunas ───────────────────────────────────────────────────
col_widths = {
    1: 4,    # #
    2: 45,   # título
    3: 16,   # link
    4: 12,   # categoria
    5: 16,   # IA
    6: 22,   # canal
    7: 12,   # língua
    8: 10,   # data
    9: 16,   # visualizações
    10: 42,  # notas
}
for col, width in col_widths.items():
    ws.column_dimensions[get_column_letter(col)].width = width

# ── Legenda no rodapé ──────────────────────────────────────────────────────
footer_row = len(data) + 6
ws.merge_cells(f"A{footer_row}:J{footer_row}")
leg = ws[f"A{footer_row}"]
leg.value = (
    "🔵 Azul = Vídeos com Claude AI  |  🔴 Rosa = Vídeos com IA Geral  |  "
    "⚠️ 'Verificar no YT' = abrir o link e anotar o nº de visualizações actual  |  "
    "Dados recolhidos em Junho 2026"
)
leg.font = Font(name="Calibri", size=8, italic=True, color="555577")
leg.fill = PatternFill("solid", fgColor="EEEEFF")
leg.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
ws.row_dimensions[footer_row].height = 28

# ── Congelar painel no cabeçalho ──────────────────────────────────────────
ws.freeze_panes = "A5"

# ── Guardar ────────────────────────────────────────────────────────────────
path = "/home/user/vibe/Tutoriais_IA_Websites.xlsx"
wb.save(path)
print(f"Ficheiro guardado: {path}")
