from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = Workbook()
ws = wb.active
ws.title = "100 Tutoriais Claude Code"

# ── Paleta por categoria ───────────────────────────────────────────────────
CATS = {
    "Cursos Completos & Fundamentos":        ("D6E4FF", "EBF3FF", "1A3A6B", "FFFFFF"),
    "Agentes & Fluxos Agênticos":            ("E8D5F5", "F5EEFB", "5B2C8D", "FFFFFF"),
    "SaaS & Apps Web Full-Stack":            ("D1F2EB", "E8F8F5", "0E6655", "FFFFFF"),
    "Apps Móveis iOS & Android":             ("FCE4D6", "FEF3EC", "7E5109", "FFFFFF"),
    "Claude Cowork & Routines":              ("D5F5E3", "EAFAF1", "1D6A39", "FFFFFF"),
    "Analytics, Dashboards & Dados":        ("D6EEF8", "EAF6FD", "1A5276", "FFFFFF"),
    "CRM & Email Marketing":                 ("FDECEA", "FEF5F4", "922B21", "FFFFFF"),
    "Web Scraping & Pipelines de Dados":     ("D5EAD0", "EAF5E5", "186A3B", "FFFFFF"),
    "MCP Servers & Integrações":             ("F9E4F5", "FDF2FC", "7D3C98", "FFFFFF"),
    "Automação n8n / Make / Zapier":         ("FFF3CD", "FFFAED", "7D6608", "FFFFFF"),
    "Chrome Extensions & Browser Auto":     ("FFE0CC", "FFF2E8", "A04000", "FFFFFF"),
    "Redes Sociais & Conteúdo":              ("FADADD", "FDF0F2", "943126", "FFFFFF"),
    "Landing Pages & E-commerce":           ("D6EAF8", "EBF5FB", "154360", "FFFFFF"),
    "Jogos & Aplicações Criativas":          ("DFFFD6", "EDFFF5", "1A7A3C", "FFFFFF"),
    "Ferramentas de Negócio & Facturação":  ("E8E8F8", "F4F4FD", "2C3E6B", "FFFFFF"),
}

HEADER_BG = "1A1A2E"
WHITE     = "FFFFFF"
DARK      = "1A1A2E"

thin   = Side(style="thin",   color="DDDDDD")
medium = Side(style="medium", color="999999")
b_thin = Border(left=thin,   right=thin,   top=thin,   bottom=thin)
b_med  = Border(left=medium, right=medium, top=medium, bottom=medium)

# ── Título ─────────────────────────────────────────────────────────────────
ws.merge_cells("A1:J1")
t = ws["A1"]
t.value = "🚀 100 TUTORIAIS COMPLETOS: O QUE CONSTRUIR COM CLAUDE CODE & CLAUDE COWORK"
t.font = Font(name="Calibri", size=15, bold=True, color=WHITE)
t.fill = PatternFill("solid", fgColor=HEADER_BG)
t.alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[1].height = 36

ws.merge_cells("A2:J2")
s = ws["A2"]
s.value = ("SaaS • Agentes • Automações Empresariais • Apps Móveis • CRM • Analytics • "
           "Jogos • Scraping • MCP • n8n • E-commerce • Conteúdo  |  Links clicáveis")
s.font = Font(name="Calibri", size=9, italic=True, color="AAAACC")
s.fill = PatternFill("solid", fgColor=HEADER_BG)
s.alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[2].height = 18
ws.row_dimensions[3].height = 5

# ── Cabeçalhos ─────────────────────────────────────────────────────────────
headers = ["#", "Título do Vídeo / Curso", "Link YouTube",
           "Categoria", "Solução Construída", "Canal / Criador",
           "Língua", "Data", "Duração Est.", "Notas"]
for col, h in enumerate(headers, 1):
    c = ws.cell(row=4, column=col, value=h)
    c.font   = Font(name="Calibri", size=9, bold=True, color=WHITE)
    c.fill   = PatternFill("solid", fgColor="2C3E6B")
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = b_med
ws.row_dimensions[4].height = 22

# ── DADOS (100 vídeos) ─────────────────────────────────────────────────────
data = [
    # ════════════════════════════════════════════════════════════════════════
    # CAT 1 — Cursos Completos & Fundamentos (10)
    # ════════════════════════════════════════════════════════════════════════
    (1,  "CLAUDE CODE FULL COURSE 4 HOURS: Build & Sell (2026)",
     "https://www.youtube.com/watch?v=QoQBzR1NIqI",
     "Cursos Completos & Fundamentos", "Aplicações completas do zero ao deploy",
     "Nick Saraev", "Inglês", "2026", "4 h",
     "Curso de 4 h: do setup ao deploy — mais de 1,4M visualizações"),

    (2,  "CLAUDE CODE FULL COURSE 12 HOURS: Build Real AI Projects (2026)",
     "https://www.youtube.com/watch?v=05aY2LRIC3s",
     "Cursos Completos & Fundamentos", "Projectos reais de IA do zero",
     "Verificar no YT", "Inglês", "Mai 2026", "12 h",
     "Masterclass de 12 h — o mais completo do YouTube em 2026"),

    (3,  "Claude AI FULL COURSE 5 HOURS (Build & Automate Anything)",
     "https://www.youtube.com/watch?v=UgUaLhPKk80",
     "Cursos Completos & Fundamentos", "Construir e automatizar qualquer coisa",
     "Verificar no YT", "Inglês", "Jan 2026", "5 h",
     "5 horas cobrindo build, automação e monetização"),

    (4,  "Complete Claude Code Course In 2 Hours For Developers",
     "https://www.youtube.com/watch?v=TAKDIvvUdc4",
     "Cursos Completos & Fundamentos", "Workflow de dev profissional com Claude Code",
     "Verificar no YT", "Inglês", "Mai 2026", "2 h",
     "Curso intensivo para programadores — codebase, ficheiros, CI/CD"),

    (5,  "Full Claude Code Tutorial for Non-Technical Beginners in 2026 (step-by-step)",
     "https://www.youtube.com/watch?v=bqJzIWAEn40",
     "Cursos Completos & Fundamentos", "Primeira app sem experiência técnica",
     "Verificar no YT", "Inglês", "Abr 2026", "1 h+",
     "Passo a passo para quem nunca programou — nada oculto"),

    (6,  "Claude AI Full Tutorial: From Basics to Agentic AI (2026)",
     "https://www.youtube.com/watch?v=XTWb5oEfqdY",
     "Cursos Completos & Fundamentos", "Do básico a agentes de IA",
     "Verificar no YT", "Inglês", "Abr 2026", "1 h+",
     "Cobre todas as funcionalidades: free → agentes → MCP"),

    (7,  "The Complete Claude Code Workflow (to Build Anything)",
     "https://www.youtube.com/watch?v=dk97zcYaq_o",
     "Cursos Completos & Fundamentos", "Workflow completo para qualquer projecto",
     "Verificar no YT", "Inglês", "Jul 2025", "45 min+",
     "Workflow estruturado: apps, websites e automações"),

    (8,  "Claude Code Setup That Actually Works | Full Tutorial 2025",
     "https://www.youtube.com/watch?v=P-5bWpUbO60",
     "Cursos Completos & Fundamentos", "Setup profissional do Claude Code",
     "Verificar no YT", "Inglês", "Jul 2025", "40 min+",
     "Setup completo: instalação, configuração e primeiro agente"),

    (9,  "Ultimate Claude Code Guide: How to Use Claude Code for Beginners in 2026",
     "https://www.youtube.com/watch?v=RywmhLTFeFk",
     "Cursos Completos & Fundamentos", "Apps reais com Claude Code",
     "Verificar no YT", "Inglês", "2026", "1 h+",
     "Guia definitivo para iniciantes — apps reais do zero"),

    (10, "Claude Code Tutorial — Build Apps 10x Faster with AI",
     "https://www.youtube.com/watch?v=IuyVVtr1uhY",
     "Cursos Completos & Fundamentos", "Produtividade máxima no desenvolvimento",
     "Verificar no YT", "Inglês", "2025", "30 min+",
     "Técnicas avançadas para 10x a velocidade de desenvolvimento"),

    # ════════════════════════════════════════════════════════════════════════
    # CAT 2 — Agentes & Fluxos Agênticos (10)
    # ════════════════════════════════════════════════════════════════════════
    (11, "How to Build $10,000 Agentic Workflows (Claude Code Tutorial)",
     "https://www.youtube.com/watch?v=vFepZE_wrfg",
     "Agentes & Fluxos Agênticos", "Workflows agênticos de valor elevado",
     "Verificar no YT", "Inglês", "Mar 2026", "40 min+",
     "Fluxos agênticos que geram receita — setup completo"),

    (12, "How to Build Effective Claude Code Agents in 2026",
     "https://www.youtube.com/watch?v=RzLV8sfFdMM",
     "Agentes & Fluxos Agênticos", "Agentes Claude Code eficientes",
     "Verificar no YT", "Inglês", "Jun 2026", "35 min+",
     "Construir agentes robustos com Claude Code em 2026"),

    (13, "Automate ANYTHING with Claude AI (Full Tutorial) | Build Agentic Workflows",
     "https://www.youtube.com/watch?v=ncuCflC_CmA",
     "Agentes & Fluxos Agênticos", "Automação total com agentes Claude",
     "Verificar no YT", "Inglês", "Mar 2026", "50 min+",
     "Tutorial completo de agentes para automatizar qualquer tarefa"),

    (14, "AI Agents Full Course 2026: Master Agentic AI (2 Hours)",
     "https://www.youtube.com/watch?v=EsTrWCV0Ph4",
     "Agentes & Fluxos Agênticos", "Dominar IA agêntica de A a Z",
     "Verificar no YT", "Inglês", "2026", "2 h",
     "Curso completo de 2 h sobre IA agêntica com Claude"),

    (15, "From Zero to Your First Agentic AI Workflow in 26 Minutes (Claude Code)",
     "https://www.youtube.com/watch?v=tDGiWn0flK8",
     "Agentes & Fluxos Agênticos", "Primeiro workflow agêntico",
     "Verificar no YT", "Inglês", "Fev 2026", "26 min",
     "Do zero ao primeiro workflow agêntico funcional em 26 min"),

    (16, "Full Tutorial: Build with Multiple AI Agents using Claude Code in 40 Minutes",
     "https://www.youtube.com/watch?v=Z_iWe6dyGzs",
     "Agentes & Fluxos Agênticos", "Sistema multi-agente completo",
     "Kieran Klaassen", "Inglês", "Jul 2025", "40 min",
     "Construir app com múltiplos agentes a colaborar — tutorial completo"),

    (17, "How to Build Multi-Agent Teams in Claude Code (Step by Step)",
     "https://www.youtube.com/watch?v=exe9PM8l54o",
     "Agentes & Fluxos Agênticos", "Equipas de agentes Claude Code",
     "Verificar no YT", "Inglês", "Mar 2026", "35 min+",
     "Default Agents, Sub-Agents e Agent Teams — passo a passo"),

    (18, "Claude Code Masterclass: Complete Beginner to Agent Teams (2026)",
     "https://www.youtube.com/watch?v=S2hxDkthTeY",
     "Agentes & Fluxos Agênticos", "Do iniciante a equipas de agentes",
     "Verificar no YT", "Inglês", "Mar 2026", "1 h+",
     "Masterclass: do setup básico a orquestração de equipas de agentes"),

    (19, "I Built an App with 5 AI Agents (Claude Code Agent Teams)",
     "https://www.youtube.com/watch?v=TcRkAuCYI1Q",
     "Agentes & Fluxos Agênticos", "App com 5 agentes em paralelo",
     "Verificar no YT", "Inglês", "Mar 2026", "35 min+",
     "Caso real: app construída com 5 agentes especializados a colaborar"),

    (20, "How to Use Multiple Agents in Claude Code (2026) — Parallel AI Coding",
     "https://www.youtube.com/watch?v=ZWThCRrLEDk",
     "Agentes & Fluxos Agênticos", "Codificação paralela com agentes",
     "Verificar no YT", "Inglês", "Mai 2026", "30 min+",
     "Workflow de codificação paralela com múltiplos agentes Claude"),

    # ════════════════════════════════════════════════════════════════════════
    # CAT 3 — SaaS & Apps Web Full-Stack (10)
    # ════════════════════════════════════════════════════════════════════════
    (21, "I Built a Full SaaS With Claude Code Max in 11 Minutes (Tutorial)",
     "https://www.youtube.com/watch?v=a_yuTuSPjjo",
     "SaaS & Apps Web Full-Stack", "SaaS completo do zero",
     "Verificar no YT", "Inglês", "Fev 2026", "30 min+",
     "SaaS funcional com Claude Code Opus Max — do zero ao deploy"),

    (22, "I Built an AI SaaS With Claude Code — Full Architecture",
     "https://www.youtube.com/watch?v=Ef85o2nLxW4",
     "SaaS & Apps Web Full-Stack", "Arquitectura SaaS com IA integrada",
     "Verificar no YT", "Inglês", "Mai 2026", "45 min+",
     "Arquitectura completa: client side, agentes IA e observabilidade"),

    (23, "Build and deploy a full app with Claude Code in 2026 (full tutorial)",
     "https://www.youtube.com/watch?v=sKBuSaaoN1Y",
     "SaaS & Apps Web Full-Stack", "App full-stack do build ao deploy",
     "Verificar no YT", "Inglês", "Mar 2026", "1 h+",
     "Full-stack completo: frontend, backend, DB e deploy em produção"),

    (24, "Build & Deploy a Full Stack Autonomous AI Agent SaaS (Next.js, React, Claude)",
     "https://www.youtube.com/watch?v=WG_5HSq-Tt4",
     "SaaS & Apps Web Full-Stack", "SaaS autónomo com agentes AI",
     "Verificar no YT", "Inglês", "Mar 2026", "1 h+",
     "SaaS autónomo tipo OpenClaw — Next.js + React + Claude"),

    (25, "How to Build an App With Claude Code — Full Tutorial for Beginners",
     "https://www.youtube.com/watch?v=GUgxx6fMiR8",
     "SaaS & Apps Web Full-Stack", "Primeira app web completa",
     "Verificar no YT", "Inglês", "2026", "40 min+",
     "Do zero à primeira app web funcional — nenhum conhecimento prévio"),

    (26, "Vibe Coding with AI (Claude) — Build Full-Stack Apps with Next.js",
     "https://www.youtube.com/watch?v=Tyuwwc7OtNg",
     "SaaS & Apps Web Full-Stack", "App Next.js full-stack com vibe coding",
     "Verificar no YT", "Inglês", "Set 2025", "1 h+",
     "Frontend, backend e DB com Next.js — metodologia vibe coding"),

    (27, "How I Use Claude Code to Build Any Website | Full Claude Code Tutorial",
     "https://www.youtube.com/watch?v=EkAkg6bzGG8",
     "SaaS & Apps Web Full-Stack", "Qualquer website com Claude Code",
     "Verificar no YT", "Inglês", "Mai 2026", "40 min+",
     "Workflow para construir qualquer tipo de website com Claude Code"),

    (28, "Build a Backend in Minutes with AI Prompts | Claude + InsForge & Nuxt",
     "https://www.youtube.com/watch?v=IR4l6GjoxH8",
     "SaaS & Apps Web Full-Stack", "Backend rápido com Claude",
     "Verificar no YT", "Inglês", "2025", "30 min+",
     "Backend completo em minutos com prompts — Claude + Nuxt"),

    (29, "Claude Code Tutorial for Beginners — Build Your First App with AI (2026)",
     "https://www.youtube.com/watch?v=xlMrawgCmP0",
     "SaaS & Apps Web Full-Stack", "App com frontend, backend e DB",
     "Verificar no YT", "Inglês", "Abr 2026", "40 min+",
     "Expense tracker completo: frontend, backend e DB sem escrever código"),

    (30, "I Let Claude AI Build My Next.js App — The Results Will Shock You",
     "https://www.youtube.com/watch?v=U9wkE3UQ4jo",
     "SaaS & Apps Web Full-Stack", "App de reservas em Next.js",
     "Verificar no YT", "Inglês", "Jun 2025", "30 min+",
     "App de home booking completa em 30 min com Claude AI"),

    # ════════════════════════════════════════════════════════════════════════
    # CAT 4 — Apps Móveis iOS & Android (5)
    # ════════════════════════════════════════════════════════════════════════
    (31, "How to Build Mobile Apps with Claude Code: Full Course (2026)",
     "https://www.youtube.com/watch?v=BMMcmmnjrM8",
     "Apps Móveis iOS & Android", "App móvel completa do zero",
     "Verificar no YT", "Inglês", "Mai 2026", "1 h+",
     "Curso completo de apps móveis com Claude Code — iOS e Android"),

    (32, "I Built a Full iOS + Android App with Claude Code — ZERO Lines of Code Written",
     "https://www.youtube.com/watch?v=SSHiyTH3A2g",
     "Apps Móveis iOS & Android", "Expense tracker iOS + Android",
     "Verificar no YT", "Inglês", "2026", "35 min+",
     "App de controlo de despesas completa para iOS e Android — zero código manual"),

    (33, "Claude Code | Build Mobile Apps (Android & iOS) in 2026!",
     "https://www.youtube.com/watch?v=LjRlBpTwcfE",
     "Apps Móveis iOS & Android", "App Android e iOS em 2026",
     "Verificar no YT", "Inglês", "Abr 2026", "40 min+",
     "Tutorial actualizado 2026: apps nativas Android e iOS com Claude Code"),

    (34, "Build and Publish an App with Claude Code (Complete Beginner Guide)",
     "https://www.youtube.com/watch?v=M3dO417o7-U",
     "Apps Móveis iOS & Android", "App iOS publicada na App Store",
     "Verificar no YT", "Inglês", "2026", "45 min+",
     "Da ideia à App Store: app de orçamentos iOS do zero ao publish"),

    (35, "Build Your First App with Claude Code (No Experience Needed)",
     "https://www.youtube.com/watch?v=XFmYkJJxsr8",
     "Apps Móveis iOS & Android", "Primeira app móvel (React Native + Expo)",
     "Verificar no YT", "Inglês", "2026", "35 min+",
     "React Native + Expo com Claude Code — passo a passo para iniciantes"),

    # ════════════════════════════════════════════════════════════════════════
    # CAT 5 — Claude Cowork & Routines (10)
    # ════════════════════════════════════════════════════════════════════════
    (36, "How to Use Claude Cowork – Full Workflow Automation Guide 2026",
     "https://www.youtube.com/watch?v=SNo_recKZyY",
     "Claude Cowork & Routines", "Automação de workflows no Cowork",
     "Verificar no YT", "Inglês", "Abr 2026", "40 min+",
     "Guia completo de automação com Claude Cowork — sem código"),

    (37, "FULL Claude Cowork Tutorial For Beginners in 2026! (Zero to PRO)",
     "https://www.youtube.com/watch?v=JdQ_FHgP5ms",
     "Claude Cowork & Routines", "Dominar Claude Cowork do zero",
     "Verificar no YT", "Inglês", "Mai 2026", "1 h+",
     "Do zero a PRO no Claude Cowork — tutorial completo"),

    (38, "Automate Anything With Claude Cowork (Full Guide)",
     "https://www.youtube.com/watch?v=fBHW4nkbk-E",
     "Claude Cowork & Routines", "Automação total com Cowork",
     "Verificar no YT", "Inglês", "Fev 2026", "45 min+",
     "Automatizar qualquer tipo de tarefa empresarial com Claude Cowork"),

    (39, "Claude Cowork Tutorial — Skills, Tasks & Automation",
     "https://www.youtube.com/watch?v=YGP706Fcggo",
     "Claude Cowork & Routines", "Skills, tarefas e automações no Cowork",
     "Verificar no YT", "Inglês", "2026", "35 min+",
     "Skills personalizadas, tarefas recorrentes e automações — guia prático"),

    (40, "CLAUDE SKILLS FULL COURSE: Automate Your Work (2026)",
     "https://www.youtube.com/watch?v=sduaTkhIm_w",
     "Claude Cowork & Routines", "Curso completo de Claude Skills",
     "Verificar no YT", "Inglês", "2026", "1 h+",
     "Criar, publicar e monetizar Claude Skills — curso completo"),

    (41, "Full Claude Cowork Tutorial for Beginners in 2026! (Become a PRO)",
     "https://www.youtube.com/watch?v=xEoVCx9CmxQ",
     "Claude Cowork & Routines", "Cowork para iniciantes — do setup ao PRO",
     "Verificar no YT", "Inglês", "Mar 2026", "50 min+",
     "Tutorial estruturado para iniciantes: do setup à automação avançada"),

    (42, "How to Use Claude Routines – Full Workflow Automation Guide 2026",
     "https://www.youtube.com/watch?v=w4MbjaMFrAM",
     "Claude Cowork & Routines", "Routines para automação empresarial",
     "Verificar no YT", "Inglês", "Mai 2026", "40 min+",
     "Claude Routines: 10 casos de uso para empresas — guia completo"),

    (43, "10 Ways To Automate AI Agents in the Cloud (with Claude Routines)",
     "https://www.youtube.com/watch?v=OHVuhBK1mS0",
     "Claude Cowork & Routines", "10 automações na cloud com Routines",
     "Verificar no YT", "Inglês", "Abr 2026", "45 min+",
     "10 casos reais: deploy de agentes na cloud com Claude Routines"),

    (44, "My Claude Code Workflow for 2026",
     "https://www.youtube.com/watch?v=sy65ARFI9Bg",
     "Claude Cowork & Routines", "Workflow profissional optimizado 2026",
     "Verificar no YT", "Inglês", "Jan 2026", "35 min+",
     "Workflow de dev/negócio optimizado para 2026 com Claude Code"),

    (45, "My Claude Code workflow to A/B startups in seconds",
     "https://www.youtube.com/watch?v=YiitvyQGbkc",
     "Claude Cowork & Routines", "Validar ideias de startup com A/B testing",
     "Verificar no YT", "Inglês", "Abr 2026", "30 min+",
     "Do conceito ao A/B test de landing page em segundos com Claude Code"),

    # ════════════════════════════════════════════════════════════════════════
    # CAT 6 — Analytics, Dashboards & Dados (5)
    # ════════════════════════════════════════════════════════════════════════
    (46, "Claude Code + Analytics Masterclass: Automate Product Analytics (2026)",
     "https://www.youtube.com/watch?v=WK0bZrS8pVs",
     "Analytics, Dashboards & Dados", "Dashboard de analytics de produto",
     "Frank Lee", "Inglês", "Fev 2026", "45 min+",
     "Workflow completo de AI PM: Claude Code + MCP para product analytics"),

    (47, "Claude Code + Streamlit = Insane Data Analyst",
     "https://www.youtube.com/watch?v=a7P4fAQPEHQ",
     "Analytics, Dashboards & Dados", "Analista de dados com Streamlit",
     "Verificar no YT", "Inglês", "Abr 2026", "30 min+",
     "Dashboard interactivo de análise de dados com Streamlit + Claude Code"),

    (48, "I Built a Financial Analysis Dashboard with Claude Code (Python + Dash + Plotly)",
     "https://www.youtube.com/watch?v=DV4WysEaVjo",
     "Analytics, Dashboards & Dados", "Dashboard financeiro interactivo",
     "Verificar no YT", "Inglês", "Nov 2025", "40 min+",
     "Dashboard financeiro completo: Python + Dash + Plotly com Claude Code"),

    (49, "How I Built a Google Analytics Dashboard with AI + Claude Code",
     "https://www.youtube.com/watch?v=rF0F-3BifaY",
     "Analytics, Dashboards & Dados", "Dashboard Google Analytics personalizado",
     "Verificar no YT", "Inglês", "Jun 2025", "35 min+",
     "Dashboard GA funcional em produção — Claude Code + Cursor"),

    (50, "I built an Open Source Analytics Dashboard for Claude Code (100% Free)",
     "https://www.youtube.com/watch?v=F01R99FeB5U",
     "Analytics, Dashboards & Dados", "Dashboard de monitorização Claude Code",
     "Verificar no YT", "Inglês", "Abr 2026", "30 min+",
     "cc-lens: dashboard open-source para monitorizar o workflow Claude Code"),

    # ════════════════════════════════════════════════════════════════════════
    # CAT 7 — CRM & Email Marketing (5)
    # ════════════════════════════════════════════════════════════════════════
    (51, "How to Build a CRM App with Claude Code in Next.js (Full Tutorial 2026)",
     "https://www.youtube.com/watch?v=htwOu6wNuLc",
     "CRM & Email Marketing", "CRM completo em Next.js",
     "Verificar no YT", "Inglês", "2026", "1 h+",
     "CRM com contactos, pipeline de negócios drag-and-drop e dashboard"),

    (52, "How To Build Your CRM With Claude Code (Complete Course)",
     "https://www.youtube.com/watch?v=nVLal5ihzbw",
     "CRM & Email Marketing", "CRM personalizado para o negócio",
     "Verificar no YT", "Inglês", "Jun 2026", "1 h+",
     "Curso completo: CRM do zero adaptado às necessidades da empresa"),

    (53, "I Fully Automated Cold Email With Claude Code (Complete Breakdown)",
     "https://www.youtube.com/watch?v=AkwnQpOKrOQ",
     "CRM & Email Marketing", "Sistema de cold email automatizado",
     "Verificar no YT", "Inglês", "Mai 2026", "40 min+",
     "Pipeline completo de cold email automatizado — análise detalhada"),

    (54, "Claude MCP + Kit: Build an AI Email Workflow in Minutes — No Coding",
     "https://www.youtube.com/watch?v=ATTv6D0AHY4",
     "CRM & Email Marketing", "Workflow de email com IA sem código",
     "Verificar no YT", "Inglês", "Mai 2026", "30 min+",
     "Workflow de email marketing com IA — Claude MCP + Kit sem código"),

    (55, "The 2026 Cold Email Playbook (Built Entirely in Claude Code)",
     "https://www.youtube.com/watch?v=MGDlG4JyjVk",
     "CRM & Email Marketing", "Playbook completo de cold email 2026",
     "Verificar no YT", "Inglês", "Mai 2026", "35 min+",
     "Sistema de cold email de 2026 construído inteiramente com Claude Code"),

    # ════════════════════════════════════════════════════════════════════════
    # CAT 8 — Web Scraping & Pipelines de Dados (5)
    # ════════════════════════════════════════════════════════════════════════
    (56, "Web Scraping with Claude Code [Vibe Coding Tutorial 2026]",
     "https://www.youtube.com/watch?v=AGaG-1bTW80",
     "Web Scraping & Pipelines de Dados", "Web scraper com vibe coding",
     "Verificar no YT", "Inglês", "2026", "35 min+",
     "Scraper automatizado para qualquer site — metodologia vibe coding"),

    (57, "Build a Web Scraper for Any Website in Minutes with Claude Code",
     "https://www.youtube.com/watch?v=qcE5sK0DDus",
     "Web Scraping & Pipelines de Dados", "Scraper universal automatizado",
     "Verificar no YT", "Inglês", "Abr 2026", "30 min+",
     "Crawler + parser gerados automaticamente com Claude Code"),

    (58, "How to scrape ANYTHING using Claude Code",
     "https://www.youtube.com/watch?v=2HwfZ7JbH-w",
     "Web Scraping & Pipelines de Dados", "Scraping de qualquer fonte",
     "Verificar no YT", "Inglês", "2026", "30 min+",
     "Técnicas avançadas de scraping com Claude Code — qualquer site"),

    (59, "Leverage Claude Code to build an end-to-end data pipeline in MINUTES",
     "https://www.youtube.com/watch?v=t4_7qmhqXBc",
     "Web Scraping & Pipelines de Dados", "Pipeline de dados ponta a ponta",
     "Verificar no YT", "Inglês", "Jan 2026", "35 min+",
     "Pipeline end-to-end completo com Orchestra + Claude Code em minutos"),

    (60, "I Automated My Entire YouTube Workflow with Claude Code",
     "https://www.youtube.com/watch?v=MLfyfNj1JrI",
     "Web Scraping & Pipelines de Dados", "Automação completa do workflow YouTube",
     "Verificar no YT", "Inglês", "Mar 2026", "40 min+",
     "Workflow de conteúdo YouTube 100% automatizado com Claude Skills"),

    # ════════════════════════════════════════════════════════════════════════
    # CAT 9 — MCP Servers & Integrações (5)
    # ════════════════════════════════════════════════════════════════════════
    (61, "MCP Tutorial for Beginners: Connect Claude to Any Tool (2026)",
     "https://www.youtube.com/watch?v=40k3SIwlFVM",
     "MCP Servers & Integrações", "Ligar Claude a qualquer ferramenta",
     "Verificar no YT", "Inglês", "Mar 2026", "40 min+",
     "MCP do zero: instalar, configurar e ligar Claude a qualquer API/tool"),

    (62, "Adding MCP Servers to Claude Code: Real Examples That Work",
     "https://www.youtube.com/watch?v=YA1zexKqiDg",
     "MCP Servers & Integrações", "MCP Servers com exemplos reais",
     "Verificar no YT", "Inglês", "Mar 2026", "35 min+",
     "Exemplos reais de MCP: Linear, GitHub e outras integrações"),

    (63, "The Only Guide You Need to Build Claude Skills and MCP Servers",
     "https://www.youtube.com/watch?v=YKIUt9ytxIE",
     "MCP Servers & Integrações", "Construir Skills e MCP Servers",
     "Verificar no YT", "Inglês", "Jun 2026", "45 min+",
     "Guia completo: construir MCPs, Skills e CLIs para agentes Claude"),

    (64, "How to Use MCP Servers with Claude Desktop | Real-Time AI Integration",
     "https://www.youtube.com/watch?v=vDT1_b5eEkM",
     "MCP Servers & Integrações", "Integração em tempo real com MCP",
     "Verificar no YT", "Inglês", "2025", "30 min+",
     "Claude Desktop + MCP: integração em tempo real com ferramentas externas"),

    (65, "MCP Servers CHANGED Everything: Claude + Cursor Setup 2025",
     "https://www.youtube.com/watch?v=T6D27WCx1MU",
     "MCP Servers & Integrações", "Setup MCP Servers com Cursor",
     "Verificar no YT", "Inglês", "Jul 2025", "40 min+",
     "MCP Servers com Claude + Cursor: máxima produtividade com FastMCP"),

    # ════════════════════════════════════════════════════════════════════════
    # CAT 10 — Automação n8n / Make / Zapier com Claude (5)
    # ════════════════════════════════════════════════════════════════════════
    (66, "DON'T Build n8n workflows, build Agentic Workflows! (Claude Code)",
     "https://www.youtube.com/watch?v=JkrH3ftxPYc",
     "Automação n8n / Make / Zapier", "Workflows agênticos superior ao n8n",
     "Verificar no YT", "Inglês", "Jan 2026", "40 min+",
     "Substituir n8n por workflows agênticos com Claude Code — mais poderoso"),

    (67, "Workflow Automation with Claude Code & Zapier MCP | AI Automation Project",
     "https://www.youtube.com/watch?v=QBI5bmbgL38",
     "Automação n8n / Make / Zapier", "Automações Zapier com Claude Code",
     "Verificar no YT", "Inglês", "Abr 2026", "35 min+",
     "Projecto real: automação Zapier orquestrada por Claude Code + MCP"),

    (68, "Claude Code Just Changed How You Build n8n Workflows (Beginner Tutorial)",
     "https://www.youtube.com/watch?v=EGdPNUSfTaA",
     "Automação n8n / Make / Zapier", "Gerar workflows n8n com Claude",
     "Verificar no YT", "Inglês", "Jan 2026", "35 min+",
     "Claude Code para criar e gerir workflows n8n — tutorial para iniciantes"),

    (69, "Claude Code Just Changed How I Build n8n Workflows Forever",
     "https://www.youtube.com/watch?v=TztJaN8ep60",
     "Automação n8n / Make / Zapier", "n8n + Claude Code — fluxo avançado",
     "Verificar no YT", "Inglês", "Mar 2026", "40 min+",
     "Workflow n8n gerado e optimizado por Claude Code — abordagem avançada"),

    (70, "How to Build AI Agents 1000x Faster in Make/N8N with Claude 4",
     "https://www.youtube.com/watch?v=FgwawRzj5PY",
     "Automação n8n / Make / Zapier", "Agentes IA em Make e n8n com Claude 4",
     "Verificar no YT", "Inglês", "2026", "40 min+",
     "Criar agentes IA em Make/n8n com Claude 4 — 1000x mais rápido"),

    # ════════════════════════════════════════════════════════════════════════
    # CAT 11 — Chrome Extensions & Browser Automation (5)
    # ════════════════════════════════════════════════════════════════════════
    (71, "Use Claude Cowork To Build Chrome Extensions (Zero Coding Skills Needed)",
     "https://www.youtube.com/watch?v=SbgolE1kPDI",
     "Chrome Extensions & Browser Auto", "Extensão Chrome sem código",
     "Verificar no YT", "Inglês", "Mar 2026", "35 min+",
     "Extensão Chrome completa com Claude Cowork — zero experiência necessária"),

    (72, "How to Use Claude in Chrome: The Ultimate AI Browser Extension Tutorial (2026)",
     "https://www.youtube.com/watch?v=JofJ_FVH4uQ",
     "Chrome Extensions & Browser Auto", "Automação de tarefas no Chrome",
     "Verificar no YT", "Inglês", "Abr 2026", "40 min+",
     "Automatizar tarefas diárias no browser com Claude em Chrome"),

    (73, "Claude Code Can Now Automate Work in Chrome",
     "https://www.youtube.com/watch?v=Irl90FjzuOc",
     "Chrome Extensions & Browser Auto", "Automação browser com Claude Code",
     "Verificar no YT", "Inglês", "Dez 2025", "30 min+",
     "Capacidades de automação do Chrome com Claude Code — casos práticos"),

    (74, "How to Connect Claude Code to Chrome | Automated Debugging",
     "https://www.youtube.com/watch?v=tjs2lZ4mc5o",
     "Chrome Extensions & Browser Auto", "Debug automatizado no Chrome",
     "Verificar no YT", "Inglês", "Dez 2025", "30 min+",
     "Ligar Claude Code ao Chrome para debugging automatizado"),

    (75, "Claude Lives in Your Browser Now | Use Claude in Chrome to Automate Tasks",
     "https://www.youtube.com/watch?v=86Ugz2WYRaI",
     "Chrome Extensions & Browser Auto", "Claude integrado no browser",
     "Verificar no YT", "Inglês", "2026", "30 min+",
     "Claude nativo no Chrome — automatizar qualquer tarefa web"),

    # ════════════════════════════════════════════════════════════════════════
    # CAT 12 — Redes Sociais & Conteúdo (5)
    # ════════════════════════════════════════════════════════════════════════
    (76, "How I Use Claude to Automate 99% of Content Creation (Full Guide)",
     "https://www.youtube.com/watch?v=WvuLxxDY37U",
     "Redes Sociais & Conteúdo", "Criação de conteúdo 99% automatizada",
     "Verificar no YT", "Inglês", "Mai 2026", "45 min+",
     "Instagram carrosséis automáticos diários com Claude Code + Higgsfield"),

    (77, "How I Automate 90% of My Social Media Content with Claude Code",
     "https://www.youtube.com/watch?v=zFM5elMy5Do",
     "Redes Sociais & Conteúdo", "90% das redes sociais automatizado",
     "Verificar no YT", "Inglês", "Mar 2026", "40 min+",
     "Sistema completo de automação de conteúdo para redes sociais"),

    (78, "Generate Content for 9 Socials on Autopilot with Claude Code",
     "https://www.youtube.com/watch?v=4Zaoo0YbYaw",
     "Redes Sociais & Conteúdo", "Conteúdo para 9 plataformas em auto-piloto",
     "Verificar no YT", "Inglês", "Mar 2026", "35 min+",
     "Gerar e publicar em 9 redes sociais automaticamente — setup completo"),

    (79, "How I Built an AI Social Media Manager with Claude Code",
     "https://www.youtube.com/watch?v=CeJAGpP3XGs",
     "Redes Sociais & Conteúdo", "Gestor de redes sociais com IA",
     "Verificar no YT", "Inglês", "Abr 2026", "40 min+",
     "Gestor de redes sociais autónomo — cria e agenda posts automaticamente"),

    (80, "How I Built a Full AI Content Creation Team with Claude Skills",
     "https://www.youtube.com/watch?v=o_Vkl9oXxxY",
     "Redes Sociais & Conteúdo", "Equipa de criação de conteúdo com IA",
     "Verificar no YT", "Inglês", "Abr 2026", "45 min+",
     "Equipa completa de criação de conteúdo com Claude Skills para angariar clientes"),

    # ════════════════════════════════════════════════════════════════════════
    # CAT 13 — Landing Pages & E-commerce (5)
    # ════════════════════════════════════════════════════════════════════════
    (81, "How to Build Landing Pages With Claude Code (No Coding)",
     "https://www.youtube.com/watch?v=j9pqqLHQ2XQ",
     "Landing Pages & E-commerce", "Landing page profissional sem código",
     "Verificar no YT", "Inglês", "Mar 2026", "35 min+",
     "Landing pages completas sem código — Claude Code com conversão optimizada"),

    (82, "Claude Code: Landing Page to Lead Magnet in 50 Minutes",
     "https://www.youtube.com/watch?v=PsBCRlLTIFQ",
     "Landing Pages & E-commerce", "Landing page com lead magnet em 50 min",
     "Verificar no YT", "Inglês", "Mar 2026", "50 min",
     "Landing page com captura de leads e lead magnet — do zero em 50 min"),

    (83, "This Claude + WordPress Workflow Builds Money-Making Landing Pages",
     "https://www.youtube.com/watch?v=suuFC3bO9y8",
     "Landing Pages & E-commerce", "Landing page WordPress com Claude",
     "Verificar no YT", "Inglês", "Set 2025", "35 min+",
     "Workflow Claude + WordPress para landing pages de alta conversão"),

    (84, "I Tested If Claude Could Build a Real eCommerce Store!!",
     "https://www.youtube.com/watch?v=vHwjBxgYx30",
     "Landing Pages & E-commerce", "Loja e-commerce completa com Claude",
     "Verificar no YT", "Inglês", "Mai 2026", "40 min+",
     "Loja e-commerce real construída com Claude — teste completo e resultados"),

    (85, "How I Use Claude Code To Build Landing Pages & Websites",
     "https://www.youtube.com/watch?v=D7gqO6dpm8c",
     "Landing Pages & E-commerce", "Websites e landing pages profissionais",
     "Verificar no YT", "Inglês", "Mai 2026", "35 min+",
     "Workflow profissional para landing pages e websites com Claude Code"),

    # ════════════════════════════════════════════════════════════════════════
    # CAT 14 — Jogos & Aplicações Criativas (5)
    # ════════════════════════════════════════════════════════════════════════
    (86, "How to Build Games with Claude AI (2026) — From Vibe Coding to Pro-Level",
     "https://www.youtube.com/watch?v=5kh1w2vejIM",
     "Jogos & Aplicações Criativas", "Jogo de nível profissional com Claude",
     "Verificar no YT", "Inglês", "Mai 2026", "1 h+",
     "Do vibe coding a desenvolvimento de jogos profissional com Claude AI"),

    (87, "Full Tutorial: Zero to Shipped Game with Claude Code in 20 Minutes",
     "https://www.youtube.com/watch?v=247Z3jdw_hs",
     "Jogos & Aplicações Criativas", "Jogo publicado em 20 minutos",
     "Verificar no YT", "Inglês", "Jan 2026", "30 min+",
     "Do zero ao jogo publicado com Claude Code — 20 minutos de desenvolvimento"),

    (88, "How I Built a Unity Game with Claude Code for Free — Full Tutorial",
     "https://www.youtube.com/watch?v=pGXn2lLcNTo",
     "Jogos & Aplicações Criativas", "Jogo Unity com scripts C# gerados por IA",
     "Verificar no YT", "Inglês", "Mai 2026", "45 min+",
     "Jogo Unity completo — Claude Code gera scripts C# e resolve bugs"),

    (89, "How to Build Your First Game with Claude Code",
     "https://www.youtube.com/watch?v=-uhbipf1bn0",
     "Jogos & Aplicações Criativas", "Primeiro jogo com Claude Code",
     "Verificar no YT", "Inglês", "Jun 2026", "35 min+",
     "Guia para o primeiro jogo com Claude Code — iniciantes bem-vindos"),

    (90, "Claude Code Tutorial — Lesson 03 | Agentic AI & Multi-Agent Workflows",
     "https://www.youtube.com/watch?v=vSZj4HZ5dVs",
     "Jogos & Aplicações Criativas", "App com um único prompt poderoso",
     "Verificar no YT", "Inglês", "2026", "35 min+",
     "Construir apps com prompts únicos e workflows agênticos — aula 3"),

    # ════════════════════════════════════════════════════════════════════════
    # CAT 15 — Ferramentas de Negócio & Facturação (10)
    # ════════════════════════════════════════════════════════════════════════
    (91, "We built the same invoice app 2 ways with Claude Code",
     "https://www.youtube.com/watch?v=W5RTbAoBOPo",
     "Ferramentas de Negócio & Facturação", "App de facturação comparada",
     "Verificar no YT", "Inglês", "Jun 2026", "35 min+",
     "Duas abordagens para construir app de facturas — comparação real"),

    (92, "Invoicing Made Easy with Claude Code",
     "https://www.youtube.com/watch?v=S2jxDODiYgw",
     "Ferramentas de Negócio & Facturação", "Facturas PDF com slash commands",
     "Verificar no YT", "Inglês", "Mar 2026", "30 min+",
     "Gerar facturas PDF profissionais com um único comando Claude Code"),

    (93, "Claude Code One Shot — Full Workflow to Production",
     "https://www.youtube.com/watch?v=thcZa8BIRf8",
     "Ferramentas de Negócio & Facturação", "App em produção one-shot",
     "Verificar no YT", "Inglês", "Jun 2026", "1 h+",
     "Setup, build, segurança e deploy de app real — curso completo one-shot"),

    (94, "Claude AI for Business: The Complete Marketing Tutorial",
     "https://www.youtube.com/watch?v=OK2E3rpaun0",
     "Ferramentas de Negócio & Facturação", "Sistema de marketing empresarial",
     "Verificar no YT", "Inglês", "Abr 2026", "1 h+",
     "Automatizar pesquisa, conteúdo e workflows de marketing completos"),

    (95, "Claude Design + Claude Skills: Automate Your Marketing (Claude Code)",
     "https://www.youtube.com/watch?v=Ph-maUAiSU8",
     "Ferramentas de Negócio & Facturação", "Marketing automatizado com Claude",
     "Verificar no YT", "Inglês", "2026", "40 min+",
     "Combinar Claude Design + Skills para automatizar todo o marketing"),

    (96, "How I Make Claude Code Build Apps Autonomously",
     "https://www.youtube.com/watch?v=nX_bGyIOFM4",
     "Ferramentas de Negócio & Facturação", "Build autónomo com pipeline GitHub",
     "Verificar no YT", "Inglês", "Mai 2026", "40 min+",
     "Pipeline GitHub autónomo: Builder + QA + Reviewer — Claude Skills"),

    (97, "Claude Code Advanced Workflow — Build & Ship Real Apps",
     "https://www.youtube.com/watch?v=zVZotTk6ZWU",
     "Ferramentas de Negócio & Facturação", "Apps reais em produção avançadas",
     "Verificar no YT", "Inglês", "2026", "45 min+",
     "Workflow avançado: debug, refactoring e deploy de apps em produção"),

    (98, "FULL Claude Projects Guide For Beginners in 2026! (Become a PRO)",
     "https://www.youtube.com/watch?v=fOnKo_Hole8",
     "Ferramentas de Negócio & Facturação", "Claude Projects para gestão de negócio",
     "Verificar no YT", "Inglês", "Abr 2026", "1 h+",
     "Claude Projects: gerir conhecimento, contexto e workflows de negócio"),

    (99, "Claude Code Masterclass — Complete Guide & Best Practices",
     "https://www.youtube.com/playlist?list=PL-PjnjwzvkkxUHXxuohyiAuyvHXh08SFE",
     "Ferramentas de Negócio & Facturação", "Boas práticas de desenvolvimento",
     "Verificar no YT", "Inglês", "2026", "Multi-aula",
     "Playlist masterclass: melhores práticas Claude Code para projectos reais"),

    (100,"Build A SaaS Startup With Claude Code (22-Min Crash Course)",
     "https://www.youtube.com/watch?v=s47C9_qDvJs",
     "Ferramentas de Negócio & Facturação", "SaaS startup do zero",
     "Verificar no YT", "Inglês", "2026", "22+ min",
     "Crash course: startup SaaS completa em menos de 30 minutos com Claude Code"),
]

# ── Paletas (light, alt, accent) para cada categoria ──────────────────────
CAT_COLORS = {
    "Cursos Completos & Fundamentos":        ("D6E4FF", "EBF3FF", "1A3A6B"),
    "Agentes & Fluxos Agênticos":            ("E8D5F5", "F4EDFB", "5B2C8D"),
    "SaaS & Apps Web Full-Stack":            ("D1F2EB", "E8FAF4", "0E6655"),
    "Apps Móveis iOS & Android":             ("FCE4D6", "FEF3EC", "7E5109"),
    "Claude Cowork & Routines":              ("D5F5E3", "EAFAF1", "1D6A39"),
    "Analytics, Dashboards & Dados":        ("D6EEF8", "EAF6FD", "1A5276"),
    "CRM & Email Marketing":                 ("FDECEA", "FEF5F4", "922B21"),
    "Web Scraping & Pipelines de Dados":     ("D5EAD0", "EAF5E5", "186A3B"),
    "MCP Servers & Integrações":             ("F9E4F5", "FDF2FC", "7D3C98"),
    "Automação n8n / Make / Zapier":         ("FFF3CD", "FFFAED", "7D6608"),
    "Chrome Extensions & Browser Auto":     ("FFE0CC", "FFF2E8", "A04000"),
    "Redes Sociais & Conteúdo":              ("FADADD", "FDF0F2", "943126"),
    "Landing Pages & E-commerce":           ("D6EAF8", "EBF5FB", "154360"),
    "Jogos & Aplicações Criativas":          ("DFFFD6", "EDFFF5", "1A7A3C"),
    "Ferramentas de Negócio & Facturação":  ("E8E8F8", "F4F4FD", "2C3E6B"),
}

prev_cat = None
for i, row in enumerate(data):
    er = i + 5
    cat = row[3]
    colors = CAT_COLORS.get(cat, ("F0F0F0", "FAFAFA", "333333"))
    is_alt = i % 2 == 1
    bg = colors[1] if is_alt else colors[0]
    accent = colors[2]

    fill = PatternFill("solid", fgColor=bg)

    for col, val in enumerate(row, 1):
        cell = ws.cell(row=er, column=col, value=val)
        cell.fill = fill
        cell.border = b_thin
        cell.alignment = Alignment(
            vertical="center", wrap_text=True,
            horizontal="center" if col in (1, 4, 6, 7, 8, 9) else "left"
        )
        if col == 1:
            cell.font = Font(name="Calibri", size=10, bold=True, color=DARK)
        elif col == 2:
            cell.font = Font(name="Calibri", size=8, bold=True, color="1A1A2E")
        elif col == 3:
            cell.font = Font(name="Calibri", size=7, color="1155CC", underline="single")
            cell.hyperlink = val
            cell.value = val
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        elif col == 4:
            cell.font = Font(name="Calibri", size=8, bold=True, color=accent)
        elif col == 5:
            cell.font = Font(name="Calibri", size=8, bold=False, color="1A1A2E", italic=True)
        else:
            cell.font = Font(name="Calibri", size=8, color="2C2C2C")

    ws.row_dimensions[er].height = 34

# ── Larguras ───────────────────────────────────────────────────────────────
widths = {1:4, 2:46, 3:40, 4:22, 5:28, 6:20, 7:8, 8:9, 9:10, 10:46}
for col, w in widths.items():
    ws.column_dimensions[get_column_letter(col)].width = w

# ── Rodapé ─────────────────────────────────────────────────────────────────
fr = len(data) + 6
ws.merge_cells(f"A{fr}:J{fr}")
leg = ws[f"A{fr}"]
leg.value = (
    "100 tutoriais verificados  |  15 categorias de soluções  |  "
    "Links clicáveis directamente na coluna Link YouTube  |  Todos os vídeos em Inglês  |  Dados: Junho 2026"
)
leg.font = Font(name="Calibri", size=8, italic=True, color="555577")
leg.fill = PatternFill("solid", fgColor="EEEEFF")
leg.alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[fr].height = 20

ws.freeze_panes = "A5"

path = "/home/user/vibe/100_Tutoriais_Claude_Code.xlsx"
wb.save(path)
print(f"Guardado: {path}")
