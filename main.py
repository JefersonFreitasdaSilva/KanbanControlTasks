import os
import json
import random
import sys
import time
import tkinter as tk
from tkinter import ttk, Text, simpledialog, messagebox, colorchooser
import threading
from pystray import Icon, MenuItem, Menu
from PIL import Image, ImageTk
import requests
from datetime import datetime
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# Configurações
REGISTRO_TAREFAS = "tasks.json"
CONFIG_ARQUIVO = "config.json"
DATA_DIR = os.path.join(os.path.expanduser("~"), "TaskManagerData")  # External directory: ~/TaskManagerData
ESTADOS_PADRAO = ["To Do", "In Progress", "Done"]
estados = ESTADOS_PADRAO.copy()
colunas = {}
tarefas_widgets = {}
CORES_COLUNAS = {}  # Already initialized globally
resize_timer = None
layout_lock = False
arrastando = None
tarefa_arrastada = None
widget_fantasma = None
coluna_arrastada = None
widget_fantasma_coluna = None
ANIMATION_DURATION = 400  # Animation duration in milliseconds
ANIMATION_STEPS = 20     # Number of animation steps for smooth movement
CORES_PASTEL = {
    "To Do": "#B3E5FC",
    "In Progress": "#C8E6C9",
    "Done": "#F3E5F5",
    "default": "#FFE0B2"
}
from PIL import Image


# Ensure the external data directory exists
if not os.path.exists(DATA_DIR):
    os.makedirs(DATA_DIR)
    print(f"[INFO] Created data directory: {DATA_DIR}")

# Função para gerar cor pastel aleatória
def gerar_cor_pastel_aleatoria():
    r = random.randint(180, 255)
    g = random.randint(180, 255)
    b = random.randint(180, 255)
    return f"#{r:02x}{g:02x}{b:02x}"

# Função para obter cor da coluna
def obter_cor_coluna(estado):
    return CORES_COLUNAS.get(estado, CORES_PASTEL.get(estado, CORES_PASTEL["default"]))


# Função para salvar configurações
def salvar_configuracoes(url, estados, cores_colunas):
    config = {"url": url, "estados": estados, "cores_colunas": cores_colunas}
    config_path = os.path.join(DATA_DIR, CONFIG_ARQUIVO)
    try:
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
        print(f"[INFO] Configurações salvas em: {config_path}")
    except Exception as e:
        print(f"[ERROR] Erro ao salvar configurações: {e}")

# Função para carregar configurações
def carregar_configuracoes():
    config_path = os.path.join(DATA_DIR, CONFIG_ARQUIVO)
    try:
        if os.path.exists(config_path):
            with open(config_path, "r", encoding="utf-8") as f:
                config = json.load(f)
                CORES_COLUNAS.update(config.get("cores_colunas", {}))
                url = config.get("url", "")
                # Validate URL
                if url and not url.startswith(('http://', 'https://')):
                    print(f"[WARNING] URL inválida em config.json: {url}. Usando URL vazia.")
                    url = ""
                print(f"[INFO] Configurações carregadas de: {config_path}")
                return url, config.get("estados", ESTADOS_PADRAO)
        else:
            print(f"[INFO] config.json não encontrado. Criando novo arquivo em {config_path}.")
            default_config = {
                "url": "",
                "estados": ESTADOS_PADRAO,
                "cores_colunas": CORES_PASTEL
            }
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump(default_config, f, indent=4, ensure_ascii=False)
            CORES_COLUNAS.update(CORES_PASTEL)
            return "", ESTADOS_PADRAO
    except (FileNotFoundError, json.JSONDecodeError) as e:
        print(f"[ERROR] Erro ao carregar configurações: {e}")
        default_config = {
            "url": "",
            "estados": ESTADOS_PADRAO,
            "cores_colunas": CORES_PASTEL
        }
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(default_config, f, indent=4, ensure_ascii=False)
        CORES_COLUNAS.update(CORES_PASTEL)
        return "", ESTADOS_PADRAO

# Carregar configurações e sincronizar URLs
GOOGLE_SHEETS_URL, estados = carregar_configuracoes()
GOOGLE_SHEETS_API_URL = GOOGLE_SHEETS_URL

# Função para salvar tarefas localmente
def salvar_tarefas(tarefas):
    registro_path = os.path.join(DATA_DIR, REGISTRO_TAREFAS)
    try:
        with open(registro_path, "w", encoding="utf-8") as f:
            json.dump(tarefas, f, indent=4, ensure_ascii=False)
        print(f"[INFO] Tarefas salvas em: {registro_path}")
    except Exception as e:
        print(f"[ERROR] Erro ao salvar tarefas: {e}")

# NEW: Função para editar o nome da coluna
def editar_nome_coluna(estado):
    global estados, CORES_COLUNAS
    novo_nome = simpledialog.askstring("Editar Estado", f"Digite o novo nome para '{estado}':", initialvalue=estado)
    if novo_nome and novo_nome != estado and novo_nome not in estados:
        # Update estados
        idx = estados.index(estado)
        estados[idx] = novo_nome
        # Update CORES_COLUNAS
        if estado in CORES_COLUNAS:
            CORES_COLUNAS[novo_nome] = CORES_COLUNAS.pop(estado)
        # Update tasks
        tarefas = carregar_tarefas()
        for task_id, task in tarefas.items():
            if task.get("estado") == estado:
                task["estado"] = novo_nome
        salvar_tarefas(tarefas)
        # Save configurations
        salvar_configuracoes(GOOGLE_SHEETS_URL, estados, CORES_COLUNAS)
        # Update UI
        reordenar_colunas()
        atualizar_tarefas()
        enviar_tarefas_planilha()
        messagebox.showinfo("Estado Renomeado", f"Estado '{estado}' renomeado para '{novo_nome}'.")
    elif novo_nome in estados:
        messagebox.showwarning("Estado Existente", "Este nome de estado já existe.")
    elif not novo_nome:
        messagebox.showwarning("Nome Inválido", "O nome do estado não pode ser vazio.")

# NEW: Função para excluir uma coluna
def excluir_coluna(estado):
    global estados, CORES_COLUNAS
    if len(estados) <= 1:
        messagebox.showwarning("Erro", "Não é possível excluir a última coluna.")
        return
    tarefas = carregar_tarefas()
    tarefas_na_coluna = [task_id for task_id, task in tarefas.items() if task.get("estado") == estado]
    if tarefas_na_coluna:
        opcao = messagebox.askyesnocancel(
            "Tarefas Encontradas",
            f"A coluna '{estado}' contém {len(tarefas_na_coluna)} tarefa(s). Deseja mover as tarefas para outra coluna ou excluí-las?",
            detail="Sim: Mover tarefas\nNão: Excluir tarefas\nCancelar: Abortar"
        )
        if opcao is None:  # Cancel
            return
        elif opcao:  # Move tasks
            outras_colunas = [e for e in estados if e != estado]
            coluna_destino = simpledialog.askstring(
                "Mover Tarefas",
                f"Escolha a coluna para mover as tarefas de '{estado}':",
                initialvalue=outras_colunas[0]
            )
            if coluna_destino not in outras_colunas:
                messagebox.showerror("Erro", "Coluna inválida selecionada.")
                return
            for task_id in tarefas_na_coluna:
                tarefas[task_id]["estado"] = coluna_destino
        else:  # Delete tasks
            for task_id in tarefas_na_coluna:
                del tarefas[task_id]
        salvar_tarefas(tarefas)
    # Remove the column
    estados.remove(estado)
    if estado in CORES_COLUNAS:
        del CORES_COLUNAS[estado]
    salvar_configuracoes(GOOGLE_SHEETS_URL, estados, CORES_COLUNAS)
    reordenar_colunas()
    atualizar_tarefas()
    enviar_tarefas_planilha()
    messagebox.showinfo("Coluna Excluída", f"Coluna '{estado}' excluída com sucesso.")

# Função para enviar tarefas e estados para a planilha (unchanged)
def enviar_tarefas_planilha():
    if not GOOGLE_SHEETS_API_URL:
        janela.after(0, lambda: texto_detalhes.insert(tk.END, "⚠️ URL do Apps Script não configurada no config.json.\n", "info"))
        return False
    tarefas = carregar_tarefas()
    payload = {"tarefas": tarefas, "estados": estados}
    session = requests.Session()
    retries = Retry(total=3, backoff_factor=0.1, status_forcelist=[500, 502, 503, 504])
    session.mount('https://', HTTPAdapter(max_retries=retries))
    try:
        print(f"[DEBUG] Enviando tarefas e estados para {GOOGLE_SHEETS_API_URL}")
        response = session.post(GOOGLE_SHEETS_API_URL, json=payload, timeout=10)
        print(f"[DEBUG] Resposta do servidor: {response.status_code}, {response.text}")
        if response.status_code == 200:
            resposta = response.json()
            if resposta.get("status") == "success":
                janela.after(0, lambda: texto_detalhes.insert(tk.END, "✅ Tarefas e estados sincronizados com a planilha.\n", "info"))
                return True
            else:
                janela.after(0, lambda: texto_detalhes.insert(tk.END, f"⚠️ Erro do servidor: {resposta.get('message', 'Desconhecido')}\n", "info"))
                return False
        else:
            janela.after(0, lambda: texto_detalhes.insert(tk.END, f"⚠️ Falha ao enviar dados: {response.status_code} - {response.text}\n", "info"))
            return False
    except requests.exceptions.RequestException as e:
        print(f"[ERROR] Erro ao enviar dados: {e}")
        janela.after(0, lambda: texto_detalhes.insert(tk.END, f"⚠️ Erro ao enviar dados: {e}\n", "info"))
        return False

# Função para carregar tarefas locais
def carregar_tarefas():
    registro_path = os.path.join(DATA_DIR, REGISTRO_TAREFAS)
    try:
        if os.path.exists(registro_path):
            with open(registro_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict) and all(isinstance(v, dict) for v in data.values()):
                    print(f"[INFO] Tarefas carregadas de: {registro_path}")
                    return data
                else:
                    print(f"[WARNING] Arquivo tasks.json contém dados inválidos. Retornando vazio.")
                    return {}
        else:
            # Create empty tasks.json if it doesn't exist
            print(f"[INFO] tasks.json não encontrado. Criando novo arquivo em {registro_path}.")
            with open(registro_path, "w", encoding="utf-8") as f:
                json.dump({}, f, indent=4, ensure_ascii=False)
            return {}
    except (FileNotFoundError, json.JSONDecodeError) as e:
        print(f"[ERROR] Erro ao carregar tarefas: {e}")
        # Create empty tasks.json in case of error
        with open(registro_path, "w", encoding="utf-8") as f:
            json.dump({}, f, indent=4, ensure_ascii=False)
        return {}

# Função para alterar a URL do Google Sheets
def alterar_url():
    global GOOGLE_SHEETS_URL, GOOGLE_SHEETS_API_URL
    nova_url = simpledialog.askstring("Alterar URL", "Digite a nova URL do Google Sheets:",
                                      initialvalue=GOOGLE_SHEETS_URL)
    if nova_url:
        GOOGLE_SHEETS_URL = nova_url
        GOOGLE_SHEETS_API_URL = nova_url
        salvar_configuracoes(GOOGLE_SHEETS_URL, estados, CORES_COLUNAS)
        messagebox.showinfo("URL Alterada", "A URL do Google Sheets foi alterada.")

# Função para adicionar novo estado
def adicionar_estado():
    global estados
    novo_estado = simpledialog.askstring("Novo Estado", "Digite o novo estado:")
    if novo_estado and novo_estado not in estados:
        cor_escolhida = colorchooser.askcolor(title="Escolher cor da coluna")[1] or gerar_cor_pastel_aleatoria()
        CORES_COLUNAS[novo_estado] = cor_escolhida
        estados.append(novo_estado)
        salvar_configuracoes(GOOGLE_SHEETS_URL, estados, CORES_COLUNAS)
        reordenar_colunas()
        atualizar_tarefas()
        enviar_tarefas_planilha()  # Sincronizar estados com a planilha
        messagebox.showinfo("Estado Adicionado", f"Estado '{novo_estado}' adicionado com cor {cor_escolhida}.")
    elif novo_estado in estados:
        messagebox.showwarning("Estado Existente", "Este estado já existe.")

# Função para enviar tarefas e estados para a planilha (assíncrona)
def enviar_tarefas_planilha():
    if not GOOGLE_SHEETS_API_URL:
        janela.after(0, lambda: texto_detalhes.insert(tk.END, "⚠️ URL do Apps Script não configurada no config.json.\n", "info"))
        return False

    tarefas = carregar_tarefas()
    payload = {
        "tarefas": tarefas,
        "estados": estados
    }

    session = requests.Session()
    retries = Retry(total=3, backoff_factor=0.1, status_forcelist=[500, 502, 503, 504])
    session.mount('https://', HTTPAdapter(max_retries=retries))

    try:
        print(f"[DEBUG] Enviando tarefas e estados para {GOOGLE_SHEETS_API_URL}")
        response = session.post(GOOGLE_SHEETS_API_URL, json=payload, timeout=10)
        print(f"[DEBUG] Resposta do servidor: {response.status_code}, {response.text}")
        if response.status_code == 200:
            resposta = response.json()
            if resposta.get("status") == "success":
                janela.after(0, lambda: texto_detalhes.insert(tk.END, "✅ Tarefas e estados sincronizados com a planilha.\n", "info"))
                return True
            else:
                janela.after(0, lambda: texto_detalhes.insert(tk.END, f"⚠️ Erro do servidor: {resposta.get('message', 'Desconhecido')}\n", "info"))
                return False
        else:
            janela.after(0, lambda: texto_detalhes.insert(tk.END, f"⚠️ Falha ao enviar dados: {response.status_code} - {response.text}\n", "info"))
            return False
    except requests.exceptions.RequestException as e:
        print(f"[ERROR] Erro ao enviar dados: {e}")
        janela.after(0, lambda: texto_detalhes.insert(tk.END, f"⚠️ Erro ao enviar dados: {e}\n", "info"))
        return False

# Função para sincronizar com Google Sheets (assíncrona)
def sincronizar_com_planilha():
    def sincronizar_thread():
        global estados
        if not GOOGLE_SHEETS_API_URL:
            janela.after(0, lambda: texto_detalhes.insert(tk.END, "⚠️ URL do Apps Script não configurada no config.json.\n", "info"))
            return
        session = requests.Session()
        retries = Retry(total=3, backoff_factor=0.1, status_forcelist=[500, 502, 503, 504])
        session.mount('https://', HTTPAdapter(max_retries=retries))
        try:
            print(f"[DEBUG] Buscando tarefas e estados de {GOOGLE_SHEETS_API_URL}")
            response = session.get(GOOGLE_SHEETS_API_URL, timeout=10)
            print(f"[DEBUG] Resposta do servidor: {response.status_code}, {response.text}")
            if response.status_code == 200:
                dados = response.json()
                if dados.get("status") == "error":
                    janela.after(0, lambda: texto_detalhes.insert(tk.END, f"⚠️ Erro do servidor: {dados.get('message', 'Desconhecido')}\n", "info"))
                    return
                tarefas_planilha = dados.get("tarefas", {})
                estados_planilha = dados.get("estados", [])
                if not isinstance(tarefas_planilha, dict) or not all(isinstance(v, dict) for v in tarefas_planilha.values()):
                    janela.after(0, lambda: texto_detalhes.insert(tk.END, "⚠️ Dados de tarefas inválidos recebidos do servidor.\n", "info"))
                    return
                if not isinstance(estados_planilha, list) or not all(isinstance(e, str) for e in estados_planilha):
                    janela.after(0, lambda: texto_detalhes.insert(tk.END, "⚠️ Dados de estados inválidos recebidos do servidor.\n", "info"))
                    estados_planilha = estados
                if set(estados_planilha) != set(estados):
                    estados[:] = estados_planilha
                    salvar_configuracoes(GOOGLE_SHEETS_URL, estados, CORES_COLUNAS)
                    janela.after(0, reordenar_colunas)
                salvar_tarefas(tarefas_planilha)
                janela.after(0, atualizar_tarefas)
                janela.after(0, lambda: texto_detalhes.insert(tk.END, "✅ Tarefas e estados sincronizados da planilha.\n", "info"))
                janela.after(0, listar_tarefas_em_execucao)
            else:
                janela.after(0, lambda: texto_detalhes.insert(tk.END, f"⚠️ Falha ao buscar dados: {response.status_code} - {response.text}\n", "info"))
        except requests.exceptions.RequestException as e:
            print(f"[ERROR] Erro ao sincronizar: {e}")
            janela.after(0, lambda err=e: texto_detalhes.insert(tk.END, f"⚠️ Erro ao sincronizar: {err}\n", "info"))
    threading.Thread(target=sincronizar_thread, daemon=True).start()

# Função para listar tarefas em execução
def listar_tarefas_em_execucao():
    try:
        tarefas = carregar_tarefas()
        tarefas_em_execucao = [f"Ticket {task_id}: {task['titulo']}" for task_id, task in tarefas.items() if task.get("estado") == "In Progress"]
        texto_detalhes.delete("1.0", tk.END)
        if tarefas_em_execucao:
            texto_detalhes.insert(tk.END, "Tarefas em Execução:\n", "subtitulo")
            for tarefa in tarefas_em_execucao:
                texto_detalhes.insert(tk.END, f"{tarefa}\n", "item")
        else:
            texto_detalhes.insert(tk.END, "Nenhuma tarefa em execução.\n", "info")
    except Exception as e:
        print(f"[ERROR] Erro ao listar tarefas em execução: {e}")
        texto_detalhes.insert(tk.END, f"⚠️ Erro ao listar tarefas: {e}\n", "info")

# Função para criar uma coluna no Kanban
def criar_coluna(estado):
    try:
        if estado not in estados:
            print(f"[WARNING] Estado {estado} não está na lista de estados. Ignorando.")
            return
        cor_coluna = obter_cor_coluna(estado)
        frame_coluna = tk.Frame(frame_kanban_interno, bg=cor_coluna, bd=0, relief="flat")
        frame_coluna.configure(width=150)
        frame_coluna.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        label_coluna = tk.Label(frame_coluna, text=estado, bg=cor_coluna, fg="#333333", font=("Arial", 12, "bold"))
        label_coluna.pack(pady=5, fill=tk.X)

        frame_canvas = tk.Frame(frame_coluna, bg=cor_coluna)
        frame_canvas.pack(fill=tk.BOTH, expand=True)
        canvas_tarefas = tk.Canvas(frame_canvas, bg=cor_coluna, highlightthickness=0, width=150)
        canvas_tarefas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_vertical = ttk.Scrollbar(frame_canvas, orient="vertical", command=canvas_tarefas.yview, style="Custom.Vertical.TScrollbar")
        scrollbar_vertical.pack(side=tk.RIGHT, fill=tk.Y)
        canvas_tarefas.configure(yscrollcommand=scrollbar_vertical.set)

        janela.update_idletasks()
        canvas_height = frame_kanban_canvas.winfo_height() or 600
        scrollbar_height = scrollbar_horizontal.winfo_height() if scrollbar_horizontal.winfo_viewable() else 0
        available_height = canvas_height - scrollbar_height - 10
        canvas_tarefas.configure(height=available_height)

        raio = 20
        canvas_tarefas.create_rectangle(raio, 2, 150, available_height - 2, fill=cor_coluna, outline="", tags="bg_rect")
        canvas_tarefas.create_oval(2, 2, 2 + 2 * raio, 2 + 2 * raio, fill=cor_coluna, outline="", tags="corner")
        frame_tarefas = tk.Frame(canvas_tarefas, bg=cor_coluna)
        canvas_tarefas.create_window((0, 0), window=frame_tarefas, anchor="nw")

        def atualizar_scrollregion(event):
            canvas_tarefas.configure(scrollregion=canvas_tarefas.bbox("all"))
        frame_tarefas.bind("<Configure>", atualizar_scrollregion)

        def scroll_canvas(event):
            if event.widget == canvas_tarefas or isinstance(event.widget, tk.Canvas):
                if event.delta:
                    canvas_tarefas.yview_scroll(-1 * (event.delta // 120), "units")
                elif event.num == 4:
                    canvas_tarefas.yview_scroll(-1, "units")
                elif event.num == 5:
                    canvas_tarefas.yview_scroll(1, "units")
                return "break"
        canvas_tarefas.bind("<Enter>", lambda e: canvas_tarefas.focus_set())
        canvas_tarefas.bind("<MouseWheel>", scroll_canvas)
        canvas_tarefas.bind("<Button-4>", scroll_canvas)
        canvas_tarefas.bind("<Button-5>", scroll_canvas)
        frame_tarefas.bind("<MouseWheel>", lambda e: "break")
        frame_tarefas.bind("<Button-4>", lambda e: "break")
        frame_tarefas.bind("<Button-5>", lambda e: "break")

        # Context menu for column header
        def show_context_menu(event):
            try:
                context_menu = tk.Menu(janela)  # Removed tearoff
                context_menu.add_command(label="Editar Nome", command=lambda: editar_nome_coluna(estado))
                context_menu.add_command(label="Alterar Cor", command=lambda: editar_cor_coluna(estado))
                context_menu.add_command(label="Excluir Coluna", command=lambda: excluir_coluna(estado))
                context_menu.post(event.x_root, event.y_root)
            except Exception as e:
                print(f"[ERROR] Erro ao criar menu de contexto: {e}")

        # Bindings
        label_coluna.bind("<Button-3>", show_context_menu)  # Right-click for context menu
        label_coluna.bind("<Double-Button-1>", lambda e: editar_nome_coluna(estado))  # Double-click to edit name
        label_coluna.bind("<Button-1>", lambda e: iniciar_arrasto_coluna(e, estado))
        label_coluna.bind("<B1-Motion>", arrastar_coluna)
        label_coluna.bind("<ButtonRelease-1>", lambda e: soltar_coluna(e, estado))

        if frame_coluna and canvas_tarefas and frame_tarefas:
            colunas[estado] = {"canvas": canvas_tarefas, "frame": frame_tarefas, "frame_coluna": frame_coluna}
            atualizar_layout_colunas()
    except Exception as e:
        print(f"[ERROR] Erro ao criar coluna {estado}: {e}")

# Função para criar widget de tarefa
def criar_widget_tarefa(task_id, task_data, frame_coluna):
    try:
        frame_tarefa = tk.Frame(frame_coluna, bg="#000000", bd=1, relief="raised")
        frame_tarefa.pack(fill=tk.X, padx=5, pady=2)

        frame_topo = tk.Frame(frame_tarefa, bg="#000000")
        frame_topo.pack(fill=tk.X)

        prioridade = task_data.get("prioridade", "Média")
        cor_prioridade = {"Alta": "#ff4d4d", "Média": "#ffcc00", "Baixa": "#00cc00"}.get(prioridade, "#ffcc00")
        label_prioridade = tk.Label(frame_topo, text=prioridade, bg="#000000", fg=cor_prioridade, font=("Arial", 8, "bold"))
        label_prioridade.pack(side=tk.LEFT, padx=5)

        btn_menu = tk.Label(frame_topo, text="⋮", bg="#000000", fg="#ffffff", font=("Arial", 12, "bold"), cursor="hand2")
        btn_menu.pack(side=tk.RIGHT, padx=5)

        coluna_width = frame_coluna.master.winfo_width() or 150
        label_tarefa = tk.Label(frame_tarefa, text=f"Ticket {task_id}: {task_data.get('titulo', 'Sem título')}",
                                bg="#000000", fg="#ffffff", font=("Arial", 10), wraplength=max(50, coluna_width - 30))
        label_tarefa.pack(pady=5, padx=5, fill=tk.X)

        tarefas_widgets[task_id] = {
            "frame": frame_tarefa,
            "label_tarefa": label_tarefa,
            "label_prioridade": label_prioridade
        }

        def abrir_janela_opcoes(event):
            try:
                janela_opcoes = tk.Toplevel(janela)
                janela_opcoes.title(f"Opções da Tarefa {task_id}")
                janela_opcoes.geometry("200x150")
                janela_opcoes.configure(bg="#2e2e2e")
                janela_opcoes.resizable(False, False)

                janela_opcoes.update_idletasks()
                width = janela_opcoes.winfo_width()
                height = janela_opcoes.winfo_height()
                x = (janela_opcoes.winfo_screenwidth() // 2) - (width // 2)
                y = (janela_opcoes.winfo_screenheight() // 2) - (height // 2)
                janela_opcoes.geometry(f"{width}x{height}+{x}+{y}")

                tk.Label(janela_opcoes, text="Escolha uma opção:", bg="#2e2e2e", fg="#ffffff", font=("Arial", 10)).pack(pady=10)

                btn_editar = tk.Button(janela_opcoes, text="Editar", command=lambda: [janela_opcoes.destroy(), editar_tarefa(task_id)],
                                       bg="#007acc", fg="#ffffff", relief="flat", width=15)
                btn_editar.pack(pady=5)

                btn_apagar = tk.Button(janela_opcoes, text="Apagar", command=lambda: [janela_opcoes.destroy(), excluir_tarefa(task_id)],
                                       bg="#ff4d4d", fg="#ffffff", relief="flat", width=15)
                btn_apagar.pack(pady=5)
            except Exception as e:
                print(f"[ERROR] Erro ao abrir janela de opções para tarefa {task_id}: {e}")

        btn_menu.bind("<Button-1>", abrir_janela_opcoes)

        for widget in (frame_tarefa, label_tarefa, label_prioridade):
            widget.bind("<Button-1>", lambda e: iniciar_arrasto(e, task_id))
            widget.bind("<B1-Motion>", arrastar_tarefa)
            widget.bind("<ButtonRelease-1>", lambda e: soltar_tarefa(e, task_id))
        label_tarefa.bind("<Double-Button-1>", lambda e: mostrar_detalhes(task_id))
    except Exception as e:
        print(f"[ERROR] Erro ao criar widget da tarefa {task_id}: {e}")

# Função para atualizar tarefas
def atualizar_tarefas():
    global layout_lock
    if layout_lock:
        return
    layout_lock = True
    try:
        tarefas = carregar_tarefas()
        widgets_existentes = set(tarefas_widgets.keys())
        tarefas_atuais = set(tarefas.keys())

        for task_id in widgets_existentes - tarefas_atuais:
            if task_id in tarefas_widgets:
                frame = tarefas_widgets[task_id].get("frame")
                if frame and frame.winfo_exists():
                    frame.destroy()
                del tarefas_widgets[task_id]

        for task_id, task in tarefas.items():
            estado = task.get("estado", "To Do")
            if estado not in colunas or not colunas[estado]["frame"].winfo_exists():
                continue
            if task_id not in tarefas_widgets:
                criar_widget_tarefa(task_id, task, colunas[estado]["frame"])
            else:
                widget_info = tarefas_widgets[task_id]
                frame = widget_info.get("frame")
                label_tarefa = widget_info.get("label_tarefa")
                label_prioridade = widget_info.get("label_prioridade")

                if not (frame and frame.winfo_exists() and label_tarefa and label_tarefa.winfo_exists() and
                        label_prioridade and label_prioridade.winfo_exists()):
                    if frame and frame.winfo_exists():
                        frame.destroy()
                    del tarefas_widgets[task_id]
                    criar_widget_tarefa(task_id, task, colunas[estado]["frame"])
                    continue

                coluna_width = colunas[estado]["canvas"].winfo_width() or 75
                label_tarefa.configure(
                    text=f"Ticket {task_id}: {task.get('titulo', 'Sem título')}",
                    wraplength=max(50, coluna_width - 30)
                )
                prioridade = task.get("prioridade", "Média")
                cor_prioridade = {"Alta": "#ff4d4d", "Média": "#ffcc00", "Baixa": "#00cc00"}.get(prioridade, "#ffcc00")
                label_prioridade.configure(text=prioridade, fg=cor_prioridade)
                if frame.master != colunas[estado]["frame"] and colunas[estado]["frame"].winfo_exists():
                    frame.pack_forget()
                    try:
                        frame.pack(in_=colunas[estado]["frame"], fill=tk.X, padx=5, pady=2)
                    except tk.TclError as e:
                        print(f"[WARNING] Falha ao empacotar tarefa {task_id}: {e}")
                        frame.destroy()
                        del tarefas_widgets[task_id]
                        criar_widget_tarefa(task_id, task, colunas[estado]["frame"])

        atualizar_layout_colunas()
    except Exception as e:
        print(f"[ERROR] Erro ao atualizar tarefas: {e}")
        texto_detalhes.delete("1.0", tk.END)
        texto_detalhes.insert(tk.END, f"⚠️ Erro ao atualizar tarefas: {e}\n", "info")
    finally:
        layout_lock = False

# Função para editar cor da coluna
def editar_cor_coluna(estado):
    cor_escolhida = colorchooser.askcolor(title=f"Escolher cor para {estado}")[1]
    if cor_escolhida:
        CORES_COLUNAS[estado] = cor_escolhida
        salvar_configuracoes(GOOGLE_SHEETS_URL, estados, CORES_COLUNAS)
        colunas[estado]["frame_coluna"].configure(bg=cor_escolhida)
        colunas[estado]["canvas"].configure(bg=cor_escolhida)
        colunas[estado]["frame"].configure(bg=cor_escolhida)
        colunas[estado]["canvas"].delete("bg_rect", "corner")
        coluna_width = colunas[estado]["canvas"].winfo_width() or 75
        raio = 20
        colunas[estado]["canvas"].create_rectangle(raio, 2, coluna_width - raio,
                                                  colunas[estado]["canvas"].winfo_height() - 2,
                                                  fill=cor_escolhida, outline="", tags="bg_rect")
        colunas[estado]["canvas"].create_rectangle(2, raio, coluna_width - 2,
                                                  colunas[estado]["canvas"].winfo_height() - raio,
                                                  fill=cor_escolhida, outline="", tags="bg_rect")
        for x, y in [(2, 2), (coluna_width - 2 * raio - 2, 2),
                     (2, colunas[estado]["canvas"].winfo_height() - 2 * raio - 2),
                     (coluna_width - 2 * raio - 2, colunas[estado]["canvas"].winfo_height() - 2 * raio - 2)]:
            colunas[estado]["canvas"].create_oval(x, y, x + 2 * raio, y + 2 * raio, fill=cor_escolhida, outline="",
                                                 tags="corner")
        colunas[estado]["canvas"].lower("bg_rect", "corner")
        messagebox.showinfo("Cor Alterada", f"Cor da coluna '{estado}' alterada para {cor_escolhida}.")
        atualizar_layout_colunas()

# Funções para arrastar e soltar tarefas
def iniciar_arrasto(event, task_id):
    global arrastando, tarefa_arrastada, widget_fantasma
    try:
        arrastando = task_id
        tarefa_arrastada = tarefas_widgets.get(task_id, {}).get("frame")
        if tarefa_arrastada:
            widget_fantasma = tk.Label(janela, text=f"Ticket {task_id}", bg="#555555", fg="#ffffff", font=("Arial", 10))
            widget_fantasma.place(x=event.x_root - janela.winfo_rootx(), y=event.y_root - janela.winfo_rooty())
    except Exception as e:
        print(f"[ERROR] Erro ao iniciar arrasto da tarefa {task_id}: {e}")
        arrastando = None
        tarefa_arrastada = None
        if widget_fantasma:
            widget_fantasma.destroy()
            widget_fantasma = None

def arrastar_tarefa(event):
    global widget_fantasma
    try:
        if widget_fantasma and arrastando:
            widget_fantasma.place(x=event.x_root - janela.winfo_rootx(), y=event.y_root - janela.winfo_rooty())
    except Exception as e:
        print(f"[ERROR] Erro ao arrastar tarefa: {e}")

def soltar_tarefa(event, task_id):
    global arrastando, tarefa_arrastada, widget_fantasma
    try:
        if not arrastando or task_id != arrastando:
            return
        x, y = event.x_root, event.y_root
        tarefa_movida = False
        for estado, coluna in colunas.items():
            canvas = coluna["canvas"]
            if not canvas.winfo_exists():
                continue
            canvas_x = canvas.winfo_rootx()
            canvas_y = canvas.winfo_rooty()
            canvas_w = canvas.winfo_width()
            canvas_h = canvas.winfo_height()
            if canvas_x <= x <= canvas_x + canvas_w and canvas_y <= y <= canvas_y + canvas_h:
                tarefas = carregar_tarefas()
                if task_id in tarefas and tarefas[task_id].get("estado") != estado:
                    tarefas[task_id]["estado"] = estado
                    salvar_tarefas(tarefas)
                    if task_id in tarefas_widgets:
                        frame = tarefas_widgets[task_id].get("frame")
                        if frame and frame.winfo_exists():
                            frame.destroy()
                        del tarefas_widgets[task_id]
                    if coluna["frame"].winfo_exists():
                        criar_widget_tarefa(task_id, tarefas[task_id], coluna["frame"])
                    atualizar_layout_colunas()
                    texto_detalhes.delete("1.0", tk.END)
                    texto_detalhes.insert(tk.END, f"Tarefa {task_id} movida para '{estado}'.\n", "info")
                    tarefa_movida = True
                    enviar_tarefas_planilha()
                    listar_tarefas_em_execucao()
                break
        if not tarefa_movida:
            texto_detalhes.delete("1.0", tk.END)
            texto_detalhes.insert(tk.END, f"Tarefa {task_id} não foi movida.\n", "info")
    except Exception as e:
        print(f"[ERROR] Erro ao soltar tarefa {task_id}: {e}")
        texto_detalhes.delete("1.0", tk.END)
        texto_detalhes.insert(tk.END, f"⚠️ Erro ao mover tarefa: {e}\n", "info")
    finally:
        if widget_fantasma:
            widget_fantasma.destroy()
            widget_fantasma = None
        arrastando = None
        tarefa_arrastada = None

# Funções para arrastar e soltar colunas
def iniciar_arrasto_coluna(event, estado):
    global coluna_arrastada, widget_fantasma_coluna
    try:
        coluna_arrastada = estado
        widget_fantasma_coluna = tk.Label(janela, text=estado, bg="#777777", fg="#ffffff", font=("Arial", 12, "bold"),
                                         bd=2, relief="raised")
        widget_fantasma_coluna.place(x=event.x_root - janela.winfo_rootx(), y=event.y_root - janela.winfo_rooty())
    except Exception as e:
        print(f"[ERROR] Erro ao iniciar arrasto da coluna {estado}: {e}")
        coluna_arrastada = None
        if widget_fantasma_coluna:
            widget_fantasma_coluna.destroy()
            widget_fantasma_coluna = None

def arrastar_coluna(event):
    global widget_fantasma_coluna
    try:
        if widget_fantasma_coluna and coluna_arrastada:
            widget_fantasma_coluna.place(x=event.x_root - janela.winfo_rootx(), y=event.y_root - janela.winfo_rooty())
    except Exception as e:
        print(f"[ERROR] Erro ao arrastar coluna: {e}")

def soltar_coluna(event, estado):
    global coluna_arrastada, widget_fantasma_coluna
    try:
        if not coluna_arrastada or estado != coluna_arrastada:
            return
        x = event.x_root
        nova_posicao = None
        colunas_ordenadas = sorted(colunas.items(), key=lambda item: item[1]["frame_coluna"].winfo_rootx())
        for est, coluna in colunas_ordenadas:
            canvas_x = coluna["frame_coluna"].winfo_rootx()
            canvas_x_end = canvas_x + coluna["frame_coluna"].winfo_width()
            if canvas_x <= x < canvas_x_end:
                nova_posicao = est
                break
        if not nova_posicao and x >= colunas_ordenadas[-1][1]["frame_coluna"].winfo_rootx():
            nova_posicao = colunas_ordenadas[-1][0]
        if nova_posicao and nova_posicao != estado:
            estados.remove(estado)
            idx = estados.index(nova_posicao)
            if x >= colunas[nova_posicao]["frame_coluna"].winfo_rootx() + colunas[nova_posicao]["frame_coluna"].winfo_width() / 2:
                idx += 1
            estados.insert(idx, estado)
            salvar_configuracoes(GOOGLE_SHEETS_URL, estados, CORES_COLUNAS)
            # Animate column reordering instead of instant repack
            animar_colunas()
            texto_detalhes.delete("1.0", tk.END)
            texto_detalhes.insert(tk.END, f"Coluna '{estado}' movida para a posição de '{nova_posicao}'.\n", "info")
            # Sync with Google Sheets
            enviar_tarefas_planilha()
    except Exception as e:
        print(f"[ERROR] Erro ao soltar coluna {estado}: {e}")
        texto_detalhes.delete("1.0", tk.END)
        texto_detalhes.insert(tk.END, f"⚠️ Erro ao mover coluna: {e}\n", "info")
    finally:
        if widget_fantasma_coluna:
            widget_fantasma_coluna.destroy()
            widget_fantasma_coluna = None
        coluna_arrastada = None

# New function to animate column transitions
def animar_colunas():
    global layout_lock
    if layout_lock:
        return
    layout_lock = True
    try:
        # Store initial and target positions
        coluna_width = 150 + 10  # Width + padding
        initial_positions = {estado: colunas[estado]["frame_coluna"].winfo_x() for estado in colunas}
        target_positions = {estado: idx * coluna_width for idx, estado in enumerate(estados)}

        # Switch to place geometry manager for animation
        for estado in colunas:
            colunas[estado]["frame_coluna"].pack_forget()
            colunas[estado]["frame_coluna"].place(x=initial_positions[estado], y=5, width=150, height=colunas[estado]["frame_coluna"].winfo_height())

        def animate_step(step):
            if step > ANIMATION_STEPS:
                # Animation complete, restore pack layout
                for estado in colunas:
                    colunas[estado]["frame_coluna"].place_forget()
                reordenar_colunas_post_animation()
                return
            fraction = step / ANIMATION_STEPS
            for estado in colunas:
                if not colunas[estado]["frame_coluna"].winfo_exists():
                    continue
                current_x = initial_positions[estado] + (target_positions[estado] - initial_positions[estado]) * fraction
                colunas[estado]["frame_coluna"].place(x=current_x, y=5)
            janela.after(ANIMATION_DURATION // ANIMATION_STEPS, lambda: animate_step(step + 1))

        # Start animation
        animate_step(1)
    except Exception as e:
        print(f"[ERROR] Erro ao animar colunas: {e}")
        texto_detalhes.delete("1.0", tk.END)
        texto_detalhes.insert(tk.END, f"⚠️ Erro ao animar colunas: {e}\n", "info")
        # Fallback to instant reordering
        reordenar_colunas_post_animation()
    finally:
        layout_lock = False

# Modified reordenar_colunas to work post-animation
def reordenar_colunas_post_animation():
    global layout_lock
    if layout_lock:
        return
    layout_lock = True
    try:
        # Repack columns in the correct order
        for estado in colunas:
            colunas[estado]["frame_coluna"].pack_forget()
        for estado in estados:
            if estado in colunas and colunas[estado]["frame_coluna"].winfo_exists():
                colunas[estado]["frame_coluna"].pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        atualizar_layout_colunas()
    except Exception as e:
        print(f"[ERROR] Erro ao reordenar colunas: {e}")
        texto_detalhes.delete("1.0", tk.END)
        texto_detalhes.insert(tk.END, f"⚠️ Erro ao reordenar colunas: {e}\n", "info")
    finally:
        layout_lock = False

# Modified reordenar_colunas to avoid destroying widgets
def reordenar_colunas():
    global layout_lock
    if layout_lock:
        return
    layout_lock = True
    try:
        # Only animate if called from soltar_coluna; otherwise, just update layout
        for estado in estados:
            if estado not in colunas:
                criar_coluna(estado)
        atualizar_layout_colunas()
    except Exception as e:
        print(f"[ERROR] Erro ao reordenar colunas: {e}")
        texto_detalhes.delete("1.0", tk.END)
        texto_detalhes.insert(tk.END, f"⚠️ Erro ao reordenar colunas: {e}\n", "info")
    finally:
        layout_lock = False

# Função para reordenar colunas
def reordenar_colunas():
    global layout_lock
    if layout_lock:
        return
    layout_lock = True
    try:
        for task_id, widget_info in list(tarefas_widgets.items()):
            frame = widget_info.get("frame")
            if frame and frame.winfo_exists():
                frame.destroy()
            del tarefas_widgets[task_id]

        for estado in list(colunas.keys()):
            frame = colunas[estado]["frame_coluna"]
            if frame and frame.winfo_exists():
                frame.destroy()
            del colunas[estado]

        for estado in estados:
            criar_coluna(estado)

        tarefas = carregar_tarefas()
        for task_id, task in tarefas.items():
            estado = task.get("estado", "To Do")
            if estado in colunas and colunas[estado]["frame"].winfo_exists():
                criar_widget_tarefa(task_id, task, colunas[estado]["frame"])

        janela.update_idletasks()
        atualizar_layout_colunas()
    except Exception as e:
        print(f"[ERROR] Erro ao reordenar colunas: {e}")
        texto_detalhes.delete("1.0", tk.END)
        texto_detalhes.insert(tk.END, f"⚠️ Erro ao reordenar colunas: {e}\n", "info")
    finally:
        layout_lock = False

# Função para mostrar detalhes da tarefa
def mostrar_detalhes(task_id):
    try:
        tarefas = carregar_tarefas()
        if task_id in tarefas:
            task = tarefas[task_id]
            texto_detalhes.delete("1.0", tk.END)
            texto_detalhes.insert(tk.END, f"Ticket {task_id}: {task.get('titulo', 'Sem título')}\n", "titulo")
            texto_detalhes.insert(tk.END, f"Estado: {task.get('estado', 'Desconhecido')}\n", "info")
            texto_detalhes.insert(tk.END, f"Prioridade: {task.get('prioridade', 'Média')}\n", "info")
            texto_detalhes.insert(tk.END, f"Descrição: {task.get('descricao', 'Sem descrição')}\n", "item")
            texto_detalhes.insert(tk.END, f"Criado em: {task.get('data_criacao', 'Desconhecido')}\n", "item")
    except Exception as e:
        print(f"[ERROR] Erro ao mostrar detalhes da tarefa {task_id}: {e}")
        texto_detalhes.delete("1.0", tk.END)
        texto_detalhes.insert(tk.END, f"⚠️ Erro ao mostrar detalhes: {e}\n", "info")

# Função para adicionar tarefa
def adicionar_tarefa():
    def salvar_nova_tarefa():
        titulo = entry_titulo.get()
        descricao = texto_descricao.get("1.0", tk.END).strip()
        estado = combo_estado.get()
        prioridade = combo_prioridade.get()
        if not titulo or not estado:
            messagebox.showwarning("Campos Obrigatórios", "Título e estado são obrigatórios.")
            return
        tarefas = carregar_tarefas()
        task_id = str(len(tarefas) + 1)
        task = {
            "titulo": titulo,
            "descricao": descricao,
            "estado": estado,
            "prioridade": prioridade,
            "data_criacao": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        tarefas[task_id] = task
        salvar_tarefas(tarefas)
        enviar_tarefas_planilha()
        atualizar_tarefas()
        janela_tarefa.destroy()
        messagebox.showinfo("Tarefa Adicionada", f"Tarefa '{titulo}' adicionada com sucesso.")
        listar_tarefas_em_execucao()

    janela_tarefa = tk.Toplevel(janela)
    janela_tarefa.title("Adicionar Tarefa")
    janela_tarefa.geometry("400x350")
    janela_tarefa.configure(bg="#2e2e2e")

    tk.Label(janela_tarefa, text="Título:", bg="#2e2e2e", fg="#ffffff").pack(pady=5)
    entry_titulo = tk.Entry(janela_tarefa, width=40, bg="#333333", fg="#ffffff")
    entry_titulo.pack(pady=5, fill=tk.X, padx=10)

    tk.Label(janela_tarefa, text="Descrição:", bg="#2e2e2e", fg="#ffffff").pack(pady=5)
    texto_descricao = Text(janela_tarefa, height=5, bg="#333333", fg="#ffffff")
    texto_descricao.pack(pady=5, fill=tk.X, padx=10)

    tk.Label(janela_tarefa, text="Estado:", bg="#2e2e2e", fg="#ffffff").pack(pady=5)
    combo_estado = ttk.Combobox(janela_tarefa, values=estados, state="readonly")
    combo_estado.set(estados[0])
    combo_estado.pack(pady=5, fill=tk.X, padx=10)

    tk.Label(janela_tarefa, text="Prioridade:", bg="#2e2e2e", fg="#ffffff").pack(pady=5)
    combo_prioridade = ttk.Combobox(janela_tarefa, values=["Alta", "Média", "Baixa"], state="readonly")
    combo_prioridade.set("Média")
    combo_prioridade.pack(pady=5, fill=tk.X, padx=10)

    btn_salvar = tk.Button(janela_tarefa, text="Salvar", command=salvar_nova_tarefa, bg="#007acc", fg="#ffffff", relief="flat")
    btn_salvar.pack(pady=10)

# Função para editar tarefa
def editar_tarefa(task_id):
    try:
        tarefas = carregar_tarefas()
        if task_id not in tarefas:
            messagebox.showerror("Erro", f"Tarefa {task_id} não encontrada.")
            return
        task = tarefas[task_id]

        def salvar_tarefa_editada():
            titulo = entry_titulo.get()
            descricao = texto_descricao.get("1.0", tk.END).strip()
            estado = combo_estado.get()
            prioridade = combo_prioridade.get()
            if not titulo or not estado:
                messagebox.showwarning("Campos Obrigatórios", "Título e estado são obrigatórios.")
                return
            tarefas[task_id] = {
                "titulo": titulo,
                "descricao": descricao,
                "estado": estado,
                "prioridade": prioridade,
                "data_criacao": task.get("data_criacao", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            }
            salvar_tarefas(tarefas)
            enviar_tarefas_planilha()
            atualizar_tarefas()
            janela_tarefa.destroy()
            messagebox.showinfo("Tarefa Atualizada", f"Tarefa '{titulo}' atualizada com sucesso.")
            listar_tarefas_em_execucao()

        janela_tarefa = tk.Toplevel(janela)
        janela_tarefa.title(f"Editar Tarefa {task_id}")
        janela_tarefa.geometry("400x350")
        janela_tarefa.configure(bg="#2e2e2e")

        tk.Label(janela_tarefa, text="Título:", bg="#2e2e2e", fg="#ffffff").pack(pady=5)
        entry_titulo = tk.Entry(janela_tarefa, width=40, bg="#333333", fg="#ffffff")
        entry_titulo.insert(0, task.get("titulo", ""))
        entry_titulo.pack(pady=5, fill=tk.X, padx=10)

        tk.Label(janela_tarefa, text="Descrição:", bg="#2e2e2e", fg="#ffffff").pack(pady=5)
        texto_descricao = Text(janela_tarefa, height=5, bg="#333333", fg="#ffffff")
        texto_descricao.insert("1.0", task.get("descricao", ""))
        texto_descricao.pack(pady=5, fill=tk.X, padx=10)

        tk.Label(janela_tarefa, text="Estado:", bg="#2e2e2e", fg="#ffffff").pack(pady=5)
        combo_estado = ttk.Combobox(janela_tarefa, values=estados, state="readonly")
        combo_estado.set(task.get("estado", estados[0]))
        combo_estado.pack(pady=5, fill=tk.X, padx=10)

        tk.Label(janela_tarefa, text="Prioridade:", bg="#2e2e2e", fg="#ffffff").pack(pady=5)
        combo_prioridade = ttk.Combobox(janela_tarefa, values=["Alta", "Média", "Baixa"], state="readonly")
        combo_prioridade.set(task.get("prioridade", "Média"))
        combo_prioridade.pack(pady=5, fill=tk.X, padx=10)

        btn_salvar = tk.Button(janela_tarefa, text="Salvar", command=salvar_tarefa_editada, bg="#007acc", fg="#ffffff", relief="flat")
        btn_salvar.pack(pady=10)
    except Exception as e:
        print(f"[ERROR] Erro ao editar tarefa {task_id}: {e}")
        messagebox.showerror("Erro", f"Erro ao editar tarefa: {e}")

# Função para excluir tarefa
def excluir_tarefa(task_id):
    try:
        if messagebox.askyesno("Confirmar Exclusão", f"Deseja excluir a tarefa {task_id}?"):
            tarefas = carregar_tarefas()
            if task_id in tarefas:
                del tarefas[task_id]
                salvar_tarefas(tarefas)
                enviar_tarefas_planilha()
                atualizar_tarefas()
                messagebox.showinfo("Tarefa Excluída", f"Tarefa {task_id} excluída com sucesso.")
                listar_tarefas_em_execucao()
            else:
                messagebox.showerror("Erro", f"Tarefa {task_id} não encontrada.")
    except Exception as e:
        print(f"[ERROR] Erro ao excluir tarefa {task_id}: {e}")
        messagebox.showerror("Erro", f"Erro ao excluir tarefa: {e}")

# Função para carregar ícone
def carregar_icone_janela():
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    print(f"[DEBUG] Base path for icon: {base_path}")
    icon_path = os.path.join(base_path, 'image', 'icone.ico')
    if os.path.exists(icon_path):
        try:
            icon_image = Image.open(icon_path)
            return ImageTk.PhotoImage(icon_image)
        except Exception as e:
            print(f"[ERROR] Erro ao carregar ícone: {e}")
    print(f"[WARNING] Ícone '{icon_path}' não encontrado.")
    return None

# Função para criar ícone da bandeja
def create_icon():
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    print(f"[DEBUG] Base path for icon: {base_path}")
    icon_path = os.path.join(base_path, 'image', 'icone.ico')
    try:
        if os.path.exists(icon_path):
            icon_image = Image.open(icon_path)
        else:
            print(f"[WARNING] Ícone '{icon_path}' não encontrado. Usando ícone padrão.")
            icon_image = Image.new('RGB', (16, 16), color='black')
    except Exception as e:
        print(f"[ERROR] Erro ao carregar ícone: {e}")
        icon_image = Image.new('RGB', (16, 16), color='black')
    icon = Icon("Task Manager", icon_image, menu=Menu(
        MenuItem('Abrir', lambda icon, item: janela.after(0, janela.deiconify)),
        MenuItem('Sair', lambda icon, item: janela.quit())
    ))
    return icon

# Função para atualizar o layout das colunas
def atualizar_layout_colunas(event=None):
    global resize_timer, layout_lock
    if layout_lock:
        return
    if resize_timer:
        janela.after_cancel(resize_timer)
    resize_timer = janela.after(300, lambda: _atualizar_layout_colunas())

def _atualizar_layout_colunas():
    global layout_lock
    if layout_lock:
        return
    layout_lock = True
    try:
        if not colunas:
            return
        current_width = frame_kanban_canvas.winfo_width() or 800
        if hasattr(_atualizar_layout_colunas, 'last_width') and abs(current_width - _atualizar_layout_colunas.last_width) < 5:
            return
        _atualizar_layout_colunas.last_width = current_width

        total_width = 0
        coluna_width = 150
        num_colunas = len(colunas)

        janela.update_idletasks()
        canvas_height = frame_kanban_canvas.winfo_height() or 600
        scrollbar_height = scrollbar_horizontal.winfo_height() if scrollbar_horizontal.winfo_viewable() else 0
        available_height = canvas_height - scrollbar_height - 10

        for estado, coluna in colunas.items():
            canvas = coluna["canvas"]
            frame_coluna = coluna["frame_coluna"]
            if not (canvas.winfo_exists() and frame_coluna.winfo_exists()):
                continue
            frame_coluna.configure(width=coluna_width)
            canvas.configure(width=coluna_width, height=available_height)
            frame_coluna.pack_forget()
            frame_coluna.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)

            cor_coluna = obter_cor_coluna(estado)
            raio = 20

            canvas.delete("bg_rect", "corner")
            canvas.create_rectangle(raio, 2, coluna_width - raio, available_height - 2,
                                   fill=cor_coluna, outline="", tags="bg_rect")
            canvas.create_rectangle(2, raio, coluna_width - 2, available_height - raio,
                                   fill=cor_coluna, outline="", tags="bg_rect")
            for x, y in [(2, 2), (coluna_width - 2 * raio - 2, 2),
                         (2, available_height - 2 * raio - 2),
                         (coluna_width - 2 * raio - 2, available_height - 2 * raio - 2)]:
                canvas.create_oval(x, y, x + 2 * raio, y + 2 * raio,
                                  fill=cor_coluna, outline="", tags="corner")
            canvas.lower("bg_rect", "corner")

            canvas.configure(scrollregion=canvas.bbox("all"))

            for widget in coluna["frame"].winfo_children():
                if not widget.winfo_exists():
                    continue
                for task_id, task_info in tarefas_widgets.items():
                    if task_info["frame"] == widget:
                        label_tarefa = task_info.get("label_tarefa")
                        if label_tarefa and label_tarefa.winfo_exists():
                            label_tarefa.configure(wraplength=max(50, coluna_width - 30))

            total_width += coluna_width + 10

        total_width = max(total_width, num_colunas * (coluna_width + 10))
        frame_kanban_interno.configure(width=total_width, height=available_height)
        frame_kanban_canvas.configure(scrollregion=(0, 0, total_width, available_height))

        if total_width <= frame_kanban_canvas.winfo_width():
            scrollbar_horizontal.pack_forget()
        else:
            if not scrollbar_horizontal.winfo_viewable():
                scrollbar_horizontal.pack(side=tk.BOTTOM, fill=tk.X)

        texto_detalhes.configure(width=max(40, janela.winfo_width() // 10), height=max(5, janela.winfo_height() // 80))
    except Exception as e:
        print(f"[ERROR] Erro ao atualizar layout das colunas: {e}")
    finally:
        layout_lock = False

# Configuração da janela principal (unchanged)
janela = tk.Tk()
janela.title("Gerenciador de Tarefas - Kanban")
janela.geometry("1200x600")
janela.configure(bg="#2e2e2e")
janela.minsize(800, 400)
icone_janela = carregar_icone_janela()
if icone_janela:
    janela.iconphoto(True, icone_janela)
cor_fundo = "#1e1e1e"
cor_texto = "#ffffff"
cor_lista = "#333333"
cor_destaque = "#007acc"
cor_config = "#00ac47"
style = ttk.Style()
style.theme_use("default")
style.configure("Custom.Horizontal.TScrollbar", troughcolor="#1e1e1e", background="#2e2e2e", arrowcolor="#4a4a4a",
                gripcount=0, width=10)
style.map("Custom.Horizontal.TScrollbar", background=[("active", "#4a4a4a")])
style.configure("Custom.Vertical.TScrollbar", troughcolor="#333333", background="#000000", arrowcolor="#000000",
                borderwidth=0, relief="flat", gripcount=0, width=8)
style.map("Custom.Vertical.TScrollbar", background=[("active", "#4a4a4a")])
frame_principal = tk.Frame(janela, bg="#2e2e2e")
frame_principal.pack(fill=tk.BOTH, expand=True)
frame_lateral = tk.Frame(frame_principal, bg="#252525", width=200)
frame_lateral.pack(side=tk.LEFT, fill=tk.Y)
frame_lateral.pack_propagate(False)
label_lateral = tk.Label(frame_lateral, text="Menu", bg="#252525", fg="#ffffff", font=("Arial", 14, "bold"))
label_lateral.pack(pady=10)
btn_adicionar_lateral = tk.Button(frame_lateral, text="Adicionar Tarefa", command=adicionar_tarefa, bg=cor_destaque,
                                  fg=cor_texto, relief="flat")
btn_adicionar_lateral.pack(fill=tk.X, padx=10, pady=5)
btn_sincronizar_lateral = tk.Button(frame_lateral, text="Sincronizar Dados", command=sincronizar_com_planilha,
                                    bg=cor_destaque, fg=cor_texto, relief="flat")
btn_sincronizar_lateral.pack(fill=tk.X, padx=10, pady=5)
btn_alterar_url_lateral = tk.Button(frame_lateral, text="Alterar URL", command=alterar_url, bg=cor_config, fg=cor_texto,
                                    relief="flat")
btn_alterar_url_lateral.pack(fill=tk.X, padx=10, pady=5)
btn_novo_estado_lateral = tk.Button(frame_lateral, text="Novo Estado", command=adicionar_estado, bg=cor_config,
                                    fg=cor_texto, relief="flat")
btn_novo_estado_lateral.pack(fill=tk.X, padx=10, pady=5)
frame_conteudo = tk.Frame(frame_principal, bg="#2e2e2e")
frame_conteudo.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
frame_kanban = tk.Frame(frame_conteudo, bg="#2e2e2e")
frame_kanban.pack(fill=tk.BOTH, expand=True, pady=10)
frame_kanban_canvas = tk.Canvas(frame_kanban, bg="#2e2e2e", highlightthickness=0)
frame_kanban_canvas.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
scrollbar_horizontal = ttk.Scrollbar(frame_kanban, orient="horizontal", command=frame_kanban_canvas.xview,
                                     style="Custom.Horizontal.TScrollbar")
scrollbar_horizontal.pack(side=tk.BOTTOM, fill=tk.X)
frame_kanban_canvas.configure(xscrollcommand=scrollbar_horizontal.set)
frame_kanban_interno = tk.Frame(frame_kanban_canvas, bg="#2e2e2e")
frame_window = frame_kanban_canvas.create_window((0, 0), window=frame_kanban_interno, anchor="nw")
def ajustar_scrollregion(event=None):
    frame_kanban_canvas.update_idletasks()
    total_width = frame_kanban_interno.winfo_reqwidth()
    total_height = frame_kanban_interno.winfo_reqheight()
    frame_kanban_canvas.configure(scrollregion=(0, 0, total_width, total_height))
    frame_kanban_canvas.itemconfig(frame_window, width=max(total_width, frame_kanban_canvas.winfo_width()))
frame_kanban_interno.bind("<Configure>", ajustar_scrollregion)
frame_kanban_canvas.bind("<Configure>", lambda e: frame_kanban_canvas.itemconfig(
    frame_window, width=max(frame_kanban_interno.winfo_reqwidth(), frame_kanban_canvas.winfo_width())))
def scroll_horizontal(event):
    if event.state & 0x1:
        frame_kanban_canvas.xview_scroll(-1 * (event.delta // 120), "units")
        return "break"
frame_kanban_canvas.bind("<Shift-MouseWheel>", scroll_horizontal)
frame_kanban_canvas.bind("<Shift-Button-4>", lambda e: frame_kanban_canvas.xview_scroll(-1, "units"))
frame_kanban_canvas.bind("<Shift-Button-5>", lambda e: frame_kanban_canvas.xview_scroll(1, "units"))
frame_detalhes = tk.Frame(frame_conteudo, bg="#2e2e2e")
frame_detalhes.pack(fill=tk.X, pady=5)
texto_detalhes = Text(frame_detalhes, bg=cor_lista, fg=cor_texto, borderwidth=0, highlightthickness=0, height=8)
texto_detalhes.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, pady=5, padx=5)
scrollbar_detalhes = ttk.Scrollbar(frame_detalhes, orient="vertical", command=texto_detalhes.yview,
                                   style="Custom.Vertical.TScrollbar")
scrollbar_detalhes.pack(side=tk.RIGHT, fill=tk.Y)
texto_detalhes.configure(yscrollcommand=scrollbar_detalhes.set)
def scroll_detalhes(event):
    if event.delta:
        texto_detalhes.yview_scroll(-1 * (event.delta // 120), "units")
    elif event.num == 4:
        texto_detalhes.yview_scroll(-1, "units")
    elif event.num == 5:
        texto_detalhes.yview_scroll(1, "units")
    return "break"
texto_detalhes.bind("<MouseWheel>", scroll_detalhes)
texto_detalhes.bind("<Button-4>", scroll_detalhes)
texto_detalhes.bind("<Button-5>", scroll_detalhes)
texto_detalhes.tag_configure("titulo", font=("Arial", 12, "bold"))
texto_detalhes.tag_configure("subtitulo", font=("Arial", 10, "bold"))
texto_detalhes.tag_configure("item", font=("Arial", 10))
texto_detalhes.tag_configure("info", font=("Arial", 10, "italic"))

# Função para inicializar a aplicação (unchanged)
def inicializar_aplicacao():
    global estados
    estados = carregar_configuracoes()[1]
    for estado in estados:
        criar_coluna(estado)
    atualizar_tarefas()
    listar_tarefas_em_execucao()
    sincronizar_com_planilha()

# Modificar a inicialização da aplicação (unchanged)
if __name__ == "__main__":
    inicializar_aplicacao()
    janela.bind("<Configure>", atualizar_layout_colunas)
    janela.mainloop()