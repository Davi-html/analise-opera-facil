import tkinter as tk
from tkinter import filedialog, ttk, messagebox, Listbox, MULTIPLE
import shutil
import os
import platform
import subprocess
from atualizar import executar_macro_atualizar, executar_macro_apresentacao
from main import main as exe_main
from analise_financeiro.financeiro import analise_financeiro

class AbrirPasta:
    @staticmethod
    def abrir(caminho):
        """Abre uma pasta no explorador de arquivos do sistema"""
        if not os.path.isdir(caminho):
            # Tenta criar a pasta se não existir
            try:
                os.makedirs(caminho, exist_ok=True)
            except Exception as e:
                print(f"Erro ao criar pasta: {e}")
                return
        
        sistema = platform.system()
        
        try:
            if sistema == "Windows":
                os.startfile(caminho)
            elif sistema == "Darwin":  # macOS
                subprocess.Popen(["open", caminho])
            else:  # Linux
                subprocess.Popen(["xdg-open", caminho])
        except Exception as e:
            print(f"Erro ao abrir a pasta: {e}")


# Cores modernas
COLORS = {
    'primary': '#3a7ca5',
    'secondary': '#2c3e50',
    'accent': '#1abc9c',
    'light': '#ecf0f1',
    'dark': '#2c3e50',
    'success': '#27ae60',
    'warning': '#f39c12',
    'danger': '#e74c3c',
    'background': '#f0f5ff'
}

# Fontes personalizadas
title_font = ("Segoe UI", 24, "bold")
label_font = ("Segoe UI", 11)
button_font = ("Segoe UI", 10, "bold")
entry_font = ("Segoe UI", 10)

# Configuração da janela principal
root = tk.Tk()
root.title("OPERA FÁCIL - Analisador de Relatórios")
root.geometry("800x600")
root.configure(bg='#f0f5ff')
root.resizable(True, True)

# ========== CRIAR CANVAS COM BARRA DE ROLAGEM ==========
canvas = tk.Canvas(root, bg=COLORS['background'], highlightthickness=0)
canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

scrollbar = ttk.Scrollbar(root, orient="vertical", command=canvas.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

canvas.configure(yscrollcommand=scrollbar.set)

main_frame = tk.Frame(canvas, bg=COLORS['background'], padx=30, pady=20)
canvas_frame = canvas.create_window((0, 0), window=main_frame, anchor="nw")

def update_canvas_width(event):
    canvas.itemconfig(canvas_frame, width=event.width)

def configure_scrollregion(event):
    canvas.configure(scrollregion=canvas.bbox("all"))
    canvas.itemconfig(canvas_frame, width=canvas.winfo_width())

canvas.bind('<Configure>', update_canvas_width)
main_frame.bind("<Configure>", configure_scrollregion)

def on_mousewheel(event):
    canvas.yview_scroll(int(-1*(event.delta/120)), "units")

canvas.bind_all("<MouseWheel>", on_mousewheel)
# ======================================================

# Cabeçalho
header_frame = tk.Frame(main_frame, bg=COLORS['background'])
header_frame.pack(fill='x', pady=(0, 30))

title_label = tk.Label(
    header_frame,
    text="📊 OPERA FÁCIL",
    font=title_font,
    bg=COLORS['background'],
    fg=COLORS['primary']
)
title_label.pack()

subtitle_label = tk.Label(
    header_frame,
    text="Analisador de Relatórios",
    font=("Segoe UI", 12),
    bg=COLORS['background'],
    fg=COLORS['secondary']
)
subtitle_label.pack(pady=(5, 0))

# Frame de instruções
instructions_frame = tk.Frame(
    main_frame,
    bg=COLORS['light'],
    relief='flat',
    borderwidth=1,
    highlightbackground='#d1d8e0',
    highlightthickness=1
)
instructions_frame.pack(fill='x', pady=(0, 25))

instructions_text = """Selecione os 3 arquivos Excel para análise:
1. Neomater
2. Neotin  
3. Pediátrico

Os arquivos serão copiados, renomeados e processados automaticamente."""
instructions_label = tk.Label(
    instructions_frame,
    text=instructions_text,
    font=("Segoe UI", 10),
    bg=COLORS['light'],
    fg=COLORS['dark'],
    justify='left',
    padx=15,
    pady=15
)
instructions_label.pack()

# ========== FRAME PRINCIPAL PARA ARQUIVOS E COMPETÊNCIAS ==========
arquivos_competencias_frame = tk.Frame(main_frame, bg=COLORS['background'])
arquivos_competencias_frame.pack(fill='x', pady=(0, 25))

# Frame para os arquivos (lado esquerdo)
files_frame = tk.Frame(arquivos_competencias_frame, bg=COLORS['background'])
files_frame.pack(side='left', fill='both', expand=True, padx=(0, 15))

# Estilo para os frames de arquivo
def create_file_frame(parent, title, select_command, row):
    frame = tk.Frame(parent, bg=COLORS['light'], relief='flat', padx=15, pady=15)
    frame.grid(row=row, column=0, columnspan=2, sticky='ew', pady=8)
    frame.grid_columnconfigure(0, weight=1)
    
    title_label = tk.Label(
        frame,
        text=title,
        font=("Segoe UI", 12, "bold"),
        bg=COLORS['light'],
        fg=COLORS['primary']
    )
    title_label.grid(row=0, column=0, sticky='w', pady=(0, 10))
    
    entry = tk.Entry(
        frame,
        font=entry_font,
        bg='white',
        fg=COLORS['dark'],
        relief='flat',
        borderwidth=1,
        highlightbackground='#bdc3c7',
        highlightthickness=1,
        highlightcolor=COLORS['accent']
    )
    entry.grid(row=1, column=0, sticky='ew', padx=(0, 10))
    
    button = tk.Button(
        frame,
        text="📁 Procurar",
        command=select_command,
        font=button_font,
        bg=COLORS['primary'],
        fg='white',
        activebackground=COLORS['accent'],
        activeforeground='white',
        relief='flat',
        padx=20,
        pady=5,
        cursor='hand2'
    )
    button.grid(row=1, column=1, padx=(0, 10))

    status_label = tk.Label(
        frame,
        text="❌ Aguardando seleção",
        font=("Segoe UI", 9),
        bg=COLORS['light'],
        fg=COLORS['danger']
    )
    status_label.grid(row=2, column=0, sticky='w', pady=(5, 0))
    
    return frame, entry, status_label, button

# Criar frames para cada arquivo

def select_file1():
    file_path = filedialog.askopenfilename(
        title="Selecione o arquivo Neomater",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")]
    )
    if file_path:
        entry1.delete(0, tk.END)
        entry1.insert(0, file_path)
        status_label1.config(text="✓ Arquivo selecionado", fg=COLORS['success'])

def select_file2():
    file_path = filedialog.askopenfilename(
        title="Selecione o arquivo Neotin",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")]
    )
    if file_path:
        entry2.delete(0, tk.END)
        entry2.insert(0, file_path)
        status_label2.config(text="✓ Arquivo selecionado", fg=COLORS['success'])

def select_file3():
    file_path = filedialog.askopenfilename(
        title="Selecione o arquivo Pediátrico",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")]
    )
    if file_path:
        entry3.delete(0, tk.END)
        entry3.insert(0, file_path)
        status_label3.config(text="✓ Arquivo selecionado", fg=COLORS['success'])

# ========== FRAME DAS COMPETÊNCIAS E ANO (lado direito) ==========

file1_frame, entry1, status_label1, button1 = create_file_frame(files_frame, "📋 NEOMATER", select_file1, 0)
file2_frame, entry2, status_label2, button2 = create_file_frame(files_frame, "📋 NEOTIN", select_file2, 1)
file3_frame, entry3, status_label3, button3 = create_file_frame(files_frame, "📋 PEDIÁTRICO", select_file3, 2)


competencias_frame = tk.Frame(
    arquivos_competencias_frame,
    bg=COLORS['light'],
    relief='flat',
    borderwidth=1,
    highlightbackground='#d1d8e0',
    highlightthickness=1
)
competencias_frame.pack(side='right', fill='both', expand=True)

# Título
competencias_title = tk.Label(
    competencias_frame,
    text="📅 COMPETÊNCIAS",
    font=("Segoe UI", 11, "bold"),
    bg=COLORS['light'],
    fg=COLORS['primary'],
    padx=15,
)
competencias_title.pack(anchor='w', pady=(15, 10))

# Frame para seleção de ano
ano_frame = tk.Frame(competencias_frame, bg=COLORS['light'], padx=15)
ano_frame.pack(fill='x', pady=(0, 15))

ano_label = tk.Label(
    ano_frame,
    text="Selecione o ano:",
    font=("Segoe UI", 10),
    bg=COLORS['light'],
    fg=COLORS['dark']
)
ano_label.pack(side='left', padx=(0, 10))

# Lista de anos (últimos 5 anos + próximo ano)
from datetime import datetime
ano_atual = datetime.now().year
anos = [str(ano_atual + i) for i in range(-2, 3)]  # -2 a +2 anos do atual

# Combobox para seleção de ano
ano_var = tk.StringVar(value=str(ano_atual))
ano_combobox = ttk.Combobox(
    ano_frame,
    textvariable=ano_var,
    values=anos,
    state="readonly",
    font=entry_font,
    width=10
)
ano_combobox.pack(side='left')

# Função para atualizar competências quando ano muda
def atualizar_competencias():
    ano_selecionado = ano_var.get()
    lista_competencias.delete(0, tk.END)
    
    competencias_mensais = [
        f"20/01 a 19/02/{ano_selecionado}",
        f"20/02 a 19/03/{ano_selecionado}",
        f"20/03 a 19/04/{ano_selecionado}",
        f"20/04 a 19/05/{ano_selecionado}",
        f"20/05 a 19/06/{ano_selecionado}",
        f"20/06 a 19/07/{ano_selecionado}",
        f"20/07 a 19/08/{ano_selecionado}",
        f"20/08 a 19/09/{ano_selecionado}",
        f"20/09 a 19/10/{ano_selecionado}",
        f"20/10 a 19/11/{ano_selecionado}",
        f"20/11 a 19/12/{ano_selecionado}",
        f"20/12 a 19/01/{int(ano_selecionado) + 1}",  # Competência que vai para o próximo ano
    ]
    
    for competencia in competencias_mensais:
        lista_competencias.insert(tk.END, competencia)

# Configurar evento quando ano é alterado
ano_var.trace('w', lambda *args: atualizar_competencias())

# Botão para atualizar competências manualmente
btn_atualizar_ano = tk.Button(
    ano_frame,
    text="🔄 Atualizar",
    command=atualizar_competencias,
    font=("Segoe UI", 9),
    bg=COLORS['primary'],
    fg='white',
    relief='flat',
    padx=10,
    pady=2,
    cursor='hand2'
)
btn_atualizar_ano.pack(side='left', padx=(10, 0))

# Frame para lista de competências
lista_container = tk.Frame(competencias_frame, bg=COLORS['light'], padx=15)
lista_container.pack(fill='both', expand=True, pady=(0, 10))

# Scrollbar
lista_scrollbar = tk.Scrollbar(lista_container)
lista_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

# Listbox para competências
lista_competencias = Listbox(
    lista_container,
    yscrollcommand=lista_scrollbar.set,
    selectmode='single',
    bg='white',
    fg=COLORS['dark'],
    font=("Segoe UI", 10),
    relief='flat',
    borderwidth=1,
    highlightbackground='#bdc3c7',
    highlightthickness=1,
    height=8
)
lista_competencias.pack(side=tk.LEFT, fill='both', expand=True)
lista_scrollbar.config(command=lista_competencias.yview)

# Inicializar competências com ano atual
atualizar_competencias()

# Botões para a lista de competências
def limpar_selecao_competencias():
    lista_competencias.selection_clear(0, tk.END)

def obter_selecao_competencias():
    selecionados = lista_competencias.curselection()
    if selecionados:
        competencias = [lista_competencias.get(i) for i in selecionados]
        ano = ano_var.get()
        messagebox.showinfo(
            "Competências Selecionadas",
            f"Ano: {ano}\n\n"
            f"Competências selecionadas ({len(competencias)}):\n\n" + 
            "\n".join(competencias)
        )
    else:
        messagebox.showwarning("Atenção", "Nenhuma competência selecionada!")

# Frame para botões
competencias_botoes_frame = tk.Frame(competencias_frame, bg=COLORS['light'], padx=15)
competencias_botoes_frame.pack(fill='x', pady=(0, 15))

btn_ver_competencias = tk.Button(
    competencias_botoes_frame,
    text="👁️ Ver Selecionadas",
    command=obter_selecao_competencias,
    font=("Segoe UI", 10),
    bg=COLORS['primary'],
    fg='white',
    relief='flat',
    padx=15,
    pady=5,
    cursor='hand2'
)

btn_limpar_competencias = tk.Button(
    competencias_botoes_frame,
    text="🗑️ Limpar Seleção",
    command=limpar_selecao_competencias,
    font=("Segoe UI", 10),
    bg=COLORS['danger'],
    fg='white',
    relief='flat',
    padx=15,
    pady=5,
    cursor='hand2'
)

btn_ver_competencias.pack(side='left', padx=(0, 10))
btn_limpar_competencias.pack(side='left')

# Texto informativo
competencias_info = tk.Label(
    competencias_frame,
    text="Selecione uma ou mais competências para processar.",
    font=("Segoe UI", 9),
    bg=COLORS['light'],
    fg=COLORS['dark'],
    padx=15,
)
competencias_info.pack(anchor='w', pady=(0, 15))
# ======================================================

results_section_frame = tk.Frame(
    main_frame,
    bg=COLORS['light'],
    relief='flat',
    borderwidth=1,
    highlightbackground='#d1d8e0',
    highlightthickness=1
)
results_section_frame.pack(fill='x', pady=(1, 1))

results_title = tk.Label(
    results_section_frame,
    text="📁 ABRIR PASTAS DE RESULTADOS",
    font=("Segoe UI", 11, "bold"),
    bg=COLORS['light'],
    fg=COLORS['primary'],
    padx=15,
)
results_title.pack(anchor='w')

results_buttons_container = tk.Frame(results_section_frame, bg=COLORS['light'], padx=15)
results_buttons_container.pack(fill='x')

def abrir_resultados_neomater():
    caminho = os.path.abspath("./Prestador/neomater/resultado")
    os.makedirs(caminho, exist_ok=True)
    AbrirPasta.abrir(caminho)

def abrir_resultados_neotin():
    caminho = os.path.abspath("./Prestador/neotin/resultado")
    os.makedirs(caminho, exist_ok=True)
    AbrirPasta.abrir(caminho)

def abrir_resultados_prontobaby():
    caminho = os.path.abspath("./Prestador/prontobaby/resultado")
    os.makedirs(caminho, exist_ok=True)
    AbrirPasta.abrir(caminho)

def criar_botao_resultado(parent, text, command, color):
    btn = tk.Button(
        parent,
        text=text,
        command=command,
        font=("Segoe UI", 10),
        bg=color,
        fg='white',
        relief='flat',
        padx=20,
        pady=10,
        cursor='hand2',
        width=15
    )
    
    def on_enter(e):
        btn['background'] = COLORS['accent']
    
    def on_leave(e):
        btn['background'] = color
    
    btn.bind("<Enter>", on_enter)
    btn.bind("<Leave>", on_leave)
    
    return btn

btn_result_neomater = criar_botao_resultado(
    results_buttons_container,
    "Neomater",
    abrir_resultados_neomater,
    "#2c3e50"
)

btn_result_neotin = criar_botao_resultado(
    results_buttons_container,
    "Neotin",
    abrir_resultados_neotin,
    "#34495e"
)

btn_result_prontobaby = criar_botao_resultado(
    results_buttons_container,
    "Pronto Baby",
    abrir_resultados_prontobaby,
    "#7f8c8d"
)

btn_result_neomater.pack(side='left', padx=(0, 10))
btn_result_neotin.pack(side='left', padx=(0, 10))
btn_result_prontobaby.pack(side='left')

results_info = tk.Label(
    results_section_frame,
    text="Clique em qualquer botão acima para abrir a pasta com os relatórios gerados.",
    font=("Segoe UI", 9),
    bg=COLORS['light'],
    fg=COLORS['dark'],
    padx=15,
)
results_info.pack(anchor='w')

# Função de envio (submit)

# Função de envio (submit) - MODIFICADA
def submit():
    file1 = entry1.get()
    file2 = entry2.get()
    file3 = entry3.get()
    
    # Obter competências selecionadas
    selecionados_indices = lista_competencias.curselection()
    
    # Verificar se há competência selecionada
    if not selecionados_indices:
        messagebox.showwarning("Atenção", "Por favor, selecione pelo menos uma competência!")
        return
    
    competencias_selecionadas = [lista_competencias.get(i) for i in selecionados_indices]
    
    # Obter ano selecionado
    ano_selecionado = ano_var.get()
    
    # Pegar apenas a primeira competência selecionada (se quiser todas, modifique)
    competencia_selecionada = competencias_selecionadas[0] if competencias_selecionadas else None
    
    # Mostrar competências selecionadas na confirmação
    competencias_text = "\n".join(competencias_selecionadas) if competencias_selecionadas else "Nenhuma competência selecionada"
    
    # Confirmar com o usuário
    confirm = messagebox.askyesno(
        "Confirmar Processamento",
        f"Deseja processar os arquivos selecionados?\n\n"
        f"• Neomater: {os.path.basename(file1) if file1 else 'Não selecionado'}\n"
        f"• Neotin: {os.path.basename(file2) if file2 else 'Não selecionado'}\n"
        f"• Pediátrico: {os.path.basename(file3) if file3 else 'Não selecionado'}\n\n"
        f"Ano base: {ano_selecionado}\n"
        f"Competências selecionadas ({len(competencias_selecionadas)}):\n{competencias_text}\n\n"
        "Esta operação pode levar alguns minutos."
    )
    
    if not confirm:
        return
    
    submit_button.config(state='disabled', text="Processando...")
    root.update()
    
    try:
        destination_folder = "./separarRelatorio"
        os.makedirs(destination_folder, exist_ok=True)
        
        files_to_process = [
            (file1, "separarNeomater"),
            (file2, "separarNeotin"),
            (file3, "separarPediatrico")
        ]

        from datetime import datetime
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        for original_file, new_name in files_to_process:
            if original_file:
                if not original_file.lower().endswith(('.xlsx', '.xls')):
                    messagebox.showerror("Erro", f"O arquivo {os.path.basename(original_file)} não é um arquivo Excel válido!")
                    submit_button.config(state='normal', text="Processar Arquivos")
                    return
                
                try:
                    ext = os.path.splitext(original_file)[1]
                    new_name_with_timestamp = f"{new_name}"
                    new_path = os.path.join(destination_folder, new_name_with_timestamp + ext)
                    
                    shutil.copy(original_file, new_path)
                    print(f"✅ Arquivo copiado: {os.path.basename(original_file)} -> {new_name_with_timestamp + ext}")
                    
                except Exception as copy_error:
                    messagebox.showerror("Erro", f"Erro ao copiar arquivo:\n{str(copy_error)}")
                    submit_button.config(state='normal', text="Processar Arquivos")
                    return
        
        try:
            # 1. PRIMEIRO processa os arquivos principais
            exe_main()
            
            # 2. DEPOIS chama a análise financeira com a competência selecionada
            try:
                # Extrair apenas a competência (sem o ano) para passar para analise_financeiro
                competencia_sem_ano = competencia_selecionada.split('/')[0] + '/' + competencia_selecionada.split('/')[1] + '/' + competencia_selecionada.split('/')[2]
                
                # Chamar análise financeira para cada prestador
                for prestador in ["Neomater", "Neotin", "Prontobaby"]:
                    try:
                        analise_financeiro(competencia_sem_ano, ano_selecionado)
                        print(f"✅ Análise financeira concluída para {prestador}")
                    except Exception as e:
                        print(f"⚠️ Erro na análise financeira para {prestador}: {str(e)}")
                        # Continua com os outros prestadores mesmo se um falhar
                
                messagebox.showinfo(
                    "Sucesso!",
                    "✅ Processamento concluído com sucesso!\n\n"
                    f"Ano base: {ano_selecionado}\n"
                    f"Competência: {competencia_selecionada}\n"
                    f"Análise financeira executada para todos os prestadores\n"
                    "Os relatórios foram gerados nas pastas:\n"
                    "'relatorios_simplificados' e pastas de resultado"
                )
                
            except Exception as financeiro_error:
                # Se a análise financeira falhar, mostra mensagem mas ainda considera sucesso o processamento principal
                messagebox.showinfo(
                    "Processamento Parcial",
                    f"✅ Processamento principal concluído!\n\n"
                    f"Análise financeira não pôde ser executada:\n{str(financeiro_error)}\n\n"
                    f"Ano base: {ano_selecionado}\n"
                    f"Competência: {competencia_selecionada}"
                )
            
            # Resetar status
            status_label1.config(text="❌ Aguardando seleção", fg=COLORS['danger'])
            status_label2.config(text="❌ Aguardando seleção", fg=COLORS['danger'])
            status_label3.config(text="❌ Aguardando seleção", fg=COLORS['danger'])
            entry1.delete(0, tk.END)
            entry2.delete(0, tk.END)
            entry3.delete(0, tk.END)
            
        except Exception as process_error:
            messagebox.showerror("Erro no Processamento", f"Erro ao processar arquivos:\n{str(process_error)}")
            
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao processar arquivos:\n{str(e)}")
    
    finally:
        submit_button.config(state='normal', text="Processar Arquivos")

def execMacro():
    executar_macro_atualizar()
    executar_macro_apresentacao()
    messagebox.showinfo(
        "Sucesso!",
        "✅ Atualização concluída com sucesso!\n\n"
        "Os relatórios foram atualizados"
    )

action_frame = tk.Frame(main_frame, bg=COLORS['background'])
action_frame.pack(fill='x', pady=(10, 0))

submit_button = tk.Button(
    action_frame,
    text="🚀 Processar Arquivos",
    command=submit,
    font=("Segoe UI", 13, "bold"),
    bg=COLORS['accent'],
    fg='white',
    activebackground='#16a085',
    activeforeground='white',
    relief='flat',
    padx=40,
    pady=15,
    cursor='hand2',
    borderwidth=0
)
submit_button.pack()

submit_enviar = tk.Button(
    action_frame,
    text="🚀 Atualizar dados",
    command=execMacro,
    font=("Segoe UI", 13, "bold"),
    bg=COLORS['warning'],
    fg='white',
    activebackground="#bd8e38",
    activeforeground='white',
    relief='flat',
    padx=40,
    pady=10,
    cursor='hand2',
    borderwidth=0
)
submit_enviar.pack(pady=(25, 0))

footer_frame = tk.Frame(main_frame, bg=COLORS['background'])
footer_frame.pack(fill='x', pady=(25, 0))

footer_label = tk.Label(
    footer_frame,
    text="© 2025 Opera Fácil - Sistema de Análise de Dados Feito por Davi",
    font=("Segoe UI", 9),
    bg=COLORS['background'],
    fg=COLORS['secondary']
)
footer_label.pack()

def on_enter(e):
    e.widget['background'] = '#16a085' if e.widget == submit_button else COLORS['accent']

def on_leave(e):
    e.widget['background'] = COLORS['accent'] if e.widget == submit_button else COLORS['primary']

submit_button.bind("<Enter>", on_enter)
submit_button.bind("<Leave>", on_leave)

for button in [button1, button2, button3]:
    button.bind("<Enter>", on_enter)
    button.bind("<Leave>", on_leave)

for btn in [btn_ver_competencias, btn_limpar_competencias, btn_atualizar_ano]:
    btn.bind("<Enter>", on_enter)
    btn.bind("<Leave>", on_leave)

root.update_idletasks()
width = root.winfo_width()
height = root.winfo_height()
x = (root.winfo_screenwidth() // 2) - (width // 2)
y = (root.winfo_screenheight() // 2) - (height // 2)
root.geometry(f'{width}x{height}+{x}+{y}')

root.mainloop()
