import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import pyautogui as pa
import time
import threading

# Configuração inicial do estilo
ctk.set_appearance_mode("Dark")  # "Light", "Dark" ou "System"
ctk.set_default_color_theme("blue")  # Temas: "blue", "green", "dark-blue"

# Variáveis globais
parar_automacao = False
caminho_arquivo = None

def selecionar_arquivo():
    global caminho_arquivo
    caminho = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=[("Arquivos Excel", "*.xlsx")]
    )
    if caminho:
        label_arquivo.configure(text=f"Arquivo selecionado:\n{caminho}")
        btn_iniciar.configure(state="normal")
        caminho_arquivo = caminho

def iniciar_automacao_thread():
    t = threading.Thread(target=iniciar_automacao)
    t.start()

def iniciar_automacao():
    global parar_automacao
    parar_automacao = False

    confirmacao = messagebox.askyesno(
        "Confirmação",
        "Certifique-se de que o sistema Simpas está aberto.\nDeseja iniciar?"
    )
    if not confirmacao:
        return

    try:
        tempo_pause = float(entry_pause.get())
        tempo_sleep = 0.2
        pa.PAUSE = tempo_pause
    except ValueError:
        messagebox.showerror("Erro", "Informe um valor numérico válido para o tempo.")
        return

    try:
        df = pd.read_excel(caminho_arquivo)

        if df.shape[1] < 2:
            messagebox.showerror("Erro", "O arquivo precisa ter ao menos duas colunas.")
            return

        messagebox.showinfo(
            "Atenção",
            "Você tem 5 segundos para colocar o cursor no campo 'Código:' no sistema Simpas."
        )
        time.sleep(5)

        progress_bar.set(0)
        total = len(df)

        for i, row in df.iterrows():
            if parar_automacao:
                messagebox.showinfo("Interrompido", "Automação interrompida pelo usuário.")
                return

            codigo = str(row[0])
            quantidade = str(row[1])

            pa.write(codigo)
            pa.press('tab')
            pa.write(quantidade)
            pa.hotkey('alt', 'i', duration=0.4)
            time.sleep(tempo_sleep)

            progress_bar.set((i + 1) / total)

        messagebox.showinfo("Concluído", "Automação finalizada com sucesso.")
        progress_bar.set(1)

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")

def parar():
    global parar_automacao
    parar_automacao = True

# --------------------
# Interface
# --------------------
janela = ctk.CTk()
janela.title("Automação Simpas com Excel")
janela.geometry("520x480")
janela.resizable(False, False)

# Caixa de instruções
frame_instrucao = ctk.CTkFrame(janela, corner_radius=10)
frame_instrucao.pack(pady=10, padx=10, fill="x")

label_instrucao = ctk.CTkLabel(
    frame_instrucao,
    text=(
        "Passos para usar:\n"
        "1. Selecione o arquivo Excel.\n"
        "2. Ajuste o tempo de lançamento.\n"
        "3. Abra o sistema Simpas.\n"
        "4. Vá até 'Saída por Pedido'.\n"
        "5. Preencha 'Almoxarifado' e 'Área Req.'.\n"
        "6. Clique em 'Iniciar Automação'."
    ),
    justify="left"
)
label_instrucao.pack(pady=10, padx=10)

# Botão de arquivo
btn_arquivo = ctk.CTkButton(janela, text="Selecionar Arquivo", command=selecionar_arquivo)
btn_arquivo.pack(pady=5)

label_arquivo = ctk.CTkLabel(janela, text="Nenhum arquivo selecionado", wraplength=400)
label_arquivo.pack(pady=10)

# Ajuste de tempo
frame_tempos = ctk.CTkFrame(janela, corner_radius=10)
frame_tempos.pack(pady=10)

ctk.CTkLabel(frame_tempos, text="Tempo de lançamento (segundos):").grid(row=0, column=0, padx=5, pady=5)
entry_pause = ctk.CTkEntry(frame_tempos, width=60)
entry_pause.insert(0, "3")
entry_pause.grid(row=0, column=1, padx=5, pady=5)

# Barra de progresso
progress_bar = ctk.CTkProgressBar(janela, width=400)
progress_bar.set(0)
progress_bar.pack(pady=10)

# Botões de ação
btn_iniciar = ctk.CTkButton(
    janela, text="Iniciar Automação", command=iniciar_automacao_thread, state="disabled", fg_color="green"
)
btn_iniciar.pack(pady=5)

btn_parar = ctk.CTkButton(
    janela, text="Parar Automação", command=parar, fg_color="red", hover_color="darkred"
)
btn_parar.pack(pady=5)

janela.mainloop()
