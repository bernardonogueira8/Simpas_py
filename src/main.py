import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import pyautogui as pa
import time
import threading

# Variável de controle para parar a automação
parar_automacao = False

def selecionar_arquivo():
    caminho = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=[("Arquivos Excel", "*.xlsx")]
    )
    if caminho:
        label_arquivo.config(text=f"Arquivo selecionado:\n{caminho}")
        btn_iniciar.config(state=tk.NORMAL)
        global caminho_arquivo
        caminho_arquivo = caminho

def iniciar_automacao_thread():
    t = threading.Thread(target=iniciar_automacao)
    t.start()

def iniciar_automacao():
    global parar_automacao
    parar_automacao = False  # Resetar flag

    confirmacao = messagebox.askyesno("Confirmação", "Certifique-se de que a sistema Simpas está aberto. Deseja iniciar?")
    if not confirmacao:
        return

    try:
        tempo_pause = float(entry_pause.get())
        # tempo_sleep = float(entry_sleep.get())
        tempo_sleep = float(0.2)
        pa.PAUSE = tempo_pause
    except ValueError:
        messagebox.showerror("Erro", "Informe valores numéricos válidos para os tempos.")
        return

    try:
        df = pd.read_excel(caminho_arquivo)

        if df.shape[1] < 2:
            messagebox.showerror("Erro", "Arquivo precisa ter ao menos duas colunas.")
            return

        messagebox.showinfo("Atenção", "Você tem 5 segundos para colocar o cursor no campo 'Código:' no sistema Simpas.")
        time.sleep(5)

        for index, row in df.iterrows():
            if parar_automacao:
                messagebox.showinfo("Interrompido", "Automação foi interrompida pelo usuário.")
                return

            codigo = str(row[0])
            quantidade = str(row[1])

            pa.write(codigo)
            pa.press('tab')
            pa.write(quantidade)
            pa.hotkey('alt', 'i', duration=0.4)
            time.sleep(tempo_sleep)

        messagebox.showinfo("Concluído", "Automação finalizada com sucesso!")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")

def parar():
    global parar_automacao
    parar_automacao = True

# Interface Gráfica
janela = tk.Tk()
janela.title("Automação com Excel")
janela.geometry("420x360")

label_instrucao = tk.Label(janela, text="1. Selecionar o arquivo\n2. Ajustar o tempo de lançamento\n3. Abrir sistema Simpas\n4. Abrir a janela 'Saída por Pedido'\n5. Preencher os campos 'Almoxarifado' e 'Area Req.'\n 6. Clique em 'Iniciar Automação'", pady=10)
label_instrucao.pack()

btn_arquivo = tk.Button(janela, text="Selecionar Arquivo", command=selecionar_arquivo)
btn_arquivo.pack()

label_arquivo = tk.Label(janela, text="Nenhum arquivo selecionado", wraplength=400, pady=10)
label_arquivo.pack()

frame_tempos = tk.Frame(janela)
frame_tempos.pack(pady=10)

tk.Label(frame_tempos, text="Tempo de lançamento em segundos:").grid(row=0, column=0, sticky='e')
entry_pause = tk.Entry(frame_tempos, width=5)
entry_pause.insert(0, "3")
entry_pause.grid(row=0, column=1)

# tk.Label(frame_tempos, text="Sleep (entre linhas):").grid(row=1, column=0, sticky='e')
# entry_sleep = tk.Entry(frame_tempos, width=5)
# entry_sleep.insert(0, "0.2")
# entry_sleep.grid(row=1, column=1)

btn_iniciar = tk.Button(janela, text="Iniciar Automação", command=iniciar_automacao_thread, state=tk.DISABLED)
btn_iniciar.pack(pady=5)

btn_parar = tk.Button(janela, text="Parar Automação", command=parar, fg="white", bg="red")
btn_parar.pack(pady=5)

janela.mainloop()
