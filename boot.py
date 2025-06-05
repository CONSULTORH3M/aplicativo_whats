import pandas as pd
import pyautogui
import pyperclip
import time
import webbrowser
import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import winsound
from tkinter import ttk
from urllib.parse import quote

df_global = None
df_personalizado = None

# Funções para aba 1 (Envio de Mensagens)
def selecionar_planilha():
    caminho = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
    if caminho:
        entry_planilha.delete(0, tk.END)
        entry_planilha.insert(0, caminho)

def carregar_dados():
    global df_global
    caminho = entry_planilha.get()
    if not caminho:
        messagebox.showwarning("Aviso", "Selecione uma planilha.")
        return

    try:
        df = pd.read_excel(caminho)
        if not {'EMPRESA', 'CONTATO', 'NUMERO', 'TIPO', 'MENSAGENS', 'ANEXO', 'ENVIAR'}.issubset(df.columns):
            messagebox.showerror("Erro", "A planilha deve conter as colunas: EMPRESA, CONTATO, NUMERO, TIPO, MENSAGENS, ANEXO e ENVIAR.")
            return

        df = df[df['ENVIAR'].astype(str).str.strip().str.upper() == 'S']
        df_global = df
        texto_mensagens.delete(1.0, tk.END)

        for index, row in df.iterrows():
            texto_mensagens.insert(tk.END, f"Para: {row['NUMERO']} - MENSAGENS: {row['MENSAGENS']} - ANEXO (link): {row['ANEXO']}\n\n")

        if df.empty:
            messagebox.showinfo("Aviso", "Nenhuma Mensagem marcada como 'S' para enviar.")
        else:
            messagebox.showinfo("Sucesso", "Mensagens carregadas com sucesso.")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao carregar planilha:\n{e}")

def enviar_mensagens():
    global df_global
    if df_global is None:
        messagebox.showwarning("Aviso", "Carregue uma planilha antes de enviar.")
        return

    imagem_global = entry_link.get().strip()
    total = len(df_global)
    progress_bar["maximum"] = total
    progress_bar["value"] = 0
    janela.update()

    for index, row in df_global.iterrows():
        numero = str(row['NUMERO']).strip()

        # Prepara dicionário com todas as colunas da linha para formatação
        row_dict = {col.strip().upper(): str(val) if pd.notna(val) else "" for col, val in row.items()}

        mensagem_template = str(row['MENSAGENS']).strip()
        try:
            mensagem = mensagem_template.format(**row_dict)
        except KeyError:
            mensagem = mensagem_template  # Se faltar alguma chave, usa o texto bruto

        anexo = str(row['ANEXO']).strip() if 'ANEXO' in row and pd.notna(row['ANEXO']) else ""

        mensagem_completa = mensagem
        if anexo:
            mensagem_completa += f"\n{anexo}"
        if imagem_global:
            mensagem_completa += f"\n{imagem_global}"

        mensagem_encoded = quote(mensagem_completa)
        link = f"https://web.whatsapp.com/send?phone={numero}&text={mensagem_encoded}"
        webbrowser.open(link)
        time.sleep(20)
        pyautogui.press('enter')
        time.sleep(5)

        progress_bar["value"] = index + 1
        janela.update()

        if index < total - 1:
            time.sleep(120)

    winsound.Beep(1000, 500)
    winsound.Beep(1200, 500)
    messagebox.showinfo("Concluído", "Mensagens e links enviados com sucesso!")

# ======================== NOVA ABA ========================

def selecionar_planilha_personalizada():
    caminho = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
    if caminho:
        entry_planilha_personalizada.delete(0, tk.END)
        entry_planilha_personalizada.insert(0, caminho)

def carregar_dados_personalizados():
    global df_personalizado
    caminho = entry_planilha_personalizada.get()
    if not caminho:
        messagebox.showwarning("Aviso", "Selecione uma planilha.")
        return

    try:
        df = pd.read_excel(caminho)
        required = {'EMPRESA', 'CONTATO', 'NUMERO', 'TIPO', 'MENSAGENS', 'ENVIAR'}
        if not required.issubset(df.columns):
            messagebox.showerror("Erro", f"A planilha deve conter as colunas: {', '.join(required)}")
            return

        df = df[df['ENVIAR'].astype(str).str.strip().str.upper() == 'S']
        df_personalizado = df
        texto_mensagens_personalizadas.delete(1.0, tk.END)

        for index, row in df.iterrows():
            msg_template = str(row['MENSAGENS'])
            msg_final = msg_template.format(**{col.strip().upper(): str(val) for col, val in row.items()})
            texto_mensagens_personalizadas.insert(tk.END, f"Para: {row['NUMERO']} - {msg_final}\n\n")

        if df.empty:
            messagebox.showinfo("Aviso", "Nenhuma mensagem marcada como 'S' para enviar.")
        else:
            messagebox.showinfo("Sucesso", "Mensagens personalizadas carregadas com sucesso.")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao carregar planilha:\n{e}")

def enviar_mensagens_personalizadas():
    global df_personalizado
    if df_personalizado is None:
        messagebox.showwarning("Aviso", "Carregue uma planilha primeiro.")
        return

    total = len(df_personalizado)
    progress_bar2["maximum"] = total
    progress_bar2["value"] = 0
    janela.update()

    for index, row in df_personalizado.iterrows():
        numero = str(row['NUMERO']).strip()
        msg_template = str(row['MENSAGENS'])
        msg_final = msg_template.format(**{col.strip().upper(): str(val) for col, val in row.items()})

        mensagem_encoded = quote(msg_final)
        link = f"https://web.whatsapp.com/send?phone={numero}&text={mensagem_encoded}"
        webbrowser.open(link)
        time.sleep(20)
        pyautogui.press('enter')
        time.sleep(5)

        progress_bar2["value"] = index + 1
        janela.update()

        if index < total - 1:
            time.sleep(120)

    winsound.Beep(1000, 500)
    winsound.Beep(1200, 500)
    messagebox.showinfo("Concluído", "Mensagens Enviadas Com Sucesso!")

# ======================== INTERFACE GRÁFICA ========================

janela = tk.Tk()
janela.title("Envio de Boletos via WhatsApp")
janela.geometry("750x630")

notebook = ttk.Notebook(janela)
notebook.pack(expand=1, fill="both")

# ABA 1 - ENVIO COM ANEXO
aba_envio = tk.Frame(notebook)
notebook.add(aba_envio, text="Envio de Mensagens")

tk.Label(aba_envio, text="Caminho da Planilha Excel:").pack(pady=5)
entry_planilha = tk.Entry(aba_envio, width=80)
entry_planilha.pack()
tk.Button(aba_envio, text="Selecionar Planilha", command=selecionar_planilha).pack(pady=5)

tk.Label(aba_envio, text="Link da Imagem/Vídeo (Drive, etc):").pack(pady=5)
entry_link = tk.Entry(aba_envio, width=100)
entry_link.pack()

tk.Button(aba_envio, text="Carregar Mensagens", command=carregar_dados).pack(pady=10)

texto_mensagens = scrolledtext.ScrolledText(aba_envio, width=85, height=15)
texto_mensagens.pack(pady=10)

progress_bar = ttk.Progressbar(aba_envio, orient="horizontal", length=600, mode="determinate")
progress_bar.pack(pady=5)

tk.Button(aba_envio, text="Enviar Boletos", command=enviar_mensagens, bg="green", fg="white").pack(pady=10)


# Aba 2
aba_personalizada = tk.Frame(notebook)
notebook.add(aba_personalizada, text="Envio Personalizado")

tk.Label(aba_personalizada, text="Caminho da Planilha Excel:").pack(pady=5)
entry_planilha_personalizada = tk.Entry(aba_personalizada, width=80)
entry_planilha_personalizada.pack()
tk.Button(aba_personalizada, text="Selecionar Planilha", command=selecionar_planilha_personalizada).pack(pady=5)

tk.Button(aba_personalizada, text="Carregar Mensagens", command=carregar_dados_personalizados).pack(pady=10)
texto_mensagens_personalizadas = scrolledtext.ScrolledText(aba_personalizada, width=85, height=15)
texto_mensagens_personalizadas.pack(pady=10)

progress_bar2 = ttk.Progressbar(aba_personalizada, orient="horizontal", length=600, mode="determinate")
progress_bar2.pack(pady=5)

tk.Button(aba_personalizada, text="Enviar Mensagens", command=enviar_mensagens_personalizadas, bg="green", fg="white").pack(pady=10)

tk.Button(janela, text="Fechar", command=janela.destroy, bg="red", fg="white").pack(pady=5)

janela.mainloop()
