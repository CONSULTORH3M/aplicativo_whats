import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import webbrowser
import pyautogui
import threading
import time
from urllib.parse import quote
import winsound

# Exemplo básico para funcionamento da aba grupo
templates_exemplo = {
    "em treinamento testando o nome grande disso": (
        ", {nome}, da empresa {empresa}, "
        "Como está o uso do *Sistema EvoluTI*? Tudo certo? Algo a relatar? "
        "Prestamos suporte técnico rápido e ativo. Qualquer coisa, pode entrar em contato comigo, o Glaucio, "
        "ou com nosso Suporte Técnico no *(55) 9119 4370* com a (Bruna). "
        "Caso não queira mais receber essa forma de contato, envie *SAIR*."
    ),
    "prospects": (
        ", {nome}, da empresa {empresa}, "
        "Nós da *GDI Informática* trabalhamos com um *Sistema de Gestão: Simples e Prático*. "
        "Montamos a mensalidade baseada nas ferramentas que realmente for utilizar. "
        "Valores a partir de R$ 69,90. Peça mais informações, comigo , o Glaucio. "
        "Se não quiser mais receber informações sobre nossos serviços, envie *SAIR*."
    ),
    "pos": (
        ", {nome}, da empresa {empresa}, "
        "Como está o uso do aplicativo na maquininha de cartões? Tudo certo? Algo a relatar? "
        "Entre em contato para mais informações no (54) 9 9104 1029. "
        "Prestamos suporte técnico rápido e ativo. "
        "Caso não queira mais receber essa forma de contato, envie *SAIR*."
    ),
    "em negociacao": (
        ", {nome}, da empresa {empresa}, "
        "Nós da *GDI Informática* trabalhamos com um *Sistema de Gestão: Simples e Prático*. "
        "Já tinhamos conversado anteriormente, chegamos a falar um pouco sobre o sistema, e até foi sugerido um valor de mensalidade para usar o sistema, " 
        "e agora temos *Novidades* e *Promoções* exclusivas para você. ""Entre em contato comigo, o Glaucio! "
        "Mas Caso não queira mais receber esse tipo de contato, envie *SAIR*."
    ),
    "parceiros": (
        ", {nome}, da empresa {empresa}, "
        "Nós da *GDI Informática* reafirmamos nossa parceria e temos *Novidades* e *Promoções* "
        "exclusivas para seus Clientes nesse novo ciclo. Ferramentas como a Emissão da NF-e, e o Cupom Eletrônico interligado com as Máquinas de Cartões. "
        "Entre em contato comigo, o Glaucio!, para mais informações. "
        "Caso não queira mais receber esse tipo de contato, envie *SAIR*."
    ),
    "contadores": (
        ", {nome}, da empresa {empresa}, "
        "Visitei seu escritório pessoalmente, e agora estamos entrando em contato novamente para apresentar uma "
        "*Solução Prática e Simples* para seus Clientes com relação ao *Sistema de Gestão*; com um Controle de Estoque e Caixa, Além da parte Fiscal, com a Emissão da Nota Eletrônica e o Cupom Eletrônico. "
        "Entre em contato comigo, o Glaucio! "
        "Caso não queira mais receber esse tipo de informação, envie *SAIR*."
    ),
    "smart": (
        ", {nome}, da empresa {empresa}, "
        "Como está o uso do *Sistema EvoluTI*, tudo certo? Sabemos que utiliza a maquininha Smart, com ligação ao sistema de Emissão do Cupom. "
        "E as vezes acontecem alguns problemas, geralmente devido a oxilação da Internet, mas agora temos a opção de ligação direta nas máquinas do Banrisul, Sicredi e Stone. "
        "Se quiser trocar para essa nova forma, ela é bem mais em conta referente a valores mensais, peça mais informações comigo, o Glaucio, ou com nosso suporte no *(55) 9119 4370* (Bruna).  "
        "Caso não queira mais receber essa forma de contato, envie *SAIR*."
    ),
    "evolutiplay": (
        ", {nome}, da empresa {empresa}, "
        "Como está o uso do *Sistema EvoluTI*, tudo certo? "
        "Temos a solução para ligação direta nas máquinas de cartão SMARTs, utilizando Banrisul, Sicredi e Stone, "
        "que atendendem às exigências fiscais atuais. Interligando a emissão do Cupom Eletrônico com as maquininhas. "
        "Entre em contato para mais informações comigo, o Glaucio, ou com nosso suporte no *(55) 9119 4370* (Bruna). "
        "Caso não queira mais receber essa forma de contato, envie *SAIR*."
    ),
    "tef": (
        ", {nome}, da empresa {empresa}, "
        "Como está o uso do *Sistema EvoluTI*, tudo certo? Sabemos que utiliza o TEF na emissão de cupom eletrônico. "
        "Às vezes surgem problemas, principalmente a questão do recebimento dos boletos,  mas agora temos a opção de ligação direta nas máquinas de cartão as SMARTs, temos com o Banrisul, Sicredi e Stone. "
        "Se quiser trocar para essa nova forma, fale comigo, o Glaucio, ou com nosso suporte no *(55) 9119 4370* (Bruna). "
        "Caso não queira mais receber essa forma de contato, envie *SAIR*."
    ),
    "nfe": (
        ", {nome}, da empresa {empresa}, "
        "Como está o uso do *Sistema EvoluTI*, tudo certo? "
        "Sabemos que utiliza principalmente a Emissão de Nota Eletrônica. "
        "Estamos sempre à disposição para ajudar. "
        "pode falar comigo, o Glaucio, ou com nosso suporte no *(55) 9119 4370* (Bruna). "
        "Caso não queira mais receber esse tipo de contato, envie *SAIR*."
    ),
    "pouco suporte": (
        ", {nome}, da empresa {empresa}, "
        "Como está o uso do *Sistema EvoluTI*? "
        "Sabemos que dificilmente precisa de ajuda. "
        "Mas estamos prontos para ajudar, caso precisar de algo. Basta falar com o nosso suporte no *(55) 9119 4370* (Bruna). "
        "Se não quiser mais receber esse tipo de contato, envie *SAIR*."
    ),
    "não usa": (
        ", {nome}, da empresa {empresa}, "
        "Como está o uso do *Sistema EvoluTI*? "
        "Sabemos que utiliza poucas opções e ferramentas. "
        "Mas estamos prontos para ajudar, caso precisar de algo. Basta falar com o nosso suporte no *(55) 9119 4370* (Bruna). "
        "Se não quiser mais receber esse tipo de contato, envie *SAIR*."
    ),
    "usa bem": (
        ", {nome}, da empresa {empresa}, "
        "Como está o uso do *Sistema EvoluTI*, tudo certo? "
        "Além de todas as opções que já utiliza, estamos sempre à disposição para ajudar. Caso precise de suporte basta falar com o nosso Whats no *(55) 9119 4370* (Bruna). "
        "Se não quiser mais receber esse tipo de contato, envie *SAIR*."
    ),
    "mei": (
        ", {nome}, da empresa {empresa}, "
        "Como está o uso do *Sistema EvoluTI*, tudo certo? "
        "Além das opções que já utiliza, temos ferramentas para a parte fiscal, incluindo emissão de NF-e e o Cupom Eletrônico. e Agora com essa questão da Ligação com as maquininhas de Cartões.  "
        "Entre em contato para mais informações comigo, o Glaucio, ou com nosso suporte no *(55) 9119 4370* (Bruna). Se não quiser mais receber esse tipo de contato, envie *SAIR*."
    ),
}


class WhatsAppMessenger:
    def __init__(self, root):
        self.root = root
        self.root.title("Envio de Mensagens Automáticas via WhatsApp_v2")
        self.root.geometry("1000x850")

        self.df_global = None
        self.df_personalizado = None
        self.df_grupo = None
        self.enviando = False

        self.init_gui()

    def cancelar_envio(self):
        if self.enviando:
            self.enviando = False

        if messagebox.askokcancel("Fechar", "Deseja realmente Fechar o Aplicativo?"):
            self.root.destroy()

    def init_gui(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=1, fill="both")

        self.init_aba_personalizado()
        
        self.init_aba_grupo()

    def init_aba_personalizado(self):
        self.aba_personalizado = tk.Frame(self.notebook)
        self.notebook.add(self.aba_personalizado, text="Enviando MENSAGENS e Links dos BOLETOS")

        tk.Label(self.aba_personalizado, text="Planilha Excel Pronta:").pack()
        self.entry_planilha = tk.Entry(self.aba_personalizado, width=80)
        self.entry_planilha.pack()
        tk.Button(self.aba_personalizado, text="Selecionar", command=self.selecionar_planilha).pack(pady=5)

        tk.Label(self.aba_personalizado, text="Link Imagem/Link Padrão(Se tiver aqui vai igual para Todos:").pack()
        self.entry_link = tk.Entry(self.aba_personalizado, width=100)
        self.entry_link.pack()

        self.texto_mensagens = scrolledtext.ScrolledText(self.aba_personalizado, width=100, height=15)
        self.texto_mensagens.pack()

        self.progress_bar = ttk.Progressbar(self.aba_personalizado, length=600, mode="determinate")
        self.progress_bar.pack(pady=5)

        tk.Button(
            self.aba_personalizado,
            text="Carregar Dados",
            command=self.carregar_dados,
            bg="blue",
            fg="white"
        ).pack(pady=5)

        tk.Button(
            self.aba_personalizado,
            text="Disparar Mensagens",
            command=lambda: threading.Thread(target=self.enviar_mensagens).start(),
            bg="green",
            fg="white"
        ).pack(pady=10)

        tk.Button(
            self.aba_personalizado,
            text="Cancelar Envio",
            command=self.cancelar_envio,
            bg="red",
            fg="white"
        ).pack(pady=5)

    
    def init_aba_grupo(self):
        from tkinter import simpledialog

        self.aba_grupo = tk.Frame(self.notebook)
        self.notebook.add(self.aba_grupo, text="MENSAGEM Personalizada por GRUPO")

        frame_top = ttk.Frame(self.aba_grupo)
        frame_top.pack(fill='x', pady=10)

        tk.Label(frame_top, text="Arquivo do Excel:").pack(side='left')
        self.entry_file = tk.Entry(frame_top, width=40)
        self.entry_file.pack(side='left', padx=5)
        ttk.Button(frame_top, text="Selecionar", command=self.select_file).pack(side='left')

        tk.Label(frame_top, text="Planilha:").pack(side='left', padx=(20, 5))
        self.combo_sheet = ttk.Combobox(frame_top, state='readonly', width=20)
        self.combo_sheet.pack(side='left')
        self.combo_sheet.bind("<<ComboboxSelected>>", self.sheet_selected)

        frame_middle = ttk.Frame(self.aba_grupo)
        frame_middle.pack(fill='x', pady=10)

        tk.Label(frame_middle, text="Tipo de Mensagem Disparar:").pack(side='left')
        self.combo_msg_type = ttk.Combobox(frame_middle, state='readonly', width=30, values=list(templates_exemplo.keys()))
        self.combo_msg_type.pack(side='left', padx=5)
        self.combo_msg_type.bind("<<ComboboxSelected>>", self.edit_template)

        tk.Label(frame_middle, text="Grupo/Categoria:").pack(side='left')
        self.combo_group = ttk.Combobox(frame_middle, state='readonly', width=30)
        self.combo_group.pack(side='left', padx=5)

        self.tree = ttk.Treeview(self.aba_grupo, columns=('Telefone', 'Mensagem'), show='headings')
        self.tree.heading('Telefone', text='Telefone')
        self.tree.column('Telefone', width=60, anchor='center')
        self.tree.heading('Mensagem', text='Mensagem')
        self.tree.column('Mensagem', width=500, anchor='w', stretch=True)
        self.tree.pack(fill='both', expand=True, padx=10, pady=10)

        self.txt_template = scrolledtext.ScrolledText(self.aba_grupo, width=100, height=10)
        self.txt_template.pack(padx=10, pady=5)

        style = ttk.Style()
        style.configure("Azul.TButton", foreground="Black", background="#00f7ff")
        style.map("Azul.TButton", background=[("active", "#3399ff")])

        style = ttk.Style()
        style.theme_use('default')

# Botão azul (Carregar Contatos)
        style.configure("Azul.TButton",
                background="#2196F3",
                foreground="white",
                padding=6,
                font=("Arial", 10, "bold"))
        style.map("Azul.TButton",
          background=[("active", "#1976D2")])

# Botão verde (Disparar Mensagens)
        style.configure("Verde.TButton",
                background="#4CAF50",
                foreground="white",
                padding=6,
                font=("Arial", 10, "bold"))
        style.map("Verde.TButton",
          background=[("active", "#388E3C")])

# Botão vermelho (Cancelar Envio)
        style.configure("Vermelho.TButton",
                background="#f44336",
                foreground="white",
                padding=6,
                font=("Arial", 10, "bold"))
        style.map("Vermelho.TButton",
          background=[("active", "#d32f2f")])
        
        # Botão laranja (Exemplo: Excluir, Editar, etc.)
        style.configure("Laranja.TButton",
                background="#FFA500",  # Laranja
                foreground="white",
                padding=6,
                font=("Arial", 10, "bold"))
        style.map("Laranja.TButton",
          background=[("active", "#FF8C00")])  # Laranja escuro no hover
        
        ttk.Button(self.aba_grupo, text="Carregar Contatos",
           command=self.load_messages, style="Azul.TButton").pack(pady=5)

        ttk.Button(self.aba_grupo, text="Disparar Mensagens",
           command=lambda: threading.Thread(target=self.enviar_mensagens_grupo).start(),
           style="Verde.TButton").pack(pady=5)

        ttk.Button(self.aba_grupo, text="Cancelar Envio",
           command=self.cancelar_envio, style="Laranja.TButton").pack(pady=5)

        # Criação do botão
        ttk.Button(
            self.aba_grupo,
            text="Excluir Mensagem Selecionada",
            command=self.excluir_mensagem_selecionada,
            style="Vermelho.TButton"
        ).pack(pady=5)

# Faz o bind da tecla Delete diretamente na Treeview
        self.tree.bind("<Delete>", self.excluir_mensagem_selecionada_evento)

    def excluir_mensagem_selecionada(self, event=None):
        item_selecionado = self.tree.selection()
        if item_selecionado:
            for item in item_selecionado:
                self.tree.delete(item)
        else:
            messagebox.showwarning("Aviso", "Nenhuma mensagem selecionada para excluir.")

    def excluir_mensagem_selecionada_evento(self, event=None):
        self.excluir_mensagem_selecionada()
        self.tree.bind("<Delete>", self.excluir_mensagem_selecionada_evento)

    def selecionar_planilha(self):
        caminho = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if caminho:
            self.entry_planilha.delete(0, tk.END)
            self.entry_planilha.insert(0, caminho)

    def carregar_dados(self):
        caminho = self.entry_planilha.get()
        if not caminho:
            messagebox.showwarning("Aviso", "Selecione uma Planilha.")
            return

        try:
            df = pd.read_excel(caminho)
            if 'MENSAGENS' not in df.columns or 'NUMERO' not in df.columns:
                messagebox.showerror("Erro", "A planilha deve conter as colunas 'MENSAGENS' e 'NUMERO'")
                return

            df = df[df['ENVIAR'].astype(str).str.upper() == 'S']
            self.df_global = df

            self.texto_mensagens.delete(1.0, tk.END)
            for _, row in df.iterrows():
                msg = row['MENSAGENS'].format(**{col.upper(): str(val) for col, val in row.items()})
                self.texto_mensagens.insert(tk.END, f"Para: {row['NUMERO']} - {msg}\n")

            messagebox.showinfo("OK", "Mensagens e Links Carregadas.")
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    def enviar_mensagens(self):
        if self.df_global is None:
            messagebox.showwarning("Aviso", "Carregue uma Planilha Primeiro.")
            return

        self.enviando = True
        total = len(self.df_global)
        self.progress_bar["maximum"] = total

        for i, row in self.df_global.iterrows():
            if not self.enviando:
                break
            numero = row['NUMERO']
            mensagem = row['MENSAGENS'].format(**{col.upper(): str(val) for col, val in row.items()})
            anexo = str(row['ANEXO']) if pd.notna(row.get('ANEXO', '')) else ''
            msg_final = f"{mensagem}\n{anexo}\n{self.entry_link.get()}"
            link = f"https://web.whatsapp.com/send?phone={numero}&text={quote(msg_final)}"
            webbrowser.open(link)
            time.sleep(5)# TESTAR ISSO SE FICAR RAPIDO É DAQUI QUE PEGA
            pyautogui.press('enter')
            time.sleep(5)
            self.progress_bar["value"] = i + 1
            self.root.update()
            if i < total - 1:
                time.sleep(30)

        self.enviando = False
        winsound.Beep(1000, 500)

    def selecionar_planilha_personalizada(self):
        caminho = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if caminho:
            self.entry_planilha_personalizada.delete(0, tk.END)
            self.entry_planilha_personalizada.insert(0, caminho)

    def carregar_dados_personalizados(self):
        try:
            df = pd.read_excel(self.entry_planilha_personalizada.get())
            if 'MENSAGENS' not in df.columns or 'NUMERO' not in df.columns:
                messagebox.showerror("Erro", "A planilha deve conter as colunas 'MENSAGENS' e 'NUMERO'")
                return

            df = df[df['ENVIAR'].astype(str).str.upper() == 'S']
            self.df_personalizado = df
            self.texto_mensagens_personalizadas.delete(1.0, tk.END)
            for _, row in df.iterrows():
                msg = row['MENSAGENS'].format(**{col.upper(): str(val) for col, val in row.items()})
                self.texto_mensagens_personalizadas.insert(tk.END, f"{row['NUMERO']} - {msg}\n")
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    def enviar_mensagens_personalizadas(self):
        if self.df_personalizado is None:
            return
        self.enviando = True
        total = len(self.df_personalizado)
        self.progress_bar2["maximum"] = total
        for i, row in self.df_personalizado.iterrows():
            if not self.enviando:
                break
            numero = row['NUMERO']
            msg = row['MENSAGENS'].format(**{col.upper(): str(val) for col, val in row.items()})
            link = f"https://web.whatsapp.com/send?phone={numero}&text={quote(msg)}"
            webbrowser.open(link)
            time.sleep(10)
            pyautogui.press('enter')
            time.sleep(5)
            self.progress_bar2["value"] = i + 1
            self.root.update()
            if i < total - 1:
                time.sleep(100)
        self.enviando = False
        winsound.Beep(1000, 500)

    def select_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if path:
            self.entry_file.delete(0, tk.END)
            self.entry_file.insert(0, path)
            xls = pd.ExcelFile(path)
            self.combo_sheet['values'] = xls.sheet_names
            if xls.sheet_names:
                self.combo_sheet.current(0)
                self.sheet_selected()

    def sheet_selected(self, _=None):
        file = self.entry_file.get()
        sheet = self.combo_sheet.get()
        self.df_grupo = pd.read_excel(file, sheet_name=sheet)
        if 'grupo' in self.df_grupo.columns:
            grupos = self.df_grupo['grupo'].dropna().unique().tolist()
            self.combo_group['values'] = grupos
            if grupos:
                self.combo_group.current(0)

    def edit_template(self, _=None):
        tipo = self.combo_msg_type.get()
        self.txt_template.delete(1.0, tk.END)
        self.txt_template.insert(tk.END, templates_exemplo.get(tipo, ""))

    def load_messages(self):
        grupo = self.combo_group.get()
        tipo = self.combo_msg_type.get()

        if self.df_grupo is None or not grupo or not tipo:
            return

        template = self.txt_template.get(1.0, tk.END).strip()
        self.tree.delete(*self.tree.get_children())

        df_filtrado = self.df_grupo[self.df_grupo['grupo'] == grupo]

        for _, row in df_filtrado.iterrows():
            nome = row.get('nome', 'Cliente')
            empresa = row.get('empresa', '')
            telefone = row.get('telefone') or row.get('telefone') or ''
            inicio = row.get('inicio', '').strip()  # Pegando a saudação

        # Compondo a mensagem com a saudação + template
            msg_formatada = template.format(nome=nome, empresa=empresa)
            mensagem_final = f"{inicio} {msg_formatada}".strip()

        # Inserindo na Treeview - CORRETAMENTE INDENTADO AGORA
            self.tree.insert('', tk.END, values=(telefone, mensagem_final))

        
    def enviar_mensagens_grupo(self):
        self.enviando = True
        for item in self.tree.get_children():
            if not self.enviando:
                break
            telefone, msg = self.tree.item(item)['values']
            link = f"https://web.whatsapp.com/send?phone={telefone}&text={quote(msg)}"
            webbrowser.open(link)
            time.sleep(10)  # Tempo para carregar a página
            pyautogui.press('enter')  # Envia a mensagem
            time.sleep(110)  # Aguarda 2 minutos antes de ir para o próximo número
        self.enviando = False


if __name__ == "__main__":
    root = tk.Tk()
    app = WhatsAppMessenger(root)
    root.mainloop()
