import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import threading
import os
import sys
from datetime import datetime

# desativa a consciencia automatica de dpi para manter escala consistente
ctk.deactivate_automatic_dpi_awareness()
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import robo

# definicao da paleta de cores para identidade visual da aplicacao
COR = {
    'primaria'      : '#710014',
    'primaria_clara': '#8B0000',   
    'primaria_hover': '#A00000',   
    'cinza_escuro'  : '#1A1A1A',
    'cinza_medio'   : '#333333',
    'fundo_claro'   : '#EAEAEA',   
    'branco'        : '#FFFFFF',
    'texto_log'     : '#1A1A1A',
    
    # cores para sinalizacao de status e resultados
    'status_ok'     : '#2D6A2D',   
    'status_aviso'  : '#EEAA22',   
    'status_erro'   : '#CC0000',   
    'bg_ok'         : '#E8F5E9',   
    'bg_erro'       : '#FFEBEE',   
    'bg_aviso'      : '#FFF8E1',
}

ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

# configuracoes de tipografia padronizadas para a interface
FONTE = {
    'titulo'    : ('Segoe UI', 18, 'bold'), 
    'secao'     : ('Segoe UI', 13, 'bold'),
    'label'     : ('Segoe UI', 11),
    'log'       : ('Consolas', 10),
    'btn_main'  : ('Segoe UI', 12, 'bold'),
}

class RoboApp(ctk.CTk):
    # classe principal que gerencia a janela e o fluxo da aplicacao rpa
    def __init__(self):
        super().__init__()
        self.title("robô de conferência fiscal v1.0")
        self.geometry("1050x760")
        
        # inicializacao de variaveis de controle e caminhos de arquivos
        self.caminho_apuracao = tk.StringVar()
        self.caminho_guias    = tk.StringVar()
        self.pasta_pdfs       = tk.StringVar()
        self.aba_selecionada  = tk.StringVar()
        self.abas_disponiveis = []   
        self.progresso        = tk.DoubleVar(value=0.0)

        self._construir_ui()

    def _construir_ui(self):
        # monta a estrutura visual dividida em navegacao e area operacional
        self.frame_navegacao = ctk.CTkFrame(self, fg_color=COR['cinza_escuro'], width=300, corner_radius=0)
        self.frame_navegacao.pack(side='left', fill='y')
        self.frame_navegacao.pack_propagate(False)

        # banner de identificacao do sistema no menu lateral
        self.frame_id = ctk.CTkFrame(self.frame_navegacao, fg_color='transparent')
        self.frame_id.pack(fill='x', pady=30)
        
        ctk.CTkLabel(self.frame_id, text="CONFERÊNCIA\nFISCAL", font=('Georgia', 20, 'bold'), text_color=COR['branco']).pack(pady=(10, 0))
        
        # secao de metadados anonimizados sobre o contexto de execucao
        info_frame = ctk.CTkFrame(self.frame_navegacao, fg_color='transparent')
        info_frame.pack(fill='x', padx=20, pady=20)
        
        infos = [
            ("🧾", "setor: financeiro / fiscal"),
            ("🏭", "empresa: sua_empresa_exemplo"),
            ("⚙️", "motor: python rpa v1.0"),
        ]
        
        for icone, texto in infos:
            linha = ctk.CTkFrame(info_frame, fg_color='transparent')
            linha.pack(fill='x', pady=10)
            ctk.CTkLabel(linha, text=icone, font=('Segoe UI', 18)).pack(side='left', padx=(0, 10))
            ctk.CTkLabel(linha, text=texto, font=('Segoe UI', 11, 'bold'), text_color=COR['branco']).pack(side='left')

        # painel operacional para interacao com o usuario e log
        self.frame_operacional = ctk.CTkFrame(self, fg_color=COR['fundo_claro'], corner_radius=0)
        self.frame_operacional.pack(side='right', fill='both', expand=True, padx=30, pady=15) 

        self._secao(self.frame_operacional, "1. arquivos e diretorios")
        self._campo_arquivo(self.frame_operacional, "📄 planilha de apuracao:", self.caminho_apuracao, self._selecionar_apuracao)
        self._campo_arquivo(self.frame_operacional, "📄 cadastro de regras:", self.caminho_guias, lambda: self._selecionar_xlsx(self.caminho_guias))
        self._campo_arquivo(self.frame_operacional, "📁 pasta de guias (pdf):", self.pasta_pdfs, self._selecionar_pasta, é_pasta=True)

        self._secao(self.frame_operacional, "2. configuracao do periodo")
        
        # botao de gatilho para inicio do processamento em segundo plano
        self.btn_iniciar = ctk.CTkButton(self.frame_operacional, text="▶  iniciar conferencia", font=FONTE['btn_main'], fg_color=COR['primaria'], command=self._iniciar_conferencia)
        self.btn_iniciar.pack(pady=20)

        # area de saida para feedback em tempo real da execucao
        self.texto_log = ScrolledText(self.frame_operacional, font=FONTE['log'], height=10, state='disabled')
        self.texto_log.pack(fill='both', expand=True)

    def _secao(self, parent, texto):
        # utilitario para criar divisores de secao na interface
        frame = ctk.CTkFrame(parent, fg_color='transparent')
        frame.pack(fill='x', pady=(10, 5))
        ctk.CTkLabel(frame, text=texto, font=FONTE['secao'], text_color=COR['primaria']).pack(side='left')

    def _campo_arquivo(self, parent, label, variavel, comando, é_pasta=False):
        # gera campos de entrada com botoes de selecao de arquivo/diretorio
        frame = ctk.CTkFrame(parent, fg_color='transparent')
        frame.pack(fill='x', pady=5)
        ctk.CTkLabel(frame, text=label, width=280, anchor='w').pack(side='left')
        ctk.CTkButton(frame, text="selecionar", command=comando).pack(side='right')

    def _selecionar_apuracao(self):
        # abre caixa de dialogo para selecao da planilha de apuracao mensal
        caminho = filedialog.askopenfilename(filetypes=[("excel", "*.xlsx *.xlsm")])
        if caminho:
            self.caminho_apuracao.set(caminho)
            self._carregar_abas()

    def _carregar_abas(self):
        # thread secundaria para ler as abas do excel sem travar a interface
        caminho = self.caminho_apuracao.get()
        if not caminho: return
        threading.Thread(target=lambda: robo.listar_abas_recentes(caminho), daemon=True).start()

    def _iniciar_conferencia(self):
        # valida campos e dispara a rotina de conferencia no motor rpa
        if not all([self.caminho_apuracao.get(), self.caminho_guias.get(), self.pasta_pdfs.get()]):
            messagebox.showerror("erro", "preencha todos os campos antes de iniciar.")
            return
        
        self.btn_iniciar.configure(state='disabled', text="processando...")
        threading.Thread(target=self._executar_processo, daemon=True).start()

    def _executar_processo(self):
        # executa a logica de negocio e retorna o feedback para a thread principal
        try:
            res = robo.conferir(
                self.caminho_guias.get(), self.caminho_apuracao.get(),
                self.pasta_pdfs.get(), self.aba_selecionada.get()
            )
            self.after(0, lambda: self._finalizar(res))
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("erro fatal", str(e)))

    def _finalizar(self, resultado):
        # restaura o estado da ui e exibe os resultados ao usuario
        self.btn_iniciar.configure(state='normal', text="▶  iniciar conferencia")
        messagebox.showinfo("concluido", "conferencia finalizada com sucesso.")

if __name__ == "__main__":
    app = RoboApp()
    app.mainloop()