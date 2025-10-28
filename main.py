# -*- coding: utf-8 -*-
import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import tempfile
from database import Database
from export_manager import ExportManager
from validate_docbr import CPF
from datetime import datetime
import re


# Cache simples para consultas frequentes
class SimpleCache:
    def __init__(self, max_size=100):
        self.cache = {}
        self.max_size = max_size

    def get(self, key):
        return self.cache.get(key)

    def set(self, key, value):
        if len(self.cache) >= self.max_size:
            self.cache.pop(next(iter(self.cache)))
        self.cache[key] = value

    def clear(self):
        self.cache.clear()


class CRUDApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Cadastro de Serviços")
        self.root.geometry("1200x700")
        self.root.minsize(1000, 600)

        self.db = Database()
        self.export_manager = ExportManager(self.db)
        self.cpf_validator = CPF()
        self.cache = SimpleCache()
        self.endereco_duplicado_id = None
        self.confirmar_duplicidade = False
        self.pagina_atual = 1
        self.itens_por_pagina = 20
        self.total_registros = 0
        self.filtros = {}

        self.style = ttk.Style()
        self.style.configure("TFrame", background="#f0f0f0")
        self.style.configure("TLabel", background="#f0f0f0", font=("Arial", 10))
        self.style.configure("TButton", font=("Arial", 10))
        self.style.configure("Header.TLabel", font=("Arial", 12, "bold"))
        self.style.configure("Title.TLabel", font=("Arial", 14, "bold"))

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.tab_lista = ttk.Frame(self.notebook)
        self.tab_cadastro = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_lista, text="Lista de Serviços")
        self.notebook.add(self.tab_cadastro, text="Cadastro de Serviço")

        self.setup_lista_tab()
        self.setup_cadastro_tab()
        self.carregar_servicos()

    def setup_lista_tab(self):
        frame_topo = ttk.Frame(self.tab_lista)
        frame_topo.pack(fill=tk.X, padx=10, pady=10)
        ttk.Label(frame_topo, text="Lista de Serviços", style="Title.TLabel").pack(side=tk.LEFT, padx=5)

        frame_botoes_topo = ttk.Frame(frame_topo)
        frame_botoes_topo.pack(side=tk.LEFT)
        ttk.Button(frame_botoes_topo, text="Novo Serviço", command=self.novo_servico).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame_botoes_topo, text="Exportar Excel", command=self.exportar_excel).pack(side=tk.LEFT, padx=5)

        frame_filtros = ttk.LabelFrame(self.tab_lista, text="Filtros")
        frame_filtros.pack(fill=tk.X, padx=10, pady=5)

        filtro_frame1 = ttk.Frame(frame_filtros)
        filtro_frame1.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(filtro_frame1, text="Nome:").pack(side=tk.LEFT, padx=5)
        self.filtro_nome = ttk.Entry(filtro_frame1, width=20)
        self.filtro_nome.pack(side=tk.LEFT, padx=5)
        ttk.Label(filtro_frame1, text="CPF:").pack(side=tk.LEFT, padx=5)
        self.filtro_cpf = ttk.Entry(filtro_frame1, width=15)
        self.filtro_cpf.pack(side=tk.LEFT, padx=5)
        ttk.Label(filtro_frame1, text="Status:").pack(side=tk.LEFT, padx=5)
        self.filtro_status = ttk.Combobox(filtro_frame1, width=15, values=('', 'Pendente', 'Concluído', 'Cancelado'))
        self.filtro_status.pack(side=tk.LEFT, padx=5)

        filtro_frame2 = ttk.Frame(frame_filtros)
        filtro_frame2.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(filtro_frame2, text="Bairro:").pack(side=tk.LEFT, padx=5)
        self.filtro_bairro = ttk.Entry(filtro_frame2, width=20)
        self.filtro_bairro.pack(side=tk.LEFT, padx=5)
        ttk.Label(filtro_frame2, text="Rua:").pack(side=tk.LEFT, padx=5)
        self.filtro_rua = ttk.Entry(filtro_frame2, width=20)
        self.filtro_rua.pack(side=tk.LEFT, padx=5)
        ttk.Button(filtro_frame2, text="Filtrar", command=self.aplicar_filtros).pack(side=tk.LEFT, padx=5)
        ttk.Button(filtro_frame2, text="Limpar Filtros", command=self.limpar_filtros).pack(side=tk.LEFT, padx=5)

        frame_tabela = ttk.Frame(self.tab_lista)
        frame_tabela.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        colunas = ("id", "data", "nome", "cpf", "telefone", "endereco", "status")
        self.tabela = ttk.Treeview(frame_tabela, columns=colunas, show="headings")

        headings = {"id": "Protocolo", "data": "Data", "nome": "Nome", "cpf": "CPF", "telefone": "Telefone",
                    "endereco": "Endereço", "status": "Status"}
        widths = {"id": 80, "data": 100, "nome": 200, "cpf": 120, "telefone": 120, "endereco": 300, "status": 100}
        for col, text in headings.items():
            self.tabela.heading(col, text=text)
            self.tabela.column(col, width=widths[col], anchor=tk.CENTER if col in ['id', 'status'] else tk.W)

        scrollbar_y = ttk.Scrollbar(frame_tabela, orient=tk.VERTICAL, command=self.tabela.yview)
        scrollbar_x = ttk.Scrollbar(frame_tabela, orient=tk.HORIZONTAL, command=self.tabela.xview)
        self.tabela.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tabela.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.tabela.bind("<Double-1>", self.abrir_servico)

        frame_paginacao = ttk.Frame(self.tab_lista)
        frame_paginacao.pack(fill=tk.X, padx=10, pady=5)
        ttk.Button(frame_paginacao, text="<<", command=lambda: self.mudar_pagina(1)).pack(side=tk.LEFT, padx=2)
        ttk.Button(frame_paginacao, text="<", command=lambda: self.mudar_pagina(self.pagina_atual - 1)).pack(
            side=tk.LEFT, padx=2)
        self.label_paginacao = ttk.Label(frame_paginacao, text="Página 1 de 1")
        self.label_paginacao.pack(side=tk.LEFT, padx=10)
        ttk.Button(frame_paginacao, text=">", command=lambda: self.mudar_pagina(self.pagina_atual + 1)).pack(
            side=tk.LEFT, padx=2)
        ttk.Button(frame_paginacao, text=">>", command=lambda: self.mudar_pagina(999999)).pack(side=tk.LEFT, padx=2)
        ttk.Label(frame_paginacao, text="Itens por página:").pack(side=tk.LEFT, padx=10)
        self.cb_itens_por_pagina = ttk.Combobox(frame_paginacao, values=["10", "20", "50", "100"], width=5)
        self.cb_itens_por_pagina.set("20")
        self.cb_itens_por_pagina.pack(side=tk.LEFT, padx=5)
        self.cb_itens_por_pagina.bind("<<ComboboxSelected>>", self.alterar_itens_por_pagina)

        frame_acoes = ttk.Frame(self.tab_lista)
        frame_acoes.pack(fill=tk.X, padx=10, pady=10)
        ttk.Button(frame_acoes, text="Editar", command=self.editar_servico).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame_acoes, text="Excluir", command=self.excluir_servico).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame_acoes, text="Gerar PDF", command=self.gerar_pdf).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame_acoes, text="Visualizar PDF", command=self.visualizar_pdf).pack(side=tk.LEFT, padx=5)

    def setup_cadastro_tab(self):
        main_frame = ttk.Frame(self.tab_cadastro)
        main_frame.pack(fill=tk.BOTH, expand=True)
        canvas = tk.Canvas(main_frame, background="#f0f0f0")
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        self.form_frame = ttk.Frame(canvas, style="TFrame")
        self.form_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.form_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        ttk.Label(self.form_frame, text="Cadastro de Serviço", style="Title.TLabel").grid(row=0, column=0, columnspan=4,
                                                                                          pady=10, sticky="w")

        self.vars = {
            "id": tk.StringVar(), "cpf": tk.StringVar(), "nome": tk.StringVar(), "inscricao": tk.StringVar(),
            "telefone": tk.StringVar(), "bairro": tk.StringVar(), "rua": tk.StringVar(), "numero": tk.StringVar(),
            "quadra": tk.StringVar(), "lote": tk.StringVar(), "referencia": tk.StringVar(), "fossas": tk.StringVar(),
            "status": tk.StringVar(value="Pendente"), "chegada": tk.StringVar(), "saida": tk.StringVar(),
            "conclusao": tk.StringVar(), "placa": tk.StringVar(), "motorista": tk.StringVar(),
            "ajudante": tk.StringVar(),
            "observacao": tk.StringVar()
        }

        ttk.Label(self.form_frame, text="Dados do Solicitante", style="Header.TLabel").grid(row=1, column=0,
                                                                                            columnspan=4, pady=(10, 5),
                                                                                            sticky="w")
        ttk.Separator(self.form_frame, orient="horizontal").grid(row=2, column=0, columnspan=4, sticky="ew", pady=5)

        ttk.Label(self.form_frame, text="ID:").grid(row=3, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(self.form_frame, textvariable=self.vars['id'], state="readonly", width=10).grid(row=3, column=1,
                                                                                                  sticky="w", padx=5,
                                                                                                  pady=5)

        ttk.Label(self.form_frame, text="CPF:*").grid(row=4, column=0, sticky="e", padx=5, pady=5)
        self.cpf_entry = ttk.Entry(self.form_frame, textvariable=self.vars['cpf'], width=20)
        self.cpf_entry.grid(row=4, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(self.form_frame, text="Nome:*").grid(row=4, column=2, sticky="e", padx=5, pady=5)
        ttk.Entry(self.form_frame, textvariable=self.vars['nome'], width=40).grid(row=4, column=3, sticky="w", padx=5,
                                                                                  pady=5)

        ttk.Label(self.form_frame, text="Inscrição Municipal:").grid(row=5, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(self.form_frame, textvariable=self.vars['inscricao'], width=20).grid(row=5, column=1, sticky="w",
                                                                                       padx=5, pady=5)

        ttk.Label(self.form_frame, text="Telefone:*").grid(row=5, column=2, sticky="e", padx=5, pady=5)
        self.telefone_entry = ttk.Entry(self.form_frame, textvariable=self.vars['telefone'], width=20)
        self.telefone_entry.grid(row=5, column=3, sticky="w", padx=5, pady=5)

        ttk.Label(self.form_frame, text="Número de Fossas:").grid(row=6, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(self.form_frame, textvariable=self.vars['fossas'], width=10).grid(row=6, column=1, sticky="w", padx=5,
                                                                                    pady=5)

        ttk.Label(self.form_frame, text="Endereço", style="Header.TLabel").grid(row=7, column=0, columnspan=4,
                                                                                pady=(10, 5), sticky="w")
        ttk.Separator(self.form_frame, orient="horizontal").grid(row=8, column=0, columnspan=4, sticky="ew", pady=5)

        ttk.Label(self.form_frame, text="Bairro:*").grid(row=9, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(self.form_frame, textvariable=self.vars['bairro'], width=30).grid(row=9, column=1, sticky="w", padx=5,
                                                                                    pady=5)
        ttk.Label(self.form_frame, text="Rua:*").grid(row=9, column=2, sticky="e", padx=5, pady=5)
        ttk.Entry(self.form_frame, textvariable=self.vars['rua'], width=40).grid(row=9, column=3, sticky="w", padx=5,
                                                                                 pady=5)
        ttk.Label(self.form_frame, text="Número:*").grid(row=10, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(self.form_frame, textvariable=self.vars['numero'], width=10).grid(row=10, column=1, sticky="w",
                                                                                    padx=5, pady=5)
        ttk.Label(self.form_frame, text="Quadra:").grid(row=10, column=2, sticky="e", padx=5, pady=5)
        ttk.Entry(self.form_frame, textvariable=self.vars['quadra'], width=10).grid(row=10, column=3, sticky="w",
                                                                                    padx=5, pady=5)
        ttk.Label(self.form_frame, text="Lote:").grid(row=11, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(self.form_frame, textvariable=self.vars['lote'], width=10).grid(row=11, column=1, sticky="w", padx=5,
                                                                                  pady=5)
        ttk.Label(self.form_frame, text="Referência:").grid(row=11, column=2, sticky="e", padx=5, pady=5)
        ttk.Entry(self.form_frame, textvariable=self.vars['referencia'], width=40).grid(row=11, column=3, sticky="w",
                                                                                        padx=5, pady=5)

        ttk.Label(self.form_frame, text="Status e Informações de Execução", style="Header.TLabel").grid(row=12,
                                                                                                        column=0,
                                                                                                        columnspan=4,
                                                                                                        pady=(10, 5),
                                                                                                        sticky="w")
        ttk.Separator(self.form_frame, orient="horizontal").grid(row=13, column=0, columnspan=4, sticky="ew", pady=5)

        ttk.Label(self.form_frame, text="Status:*").grid(row=14, column=0, sticky="e", padx=5, pady=5)
        status_combo = ttk.Combobox(self.form_frame, textvariable=self.vars['status'], width=15,
                                    values=('Pendente', 'Concluído', 'Cancelado'))
        status_combo.grid(row=14, column=1, sticky="w", padx=5, pady=5)
        status_combo.bind("<<ComboboxSelected>>", self.toggle_execucao_fields)

        self.frame_execucao = ttk.Frame(self.form_frame, style="TFrame")
        self.frame_execucao.grid(row=15, column=0, columnspan=4, sticky="ew", padx=5, pady=5)

        ttk.Label(self.frame_execucao, text="Hora de Chegada:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        self.chegada_entry = ttk.Entry(self.frame_execucao, textvariable=self.vars['chegada'], width=10)
        self.chegada_entry.grid(row=0, column=1, sticky="w", padx=5, pady=5)
        ttk.Label(self.frame_execucao, text="Hora de Saída:").grid(row=0, column=2, sticky="e", padx=5, pady=5)
        self.saida_entry = ttk.Entry(self.frame_execucao, textvariable=self.vars['saida'], width=10)
        self.saida_entry.grid(row=0, column=3, sticky="w", padx=5, pady=5)
        ttk.Label(self.frame_execucao, text="Data de Conclusão:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        self.conclusao_entry = ttk.Entry(self.frame_execucao, textvariable=self.vars['conclusao'], width=15)
        self.conclusao_entry.grid(row=1, column=1, sticky="w", padx=5, pady=5)
        ttk.Label(self.frame_execucao, text="Placa do Veículo:").grid(row=1, column=2, sticky="e", padx=5, pady=5)
        ttk.Entry(self.frame_execucao, textvariable=self.vars['placa'], width=15).grid(row=1, column=3, sticky="w",
                                                                                       padx=5, pady=5)
        ttk.Label(self.frame_execucao, text="Motorista:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(self.frame_execucao, textvariable=self.vars['motorista'], width=30).grid(row=2, column=1, sticky="w",
                                                                                           padx=5, pady=5)
        ttk.Label(self.frame_execucao, text="Ajudante:").grid(row=2, column=2, sticky="e", padx=5, pady=5)
        ttk.Entry(self.frame_execucao, textvariable=self.vars['ajudante'], width=30).grid(row=2, column=3, sticky="w",
                                                                                          padx=5, pady=5)
        ttk.Label(self.frame_execucao, text="Observação:").grid(row=3, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(self.frame_execucao, textvariable=self.vars['observacao'], width=70).grid(row=3, column=1,
                                                                                            columnspan=3, sticky="w",
                                                                                            padx=5, pady=5)

        frame_botoes = ttk.Frame(self.form_frame, style="TFrame")
        frame_botoes.grid(row=16, column=0, columnspan=4, pady=20, sticky="w")
        ttk.Button(frame_botoes, text="Salvar", command=self.salvar_servico).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame_botoes, text="Limpar", command=self.limpar_formulario).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame_botoes, text="Cancelar", command=self.cancelar_edicao).pack(side=tk.LEFT, padx=5)

        self.toggle_execucao_fields()
        self.configurar_validacoes()

        self.progress_frame = ttk.Frame(self.root)
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.progress_frame, variable=self.progress_var, maximum=100)
        self.progress_label = ttk.Label(self.progress_frame, text="Processando...")

    def configurar_validacoes(self):
        self.cpf_entry.bind("<FocusOut>", self.validar_e_formatar_cpf)
        self.telefone_entry.bind("<FocusOut>", self.formatar_telefone)
        self.chegada_entry.bind("<FocusOut>", lambda e, v=self.vars['chegada']: self.formatar_hora(e, v))
        self.saida_entry.bind("<FocusOut>", lambda e, v=self.vars['saida']: self.formatar_hora(e, v))
        self.conclusao_entry.bind("<FocusOut>", self.formatar_data)

    def validar_e_formatar_cpf(self, event=None):
        cpf = self.vars['cpf'].get()
        cpf_limpo = re.sub(r'\D', '', cpf)
        if cpf_limpo:
            if not self.cpf_validator.validate(cpf_limpo):
                messagebox.showerror("Erro de Validação", "CPF inválido!")
                return False
            self.vars['cpf'].set(self.cpf_validator.mask(cpf_limpo))
        return True

    def formatar_telefone(self, event=None):
        tel = re.sub(r'\D', '', self.vars['telefone'].get())
        if tel:
            if len(tel) == 11:
                self.vars['telefone'].set(f"({tel[:2]}) {tel[2:7]}-{tel[7:]}")
            elif len(tel) == 10:
                self.vars['telefone'].set(f"({tel[:2]}) {tel[2:6]}-{tel[6:]}")

    def formatar_hora(self, event, var):
        hora = re.sub(r'\D', '', var.get())
        if len(hora) >= 4:
            var.set(f"{hora[:2]}:{hora[2:4]}")

    def formatar_data(self, event=None):
        data = re.sub(r'\D', '', self.vars['conclusao'].get())
        if len(data) >= 8:
            self.vars['conclusao'].set(f"{data[:2]}/{data[2:4]}/{data[4:8]}")

    def toggle_execucao_fields(self, event=None):
        state = 'normal' if self.vars['status'].get() == "Concluído" else 'disabled'
        for widget in self.frame_execucao.winfo_children():
            if isinstance(widget, (ttk.Entry, ttk.Combobox)):
                widget.config(state=state)

    def run_in_thread(self, target_func, callback, *args):
        def threaded_func():
            result = target_func(*args)
            self.root.after(0, callback, result)

        threading.Thread(target=threaded_func).start()

    def mostrar_progresso(self, mostrar=True, texto="Processando..."):
        if mostrar:
            self.progress_label.config(text=texto)
            self.progress_frame.pack(fill=tk.X, padx=10, pady=5, side=tk.BOTTOM)
            self.progress_bar.pack(fill=tk.X, padx=10, pady=5)
            self.progress_var.set(0)
            self.root.update_idletasks()
        else:
            self.progress_frame.pack_forget()
            self.progress_bar.pack_forget()

    def carregar_servicos(self):
        for item in self.tabela.get_children():
            self.tabela.delete(item)

        servicos, self.total_registros = self.db.listar_servicos(
            filtros=self.filtros, pagina=self.pagina_atual, itens_por_pagina=self.itens_por_pagina
        )
        total_paginas = max(1, (self.total_registros + self.itens_por_pagina - 1) // self.itens_por_pagina)
        self.label_paginacao.config(
            text=f"Página {self.pagina_atual} de {total_paginas} (Total: {self.total_registros} registros)")

        for servico in servicos:
            data = datetime.strptime(servico['data_solicitacao'], '%Y-%m-%d %H:%M:%S').strftime('%d/%m/%Y')
            endereco = f"{servico.get('bairro', '')}, {servico.get('rua', '')}, {servico.get('numero', '')}"
            self.tabela.insert("", "end",
                               values=(servico['id'], data, servico['nome'], servico['cpf'], servico['telefone'],
                                       endereco, servico['status']))

    def aplicar_filtros(self):
        self.filtros = {
            'nome': self.filtro_nome.get().strip(),
            'cpf': self.filtro_cpf.get().strip(),
            'status': self.filtro_status.get().strip(),
            'bairro': self.filtro_bairro.get().strip(),
            'rua': self.filtro_rua.get().strip()
        }
        self.filtros = {k: v for k, v in self.filtros.items() if v}
        self.pagina_atual = 1
        self.carregar_servicos()

    def limpar_filtros(self):
        self.filtro_nome.delete(0, tk.END)
        self.filtro_cpf.delete(0, tk.END)
        self.filtro_status.set("")
        self.filtro_bairro.delete(0, tk.END)
        self.filtro_rua.delete(0, tk.END)
        self.filtros = {}
        self.pagina_atual = 1
        self.carregar_servicos()

    def mudar_pagina(self, pagina):
        total_paginas = max(1, (self.total_registros + self.itens_por_pagina - 1) // self.itens_por_pagina)
        self.pagina_atual = max(1, min(pagina, total_paginas))
        self.carregar_servicos()

    def alterar_itens_por_pagina(self, event=None):
        self.itens_por_pagina = int(self.cb_itens_por_pagina.get())
        self.pagina_atual = 1
        self.carregar_servicos()

    def _get_selected_id(self, action):
        item = self.tabela.selection()
        if not item:
            messagebox.showwarning("Aviso", f"Selecione um serviço para {action}.")
            return None
        return self.tabela.item(item[0], "values")[0]

    def novo_servico(self):
        self.limpar_formulario()
        self.notebook.select(self.tab_cadastro)

    def abrir_servico(self, event=None):
        servico_id = self._get_selected_id("abrir")
        if servico_id:
            self.carregar_servico(servico_id)
            self.notebook.select(self.tab_cadastro)


    def editar_servico(self):
        self.abrir_servico()

    def excluir_servico(self):
        servico_id = self._get_selected_id("excluir")
        if servico_id and messagebox.askyesno("Confirmar Exclusão",
                                              f"Tem certeza que deseja excluir o serviço {servico_id}?"):
            self.db.excluir_servico(servico_id)
            self.cache.clear()
            self.carregar_servicos()
            messagebox.showinfo("Sucesso", "Serviço excluído com sucesso!")

    def _gerar_pdf_callback(self, result):
        self.mostrar_progresso(False)
        success, caminho_arquivo = result
        if success:
            messagebox.showinfo("Sucesso", f"PDF salvo em: {caminho_arquivo}")
            if messagebox.askyesno("Abrir PDF", "Deseja abrir o PDF gerado?"):
                os.startfile(caminho_arquivo)
        else:
            messagebox.showerror("Erro", "Erro ao gerar PDF.")

    def gerar_pdf(self):
        servico_id = self._get_selected_id("gerar PDF")
        if servico_id:
            caminho_arquivo = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")],
                                                           title="Salvar PDF")
            if caminho_arquivo:
                self.mostrar_progresso(True, "Gerando PDF...")
                self.run_in_thread(
                    lambda: (self.export_manager.gerar_pdf(servico_id, caminho_arquivo), caminho_arquivo),
                    self._gerar_pdf_callback)

    def _visualizar_pdf_callback(self, success):
        self.mostrar_progresso(False)
        if not success:
            messagebox.showerror("Erro", "Erro ao visualizar o PDF.")

    def visualizar_pdf(self):
        servico_id = self._get_selected_id("visualizar PDF")
        if servico_id:
            self.mostrar_progresso(True, "Gerando visualização do PDF...")
            self.run_in_thread(self.export_manager.visualizar_pdf, self._visualizar_pdf_callback, servico_id)

    def _exportar_excel_callback(self, result):
        self.mostrar_progresso(False)
        success, caminho_arquivo = result
        if success:
            messagebox.showinfo("Sucesso", f"Excel exportado com sucesso!\nSalvo em: {caminho_arquivo}")
            if messagebox.askyesno("Abrir Excel", "Deseja abrir o arquivo Excel gerado?"):
                os.startfile(caminho_arquivo)
        else:
            messagebox.showerror("Erro", "Erro ao exportar para Excel.")

    def exportar_excel(self):
        caminho_arquivo = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if caminho_arquivo:
            self.mostrar_progresso(True, "Exportando para Excel...")
            self.run_in_thread(
                lambda: (self.export_manager.exportar_excel(caminho_arquivo, self.filtros), caminho_arquivo),
                self._exportar_excel_callback)

    def carregar_servico(self, id_servico):
        servico = self.db.obter_servico(id_servico)
        if servico:
            self.vars['id'].set(servico.get('id', ''))
            self.vars['cpf'].set(servico.get('cpf', ''))
            self.vars['nome'].set(servico.get('nome', ''))
            self.vars['inscricao'].set(servico.get('inscricao_municipal', '') or '')
            self.vars['telefone'].set(servico.get('telefone', ''))
            self.vars['bairro'].set(servico.get('bairro', ''))
            self.vars['rua'].set(servico.get('rua', ''))
            self.vars['numero'].set(servico.get('numero', ''))
            self.vars['quadra'].set(servico.get('quadra', '') or '')
            self.vars['lote'].set(servico.get('lote', '') or '')
            self.vars['referencia'].set(servico.get('referencia', '') or '')
            self.vars['fossas'].set(servico.get('numero_fossas', '') or '')
            self.vars['status'].set(servico.get('status', 'Pendente'))
            self.vars['chegada'].set(servico.get('data_chegada', '') or '')
            self.vars['saida'].set(servico.get('data_saida', '') or '')
            self.vars['conclusao'].set(servico.get('data_conclusao', '') or '')
            self.vars['placa'].set(servico.get('placa_veiculo', '') or '')
            self.vars['motorista'].set(servico.get('motorista', '') or '')
            self.vars['ajudante'].set(servico.get('ajudante', '') or '')
            self.vars['observacao'].set(servico.get('observacao_empresa', '') or '')
            self.toggle_execucao_fields()

    def salvar_servico(self):
        if not self.validar_campos_obrigatorios():
            return

        id_atual = self.vars['id'].get()
        endereco_duplicado_id = self.db.verificar_endereco_duplicado(
            self.vars['bairro'].get(), self.vars['rua'].get(), self.vars['numero'].get(),
            self.vars['quadra'].get(), self.vars['lote'].get(), id_atual
        )

        if endereco_duplicado_id and not self.confirmar_duplicidade:
            if messagebox.askyesno("Endereço Duplicado",
                                   f"Endereço já cadastrado (Protocolo: {endereco_duplicado_id}). Continuar?"):
                self.confirmar_duplicidade = True
            else:
                return

        dados = {
            'cpf': self.vars['cpf'].get(), 'nome': self.vars['nome'].get(),
            'inscricao_municipal': self.vars['inscricao'].get() or None,
            'telefone': self.vars['telefone'].get(), 'bairro': self.vars['bairro'].get(),
            'rua': self.vars['rua'].get(), 'numero': self.vars['numero'].get(),
            'quadra': self.vars['quadra'].get() or None, 'lote': self.vars['lote'].get() or None,
            'referencia': self.vars['referencia'].get() or None,
            'numero_fossas': self.vars['fossas'].get() or None, 'status': self.vars['status'].get()
        }

        if self.vars['status'].get() == "Concluído":
            dados.update({
                'data_chegada': self.vars['chegada'].get() or None, 'data_saida': self.vars['saida'].get() or None,
                'data_conclusao': self.vars['conclusao'].get() or None,
                'placa_veiculo': self.vars['placa'].get() or None,
                'motorista': self.vars['motorista'].get() or None, 'ajudante': self.vars['ajudante'].get() or None,
                'observacao_empresa': self.vars['observacao'].get() or None
            })

        if id_atual:
            self.db.atualizar_servico(id_atual, dados)
            messagebox.showinfo("Sucesso", "Serviço atualizado com sucesso!")
        else:
            servico_id = self.db.inserir_servico(dados)
            messagebox.showinfo("Sucesso", f"Serviço cadastrado com sucesso! Protocolo: {servico_id}")

        self.cache.clear()
        self.carregar_servicos()
        self.notebook.select(self.tab_lista)
        self.limpar_formulario()

    def validar_campos_obrigatorios(self):
        for var_name, field_name in [('cpf', 'CPF'), ('nome', 'Nome'), ('telefone', 'Telefone'), ('bairro', 'Bairro'),
                                     ('rua', 'Rua'), ('numero', 'Número')]:
            if not self.vars[var_name].get().strip():
                messagebox.showerror("Erro", f"O campo {field_name} é obrigatório.")
                return False
        return True

    def limpar_formulario(self):
        for var in self.vars.values():
            var.set("")
        self.vars['status'].set("Pendente")
        self.confirmar_duplicidade = False
        self.endereco_duplicado_id = None
        self.toggle_execucao_fields()

    def cancelar_edicao(self):
        self.limpar_formulario()
        self.notebook.select(self.tab_lista)


if __name__ == "__main__":
    root = tk.Tk()
    app = CRUDApp(root)
    root.mainloop()