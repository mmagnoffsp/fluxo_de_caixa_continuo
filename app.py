import csv
import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import subprocess
import sys
import shutil

ARQUIVO_CSV = "fluxo_de_caixa_continuo.csv"
PASTA_BACKUP = "backups"

class AppFluxoCaixa:
    def __init__(self, root):
        self.root = root
        self.root.title("💰 Fluxo de Caixa - Filtros Estilo Excel")
        self.root.geometry("1250x850")
        self.root.minsize(1100, 700)

        self.style = ttk.Style()
        self.style.theme_use("clam")

        # Variáveis
        self.tipo_var = tk.StringVar(value="Entrada")
        self.categoria_var = tk.StringVar()
        self.valor_var = tk.StringVar()
        self.data_var = tk.StringVar(value=datetime.now().strftime("%d/%m/%Y %H:%M"))
        self.descricao_var = tk.StringVar()
        self.obs_var = tk.StringVar()

        # Variáveis dos FILTROS ESTILO EXCEL
        self.filtro_data_inicio = tk.StringVar()
        self.filtro_data_fim = tk.StringVar()
        self.filtro_tipo = tk.StringVar(value="Todos")
        self.filtro_categoria = tk.StringVar(value="Todas")
        self._todos_dados = []  # Guarda todos os registros para filtrar

        self.criar_pasta_backup()
        self.criar_widgets()
        self.carregar_dados()
        self.atualizar_categorias()

    def criar_pasta_backup(self):
        if not os.path.exists(PASTA_BACKUP):
            os.makedirs(PASTA_BACKUP)

    def fazer_backup(self):
        if os.path.exists(ARQUIVO_CSV):
            data_agora = datetime.now().strftime("%Y%m%d_%H%M%S")
            nome_backup = f"{PASTA_BACKUP}/fluxo_{data_agora}.csv"
            shutil.copy2(ARQUIVO_CSV, nome_backup)
            return True
        return False

    def criar_widgets(self):
        # === CABEÇALHO ===
        header_frame = tk.Frame(self.root, bg="#2c3e50", padx=15, pady=12)
        header_frame.pack(fill="x")
        tk.Label(header_frame, text="💰 CONTROLE DE FLUXO DE CAIXA", 
                 font=("Helvetica", 18, "bold"), fg="white", bg="#2c3e50").pack()
        tk.Label(header_frame, text="Filtros por Data, Período, Tipo e Categoria", 
                 font=("Helvetica", 10), fg="#bdc3c7", bg="#2c3e50").pack()

        # === ÁREA DE LANÇAMENTO ===
        form_frame = ttk.LabelFrame(self.root, text="📝 Novo Lançamento")
        form_frame.pack(fill="x", padx=15, pady=5)

        ttk.Label(form_frame, text="Tipo:").grid(row=0, column=0, sticky="w", padx=8, pady=6)
        self.tipo_combo = ttk.Combobox(form_frame, textvariable=self.tipo_var,
                                       values=["Entrada", "Saída"], state="readonly", width=18)
        self.tipo_combo.grid(row=0, column=1, padx=8, pady=6)
        self.tipo_combo.bind("<<ComboboxSelected>>", self.atualizar_categorias)

        ttk.Label(form_frame, text="Categoria:").grid(row=0, column=2, sticky="w", padx=8, pady=6)
        self.cat_combo = ttk.Combobox(form_frame, textvariable=self.categoria_var, state="readonly", width=25)
        self.cat_combo.grid(row=0, column=3, padx=8, pady=6)

        ttk.Label(form_frame, text="Valor R$:").grid(row=0, column=4, sticky="w", padx=8, pady=6)
        ttk.Entry(form_frame, textvariable=self.valor_var, width=18).grid(row=0, column=5, padx=8, pady=6)

        ttk.Label(form_frame, text="Data/Hora:").grid(row=1, column=0, sticky="w", padx=8, pady=6)
        ttk.Entry(form_frame, textvariable=self.data_var, width=18).grid(row=1, column=1, padx=8, pady=6)

        ttk.Label(form_frame, text="Descrição:").grid(row=1, column=2, sticky="w", padx=8, pady=6)
        ttk.Entry(form_frame, textvariable=self.descricao_var, width=40).grid(row=1, column=3, columnspan=2, padx=8, pady=6)

        ttk.Label(form_frame, text="Obs:").grid(row=1, column=5, sticky="w", padx=8, pady=6)
        ttk.Entry(form_frame, textvariable=self.obs_var, width=25).grid(row=1, column=6, padx=8, pady=6)

        btn_form = ttk.Frame(form_frame)
        btn_form.grid(row=2, column=0, columnspan=7, pady=8)
        ttk.Button(btn_form, text="✅ Adicionar", command=self.adicionar_registro).pack(side="left", padx=4)
        ttk.Button(btn_form, text="✏️ Editar", command=self.editar_registro).pack(side="left", padx=4)
        ttk.Button(btn_form, text="🗑️ Excluir", command=self.excluir_registro).pack(side="left", padx=4)

        # === FILTROS ESTILO EXCEL ===
        filtro_frame = ttk.LabelFrame(self.root, text="🔍 FILTROS — Selecione para filtrar e gerar relatório")
        filtro_frame.pack(fill="x", padx=15, pady=5)

        # Filtro de PERÍODO / DATA
        ttk.Label(filtro_frame, text="📅 Data Início:").grid(row=0, column=0, sticky="w", padx=8, pady=6)
        ttk.Entry(filtro_frame, textvariable=self.filtro_data_inicio, width=15).grid(row=0, column=1, padx=5, pady=6)
        ttk.Label(filtro_frame, text="(dd/mm/aaaa)").grid(row=0, column=2, sticky="w", padx=0, pady=6)

        ttk.Label(filtro_frame, text="📅 Data Fim:").grid(row=0, column=3, sticky="w", padx=15, pady=6)
        ttk.Entry(filtro_frame, textvariable=self.filtro_data_fim, width=15).grid(row=0, column=4, padx=5, pady=6)
        ttk.Label(filtro_frame, text="(deixe vazio = todos)").grid(row=0, column=5, sticky="w", padx=0, pady=6)

        # Filtro TIPO
        ttk.Label(filtro_frame, text="Tipo:").grid(row=1, column=0, sticky="w", padx=8, pady=6)
        filt_tipo = ttk.Combobox(filtro_frame, textvariable=self.filtro_tipo,
                                 values=["Todos", "Entrada", "Saída"], state="readonly", width=15)
        filt_tipo.grid(row=1, column=1, padx=5, pady=6)
        filt_tipo.bind("<<ComboboxSelected>>", self.aplicar_filtros)

        # Filtro CATEGORIA
        ttk.Label(filtro_frame, text="Categoria:").grid(row=1, column=3, sticky="w", padx=15, pady=6)
        self.filt_categoria = ttk.Combobox(filtro_frame, textvariable=self.filtro_categoria,
                                            state="readonly", width=22)
        self.filt_categoria.grid(row=1, column=4, padx=5, pady=6)
        self.filt_categoria.bind("<<ComboboxSelected>>", self.aplicar_filtros)

        # Botões dos filtros
        btn_filtro = ttk.Frame(filtro_frame)
        btn_filtro.grid(row=0, column=6, rowspan=2, padx=15, pady=6)
        ttk.Button(btn_filtro, text="🔍 Aplicar Filtro", command=self.aplicar_filtros).pack(fill="x", pady=3)
        ttk.Button(btn_filtro, text="🔄 Limpar Filtros", command=self.limpar_filtros).pack(fill="x", pady=3)
        ttk.Button(btn_filtro, text="📄 Extrato Filtrado", command=self.gerar_extrato_filtrado).pack(fill="x", pady=3)
        ttk.Button(btn_filtro, text="📋 PDF Filtrado", command=self.gerar_pdf_filtrado).pack(fill="x", pady=3)

        # === TABELA ===
        tabela_frame = ttk.Frame(self.root)
        tabela_frame.pack(fill="both", expand=True, padx=15, pady=5)

        colunas = ("Data", "Tipo", "Categoria", "Valor", "Descrição", "Observação")
        self.tabela = ttk.Treeview(tabela_frame, columns=colunas, show="headings", selectmode="browse", height=14)
        self.tabela.bind("<Double-1>", lambda e: self.editar_registro())
        self.tabela.tag_configure("credito", foreground="#27ae60", font=("Helvetica", 9, "bold"))
        self.tabela.tag_configure("debito", foreground="#e74c3c", font=("Helvetica", 9, "bold"))

        for col in colunas:
            self.tabela.heading(col, text=col)
            if col == "Valor":
                self.tabela.column(col, width=130, anchor="e")
            elif col in ["Tipo", "Categoria"]:
                self.tabela.column(col, width=150, anchor="center")
            elif col == "Data":
                self.tabela.column(col, width=160, anchor="center")
            else:
                self.tabela.column(col, width=220, anchor="w")

        scroll = ttk.Scrollbar(tabela_frame, orient="vertical", command=self.tabela.yview)
        self.tabela.configure(yscrollcommand=scroll.set)
        self.tabela.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        # === RESUMO / SALDO ===
        resumo_frame = tk.Frame(self.root, bg="#ecf0f1", padx=15, pady=12)
        resumo_frame.pack(fill="x")

        self.lbl_entradas = tk.Label(resumo_frame, text="📈 Total de Entradas: R$ 0,00",
                                     font=("Helvetica", 11, "bold"), fg="#27ae60", bg="#ecf0f1")
        self.lbl_entradas.pack(anchor="w")

        self.lbl_saidas = tk.Label(resumo_frame, text="📉 Total de Saídas: R$ 0,00",
                                   font=("Helvetica", 11, "bold"), fg="#e74c3c", bg="#ecf0f1")
        self.lbl_saidas.pack(anchor="w")

        self.lbl_saldo = tk.Label(resumo_frame, text="💰 SALDO ATUAL: R$ 0,00",
                                  font=("Helvetica", 14, "bold"), bg="#ecf0f1")
        self.lbl_saldo.pack(anchor="w", pady=5)

        self.lbl_status_filtro = tk.Label(resumo_frame, text="✅ Exibindo: TODOS os registros",
                                          font=("Helvetica", 9, "italic"), fg="#3498db", bg="#ecf0f1")
        self.lbl_status_filtro.pack(anchor="w")

    def atualizar_categorias(self, event=None):
        if self.tipo_var.get() == "Entrada":
            cats = ["Salário", "Rendimento", "Resgate", "Reembolso", "Venda", "Outras Receitas"]
        else:
            cats = ["Moradia", "Alimentação", "Transporte", "Saúde", "Educação",
                    "Lazer", "Cartão", "Empréstimo", "Contas", "Outras Despesas"]
        self.cat_combo['values'] = cats
        if not self.categoria_var.get():
            self.categoria_var.set(cats[0])
        
        # Atualizar filtro de categorias com TODAS as opções
        todas_cats = ["Todas", "Salário", "Rendimento", "Resgate", "Reembolso", "Venda", "Outras Receitas",
                      "Moradia", "Alimentação", "Transporte", "Saúde", "Educação", "Lazer",
                      "Cartão", "Empréstimo", "Contas", "Outras Despesas"]
        self.filt_categoria['values'] = todas_cats

    def _parse_data(self, texto_data):
        """Converte data do formato dd/mm/aaaa hh:mm para objeto datetime"""
        try:
            if " " in texto_data:
                return datetime.strptime(texto_data, "%d/%m/%Y %H:%M")
            else:
                return datetime.strptime(texto_data, "%d/%m/%Y")
        except ValueError:
            return None

    def aplicar_filtros(self, event=None):
        """Aplica todos os filtros e atualiza a tabela"""
        # Limpar tabela
        for linha in self.tabela.get_children():
            self.tabela.delete(linha)

        # Pegar valores dos filtros
        data_inicio_txt = self.filtro_data_inicio.get().strip()
        data_fim_txt = self.filtro_data_fim.get().strip()
        tipo_filtro = self.filtro_tipo.get()
        cat_filtro = self.filtro_categoria.get()

        # Converter datas
        data_inicio = self._parse_data(data_inicio_txt) if data_inicio_txt else None
        data_fim = self._parse_data(data_fim_txt) if data_fim_txt else None
        if data_fim and len(data_fim_txt) <= 10:  # Só data, sem hora → inclui todo o dia
            data_fim = data_fim.replace(hour=23, minute=59, second=59)

        # Filtrar e reinserir
        qtd_mostrada = 0
        for vals, tags in self._todos_dados:
            data_reg = self._parse_data(vals[0])
            if not data_reg:
                continue

            # Filtro de data
            if data_inicio and data_reg < data_inicio:
                continue
            if data_fim and data_reg > data_fim:
                continue

            # Filtro de tipo
            if tipo_filtro != "Todos" and vals[1] != tipo_filtro:
                continue

            # Filtro de categoria
            if cat_filtro != "Todas" and vals[2] != cat_filtro:
                continue

            self.tabela.insert("", "end", values=vals, tags=tags)
            qtd_mostrada += 1

        # Atualizar status
        status = f"✅ Exibindo {qtd_mostrada} registro(s)"
        if data_inicio_txt or data_fim_txt or tipo_filtro != "Todos" or cat_filtro != "Todas":
            status += " — FILTRADO"
        self.lbl_status_filtro.config(text=status)

        self.atualizar_saldo_exibicao()

    def limpar_filtros(self):
        """Remove todos os filtros e mostra tudo"""
        self.filtro_data_inicio.set("")
        self.filtro_data_fim.set("")
        self.filtro_tipo.set("Todos")
        self.filtro_categoria.set("Todas")
        self.aplicar_filtros()

    def calcular_totais_tabela(self):
        """Calcula totais da tabela que está VISÍVEL no momento"""
        entradas = 0.0
        saidas = 0.0
        for linha in self.tabela.get_children():
            vals = self.tabela.item(linha, "values")
            tipo = vals[1]
            val_str = vals[3].replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
            try:
                valor = float(val_str)
                if tipo == "Entrada":
                    entradas += valor
                else:
                    saidas += valor
            except ValueError:
                continue
        return entradas, saidas, entradas - saidas

    def atualizar_saldo_exibicao(self, mensagem_extra=""):
        entradas, saidas, saldo = self.calcular_totais_tabela()

        self.lbl_entradas.config(
            text=f"📈 Total de Entradas: R$ {entradas:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        )
        self.lbl_saidas.config(
            text=f"📉 Total de Saídas: R$ {saidas:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        )

        saldo_fmt = f"💰 SALDO ATUAL: R$ {saldo:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        self.lbl_saldo.config(text=saldo_fmt)
        if saldo >= 0:
            self.lbl_saldo.config(fg="#27ae60")
        else:
            self.lbl_saldo.config(fg="#e74c3c")

    def adicionar_registro(self):
        tipo = self.tipo_var.get()
        cat = self.categoria_var.get()
        data = self.data_var.get().strip() or datetime.now().strftime("%d/%m/%Y %H:%M")
        valor_bruto = self.valor_var.get().strip()
        desc = self.descricao_var.get().strip()
        obs = self.obs_var.get().strip()

        if not valor_bruto or not desc:
            messagebox.showerror("Erro", "Preencha Valor e Descrição!")
            return

        valor_str = valor_bruto.replace("R$", "").strip().replace(".", "").replace(",", ".")
        try:
            valor = float(valor_str)
            if valor <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Erro", "Valor inválido! Ex: 1250,50")
            return

        valor_fmt = f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        tag = "credito" if tipo == "Entrada" else "debito"
        vals = (data, tipo, cat, valor_fmt, desc, obs)

        self.fazer_backup()
        self._todos_dados.append((vals, (tag,)))
        self.salvar_dados_csv()
        self.aplicar_filtros()  # Reaplica filtro e atualiza

        # Limpar
        self.valor_var.set("")
        self.descricao_var.set("")
        self.obs_var.set("")
        self.data_var.set(datetime.now().strftime("%d/%m/%Y %H:%M"))

    def editar_registro(self):
        selecionado = self.tabela.selection()
        if not selecionado:
            messagebox.showinfo("Dica", "Selecione um lançamento na tabela")
            return

        item_id = selecionado[0]
        idx_visivel = self.tabela.index(item_id)
        
        # Encontrar no banco completo
        vals_atuais = self.tabela.item(item_id, "values")
        
        janela = tk.Toplevel(self.root)
        janela.title("✏️ Editar Lançamento")
        janela.geometry("520x450")
        janela.grab_set()

        e_data = tk.StringVar(value=vals_atuais[0])
        e_tipo = tk.StringVar(value=vals_atuais[1])
        e_cat = tk.StringVar(value=vals_atuais[2])
        e_val = tk.StringVar(value=vals_atuais[3].replace("R$", "").strip())
        e_desc = tk.StringVar(value=vals_atuais[4])
        e_obs = tk.StringVar(value=vals_atuais[5])

        ttk.Label(janela, text="Editar Lançamento", font=("Helvetica", 13, "bold")).pack(pady=10)
        frame = ttk.Frame(janela, padding=15)
        frame.pack(fill="both", expand=True)

        campos = [
            ("Data/Hora:", e_data),
            ("Tipo:", e_tipo),
            ("Categoria:", e_cat),
            ("Valor:", e_val),
            ("Descrição:", e_desc),
            ("Observação:", e_obs),
        ]
        for i, (rot, var) in enumerate(campos):
            ttk.Label(frame, text=rot).grid(row=i, column=0, sticky="w", pady=6)
            if rot == "Tipo:":
                ttk.Combobox(frame, textvariable=var, values=["Entrada", "Saída"], state="readonly", width=35).grid(row=i, column=1, pady=6)
            else:
                ttk.Entry(frame, textvariable=var, width=35).grid(row=i, column=1, pady=6)

        def salvar():
            v_str = e_val.get().strip().replace(".", "").replace(",", ".")
            try:
                v_num = float(v_str)
                if v_num <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Erro", "Valor inválido!", parent=janela)
                return

            v_fmt = f"R$ {v_num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            tag = "credito" if e_tipo.get() == "Entrada" else "debito"
            novos_vals = (e_data.get(), e_tipo.get(), e_cat.get(), v_fmt, e_desc.get(), e_obs.get())

            self.fazer_backup()
            # Atualizar na lista mestra
            for i, (dados, _) in enumerate(self._todos_dados):
                if dados == vals_atuais:
                    self._todos_dados[i] = (novos_vals, (tag,))
                    break

            self.salvar_dados_csv()
            self.aplicar_filtros()
            messagebox.showinfo("Sucesso", "Alterações salvas!", parent=janela)
            janela.destroy()

        ttk.Button(frame, text="💾 Salvar", command=salvar).grid(row=6, column=0, columnspan=2, pady=15)

    def excluir_registro(self):
        selecionado = self.tabela.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um lançamento")
            return

        vals_excluir = self.tabela.item(selecionado[0], "values")

        if messagebox.askyesno("Confirmação", "Excluir este lançamento?\nBackup será feito."):
            self.fazer_backup()
            # Remover da lista mestra
            self._todos_dados = [(d, t) for d, t in self._todos_dados if d != vals_excluir]
            self.salvar_dados_csv()
            self.aplicar_filtros()

    def gerar_extrato_filtrado(self):
        """Mostra na tela SOMENTE os dados que estão filtrados"""
        itens = self.tabela.get_children()
        if not itens:
            messagebox.showwarning("Aviso", "Sem dados para exibir — ajuste os filtros")
            return

        janela = tk.Toplevel(self.root)
        janela.title("📄 EXTRATO FILTRADO")
        janela.geometry("1050x700")

        periodo_txt = "PERÍODO: "
        if self.filtro_data_inicio.get():
            periodo_txt += f"De {self.filtro_data_inicio.get()} "
        if self.filtro_data_fim.get():
            periodo_txt += f"Até {self.filtro_data_fim.get()}"
        if not self.filtro_data_inicio.get() and not self.filtro_data_fim.get():
            periodo_txt += "Todos os registros"

        tk.Label(janela, text="📄 EXTRATO FILTRADO", font=("Helvetica", 16, "bold"), fg="#2c3e50").pack(pady=(15,5))
        tk.Label(janela, text=periodo_txt, font=("Helvetica", 11, "bold"), fg="#3498db").pack()
        tk.Label(janela, text=f"Emitido em: {datetime.now().strftime('%d/%m/%Y %H:%M')}", font=("Helvetica", 9), fg="#7f8c8d").pack()

        frame_tab = ttk.Frame(janela)
        frame_tab.pack(fill="both", expand=True, padx=20, pady=15)

        cols = ("Data", "Categoria", "Descrição", "Entrada (+)", "Saída (-)", "Saldo Parcial")
        tree = ttk.Treeview(frame_tab, columns=cols, show="headings", height=15)
        for c in cols:
            tree.heading(c, text=c)
            if c in ["Entrada (+)", "Saída (-)", "Saldo Parcial"]:
                tree.column(c, width=140, anchor="e")
            elif c == "Data":
                tree.column(c, width=160, anchor="center")
            else:
                tree.column(c, width=220, anchor="w")

        scroll = ttk.Scrollbar(frame_tab, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        saldo = 0.0
        total_ent = 0.0
        total_sai = 0.0

        for linha in itens:
            vals = self.tabela.item(linha, "values")
            data, tipo, cat, val_str, desc, obs = vals
            limpo = val_str.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
            try:
                v = float(limpo)
            except:
                v = 0.0
            if tipo == "Entrada":
                saldo += v
                total_ent += v
                col_e = f"+ R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                col_s = "-"
            else:
                saldo -= v
                total_sai += v
                col_e = "-"
                col_s = f"- R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            saldo_fmt = f"R$ {saldo:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            hist = f"{desc} {f'| {obs}' if obs else ''}"
            tree.insert("", "end", values=(data, cat, hist, col_e, col_s, saldo_fmt))

        rodape = ttk.Frame(janela, padding=20)
        rodape.pack(fill="x")

        tk.Label(rodape, text=f"📈 TOTAL DE ENTRADAS: R$ {total_ent:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                 font=("Helvetica", 11, "bold"), fg="#27ae60").pack(anchor="w")
        tk.Label(rodape, text=f"📉 TOTAL DE SAÍDAS: R$ {total_sai:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                 font=("Helvetica", 11, "bold"), fg="#e74c3c").pack(anchor="w")
        tk.Label(rodape, text=f"💰 SALDO FINAL DO PERÍODO: R$ {saldo:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                 font=("Helvetica", 14, "bold"), fg="#27ae60" if saldo >= 0 else "#e74c3c").pack(anchor="w", pady=8)

    def gerar_pdf_filtrado(self):
        """Gera PDF SOMENTE com os dados filtrados na tela"""
        itens = self.tabela.get_children()
        if not itens:
            messagebox.showwarning("Aviso", "Sem dados para gerar PDF — ajuste os filtros")
            return

        # Montar nome do arquivo com período
        data_agora = datetime.now().strftime("%Y%m%d_%H%M%S")
        periodo_nome = ""
        if self.filtro_data_inicio.get():
            periodo_nome = f"_{self.filtro_data_inicio.get().replace('/', '-')}"
        if self.filtro_data_fim.get():
            periodo_nome += f"_a_{self.filtro_data_fim.get().replace('/', '-')}"

        caminho = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf"), ("Todos", "*.*")],
            initialfile=f"extrato_filtrado{periodo_nome}_{data_agora}.pdf",
            title="Salvar PDF Filtrado"
        )
        if not caminho:
            return

        try:
            doc = SimpleDocTemplate(caminho, pagesize=landscape(letter),
                                     leftMargin=25, rightMargin=25, topMargin=25, bottomMargin=25)
            elementos = []
            estilos = getSampleStyleSheet()

            # Título
            titulo = ParagraphStyle("Titulo", parent=estilos["Heading1"], fontSize=18, alignment=1,
                                    textColor=colors.HexColor("#2c3e50"), spaceAfter=6)
            subtitulo = ParagraphStyle("Sub", parent=estilos["Normal"], fontSize=11, alignment=1,
                                        textColor=colors.HexColor("#3498db"), spaceAfter=6)
            rodape_estilo = ParagraphStyle("Rodape", parent=estilos["Normal"], fontSize=9, alignment=1, textColor=colors.grey)

            elementos.append(Paragraph("EXTRATO DE CONTA CORRENTE — FILTRADO", titulo))
            
            periodo_txt = ""
            if self.filtro_data_inicio.get() or self.filtro_data_fim.get():
                if self.filtro_data_inicio.get():
                    periodo_txt += f"De: {self.filtro_data_inicio.get()} "
                if self.filtro_data_fim.get():
                    periodo_txt += f"Até: {self.filtro_data_fim.get()}"
            else:
                periodo_txt = "Todos os registros"
            if self.filtro_tipo.get() != "Todos":
                periodo_txt += f" | Tipo: {self.filtro_tipo.get()}"
            if self.filtro_categoria.get() != "Todas":
                periodo_txt += f" | Categoria: {self.filtro_categoria.get()}"

            elementos.append(Paragraph(periodo_txt, subtitulo))
            elementos.append(Paragraph(f"Emitido em: {datetime.now().strftime('%d/%m/%Y às %H:%M')}", rodape_estilo))
            elementos.append(Spacer(1, 15))

            # Tabela
            cel = ParagraphStyle("Cel", parent=estilos["Normal"], fontSize=9, leading=11)
            cab = ParagraphStyle("Cab", parent=estilos["Normal"], fontSize=10, fontName="Helvetica-Bold",
                                  textColor=colors.whitesmoke, alignment=1)

            dados = [[Paragraph(h, cab) for h in ["Data", "Tipo", "Categoria", "Valor", "Descrição", "Obs"]]]
            for linha in itens:
                vals = self.tabela.item(linha, "values")
                dados.append([Paragraph(str(v), cel) for v in vals])

            tabela = Table(dados, colWidths=[120, 90, 150, 110, 220, 120])
            tabela.setStyle(TableStyle([
                ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#2c3e50")),
                ("ALIGN", (0,0), (-1,-1), "CENTER"),
                ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                ("GRID", (0,0), (-1,-1), 0.5, colors.grey),
                ("BOTTOMPADDING", (0,0), (-1,-1), 8),
                ("TOPPADDING", (0,0), (-1,-1), 8),
            ]))
            elementos.append(tabela)
            elementos.append(Spacer(1, 25))

            # Resumo do período
            ent, sai, saldo = self.calcular_totais_tabela()
            resumo = ParagraphStyle("Resumo", parent=estilos["Normal"], fontSize=12, fontName="Helvetica-Bold", spaceAfter=8)
            cor_saldo = "#27ae60" if saldo >= 0 else "#e74c3c"
            
            elementos.append(Paragraph(f"Total de Entradas no período: R$ {ent:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), resumo))
            elementos.append(Paragraph(f"Total de Saídas no período: R$ {sai:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), resumo))
            elementos.append(Paragraph(f"<font color='{cor_saldo}'>SALDO FINAL DO PERÍODO: R$ {saldo:,.2f}</font>".replace(",", "X").replace(".", ",").replace("X", "."), resumo))

            doc.build(elementos)
            messagebox.showinfo("✅ Sucesso", f"PDF salvo em:\n{caminho}")

            try:
                if sys.platform == "win32":
                    os.startfile(caminho)
            except:
                pass
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao gerar PDF:\n{str(e)}")

    def salvar_dados_csv(self):
        try:
            with open(ARQUIVO_CSV, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(["Data", "Tipo", "Categoria", "Valor", "Descrição", "Observação"])
                for vals, _ in self._todos_dados:
                    w.writerow(vals)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao salvar: {str(e)}")

    def carregar_dados(self):
        self._todos_dados = []
        if os.path.exists(ARQUIVO_CSV):
            try:
                with open(ARQUIVO_CSV, "r", encoding="utf-8") as f:
                    r = csv.reader(f)
                    next(r, None)
                    for linha in r:
                        if linha and len(linha) >= 4:
                            if len(linha) == 5:
                                d, t, v, desc, obs = linha
                                linha = [d, t, "Geral", v, desc, obs]
                            tag = "credito" if linha[1] == "Entrada" else "debito"
                            self._todos_dados.append((tuple(linha), (tag,)))
                self.aplicar_filtros()
            except Exception as e:
                messagebox.showerror("Erro", f"Falha ao carregar: {str(e)}")
        else:
            self.aplicar_filtros()


if __name__ == "__main__":
    root = tk.Tk()
    app = AppFluxoCaixa(root)
    root.mainloop()