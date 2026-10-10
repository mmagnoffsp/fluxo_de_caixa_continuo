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
        self.root.title("💰 Fluxo de Caixa - Contas e Poupanças")
        self.root.geometry("1280x900")
        self.root.minsize(1100, 700)
        self.style = ttk.Style()
        self.style.theme_use("clam")
        
        self.tipo_var = tk.StringVar(value="Entrada")
        self.categoria_var = tk.StringVar()
        self.valor_var = tk.StringVar()
        self.data_var = tk.StringVar(value=datetime.now().strftime("%d/%m/%Y %H:%M"))
        self.descricao_var = tk.StringVar()
        self.obs_var = tk.StringVar()
        self.conta_var = tk.StringVar()
        self.filtro_data_inicio = tk.StringVar()
        self.filtro_data_fim = tk.StringVar()
        self.filtro_tipo = tk.StringVar(value="Todos")
        self.filtro_categoria = tk.StringVar(value="Todas")
        
        self.conta_atual = None  # None = TODAS as contas
        self._todos_dados = []
        
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

    def obter_lista_contas(self):
        """Retorna contas únicas encontradas nos dados"""
        contas = set()
        for vals, _ in self._todos_dados:
            if len(vals) >= 7 and vals[6].strip():
                contas.add(vals[6].strip())
        return sorted(list(contas))

    def formatar_valor(self, valor):
        return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    def calcular_saldo_conta(self, conta_filtro):
        """Calcula saldo de uma conta ou de todas"""
        entradas = saidas = 0.0
        for vals, _ in self._todos_dados:
            if conta_filtro is not None:
                if len(vals) < 7 or vals[6] != conta_filtro:
                    continue
            try:
                val_str = vals[3].replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
                valor = float(val_str)
                if vals[1] == "Entrada":
                    entradas += valor
                else:
                    saidas += valor
            except:
                pass
        return entradas - saidas

    def criar_widgets(self):
        # Cabeçalho
        header_frame = tk.Frame(self.root, bg="#2c3e50", padx=15, pady=12)
        header_frame.pack(fill="x")
        tk.Label(header_frame, text="💰 CONTROLE DE FLUXO DE CAIXA", 
                 font=("Helvetica", 18, "bold"), fg="white", bg="#2c3e50").pack()
        tk.Label(header_frame, text="Clique em uma conta para filtrar os lançamentos", 
                 font=("Helvetica", 10), fg="#bdc3c7", bg="#2c3e50").pack()

        # === BOTÕES DAS CONTAS ===
        self.frame_botoes_contas = tk.LabelFrame(
            self.root,
            text="🏦 SUAS CONTAS / POUPANÇAS — Clique para visualizar",
            padx=10, pady=8,
            font=("Helvetica", 11, "bold")
        )
        self.frame_botoes_contas.pack(fill="x", padx=15, pady=(10, 5))

        # Título do filtro ativo
        self.lbl_conta_selecionada = tk.Label(
            self.root,
            text="📊 EXTRATO GERAL — Todas as contas",
            font=("Helvetica", 14, "bold"),
            bg="#d4edda", fg="#27ae60", padx=15, pady=10
        )
        self.lbl_conta_selecionada.pack(fill="x", padx=15, pady=(0, 5))

        # Formulário de lançamento
        form_frame = ttk.LabelFrame(self.root, text="📝 Novo Lançamento")
        form_frame.pack(fill="x", padx=15, pady=5)

        ttk.Label(form_frame, text="Conta/Poupança:").grid(row=0, column=0, sticky="w", padx=8, pady=6)
        self.conta_combo = ttk.Combobox(form_frame, textvariable=self.conta_var, width=30)
        self.conta_combo.grid(row=0, column=1, padx=8, pady=6)
        ttk.Label(form_frame, text="(deixe vazio = Geral)").grid(row=0, column=2, sticky="w", padx=0, pady=6)

        ttk.Label(form_frame, text="Tipo:").grid(row=1, column=0, sticky="w", padx=8, pady=6)
        self.tipo_combo = ttk.Combobox(form_frame, textvariable=self.tipo_var,
                                       values=["Entrada", "Saída"], state="readonly", width=18)
        self.tipo_combo.grid(row=1, column=1, padx=8, pady=6)
        self.tipo_combo.bind("<<ComboboxSelected>>", self.atualizar_categorias)

        ttk.Label(form_frame, text="Categoria:").grid(row=1, column=3, sticky="w", padx=8, pady=6)
        self.cat_combo = ttk.Combobox(form_frame, textvariable=self.categoria_var, state="readonly", width=35)
        self.cat_combo.grid(row=1, column=4, padx=8, pady=6)

        ttk.Label(form_frame, text="Valor R$:").grid(row=1, column=5, sticky="w", padx=8, pady=6)
        ttk.Entry(form_frame, textvariable=self.valor_var, width=18).grid(row=1, column=6, padx=8, pady=6)

        ttk.Label(form_frame, text="Data/Hora:").grid(row=2, column=0, sticky="w", padx=8, pady=6)
        ttk.Entry(form_frame, textvariable=self.data_var, width=18).grid(row=2, column=1, padx=8, pady=6)

        ttk.Label(form_frame, text="Descrição:").grid(row=2, column=3, sticky="w", padx=8, pady=6)
        ttk.Entry(form_frame, textvariable=self.descricao_var, width=40).grid(row=2, column=4, padx=8, pady=6)

        ttk.Label(form_frame, text="Obs:").grid(row=2, column=5, sticky="w", padx=8, pady=6)
        ttk.Entry(form_frame, textvariable=self.obs_var, width=25).grid(row=2, column=6, padx=8, pady=6)

        btn_form = ttk.Frame(form_frame)
        btn_form.grid(row=3, column=0, columnspan=7, pady=8)
        ttk.Button(btn_form, text="✅ Adicionar", command=self.adicionar_registro).pack(side="left", padx=4)
        ttk.Button(btn_form, text="✏️ Editar", command=self.editar_registro).pack(side="left", padx=4)
        ttk.Button(btn_form, text="🗑️ Excluir", command=self.excluir_registro).pack(side="left", padx=4)

        # Filtros
        filtro_frame = ttk.LabelFrame(self.root, text="🔍 FILTROS — Selecione para filtrar e gerar relatório")
        filtro_frame.pack(fill="x", padx=15, pady=5)

        ttk.Label(filtro_frame, text="📅 Data Início:").grid(row=0, column=0, sticky="w", padx=8, pady=6)
        ttk.Entry(filtro_frame, textvariable=self.filtro_data_inicio, width=15).grid(row=0, column=1, padx=5, pady=6)
        ttk.Label(filtro_frame, text="📅 Data Fim:").grid(row=0, column=2, sticky="w", padx=15, pady=6)
        ttk.Entry(filtro_frame, textvariable=self.filtro_data_fim, width=15).grid(row=0, column=3, padx=5, pady=6)

        ttk.Label(filtro_frame, text="Tipo:").grid(row=0, column=4, sticky="w", padx=15, pady=6)
        filt_tipo = ttk.Combobox(filtro_frame, textvariable=self.filtro_tipo,
                                 values=["Todos", "Entrada", "Saída"], state="readonly", width=15)
        filt_tipo.grid(row=0, column=5, padx=5, pady=6)
        filt_tipo.bind("<<ComboboxSelected>>", self.aplicar_filtros)

        ttk.Label(filtro_frame, text="Categoria:").grid(row=1, column=0, sticky="w", padx=8, pady=6)
        self.filt_categoria = ttk.Combobox(filtro_frame, textvariable=self.filtro_categoria,
                                            state="readonly", width=35)
        self.filt_categoria.grid(row=1, column=1, padx=5, pady=6)
        self.filt_categoria.bind("<<ComboboxSelected>>", self.aplicar_filtros)

        btn_filtro = ttk.Frame(filtro_frame)
        btn_filtro.grid(row=0, column=6, rowspan=2, padx=15, pady=6)
        ttk.Button(btn_filtro, text="🔄 Ver Todas as Contas", command=self.ver_todas_contas).pack(fill="x", pady=3)
        ttk.Button(btn_filtro, text="📄 Extrato", command=self.gerar_extrato_filtrado).pack(fill="x", pady=3)
        ttk.Button(btn_filtro, text="📋 Gerar PDF", command=self.gerar_pdf_filtrado).pack(fill="x", pady=3)

        # Tabela
        tabela_frame = ttk.Frame(self.root)
        tabela_frame.pack(fill="both", expand=True, padx=15, pady=5)
        colunas = ("Data", "Conta", "Tipo", "Categoria", "Valor", "Descrição", "Observação")
        self.tabela = ttk.Treeview(tabela_frame, columns=colunas, show="headings", selectmode="browse", height=12)
        self.tabela.bind("<Double-1>", lambda e: self.editar_registro())
        self.tabela.tag_configure("credito", foreground="#27ae60", font=("Helvetica", 9, "bold"))
        self.tabela.tag_configure("debito", foreground="#e74c3c", font=("Helvetica", 9, "bold"))

        for col in colunas:
            self.tabela.heading(col, text=col)
            if col == "Valor":
                self.tabela.column(col, width=130, anchor="e")
            elif col == "Conta":
                self.tabela.column(col, width=200, anchor="w")
            elif col in ["Tipo", "Categoria"]:
                self.tabela.column(col, width=180, anchor="center")
            elif col == "Data":
                self.tabela.column(col, width=160, anchor="center")
            else:
                self.tabela.column(col, width=220, anchor="w")

        scroll = ttk.Scrollbar(tabela_frame, orient="vertical", command=self.tabela.yview)
        self.tabela.configure(yscrollcommand=scroll.set)
        self.tabela.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        # Resumo
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

    def atualizar_botoes_contas(self):
        """Recria os botões de todas as contas com seus saldos"""
        for w in self.frame_botoes_contas.winfo_children():
            w.destroy()

        # Botão Geral
        saldo_geral = self.calcular_saldo_conta(None)
        cor_geral = "#27ae60" if saldo_geral >= 0 else "#e74c3c"
        tk.Button(
            self.frame_botoes_contas,
            text=f"📊 GERAL\n{self.formatar_valor(saldo_geral)}",
            bg=cor_geral, fg="white", font=("Helvetica", 10, "bold"),
            padx=15, pady=10, width=22,
            command=lambda: self.selecionar_conta(None)
        ).pack(side=tk.LEFT, padx=5, pady=3)

        # Botões das contas individuais
        lista_contas = self.obter_lista_contas()
        for nome_conta in lista_contas:
            saldo = self.calcular_saldo_conta(nome_conta)
            cor = "#27ae60" if saldo >= 0 else "#e74c3c"
            icone = "🏦" if "poup" in nome_conta.lower() else "💳"
            tk.Button(
                self.frame_botoes_contas,
                text=f"{icone} {nome_conta}\n{self.formatar_valor(saldo)}",
                bg=cor, fg="white", font=("Helvetica", 10, "bold"),
                padx=15, pady=10, width=22,
                command=lambda c=nome_conta: self.selecionar_conta(c)
            ).pack(side=tk.LEFT, padx=5, pady=3)

        # Atualizar opções do combo
        self.conta_combo['values'] = [""] + lista_contas

    def selecionar_conta(self, nome_conta):
        """Filtra lançamentos de uma conta específica"""
        self.conta_atual = nome_conta
        
        if nome_conta is None:
            self.lbl_conta_selecionada.config(
                text="📊 EXTRATO GERAL — Todas as contas",
                bg="#d4edda", fg="#27ae60"
            )
        else:
            icone = "🏦" if "poup" in nome_conta.lower() else "💳"
            self.lbl_conta_selecionada.config(
                text=f"{icone} {nome_conta}",
                bg="#e3f2fd", fg="#1565c0"
            )
        
        self.aplicar_filtros()

    def ver_todas_contas(self):
        self.conta_atual = None
        self.filtro_data_inicio.set("")
        self.filtro_data_fim.set("")
        self.filtro_tipo.set("Todos")
        self.filtro_categoria.set("Todas")
        self.selecionar_conta(None)

    def atualizar_categorias(self, event=None):
        cats_entrada = [
            "Disponível em Conta Corrente",
            "Salário",
            "13º Salário - 1ª Parcela",
            "13º Salário - 2ª Parcela",
            "Férias",
            "Pagamento antes das Férias",
            "Pagamento após as Férias",
            "Adiantamento",
            "Vale",
            "Vale após as Férias",
            "Resgate Poupança",
            "Rendimento",
            "Resgate",
            "Reembolso",
            "Venda",
            "Outras Receitas"
        ]
        
        cats_saida = [
            "Disponível em Conta Corrente",
            "Pagamento",
            "Adiantamento",
            "Vale",
            "Vale após as Férias",
            "13º Salário - 1ª Parcela",
            "13º Salário - 2ª Parcela",
            "Férias",
            "Pagamento antes das Férias",
            "Pagamento após as Férias",
            "Transferência Poupança",
            "Moradia",
            "Alimentação",
            "Transporte",
            "Saúde",
            "Educação",
            "Lazer",
            "Cartão",
            "Empréstimo",
            "Contas",
            "Outras Despesas"
        ]
        
        if self.tipo_var.get() == "Entrada":
            cats = cats_entrada
        else:
            cats = cats_saida
            
        self.cat_combo['values'] = cats
        if not self.categoria_var.get() or self.categoria_var.get() not in cats:
            self.categoria_var.set(cats[0])
            
        todas_cats = ["Todas"] + sorted(set(cats_entrada + cats_saida))
        self.filt_categoria['values'] = todas_cats
        
        self.atualizar_botoes_contas()

    def _parse_data(self, texto_data):
        try:
            if " " in texto_data:
                return datetime.strptime(texto_data, "%d/%m/%Y %H:%M")
            else:
                return datetime.strptime(texto_data, "%d/%m/%Y")
        except ValueError:
            return None

    def aplicar_filtros(self, event=None):
        for linha in self.tabela.get_children():
            self.tabela.delete(linha)

        data_inicio_txt = self.filtro_data_inicio.get().strip()
        data_fim_txt = self.filtro_data_fim.get().strip()
        tipo_filtro = self.filtro_tipo.get()
        cat_filtro = self.filtro_categoria.get()

        data_inicio = self._parse_data(data_inicio_txt) if data_inicio_txt else None
        data_fim = self._parse_data(data_fim_txt) if data_fim_txt else None
        if data_fim and len(data_fim_txt) <= 10:
            data_fim = data_fim.replace(hour=23, minute=59, second=59)

        qtd_mostrada = 0
        for vals, tags in self._todos_dados:
            # Filtro de conta
            if self.conta_atual is not None:
                if len(vals) < 7 or vals[6] != self.conta_atual:
                    continue
            # Filtro de data
            data_reg = self._parse_data(vals[0])
            if not data_reg:
                continue
            if data_inicio and data_reg < data_inicio:
                continue
            if data_fim and data_reg > data_fim:
                continue
            # Filtro de tipo e categoria
            if tipo_filtro != "Todos" and vals[2] != tipo_filtro:
                continue
            if cat_filtro != "Todas" and vals[3] != cat_filtro:
                continue
                
            self.tabela.insert("", "end", values=vals, tags=tags)
            qtd_mostrada += 1

        # Status
        if self.conta_atual is None:
            status = f"✅ Exibindo {qtd_mostrada} registro(s) — TODAS as contas"
        else:
            status = f"✅ Exibindo {qtd_mostrada} registro(s) | Conta: {self.conta_atual}"
        if data_inicio_txt or data_fim_txt or tipo_filtro != "Todos" or cat_filtro != "Todas":
            status += " | FILTRADO"
        self.lbl_status_filtro.config(text=status)
        self.atualizar_saldo_exibicao()
        self.atualizar_botoes_contas()

    def calcular_totais_tabela(self):
        entradas = 0.0
        saidas = 0.0
        for linha in self.tabela.get_children():
            vals = self.tabela.item(linha, "values")
            tipo = vals[2]
            val_str = vals[4].replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
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
            text=f"📈 Total de Entradas: {self.formatar_valor(entradas)}"
        )
        self.lbl_saidas.config(
            text=f"📉 Total de Saídas: {self.formatar_valor(saidas)}"
        )
        saldo_fmt = f"💰 SALDO ATUAL: {self.formatar_valor(saldo)}"
        self.lbl_saldo.config(text=saldo_fmt)
        self.lbl_saldo.config(fg="#27ae60" if saldo >= 0 else "#e74c3c")

    def adicionar_registro(self):
        conta = self.conta_var.get().strip() or "Geral"
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

        valor_fmt = self.formatar_valor(valor)
        tag = "credito" if tipo == "Entrada" else "debito"
        vals = (data, conta, tipo, cat, valor_fmt, desc, obs)

        self.fazer_backup()
        self._todos_dados.append((vals, (tag,)))
        self.salvar_dados_csv()
        self.atualizar_botoes_contas()
        self.aplicar_filtros()

        self.valor_var.set("")
        self.descricao_var.set("")
        self.obs_var.set("")
        self.data_var.set(datetime.now().strftime("%d/%m/%Y %H:%M"))

    def editar_registro(self):
        selecionado = self.tabela.selection()
        if not selecionado:
            messagebox.showinfo("Dica", "Selecione um lançamento na tabela")
            return
        vals_atuais = self.tabela.item(selecionado[0], "values")

        janela = tk.Toplevel(self.root)
        janela.title("✏️ Editar Lançamento")
        janela.geometry("520x520")
        janela.grab_set()

        e_data = tk.StringVar(value=vals_atuais[0])
        e_conta = tk.StringVar(value=vals_atuais[1])
        e_tipo = tk.StringVar(value=vals_atuais[2])
        e_cat = tk.StringVar(value=vals_atuais[3])
        e_val = tk.StringVar(value=vals_atuais[4].replace("R$", "").strip())
        e_desc = tk.StringVar(value=vals_atuais[5])
        e_obs = tk.StringVar(value=vals_atuais[6])

        ttk.Label(janela, text="Editar Lançamento", font=("Helvetica", 13, "bold")).pack(pady=10)
        frame = ttk.Frame(janela, padding=15)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Conta:").grid(row=0, column=0, sticky="w", pady=6)
        ttk.Entry(frame, textvariable=e_conta, width=40).grid(row=0, column=1, pady=6)
        ttk.Label(frame, text="Data/Hora:").grid(row=1, column=0, sticky="w", pady=6)
        ttk.Entry(frame, textvariable=e_data, width=40).grid(row=1, column=1, pady=6)
        ttk.Label(frame, text="Tipo:").grid(row=2, column=0, sticky="w", pady=6)
        tipo_combo_edit = ttk.Combobox(frame, textvariable=e_tipo, values=["Entrada", "Saída"], state="readonly", width=37)
        tipo_combo_edit.grid(row=2, column=1, pady=6)
        ttk.Label(frame, text="Categoria:").grid(row=3, column=0, sticky="w", pady=6)
        ttk.Entry(frame, textvariable=e_cat, width=40).grid(row=3, column=1, pady=6)
        ttk.Label(frame, text="Valor:").grid(row=4, column=0, sticky="w", pady=6)
        ttk.Entry(frame, textvariable=e_val, width=40).grid(row=4, column=1, pady=6)
        ttk.Label(frame, text="Descrição:").grid(row=5, column=0, sticky="w", pady=6)
        ttk.Entry(frame, textvariable=e_desc, width=40).grid(row=5, column=1, pady=6)
        ttk.Label(frame, text="Observação:").grid(row=6, column=0, sticky="w", pady=6)
        ttk.Entry(frame, textvariable=e_obs, width=40).grid(row=6, column=1, pady=6)

        def salvar():
            v_str = e_val.get().strip().replace(".", "").replace(",", ".")
            try:
                v_num = float(v_str)
                if v_num <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Erro", "Valor inválido!", parent=janela)
                return
            v_fmt = self.formatar_valor(v_num)
            tag = "credito" if e_tipo.get() == "Entrada" else "debito"
            novos_vals = (e_data.get(), e_conta.get(), e_tipo.get(), e_cat.get(), v_fmt, e_desc.get(), e_obs.get())
            
            self.fazer_backup()
            for i, (dados, _) in enumerate(self._todos_dados):
                if dados == vals_atuais:
                    self._todos_dados[i] = (novos_vals, (tag,))
                    break
            self.salvar_dados_csv()
            self.atualizar_botoes_contas()
            self.aplicar_filtros()
            messagebox.showinfo("Sucesso", "Alterações salvas!", parent=janela)
            janela.destroy()

        ttk.Button(frame, text="💾 Salvar", command=salvar).grid(row=7, column=0, columnspan=2, pady=15)

    def excluir_registro(self):
        selecionado = self.tabela.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um lançamento")
            return
        vals_excluir = self.tabela.item(selecionado[0], "values")
        if messagebox.askyesno("Confirmação", "Excluir este lançamento?\nBackup será feito."):
            self.fazer_backup()
            self._todos_dados = [(d, t) for d, t in self._todos_dados if d != vals_excluir]
            self.salvar_dados_csv()
            self.atualizar_botoes_contas()
            self.aplicar_filtros()

    def gerar_extrato_filtrado(self):
        itens = self.tabela.get_children()
        if not itens:
            messagebox.showwarning("Aviso", "Sem dados para exibir — ajuste os filtros")
            return

        janela = tk.Toplevel(self.root)
        titulo = f"📄 EXTRATO — {self.conta_atual}" if self.conta_atual else "📄 EXTRATO GERAL"
        janela.title(titulo)
        janela.geometry("1100x700")

        periodo_txt = "PERÍODO: "
        if self.filtro_data_inicio.get():
            periodo_txt += f"De {self.filtro_data_inicio.get()} "
        if self.filtro_data_fim.get():
            periodo_txt += f"Até {self.filtro_data_fim.get()}"
        if not self.filtro_data_inicio.get() and not self.filtro_data_fim.get():
            periodo_txt += "Todos os registros"
        if self.conta_atual:
            periodo_txt += f" | Conta: {self.conta_atual}"

        tk.Label(janela, text=titulo, font=("Helvetica", 16, "bold"), fg="#2c3e50").pack(pady=(15,5))
        tk.Label(janela, text=periodo_txt, font=("Helvetica", 11, "bold"), fg="#3498db").pack()
        tk.Label(janela, text=f"Emitido em: {datetime.now().strftime('%d/%m/%Y %H:%M')}", font=("Helvetica", 9), fg="#7f8c8d").pack()

        frame_tab = ttk.Frame(janela)
        frame_tab.pack(fill="both", expand=True, padx=20, pady=15)
        cols = ("Data", "Conta", "Categoria", "Descrição", "Entrada (+)", "Saída (-)", "Saldo Parcial")
        tree = ttk.Treeview(frame_tab, columns=cols, show="headings", height=15)
        
        for c in cols:
            tree.heading(c, text=c)
            if c in ["Entrada (+)", "Saída (-)", "Saldo Parcial"]:
                tree.column(c, width=140, anchor="e")
            elif c == "Data":
                tree.column(c, width=160, anchor="center")
            else:
                tree.column(c, width=200, anchor="w")

        scroll = ttk.Scrollbar(frame_tab, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        saldo = 0.0
        total_ent = 0.0
        total_sai = 0.0

        for linha in itens:
            vals = self.tabela.item(linha, "values")
            data, conta, tipo, cat, val_str, desc, obs = vals
            limpo = val_str.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
            try:
                v = float(limpo)
            except:
                v = 0.0

            if tipo == "Entrada":
                saldo += v
                total_ent += v
                col_e = self.formatar_valor(v)
                col_s = "-"
            else:
                saldo -= v
                total_sai += v
                col_e = "-"
                col_s = self.formatar_valor(v)

            saldo_fmt = self.formatar_valor(saldo)
            hist = f"{desc} {f'| {obs}' if obs else ''}"
            tree.insert("", "end", values=(data, conta, cat, hist, col_e, col_s, saldo_fmt))

        rodape = ttk.Frame(janela, padding=20)
        rodape.pack(fill="x")
        tk.Label(rodape, text=f"📈 TOTAL DE ENTRADAS: {self.formatar_valor(total_ent)}",
                 font=("Helvetica", 11, "bold"), fg="#27ae60").pack(anchor="w")
        tk.Label(rodape, text=f"📉 TOTAL DE SAÍDAS: {self.formatar_valor(total_sai)}",
                 font=("Helvetica", 11, "bold"), fg="#e74c3c").pack(anchor="w")
        tk.Label(rodape, text=f"💰 SALDO FINAL: {self.formatar_valor(saldo)}",
                 font=("Helvetica", 14, "bold"), fg="#27ae60" if saldo >= 0 else "#e74c3c").pack(anchor="w", pady=8)

    def gerar_pdf_filtrado(self):
        itens = self.tabela.get_children()
        if not itens:
            messagebox.showwarning("Aviso", "Sem dados para gerar PDF — ajuste os filtros")
            return

        data_agora = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_conta_arquivo = f"_{self.conta_atual.replace(' ', '_')}" if self.conta_atual else "_geral"
        
        caminho = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf"), ("Todos", "*.*")],
            initialfile=f"extrato{nome_conta_arquivo}_{data_agora}.pdf",
            title="Salvar PDF"
        )
        if not caminho:
            return

        try:
            doc = SimpleDocTemplate(caminho, pagesize=landscape(letter),
                                     leftMargin=25, rightMargin=25, topMargin=25, bottomMargin=25)
            elementos = []
            estilos = getSampleStyleSheet()

            titulo_texto = f"EXTRATO — {self.conta_atual}" if self.conta_atual else "EXTRATO GERAL"
            titulo = ParagraphStyle("Titulo", parent=estilos["Heading1"], fontSize=18, alignment=1,
                                    textColor=colors.HexColor("#2c3e50"), spaceAfter=6)
            subtitulo = ParagraphStyle("Sub", parent=estilos["Normal"], fontSize=11, alignment=1,
                                        textColor=colors.HexColor("#3498db"), spaceAfter=6)
            rodape_estilo = ParagraphStyle("Rodape", parent=estilos["Normal"], fontSize=9, alignment=1, textColor=colors.grey)

            elementos.append(Paragraph(titulo_texto, titulo))

            periodo_txt = ""
            if self.filtro_data_inicio.get() or self.filtro_data_fim.get():
                if self.filtro_data_inicio.get():
                    periodo_txt += f"De: {self.filtro_data_inicio.get()} "
                if self.filtro_data_fim.get():
                    periodo_txt += f"Até: {self.filtro_data_fim.get()}"
            else:
                periodo_txt = "Todos os registros"
            if self.conta_atual:
                periodo_txt += f" | Conta: {self.conta_atual}"
            if self.filtro_tipo.get() != "Todos":
                periodo_txt += f" | Tipo: {self.filtro_tipo.get()}"
            if self.filtro_categoria.get() != "Todas":
                periodo_txt += f" | Categoria: {self.filtro_categoria.get()}"

            elementos.append(Paragraph(periodo_txt, subtitulo))
            elementos.append(Paragraph(f"Emitido em: {datetime.now().strftime('%d/%m/%Y às %H:%M')}", rodape_estilo))
            elementos.append(Spacer(1, 15))

            cel = ParagraphStyle("Cel", parent=estilos["Normal"], fontSize=9, leading=11)
            cab = ParagraphStyle("Cab", parent=estilos["Normal"], fontSize=10, fontName="Helvetica-Bold",
                                  textColor=colors.whitesmoke, alignment=1)

            dados = [[Paragraph(h, cab) for h in ["Data", "Conta", "Tipo", "Categoria", "Valor", "Descrição", "Obs"]]]
            for linha in itens:
                vals = self.tabela.item(linha, "values")
                dados.append([Paragraph(str(v), cel) for v in vals])

            tabela = Table(dados, colWidths=[100, 140, 90, 180, 110, 200, 100])
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

            ent, sai, saldo = self.calcular_totais_tabela()
            resumo = ParagraphStyle("Resumo", parent=estilos["Normal"], fontSize=12, fontName="Helvetica-Bold", spaceAfter=8)
            cor_saldo = "#27ae60" if saldo >= 0 else "#e74c3c"

            elementos.append(Paragraph(f"Total de Entradas: {self.formatar_valor(ent)}", resumo))
            elementos.append(Paragraph(f"Total de Saídas: {self.formatar_valor(sai)}", resumo))
            elementos.append(Paragraph(f"<font color='{cor_saldo}'>SALDO FINAL: {self.formatar_valor(saldo)}</font>", resumo))

            doc.build(elementos)
            messagebox.showinfo("✅ Sucesso", f"PDF salvo com sucesso!")
            
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
                w.writerow(["Data", "Conta", "Tipo", "Categoria", "Valor", "Descrição", "Observação"])
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
                            if len(linha) == 6:
                                linha.insert(1, "Geral")
                            tag = "credito" if linha[2] == "Entrada" else "debito"
                            self._todos_dados.append((tuple(linha), (tag,)))
                self.atualizar_botoes_contas()
                self.selecionar_conta(None)
            except Exception as e:
                messagebox.showerror("Erro", f"Falha ao carregar: {str(e)}")
        else:
            self.atualizar_botoes_contas()
            self.selecionar_conta(None)

if __name__ == "__main__":
    root = tk.Tk()
    app = AppFluxoCaixa(root)
    root.mainloop()