import csv
import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import shutil

ARQUIVO_CSV = "fluxo_de_caixa_continuo.csv"
PASTA_BACKUP = "backups"

class AppFluxoCaixa:
    def __init__(self, root):
        self.root = root
        self.root.title("💰 Fluxo de Caixa - Contas e Poupanças")
        
        largura_janela = 1280
        altura_janela = 850
        largura_tela = self.root.winfo_screenwidth()
        altura_tela = self.root.winfo_screenheight()
        pos_x = (largura_tela // 2) - (largura_janela // 2)
        pos_y = (altura_tela // 2) - (altura_janela // 2)
        self.root.geometry(f"{largura_janela}x{altura_janela}+{pos_x}+{pos_y}")
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
        
        # Variáveis para o filtro por data
        self.filtro_data_inicio = tk.StringVar()
        self.filtro_data_fim = tk.StringVar()
        
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

    def carregar_dados(self):
        self._todos_dados = []
        if os.path.exists(ARQUIVO_CSV):
            try:
                with open(ARQUIVO_CSV, mode="r", encoding="utf-8") as f:
                    leitor = csv.reader(f)
                    for linha in leitor:
                        if len(linha) >= 7:
                            vals = tuple(linha[:7])
                            tipo = vals[2]
                            tag = "credito" if tipo == "Entrada" else "debito"
                            self._todos_dados.append((vals, (tag,)))
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao carregar dados:\n{e}")
        self.aplicar_filtro_data()

    def salvar_dados_csv(self):
        try:
            with open(ARQUIVO_CSV, mode="w", newline="", encoding="utf-8") as f:
                escritor = csv.writer(f)
                for vals, _ in self._todos_dados:
                    escritor.writerow(vals)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar dados:\n{e}")

    def obter_lista_contas(self):
        contas = set()
        for vals, _ in self._todos_dados:
            if len(vals) >= 7 and vals[1].strip():
                contas.add(vals[1].strip())
        return sorted(list(contas))

    def formatar_valor(self, valor):
        return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    def criar_widgets(self):
        header_frame = tk.Frame(self.root, bg="#2c3e50", padx=15, pady=8)
        header_frame.pack(fill="x")
        tk.Label(header_frame, text="💰 CONTROLE DE FLUXO DE CAIXA", 
                 font=("Helvetica", 16, "bold"), fg="white", bg="#2c3e50").pack()
        tk.Label(header_frame, text="Gerenciamento de Lançamentos e Extratos", 
                 font=("Helvetica", 9), fg="#bdc3c7", bg="#2c3e50").pack()

        form_frame = ttk.LabelFrame(self.root, text="📝 Novo Lançamento e Ações")
        form_frame.pack(fill="x", padx=15, pady=5)

        ttk.Label(form_frame, text="Conta Corrente/Poupança:").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        self.conta_combo = ttk.Combobox(form_frame, textvariable=self.conta_var, width=32)
        self.conta_combo.grid(row=0, column=1, padx=6, pady=4, sticky="w")
        ttk.Label(form_frame, text="(deixe vazio = Geral)").grid(row=0, column=2, sticky="w", padx=0, pady=4)

        ttk.Label(form_frame, text="Data/Hora:").grid(row=1, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(form_frame, textvariable=self.data_var, width=20).grid(row=1, column=1, padx=6, pady=4, sticky="w")

        ttk.Label(form_frame, text="Tipo:").grid(row=1, column=2, sticky="w", padx=6, pady=4)
        self.tipo_combo = ttk.Combobox(form_frame, textvariable=self.tipo_var,
                                       values=["Entrada", "Saída"], state="readonly", width=18)
        self.tipo_combo.grid(row=1, column=3, padx=6, pady=4, sticky="w")
        self.tipo_combo.bind("<<ComboboxSelected>>", self.atualizar_categorias)

        ttk.Label(form_frame, text="Categoria:").grid(row=1, column=4, sticky="w", padx=6, pady=4)
        self.cat_combo = ttk.Combobox(form_frame, textvariable=self.categoria_var, state="readonly", width=35)
        self.cat_combo.grid(row=1, column=5, padx=6, pady=4, sticky="w")

        ttk.Label(form_frame, text="Valor R$:").grid(row=2, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(form_frame, textvariable=self.valor_var, width=20).grid(row=2, column=1, padx=6, pady=4, sticky="w")

        ttk.Label(form_frame, text="Descrição:").grid(row=2, column=2, sticky="w", padx=6, pady=4)
        ttk.Entry(form_frame, textvariable=self.descricao_var, width=42).grid(row=2, column=3, columnspan=3, padx=6, pady=4, sticky="w")

        ttk.Label(form_frame, text="Observação:").grid(row=3, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(form_frame, textvariable=self.obs_var, width=42).grid(row=3, column=1, columnspan=3, padx=6, pady=4, sticky="w")

        # Linha de Filtro por Data integrada na parte inferior do painel de ações
        filtro_data_frame = ttk.Frame(form_frame)
        filtro_data_frame.grid(row=4, column=0, columnspan=6, sticky="w", padx=6, pady=4)
        
        ttk.Label(filtro_data_frame, text="📅 Filtrar Período — Início:").pack(side="left", padx=(0, 2))
        ttk.Entry(filtro_data_frame, textvariable=self.filtro_data_inicio, width=12).pack(side="left", padx=2)
        ttk.Label(filtro_data_frame, text="Até:").pack(side="left", padx=5)
        ttk.Entry(filtro_data_frame, textvariable=self.filtro_data_fim, width=12).pack(side="left", padx=2)
        ttk.Button(filtro_data_frame, text="🔍 Filtrar", command=self.aplicar_filtro_data).pack(side="left", padx=8)
        ttk.Button(filtro_data_frame, text="🔄 Limpar Filtro", command=self.limpar_filtro_data).pack(side="left", padx=2)

        # Botões principais
        btn_form = ttk.Frame(form_frame)
        btn_form.grid(row=5, column=0, columnspan=6, pady=6)
        ttk.Button(btn_form, text="✅ Adicionar", command=self.adicionar_registro).pack(side="left", padx=5)
        ttk.Button(btn_form, text="✏️ Editar", command=self.editar_registro).pack(side="left", padx=5)
        ttk.Button(btn_form, text="🗑️ Excluir", command=self.excluir_registro).pack(side="left", padx=5)
        ttk.Button(btn_form, text="📄 Extrato", command=self.gerar_extrato).pack(side="left", padx=5)
        ttk.Button(btn_form, text="📋 Gerar PDF", command=self.gerar_pdf).pack(side="left", padx=5)

        tabela_frame = ttk.Frame(self.root)
        tabela_frame.pack(fill="both", expand=True, padx=15, pady=5)
        colunas = ("Data", "Conta", "Tipo", "Categoria", "Valor", "Descrição", "Observação")
        self.tabela = ttk.Treeview(tabela_frame, columns=colunas, show="headings", selectmode="browse", height=12)
        self.tabela.bind("<Double-1>", lambda e: self.editar_registro())
        self.tabela.tag_configure("credito", foreground="#27ae60", font=("Helvetica", 9, "bold"))
        self.tabela.tag_configure("debito", foreground="#e74c3c", font=("Helvetica", 9, "bold"))

        for col in colunas:
            self.tabela.heading(col, text=col)
            if col == "Data":
                self.tabela.column(col, width=140, anchor="center")
            elif col == "Conta":
                self.tabela.column(col, width=170, anchor="w")
            elif col == "Tipo":
                self.tabela.column(col, width=90, anchor="center")
            elif col == "Categoria":
                self.tabela.column(col, width=210, anchor="w")
            elif col == "Valor":
                self.tabela.column(col, width=120, anchor="e")
            elif col == "Descrição":
                self.tabela.column(col, width=260, anchor="w")
            elif col == "Observação":
                self.tabela.column(col, width=230, anchor="w")

        scroll = ttk.Scrollbar(tabela_frame, orient="vertical", command=self.tabela.yview)
        self.tabela.configure(yscrollcommand=scroll.set)
        self.tabela.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        resumo_frame = tk.Frame(self.root, bg="#ecf0f1", padx=15, pady=8)
        resumo_frame.pack(fill="x")
        self.lbl_entradas = tk.Label(resumo_frame, text="📈 Total de Entradas: R$ 0,00",
                                     font=("Helvetica", 10, "bold"), fg="#27ae60", bg="#ecf0f1")
        self.lbl_entradas.pack(anchor="w")
        self.lbl_saidas = tk.Label(resumo_frame, text="📉 Total de Saídas: R$ 0,00",
                                   font=("Helvetica", 10, "bold"), fg="#e74c3c", bg="#ecf0f1")
        self.lbl_saidas.pack(anchor="w")
        self.lbl_saldo = tk.Label(resumo_frame, text="💰 SALDO ATUAL: R$ 0,00",
                                  font=("Helvetica", 12, "bold"), bg="#ecf0f1")
        self.lbl_saldo.pack(anchor="w", pady=3)
        self.lbl_status = tk.Label(resumo_frame, text="✅ Exibindo: TODOS os registros",
                                   font=("Helvetica", 9, "italic"), fg="#3498db", bg="#ecf0f1")
        self.lbl_status.pack(anchor="w")

    def atualizar_categorias(self, event=None):
        cats_entrada = ["Disponível em Conta Corrente", "Salário", "13º Salário - 1ª Parcela", "13º Salário - 2ª Parcela", "Férias", "Pagamento antes das Férias", "Pagamento após das Férias", "Adiantamento", "Vale", "Vale após as Férias", "Resgate Poupança", "Rendimento", "Resgate", "Reembolso", "Venda", "Outras Receitas"]
        cats_saida = ["Disponível em Conta Corrente", "Pagamento", "Adiantamento", "Vale", "Vale após as Férias", "13º Salário - 1ª Parcela", "13º Salário - 2ª Parcela", "Férias", "Pagamento antes das Férias", "Pagamento após das Férias", "Transferência Poupança", "Moradia", "Alimentação", "Transporte", "Saúde", "Educação", "Lazer", "Cartão", "Empréstimo", "Contas", "Outras Despesas"]
        
        cats = cats_entrada if self.tipo_var.get() == "Entrada" else cats_saida
        self.cat_combo['values'] = cats
        if not self.categoria_var.get() or self.categoria_var.get() not in cats:
            self.categoria_var.set(cats[0])
            
        lista_contas = self.obter_lista_contas()
        self.conta_combo['values'] = [""] + lista_contas

    def _parse_data(self, texto_data):
        try:
            if " " in texto_data:
                return datetime.strptime(texto_data, "%d/%m/%Y %H:%M")
            else:
                return datetime.strptime(texto_data, "%d/%m/%Y")
        except ValueError:
            return None

    def aplicar_filtro_data(self):
        for linha in self.tabela.get_children():
            self.tabela.delete(linha)

        d_ini_txt = self.filtro_data_inicio.get().strip()
        d_fim_txt = self.filtro_data_fim.get().strip()

        d_ini = self._parse_data(d_ini_txt) if d_ini_txt else None
        d_fim = self._parse_data(d_fim_txt) if d_fim_txt else None
        if d_fim and len(d_fim_txt) <= 10:
            d_fim = d_fim.replace(hour=23, minute=59, second=59)

        qtd = 0
        for vals, tags in self._todos_dados:
            data_reg = self._parse_data(vals[0])
            if data_reg:
                if d_ini and data_reg < d_ini:
                    continue
                if d_fim and data_reg > d_fim:
                    continue
            self.tabela.insert("", "end", values=vals, tags=tags)
            qtd += 1

        status = f"✅ Exibindo {qtd} registro(s)"
        if d_ini_txt or d_fim_txt:
            status += " (FILTRADO POR PERÍODO)"
        self.lbl_status.config(text=status)
        self.atualizar_saldo_exibicao()
        
        lista_contas = self.obter_lista_contas()
        self.conta_combo['values'] = [""] + lista_contas

    def limpar_filtro_data(self):
        self.filtro_data_inicio.set("")
        self.filtro_data_fim.set("")
        self.aplicar_filtro_data()

    def calcular_totais(self):
        entradas = 0.0
        saidas = 0.0
        for vals, _ in self._todos_dados:
            try:
                tipo = vals[2]
                val_str = vals[4].replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
                valor = float(val_str)
                if tipo == "Entrada":
                    entradas += valor
                else:
                    saidas += valor
            except (ValueError, IndexError):
                continue
        return entradas, saidas, entradas - saidas

    def atualizar_saldo_exibicao(self):
        entradas, saidas, saldo = self.calcular_totais()
        self.lbl_entradas.config(text=f"📈 Total de Entradas: {self.formatar_valor(entradas)}")
        self.lbl_saidas.config(text=f"📉 Total de Saídas: {self.formatar_valor(saidas)}")
        self.lbl_saldo.config(text=f"💰 SALDO ATUAL: {self.formatar_valor(saldo)}", fg="#27ae60" if saldo >= 0 else "#e74c3c")

    def adicionar_registro(self):
        conta = self.conta_var.get().strip() or "Geral"
        data = self.data_var.get().strip() or datetime.now().strftime("%d/%m/%Y %H:%M")
        tipo = self.tipo_var.get()
        cat = self.categoria_var.get()
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
        self.aplicar_filtro_data()

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

        ja = tk.Toplevel(self.root)
        ja.title("✏️ Editar Lançamento")
        ja.geometry("580x520")
        ja.grab_set()

        e_data = tk.StringVar(value=vals_atuais[0])
        e_conta = tk.StringVar(value=vals_atuais[1])
        e_tipo = tk.StringVar(value=vals_atuais[2])
        e_cat = tk.StringVar(value=vals_atuais[3])
        e_val = tk.StringVar(value=vals_atuais[4].replace("R$", "").strip())
        e_desc = tk.StringVar(value=vals_atuais[5])
        e_obs = tk.StringVar(value=vals_atuais[6])

        tk.Label(ja, text="Editar Lançamento", font=("Helvetica", 12, "bold")).pack(pady=8)
        frame = ttk.Frame(ja, padding=12)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Data/Hora:").grid(row=0, column=0, sticky="w", pady=5)
        ttk.Entry(frame, textvariable=e_data, width=42).grid(row=0, column=1, pady=5, sticky="w")
        
        ttk.Label(frame, text="Conta Corrente/Poupança:").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(frame, textvariable=e_conta, width=42).grid(row=1, column=1, pady=5, sticky="w")
        
        ttk.Label(frame, text="Tipo:").grid(row=2, column=0, sticky="w", pady=5)
        ttk.Combobox(frame, textvariable=e_tipo, values=["Entrada", "Saída"], state="readonly", width=40).grid(row=2, column=1, pady=5, sticky="w")
        
        ttk.Label(frame, text="Categoria:").grid(row=3, column=0, sticky="w", pady=5)
        ttk.Entry(frame, textvariable=e_cat, width=42).grid(row=3, column=1, pady=5, sticky="w")
        
        ttk.Label(frame, text="Valor:").grid(row=4, column=0, sticky="w", pady=5)
        ttk.Entry(frame, textvariable=e_val, width=42).grid(row=4, column=1, pady=5, sticky="w")
        
        ttk.Label(frame, text="Descrição:").grid(row=5, column=0, sticky="w", pady=5)
        ttk.Entry(frame, textvariable=e_desc, width=42).grid(row=5, column=1, pady=5, sticky="w")
        
        ttk.Label(frame, text="Observação:").grid(row=6, column=0, sticky="w", pady=5)
        ttk.Entry(frame, textvariable=e_obs, width=42).grid(row=6, column=1, pady=5, sticky="w")

        def salvar():
            v_str = e_val.get().strip().replace(".", "").replace(",", ".")
            try:
                v_num = float(v_str)
                if v_num <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Erro", "Valor inválido!", parent=ja)
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
            self.aplicar_filtro_data()
            messagebox.showinfo("Sucesso", "Alterações salvas!", parent=ja)
            ja.destroy()

        ttk.Button(frame, text="💾 Salvar Alterações", command=salvar).grid(row=7, column=0, columnspan=2, pady=12)

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
            self.aplicar_filtro_data()

    def gerar_extrato(self):
        itens = self.tabela.get_children()
        if not itens:
            messagebox.showwarning("Aviso", "Sem dados para exibir")
            return

        janela = tk.Toplevel(self.root)
        titulo = "📄 EXTRATO"
        janela.title(titulo)
        janela.geometry("1100x700")

        periodo_txt = "PERÍODO: "
        if self.filtro_data_inicio.get():
            periodo_txt += f"De {self.filtro_data_inicio.get()} "
        if self.filtro_data_fim.get():
            periodo_txt += f"Até {self.filtro_data_fim.get()}"
        if not self.filtro_data_inicio.get() and not self.filtro_data_fim.get():
            periodo_txt += "Todos os registros"

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

        saldo = total_ent = total_sai = 0.0
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
                col_e, col_s = self.formatar_valor(v), "-"
            else:
                saldo -= v
                total_sai += v
                col_e, col_s = "-", self.formatar_valor(v)

            saldo_fmt = self.formatar_valor(saldo)
            hist = f"{desc} {f'| {obs}' if obs else ''}"
            tree.insert("", "end", values=(data, conta, cat, hist, col_e, col_s, saldo_fmt))

        rodape = ttk.Frame(janela, padding=20)
        rodape.pack(fill="x")
        tk.Label(rodape, text=f"📈 TOTAL DE ENTRADAS: {self.formatar_valor(total_ent)}", font=("Helvetica", 11, "bold"), fg="#27ae60").pack(anchor="w")
        tk.Label(rodape, text=f"📉 TOTAL DE SAÍDAS: {self.formatar_valor(total_sai)}", font=("Helvetica", 11, "bold"), fg="#e74c3c").pack(anchor="w")
        tk.Label(rodape, text=f"💰 SALDO FINAL: {self.formatar_valor(saldo)}", font=("Helvetica", 14, "bold"), fg="#27ae60" if saldo >= 0 else "#e74c3c").pack(anchor="w", pady=8)

    def gerar_pdf(self):
        itens = self.tabela.get_children()
        if not itens:
            messagebox.showwarning("Aviso", "Sem dados para gerar PDF")
            return

        data_agora = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        caminho = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf"), ("Todos", "*.*")],
            initialfile=f"extrato_{data_agora}.pdf",
            title="Salvar PDF"
        )
        if not caminho:
            return

        try:
            doc = SimpleDocTemplate(caminho, pagesize=landscape(letter), leftMargin=25, rightMargin=25, topMargin=25, bottomMargin=25)
            elementos = []
            estilos = getSampleStyleSheet()

            titulo_texto = "EXTRATO DE LANÇAMENTOS"
            titulo = ParagraphStyle("Titulo", parent=estilos["Heading1"], fontSize=18, alignment=1, textColor=colors.HexColor("#2c3e50"), spaceAfter=6)
            subtitulo = ParagraphStyle("Sub", parent=estilos["Normal"], fontSize=11, alignment=1, textColor=colors.HexColor("#3498db"), spaceAfter=6)
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

            elementos.append(Paragraph(periodo_txt, subtitulo))
            elementos.append(Paragraph(f"Emitido em: {datetime.now().strftime('%d/%m/%Y às %H:%M')}", rodape_estilo))
            elementos.append(Spacer(1, 15))

            cel = ParagraphStyle("Cel", parent=estilos["Normal"], fontSize=9, leading=11)
            cab = ParagraphStyle("Cab", parent=estilos["Normal"], fontSize=10, fontName="Helvetica-Bold", textColor=colors.whitesmoke, alignment=1)

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

            ent, sai, saldo = self.calcular_totais()
            resumo = ParagraphStyle("Resumo", parent=estilos["Normal"], fontSize=12, fontName="Helvetica-Bold", spaceAfter=8)
            cor_saldo = "#27ae60" if saldo >= 0 else "#e74c3c"

            elementos.append(Paragraph(f"Total de Entradas: {self.formatar_valor(ent)}", resumo))
            elementos.append(Paragraph(f"Total de Saídas: {self.formatar_valor(sai)}", resumo))
            elementos.append(Paragraph(f"<font color='{cor_saldo}'>SALDO FINAL: {self.formatar_valor(saldo)}</font>", resumo))

            doc.build(elementos)
            messagebox.showinfo("Sucesso", f"PDF gerado com sucesso em:\n{caminho}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar PDF:\n{e}")

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = AppFluxoCaixa(root)
        root.mainloop()
    except Exception as e:
        print(f"Erro fatal: {e}")
        input("Pressione Enter para fechar...")