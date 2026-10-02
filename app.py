import csv
import os
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
from datetime import datetime

# Importações do ReportLab para geração de PDF
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# Importações para abrir o PDF automaticamente após salvar
import subprocess
import sys

# Nome do arquivo CSV para persistência dos dados
ARQUIVO_CSV = "fluxo_de_caixa_continuo.csv"

class AppFluxoCaixa:
    def __init__(self, root):
        self.root = root
        self.root.title("Controle de Fluxo de Caixa & Balancete")
        self.root.geometry("1100x780")
        self.root.minsize(950, 660)

        # Configuração de Estilo
        self.style = ttk.Style()
        self.style.theme_use("clam")

        # Variáveis de controle
        self.tipo_var = tk.StringVar(value="Entrada")
        self.categoria_var = tk.StringVar(value="Geral")
        self.valor_var = tk.StringVar()
        self.data_var = tk.StringVar(value=datetime.now().strftime("%d/%m/%Y %H:%M"))
        self.descricao_var = tk.StringVar()
        self.obs_var = tk.StringVar()
        self.conta_poupanca_var = tk.StringVar(value="PAGAMENTO")

        # Criar interface
        self.criar_widgets()
        self.carregar_dados()

    def criar_widgets(self):
        # Título Principal
        titulo_label = tk.Label(self.root, text="Controle de Fluxo de Caixa & Balancete por Crédito", font=("Helvetica", 18, "bold"), fg="#2c3e50")
        titulo_label.pack(pady=10)

        # Frame de Formulário
        form_frame = ttk.LabelFrame(self.root, text=" Novo Lançamento ", padding=15)
        form_frame.pack(fill="x", padx=20, pady=10)

        # Grid do Formulário - Linha 0
        ttk.Label(form_frame, text="Tipo:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.tipo_combo = ttk.Combobox(form_frame, textvariable=self.tipo_var, values=["Entrada", "Saída"], state="readonly", width=15)
        self.tipo_combo.grid(row=0, column=1, sticky="w", padx=5, pady=5)
        self.tipo_combo.bind("<<ComboboxSelected>>", self.atualizar_categorias)

        ttk.Label(form_frame, text="Categoria:").grid(row=0, column=2, sticky="w", padx=5, pady=5)
        self.cat_combo = ttk.Combobox(form_frame, textvariable=self.categoria_var, state="readonly", width=25)
        self.cat_combo.grid(row=0, column=3, sticky="w", padx=5, pady=5)

        # Linha 1
        ttk.Label(form_frame, text="Valor (R$):").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        valor_entry = ttk.Entry(form_frame, textvariable=self.valor_var, width=15)
        valor_entry.grid(row=1, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(form_frame, text="Conta Corrente / Poupança:").grid(row=1, column=2, sticky="w", padx=5, pady=5)
        
        opcoes_pagamento = [
            "PAGAMENTO",
            "ADIANTAMENTO VALE",
            "13 SALARIO 1 PARCELA",
            "13 SALARIO 2 PARCELA",
            "FERIAS",
            "PAGAMENTO ANTES DAS FERIAS",
            "FERIAS TRABALHADAS 10 DIAS",
            "PAGAMENTO APOS AS FERIAS",
            "VALE APOS FERIAS",
            "DESPESAS DIVERSAS",
            "OUTRA CONTA (DIGITAR MANUALMENTE)"
        ]
        self.conta_poupanca_combo = ttk.Combobox(form_frame, textvariable=self.conta_poupanca_var, values=opcoes_pagamento, state="readonly", width=35)
        self.conta_poupanca_combo.grid(row=1, column=3, sticky="w", padx=5, pady=5)
        self.conta_poupanca_combo.bind("<<ComboboxSelected>>", self.verificar_conta_manual)

        # Inicializar categorias
        self.atualizar_categorias()

        # Linha 2 - Data e Hora
        ttk.Label(form_frame, text="Data e Hora:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        data_entry = ttk.Entry(form_frame, textvariable=self.data_var, width=20)
        data_entry.grid(row=2, column=1, sticky="w", padx=5, pady=5)
        tk.Label(form_frame, text="(Formato: DD/MM/AAAA HH:MM)", font=("Helvetica", 8), fg="#7f8c8d").grid(row=2, column=2, sticky="w", padx=5, pady=5)

        # Linha 3 - Descrição
        ttk.Label(form_frame, text="Descrição:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.desc_entry = ttk.Entry(form_frame, textvariable=self.descricao_var, width=50)
        self.desc_entry.grid(row=3, column=1, columnspan=3, sticky="w", padx=5, pady=5)

        # Linha 4 - Observação
        ttk.Label(form_frame, text="Observação:").grid(row=4, column=0, sticky="w", padx=5, pady=5)
        obs_entry = ttk.Entry(form_frame, textvariable=self.obs_var, width=50)
        obs_entry.grid(row=4, column=1, columnspan=3, sticky="w", padx=5, pady=5)

        # Botões de Ação
        btn_frame = ttk.Frame(form_frame)
        btn_frame.grid(row=5, column=0, columnspan=4, pady=10)

        btn_adicionar = ttk.Button(btn_frame, text="Adicionar Lançamento", command=self.adicionar_registro)
        btn_adicionar.pack(side="left", padx=5)

        btn_editar = ttk.Button(btn_frame, text="Editar Selecionado", command=self.carregar_registro_para_edicao)
        btn_editar.pack(side="left", padx=5)

        btn_excluir = ttk.Button(btn_frame, text="Excluir Selecionado", command=self.excluir_registro)
        btn_excluir.pack(side="left", padx=5)

        btn_excluir_todos = ttk.Button(btn_frame, text="Excluir Todos", command=self.excluir_todos_registros)
        btn_excluir_todos.pack(side="left", padx=5)

        btn_balancete = ttk.Button(btn_frame, text="Balancete por Crédito", command=self.gerar_balancete_creditos)
        btn_balancete.pack(side="left", padx=5)

        btn_extrato = ttk.Button(btn_frame, text="Extrato Conta", command=self.gerar_extrato_conta_corrente)
        btn_extrato.pack(side="left", padx=5)

        btn_pdf = ttk.Button(btn_frame, text="Relatório PDF", command=self.gerar_pdf)
        btn_pdf.pack(side="left", padx=5)

        # Frame de Resumo / Saldo
        resumo_frame = ttk.LabelFrame(self.root, text=" Resumo Financeiro ", padding=10)
        resumo_frame.pack(fill="x", padx=20, pady=5)

        self.lbl_saldo = tk.Label(resumo_frame, text="Saldo Atual: R$ 0,00", font=("Helvetica", 12, "bold"), fg="#27ae60")
        self.lbl_saldo.pack(anchor="w", padx=10)

        # Tabela (Treeview) para exibir os dados
        tabela_frame = ttk.Frame(self.root)
        tabela_frame.pack(fill="both", expand=True, padx=20, pady=10)

        colunas = ("Data", "Tipo", "Categoria", "Valor", "Descrição", "Observação")
        self.tabela = ttk.Treeview(tabela_frame, columns=colunas, show="headings", selectmode="browse")
        self.tabela.bind("<Double-1>", lambda event: self.carregar_registro_para_edicao())

        for col in colunas:
            self.tabela.heading(col, text=col)
            if col == "Valor":
                self.tabela.column(col, width=100, anchor="e")
            elif col in ["Tipo", "Categoria"]:
                self.tabela.column(col, width=155, anchor="center")
            elif col == "Data":
                self.tabela.column(col, width=130, anchor="center")
            else:
                self.tabela.column(col, width=170, anchor="w")

        scrollbar = ttk.Scrollbar(tabela_frame, orient="vertical", command=self.tabela.yview)
        self.tabela.configure(yscrollcommand=scrollbar.set)

        self.tabela.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def atualizar_categorias(self, event=None):
        tipo = self.tipo_var.get()
        if tipo == "Entrada":
            cats = ["Resgate da Poupança", "Transferência para Poupança", "Conta Corrente", "Outras Despesas", "Outras Receitas"]
        else:
            cats = ["Saída", "Débito", "Transferência para Poupança", "Conta Corrente", "Despesa Operacional", "Outras Despesas"]
        self.cat_combo['values'] = cats
        self.categoria_var.set(cats[0])

    def verificar_conta_manual(self, event=None):
        if self.conta_poupanca_var.get() == "OUTRA CONTA (DIGITAR MANUALMENTE)":
            conta_manual = simpledialog.askstring("Conta Manual", "Digite o número ou nome da conta resgatada/utilizada:")
            if conta_manual and conta_manual.strip():
                self.conta_manual_custom = conta_manual.strip().upper()
            else:
                self.conta_manual_custom = "OUTRA CONTA"
        else:
            self.conta_manual_custom = None

    def adicionar_registro(self):
        tipo = self.tipo_var.get()
        categoria = self.categoria_var.get()
        
        data_lancamento = self.data_var.get().strip()
        if not data_lancamento:
            data_lancamento = datetime.now().strftime("%d/%m/%Y %H:%M")

        valor_bruto = self.valor_var.get().strip()
        if not valor_bruto:
            messagebox.showerror("Erro", "Preencha o campo de Valor!")
            return

        descricao = self.descricao_var.get().strip()
        if not descricao:
            messagebox.showerror("Erro", "Preencha o campo de Descrição!")
            return

        if categoria == "Resgate da Poupança":
            qual_poupanca = simpledialog.askstring("Qual Poupança?", "Informe de qual poupança foi feito o resgate (Ex: Poupança Caixa, Nubank):")
            if qual_poupanca:
                descricao = f"Resgate da {qual_poupanca.strip()} - {descricao}"

        conta = self.conta_poupanca_var.get().strip()
        if conta == "OUTRA CONTA (DIGITAR MANUALMENTE)":
            if hasattr(self, 'conta_manual_custom') and self.conta_manual_custom:
                conta = self.conta_manual_custom
            else:
                conta_manual = simpledialog.askstring("Conta Manual", "Digite o número ou nome da conta resgatada/utilizada:")
                conta = conta_manual.strip().upper() if conta_manual and conta_manual.strip() else "OUTRA CONTA"

        valor_str = valor_bruto.replace("R$", "").strip()
        if "," in valor_str:
            valor_str = valor_str.replace(".", "").replace(",", ".")
        
        try:
            valor = float(valor_str)
            if valor <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Erro", "Insira um valor numérico válido maior que zero (ex: 7033,70)!")
            return

        obs = self.obs_var.get().strip()

        descricao_final = f"[{conta}] - {descricao}" if conta and not descricao.startswith("[") else descricao
        valor_formatado = f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

        self.tabela.insert("", "end", values=(data_lancamento, tipo, categoria, valor_formatado, descricao_final, obs))
        self.salvar_dados_csv()

        self.valor_var.set("")
        self.descricao_var.set("")
        self.obs_var.set("")
        self.data_var.set(datetime.now().strftime("%d/%m/%Y %H:%M"))
        self.conta_manual_custom = None
        self.atualizar_saldo()

    def carregar_registro_para_edicao(self):
        selecionado = self.tabela.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um registro na tabela para editar (ou clique duas vezes sobre ele).")
            return

        item_id = selecionado[0]
        vals = self.tabela.item(item_id, "values")
        data, tipo, cat, val_str, desc, obs = vals

        self.edit_janela = tk.Toplevel(self.root)
        self.edit_janela.title("Editar Lançamento")
        self.edit_janela.geometry("500x440")
        self.edit_janela.grab_set()

        ttk.Label(self.edit_janela, text="Editar Dados do Lançamento", font=("Helvetica", 12, "bold")).pack(pady=10)

        frame_edicao = ttk.Frame(self.edit_janela, padding=15)
        frame_edicao.pack(fill="both", expand=True)

        ttk.Label(frame_edicao, text="Data e Hora (DD/MM/AAAA HH:MM):").grid(row=0, column=0, sticky="w", pady=5)
        e_data_var = tk.StringVar(value=data)
        e_data_entry = ttk.Entry(frame_edicao, textvariable=e_data_var, width=27)
        e_data_entry.grid(row=0, column=1, sticky="w", pady=5)

        ttk.Label(frame_edicao, text="Tipo:").grid(row=1, column=0, sticky="w", pady=5)
        e_tipo_var = tk.StringVar(value=tipo)
        e_tipo_combo = ttk.Combobox(frame_edicao, textvariable=e_tipo_var, values=["Entrada", "Saída"], state="readonly", width=25)
        e_tipo_combo.grid(row=1, column=1, sticky="w", pady=5)

        ttk.Label(frame_edicao, text="Categoria:").grid(row=2, column=0, sticky="w", pady=5)
        e_cat_var = tk.StringVar(value=cat)
        e_cat_combo = ttk.Combobox(frame_edicao, textvariable=e_cat_var, values=self.cat_combo['values'], state="readonly", width=25)
        e_cat_combo.grid(row=2, column=1, sticky="w", pady=5)

        def mudar_cats_edicao(event=None):
            if e_tipo_var.get() == "Entrada":
                e_cat_combo['values'] = ["Resgate da Poupança", "Transferência para Poupança", "Conta Corrente", "Outras Despesas", "Outras Receitas"]
            else:
                e_cat_combo['values'] = ["Saída", "Débito", "Transferência para Poupança", "Conta Corrente", "Despesa Operacional", "Outras Despesas"]
        e_tipo_combo.bind("<<ComboboxSelected>>", mudar_cats_edicao)

        ttk.Label(frame_edicao, text="Valor (R$):").grid(row=3, column=0, sticky="w", pady=5)
        limpa_val = val_str.replace("R$", "").strip()
        e_val_var = tk.StringVar(value=limpa_val)
        e_val_entry = ttk.Entry(frame_edicao, textvariable=e_val_var, width=27)
        e_val_entry.grid(row=3, column=1, sticky="w", pady=5)

        ttk.Label(frame_edicao, text="Descrição:").grid(row=4, column=0, sticky="w", pady=5)
        e_desc_var = tk.StringVar(value=desc)
        e_desc_entry = ttk.Entry(frame_edicao, textvariable=e_desc_var, width=38)
        e_desc_entry.grid(row=4, column=1, sticky="w", pady=5)

        ttk.Label(frame_edicao, text="Observação:").grid(row=5, column=0, sticky="w", pady=5)
        e_obs_var = tk.StringVar(value=obs)
        e_obs_entry = ttk.Entry(frame_edicao, textvariable=e_obs_var, width=38)
        e_obs_entry.grid(row=5, column=1, sticky="w", pady=5)

        def salvar_edicao():
            nova_data = e_data_var.get().strip()
            novo_tipo = e_tipo_var.get()
            nova_cat = e_cat_var.get()
            novo_val_bruto = e_val_var.get().strip()
            nova_desc = e_desc_var.get().strip()
            nova_obs = e_obs_var.get().strip()

            if not nova_data or not novo_val_bruto or not nova_desc:
                messagebox.showerror("Erro", "Data, Valor e Descrição não podem ficar vazios!", parent=self.edit_janela)
                return

            v_str = novo_val_bruto.replace("R$", "").strip()
            if "," in v_str:
                v_str = v_str.replace(".", "").replace(",", ".")
            
            try:
                v_num = float(v_str)
                if v_num <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Erro", "Insira um valor numérico válido!", parent=self.edit_janela)
                return

            novo_val_formatado = f"R$ {v_num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

            self.tabela.item(item_id, values=(nova_data, novo_tipo, nova_cat, novo_val_formatado, nova_desc, nova_obs))
            self.salvar_dados_csv()
            self.atualizar_saldo()

            messagebox.showinfo("Sucesso", "Lançamento atualizado com sucesso!", parent=self.edit_janela)
            self.edit_janela.destroy()

        btn_salvar_edit = ttk.Button(frame_edicao, text="Salvar Alterações", command=salvar_edicao)
        btn_salvar_edit.grid(row=6, column=0, columnspan=2, pady=15)

    def excluir_registro(self):
        selecionado = self.tabela.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um registro na tabela para excluir.")
            return

        if messagebox.askyesno("Confirmação", "Deseja realmente excluir o lançamento selecionado?"):
            self.tabela.delete(selecionado)
            self.salvar_dados_csv()
            self.atualizar_saldo()

    def excluir_todos_registros(self):
        itens = self.tabela.get_children()
        if not itens:
            messagebox.showinfo("Aviso", "A tabela já está vazia.")
            return

        if messagebox.askyesno("ATENÇÃO", "Deseja realmente excluir TODOS os lançamentos?"):
            for item in itens:
                self.tabela.delete(item)
            self.salvar_dados_csv()
            self.atualizar_saldo()
            messagebox.showinfo("Sucesso", "Todos os lançamentos foram apagados.")

    def atualizar_saldo(self):
        saldo = 0.0
        for linha in self.tabela.get_children():
            vals = self.tabela.item(linha, "values")
            tipo = vals[1]
            val_str = vals[3].replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
            try:
                valor = float(val_str)
                if tipo == "Entrada":
                    saldo += valor
                else:
                    saldo -= valor
            except ValueError:
                pass

        self.lbl_saldo.config(text=f"Saldo Atual: R$ {saldo:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        if saldo >= 0:
            self.lbl_saldo.config(fg="#27ae60")
        else:
            self.lbl_saldo.config(fg="#c0392b")

    def gerar_balancete_creditos(self):
        itens = self.tabela.get_children()
        if not itens:
            messagebox.showwarning("Aviso", "Não há lançamentos para gerar o balancete.")
            return

        bal_janela = tk.Toplevel(self.root)
        bal_janela.title("Balancete por Crédito e Abatimento de Despesas")
        bal_janela.geometry("1050x650")

        header_frame = ttk.Frame(bal_janela, padding=10)
        header_frame.pack(fill="x")

        tk.Label(header_frame, text="BALANCETE ANALÍTICO: CRÉDITOS VS DESPESAS", font=("Helvetica", 14, "bold"), fg="#2c3e50").pack(anchor="w")
        tk.Label(header_frame, text="Demonstração do consumo de despesas por crédito e saldo real restante para uso.", font=("Helvetica", 9), fg="#7f8c8d").pack(anchor="w")

        tabela_frame_bal = ttk.Frame(bal_janela)
        tabela_frame_bal.pack(fill="both", expand=True, padx=10, pady=5)

        cols = ("Data", "Tipo / Evento", "Descrição / Histórico", "Valor Original", "Valor Abatido / Gasto", "Saldo Restante do Crédito")
        tree_bal = ttk.Treeview(tabela_frame_bal, columns=cols, show="headings", selectmode="none")

        for col in cols:
            tree_bal.heading(col, text=col)
            if col in ["Valor Original", "Valor Abatido / Gasto", "Saldo Restante do Crédito"]:
                tree_bal.column(col, width=125, anchor="e")
            elif col in ["Tipo / Evento"]:
                tree_bal.column(col, width=130, anchor="center")
            elif col == "Data":
                tree_bal.column(col, width=120, anchor="center")
            else:
                tree_bal.column(col, width=220, anchor="w")

        scroll_bal = ttk.Scrollbar(tabela_frame_bal, orient="vertical", command=tree_bal.yview)
        tree_bal.configure(yscrollcommand=scroll_bal.set)

        tree_bal.pack(side="left", fill="both", expand=True)
        scroll_bal.pack(side="right", fill="y")

        creditos_ativos = []
        linhas_balancete = []
        despesas_nao_alocadas = 0.0

        for linha in itens:
            vals = self.tabela.item(linha, "values")
            data, tipo, cat, val_str, desc, obs = vals
            limpo = val_str.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
            try:
                v_num = float(limpo)
            except ValueError:
                v_num = 0.0

            historico = f"{desc} {f'({obs})' if obs else ''}"

            if tipo == "Entrada":
                creditos_ativos.append({
                    'data': data,
                    'desc': historico,
                    'valor_inicial': v_num,
                    'saldo': v_num
                })
                linhas_balancete.append((data, f"[ENTRADA] {cat}", historico, f"+ R$ {v_num:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), "-", "-"))
            else:
                valor_gasto_restante = v_num
                detalhe_abatimento = []

                while valor_gasto_restante > 0 and creditos_ativos:
                    credito_atual = creditos_ativos[0]
                    if credito_atual['saldo'] >= valor_gasto_restante:
                        credito_atual['saldo'] -= valor_gasto_restante
                        detalhe_abatimento.append(f"Abatido de '{credito_atual['desc'][:20]}...': R$ {valor_gasto_restante:,.2f}")
                        valor_gasto_restante = 0.0
                    else:
                        consumido = credito_atual['saldo']
                        valor_gasto_restante -= consumido
                        credito_atual['saldo'] = 0.0
                        detalhe_abatimento.append(f"Esgotou '{credito_atual['desc'][:20]}...': R$ {consumido:,.2f}")
                        creditos_ativos.pop(0)

                if valor_gasto_restante > 0:
                    despesas_nao_alocadas += valor_gasto_restante

                sinal_saida = f"- R$ {v_num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                soma_saldos_atuais = sum(c['saldo'] for c in creditos_ativos) - despesas_nao_alocadas
                s_rest_str = f"R$ {soma_saldos_atuais:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                
                linhas_balancete.append((data, f"[SAÍDA] {cat}", f"{historico} | " + " | ".join(detalhe_abatimento), "-", sinal_saida, s_rest_str))

        for l in linhas_balancete:
            tree_bal.insert("", "end", values=l)

        footer_frame = ttk.Frame(bal_janela, padding=10)
        footer_frame.pack(fill="x")

        saldo_real_livre = sum(c['saldo'] for c in creditos_ativos) - despesas_nao_alocadas
        t_sld_real = f"REAL SALDO DO QUE SOBROU PARA USO: R$ {saldo_real_livre:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

        lbl_real = tk.Label(footer_frame, text=t_sld_real, fg="#27ae60" if saldo_real_livre >= 0 else "#c0392b", font=("Helvetica", 11, "bold"))
        lbl_real.pack(anchor="w")

    def gerar_extrato_conta_corrente(self):
        itens = self.tabela.get_children()
        if not itens:
            messagebox.showwarning("Aviso", "Não há movimentações para gerar o extrato.")
            return

        extrato_janela = tk.Toplevel(self.root)
        extrato_janela.title("Extrato de Movimentações - Conta Corrente")
        extrato_janela.geometry("1000x600")

        header_frame = ttk.Frame(extrato_janela, padding=10)
        header_frame.pack(fill="x")

        tk.Label(header_frame, text="EXTRATO BANCÁRIO DE CONTA CORRENTE", font=("Helvetica", 14, "bold"), fg="#2c3e50").pack(anchor="w")
        tk.Label(header_frame, text=f"Emitido em: {datetime.now().strftime('%d/%m/%Y às %H:%M')}", font=("Helvetica", 9), fg="#7f8c8d").pack(anchor="w")

        tabela_frame_ext = ttk.Frame(extrato_janela)
        tabela_frame_ext.pack(fill="both", expand=True, padx=10, pady=5)

        cols = ("Data", "Categoria", "Histórico", "Tipo", "Valor (R$)", "Saldo Parcial")
        tree_extrato = ttk.Treeview(tabela_frame_ext, columns=cols, show="headings", selectmode="none")

        for col in cols:
            tree_extrato.heading(col, text=col)
            if col in ["Valor (R$)", "Saldo Parcial"]:
                tree_extrato.column(col, width=105, anchor="e")
            elif col in ["Tipo", "Data"]:
                tree_extrato.column(col, width=110, anchor="center")
            elif col == "Categoria":
                tree_extrato.column(col, width=130, anchor="center")
            else:
                tree_extrato.column(col, width=200, anchor="w")

        scroll_ext = ttk.Scrollbar(tabela_frame_ext, orient="vertical", command=tree_extrato.yview)
        tree_extrato.configure(yscrollcommand=scroll_ext.set)

        tree_extrato.pack(side="left", fill="both", expand=True)
        scroll_ext.pack(side="right", fill="y")

        saldo_acumulado = 0.0
        total_entradas = 0.0
        total_saidas = 0.0

        for linha in itens:
            vals = self.tabela.item(linha, "values")
            data, tipo, cat, val_str, desc, obs = vals

            limpo = val_str.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
            try:
                v_num = float(limpo)
            except ValueError:
                v_num = 0.0

            if tipo == "Entrada":
                saldo_acumulado += v_num
                total_entradas += v_num
                sinal_val = f"+ R$ {v_num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            else:
                saldo_acumulado -= v_num
                total_saidas += v_num
                sinal_val = f"- R$ {v_num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

            s_parcial = f"R$ {saldo_acumulado:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            historico = f"{desc} {f'({obs})' if obs else ''}"

            tree_extrato.insert("", "end", values=(data, cat, historico, tipo, sinal_val, s_parcial))

        footer_frame = ttk.Frame(extrato_janela, padding=10)
        footer_frame.pack(fill="x")

        t_ent = f"Total Entradas: R$ {total_entradas:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        t_sai = f"Total Saídas: R$ {total_saidas:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        t_sld = f"Saldo Final em Conta: R$ {saldo_acumulado:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

        tk.Label(footer_frame, text=t_ent, fg="#27ae60", font=("Helvetica", 10, "bold")).pack(anchor="w")
        tk.Label(footer_frame, text=t_sai, fg="#c0392b", font=("Helvetica", 10, "bold")).pack(anchor="w")
        tk.Label(footer_frame, text=t_sld, fg="#2c3e50", font=("Helvetica", 11, "bold")).pack(anchor="w")

    def salvar_dados_csv(self):
        try:
            with open(ARQUIVO_CSV, mode="w", newline="", encoding="utf-8") as f:
                escritor = csv.writer(f)
                escritor.writerow(["Data", "Tipo", "Categoria", "Valor", "Descrição", "Observação"])
                for linha in self.tabela.get_children():
                    escritor.writerow(self.tabela.item(linha, "values"))
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar arquivo CSV:\n{str(e)}")

    def carregar_dados(self):
        if os.path.exists(ARQUIVO_CSV):
            try:
                with open(ARQUIVO_CSV, mode="r", encoding="utf-8") as f:
                    leitor = csv.reader(f)
                    next(leitor, None)
                    for linha in leitor:
                        if linha:
                            if len(linha) == 5:
                                d, t, v, desc, obs = linha
                                linha = [d, t, "Geral", v, desc, obs]
                            self.tabela.insert("", "end", values=linha)
                self.atualizar_saldo()
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao carregar dados salvos:\n{str(e)}")

    def gerar_pdf(self):
        if not self.tabela.get_children():
            messagebox.showwarning("Aviso", "Não há dados suficientes para gerar o relatório PDF.")
            return

        arquivo_caminho = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("Arquivos PDF", "*.pdf"), ("Todos os arquivos", "*.*")],
            initialfile="relatorio_fluxo_caixa.pdf",
            title="Salvar Relatório PDF"
        )
        
        if not arquivo_caminho:
            return

        try:
            doc = SimpleDocTemplate(arquivo_caminho, pagesize=landscape(letter), rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
            elementos = []
            estilos = getSampleStyleSheet()

            titulo_estilo = ParagraphStyle('TituloRelatorio', parent=estilos['Heading1'], fontSize=16, alignment=1, spaceAfter=20)
            elementos.append(Paragraph("Relatório de Fluxo de Caixa", titulo_estilo))

            estilo_celula = ParagraphStyle('CelulaTabela', parent=estilos['Normal'], fontSize=8, leading=10, fontName='Helvetica')
            estilo_cabecalho = ParagraphStyle('CabecalhoTabela', parent=estilos['Normal'], fontSize=9, leading=11, fontName='Helvetica-Bold', textColor=colors.whitesmoke, alignment=1)

            dados = [[Paragraph(h, estilo_cabecalho) for h in ["Data", "Tipo", "Categoria", "Valor", "Descrição", "Observação"]]]

            for linha in self.tabela.get_children():
                vals = self.tabela.item(linha, "values")
                linha_paragrafos = [Paragraph(str(v), estilo_celula) for v in vals]
                dados.append(linha_paragrafos)

            tabela_pdf = Table(dados, colWidths=[110, 80, 130, 95, 165, 150])
            tabela_pdf.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2c3e50')),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('BOTTOMPADDING', (0,0), (-1,0), 8),
                ('TOPPADDING', (0,0), (-1,0), 8),
                ('BACKGROUND', (0,1), (-1,-1), colors.HexColor('#f9f9f9')),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#d3d3d3')),
                ('BOTTOMPADDING', (0,1), (-1,-1), 6),
                ('TOPPADDING', (0,1), (-1,-1), 6),
            ]))

            elementos.append(tabela_pdf)
            doc.build(elementos)

            messagebox.showinfo("Sucesso", f"Relatório PDF gerado com sucesso!\nSalvo em:\n{arquivo_caminho}")

            try:
                if sys.platform == "win32":
                    os.startfile(arquivo_caminho)
                elif sys.platform == "darwin":
                    subprocess.call(["open", arquivo_caminho])
                else:
                    subprocess.call(["xdg-open", arquivo_caminho])
            except Exception as e:
                print(f"Não foi possível abrir o arquivo automaticamente: {e}")

        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao gerar o PDF:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = AppFluxoCaixa(root)
    root.mainloop()