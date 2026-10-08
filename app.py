import csv
import os
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
from datetime import datetime
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import subprocess
import sys

ARQUIVO_CSV = "fluxo_de_caixa_continuo.csv"


class AppFluxoCaixa:
    def __init__(self, root):
        self.root = root
        self.root.title("Controle de Fluxo de Caixa - Conta Corrente")
        self.root.geometry("1150x820")
        self.root.minsize(1000, 700)

        self.style = ttk.Style()
        self.style.theme_use("clam")

        self.tipo_var = tk.StringVar(value="Entrada")
        self.categoria_var = tk.StringVar(value="Geral")
        self.valor_var = tk.StringVar()
        self.data_var = tk.StringVar(value=datetime.now().strftime("%d/%m/%Y %H:%M"))
        self.descricao_var = tk.StringVar()
        self.obs_var = tk.StringVar()

        self.criar_widgets()
        self.carregar_dados()

    def criar_widgets(self):
        titulo_label = tk.Label(
            self.root,
            text="CONTROLE DE FLUXO DE CAIXA - CONTA CORRENTE",
            font=("Helvetica", 18, "bold"),
            fg="#2c3e50"
        )
        titulo_label.pack(pady=10)

        form_frame = ttk.LabelFrame(self.root, text=" Novo Lançamento ", padding=15)
        form_frame.pack(fill="x", padx=20, pady=5)

        ttk.Label(form_frame, text="Tipo:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.tipo_combo = ttk.Combobox(
            form_frame,
            textvariable=self.tipo_var,
            values=["Entrada", "Saída"],
            state="readonly",
            width=15
        )
        self.tipo_combo.grid(row=0, column=1, sticky="w", padx=5, pady=5)
        self.tipo_combo.bind("<<ComboboxSelected>>", self.atualizar_categorias)

        ttk.Label(form_frame, text="Categoria:").grid(row=0, column=2, sticky="w", padx=5, pady=5)
        self.cat_combo = ttk.Combobox(form_frame, textvariable=self.categoria_var, state="readonly", width=25)
        self.cat_combo.grid(row=0, column=3, sticky="w", padx=5, pady=5)

        ttk.Label(form_frame, text="Valor (R$):").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        valor_entry = ttk.Entry(form_frame, textvariable=self.valor_var, width=15)
        valor_entry.grid(row=1, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(form_frame, text="Data e Hora:").grid(row=1, column=2, sticky="w", padx=5, pady=5)
        data_entry = ttk.Entry(form_frame, textvariable=self.data_var, width=20)
        data_entry.grid(row=1, column=3, sticky="w", padx=5, pady=5)
        tk.Label(
            form_frame,
            text="(DD/MM/AAAA HH:MM)",
            font=("Helvetica", 8),
            fg="#7f8c8d"
        ).grid(row=1, column=4, sticky="w", padx=5, pady=5)

        ttk.Label(form_frame, text="Descrição:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.desc_entry = ttk.Entry(form_frame, textvariable=self.descricao_var, width=50)
        self.desc_entry.grid(row=2, column=1, columnspan=4, sticky="w", padx=5, pady=5)

        ttk.Label(form_frame, text="Observação:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        obs_entry = ttk.Entry(form_frame, textvariable=self.obs_var, width=50)
        obs_entry.grid(row=3, column=1, columnspan=4, sticky="w", padx=5, pady=5)

        btn_frame = ttk.Frame(form_frame)
        btn_frame.grid(row=4, column=0, columnspan=5, pady=10)
        ttk.Button(btn_frame, text="Adicionar Lançamento", command=self.adicionar_registro).pack(side="left", padx=3)
        ttk.Button(btn_frame, text="Editar Selecionado", command=self.carregar_registro_para_edicao).pack(side="left", padx=3)
        ttk.Button(btn_frame, text="Excluir Selecionado", command=self.excluir_registro).pack(side="left", padx=3)
        ttk.Button(btn_frame, text="Excluir Todos", command=self.excluir_todos_registros).pack(side="left", padx=3)
        ttk.Button(btn_frame, text="Extrato Único", command=self.gerar_extrato_unico).pack(side="left", padx=3)
        ttk.Button(btn_frame, text="Relatório PDF", command=self.gerar_pdf).pack(side="left", padx=3)

        tabela_frame = ttk.Frame(self.root)
        tabela_frame.pack(fill="both", expand=True, padx=20, pady=5)
        colunas = ("Data", "Tipo", "Categoria", "Valor", "Descrição", "Observação")
        self.tabela = ttk.Treeview(tabela_frame, columns=colunas, show="headings", selectmode="browse")
        self.tabela.bind("<Double-1>", lambda event: self.carregar_registro_para_edicao())
        self.tabela.tag_configure("credito", foreground="#27ae60")
        self.tabela.tag_configure("debito", foreground="#c0392b")
        for col in colunas:
            self.tabela.heading(col, text=col)
            if col == "Valor":
                self.tabela.column(col, width=120, anchor="e")
            elif col in ["Tipo", "Categoria"]:
                self.tabela.column(col, width=150, anchor="center")
            elif col == "Data":
                self.tabela.column(col, width=150, anchor="center")
            else:
                self.tabela.column(col, width=220, anchor="w")
        scrollbar = ttk.Scrollbar(tabela_frame, orient="vertical", command=self.tabela.yview)
        self.tabela.configure(yscrollcommand=scrollbar.set)
        self.tabela.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        resumo_frame = ttk.LabelFrame(self.root, text=" SALDO DA CONTA CORRENTE ", padding=12)
        resumo_frame.pack(fill="x", padx=20, pady=5)
        
        self.lbl_saldo = tk.Label(
            resumo_frame,
            text="Saldo Atual: R$ 0,00",
            font=("Helvetica", 14, "bold"),
            fg="#27ae60"
        )
        self.lbl_saldo.pack(anchor="w", padx=10)

        self.lbl_saldo_apos = tk.Label(
            resumo_frame,
            text="",
            font=("Helvetica", 12, "bold", "italic"),
            fg="#2980b9"
        )
        self.lbl_saldo_apos.pack(anchor="w", padx=10, pady=5)

    def atualizar_categorias(self, event=None):
        tipo = self.tipo_var.get()
        if tipo == "Entrada":
            cats = ["Salário", "Rendimento", "Resgate", "Reembolso", "Outras Receitas"]
        else:
            cats = ["Conta", "Cartão", "Empréstimo", "Alimentação", "Transporte", "Moradia", "Outras Despesas"]
        self.cat_combo['values'] = cats
        self.categoria_var.set(cats[0])

    def calcular_saldo_atual(self):
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
        return saldo

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
        valor_str = valor_bruto.replace("R$", "").strip()
        if "," in valor_str:
            valor_str = valor_str.replace(".", "").replace(",", ".")
        try:
            valor = float(valor_str)
            if valor <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Erro", "Insira um valor válido (ex: 1500,50)!")
            return
        obs = self.obs_var.get().strip()
        valor_formatado = f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        tag_cor = "credito" if tipo == "Entrada" else "debito"
        
        self.tabela.insert("", "end", values=(data_lancamento, tipo, categoria, valor_formatado, descricao, obs), tags=(tag_cor,))
        self.salvar_dados_csv()
        
        self.valor_var.set("")
        self.descricao_var.set("")
        self.obs_var.set("")
        self.data_var.set(datetime.now().strftime("%d/%m/%Y %H:%M"))
        
        self.atualizar_saldo_exibicao()
        
        saldo_atual = self.calcular_saldo_atual()
        saldo_str = f"✅ Saldo após lançamento: R$ {saldo_atual:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        self.lbl_saldo_apos.config(
            text=saldo_str,
            fg="#27ae60" if saldo_atual >= 0 else "#c0392b"
        )

    def carregar_registro_para_edicao(self):
        selecionado = self.tabela.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um registro na tabela para editar.")
            return
        item_id = selecionado[0]
        vals = self.tabela.item(item_id, "values")
        data, tipo, cat, val_str, desc, obs = vals
        
        self.edit_janela = tk.Toplevel(self.root)
        self.edit_janela.title("Editar Lançamento")
        self.edit_janela.geometry("500x400")
        self.edit_janela.grab_set()
        
        ttk.Label(self.edit_janela, text="Editar Dados", font=("Helvetica", 12, "bold")).pack(pady=10)
        frame_edicao = ttk.Frame(self.edit_janela, padding=15)
        frame_edicao.pack(fill="both", expand=True)
        
        ttk.Label(frame_edicao, text="Data e Hora:").grid(row=0, column=0, sticky="w", pady=5)
        e_data_var = tk.StringVar(value=data)
        ttk.Entry(frame_edicao, textvariable=e_data_var, width=30).grid(row=0, column=1, sticky="w", pady=5)
        
        ttk.Label(frame_edicao, text="Tipo:").grid(row=1, column=0, sticky="w", pady=5)
        e_tipo_var = tk.StringVar(value=tipo)
        e_tipo_combo = ttk.Combobox(frame_edicao, textvariable=e_tipo_var, values=["Entrada", "Saída"], state="readonly", width=27)
        e_tipo_combo.grid(row=1, column=1, sticky="w", pady=5)
        
        ttk.Label(frame_edicao, text="Categoria:").grid(row=2, column=0, sticky="w", pady=5)
        e_cat_var = tk.StringVar(value=cat)
        e_cat_combo = ttk.Combobox(frame_edicao, textvariable=e_cat_var, values=self.cat_combo['values'], state="readonly", width=27)
        e_cat_combo.grid(row=2, column=1, sticky="w", pady=5)
        
        ttk.Label(frame_edicao, text="Valor:").grid(row=3, column=0, sticky="w", pady=5)
        limpa_val = val_str.replace("R$", "").strip()
        e_val_var = tk.StringVar(value=limpa_val)
        ttk.Entry(frame_edicao, textvariable=e_val_var, width=30).grid(row=3, column=1, sticky="w", pady=5)
        
        ttk.Label(frame_edicao, text="Descrição:").grid(row=4, column=0, sticky="w", pady=5)
        e_desc_var = tk.StringVar(value=desc)
        ttk.Entry(frame_edicao, textvariable=e_desc_var, width=30).grid(row=4, column=1, sticky="w", pady=5)
        
        ttk.Label(frame_edicao, text="Observação:").grid(row=5, column=0, sticky="w", pady=5)
        e_obs_var = tk.StringVar(value=obs)
        ttk.Entry(frame_edicao, textvariable=e_obs_var, width=30).grid(row=5, column=1, sticky="w", pady=5)
        
        def salvar_edicao():
            nova_data = e_data_var.get().strip()
            novo_tipo = e_tipo_var.get()
            nova_cat = e_cat_var.get()
            novo_val_bruto = e_val_var.get().strip()
            nova_desc = e_desc_var.get().strip()
            nova_obs = e_obs_var.get().strip()
            
            if not nova_data or not novo_val_bruto or not nova_desc:
                messagebox.showerror("Erro", "Preencha todos os campos obrigatórios!", parent=self.edit_janela)
                return
            
            v_str = novo_val_bruto.replace("R$", "").strip()
            if "," in v_str:
                v_str = v_str.replace(".", "").replace(",", ".")
            try:
                v_num = float(v_str)
                if v_num <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Erro", "Valor inválido!", parent=self.edit_janela)
                return
            
            novo_val_formatado = f"R$ {v_num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            tag_cor = "credito" if novo_tipo == "Entrada" else "debito"
            
            self.tabela.item(item_id, values=(nova_data, novo_tipo, nova_cat, novo_val_formatado, nova_desc, nova_obs), tags=(tag_cor,))
            self.salvar_dados_csv()
            self.atualizar_saldo_exibicao()
            
            saldo_atual = self.calcular_saldo_atual()
            saldo_str = f"✅ Saldo após alteração: R$ {saldo_atual:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            self.lbl_saldo_apos.config(
                text=saldo_str,
                fg="#27ae60" if saldo_atual >= 0 else "#c0392b"
            )
            
            messagebox.showinfo("Sucesso", "Lançamento atualizado!", parent=self.edit_janela)
            self.edit_janela.destroy()
        
        ttk.Button(frame_edicao, text="Salvar", command=salvar_edicao).grid(row=6, column=0, columnspan=2, pady=15)

    def excluir_registro(self):
        selecionado = self.tabela.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um registro para excluir.")
            return
        if messagebox.askyesno("Confirmação", "Excluir este lançamento?"):
            self.tabela.delete(selecionado)
            self.salvar_dados_csv()
            self.atualizar_saldo_exibicao()
            
            saldo_atual = self.calcular_saldo_atual()
            saldo_str = f"✅ Saldo após exclusão: R$ {saldo_atual:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            self.lbl_saldo_apos.config(
                text=saldo_str,
                fg="#27ae60" if saldo_atual >= 0 else "#c0392b"
            )

    def excluir_todos_registros(self):
        if not self.tabela.get_children():
            messagebox.showinfo("Aviso", "A tabela já está vazia.")
            return
        if messagebox.askyesno("ATENÇÃO", "Excluir TODOS os lançamentos?"):
            for item in self.tabela.get_children():
                self.tabela.delete(item)
            self.salvar_dados_csv()
            self.atualizar_saldo_exibicao()
            self.lbl_saldo_apos.config(text="")
            messagebox.showinfo("Sucesso", "Todos os lançamentos foram apagados.")

    def atualizar_saldo_exibicao(self):
        saldo = self.calcular_saldo_atual()
        saldo_formatado = f"Saldo Atual: R$ {saldo:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        self.lbl_saldo.config(text=saldo_formatado)
        if saldo >= 0:
            self.lbl_saldo.config(fg="#27ae60")
        else:
            self.lbl_saldo.config(fg="#c0392b")

    def gerar_extrato_unico(self):
        itens = self.tabela.get_children()
        if not itens:
            messagebox.showwarning("Aviso", "Não há lançamentos para gerar o extrato.")
            return
        
        ext_janela = tk.Toplevel(self.root)
        ext_janela.title("EXTRATO ÚNICO - Conta Corrente")
        ext_janela.geometry("1000x650")
        
        header_frame = ttk.Frame(ext_janela, padding=15)
        header_frame.pack(fill="x")
        tk.Label(
            header_frame,
            text="EXTRATO ÚNICO CONSOLIDADO - CONTA CORRENTE",
            font=("Helvetica", 15, "bold"),
            fg="#2c3e50"
        ).pack(anchor="w")
        tk.Label(
            header_frame,
            text="Histórico completo de movimentações e saldo final",
            font=("Helvetica", 10),
            fg="#7f8c8d"
        ).pack(anchor="w")
        
        tabela_frame = ttk.Frame(ext_janela)
        tabela_frame.pack(fill="both", expand=True, padx=15, pady=10)
        
        cols = ("Data", "Categoria", "Descrição", "Entrada (+)", "Saída (-)", "Saldo Parcial")
        tree = ttk.Treeview(tabela_frame, columns=cols, show="headings", selectmode="none")
        
        for col in cols:
            tree.heading(col, text=col)
            if col in ["Entrada (+)", "Saída (-)", "Saldo Parcial"]:
                tree.column(col, width=130, anchor="e")
            elif col == "Data":
                tree.column(col, width=150, anchor="center")
            elif col == "Categoria":
                tree.column(col, width=160, anchor="center")
            else:
                tree.column(col, width=250, anchor="w")
        
        scroll = ttk.Scrollbar(tabela_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        
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
                col_entrada = f"+ R$ {v_num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                col_saida = "-"
            else:
                saldo_acumulado -= v_num
                total_saidas += v_num
                col_entrada = "-"
                col_saida = f"- R$ {v_num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            
            saldo_parcial_str = f"R$ {saldo_acumulado:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            historico = f"{desc} {f'| {obs}' if obs else ''}"
            
            tree.insert("", "end", values=(data, cat, historico, col_entrada, col_saida, saldo_parcial_str))
        
        footer_frame = ttk.Frame(ext_janela, padding=15)
        footer_frame.pack(fill="x")
        
        t_ent = f"📈 TOTAL DE ENTRADAS: R$ {total_entradas:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        t_sai = f"📉 TOTAL DE SAÍDAS: R$ {total_saidas:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        t_sld = f"💰 SALDO FINAL REALMENTE SOBROU: R$ {saldo_acumulado:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        
        tk.Label(footer_frame, text=t_ent, fg="#27ae60", font=("Helvetica", 11, "bold")).pack(anchor="w")
        tk.Label(footer_frame, text=t_sai, fg="#c0392b", font=("Helvetica", 11, "bold")).pack(anchor="w")
        tk.Label(
            footer_frame,
            text=t_sld,
            fg="#27ae60" if saldo_acumulado >= 0 else "#c0392b",
            font=("Helvetica", 13, "bold")
        ).pack(anchor="w", pady=8)

    def salvar_dados_csv(self):
        try:
            with open(ARQUIVO_CSV, mode="w", newline="", encoding="utf-8") as f:
                escritor = csv.writer(f)
                escritor.writerow(["Data", "Tipo", "Categoria", "Valor", "Descrição", "Observação"])
                for linha in self.tabela.get_children():
                    escritor.writerow(self.tabela.item(linha, "values"))
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar:\n{str(e)}")

    def carregar_dados(self):
        if os.path.exists(ARQUIVO_CSV):
            try:
                with open(ARQUIVO_CSV, mode="r", encoding="utf-8") as f:
                    leitor = csv.reader(f)
                    next(leitor, None)
                    for linha in leitor:
                        if linha and len(linha) >= 4:
                            if len(linha) == 5:
                                d, t, v, desc, obs = linha
                                linha = [d, t, "Geral", v, desc, obs]
                            tipo = linha[1]
                            tag_cor = "credito" if tipo == "Entrada" else "debito"
                            self.tabela.insert("", "end", values=linha, tags=(tag_cor,))
                    self.atualizar_saldo_exibicao()
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao carregar dados:\n{str(e)}")

    def gerar_pdf(self):
        if not self.tabela.get_children():
            messagebox.showwarning("Aviso", "Não há dados para gerar o PDF.")
            return
        
        caminho = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf"), ("Todos", "*.*")],
            initialfile="extrato_conta_corrente.pdf",
            title="Salvar Extrato em PDF"
        )
        if not caminho:
            return
        
        try:
            doc = SimpleDocTemplate(caminho, pagesize=landscape(letter), rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
            elementos = []
            estilos = getSampleStyleSheet()
            
            titulo_estilo = ParagraphStyle(
                'Titulo',
                parent=estilos['Heading1'],
                fontSize=16,
                alignment=1,
                spaceAfter=20
            )
            elementos.append(Paragraph("EXTRATO CONTA CORRENTE", titulo_estilo))
            
            cel_estilo = ParagraphStyle('Celula', parent=estilos['Normal'], fontSize=9, leading=11)
            cab_estilo = ParagraphStyle('Cabecalho', parent=estilos['Normal'], fontSize=10, fontName='Helvetica-Bold', textColor=colors.whitesmoke, alignment=1)
            
            dados = [[Paragraph(h, cab_estilo) for h in ["Data", "Tipo", "Categoria", "Valor", "Descrição", "Obs"]]]
            for linha in self.tabela.get_children():
                vals = self.tabela.item(linha, "values")
                dados.append([Paragraph(str(v), cel_estilo) for v in vals])
            
            tabela = Table(dados, colWidths=[110, 80, 130, 100, 200, 130])
            tabela.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2c3e50')),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                ('BOTTOMPADDING', (0,0), (-1,-1), 8),
                ('TOPPADDING', (0,0), (-1,-1), 8),
            ]))
            elementos.append(tabela)
            
            saldo_final = self.calcular_saldo_atual()
            resumo_estilo = ParagraphStyle('Resumo', parent=estilos['Normal'], fontSize=12, spaceBefore=20, fontName='Helvetica-Bold')
            cor = '#27ae60' if saldo_final >= 0 else '#c0392b'
            elementos.append(Paragraph(f"<font color='{cor}'>SALDO FINAL: R$ {saldo_final:,.2f}</font>", resumo_estilo))
            
            doc.build(elementos)
            messagebox.showinfo("Sucesso", f"PDF salvo em:\n{caminho}")
            
            try:
                if sys.platform == "win32":
                    os.startfile(caminho)
                elif sys.platform == "darwin":
                    subprocess.call(["open", caminho])
                else:
                    subprocess.call(["xdg-open", caminho])
            except:
                pass
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao gerar PDF:\n{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = AppFluxoCaixa(root)
    root.mainloop()