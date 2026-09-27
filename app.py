import csv
from datetime import datetime
import os
import subprocess
import sys
import tkinter as tk
from tkinter import messagebox, ttk
import urllib.parse
import webbrowser

# Importações para a geração do PDF
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

# Nome do arquivo de dados padrão
ARQUIVO_DADOS = "fluxo_de_caixa_continuo.csv"


class AppFluxoCaixa:

  def __init__(self, root):
    self.root = root
    self.root.title("Fluxo de Caixa Contínuo")
    self.root.geometry("1150x1000")
    self.root.config(bg="#f0f0f0")

    # Variáveis para controle das máscaras e edição
    self.raw_data_digits = ""
    self.raw_valor_digits = "0"
    self.id_em_edicao = None

    # Lista das despesas/tabulações solicitadas para Débitos (incluindo "DIVERSOS")
    self.sub_categorias = [
        "FINANCIAMENTO APTO",
        "CONDOMINIO",
        "COMPRAS PARCELADAS NO CARTAO",
        "FACULDADE",
        "DESPESAS MEDICAS",
        "VALE REFEIÇAO",
        "CONTA DE ENERGIA",
        "CONTA DE INTERNET",
        "SUPERMERCADO",
        "BARBEARIA",
        "AGUA POTAVEL",
        "AÇOUGUE",
        "FATURA VIVO",
        "RAÇAO DE GATO",
        "AREIA DE GATO",
        "VACINAS NOS GATOS",
        "DIVERSOS",
    ]

    # Lista de botões/operações solicitadas (Disponíveis em Crédito e Débito)
    self.categorias = [
        "PAGAMENTO RECEBIDO",
        "ADIANTAMENTO VALE",
        "SALARIO ANTES DE SAIR DE FERIAS",
        "FERIAS",
        "SALARIO 10 DIAS TRABALHADOS DAS FERIAS",
        "VALE APOS FERIAS",
        "PAGAMENTO APOS FERIAS",
        "DESPESAS PAGAS COM FERIAS",
        "RESERVA POUPANÇA",
        "BAIXA RESERVA POUPANCA",
        "SALDO APOS BAIXA RESERVA POUPANÇA",
        "IRPF RECEBIDO",
        "SALDO IRPF USADO",
        "LIQUIDO IRPF APOS USO",
    ]

    self.criar_componentes()
    self.carregar_dados()

  def criar_componentes(self):
    # Título
    titulo = tk.Label(
        self.root,
        text="Fluxo de Caixa Contínuo",
        font=("Arial", 15, "bold"),
        bg="#f0f0f0",
        fg="#333",
    )
    titulo.pack(pady=4)

    # Painel de Resumo do Saldo em tempo real
    self.frame_resumo = tk.LabelFrame(
        self.root,
        text=" Resumo do Caixa (Saldo Atual Após Abatimento) ",
        font=("Arial", 9, "bold"),
        bg="#e8f4f8",
    )
    self.frame_resumo.pack(fill="x", padx=10, pady=2)

    self.lbl_resumo_creditos = tk.Label(
        self.frame_resumo,
        text="Total Entradas: R$ 0,00",
        font=("Arial", 9, "bold"),
        bg="#e8f4f8",
        fg="#2e7d32",
    )
    self.lbl_resumo_creditos.pack(side="left", padx=15, pady=4)

    self.lbl_resumo_debitos = tk.Label(
        self.frame_resumo,
        text="Total Saídas: R$ 0,00",
        font=("Arial", 9, "bold"),
        bg="#e8f4f8",
        fg="#c62828",
    )
    self.lbl_resumo_debitos.pack(side="left", padx=15, pady=4)

    self.lbl_resumo_saldo = tk.Label(
        self.frame_resumo,
        text="Saldo Restante no Caixa: R$ 0,00",
        font=("Arial", 9, "bold"),
        bg="#e8f4f8",
        fg="#0d47a1",
    )
    self.lbl_resumo_saldo.pack(side="right", padx=15, pady=4)

    # Container Principal para os Painéis de Crédito e Débito lado a lado
    frame_lancamentos = tk.Frame(self.root, bg="#f0f0f0")
    frame_lancamentos.pack(fill="x", padx=10, pady=2)

    # Frame Esquerdo: Créditos
    frame_credito_container = tk.Frame(frame_lancamentos, bg="#f0f0f0")
    frame_credito_container.pack(side="left", fill="both", expand=True, padx=2)
    self.construir_painel_lancamento(frame_credito_container, "CRÉDITO")

    # Frame Direito: Débitos
    frame_debito_container = tk.Frame(frame_lancamentos, bg="#f0f0f0")
    frame_debito_container.pack(side="right", fill="both", expand=True, padx=2)
    self.construir_painel_lancamento(frame_debito_container, "DÉBITO")

    # Tabela de Histórico (Treeview)
    frame_tabela = tk.LabelFrame(
        self.root, text=" Histórico de Lançamentos ", font=("Arial", 9, "bold")
    )
    frame_tabela.pack(fill="both", expand=True, padx=10, pady=4)

    colunas = (
        "ID",
        "Data",
        "Fluxo",
        "Tipo / Operação",
        "Despesa / Detalhe",
        "Valor (R$)",
        "Observação",
    )
    self.tabela = ttk.Treeview(
        frame_tabela, columns=colunas, show="headings", selectmode="browse", height=7
    )

    for col in colunas:
      self.tabela.heading(col, text=col)
      if col == "ID":
        self.tabela.column(col, width=35, anchor="center")
      elif col == "Data":
        self.tabela.column(col, width=75, anchor="center")
      elif col == "Fluxo":
        self.tabela.column(col, width=65, anchor="center")
      elif col == "Tipo / Operação":
        self.tabela.column(col, width=160, anchor="w")
      elif col == "Despesa / Detalhe":
        self.tabela.column(col, width=140, anchor="w")
      elif col == "Valor (R$)":
        self.tabela.column(col, width=85, anchor="e")
      else:
        self.tabela.column(col, width=130, anchor="w")

    scrollbar = ttk.Scrollbar(
        frame_tabela, orient="vertical", command=self.tabela.yview
    )
    self.tabela.configure(yscrollcommand=scrollbar.set)

    self.tabela.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # Evento ao selecionar um item na tabela para edição
    self.tabela.bind("<<TreeviewSelect>>", self.carregar_para_edicao)

    # Frame Inferior para os botões de Ação Geral
    frame_acoes = tk.Frame(self.root, bg="#f0f0f0")
    frame_acoes.pack(fill="x", padx=10, pady=4)

    btn_pdf = tk.Button(
        frame_acoes,
        text="Gerar Relatório Geral em PDF",
        bg="#4CAF50",
        fg="white",
        font=("Arial", 9, "bold"),
        command=lambda: self.gerar_pdf(None),
    )
    btn_pdf.pack(side="left")

    self.btn_salvar_edicao = tk.Button(
        frame_acoes,
        text="Salvar Alterações da Edição",
        bg="#FF9800",
        fg="white",
        font=("Arial", 9, "bold"),
        state="disabled",
        command=self.salvar_edicao,
    )
    self.btn_salvar_edicao.pack(side="left", padx=10)

    btn_excluir = tk.Button(
        frame_acoes,
        text="Excluir Lançamento Selecionado",
        bg="#f44336",
        fg="white",
        font=("Arial", 9, "bold"),
        command=self.excluir_lancamento,
    )
    btn_excluir.pack(side="right")

  def construir_painel_lancamento(self, parent_frame, tipo_fluxo):
    frame_form = tk.LabelFrame(
        parent_frame,
        text=f" Lançamento de {tipo_fluxo} ",
        font=("Arial", 9, "bold"),
    )
    frame_form.pack(fill="x", padx=2, pady=2)

    self.criar_campos_formulario(frame_form, tipo_fluxo)

    frame_botoes = tk.LabelFrame(
        parent_frame,
        text=f" Selecione a Operação ({tipo_fluxo}) ",
        font=("Arial", 9, "bold"),
    )
    frame_botoes.pack(fill="x", padx=2, pady=2)

    for i, categoria in enumerate(self.categorias):
      linha = i // 2
      coluna = i % 2
      btn = tk.Button(
          frame_botoes,
          text=categoria,
          bg="#2196F3" if tipo_fluxo == "CRÉDITO" else "#E91E63",
          fg="white",
          font=("Arial", 7, "bold"),
          width=26,
          command=lambda cat=categoria, tf=tipo_fluxo: self.adicionar_lancamento(
              cat, tf
          ),
      )
      btn.grid(row=linha, column=coluna, padx=2, pady=2, sticky="ew")

    # Botão individual de PDF para cada painel
    btn_pdf_painel = tk.Button(
        parent_frame,
        text=f"📄 Gerar e Visualizar PDF ({tipo_fluxo})",
        bg="#009688",
        fg="white",
        font=("Arial", 8, "bold"),
        command=lambda tf=tipo_fluxo: self.gerar_pdf(filtro_fluxo=tf),
    )
    btn_pdf_painel.pack(fill="x", padx=2, pady=3)

  def criar_campos_formulario(self, frame_form, tipo_fluxo):
    tk.Label(frame_form, text="Data:", font=("Arial", 8)).grid(
        row=0, column=0, sticky="w", padx=3, pady=2
    )

    entry_data = tk.Entry(frame_form, font=("Arial", 8), width=11)
    entry_data.grid(row=0, column=1, sticky="w", padx=3, pady=2)
    hoje_str = datetime.now().strftime("%d/%m/%Y")
    entry_data.insert(0, hoje_str)
    entry_data.bind("<KeyRelease>", self.ao_digitar_data)

    tk.Label(frame_form, text="Valor (R$):", font=("Arial", 8)).grid(
        row=0, column=2, sticky="w", padx=3, pady=2
    )

    entry_valor = tk.Entry(frame_form, font=("Arial", 8), width=12)
    entry_valor.grid(row=0, column=3, sticky="w", padx=3, pady=2)
    entry_valor.insert(0, "0,00")
    entry_valor.bind("<KeyRelease>", self.ao_digitar_valor)

    combo_despesa = None
    if tipo_fluxo == "DÉBITO":
      tk.Label(frame_form, text="Despesa:", font=("Arial", 8)).grid(
          row=1, column=0, sticky="w", padx=3, pady=2
      )

      combo_despesa = ttk.Combobox(
          frame_form,
          values=self.sub_categorias,
          font=("Arial", 8),
          width=27,
          state="readonly",
      )
      combo_despesa.grid(row=1, column=1, columnspan=3, sticky="w", padx=3, pady=2)
      if self.sub_categorias:
        combo_despesa.set(self.sub_categorias[0])

      tk.Label(frame_form, text="Obs:", font=("Arial", 8)).grid(
          row=2, column=0, sticky="w", padx=3, pady=2
      )
      entry_obs = tk.Entry(frame_form, font=("Arial", 8), width=32)
      entry_obs.grid(row=2, column=1, columnspan=3, sticky="w", padx=3, pady=2)
    else:
      tk.Label(frame_form, text="Obs:", font=("Arial", 8)).grid(
          row=1, column=0, sticky="w", padx=3, pady=2
      )
      entry_obs = tk.Entry(frame_form, font=("Arial", 8), width=32)
      entry_obs.grid(row=1, column=1, columnspan=3, sticky="w", padx=3, pady=2)
      tk.Label(frame_form, text="", font=("Arial", 8)).grid(row=2, column=0, pady=2)

    if tipo_fluxo == "CRÉDITO":
      self.entry_data_c = entry_data
      self.entry_valor_c = entry_valor
      self.combo_despesa_c = None
      self.entry_obs_c = entry_obs
    else:
      self.entry_data_d = entry_data
      self.entry_valor_d = entry_valor
      self.combo_despesa_d = combo_despesa
      self.entry_obs_d = entry_obs

  def ao_digitar_data(self, event):
    if event.keysym in (
        "BackSpace",
        "Delete",
        "Left",
        "Right",
        "Tab",
        "Shift_L",
        "Shift_R",
    ):
      return
    widget = event.widget
    texto_atual = widget.get()
    digitos = "".join(filter(str.isdigit, texto_atual))[:8]
    formatado = ""
    for i, d in enumerate(digitos):
      if i == 2 or i == 4:
        formatado += "/"
      formatado += d
    widget.delete(0, tk.END)
    widget.insert(0, formatado)
    widget.icursor(tk.END)

  def ao_digitar_valor(self, event):
    if event.keysym == "BackSpace":
      if len(self.raw_valor_digits) > 1:
        self.raw_valor_digits = self.raw_valor_digits[:-1]
      else:
        self.raw_valor_digits = "0"
    else:
      char = event.char
      if char.isdigit():
        if self.raw_valor_digits == "0":
          self.raw_valor_digits = char
        else:
          self.raw_valor_digits += char

    try:
      cents_val = int(self.raw_valor_digits)
    except ValueError:
      cents_val = 0

    valor_float = cents_val / 100.0
    valor_fmt = (
        f"{valor_float:,.2f}"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )

    widget_focado = self.root.focus_get()
    if widget_focado == getattr(self, "entry_valor_c", None):
      self.entry_valor_c.delete(0, tk.END)
      self.entry_valor_c.insert(0, valor_fmt)
      self.entry_valor_c.icursor(tk.END)
    elif widget_focado == getattr(self, "entry_valor_d", None):
      self.entry_valor_d.delete(0, tk.END)
      self.entry_valor_d.insert(0, valor_fmt)
      self.entry_valor_d.icursor(tk.END)
    else:
      self.entry_valor_c.delete(0, tk.END)
      self.entry_valor_c.insert(0, valor_fmt)
      self.entry_valor_d.delete(0, tk.END)
      self.entry_valor_d.insert(0, valor_fmt)

  def adicionar_lancamento(self, tipo_operacao, tipo_fluxo):
    if self.id_em_edicao is not None:
      messagebox.showwarning(
          "Aviso",
          "Você está editando um item. Clique em 'Salvar Alterações da Edição'"
          " antes de criar um novo.",
      )
      return

    if tipo_fluxo == "CRÉDITO":
      data = self.entry_data_c.get().strip()
      valor_texto = self.entry_valor_c.get().strip()
      despesa = ""
      obs = self.entry_obs_c.get().strip()
    else:
      data = self.entry_data_d.get().strip()
      valor_texto = self.entry_valor_d.get().strip()
      despesa = self.combo_despesa_d.get().strip()
      obs = self.entry_obs_d.get().strip()

    if len(data) < 10:
      messagebox.showerror(
          "Erro", "A data está incompleta. Use o formato DD/MM/AAAA."
      )
      return

    try:
      val_limpo = valor_texto.replace(".", "").replace(",", ".")
      valor = float(val_limpo)
    except ValueError:
      messagebox.showerror("Erro", "O valor digitado é inválido.")
      return

    id_lancamento = datetime.now().strftime("%Y%m%d%H%M%S")

    self.tabela.insert(
        "",
        "end",
        values=(
            id_lancamento,
            data,
            tipo_fluxo,
            tipo_operacao,
            despesa,
            f"R$ {valor:,.2f}"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", "."),
            obs,
        ),
    )

    self.salvar_no_csv_novo(
        id_lancamento, data, tipo_fluxo, tipo_operacao, despesa, valor, obs
    )
    self.atualizar_resumo_caixa()
    self.limpar_campos()

    if tipo_fluxo == "DÉBITO":
      total_c, total_d, saldo = self.calcular_totais()
      saldo_fmt = (
          f"R$ {saldo:,.2f}"
          .replace(",", "X")
          .replace(".", ",")
          .replace("X", ".")
      )
      messagebox.showinfo(
          "Débito Lançado",
          f"Débito registrado com sucesso!\n\nNovo Saldo Restante no Caixa:"
          f" {saldo_fmt}",
      )
    else:
      messagebox.showinfo(
          "Sucesso",
          f"Lançamento de {tipo_fluxo} ({tipo_operacao}) registrado com"
          " sucesso!",
      )

  def salvar_no_csv_novo(self, id_l, data, fluxo, tipo, despesa, valor, obs):
    modo = "a" if os.path.exists(ARQUIVO_DADOS) else "w"
    with open(ARQUIVO_DADOS, mode=modo, newline="", encoding="utf-8") as arquivo:
      escritor = csv.writer(arquivo)
      if modo == "w":
        escritor.writerow(
            ["ID", "Data", "Fluxo", "Tipo", "Despesa", "Valor", "Observacao"]
        )
      escritor.writerow([id_l, data, fluxo, tipo, despesa, valor, obs])

  def carregar_dados(self):
    if os.path.exists(ARQUIVO_DADOS):
      with open(ARQUIVO_DADOS, mode="r", newline="", encoding="utf-8") as arquivo:
        leitor = csv.reader(arquivo)
        next(leitor, None)
        for linha in leitor:
          if len(linha) == 7:
            id_l, data, fluxo, tipo, despesa, valor, obs = linha
            try:
              val_float = float(valor)
              valor_formatado = (
                  f"R$ {val_float:,.2f}"
                  .replace(",", "X")
                  .replace(".", ",")
                  .replace("X", ".")
              )
            except:
              valor_formatado = valor
            self.tabela.insert(
                "",
                "end",
                values=(
                    id_l,
                    data,
                    fluxo,
                    tipo,
                    despesa,
                    valor_formatado,
                    obs,
                ),
            )
          elif len(linha) == 6:
            id_l, data, tipo, despesa, valor, obs = linha
            fluxo = "CRÉDITO"
            try:
              val_float = float(valor)
              valor_formatado = (
                  f"R$ {val_float:,.2f}"
                  .replace(",", "X")
                  .replace(".", ",")
                  .replace("X", ".")
              )
            except:
              valor_formatado = valor
            self.tabela.insert(
                "",
                "end",
                values=(
                    id_l,
                    data,
                    fluxo,
                    tipo,
                    despesa,
                    valor_formatado,
                    obs,
                ),
            )
    self.atualizar_resumo_caixa()

  def calcular_totais(self):
    total_creditos = 0.0
    total_debitos = 0.0
    for child in self.tabela.get_children():
      vals = self.tabela.item(child)["values"]
      fluxo = vals[2]
      val_str = vals[5]
      try:
        val_limpo = (
            val_str.replace("R$", "")
            .replace(".", "")
            .replace(",", ".")
            .strip()
        )
        val_float = float(val_limpo)
      except:
        val_float = 0.0

      if fluxo == "CRÉDITO":
        total_creditos += val_float
      else:
        total_debitos += val_float

    saldo_restante = total_creditos - total_debitos
    return total_creditos, total_debitos, saldo_restante

  def atualizar_resumo_caixa(self):
    t_cred, t_deb, saldo = self.calcular_totais()
    c_fmt = (
        f"R$ {t_cred:,.2f}"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )
    d_fmt = (
        f"R$ {t_deb:,.2f}"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )
    s_fmt = (
        f"R$ {saldo:,.2f}"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )

    self.lbl_resumo_creditos.config(text=f"Total Entradas: {c_fmt}")
    self.lbl_resumo_debitos.config(text=f"Total Saídas: {d_fmt}")
    self.lbl_resumo_saldo.config(text=f"Saldo Restante no Caixa: {s_fmt}")

    if saldo < 0:
      self.lbl_resumo_saldo.config(fg="#c62828")
    else:
      self.lbl_resumo_saldo.config(fg="#0d47a1")

  def carregar_para_edicao(self, event):
    selecionado = self.tabela.selection()
    if not selecionado:
      return

    item = self.tabela.item(selecionado)
    valores = item["values"]

    self.id_em_edicao = str(valores[0])
    data = valores[1]
    fluxo = valores[2]
    despesa = valores[4]
    val_str = valores[5]
    obs = valores[6]

    if fluxo == "CRÉDITO":
      self.entry_data_c.delete(0, tk.END)
      self.entry_data_c.insert(0, data)
      self.entry_valor_c.delete(0, tk.END)
      self.entry_valor_c.insert(0, val_str.replace("R$", "").strip())
      self.entry_obs_c.delete(0, tk.END)
      self.entry_obs_c.insert(0, obs if obs != "None" else "")
    else:
      self.entry_data_d.delete(0, tk.END)
      self.entry_data_d.insert(0, data)
      if despesa in self.sub_categorias:
        self.combo_despesa_d.set(despesa)
      self.entry_valor_d.delete(0, tk.END)
      self.entry_valor_d.insert(0, val_str.replace("R$", "").strip())
      self.entry_obs_d.delete(0, tk.END)
      self.entry_obs_d.insert(0, obs if obs != "None" else "")

    val_limpo = (
        val_str.replace("R$", "").replace(".", "").replace(",", ".").strip()
    )
    try:
      val_float = float(val_limpo)
      self.raw_valor_digits = str(int(val_float * 100))
    except:
      self.raw_valor_digits = "0"

    self.btn_salvar_edicao.config(state="normal")

  def salvar_edicao(self):
    if not self.id_em_edicao:
      return

    selecionado = self.tabela.selection()
    if not selecionado:
      return
    item = self.tabela.item(selecionado)
    fluxo_atual = item["values"][2]
    tipo_atual = item["values"][3]

    if fluxo_atual == "CRÉDITO":
      nova_data = self.entry_data_c.get().strip()
      nova_despesa = ""
      valor_texto = self.entry_valor_c.get().strip()
      nova_obs = self.entry_obs_c.get().strip()
    else:
      nova_data = self.entry_data_d.get().strip()
      nova_despesa = self.combo_despesa_d.get().strip()
      valor_texto = self.entry_valor_d.get().strip()
      nova_obs = self.entry_obs_d.get().strip()

    if len(nova_data) < 10:
      messagebox.showerror("Erro", "A data está incompleta.")
      return

    try:
      val_limpo = valor_texto.replace(".", "").replace(",", ".")
      novo_valor = float(val_limpo)
    except ValueError:
      messagebox.showerror("Erro", "Valor inválido.")
      return

    valor_fmt = (
        f"R$ {novo_valor:,.2f}"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )
    self.tabela.item(
        selecionado,
        values=(
            self.id_em_edicao,
            nova_data,
            fluxo_atual,
            tipo_atual,
            nova_despesa,
            valor_fmt,
            nova_obs,
        ),
    )

    linhas = []
    if os.path.exists(ARQUIVO_DADOS):
      with open(ARQUIVO_DADOS, mode="r", newline="", encoding="utf-8") as arquivo:
        leitor = csv.reader(arquivo)
        cabecalho = next(leitor, None)
        for linha in leitor:
          if linha:
            if str(linha[0]) == str(self.id_em_edicao):
              linhas.append([
                  self.id_em_edicao,
                  nova_data,
                  fluxo_atual,
                  linha[3],
                  nova_despesa,
                  novo_valor,
                  nova_obs,
              ])
            else:
              linhas.append(linha)

      with open(ARQUIVO_DADOS, mode="w", newline="", encoding="utf-8") as arquivo:
        escritor = csv.writer(arquivo)
        if cabecalho:
          escritor.writerow(
              ["ID", "Data", "Fluxo", "Tipo", "Despesa", "Valor", "Observacao"]
          )
        escritor.writerows(linhas)

    self.atualizar_resumo_caixa()
    messagebox.showinfo("Sucesso", "Lançamento editado com sucesso!")
    self.limpar_campos()

  def excluir_lancamento(self):
    selecionado = self.tabela.selection()
    if not selecionado:
      messagebox.showwarning(
          "Aviso", "Selecione um lançamento na tabela para excluir."
      )
      return

    item = self.tabela.item(selecionado)
    id_alvo = item["values"][0]

    self.tabela.delete(selecionado)

    linhas_mantidas = []
    if os.path.exists(ARQUIVO_DADOS):
      with open(ARQUIVO_DADOS, mode="r", newline="", encoding="utf-8") as arquivo:
        leitor = csv.reader(arquivo)
        cabecalho = next(leitor, None)
        for linha in leitor:
          if linha and str(linha[0]) != str(id_alvo):
            linhas_mantidas.append(linha)

      with open(ARQUIVO_DADOS, mode="w", newline="", encoding="utf-8") as arquivo:
        escritor = csv.writer(arquivo)
        if cabecalho:
          escritor.writerow(
              ["ID", "Data", "Fluxo", "Tipo", "Despesa", "Valor", "Observacao"]
          )
        escritor.writerows(linhas_mantidas)

    self.atualizar_resumo_caixa()
    self.limpar_campos()
    messagebox.showinfo("Sucesso", "Lançamento excluído com sucesso!")

  def limpar_campos(self):
    self.id_em_edicao = None
    self.btn_salvar_edicao.config(state="disabled")

    hoje = datetime.now().strftime("%d/%m/%Y")
    for entry in [
        getattr(self, "entry_data_c", None),
        getattr(self, "entry_data_d", None),
    ]:
      if entry:
        entry.delete(0, tk.END)
        entry.insert(0, hoje)

    for entry in [
        getattr(self, "entry_valor_c", None),
        getattr(self, "entry_valor_d", None),
    ]:
      if entry:
        entry.delete(0, tk.END)
        entry.insert(0, "0,00")

    if getattr(self, "combo_despesa_d", None) and self.sub_categorias:
      self.combo_despesa_d.set(self.sub_categorias[0])

    for entry in [
        getattr(self, "entry_obs_c", None),
        getattr(self, "entry_obs_d", None),
    ]:
      if entry:
        entry.delete(0, tk.END)

    self.raw_valor_digits = "0"

  def gerar_pdf(self, filtro_fluxo=None):
    if not os.path.exists(ARQUIVO_DADOS):
      messagebox.showwarning("Aviso", "Nenhum dado encontrado para gerar o PDF.")
      return

    try:
      if filtro_fluxo:
        nome_arquivo_pdf = f"relatorio_fluxo_{filtro_fluxo.lower()}.pdf"
      else:
        nome_arquivo_pdf = "relatorio_fluxo_de_caixa.pdf"

      doc = SimpleDocTemplate(
          nome_arquivo_pdf,
          pagesize=letter,
          rightMargin=20,
          leftMargin=20,
          topMargin=25,
          bottomMargin=25,
      )
      elementos = []

      styles = getSampleStyleSheet()
      if filtro_fluxo:
        titulo_texto = f"Relatório - Fluxo de Caixa ({filtro_fluxo})"
      else:
        titulo_texto = "Relatório - Fluxo de Caixa Contínuo"

      titulo_estilo = ParagraphStyle(
          "TituloEstilo",
          parent=styles["Heading1"],
          fontSize=16,
          alignment=1,
          textColor=colors.HexColor(
              "#2196F3"
              if not filtro_fluxo
              else (
                  "#2e7d32" if filtro_fluxo == "CRÉDITO" else "#c62828"
              )
          ),
          spaceAfter=10,
      )

      elementos.append(Paragraph(titulo_texto, titulo_estilo))
      elementos.append(
          Paragraph(
              f"Gerado em: {datetime.now().strftime('%d/%m/%Y às %H:%M')}",
              styles["Normal"],
          )
      )
      elementos.append(Spacer(1, 10))

      t_cred, t_deb, saldo = self.calcular_totais()
      c_fmt = (
          f"R$ {t_cred:,.2f}"
          .replace(",", "X")
          .replace(".", ",")
          .replace("X", ".")
      )
      d_fmt = (
          f"R$ {t_deb:,.2f}"
          .replace(",", "X")
          .replace(".", ",")
          .replace("X", ".")
      )
      s_fmt = (
          f"R$ {saldo:,.2f}"
          .replace(",", "X")
          .replace(".", ",")
          .replace("X", ".")
      )

      resumo_texto = f"<b>Total de Entradas:</b> {c_fmt} | <b>Total de Saídas:</b> {d_fmt} | <b>Saldo Restante:</b> {s_fmt}"
      elementos.append(Paragraph(resumo_texto, styles["Normal"]))
      elementos.append(Spacer(1, 10))

      dados_tabela = [[
          "Data",
          "Fluxo",
          "Tipo / Operação",
          "Despesa / Detalhe",
          "Valor (R$)",
          "Observação",
      ]]

      with open(ARQUIVO_DADOS, mode="r", newline="", encoding="utf-8") as arquivo:
        leitor = csv.reader(arquivo)
        next(leitor, None)
        for linha in leitor:
          if len(linha) == 7:
            _, data, fluxo, tipo, despesa, valor, obs = linha
            if filtro_fluxo and fluxo != filtro_fluxo:
              continue
            try:
              val_float = float(valor)
              valor_fmt = (
                  f"R$ {val_float:,.2f}"
                  .replace(",", "X")
                  .replace(".", ",")
                  .replace("X", ".")
              )
            except:
              valor_fmt = valor
            dados_tabela.append([data, fluxo, tipo, despesa, valor_fmt, obs])
          elif len(linha) == 6:
            _, data, tipo, despesa, valor, obs = linha
            fluxo = "CRÉDITO"
            if filtro_fluxo and fluxo != filtro_fluxo:
              continue
            try:
              val_float = float(valor)
              valor_fmt = (
                  f"R$ {val_float:,.2f}"
                  .replace(",", "X")
                  .replace(".", ",")
                  .replace("X", ".")
              )
            except:
              valor_fmt = valor
            dados_tabela.append([data, fluxo, tipo, despesa, valor_fmt, obs])

      if len(dados_tabela) == 1:
        messagebox.showwarning(
            "Aviso", f"Nenhum registro encontrado para o filtro: {filtro_fluxo}"
        )
        return

      tabela_pdf = Table(dados_tabela, colWidths=[60, 65, 120, 110, 80, 115])
      tabela_pdf.setStyle(
          TableStyle([
              ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2196F3")),
              ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
              ("ALIGN", (0, 0), (-1, -1), "LEFT"),
              ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
              ("FONTSIZE", (0, 0), (-1, 0), 9),
              ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
              ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#f9f9f9")),
              ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#d3d3d3")),
              ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
              ("FONTSIZE", (0, 1), (-1, -1), 8),
          ])
      )

      elementos.append(tabela_pdf)
      doc.build(elementos)

      # 1. Abre o arquivo PDF automaticamente para visualização
      caminho_absoluto = os.path.abspath(nome_arquivo_pdf)
      try:
        if sys.platform == "win32":
          os.startfile(caminho_absoluto)
        elif sys.platform == "darwin":
          subprocess.run(["open", caminho_absoluto])
        else:
          subprocess.run(["xdg-open", caminho_absoluto])
      except Exception as ex:
        print(f"Não foi possível abrir automaticamente: {ex}")

      # 2. Pergunta se deseja enviar por WhatsApp Web após conferir
      resposta = messagebox.askyesno(
          "Relatório Gerado e Aberto",
          f"O relatório foi gerado e aberto para conferência!\nSalvo em:"
          f" {nome_arquivo_pdf}\n\nDeseja abrir o WhatsApp Web para enviar este"
          " relatório?",
      )
      if resposta:
        texto_msg = urllib.parse.quote(
            f"Olá! Segue em anexo/resumo o relatório do Fluxo de Caixa"
            f" ({'Geral' if not filtro_fluxo else filtro_fluxo}). O arquivo PDF"
            f" salvo no computador é: {nome_arquivo_pdf}"
        )
        url_whats = f"https://wa.me/?text={texto_msg}"
        webbrowser.open(url_whats)

    except Exception as e:
      messagebox.showerror(
          "Erro", f"Ocorreu um erro ao gerar o PDF:\n{str(e)}"
      )


if __name__ == "__main__":
  root = tk.Tk()
  app = AppFluxoCaixa(root)
  root.mainloop()