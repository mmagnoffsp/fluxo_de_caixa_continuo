import csv
from datetime import datetime
import os
import tkinter as tk
from tkinter import messagebox, ttk

# Nome do arquivo onde os dados serão salvos (agora com o nome solicitado)
ARQUIVO_DADOS = "fluxo_de_caixa_continuo.csv"


class AppFluxoCaixa:

  def __init__(self, root):
    self.root = root
    self.root.title("Fluxo de Caixa Contínuo")
    self.root.geometry("950x650")
    self.root.config(bg="#f0f0f0")

    # Lista exata de botões/operações solicitadas
    self.categorias = [
        "PAGAMENTO RECEBIDO",
        "ADIANTAMENTO VALE",
        "SALARIO ANTES DE SAIR DE FERIAS",
        "FERIAS",
        "SALARIO 10 DIAS TRABALHADOS DAS FERIAS",
        "VALE APOS FERIAS",
        "PAGAMENTO APOS FERIAS",
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
        font=("Arial", 16, "bold"),
        bg="#f0f0f0",
        fg="#333",
    )
    titulo.pack(pady=10)

    # Frame de Entrada de Dados
    frame_form = tk.LabelFrame(
        self.root, text=" Novo Lançamento ", font=("Arial", 11, "bold"), bg="#f0f0f0"
    )
    frame_form.pack(fill="x", padx=15, pady=5)

    # Data
    tk.Label(
        frame_form, text="Data (DD/MM/AAAA):", bg="#f0f0f0", font=("Arial", 10)
    ).grid(row=0, column=0, sticky="w", padx=5, pady=5)
    self.entry_data = tk.Entry(frame_form, font=("Arial", 10), width=15)
    self.entry_data.grid(row=0, column=1, sticky="w", padx=5, pady=5)
    self.entry_data.insert(0, datetime.now().strftime("%d/%m/%Y"))

    # Valor
    tk.Label(
        frame_form, text="Valor (R$):", bg="#f0f0f0", font=("Arial", 10)
    ).grid(row=0, column=2, sticky="w", padx=5, pady=5)
    self.entry_valor = tk.Entry(frame_form, font=("Arial", 10), width=15)
    self.entry_valor.grid(row=0, column=3, sticky="w", padx=5, pady=5)

    # Descrição / Observação Opcional
    tk.Label(
        frame_form, text="Observação:", bg="#f0f0f0", font=("Arial", 10)
    ).grid(row=1, column=0, sticky="w", padx=5, pady=5)
    self.entry_obs = tk.Entry(frame_form, font=("Arial", 10), width=45)
    self.entry_obs.grid(
        row=1, column=1, columnspan=3, sticky="w", padx=5, pady=5
    )

    # Frame dos Botões (Sua lista de botões personalizados)
    frame_botoes = tk.LabelFrame(
        self.root,
        text=" Selecione a Operação (Clique no botão desejado) ",
        font=("Arial", 11, "bold"),
        bg="#f0f0f0",
    )
    frame_botoes.pack(fill="x", padx=15, pady=10)

    # Criando um botão para cada item da sua lista
    for i, categoria in enumerate(self.categorias):
      linha = i // 3
      coluna = i % 3
      btn = tk.Button(
          frame_botoes,
          text=categoria,
          bg="#2196F3",
          fg="white",
          font=("Arial", 9, "bold"),
          width=35,
          command=lambda cat=categoria: self.adicionar_lancamento(cat),
      )
      btn.grid(row=linha, column=coluna, padx=5, pady=5, sticky="ew")

    # Tabela de Histórico (Treeview)
    frame_tabela = tk.Frame(self.root)
    frame_tabela.pack(fill="both", expand=True, padx=15, pady=10)

    colunas = ("ID", "Data", "Tipo / Operação", "Valor (R$)", "Observação")
    self.tabela = ttk.Treeview(
        frame_tabela, columns=colunas, show="headings", selectmode="browse"
    )

    for col in colunas:
      self.tabela.heading(col, text=col)
      if col == "ID":
        self.tabela.column(col, width=40, anchor="center")
      elif col == "Data":
        self.tabela.column(col, width=90, anchor="center")
      elif col == "Valor (R$)":
        self.tabela.column(col, width=100, anchor="e")
      else:
        self.tabela.column(col, width=250, anchor="w")

    scrollbar = ttk.Scrollbar(
        frame_tabela, orient="vertical", command=self.tabela.yview
    )
    self.tabela.configure(yscrollcommand=scrollbar.set)

    self.tabela.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # Botão para excluir registro selecionado
    btn_excluir = tk.Button(
        self.root,
        text="Excluir Lançamento Selecionado",
        bg="#f44336",
        fg="white",
        font=("Arial", 10, "bold"),
        command=self.excluir_lancamento,
    )
    btn_excluir.pack(anchor="e", padx=15, pady=5)

  def adicionar_lancamento(self, tipo_operacao):
    data = self.entry_data.get().strip()
    valor_str = self.entry_valor.get().strip().replace(",", ".")
    obs = self.entry_obs.get().strip()

    # Validações básicas
    if not data or not valor_str:
      messagebox.showerror(
          "Erro", "Preencha a Data e o Valor antes de clicar no botão!"
      )
      return

    try:
      valor = float(valor_str)
    except ValueError:
      messagebox.showerror(
          "Erro", "O valor digitado é inválido. Use apenas números."
      )
      return

    # Gerar ID único baseado no timestamp atual
    id_lancamento = datetime.now().strftime("%Y%m%d%H%M%S")

    # Inserir na tabela visual
    self.tabela.insert(
        "",
        "end",
        values=(
            id_lancamento,
            data,
            tipo_operacao,
            f"R$ {valor:,.2f}".replace(",", "X")
            .replace(".", ",")
            .replace("X", "."),
            obs,
        ),
    )

    # Salvar no arquivo CSV
    self.salvar_no_csv(id_lancamento, data, tipo_operacao, valor, obs)

    # Limpar valor e obs para o próximo lançamento
    self.entry_valor.delete(0, tk.END)
    self.entry_obs.delete(0, tk.END)
    messagebox.showinfo(
        "Sucesso", f"Lançamento de '{tipo_operacao}' registrado com sucesso!"
    )

  def salvar_no_csv(self, id_l, data, tipo, valor, obs):
    modo = "a" if os.path.exists(ARQUIVO_DADOS) else "w"
    with open(
        ARQUIVO_DADOS, mode=modo, newline="", encoding="utf-8"
    ) as arquivo:
      escritor = csv.writer(arquivo)
      if modo == "w":
        escritor.writerow(["ID", "Data", "Tipo", "Valor", "Observacao"])
      escritor.writerow([id_l, data, tipo, valor, obs])

  def carregar_dados(self):
    if os.path.exists(ARQUIVO_DADOS):
      with open(
          ARQUIVO_DADOS, mode="r", newline="", encoding="utf-8"
      ) as arquivo:
        leitor = csv.reader(arquivo)
        next(leitor, None)  # Pular cabeçalho
        for linha in leitor:
          if len(linha) == 5:
            id_l, data, tipo, valor, obs = linha
            try:
              val_float = float(valor)
              valor_formatado = (
                  f"R$ {val_float:,.2f}".replace(",", "X")
                  .replace(".", ",")
                  .replace("X", ".")
              )
            except:
              valor_formatado = valor

            self.tabela.insert(
                "",
                "end",
                values=(id_l, data, tipo, valor_formatado, obs),
            )

  def excluir_lancamento(self):
    selecionado = self.tabela.selection()
    if not selecionado:
      messagebox.showwarning(
          "Aviso", "Selecione um lançamento na tabela para excluir."
      )
      return

    item = self.tabela.item(selecionado)
    id_alvo = item["values"][0]

    # Remover da tabela visual
    self.tabela.delete(selecionado)

    # Reconstruir o CSV sem o item excluído
    linhas_mantidas = []
    if os.path.exists(ARQUIVO_DADOS):
      with open(
          ARQUIVO_DADOS, mode="r", newline="", encoding="utf-8"
      ) as arquivo:
        leitor = csv.reader(arquivo)
        cabecalho = next(leitor, None)
        for linha in leitor:
          if linha and str(linha[0]) != str(id_alvo):
            linhas_mantidas.append(linha)

      with open(
          ARQUIVO_DADOS, mode="w", newline="", encoding="utf-8"
      ) as arquivo:
        escritor = csv.writer(arquivo)
        if cabecalho:
          escritor.writerow(cabecalho)
        escritor.writerows(linhas_mantidas)

    messagebox.showinfo("Sucesso", "Lançamento excluído com sucesso!")


if __name__ == "__main__":
  root = tk.Tk()
  app = AppFluxoCaixa(root)
  root.mainloop()