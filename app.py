import tkinter as tk
from tkinter import ttk
import openpyxl
import os
from tkinter import messagebox

# CARREGANDO A TABELA
def load_workbook(wb_path):
    if os.path.exists(wb_path):
        return openpyxl.load_workbook(wb_path)
    else:
        return messagebox.showerror ("Erro", "arquivo não encontrado")
    
wb_path = "produtos.xlsx"
wb = load_workbook(wb_path)
sheet = wb["Sheet"]
sheet_obj = wb.active
list_values = list(sheet_obj.values)

# CRIANDO AS FUNÇÕES DOS BOTÕES
def insert_row():
    # Pegando as informações dos widgets
    code = int(code_spinbox.get())
    product = product_entry.get()
    stock_status = status_combobox.get()
    avaliability = "Saiu de linha" if a.get() else "Disponível"

    # Inserindo as informações na tabela .xlsx
    wb_path = "produtos.xlsx"
    wb = openpyxl.load_workbook(wb_path)
    sheet = wb["Sheet"]
    sheet_obj = wb.active
    row_values = [code, product, stock_status, avaliability]
    sheet_obj.append(row_values)
    wb.save(wb_path)

    # Inserindo na visualização da aplicação
    treeview.insert('', tk.END, values=row_values)

    # Resetando as configurações dos widgets
    code_spinbox.delete(0, "end")
    code_spinbox.insert(0, "Código do produto")
    product_entry.delete(0, "end")
    product_entry.insert(0, "Nome do Produto")
    status_combobox.delete(0, "end")
    status_combobox.insert(0, "Status de estoque")
    a.set(False)

# CONFIGURANDO A INTERFACE GRÁFICA - GUI
root = tk.Tk()
root.title("Sistema de Cadastro de Produtos")

frame = tk.Frame(root)
frame.pack()

# Moldura dos campos de preenchimento
widgets_frame = ttk.LabelFrame(frame, text="Dados do Produto")
widgets_frame.grid(row=0, column=0, padx=20, pady=10)

# Campo do código
code_spinbox = ttk.Spinbox(widgets_frame, from_="0000", to="9999")
code_spinbox.insert(0, "Código do produto")
code_spinbox.bind("<FocusIn>", lambda e: product_entry.delete(0, 'end') )
code_spinbox.grid(row=0 , column=0, padx=5, pady=5, sticky="ew")

# Campo do produto
product_entry = ttk.Entry(widgets_frame)
product_entry.insert(0, "Nome do Produto")
product_entry.bind("<FocusIn>", lambda e: product_entry.delete(0, 'end') )
product_entry.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

# Status do estoque
combo_list = ("Possui estoque", "Em falta")
status_combobox = ttk.Combobox(widgets_frame, values=combo_list)
status_combobox.insert(0, "Status de estoque")
status_combobox.grid(row=2, column=0, padx=5, pady=5, sticky="ew")

# Disponibilidade do produto
a = tk.BooleanVar()
checkbutton = ttk.Checkbutton(widgets_frame, text="Saiu de linha", variable=a)
checkbutton.grid(row=3, column=0, padx=5, pady=5, sticky="nsew")

# Botão para inserir
insert_button = ttk.Button(widgets_frame, text="Salvar produto", command= insert_row)
insert_button.grid(row=4, column=0, padx=5, pady=5, sticky="nsew")

# Separar as funções de atualizar e deletar
separator = ttk.Separator(widgets_frame)
separator.grid(row=5, column=0, padx=10, pady=10, sticky="ew")

# Botão para editar produto
edit_button = ttk.Button(widgets_frame, text="Editar produto já cadastrado")
edit_button.grid(row=6, column=0, padx=5, pady=5, sticky="nsew")

# Visualização da tabela Excel
treeFrame = ttk.Frame(frame)
treeFrame.grid(row=0, column=1, pady=10)
treeScroll = ttk.Scrollbar(treeFrame)
treeScroll.pack(side="right", fill="y")

view_xl = ("CÓDIGO", "PRODUTO", "STATUS", "DISPONIBILIDADE")
treeview = ttk.Treeview(treeFrame, show="headings",
                        yscrollcommand=treeScroll.set, columns=view_xl, height=13)
treeview.column("CÓDIGO", width=80)
treeview.column("PRODUTO", width=150)
treeview.column("STATUS", width=150)
treeview.column("DISPONIBILIDADE", width=150)
treeview.pack()
treeScroll.config(command=treeview.yview)

# Inserindo os valores na janela da aplicação
for col_name in list_values[0]:
    treeview.heading(col_name, text=col_name)
for value_tuple in list_values[1:]:
    treeview.insert('', tk.END, values=value_tuple)



root.mainloop()