import pandas as pd
import tkinter as tk
from tkinter import ttk

# Função para adicionar um novo registro ao DataFrame e salvar no Excel
def adicionar_registro():
    nome = nome_entry.get()
    presenca = int(presenca_entry.get())
    falta = int(falta_entry.get())
    nota11 = float(nota11_entry.get())
    nota12 = float(nota12_entry.get())
    nota21 = float(nota21_entry.get())
    nota22 = float(nota22_entry.get())

    media = (nota11 + nota12 + nota21 + nota22) / 4
    situacao = 'Aprovado' if media >= 6 else 'Reprovado'
    
    novo_registro = pd.DataFrame({'Nome': [nome], 'Presença': [presenca], 'Falta': [falta], 'Nota1.1': [nota11], 'Nota2.1': [nota12], 'Nota1.2': [nota21], 'Nota2.2': [nota22], 'Média': [media], 'Situação': [situacao]})
    
    global dados
    dados = pd.concat([dados, novo_registro], ignore_index=True)
    
    # Salvar os dados em um arquivo Excel
    dados.to_excel('dados_alunos.xlsx', index=False)
    
    # Atualiza a exibição na tabela
    tree.delete(*tree.get_children())
    for index, row in dados.iterrows():
        tree.insert("", "end", values=(row['Nome'], row['Presença'], row['Falta'], row['Nota1.1'], row['Nota2.1'], row['Nota1.2'], row['Nota2.2'], row['Média'], row['Situação']))
    
    # Limpa os campos de entrada
    nome_entry.delete(0, 'end')
    presenca_entry.delete(0, 'end')
    falta_entry.delete(0, 'end')
    nota11_entry.delete(0, 'end')
    nota12_entry.delete(0, 'end')
    nota21_entry.delete(0, 'end')
    nota22_entry.delete(0, 'end')

# Crie um DataFrame vazio para armazenar os dados
dados = pd.DataFrame(columns=['Nome', 'Presença', 'Falta', 'Nota1.1', 'Nota2.1', 'Nota1.2', 'Nota2.2', 'Média', 'Situação'])

# Crie a janela principal
root = tk.Tk()
root.title("Cadastro de Alunos")

# Crie um frame para a entrada de dados
data_frame = ttk.Frame(root)
data_frame.pack(padx=20, pady=20)

# Labels e campos de entrada
ttk.Label(data_frame, text="Nome:").grid(row=0, column=0, sticky="w")
nome_entry = ttk.Entry(data_frame)
nome_entry.grid(row=0, column=1, padx=5, pady=5)

ttk.Label(data_frame, text="Presença:").grid(row=1, column=0, sticky="w")
presenca_entry = ttk.Entry(data_frame)
presenca_entry.grid(row=1, column=1, padx=5, pady=5)

ttk.Label(data_frame, text="Falta:").grid(row=2, column=0, sticky="w")
falta_entry = ttk.Entry(data_frame)
falta_entry.grid(row=2, column=1, padx=5, pady=5)

ttk.Label(data_frame, text="Nota 1.1:").grid(row=3, column=0, sticky="w")
nota11_entry = ttk.Entry(data_frame)
nota11_entry.grid(row=3, column=1, padx=5, pady=5)

ttk.Label(data_frame, text="Nota 2.1:").grid(row=4, column=0, sticky="w")
nota12_entry = ttk.Entry(data_frame)
nota12_entry.grid(row=4, column=1, padx=5, pady=5)

ttk.Label(data_frame, text="Nota 1.2:").grid(row=5, column=0, sticky="w")
nota21_entry = ttk.Entry(data_frame)
nota21_entry.grid(row=5, column=1, padx=5, pady=5)

ttk.Label(data_frame, text="Nota 2.2:").grid(row=6, column=0, sticky="w")
nota22_entry = ttk.Entry(data_frame)
nota22_entry.grid(row=6, column=1, padx=5, pady=5)

# Botão para adicionar registro
adicionar_button = ttk.Button(data_frame, text="Adicionar Registro", command=adicionar_registro)
adicionar_button.grid(row=7, columnspan=2, padx=5, pady=10)

# Crie uma tabela para exibir os dados
columns = ['Nome', 'Presença', 'Falta', 'Nota1.1', 'Nota2.1', 'Nota1.2', 'Nota2.2', 'Média', 'Situação']
tree = ttk.Treeview(root, columns=columns, show='headings')
for col in columns:
    tree.heading(col, text=col)
tree.pack(padx=20, pady=20)

# Execute a interface gráfica
root.mainloop()
