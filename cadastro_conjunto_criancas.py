import tkinter as tk
from tkinter import messagebox
import psycopg2
import pandas as pd

# Função para conectar ao banco de dados PostgreSQL
def connect_db():
    try:
        conn = psycopg2.connect(
            dbname="seu_db",
            user="seu_usuario",
            password="sua_senha",
            host="localhost"
        )
        return conn
    except Exception as e:
        messagebox.showerror("Erro", f"Não foi possível conectar ao banco de dados: {e}")
        return None

# Função para salvar dados em uma planilha Excel
def save_to_excel(data):
    df = pd.DataFrame(data, columns=["Nome Igreja", "Nome Pastor", "Nome Conjunto", "Qtd Masculino", "Qtd Feminina"])
    try:
        df.to_excel("conjunto_criancas_temp.xlsx", index=False)
        messagebox.showinfo("Sucesso", "Dados salvos em 'conjunto_criancas_temp.xlsx'")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao salvar dados em Excel: {e}")

# Função para adicionar uma nova criança
def add_child():
    nome_igreja = entry_nome_igreja.get()
    nome_pastor = entry_nome_pastor.get()
    nome_conjunto = entry_nome_conjunto.get()
    qtd_masculino = entry_qtd_masculino.get()
    qtd_feminina = entry_qtd_feminina.get()

    if not nome_igreja or not nome_pastor or not nome_conjunto or not qtd_masculino or not qtd_feminina:
        messagebox.showwarning("Campos vazios", "Todos os campos devem ser preenchidos")
        return

    # Salvar os dados em uma planilha Excel antes de adicionar ao banco de dados
    save_to_excel([(nome_igreja, nome_pastor, nome_conjunto, qtd_masculino, qtd_feminina)])

    conn = connect_db()
    if conn:
        try:
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO conjunto_criancas (nome_igreja, nome_pastor, nome_conjunto, qtd_masculino, qtd_feminina)
                VALUES (%s, %s, %s, %s, %s)
            """, (nome_igreja, nome_pastor, nome_conjunto, qtd_masculino, qtd_feminina))
            conn.commit()
            cur.close()
            conn.close()
            messagebox.showinfo("Sucesso", "Dados adicionados com sucesso")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao adicionar dados: {e}")
    else:
        messagebox.showerror("Erro", "Não foi possível conectar ao banco de dados")

# Função para visualizar e exportar os dados
def view_and_export_data():
    conn = connect_db()
    if conn:
        try:
            cur = conn.cursor()
            cur.execute("SELECT * FROM conjunto_criancas")
            rows = cur.fetchall()
            cur.close()
            conn.close()

            # Limpar a caixa de texto
            text_data.delete(1.0, tk.END)
            # Adicionar dados ao widget de texto
            for row in rows:
                text_data.insert(tk.END, f"{row}\n")
            
            # Exportar dados para uma planilha Excel
            df = pd.DataFrame(rows, columns=["ID", "Nome Igreja", "Nome Pastor", "Nome Conjunto", "Qtd Masculino", "Qtd Feminina"])
            df.to_excel("conjunto_criancas.xlsx", index=False)
            messagebox.showinfo("Sucesso", "Dados exportados para 'conjunto_criancas.xlsx' com sucesso")

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao visualizar e exportar dados: {e}")
    else:
        messagebox.showerror("Erro", "Não foi possível conectar ao banco de dados")

# Criar a interface gráfica
root = tk.Tk()
root.title("Cadastro de Conjunto de Crianças")

# Labels e entradas para os dados
tk.Label(root, text="Nome da Igreja:").grid(row=0, column=0, padx=10, pady=5)
entry_nome_igreja = tk.Entry(root)
entry_nome_igreja.grid(row=0, column=1, padx=10, pady=5)

tk.Label(root, text="Nome do Pastor:").grid(row=1, column=0, padx=10, pady=5)
entry_nome_pastor = tk.Entry(root)
entry_nome_pastor.grid(row=1, column=1, padx=10, pady=5)

tk.Label(root, text="Nome do Conjunto:").grid(row=2, column=0, padx=10, pady=5)
entry_nome_conjunto = tk.Entry(root)
entry_nome_conjunto.grid(row=2, column=1, padx=10, pady=5)

tk.Label(root, text="Quantidade de Crianças Masculino:").grid(row=3, column=0, padx=10, pady=5)
entry_qtd_masculino = tk.Entry(root)
entry_qtd_masculino.grid(row=3, column=1, padx=10, pady=5)

tk.Label(root, text="Quantidade de Crianças Feminina:").grid(row=4, column=0, padx=10, pady=5)
entry_qtd_feminina = tk.Entry(root)
entry_qtd_feminina.grid(row=4, column=1, padx=10, pady=5)

# Botões
tk.Button(root, text="Adicionar Criança", command=add_child).grid(row=5, column=0, padx=10, pady=10)
tk.Button(root, text="Visualizar e Exportar Dados", command=view_and_export_data).grid(row=5, column=1, padx=10, pady=10)

# Área de texto para exibir os dados
text_data = tk.Text(root, height=10, width=50)
text_data.grid(row=6, column=0, columnspan=2, padx=10, pady=10)

root.mainloop()
