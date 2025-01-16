import tkinter as tk
from tkinter import filedialog
import pandas as pd
import pyperclip
import difflib 

sugestoes_correcao = []

def carregar_arquivo():
    file_path = filedialog.askopenfilename(title="Selecionar Arquivo", filetypes=[("Excel Files", "*.xlsx;*.csv")])
    if file_path:
        entrada_arquivo.delete(0, tk.END)
        entrada_arquivo.insert(0, file_path)

def gerar_sql():
    global sugestoes_correcao
    sugestoes_correcao = []
    file_path = entrada_arquivo.get()
    query_template = entrada_query.get("1.0", tk.END).strip()

    campos = [entrada_param_1.get(), entrada_param_2.get(), entrada_param_3.get(), entrada_param_4.get(), entrada_param_5.get()]

    if not file_path or not query_template or not any(campos):
        aviso_label.config(text="Preencha todos os campos e selecione um arquivo.", fg="red")
        return

    try:
        data = pd.read_excel(file_path) if file_path.endswith('.xlsx') else pd.read_csv(file_path)

        colunas_existentes = data.columns.tolist()
        colunas_invalidas = []

        for i, campo in enumerate(campos):
            if campo and campo not in colunas_existentes:
                sugestao = difflib.get_close_matches(campo, colunas_existentes, n=1, cutoff=0.6)
                if sugestao:
                    colunas_invalidas.append((i, campo, sugestao[0]))
                    sugestoes_correcao.append((i, sugestao[0]))
                else:
                    colunas_invalidas.append((i, campo, None))

        if colunas_invalidas:
            sugestoes = []
            for _, campo, sugestao in colunas_invalidas:
                if sugestao:
                    sugestoes.append(f"Você quis dizer a coluna '{sugestao}'?")
                else:
                    sugestoes.append(f"A coluna '{campo}' não foi encontrada e não conseguimos sugerir uma alternativa.")
            aviso_label.config(text="Erro nas colunas: " + ", ".join(sugestoes), fg="red")
            botao_aplicar_correcoes.grid(row=9, column=0, columnspan=3, pady=5)
            return

        botao_aplicar_correcoes.grid_forget()

        sql_queries = []
        for _, row in data.iterrows():
            sql_query = query_template
            for i, campo in enumerate(campos):
                if campo and campo in data.columns:
                    valor = str(row[campo]) if pd.notna(row[campo]) else ''
                    sql_query = sql_query.replace(f"${{{i+1}}}", valor)

            sql_queries.append(sql_query)

        sql_output = "\n".join(sql_queries)
        with open("gerar_queries.sql", "w") as file:
            file.write(sql_output)

        aviso_label.config(text="Arquivo SQL gerado com sucesso em 'gerar_queries.sql'.", fg="green")
    
    except Exception as e:
        aviso_label.config(text=f"Erro ao processar o arquivo: {str(e)}", fg="red")

def aplicar_correcao():
    global sugestoes_correcao
    entradas = [entrada_param_1, entrada_param_2, entrada_param_3, entrada_param_4, entrada_param_5]

    for i, sugestao in sugestoes_correcao:
        entradas[i].delete(0, tk.END)
        entradas[i].insert(0, sugestao)

    aviso_label.config(text="Correções aplicadas. Clique em 'Gerar SQL' novamente.", fg="green")
    botao_aplicar_correcoes.grid_forget()

def copiar_query():
    query_exemplo = "UPDATE usuarios SET email = '${1}', nome = '${2}' WHERE id = '${3}';"
    if query_exemplo:
        pyperclip.copy(query_exemplo)
        aviso_label.config(text="Query de exemplo copiada para a área de transferência.", fg="green")
    else:
        aviso_label.config(text="Nenhuma query de exemplo para copiar.", fg="red")

root = tk.Tk()
root.title("Gerador de SQL com Substituição de Parâmetros (DEPARA)")

tk.Label(root, text="1. Selecione o arquivo (xlsx/csv) com os dados:", anchor="w").grid(row=0, column=0, sticky=tk.W, padx=10, pady=5)
entrada_arquivo = tk.Entry(root, width=50)
entrada_arquivo.grid(row=0, column=1, padx=10)

botao_carregar = tk.Button(root, text="Carregar Arquivo", command=carregar_arquivo)
botao_carregar.grid(row=0, column=2)

tk.Label(root, text="2. Insira a Query SQL com placeholders (ex: ${1}, ${2}, ...):", anchor="w").grid(row=1, column=0, sticky=tk.W, padx=10, pady=5)
entrada_query = tk.Text(root, height=5, width=50)
entrada_query.grid(row=1, column=1, columnspan=2, padx=10, pady=5)

tk.Label(root, text="3. Informe os nomes das colunas para substituir os placeholders:", anchor="w").grid(row=2, column=0, sticky=tk.W, padx=10, pady=5)

tk.Label(root, text="Parametro 1 (ex: email):").grid(row=3, column=0, sticky=tk.W, padx=10, pady=5)
entrada_param_1 = tk.Entry(root, width=30)
entrada_param_1.grid(row=3, column=1, padx=10)

tk.Label(root, text="Parametro 2 (ex: nome):").grid(row=4, column=0, sticky=tk.W, padx=10, pady=5)
entrada_param_2 = tk.Entry(root, width=30)
entrada_param_2.grid(row=4, column=1, padx=10)

tk.Label(root, text="Parametro 3 (ex: telefone):").grid(row=5, column=0, sticky=tk.W, padx=10, pady=5)
entrada_param_3 = tk.Entry(root, width=30)
entrada_param_3.grid(row=5, column=1, padx=10)

tk.Label(root, text="Parametro 4 (ex: id):").grid(row=6, column=0, sticky=tk.W, padx=10, pady=5)
entrada_param_4 = tk.Entry(root, width=30)
entrada_param_4.grid(row=6, column=1, padx=10)

tk.Label(root, text="Parametro 5 (opcional):").grid(row=7, column=0, sticky=tk.W, padx=10, pady=5)
entrada_param_5 = tk.Entry(root, width=30)
entrada_param_5.grid(row=7, column=1, padx=10)

botao_gerar = tk.Button(root, text="Gerar SQL", command=gerar_sql)
botao_gerar.grid(row=8, column=0, columnspan=3, pady=15)

botao_copiar = tk.Button(root, text="Copiar Query", command=copiar_query)
botao_copiar.grid(row=9, column=0, columnspan=3, pady=5)

botao_aplicar_correcoes = tk.Button(root, text="Aplicar Correções", command=aplicar_correcao)
botao_aplicar_correcoes.grid_forget()

aviso_label = tk.Label(root, text="", fg="red")
aviso_label.grid(row=10, column=0, columnspan=3)

depara_label = tk.Label(root, text="Utilize a aplicação para fazer o 'DEPARA' entre os dados da planilha e os parâmetros da query SQL.", anchor="w", font=("Arial", 10, "italic"))
depara_label.grid(row=11, column=0, columnspan=3, padx=10, pady=5)

instrucoes_label = tk.Label(root, text="Exemplo de Query SQL: UPDATE usuarios SET email = '${1}', nome = '${2}' WHERE id = '${3}';", anchor="w", font=("Arial", 10, "italic"))
instrucoes_label.grid(row=12, column=0, columnspan=3, padx=10, pady=5)

root.mainloop()
