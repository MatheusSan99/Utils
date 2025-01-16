import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import threading

def dividir_arquivos_excel():
    def selecionar_arquivo():
        excel_path = filedialog.askopenfilename(title="Selecione o arquivo Excel", filetypes=[("Excel Files", "*.xlsx")])
        if excel_path:
            entrada_arquivo.delete(0, tk.END)
            entrada_arquivo.insert(0, excel_path)

    def selecionar_diretorio_saida():
        output_folder = filedialog.askdirectory(title="Selecione o diretório de saída")
        if output_folder:
            entrada_saida.delete(0, tk.END)
            entrada_saida.insert(0, output_folder)

    def validar_inputs():
        excel_path = entrada_arquivo.get().strip()
        output_prefix = entrada_prefixo.get().strip()
        output_folder = entrada_saida.get().strip()

        if not entrada_linhas.get().strip():
            messagebox.showerror("Erro", "O número de linhas por arquivo não pode ser vazio.")
            return False
        
        try:
            lines_per_file = int(entrada_linhas.get().strip())
            if lines_per_file <= 0:
                raise ValueError("O número de linhas deve ser maior que 0.")
        except ValueError as e:
            messagebox.showerror("Erro", f"Entrada inválida para o número de linhas: {e}")
            return False

        if not os.path.isfile(excel_path):
            messagebox.showerror("Erro", "O arquivo Excel não foi encontrado.")
            return False

        if not excel_path.lower().endswith('.xlsx'):
            messagebox.showerror("Erro", "O arquivo informado não é um arquivo Excel válido (.xlsx).")
            return False

        if not os.path.isdir(output_folder):
            messagebox.showerror("Erro", "O diretório de saída não foi encontrado.")
            return False
        
        if not output_prefix:
            messagebox.showerror("Erro", "O prefixo para os arquivos gerados não pode ser vazio.")
            return False

        return True

    def dividir():
        if not validar_inputs():
            return

        excel_path = entrada_arquivo.get().strip()
        output_prefix = entrada_prefixo.get().strip()
        output_folder = entrada_saida.get().strip()
        botao_dividir.config(state=tk.DISABLED)
        progress_label.config(text="Processando... Aguarde.")

        try:
            lines_per_file = int(entrada_linhas.get().strip())
            if lines_per_file <= 0:
                raise ValueError("O número de linhas deve ser maior que 0.")
        except ValueError as e:
            messagebox.showerror("Erro", f"Entrada inválida para o número de linhas: {e}")
            botao_dividir.config(state=tk.NORMAL)
            progress_label.config(text="")
            return

        def processar():
            try:
                df_dict = pd.read_excel(excel_path, sheet_name=None)
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao ler o arquivo Excel: {e}")
                botao_dividir.config(state=tk.NORMAL)
                progress_label.config(text="")
                return

            def split_and_save(df, sheet_name, output_folder):
                num_files = (len(df) // lines_per_file) + (1 if len(df) % lines_per_file != 0 else 0)
                for i in range(num_files):
                    start_row = i * lines_per_file
                    end_row = start_row + lines_per_file
                    chunk = df[start_row:end_row]
                    output_file = os.path.join(output_folder, f"{output_prefix}_{sheet_name}_{i+1}.xlsx")
                    chunk.to_excel(output_file, index=False)

            for sheet_name, df in df_dict.items():
                split_and_save(df, sheet_name, output_folder)

            messagebox.showinfo("Sucesso", "Divisão completa! Arquivos gerados.")
            progress_label.config(text="")
            botao_dividir.config(state=tk.NORMAL)

        thread = threading.Thread(target=processar)
        thread.start()

    root = tk.Tk()
    root.title("Divisor de Arquivos Excel")

    tk.Label(root, text="Selecione o arquivo Excel:", anchor="w").grid(row=0, column=0, sticky=tk.W, padx=10, pady=5)
    entrada_arquivo = tk.Entry(root, width=50)
    entrada_arquivo.grid(row=0, column=1, padx=10)
    botao_selecionar = tk.Button(root, text="Selecionar Arquivo", command=selecionar_arquivo)
    botao_selecionar.grid(row=0, column=2)

    tk.Label(root, text="Digite o prefixo para os arquivos gerados:", anchor="w").grid(row=1, column=0, sticky=tk.W, padx=10, pady=5)
    entrada_prefixo = tk.Entry(root, width=50)
    entrada_prefixo.grid(row=1, column=1, padx=10)

    tk.Label(root, text="Número de linhas por arquivo:", anchor="w").grid(row=2, column=0, sticky=tk.W, padx=10, pady=5)
    entrada_linhas = tk.Entry(root, width=50)
    entrada_linhas.grid(row=2, column=1, padx=10)

    tk.Label(root, text="Selecione o diretório de saída:", anchor="w").grid(row=3, column=0, sticky=tk.W, padx=10, pady=5)
    entrada_saida = tk.Entry(root, width=50)
    entrada_saida.grid(row=3, column=1, padx=10)
    botao_selecionar_saida = tk.Button(root, text="Selecionar Diretório", command=selecionar_diretorio_saida)
    botao_selecionar_saida.grid(row=3, column=2)

    botao_dividir = tk.Button(root, text="Dividir Arquivo", command=dividir)
    botao_dividir.grid(row=4, column=0, columnspan=3, pady=15)

    progress_label = tk.Label(root, text="", anchor="w", fg="red")
    progress_label.grid(row=5, column=0, columnspan=3)

    root.mainloop()

dividir_arquivos_excel()
