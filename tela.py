import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import processa
import os

def busca_dados(inicio_dado, fim_dado, salvar, local):
    folder = salvar.get()
    data_inicio = inicio_dado.get()
    data_fim = fim_dado.get()
    local_pesquisa = local.get()
    inicio = datetime.strptime(data_inicio, "%d/%m/%Y")
    fim = datetime.strptime(data_fim, "%d/%m/%Y")

    formata_inicio = inicio.strftime("%Y-%m-%d")
    formata_fim = fim.strftime("%Y-%m-%d")

    dados = processa.captura_dados(local_pesquisa,formata_inicio, formata_fim)
    if not data_inicio or not data_fim:
        messagebox.showwarning("Atenção","Por favor, escolha um período para continuar")

    if dados:
        caminho = os.path.join(folder,"BaseDados.xlsx")
        try:
            processa.save_excel(dados,caminho)
            messagebox.showinfo("Salvo",f"Salvo o arquivo na pasta {caminho}")
        except PermissionError:
            messagebox.showwarning("Atenção",
                                 f"O arquivo {os.path.basename(caminho)} está aberto ou não pode ser acessado."
                                 f"Por favor, verifique se o arquivo está aberto tente novamente.")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao salvar o arquivo: {e}")
    else:
        messagebox.showwarning("Atenção","Por favor, verifique o período selecionado e tente novamente")
        print("Erro ao capturar as datas")


def browser(entry):
    seleciona = filedialog.askdirectory()
    if seleciona:
        entry.delete(0, tk.END)
        entry.insert(tk.END, seleciona)
    else:
        messagebox.showwarning("Atenção","Por favor, selecione um diretório antes de continuar.")

def main():
    janela = tk.Tk()
    janela.title("Previsão do Tempo")
    janela.iconbitmap('tails_ico.ico')

    local_label = tk.Label(janela, text="Cidade")
    local_label.grid(row=0,column=0, padx=15, pady=5, sticky="w")

    inicio_label = tk.Label(janela, text='Data de Início')
    inicio_label.grid(row=1, column=0, padx=15, pady=5, sticky="w")

    fim_label = tk.Label(janela, text='Data Fim')
    fim_label.grid(row=2, column=0, padx=15, pady=5, sticky="w")

    salvar_label = tk.Label(janela, text="Diretório")
    salvar_label.grid(row=3, column=0, padx=15, pady=5, sticky="w")

    def handle_keypress_nohour(event, entry_widget):
        # responsavel por alterar o campo para dia/mes/ano
        if event.char.isdigit() or event.keysym == 'BackSpace' or event.keysym == 'Tab':  # or event.char in ["/", ":", " "]:
            resultado = entry_widget.get()
            if len(resultado) == 2 and event.char.isdigit():
                entry_widget.insert(tk.END, '/')
            elif len(resultado) == 5 and event.char.isdigit():
                entry_widget.insert(tk.END, '/')
            elif len(resultado) >= 9 and event.char.isdigit():
                entry_widget.delete(9, tk.END)
            return True
        return "break"

    def handle_nokeypress(event):
        if event.keysym in ['BackSpace','Tab']:
            return True
        return "break"

    def handle_keypress_nochar(event, entry_widget):
        if event.char.isalpha() or event.keysym in ['BackSpace','Tab']:
            return True
        return "break"

    local = tk.Entry(janela, width=12)
    local.grid(row=0, column=1, padx=5, pady=5, sticky="w")
    local.bind("<Key>",lambda event:handle_keypress_nochar(event,local))

    inicio_dado = tk.Entry(janela, width=12)
    inicio_dado.grid(row=1, column=1, padx=5, pady=5, sticky="w")
    inicio_dado.bind("<Key>", lambda event: handle_keypress_nohour(event, inicio_dado))

    fim_dado = tk.Entry(janela, width=12)
    fim_dado.grid(row=2, column=1, padx=5, pady=5, sticky="w")
    fim_dado.bind("<Key>", lambda event: handle_keypress_nohour(event, fim_dado))

    salvar = tk.Entry(janela, width = 20)
    salvar.grid(row=3, column=1, padx=5, pady=5, sticky="w")
    salvar.bind("<Key>", lambda event: handle_nokeypress(event))

    procura_diret = tk.Button(janela, text="Procurar", command= lambda: browser(salvar))
    procura_diret.grid(row=3, column=2, padx=5, pady=5)

    botao_start = tk.Button(janela, text="Começar", command=lambda:busca_dados(inicio_dado,fim_dado,salvar,local))
    botao_start.grid(row=5, columnspan=3, padx=5, pady=5)

    janela.protocol("WM_DELETE_WINDOW", janela.quit)
    janela.mainloop()

if __name__ == "__main__":
    main()