import customtkinter as ctk
import tkinter as tk
import tkinter.messagebox as messagebox
from tkcalendar import DateEntry
from excel_handler import add_multiple_records, search_by_date, delete_record

# Settings
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("FALTAS- INFORMÁTICA PARA INTERNET")
        self.geometry("600x750")

        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        self.alunos_temp = [] # Lista de dicionários
        self.resultados_busca = [] # Para armazenar a última busca e ordenar

        # ====== CADASTRO ======
        self.label_titulo = ctk.CTkLabel(self, text="CADASTRO DE FALTAS", font=ctk.CTkFont(size=20, weight="bold"))
        self.label_titulo.grid(row=0, column=0, columnspan=2, padx=20, pady=(15, 10))

        # Config frame de inputs para melhor alinhamento
        self.frame_inputs = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_inputs.grid(row=1, column=0, columnspan=2, padx=20, pady=5, sticky="ew")
        self.frame_inputs.grid_columnconfigure(1, weight=1)

        # Data
        ctk.CTkLabel(self.frame_inputs, text="DATA:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.entry_data = DateEntry(self.frame_inputs, width=12, background='darkblue',
                                    foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
        self.entry_data.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        # Turma
        ctk.CTkLabel(self.frame_inputs, text="TURMA:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.entry_turma = ctk.CTkComboBox(self.frame_inputs, values=["200 INF", "300 INF", "301 INF"])
        self.entry_turma.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        # Nome
        ctk.CTkLabel(self.frame_inputs, text="NOME:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.entry_nome = ctk.CTkEntry(self.frame_inputs, placeholder_text="Nome do Aluno")
        self.entry_nome.grid(row=2, column=1, padx=5, pady=5, sticky="ew")

        # Botão +
        self.btn_adicionar = ctk.CTkButton(self.frame_inputs, text="ADICIONAR", width=150, command=self.adicionar_lista)
        self.btn_adicionar.grid(row=3, column=0, columnspan=2, padx=5, pady=10)

        # Scrollable Frame para mostrar os nomes adicionados
        self.label_lista = ctk.CTkLabel(self, text="Alunos a Registrar:", font=ctk.CTkFont(weight="bold"))
        self.label_lista.grid(row=2, column=0, columnspan=2, padx=20, pady=(0, 0), sticky="w")
        
        self.frame_lista = ctk.CTkScrollableFrame(self, height=80)
        self.frame_lista.grid(row=3, column=0, columnspan=2, padx=20, pady=5, sticky="ew")

        self.labels_alunos = [] # Guarda referencias para destruir depois

        # Botão Salvar Todas (Excel)
        self.btn_salvar = ctk.CTkButton(self, text="REGISTRAR FALTAS", fg_color="green", hover_color="darkgreen", command=self.salvar_excel)
        self.btn_salvar.grid(row=4, column=0, columnspan=2, padx=20, pady=10)

        # ====== DIVISOR ======
        self.frame_divisor = ctk.CTkFrame(self, height=2, fg_color="gray")
        self.frame_divisor.grid(row=5, column=0, columnspan=2, padx=20, pady=10, sticky="ew")

        # ====== PESQUISA ======
        self.label_pesquisa = ctk.CTkLabel(self, text="PESQUISAR FALTAS", font=ctk.CTkFont(size=20, weight="bold"))
        self.label_pesquisa.grid(row=6, column=0, columnspan=2, padx=20, pady=(0, 10))

        self.frame_pesquisa = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_pesquisa.grid(row=7, column=0, columnspan=2, padx=20, pady=5, sticky="ew")
        
        ctk.CTkLabel(self.frame_pesquisa, text="Data:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.entry_pesquisa_data = DateEntry(self.frame_pesquisa, width=12, background='darkblue',
                                             foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
        self.entry_pesquisa_data.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        self.btn_pesquisar = ctk.CTkButton(self.frame_pesquisa, text="Pesquisar", command=self.pesquisar)
        self.btn_pesquisar.grid(row=0, column=2, padx=15, pady=5)

        # Resultados em colunas scrollaveis
        self.frame_resultados = ctk.CTkFrame(self)
        self.frame_resultados.grid(row=8, column=0, columnspan=2, padx=20, pady=(10, 20), sticky="nsew")
        self.frame_resultados.grid_columnconfigure((0,1,2), weight=1)
        self.grid_rowconfigure(8, weight=1) # Permite o frame esticar
        self.frame_resultados.grid_rowconfigure(1, weight=1) # Permite que scrollables estiquem no Y

        ctk.CTkLabel(self.frame_resultados, text="200 INF", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=5, pady=5)
        ctk.CTkLabel(self.frame_resultados, text="300 INF", font=ctk.CTkFont(weight="bold")).grid(row=0, column=1, padx=5, pady=5)
        ctk.CTkLabel(self.frame_resultados, text="301 INF", font=ctk.CTkFont(weight="bold")).grid(row=0, column=2, padx=5, pady=5)

        self.caixa_200 = ctk.CTkScrollableFrame(self.frame_resultados)
        self.caixa_200.grid(row=1, column=0, padx=5, pady=2, sticky="nsew")
        
        self.caixa_300 = ctk.CTkScrollableFrame(self.frame_resultados)
        self.caixa_300.grid(row=1, column=1, padx=5, pady=2, sticky="nsew")
        
        self.caixa_301 = ctk.CTkScrollableFrame(self.frame_resultados)
        self.caixa_301.grid(row=1, column=2, padx=5, pady=2, sticky="nsew")

    def adicionar_lista(self):
        data = self.entry_data.get().strip()
        turma = self.entry_turma.get().strip()
        nome = self.entry_nome.get().strip()

        if not data or not turma or not nome:
            messagebox.showwarning("Aviso", "Preencha a data, turma e nome para adicionar à lista!")
            return

        aluno = {"data": data, "turma": turma, "nome": nome}
        self.alunos_temp.append(aluno)
        
        lbl_texto = f"Data: {data} | Turma: {turma} | Nome: {nome}"
        lbl = ctk.CTkLabel(self.frame_lista, text=lbl_texto)
        lbl.pack(anchor="w", padx=5, pady=2)
        self.labels_alunos.append(lbl)

        self.entry_nome.delete(0, 'end')

    def salvar_excel(self):
        if not self.alunos_temp:
            messagebox.showwarning("Aviso", "Adicione alunos à lista primeiro clicando em +")
            return
            
        sucesso = add_multiple_records(self.alunos_temp)
        if sucesso:
            messagebox.showinfo("Sucesso", f"{len(self.alunos_temp)} registros salvos com sucesso!")
            self.alunos_temp.clear()
            for lbl in self.labels_alunos:
                lbl.destroy()
            self.labels_alunos.clear()
        else:
            messagebox.showerror("Erro", "Erro ao salvar registros. O arquivo Excel pode estar aberto.")

    def pesquisar(self):
        data_pesquisa = self.entry_pesquisa_data.get().strip()
        if not data_pesquisa:
            return

        self.resultados_busca = search_by_date(data_pesquisa)
        self.atualizar_textbox_resultados()

    def atualizar_textbox_resultados(self):
        # Limpar scrollables destruindo filhos
        for cx in [self.caixa_200, self.caixa_300, self.caixa_301]:
            for child in cx.winfo_children():
                child.destroy()

        # Ordem alfa automática
        self.resultados_busca.sort(key=lambda x: x["Nome"].lower())
        
        # Preencher
        for aluno in self.resultados_busca:
            
            # Escolhe a coluna pai
            if aluno['Turma'] == "200 INF":
                parent = self.caixa_200
            elif aluno['Turma'] == "300 INF":
                parent = self.caixa_300
            elif aluno['Turma'] == "301 INF":
                parent = self.caixa_301
            else:
                continue
                
            frame_linha = ctk.CTkFrame(parent, fg_color="transparent")
            frame_linha.pack(fill="x", pady=2, padx=2)
            
            lbl = ctk.CTkLabel(frame_linha, text=f"- {aluno['Nome']}")
            lbl.pack(side="left")
            
            btn_del = ctk.CTkButton(frame_linha, text="X", width=20, height=20, 
                                    fg_color="red", hover_color="darkred", 
                                    command=lambda a=aluno: self.excluir_inline(a))
            btn_del.pack(side="right")

    def excluir_inline(self, aluno):
        data_pesquisa = self.entry_pesquisa_data.get().strip()
        resp = messagebox.askyesno("Confirmar Exclusão", f"Tem certeza que deseja apagar a falta de:\n\n{aluno['Nome']} (Turma {aluno['Turma']})?")
        if resp:
            sucesso = delete_record(data_pesquisa, aluno['Turma'], aluno['Nome'])
            if sucesso:
                self.pesquisar() 
            else:
                messagebox.showerror("Erro", "Registro não encontrado ou arquivo em uso.")

if __name__ == "__main__":
    app = App()
    app.mainloop()
