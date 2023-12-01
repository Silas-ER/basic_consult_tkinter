import pandas as pd
import pyodbc 
import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk
from tkcalendar import DateEntry
from datetime import datetime

with open(r'\\servidor\TI\ADM_PROGRAM\config.txt', 'r') as file:
    lines = file.readlines()
    connection_string = ';'.join([line.strip() for line in lines])

try:
    connect = pyodbc.connect(connection_string)
    cursor = connect.cursor()
except pyodbc.Error as e:
    messagebox.showerror("Erro ao conectar no banco", f"Ocorreu um erro ao carregar o banco:\n{str(e)}")


button_style = {"fg": "black", "font": ("Arial", 10)}
logo_path = r"\\servidor\TI\ADM_PROGRAM\logo.png"

try:
    class App(tk.Tk):

        def __init__(self):
            super().__init__()   
            self.title("Cartas de Ordem")  
            self.chave_entry = None

            self.widgets()   
            self.geometry("900x600")

        def head(self):
            try:
                logo = Image.open(logo_path)
                logo = logo.resize((550, 150), Image.LANCZOS)  
                self.logo_tk = ImageTk.PhotoImage(logo)

                label_logo = tk.Label(self, image=self.logo_tk)
                label_logo.pack()
            except Exception as e:
                messagebox.showerror("Erro ao carregar imagem", f"Ocorreu um erro ao carregar a imagem:\n{str(e)}")

        def widgets(self):
            self.limpar_tela()

            self.head()
            frame_botoes = tk.Frame(self)
            frame_botoes.pack()

            botao_materiais = tk.Button(frame_botoes, text="Consultar Cartas de Ordem", command=self.cartas)
            botao_materiais.config(**button_style)
            botao_materiais.pack(padx=10, pady=10)

            botao_sair = tk.Button(frame_botoes, text="Sair do Programa", command=self.fechar_programa)
            botao_sair.config(**button_style)
            botao_sair.pack(padx=10, pady=10)

        def gerar_cartas(self):
            data1m = self.data_inicials.get()
            data2m = self.data_finals.get()

            data1m_formatada = datetime.strptime(data1m, "%d/%m/%Y").strftime("%Y-%m-%d")
            data2m_formatada = datetime.strptime(data2m, "%d/%m/%Y").strftime("%Y-%m-%d")

            print(data1m_formatada)

            consulta_mat = """
                SELECT 
                    TBE.CHAVE_FATO,
                    CONVERT(NVARCHAR, TBE.DATA_MOVTO, 103) AS DATA,
                    CASE
                        WHEN CHARINDEX('NR:', TEO.DESC_MENSAGEM4) > 0 THEN 
                            SUBSTRING(TEO.DESC_MENSAGEM4, CHARINDEX('NR:', TEO.DESC_MENSAGEM4) + LEN('NR:'), 5)
                        ELSE 'NA' 
                    END AS N_CARTA_ORDEM,
                    TBE.COD_TIPO_MV, 
                    TBE.PESO_BRUTO AS PESO,
                    TCG2.APELIDO AS ATRAVESSADOR,
                    TCG.APELIDO AS NOME_DO_FORNECEDOR_NF,
                    TBE.NUM_DOCTO AS NF,  
                    TEO.DESC_MENSAGEM4 AS REFERENCIA,
                    TEO.OBSERVACAO AS BARCO
                FROM 
                    TBENTRADAS TBE
                    LEFT JOIN TBENTRADASOBS TEO ON (TBE.CHAVE_FATO = TEO.CHAVE_FATO)
                    LEFT JOIN TBCADASTROGERAL TCG ON (TBE.COD_CLI_FOR = TCG.COD_CADASTRO)
                    LEFT JOIN TBCADASTROGERAL TCG2 ON (TBE.COD_VEND_COMP = TCG2.COD_CADASTRO)
                WHERE 
                    TBE.COD_TIPO_MV IN ('T221', 'T223') 
                    AND TBE.STATUS NOT IN ('C')
                    AND (
                    (CHARINDEX('NR:', TEO.DESC_MENSAGEM4) > 0 
                    AND SUBSTRING(TEO.DESC_MENSAGEM4, CHARINDEX('NR:', TEO.DESC_MENSAGEM4) + LEN('NR:'), 5) NOT IN ('NA'))
                    )
                    AND DATA_MOVTO >= '{}' AND DATA_MOVTO < '{}'  
                """.format(data1m_formatada, data2m_formatada)
            
            try:
                result = pd.read_sql(consulta_mat, connect)
                result.to_excel(r"\\servidor\Exportação2\Beto\CARTAS_DE_ORDEM\RELACAO_NF_CARTAS_{}.xlsx".format(data2m_formatada), index=False)
                messagebox.showinfo("Sucesso","Consulta realizada com sucesso!") 
            except Exception as e:
                messagebox.showerror("Erro", f"Ocorreu um erro ao executar a consulta:\n{str(e)}")

        
        def cartas(self):
            try:
                self.limpar_tela()
                self.head()

                label = tk.Label(self, text="Consulta de Cartas", font=("Helvetica", 14, "bold"))
                label.pack(pady=10)

                label = tk.Label(self, text="Insira a data inicial: ")
                label.pack(pady=10)
                self.data_inicials = DateEntry(self, date_pattern="dd/mm/yyyy", datefont=('Helvetica', 10), width=12, selectbackground='gray80', locale='pt_BR')
                self.data_inicials.pack(pady=10) 

                label = tk.Label(self, text="Insira a data final: ")
                label.pack(pady=10)
                self.data_finals = DateEntry(self, date_pattern="dd/mm/yyyy", datefont=('Helvetica', 10), width=12, selectbackground='gray80', locale='pt_BR')
                self.data_finals.pack(pady=10) 
                
                botao_materiais = tk.Button(text="Gerar relatório", command=self.gerar_cartas)
                botao_materiais.config(**button_style)
                botao_materiais.pack(pady=10)

                back_button = tk.Button(self, text="<< Back to Menu", command=self.widgets)
                back_button.pack(pady=10)

            except Exception as e: 
                messagebox.showerror("Erro", f"Ocorreu um erro ao carregar calendário\n{str(e)}")
            
        def limpar_tela(self):
            for widget in self.winfo_children():
                widget.destroy()
            self.update()

        def fechar_programa(self):
            self.quit()

except pyodbc.Error as e:
    print(f"Erro ao executar programa: {e}")

if __name__ == "__main__":
    app = App()
    app.mainloop()
