import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import pandas as pd
from excel_logic import ExcelHandler
from sap_logic import SAPAutomation
import os

class SAPAutomationApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Automação SAP - Conversão VBA para Python")
        self.geometry("500x400")

        self.excel_path = ""
        self.running = False
        self.stop_requested = False
        
        self.sap = SAPAutomation()
        self.excel = None

        # Layout da Interface
        self.label_title = ctk.CTkLabel(self, text="Controle de Automação SAP", font=("Arial", 20))
        self.label_title.pack(pady=20)

        self.btn_select_file = ctk.CTkButton(self, text="Selecionar Excel", command=self.select_file)
        self.btn_select_file.pack(pady=10)

        self.label_file = ctk.CTkLabel(self, text="Nenhum arquivo selecionado", wraplength=400)
        self.label_file.pack(pady=5)

        self.btn_start = ctk.CTkButton(self, text="Iniciar Processamento", command=self.start_processing, state="disabled")
        self.btn_start.pack(pady=10)

        self.label_status = ctk.CTkLabel(self, text="Status: Aguardando arquivo...")
        self.label_status.pack(pady=10)

        self.label_counters = ctk.CTkLabel(self, text="Processados: 0 | Erros: 0")
        self.label_counters.pack(pady=10)

        self.btn_stop = ctk.CTkButton(self, text="Parar", command=self.stop_processing, state="disabled", fg_color="red")
        self.btn_stop.pack(pady=10)

        self.processed_count = 0
        self.error_count = 0

    def select_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xlsm")])
        if file_path:
            self.excel_path = file_path
            self.label_file.configure(text=os.path.basename(file_path))
            self.btn_start.configure(state="normal")
            self.excel = ExcelHandler(file_path)

    def update_counters(self):
        self.label_counters.configure(text=f"Processados: {self.processed_count} | Erros: {self.error_count}")

    def stop_processing(self):
        self.stop_requested = True
        self.label_status.configure(text="Status: Parando...")

    def start_processing(self):
        if not self.running:
            self.running = True
            self.stop_requested = False
            self.btn_start.configure(state="disabled")
            self.btn_stop.configure(state="normal")
            threading.Thread(target=self.process_loop, daemon=True).start()

    def process_loop(self):
        try:
            df = self.excel.get_data()
            total_rows = len(df)
            current_row = 0

            while current_row < total_rows:
                if self.stop_requested:
                    break

                # Conexão SAP
                success, msg = self.sap.connect()
                if not success:
                    self.label_status.configure(text=f"Status: {msg}")
                    break

                row_data = df.iloc[current_row]
                status_val = str(row_data['U']).strip()

                # Pular linhas vazias ou já processadas
                if pd.isna(row_data['A']) or str(row_data['A']).strip() == "" or status_val == "OK":
                    current_row += 1
                    continue

                # Agrupamento por Nota (Coluna Q)
                nota_atual = row_data['Q']
                group_end = current_row
                while group_end + 1 < total_rows and df.iloc[group_end + 1]['Q'] == nota_atual and str(df.iloc[group_end + 1]['U']) != "OK":
                    group_end += 1
                
                df_group = df.iloc[current_row : group_end + 1]
                self.label_status.configure(text=f"Processando Nota: {nota_atual}")
                
                # Executa Lógica SAP
                result = self.sap.process_row_group(df_group, self.excel, current_row, group_end)

                if result == "MANUAL_ACTION_REQUIRED":
                    ans = messagebox.askokcancel("Ação Manual", "Vá até a J1B3N, imprima e clique OK para continuar.")
                    if ans:
                        continue 
                    else:
                        break

                # Atualiza Excel e Interface
                for i in range(current_row, group_end + 1):
                    self.excel.update_status(i, result)
                    df.at[i, 'U'] = result

                if result == "OK":
                    self.processed_count += 1
                elif result != "AUTCARR_OK":
                    self.error_count += 1
                
                self.update_counters()
                
                if result == "AUTCARR_OK":
                    pass # Re-processa o grupo para o Passo 2 (NFR)
                else:
                    current_row = group_end + 1

            self.label_status.configure(text="Status: Finalizado")
            messagebox.showinfo("Fim", "Processamento concluído!")

        except Exception as e:
            self.label_status.configure(text=f"Erro: {str(e)}")
        finally:
            self.running = False
            self.btn_start.configure(state="normal")
            self.btn_stop.configure(state="disabled")

if __name__ == "__main__":
    app = SAPAutomationApp()
    app.mainloop()
