import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import numpy as np

# --- DESIGN ---
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Reconciler V8: Totals & Sort")
        self.geometry("700x750") 
        self.resizable(False, False)

        # HEADER
        self.label_title = ctk.CTkLabel(self, text="Сверка с Итогами", font=("Roboto", 24, "bold"))
        self.label_title.pack(pady=15)

        # === FILEBLOCK ===
        self.frame_files = ctk.CTkFrame(self)
        self.frame_files.pack(pady=10, padx=20, fill="x")

        # FILE 1
        self.btn_f1 = ctk.CTkButton(self.frame_files, text="Файл 1 (Основной)", width=140, command=lambda: self.select_file(self.entry_f1))
        self.btn_f1.grid(row=0, column=0, padx=10, pady=10)
        self.entry_f1 = ctk.CTkEntry(self.frame_files, placeholder_text="Путь к файлу...", width=400)
        self.entry_f1.grid(row=0, column=1, padx=10, pady=10)

        # FILE 2
        self.btn_f2 = ctk.CTkButton(self.frame_files, text="Файл 2 (Для сверки)", width=140, command=lambda: self.select_file(self.entry_f2))
        self.btn_f2.grid(row=1, column=0, padx=10, pady=10)
        self.entry_f2 = ctk.CTkEntry(self.frame_files, placeholder_text="Путь к файлу...", width=400)
        self.entry_f2.grid(row=1, column=1, padx=10, pady=10)

        # === COLUMN SETTINGS BLOCK ===
        
        # -- FILE 1 SETTINGS --
        self.frame_set1 = ctk.CTkFrame(self)
        self.frame_set1.pack(pady=5, padx=20, fill="x")
        ctk.CTkLabel(self.frame_set1, text="Поиск в Файле 1:", font=("Roboto", 14, "bold")).pack(anchor="w", padx=10, pady=5)
        
        self.frame_row1 = ctk.CTkFrame(self.frame_set1, fg_color="transparent")
        self.frame_row1.pack(fill="x")
        
        self.entry_key1 = ctk.CTkEntry(self.frame_row1, placeholder_text="БИН", width=200)
        self.entry_key1.insert(0, "БИН") 
        self.entry_key1.pack(side="left", padx=10, pady=5)
        
        self.entry_val1 = ctk.CTkEntry(self.frame_row1, placeholder_text="Сумма", width=200)
        self.entry_val1.insert(0, "Дебет") 
        self.entry_val1.pack(side="left", padx=10, pady=5)

        self.entry_name1 = ctk.CTkEntry(self.frame_set1, placeholder_text="Контрагент", width=420)
        self.entry_name1.insert(0, "Контрагенты") 
        self.entry_name1.pack(padx=10, pady=10, anchor="w")

        # -- SETTINGS FOR FILE 2 --
        self.frame_set2 = ctk.CTkFrame(self)
        self.frame_set2.pack(pady=5, padx=20, fill="x")
        ctk.CTkLabel(self.frame_set2, text="Поиск в Файле 2:", font=("Roboto", 14, "bold")).pack(anchor="w", padx=10, pady=5)

        self.entry_key2 = ctk.CTkEntry(self.frame_set2, placeholder_text="БИН", width=250)
        self.entry_key2.insert(0, "БИН")
        self.entry_key2.pack(side="left", padx=10, pady=10)

        self.entry_val2 = ctk.CTkEntry(self.frame_set2, placeholder_text="Сумма", width=250)
        self.entry_val2.insert(0, "Дебет")
        self.entry_val2.pack(side="left", padx=10, pady=10)

        # === BUTTON ===
        self.btn_run = ctk.CTkButton(self, text="СФОРМИРОВАТЬ ОТЧЕТ", font=("Roboto", 16, "bold"), height=50, fg_color="#2CC985", hover_color="#229A65", command=self.process_data)
        self.btn_run.pack(pady=20, padx=20, fill="x")

    def select_file(self, entry_widget):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filename:
            entry_widget.delete(0, "end")
            entry_widget.insert(0, filename)

    def find_best_header_row(self, filepath, keywords):
        try:
            df_preview = pd.read_excel(filepath, header=None, nrows=40)
            best_row = 0
            max_matches = 0
            for idx, row in df_preview.iterrows():
                row_text = row.astype(str).str.lower().str.cat(sep=' ')
                matches = 0
                for kw in keywords:
                    if kw.lower().strip() in row_text:
                        matches += 1
                if matches > max_matches:
                    max_matches = matches
                    best_row = idx
            return best_row
        except Exception:
            return 0

    def get_column_name(self, df, search_term):
        search_term = search_term.lower().strip()
        for col in df.columns:
            col_str = str(col).lower().strip()
            if search_term in col_str:
                return col
        return None

    def process_data(self):
        f1, f2 = self.entry_f1.get(), self.entry_f2.get()
        k1_in, v1_in, n1_in = self.entry_key1.get(), self.entry_val1.get(), self.entry_name1.get()
        k2_in, v2_in = self.entry_key2.get(), self.entry_val2.get()

        if not f1 or not f2:
            messagebox.showwarning("Ошибка", "Выберите оба файла!")
            return

        try:
            # SEARCH
            row_1 = self.find_best_header_row(f1, [k1_in, v1_in, n1_in])
            row_2 = self.find_best_header_row(f2, [k2_in, v2_in])
            
            # READING
            df1 = pd.read_excel(f1, header=row_1, dtype=str)
            df2 = pd.read_excel(f2, header=row_2, dtype=str)
            df1.columns = df1.columns.str.strip()
            df2.columns = df2.columns.str.strip()

            # COLUMNS
            k1 = self.get_column_name(df1, k1_in)
            v1 = self.get_column_name(df1, v1_in)
            n1 = self.get_column_name(df1, n1_in)
            k2 = self.get_column_name(df2, k2_in)
            v2 = self.get_column_name(df2, v2_in)

            # REVISE
            if not k1 or not v1 or not n1:
                messagebox.showerror("Ошибка Файл 1", "Не найдены колонки (БИН, Сумма или Контрагент)")
                return
            if not k2 or not v2:
                messagebox.showerror("Ошибка Файл 2", "Не найдены колонки (БИН или Сумма)")
                return

            # NUMBERS SORTING
            try:
                df1[v1] = pd.to_numeric(df1[v1].str.replace(r'\s+', '', regex=True).str.replace(',', '.'), errors='coerce').fillna(0)
                df2[v2] = pd.to_numeric(df2[v2].str.replace(r'\s+', '', regex=True).str.replace(',', '.'), errors='coerce').fillna(0)
            except:
                 pass

            # === LOGIC ===
            # SAVING CURRENT ORDER
            df1['__orig_order__'] = df1.index

            # GROUPING
            df1_g = df1.groupby(k1, as_index=False).agg({
                v1: 'sum', 
                n1: 'first',
                '__orig_order__': 'min'
            }).rename(columns={k1: 'MERGE_KEY', v1: 'Sum_Base1', n1: 'Контрагент'})

            df2_g = df2.groupby(k2, as_index=False)[v2].sum().rename(columns={k2: 'MERGE_KEY', v2: 'Sum_Base2'})

            # MERGER
            merged = pd.merge(df1_g, df2_g, on='MERGE_KEY', how='outer')
            
            merged['Sum_Base1'] = merged['Sum_Base1'].fillna(0)
            merged['Sum_Base2'] = merged['Sum_Base2'].fillna(0)
            merged['Контрагент'] = merged['Контрагент'].fillna("Нет в Базе 1")
            
            merged['Отклонение'] = merged['Sum_Base1'] - merged['Sum_Base2']
            merged['Отклонение'] = merged['Отклонение'].round(2)

            # SORTING FROM ORDER 
            merged = merged.sort_values(by='__orig_order__')
            
            # --- COUNTING OVERALL ---
            total_sum1 = merged['Sum_Base1'].sum()
            total_sum2 = merged['Sum_Base2'].sum()
            total_diff = merged['Отклонение'].sum()

            # CREATING OVERALL STRING
            # COUNTERAGENT CALLED "ИТОГО:", БИН EMPTY
            total_row = pd.DataFrame([{
                'Контрагент': 'ИТОГО ПО ВСЕМУ ОТЧЕТУ:',
                'MERGE_KEY': '', 
                'Sum_Base1': total_sum1,
                'Sum_Base2': total_sum2,
                'Отклонение': total_diff
            }])

            # (concat)
            final_df = pd.concat([merged, total_row], ignore_index=True)

            # SAVING
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if save_path:
                # CHANGING COLUMN ORDER: Сначала Контрагент, потом БИН (MERGE_KEY)
                final_cols = ['Контрагент', 'MERGE_KEY', 'Sum_Base1', 'Sum_Base2', 'Отклонение']
                
                final_df[final_cols].to_excel(save_path, index=False, header=['Наименование', 'БИН', 'Сумма (База 1)', 'Сумма (База 2)', 'Разница'])
                messagebox.showinfo("Успех", f"Отчет готов!\nСтрока итогов добавлена в конец.")

        except Exception as e:
            messagebox.showerror("Критическая ошибка", f"{e}")

if __name__ == "__main__":
    app = App()
    app.mainloop()