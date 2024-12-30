# Program Matris Uygulaması
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import openpyxl

class ProgramMatrisUygulamasi:
    def __init__(self, root):
        self.root = root
        self.root.title("Program Çıktıları İlişki Matrisi")
        self.root.geometry("1200x800")

        # Excel dosyalarını tutan bir sözlük
        self.excel_files = {
            'Ogrenci Notlari': 'ogrenci_notlari.xlsx',
        }

        # Notebook widget (tablı görünüm için)
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(expand=True, fill='both', padx=10, pady=5)
        self.create_tabs()

    # Sekmeleri oluşturur
    def create_tabs(self):
        for tab_name, file_name in self.excel_files.items():
            tab_frame = ttk.Frame(self.notebook)
            self.notebook.add(tab_frame, text=tab_name.upper())
            
            if tab_name == 'Ogrenci Notlari':
                # Öğrenci notları sekmesini özel olarak oluşturur
                self.create_student_notes_tab(tab_frame, file_name)
            else:
                # Diğer sekmeleri Excel dosyasına göre yükler
                self.load_excel_to_tab(tab_frame, file_name)

    # Öğrenci notları sekmesini oluşturur
    def create_student_notes_tab(self, tab_frame, excel_file):
        try:
            # Excel dosyasını yükler
            df = pd.read_excel(excel_file)
            self.weights = {}
            
            try:
                # Ağırlıklar dosyasını yükler
                weights_df = pd.read_excel("weights.xlsx")
                for col in df.columns[1:]:
                    if col != "Ortalama":
                        try:
                            self.weights[col] = weights_df[weights_df["Assignment"] == col]["Weight"].iloc[0] / 100
                        except IndexError:
                            messagebox.showerror("Hata", f"weights.xlsx dosyası '{col}' ödevini içermiyor.")
                            self.weights[col] = 0
            except FileNotFoundError:
                # Eğer ağırlık dosyası yoksa varsayılan ağırlık hesaplar
                num_assignments = len([col for col in df.columns[1:] if col != "Ortalama"])
                default_weight = 1 / num_assignments if num_assignments > 0 else 0
                for col in df.columns[1:]:
                    if col != "Ortalama":
                        self.weights[col] = default_weight

            # Eğer "Ortalama" sütunu yoksa ekler
            if 'Ortalama' not in df.columns:
                df['Ortalama'] = 0.0

            # Ağırlıklı ortalamayı hesaplar
            df['Ortalama'] = self.calculate_weighted_average(df)

            # Treeview widget'i oluşturur
            tree = ttk.Treeview(tab_frame, columns=list(df.columns), show='headings')
            tree.grid(row=0, column=0, columnspan=2, sticky='nsew')

            # Sütun başlıklarını ekler
            for col in df.columns:
                tree.heading(col, text=col)
                tree.column(col, width=100, anchor='center')

            # Treeview'a verileri yükler
            def load_data_to_treeview():
                tree.delete(*tree.get_children())
                for _, row in df.iterrows():
                    row_values = [str(row[col]) if col == df.columns[0] else row[col] for col in df.columns]
                    tree.insert('', 'end', values=row_values)

            # Hücre güncelleme işlemini yönetir
            def update_cell(event, item, column):
                if not item:
                    return
                
                entry = event.widget
                col_index = int(column.replace('#', '')) - 1
                row_id = tree.index(item)
                
                new_value = entry.get()
                
                if df.columns[col_index] == df.columns[0]:  # Eğer "Öğrenci No" ise
                    df.iloc[row_id, col_index] = str(new_value)  # Metin formatında sakla (yoksa ,0 oluyor)
                else:
                    try:
                        new_value = float(new_value)
                        df.iloc[row_id, col_index] = new_value
                        df.loc[row_id, 'Ortalama'] = self.calculate_weighted_average(df.iloc[[row_id]]).iloc[0]
                    except ValueError:
                        messagebox.showerror("Hata", "Geçersiz değer")
                load_data_to_treeview()

            # Hücreye çift tıklama olayını yönetir
            def on_double_click(event):
                item = tree.identify_row(event.y)
                column = tree.identify_column(event.x)
                
                if not item:
                    return
                    
                col_index = int(column.replace('#', '')) - 1
                row_id = tree.index(item)
                cell_value = df.iloc[row_id, col_index]
                
                entry = ttk.Entry(tab_frame)
                entry.insert(0, cell_value)
                entry.place(x=event.x_root - tab_frame.winfo_rootx(), 
                        y=event.y_root - tab_frame.winfo_rooty())
                entry.focus_set()
                
                entry.bind('<Return>', lambda e: [update_cell(e, item, column), entry.destroy()])
                entry.bind('<FocusOut>', lambda e: entry.destroy())

            tree.bind('<Double-1>', on_double_click)
            load_data_to_treeview()

            # Düzenleme çerçevesi oluşturur
            edit_frame = ttk.Frame(tab_frame)
            edit_frame.grid(row=1, column=0, columnspan=2, pady=10)

            # Kaydet butonu
            ttk.Button(edit_frame, text="Kaydet", 
                      command=lambda: self.save_excel(df, excel_file)).grid(row=0, column=0, padx=5)

            # Satır ekleme fonksiyonu
            def add_row():
                new_row = {col: '' if col == df.columns[0] else 0 for col in df.columns}
                df.loc[len(df)] = new_row
                load_data_to_treeview()

            # Satır silme fonksiyonu
            def delete_row():
                selected_item = tree.focus()
                if selected_item:
                    row_id = tree.index(selected_item)
                    df.drop(df.index[row_id], inplace=True)
                    df.reset_index(drop=True, inplace=True)
                    load_data_to_treeview()

            # Ağırlık düzenleme butonları
            ttk.Button(edit_frame, text="Satır Ekle", 
                      command=add_row).grid(row=0, column=1, padx=5)
            ttk.Button(edit_frame, text="Satır Sil", 
                      command=delete_row).grid(row=0, column=2, padx=5)
            ttk.Button(edit_frame, text="Ağırlıkları Düzenle", 
                      command=lambda: self.edit_weights(df, load_data_to_treeview)).grid(row=0, column=3, padx=5)

        except Exception as e:
            messagebox.showerror("Hata", f"Excel dosyası yüklenirken hata oluştu: {str(e)}")

    # Ağırlık düzenleme arayüzü
    def edit_weights(self, df, refresh_callback):
        top = tk.Toplevel(self.root)
        top.title("Ağırlıklar Düzenle")

        tree_weights = ttk.Treeview(top, columns=["Ödev", "Ağırlık (%)"], show="headings")
        tree_weights.pack()

        tree_weights.heading("Ödev", text="Ödev")
        tree_weights.heading("Ağırlık (%)", text="Ağırlık (%)")

        for col in df.columns[1:]:
            if col != "Ortalama":
                tree_weights.insert("", "end", values=(col, int(self.weights[col] * 100)))

        def on_double_click_weights(event):
            item = tree_weights.identify_row(event.y)
            column = tree_weights.identify_column(event.x)
            if column == '#2':
                x, y, width, height = tree_weights.bbox(item, column)
                entry = ttk.Entry(top)
                entry.place(x=x, y=y, width=width, height=height)
                entry.insert(0, tree_weights.set(item, column))

                def save_edit(event=None):
                    try:
                        new_value = int(entry.get())
                        tree_weights.set(item, column, new_value)
                    except ValueError:
                        messagebox.showerror("Hata", "Lütfen bir tamsayı girin.")
                    entry.destroy()

                entry.bind('<Return>', save_edit)
                entry.bind('<FocusOut>', save_edit)
                entry.focus_set()

        tree_weights.bind("<Double-1>", on_double_click_weights)

        def save_weights():
            try:
                new_weights = {}
                total_weight = 0
                for item in tree_weights.get_children():
                    assignment, weight_str = tree_weights.item(item, 'values')
                    weight = float(weight_str) / 100
                    new_weights[assignment] = weight
                    total_weight += weight

                if not (0.99 <= total_weight <= 1.01):  # Allow small floating point differences
                    messagebox.showwarning("Uyarı", "Toplam ağırlık 100 olmalıdır.")
                    return

                self.weights = new_weights
                df['Ortalama'] = self.calculate_weighted_average(df)
                refresh_callback()

                weights_df = pd.DataFrame({
                    "Assignment": list(self.weights.keys()), 
                    "Weight": [int(v * 100) for v in self.weights.values()]
                })
                weights_df.to_excel("weights.xlsx", index=False)
                top.destroy()

            except ValueError:
                messagebox.showerror("Hata", "Ağırlık değeri sayısal olmalıdır.")

        ttk.Button(top, text="Kaydet", command=save_weights).pack(pady=10)

    # Excel dosyasını sekmeye yükler
    def load_excel_to_tab(self, tab_frame, excel_file):
        try:
            df = pd.read_excel(excel_file)
            
            tree = ttk.Treeview(tab_frame, columns=list(df.columns), show="headings")
            for col in df.columns:
                tree.heading(col, text=col)
                tree.column(col, width=100)

            for _, row in df.iterrows():
                tree.insert('', 'end', values=list(row))

            tree.grid(row=0, column=0, sticky='nsew', padx=10, pady=10)
            
            ttk.Button(tab_frame, text="Kaydet", 
                      command=lambda: self.save_excel(df, excel_file)).grid(row=1, column=0, pady=5)

            tab_frame.grid_columnconfigure(0, weight=1)
            tab_frame.grid_rowconfigure(0, weight=1)

        except Exception as e:
            messagebox.showerror("Hata", f"Excel dosyasını yüklenirken hata oluştu: {str(e)}")

    # Verileri Excel dosyasına kaydeder
    def save_excel(self, df, excel_file):
        try:
            df.to_excel(excel_file, index=False)
            messagebox.showinfo("Başarılı", f"Veriler {excel_file} dosyasına kaydedildi.")
        except Exception as e:
            messagebox.showerror("Hata", f"Excel'e kaydetme hatası: {str(e)}")

    def calculate_weighted_average(self, df):
        weighted_sum = pd.Series(0, index=df.index)
        for col in df.columns[1:]:
            if col != "Ortalama":
                try:
                    weighted_sum += df[col] * self.weights.get(col, 0)
                except Exception:
                    pass
        return weighted_sum

if __name__ == "__main__":
    root = tk.Tk()
    app = ProgramMatrisUygulamasi(root)
    root.mainloop()