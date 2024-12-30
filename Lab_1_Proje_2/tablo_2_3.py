#TABLO 2 DERS CIKTISI DEGERLENDIRMELER 
import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
import pandas as pd
import os

class CourseOutputMatrixApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Ders Çıktısı Değerlendirme Matrisi")
        self.root.geometry("1200x800")
        
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        self.weights_frame = ttk.Frame(self.main_frame)
        self.weights_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.matrix_frame = ttk.Frame(self.main_frame)
        self.matrix_frame.pack(fill=tk.BOTH, expand=True)
        
        # Ayırıcı çizgi
        ttk.Separator(self.main_frame, orient='horizontal').pack(fill='x', pady=15)
        
        self.weighted_matrix_frame = ttk.Frame(self.main_frame)
        self.weighted_matrix_frame.pack(fill=tk.BOTH, expand=True)
        
        self.assignments = ["Ödev1", "Ödev2", "Quiz", "Vize", "Final"]
        self.relation_entries = {}
        self.weight_entries = {}
        self.weighted_labels = {}  # Tablo 3 için etiketler
        
        self.create_weights_row()
        self.create_matrix()
        self.create_weighted_matrix()  # Tablo 3'ü oluştur
        ttk.Button(self.main_frame, text="Kaydet", command=self.save_to_excel).pack(pady=10)

    def validate_weight(self, P):
        if P == "":
            return True
        try:
            value = float(P.replace(',', '.'))
            return True
        except ValueError:
            return False

    def validate_matrix_value(self, P):
        if P == "":
            return True
        try:
            value = float(P.replace(',', '.'))
            return 0 <= value <= 1
        except ValueError:
            return False

    def create_weights_row(self):
        ttk.Label(self.weights_frame, text="Ağırlıklar (%)", font=('Arial', 10, 'bold')).grid(row=0, column=0, padx=5, pady=5)
        
        vcmd_weight = (self.root.register(self.validate_weight), '%P')
        
        for idx, assignment in enumerate(self.assignments, 1):
            entry = ttk.Entry(self.weights_frame, width=5, validate='key', validatecommand=vcmd_weight)
            entry.insert(0, "20")
            entry.grid(row=0, column=idx, padx=2, pady=5)
            entry.bind('<FocusIn>', lambda e, entry=entry: self.on_entry_focus_in(entry))
            entry.bind('<KeyRelease>', lambda e: (self.check_weights_sum(), self.update_weighted_matrix()))
            self.weight_entries[assignment] = entry

        self.weights_sum_label = ttk.Label(self.weights_frame, text="Toplam: 100")
        self.weights_sum_label.grid(row=0, column=len(self.assignments)+1, padx=5, pady=5)

    def check_weights_sum(self, event=None):
        total = 0
        for entry in self.weight_entries.values():
            try:
                value = float(entry.get().replace(',', '.')) if entry.get() else 0
                total += value
            except ValueError:
                pass
        
        self.weights_sum_label.config(text=f"Toplam: {total}")
        if total != 100:
            self.weights_sum_label.config(foreground='red')
        else:
            self.weights_sum_label.config(foreground='black')

    def get_course_outputs(self):
        conn = sqlite3.connect("ders_ciktilari.db")
        cursor = conn.cursor()
        cursor.execute("SELECT sira_no, aciklama FROM ders_verileri ORDER BY sira_no")
        outputs = cursor.fetchall()
        conn.close()
        return outputs

    def on_entry_focus_in(self, entry):
        if entry.get() == "0":
            entry.delete(0, tk.END)

    def create_matrix(self):
        outputs = self.get_course_outputs()
        
        ttk.Label(self.matrix_frame, text="TABLO 2", font=('Arial', 12, 'bold')).grid(row=0, column=0, columnspan=len(self.assignments)+2, pady=10)
        ttk.Label(self.matrix_frame, text="Ders Çıktısı", font=('Arial', 10, 'bold')).grid(row=1, column=0, padx=5, pady=5)
        
        vcmd_matrix = (self.root.register(self.validate_matrix_value), '%P')
        
        for idx, assignment in enumerate(self.assignments, 1):
            ttk.Label(self.matrix_frame, text=assignment, font=('Arial', 10, 'bold')).grid(row=1, column=idx, padx=5, pady=5)
        
        ttk.Label(self.matrix_frame, text="TOPLAM", font=('Arial', 10, 'bold')).grid(row=1, column=len(self.assignments)+1, padx=5, pady=5)
        
        self.sum_labels = {}
        for row_idx, (output_no, _) in enumerate(outputs, 2):
            ttk.Label(self.matrix_frame, text=str(output_no), font=('Arial', 10)).grid(row=row_idx, column=0, padx=5, pady=5)
            
            for col_idx, assignment in enumerate(self.assignments, 1):
                entry = ttk.Entry(self.matrix_frame, width=5, validate='key', validatecommand=vcmd_matrix)
                entry.insert(0, "0")
                entry.grid(row=row_idx, column=col_idx, padx=2, pady=2)
                entry.bind('<FocusIn>', lambda e, entry=entry: self.on_entry_focus_in(entry))
                entry.bind('<KeyRelease>', lambda e: (self.calculate_sum(), self.update_weighted_matrix()))
                self.relation_entries[(output_no, assignment)] = entry
            
            sum_label = ttk.Label(self.matrix_frame, text="0")
            sum_label.grid(row=row_idx, column=len(self.assignments)+1, padx=5, pady=5)
            self.sum_labels[output_no] = sum_label

    def create_weighted_matrix(self):
        outputs = self.get_course_outputs()
        
        ttk.Label(self.weighted_matrix_frame, text="TABLO 3", font=('Arial', 12, 'bold')).grid(row=0, column=0, columnspan=len(self.assignments)+2, pady=10)
        ttk.Label(self.weighted_matrix_frame, text="Ders Çıktısı", font=('Arial', 10, 'bold')).grid(row=1, column=0, padx=5, pady=5)
        
        for idx, assignment in enumerate(self.assignments, 1):
            ttk.Label(self.weighted_matrix_frame, text=assignment, font=('Arial', 10, 'bold')).grid(row=1, column=idx, padx=5, pady=5)
        
        ttk.Label(self.weighted_matrix_frame, text="TOPLAM", font=('Arial', 10, 'bold')).grid(row=1, column=len(self.assignments)+1, padx=5, pady=5)
        
        self.weighted_sum_labels = {}
        for row_idx, (output_no, _) in enumerate(outputs, 2):
            ttk.Label(self.weighted_matrix_frame, text=str(output_no), font=('Arial', 10)).grid(row=row_idx, column=0, padx=5, pady=5)
            
            row_labels = {}
            for col_idx, assignment in enumerate(self.assignments, 1):
                label = ttk.Label(self.weighted_matrix_frame, text="0.0")
                label.grid(row=row_idx, column=col_idx, padx=2, pady=2)
                row_labels[assignment] = label
            
            self.weighted_labels[output_no] = row_labels
            
            sum_label = ttk.Label(self.weighted_matrix_frame, text="0.0")
            sum_label.grid(row=row_idx, column=len(self.assignments)+1, padx=5, pady=5)
            self.weighted_sum_labels[output_no] = sum_label

    def calculate_sum(self, event=None):
        for output_no, sum_label in self.sum_labels.items():
            total = 0
            for assignment in self.assignments:
                entry = self.relation_entries[(output_no, assignment)]
                try:
                    value = float(entry.get().replace(',', '.')) if entry.get() else 0
                    total += value
                except ValueError:
                    entry.delete(0, tk.END)
                    entry.insert(0, "0")
            sum_label.config(text=f"{total:.2f}")

    def update_weighted_matrix(self, event=None):
        for output_no in self.weighted_labels:
            weighted_sum = 0
            for assignment in self.assignments:
                try:
                    weight = float(self.weight_entries[assignment].get().replace(',', '.')) / 100
                    value = float(self.relation_entries[(output_no, assignment)].get().replace(',', '.'))
                    weighted_value = weight * value
                    self.weighted_labels[output_no][assignment].config(text=f"{weighted_value:.2f}")
                    weighted_sum += weighted_value
                except (ValueError, AttributeError):
                    self.weighted_labels[output_no][assignment].config(text="0.00")
            
            self.weighted_sum_labels[output_no].config(text=f"{weighted_sum:.2f}")

    def save_to_excel(self):
        try:
            # Ağırlıkların toplamını kontrol et
            total_weight = sum(float(entry.get().replace(',', '.')) for entry in self.weight_entries.values())
            if total_weight != 100:
                messagebox.showerror("Hata", "Ağırlıkların toplamı 100 olmalıdır!")
                return

            data_table2 = []
            data_table3 = []
            weights_data = {'Değerlendirmeler': 'Ağırlıklar (%)'}
            outputs = self.get_course_outputs()
            
            # Ağırlıkları kaydet
            for assignment in self.assignments:
                weights_data[assignment] = float(self.weight_entries[assignment].get().replace(',', '.'))
            
            # Tablo 2'yi kaydet
            for output_no, _ in outputs:
                row_data = {'Ders Çıktısı': output_no}
                for assignment in self.assignments:
                    entry = self.relation_entries[(output_no, assignment)]
                    row_data[assignment] = float(entry.get().replace(',', '.')) if entry.get() else 0
                row_data['TOPLAM'] = float(self.sum_labels[output_no].cget('text'))
                data_table2.append(row_data)
            
            # Tablo 3'ü kaydet
            for output_no, _ in outputs:
                row_data = {'Ders Çıktısı': output_no}
                for assignment in self.assignments:
                    label = self.weighted_labels[output_no][assignment]
                    row_data[assignment] = float(label.cget('text'))
                row_data['TOPLAM'] = float(self.weighted_sum_labels[output_no].cget('text'))
                data_table3.append(row_data)
            
            # Excel'e kaydet
            with pd.ExcelWriter('tablo_2_3.xlsx', engine='openpyxl') as writer:
                # Tablo 2'yi kaydet
                pd.DataFrame([weights_data]).to_excel(writer, sheet_name='Tablo 2', index=False)
                pd.DataFrame(data_table2).to_excel(writer, sheet_name='Tablo 2', startrow=2, index=False)
                # Tablo 3'ü kaydet
                pd.DataFrame(data_table3).to_excel(writer, sheet_name='Tablo 3', index=False)
            
            messagebox.showinfo("Başarılı", "Veriler degerler.xlsx dosyasına kaydedildi!")
            
        except Exception as e:
            messagebox.showerror("Hata", f"Kaydetme hatası: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = CourseOutputMatrixApp(root)
    root.mainloop()