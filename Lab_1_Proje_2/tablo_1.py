#Program-Ders Cıktısı İlişki Matrisi Uygulaması
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import sqlite3
import pandas as pd
import os

class DBComparisonApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Program-Ders Çıktıları İlişki Matrisi")
        self.root.geometry("1200x800")
        
        # Ana pencere çerçevesi
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Matrisin bulunduğu çerçeve
        self.matrix_frame = ttk.Frame(self.main_frame)
        self.matrix_frame.pack(fill=tk.BOTH, expand=True)
        
        # İlişkiler için giriş alanlarını saklayan bir sözlük
        self.relation_entries = {}
        
        # Excel'e aktar butonu
        ttk.Button(self.main_frame, text="Excel'e Aktar", command=self.export_to_excel).pack(pady=10)
        
        # Veritabanı dosyalarının bulunduğu dizin
        current_dir = os.path.dirname(os.path.abspath(__file__))
        self.db1_path = os.path.join(current_dir, "ders_ciktilari.db")
        self.db2_path = os.path.join(current_dir, "program_ciktilari.db")
        self.load_data()

    # Veritabanından tabloları okuma fonksiyonu
    def get_tables(self, db_path, is_first_db=True):
        try:
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            # İlk veritabanı ders verileri, diğeri program verileri
            table_name = "ders_verileri" if is_first_db else "program_verileri"
            cursor.execute(f"SELECT sira_no, aciklama FROM {table_name} ORDER BY sira_no")
            tables = cursor.fetchall()
            conn.close()
            return tables
        
        except Exception as e:
            messagebox.showerror("Hata", f"Veritabanı okuma hatası: {str(e)}")
            return []

    # Verileri yükleme ve matris oluşturma fonksiyonu
    def load_data(self):
        if os.path.exists(self.db1_path) and os.path.exists(self.db2_path):
            self.update_matrix()
        else:
            messagebox.showerror("Hata", "Veritabanı dosyaları bulunamadı!")

    # Belirli bir satırın ilişki değerini hesaplama
    def calculate_relation(self, row_no):
        total = 0
        count = 0
        # Girilen değerleri kontrol ederek toplam ve sayıyı bul
        for (r, c), entry in self.relation_entries.items():
            if r == row_no and entry.get() and entry.get() != '-':
                try:
                    value = float(entry.get())
                    if 1 <= c <= 10:  # Sadece 1-10 arası program çıktıları için hesaplama yap
                        total += value
                        count += 1
                except ValueError:
                    pass
        return total / count if count > 0 else 0

    # Giriş alanında değişiklik olduğunda tetiklenen fonksiyon
    def on_entry_change(self, event):
        entry = event.widget
        if not entry.get():
            entry.insert(0, '-')
        self.update_relation_labels()

    # Tüm ilişki değeri etiketlerini güncelleme
    def update_relation_labels(self):
        for row_no, label in self.relation_labels.items():
            relation = self.calculate_relation(row_no)
            label.config(text=f"{relation:.2f}")

    # Matrisi yeniden oluşturmak
    def update_matrix(self):
        for widget in self.matrix_frame.winfo_children():
            widget.destroy()
        
        # Program ve ders çıktılarını getir
        rows = self.get_tables(self.db2_path, False)  # Program çıktıları
        cols = self.get_tables(self.db1_path, True)   # Ders çıktıları
        
        # Matris başlığı
        ttk.Label(self.matrix_frame, text="Program Çıktısı/\nDers Çıktısı", 
                font=('Arial', 10, 'bold')).grid(row=0, column=0, padx=5, pady=5)
        
        # Ders çıktıları (sütunlar)
        for col_idx, (no, desc) in enumerate(cols, 1):
            ttk.Label(self.matrix_frame, text=str(no), 
                    font=('Arial', 10, 'bold')).grid(row=0, column=col_idx, padx=5, pady=5)
        
        # İlişki Değeri başlığı
        ttk.Label(self.matrix_frame, text="İlişki\nDeğeri", 
                font=('Arial', 10, 'bold')).grid(row=0, column=len(cols)+1, padx=5, pady=5)
        
        self.relation_labels = {}
        
        # Program çıktıları (satırlar)
        for row_idx, (row_no, row_desc) in enumerate(rows, 1):
            ttk.Label(self.matrix_frame, text=str(row_no), 
                    font=('Arial', 10, 'bold')).grid(row=row_idx, column=0, padx=5, pady=5)
            
            for col_idx, (col_no, col_desc) in enumerate(cols, 1):
                # Giriş alanlarını oluşturma
                entry = ttk.Entry(self.matrix_frame, width=5)
                entry.insert(0, '-')
                entry.grid(row=row_idx, column=col_idx, padx=2, pady=2)
                entry.bind('<KeyRelease>', self.on_entry_change)
                entry.bind('<FocusIn>', lambda e: e.widget.delete(0, tk.END))
                self.relation_entries[(row_no, col_no)] = entry
            
            # İlişki değeri etiketi
            relation_label = ttk.Label(self.matrix_frame, text="0.00")
            relation_label.grid(row=row_idx, column=len(cols)+1, padx=5, pady=5)
            self.relation_labels[row_no] = relation_label

    # Matris verilerini Excel'e aktarım fonksiyonu
    def export_to_excel(self):
        try:
            matrix_data = []
            rows = self.get_tables(self.db2_path, False)
            cols = self.get_tables(self.db1_path, True)
            
            for row_no, row_desc in rows:
                row_data = {'Program Çıktısı': row_no, 'Açıklama': row_desc}
                for col_no, col_desc in cols:
                    entry = self.relation_entries.get((row_no, col_no))
                    value = entry.get() if entry and entry.get() != '-' else '0'
                    row_data[f'{col_no}'] = value
                row_data['İlişki Değeri'] = f"{self.calculate_relation(row_no):.2f}"
                matrix_data.append(row_data)
            
            # Excel dosyasını kaydetme
            df = pd.DataFrame(matrix_data)
            current_dir = os.path.dirname(os.path.abspath(__file__))
            save_path = os.path.join(current_dir, "tablo_1.xlsx")
            df.to_excel(save_path, index=False)
            messagebox.showinfo("Başarılı", "İlişki matrisi tablo_1.xlsx olarak kaydedildi!")
        
        except Exception as e:
            messagebox.showerror("Hata", f"Excel kaydetme hatası: {str(e)}")

# Ana uygulama
if __name__ == "__main__":
    root = tk.Tk()
    app = DBComparisonApp(root)
    root.mainloop()