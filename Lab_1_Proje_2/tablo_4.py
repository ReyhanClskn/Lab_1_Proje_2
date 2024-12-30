import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import sqlite3


class StudentOutputCalculator:
    def __init__(self, root):
        # Uygulama penceresinin başlatılması
        self.root = root
        self.root.title("Öğrenci Ders Çıktıları Hesaplama")
        self.root.geometry("1200x800")

        # Tablar için notebook widget'ını oluşturuyoruz
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Tabloları depolayacak liste
        self.trees = []

        # Veri yükleme işlemi
        self.load_data()

    def load_data(self):
        try:
            # Öğrenci notlarını Excel dosyasından okuma
            self.student_grades = pd.read_excel('ogrenci_notlari.xlsx')

            # Öğrenci numaralarını tamsayıya dönüştürme
            self.student_grades['Ogrenci_No'] = self.student_grades['Ogrenci_No'].astype(int)

            # Ders çıktıları matrisini Excel dosyasından okuma
            self.output_matrix = pd.read_excel('tablo_2_3.xlsx', sheet_name='Tablo 3')

            # Ders çıktıları veritabanından alma
            self.ders_ciktilari = self.get_ders_ciktilari_from_db()

            # Öğrenciler için hesaplama yapma
            self.calculate_for_all_students()

        except FileNotFoundError:
            messagebox.showerror("Hata", "Dosya bulunamadı: ogrenci_notlari.xlsx veya tablo_2_3.xlsx. Lütfen dosyaların doğru konumda olduğunu kontrol edin.")
        except Exception as e:
            messagebox.showerror("Hata", f"Dosya okuma hatası: {e}")

    def get_ders_ciktilari_from_db(self):
        try:
            # Veritabanına bağlanma ve ders çıktılarının alınması
            conn = sqlite3.connect('ders_ciktilari.db')
            cursor = conn.cursor()

            # SQL sorgusu ile ders çıktıları verisini alıyoruz
            cursor.execute("SELECT aciklama FROM ders_verileri ORDER BY sira_no")
            results = [row[0] for row in cursor.fetchall()]

            conn.close()
            return results
        except Exception as e:
            messagebox.showerror("Hata", f"Ders çıktıları veritabanından alınamadı: {e}")
            return []

    def calculate_for_all_students(self):
        try:
            # Her öğrenci için hesaplama yapıp tablo oluşturma
            for idx, student in self.student_grades.iterrows():
                student_no = student['Ogrenci_No']
                tab_frame = ttk.Frame(self.notebook)
                self.notebook.add(tab_frame, text=f"Öğrenci {student_no}")

                # Tablo başlıkları
                columns = ("Ders Çıktısı", "Ödev1", "Ödev2", "Quiz", "Vize", "Final", "TOPLAM", "MAX", "% Başarı")
                tree = ttk.Treeview(tab_frame, columns=columns, show="headings")

                # Kolon başlıklarını ayarlama
                for col in columns:
                    tree.heading(col, text=col)
                    tree.column(col, width=100, anchor="center")

                tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

                # Hesaplama ve tablonun oluşturulması
                self.calculate_and_display(tree, student)

                # Tabloyu listeye ekleyerek daha sonra kaydetmek için saklıyoruz
                self.trees.append((student_no, tree))

            # Tüm tabloları Excel'e kaydetmek için bir buton ekliyoruz
            ttk.Button(self.root, text="Tüm Tabloları Excel'e Kaydet", command=self.save_all_to_excel).pack(pady=5)

        except Exception as e:
            messagebox.showerror("Hata", f"Hesaplama hatası: {e}")

    def calculate_and_display(self, tree, student):
        try:
            # Önceden mevcut tüm verileri tablodan temizle
            for item in tree.get_children():
                tree.delete(item)

            # Derecelendirme için kolon eşlemeleri
            column_mapping = {
                'Ödev1': 'Ödev1',
                'Ödev2': 'Ödev2',
                'Quiz': 'Quiz',
                'Vize': 'Vize',
                'Final': 'Final'
            }

            # Tablo 2'yi okuyoruz
            table2_df = pd.read_excel('tablo_2_3.xlsx', sheet_name='Tablo 3')

            # Her ders çıktısı için hesaplama yapma
            for idx, output_row in self.output_matrix.iterrows():
                output_no = output_row['Ders Çıktısı']
                ders_cikti = self.ders_ciktilari[idx] if idx < len(self.ders_ciktilari) else "Bilinmiyor"

                # Öğrenci notları ile hesaplama yapma
                total = 0
                for matrix_col, grade_col in column_mapping.items():
                    if matrix_col in self.output_matrix.columns:
                        weight = float(output_row[matrix_col])
                        grade = float(student[grade_col])
                        total += weight * grade

                # Max değeri alıp başarı oranını hesaplama
                max_value = float(table2_df[table2_df['Ders Çıktısı'] == output_no]['TOPLAM'].iloc[0]) * 100
                success_rate = (total / max_value) * 100 if max_value > 0 else 0

                # Hesaplanan verileri tabloya ekleme
                tree.insert('', 'end', values=(
                    ders_cikti, student['Ödev1'], student['Ödev2'], student['Quiz'],
                    student['Vize'], student['Final'], f"{total:.1f}",
                    f"{max_value:.1f}", f"{success_rate:.1f}"
                ))

        except Exception as e:
            messagebox.showerror("Hata", f"Hesaplama veya görüntüleme hatası: {e}")

    def save_all_to_excel(self):
        try:
            # Her tablodaki verileri alıp Excel'e kaydetme
            for student_no, tree in self.trees:
                data = []
                for item in tree.get_children():
                    data.append(tree.item(item)['values'])

                # DataFrame oluşturulup Excel dosyasına kaydediliyor
                df = pd.DataFrame(data, columns=[
                    "Ders Çıktısı", "Ödev1", "Ödev2", "Quiz", "Vize", "Final",
                    "TOPLAM", "MAX", "% Başarı"
                ])
                filename = f'tablo_4_{int(student_no)}.xlsx'
                df.to_excel(filename, index=False)

            messagebox.showinfo("Başarılı", "Tüm veriler Excel dosyalarına kaydedildi!")

        except Exception as e:
            messagebox.showerror("Hata", f"Excel'e kaydetme hatası: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = StudentOutputCalculator(root)
    root.mainloop()