import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import os

class StudentOutputCalculator:
    def __init__(self, root):
        self.root = root
        self.root.title("Öğrenci Ders Çıktıları Hesaplama")
        self.root.geometry("800x600")

        self.process_button = ttk.Button(
            root, text="Tüm Tablo 5'leri Oluştur", command=self.process_all_students
        )
        self.process_button.pack(pady=20)

    def safe_float_convert(self, value):
        """Güvenli float dönüşümü"""
        try:
            if isinstance(value, str):
                value = value.replace(',', '.')
            return float(value)
        except (ValueError, TypeError):
            return 0.0

    def get_student_numbers(self):
        try:
            df = pd.read_excel("ogrenci_notlari.xlsx")
            return df.iloc[:, 0].astype(str).tolist()
        except Exception as e:
            messagebox.showerror("Hata", f"Öğrenci notları dosyası okuma hatası: {str(e)}")
            return []

    def read_tablo1(self):
        try:
            df = pd.read_excel("tablo_1.xlsx")
            numeric_data = [[self.safe_float_convert(val) for val in row]
                          for row in df.iloc[:, 2:-1].values.tolist()]
            iliski_degerleri = [self.safe_float_convert(val) for val in df.iloc[:, -1].tolist()]
            return numeric_data, iliski_degerleri
        except Exception as e:
            messagebox.showerror("Hata", f"Tablo 1 okuma hatası: {str(e)}")
            return [], []

    def process_all_students(self):
        tablo1_values, iliski_degerleri = self.read_tablo1()
        if not tablo1_values:
            return

        student_numbers = self.get_student_numbers()
        if not student_numbers:
            return

        for student_no in student_numbers:
            try:
                tablo4_filename = f"tablo_4_{student_no}.xlsx"
                if os.path.exists(tablo4_filename):
                    self.create_tablo_5(tablo4_filename, student_no, tablo1_values, iliski_degerleri)
                else:
                    messagebox.showwarning(
                        "Uyarı",
                        f"Tablo 4 dosyası bulunamadı: {student_no}"
                    )
            except Exception as e:
                messagebox.showerror(
                    "Hata",
                    f"İşlem hatası - Öğrenci: {student_no} - Hata: {str(e)}"
                )

        messagebox.showinfo("Başarılı", "Tüm öğrenciler için TABLO 5 oluşturuldu.")

    def create_tablo_5(self, tablo4_filename, student_no, tablo1_values, iliski_degerleri):
        try:
            # Tablo 4'ü oku - sadece % Başarı sütunu
            tablo_4_df = pd.read_excel(tablo4_filename)
            basari_yuzdeleri = [self.safe_float_convert(val) for val in tablo_4_df["% Başarı"].fillna(0).tolist()]

            # Tablo 5 için veri hazırla
            tablo_5_data = []

            # Her program çıktısı için hesaplama
            for row_idx, (tablo1_row, iliski_degeri) in enumerate(zip(tablo1_values, iliski_degerleri), 1):
                row_data = {"Prg Çıktı": row_idx}

                # Her ders çıktısı için çarpma işlemi
                satir_degerleri = []  # Bu satırdaki tüm değerleri toplamak için
                for ders_idx, (basari, tablo1_val) in enumerate(zip(basari_yuzdeleri, tablo1_row), 1):
                    carpim = basari * tablo1_val
                    row_data[f"Ders çıktısı {ders_idx}"] = f"{carpim:.1f}"
                    satir_degerleri.append(carpim)

                # Başarı oranı = O satırdaki değerlerin toplamı / İlişki değeri
                satir_toplami = sum(satir_degerleri)
                basari_orani = satir_toplami / 5 * iliski_degeri if iliski_degeri != 0 else 0
                row_data["Başarı Oranı"] = f"{basari_orani:.1f}"

                tablo_5_data.append(row_data)

            # DataFrame oluştur ve kaydet
            tablo_5_df = pd.DataFrame(tablo_5_data)
            output_filename = f"tablo_5_{student_no}.xlsx"
            tablo_5_df.to_excel(output_filename, index=False)

        except Exception as e:
            raise Exception(f"Tablo 5 oluşturma hatası: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = StudentOutputCalculator(root)
    root.mainloop()
