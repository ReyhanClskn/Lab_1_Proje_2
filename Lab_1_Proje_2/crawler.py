#CRAWLER
import sqlite3
import requests
from bs4 import BeautifulSoup
import contextlib

def add_data_to_sqlite(url, veritabani_adi, tablo_adi, table_id, timeout=10):
    
    #belirtilen URL'deki tablo verilerini SQLite veritabanına ekler
    #veritabani_adi: Veritabanı dosyasının adı
    #tablo_adi: Verilerin ekleneceği tablonun adı
    #table_id: HTML tablosunun ID'si
    #timeout: HTTP isteği zaman aşımı süresi (saniye)

    with contextlib.closing(sqlite3.connect(veritabani_adi)) as conn:
        with contextlib.closing(conn.cursor()) as cursor: 
            try:
                response = requests.get(url, timeout=timeout)
                response.raise_for_status()

                soup = BeautifulSoup(response.content, "html.parser")
                table = soup.find("table", {"id": table_id})

                if not table:
                    raise ValueError(f"ID'si '{table_id}' olan tablo bulunamadı.")

                rows = table.find_all("tr", class_=["dxgvDataRow_Moderno", "dxgvDataRowAlt_Moderno"])
                data = []
                for row in rows:
                    cells = row.find_all("td")
                    if len(cells) >= 2:
                        try:
                            sira_no = int(cells[0].text.strip())
                            aciklama = cells[1].text.strip()
                            data.append((sira_no, aciklama))
                        except (ValueError, IndexError) as e:
                            print(f"Veri ayrıştırma hatası (satır atlanıyor): {e}, Satır: {cells}, URL: {url}")
                            continue

                cursor.execute(f'''
                    CREATE TABLE IF NOT EXISTS {tablo_adi} (
                        sira_no INTEGER PRIMARY KEY,
                        aciklama TEXT
                    )
                ''')

                cursor.executemany(f"INSERT OR IGNORE INTO {tablo_adi} (sira_no, aciklama) VALUES (?, ?)", data)
                conn.commit()  #commit cursor kapandıktan sonra değil, içindeyken yapılması gerekir 
                print(f"{len(data)} kayıt işlendi, {cursor.rowcount} kayıt {veritabani_adi} veritabanındaki {tablo_adi} tablosuna eklendi. URL: {url}")

            except requests.exceptions.RequestException as e:
                print(f"HTTP isteği hatası: {e}, URL: {url}")
            except (ValueError, sqlite3.Error) as e: 
                print(f"Veritabanı veya veri işleme hatası: {e}, URL: {url}")
            except Exception as e:
                print(f"Beklenmedik hata: {e}, URL: {url}")

#fonksiyon çağrıları 
url1 = "https://ebs.kocaelisaglik.edu.tr/Pages/LearningOutcomesOfProgram.aspx?lang=tr-TR&academicYear=2024&facultyId=5&programId=1&menuType=course&catalogId=2227"
add_data_to_sqlite(url1, "program_ciktilari.db", "program_verileri", "Content_Content_grid_LearningOutComes")

url2 = "https://ebs.kocaelisaglik.edu.tr/Pages/CourseDetail.aspx?lang=tr-TR&academicYear=2024&facultyId=5&programId=1&menuType=course&catalogId=2227"
add_data_to_sqlite(url2, "ders_ciktilari.db", "ders_verileri", "Content_Content_LearningOutcomes_gridLearningOutComes")