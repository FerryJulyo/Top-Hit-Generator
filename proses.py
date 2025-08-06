import os
import io
import msoffcrypto
import tkinter as tk
import re
import hashlib
import threading
import datetime
import pandas as pd
import xlrd
import tempfile

from openpyxl import load_workbook
from tempfile import NamedTemporaryFile
from xlrd.biffh import XLRDError
from tkinter import filedialog, messagebox, ttk
from Crypto.Cipher import AES
from Crypto.Util.Padding import pad, unpad
from collections import defaultdict

class BatchDecryptGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Generate Top Hit")
        self.root.geometry("750x600")

        self.source_folder = tk.StringVar()
        self.destination_folder = tk.StringVar()
        self.database_file = tk.StringVar()
        self.encryption_key = tk.StringVar(value="60132323abcd")

        self.total_files = 0
        self.success_count = 0
        self.failure_count = 0

        self.df_db = None

        self.setup_ui()

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="Folder Sumber:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(main_frame, textvariable=self.source_folder, width=60).grid(row=0, column=1, padx=10)
        self.btn_browse_source = ttk.Button(main_frame, text="Pilih", command=self.browse_source)
        self.btn_browse_source.grid(row=0, column=2)

        ttk.Label(main_frame, text="Folder Tujuan:").grid(row=1, column=0, sticky=tk.W, pady=(10, 0))
        ttk.Entry(main_frame, textvariable=self.destination_folder, width=60).grid(row=1, column=1, padx=10, pady=(10, 0))
        self.btn_browse_destination = ttk.Button(main_frame, text="Pilih", command=self.browse_destination)
        self.btn_browse_destination.grid(row=1, column=2, pady=(10, 0))

        ttk.Label(main_frame, text="File Database (xlsx):").grid(row=2, column=0, sticky=tk.W, pady=(10, 0))
        ttk.Entry(main_frame, textvariable=self.database_file, width=60).grid(row=2, column=1, padx=10, pady=(10, 0))
        self.btn_browse_database = ttk.Button(main_frame, text="Pilih", command=self.browse_database)
        self.btn_browse_database.grid(row=2, column=2, pady=(10, 0))

        self.btn_start = ttk.Button(main_frame, text="Mulai Proses", command=self.start_decryption)
        self.btn_start.grid(row=3, column=0, columnspan=3, pady=20)

        self.progress = ttk.Progressbar(main_frame, orient="horizontal", length=400, mode="determinate")
        self.progress.grid(row=4, column=0, columnspan=3, pady=(0, 10))

        self.progress_label = ttk.Label(main_frame, text="")
        self.progress_label.grid(row=5, column=0, columnspan=3)

        progress_log_frame = ttk.LabelFrame(main_frame, text="Log Progres File", padding="10")
        progress_log_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        progress_log_frame.columnconfigure(0, weight=1)
        progress_log_frame.rowconfigure(0, weight=1)

        self.progress_text = tk.Text(progress_log_frame, height=10, wrap=tk.WORD)
        self.progress_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        scrollbar = ttk.Scrollbar(progress_log_frame, orient="vertical", command=self.progress_text.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.progress_text.configure(yscrollcommand=scrollbar.set)

        log_frame = ttk.LabelFrame(main_frame, text="Log Status", padding="10")
        log_frame.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = tk.Text(log_frame, height=6, wrap=tk.WORD)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.log_text.configure(yscrollcommand=scrollbar.set)

    def browse_source(self):
        folder = filedialog.askdirectory()
        if folder:
            self.source_folder.set(folder)

    def browse_destination(self):
        folder = filedialog.askdirectory()
        if folder:
            self.destination_folder.set(folder)

    def browse_database(self):
        file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file:
            self.database_file.set(file)

    def log_message(self, message):
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def log_progress(self, message):
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        self.progress_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.progress_text.see(tk.END)
        self.root.update_idletasks()

    def start_decryption(self):
        self.set_buttons_state(tk.DISABLED)
        thread = threading.Thread(target=self.decrypt_all_files)
        thread.start()

    def set_buttons_state(self, state):
        self.btn_browse_source.config(state=state)
        self.btn_browse_destination.config(state=state)
        self.btn_browse_database.config(state=state)
        self.btn_start.config(state=state)


    def decrypt_all_files(self):
        src = self.source_folder.get()
        dst = self.destination_folder.get()
        db_path = self.database_file.get()

        if not os.path.isdir(src):
            messagebox.showerror("Error", "Folder sumber tidak valid.")
            return
        if not os.path.isdir(dst):
            messagebox.showerror("Error", "Folder tujuan tidak valid.")
            return
        if not os.path.isfile(db_path):
            messagebox.showerror("Error", "File database tidak ditemukan.")
            return

        try:
            self.log_message("Sedang membaca file master...")
            self.df_db = pd.read_excel(db_path, sheet_name="Song", usecols=["SongId", "Song", "Sing1", "Sing2", "Sing3", "Sing4", "Sing5"])
        except Exception as e:
            self.log_message(f"Gagal membaca database: {e}")
            return

        self.total_files = 0
        self.success_count = 0
        self.failure_count = 0
        self.progress_label.config(text="")
        self.progress_text.delete(1.0, tk.END)
        self.log_text.delete(1.0, tk.END)

        files = [f for f in os.listdir(src) if f.endswith(".enc")]
        self.total_files = len(files)

        if self.total_files == 0:
            self.log_message("Tidak ada file .enc ditemukan di folder sumber.")
            return

        total_export_steps = len(files) + len(set(f[:5] for f in files)) + 1
        self.progress["maximum"] = total_export_steps

        decrypted_files = []

        for idx, file_name in enumerate(files):
            enc_file = os.path.join(src, file_name)
            dec_file = os.path.join(dst, file_name[:-4] + ".xls")
            try:
                self.decrypt_file(enc_file, dec_file, self.encryption_key.get())
                decrypted_files.append(dec_file)
                self.success_count += 1
                self.log_message(f"✓ Berhasil: {os.path.basename(enc_file)}")
                self.log_progress(f"✓ {os.path.basename(enc_file)} berhasil didekripsi")
            except Exception as e:
                self.failure_count += 1
                self.log_message(f"✗ Gagal: {os.path.basename(enc_file)} - {str(e)}")
                self.log_progress(f"✗ {os.path.basename(enc_file)} gagal: {str(e)}")

            self.progress["value"] = idx + 1
            self.progress_label.config(text=f"{idx + 1}/{self.total_files} file diproses")
            self.root.update_idletasks()

        self.process_and_merge_data(decrypted_files, dst, progress_update_callback=self.update_progress)
        self.progress["value"] = self.progress["maximum"]


        summary = f"Selesai! {self.success_count} berhasil, {self.failure_count} gagal dari {self.total_files} file."
        self.set_buttons_state(tk.NORMAL)
        self.log_message(summary)
        messagebox.showinfo("Selesai", summary)

    def decrypt_file(self, in_file, out_file, key):
        hashed_key = hashlib.sha256(key.encode()).digest()
        cipher = AES.new(hashed_key, AES.MODE_ECB)

        with open(in_file, "rb") as f:
            ciphertext = f.read()
        plaintext = unpad(cipher.decrypt(ciphertext), AES.block_size)

        with open(out_file, "wb") as f:
            f.write(plaintext)

    def process_and_merge_data(self, file_list, output_path, progress_update_callback=None):
        result = defaultdict(lambda: defaultdict(int))
        all_data = defaultdict(lambda: defaultdict(int))  # 🆕 Rekap semua data

        # Buat dictionary lookup dari master
        song_dict = {}
        if self.df_db is not None:
            for _, row in self.df_db.iterrows():
                song_id = str(row["SongId"]).strip()
                song_name = str(row["Song"]).strip() if not pd.isna(row["Song"]) else ""
                singers = [str(row[col]).strip() for col in ["Sing1", "Sing2", "Sing3", "Sing4", "Sing5"] if not pd.isna(row[col]) and str(row[col]).strip()]
                song_dict[song_id] = {
                    "Song": song_name,
                    "Singer": " - ".join(singers)
                }

        # Proses semua file terenkripsi
        for file in file_list:
            try:
                # 1. Verifikasi file exists
                if not os.path.exists(file):
                    self.log_message(f"File tidak ditemukan: {file}")
                    continue

                # 2. Coba baca langsung dengan pandas (untuk file tidak terenkripsi)
                try:
                    # 1. Coba baca langsung tanpa dekripsi
                    df = pd.read_excel(file, sheet_name="Lap1")
                    self.log_message(f"File {file} berhasil dibaca tanpa dekripsi")
                except:
                    # 2. Jika gagal, coba dekripsi dengan msoffcrypto
                    try:
                        with open(file, "rb") as f:
                            office_file = msoffcrypto.OfficeFile(f)
                            office_file.load_key(password="secret")  # Ganti dengan password yang sesuai
                            
                            # Decrypt file ke dalam memory
                            decrypted = io.BytesIO()
                            office_file.decrypt(decrypted)

                        # Simpan ke temporary file dari hasil decrypt
                        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as temp_file:
                            temp_file.write(decrypted.getvalue())
                            temp_path = temp_file.name

                        try:
                            # 3. Coba baca dari hasil decrypt
                            df = pd.read_excel(temp_path, sheet_name="Lap1")
                            self.log_message(f"File {file} berhasil dibaca setelah dekripsi")
                        except Exception as e:
                            self.log_message(f"Gagal baca file setelah dekripsi: {str(e)}")
                            continue
                        finally:
                            # 4. Hapus file sementara
                            try:
                                os.unlink(temp_path)
                            except:
                                pass

                    except Exception as e:
                        self.log_message(f"Gagal proses file terenkripsi {file}: {str(e)}")
                        continue

                # 4. Proses data dari DataFrame
                group_key = os.path.basename(file)[:5]
                
                for index, row in df.iterrows():
                    try:
                        # Gunakan iloc untuk akses kolom by index
                        song_id = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
                        jumlah_cell = row.iloc[1] if len(row) > 1 else None
                        
                        if not song_id or pd.isna(jumlah_cell):
                            continue
                            
                        try:
                            # Handle berbagai format angka
                            jumlah = int(float(str(jumlah_cell))) if str(jumlah_cell).strip() else 0
                            if jumlah == 0:
                                continue
                                
                            # Update hasil
                            result[group_key][song_id] += jumlah
                            all_data["ALL"][song_id] += jumlah
                        except ValueError:
                            print(f"[WARN] Baris {index+1} di file {file}: jumlah tidak valid → {jumlah_cell}")
                            
                    except Exception as e:
                        print(f"[ERROR] Baris {index+1} di file {file}: {str(e)}")
                        continue

            except Exception as e:
                self.log_message(f"Error utama saat proses file {file}: {str(e)}")

        # Simpan output per group
        for group, items in result.items():
            lang_data = defaultdict(list)

            for song_id_raw, jumlah in items.items():
                # 1️⃣ Normalisasi SongId
                song_id_clean = re.sub(r"[A-Za-z]", "", song_id_raw)  # Hapus huruf
                song_id_clean = song_id_clean.lstrip("0")             # Hapus prefix 0

                # 2️⃣ Kategorisasi berdasarkan awal angka
                lang = "Lain-Lain"
                if song_id_clean.startswith(("10", "11", "12", "13", "14", "15", "16", "17", "19")):
                    lang = "Indonesia Pop"
                elif song_id_clean.startswith("18"):
                    lang = "Indonesia Daerah"
                elif song_id_clean.startswith("2"):
                    lang = "English"
                elif song_id_clean.startswith("3"):
                    lang = "Mandarin"
                elif song_id_clean.startswith("4"):
                    lang = "Jepang"
                elif song_id_clean.startswith("5"):
                    lang = "Korea"

                # 3️⃣ Cari kecocokan master SongId (LIKE match)
                matched_info = {"Song": "", "Singer": ""}
                final_song_id = song_id_clean  # default
                for master_id, info in song_dict.items():
                    if master_id and song_id_clean in master_id:
                        matched_info = info
                        if master_id != song_id_clean:
                            final_song_id = master_id  # ganti jika tidak persis sama
                        break

                # 4️⃣ Tambahkan ke kategori lang
                lang_data[lang].append({
                    "Judul Lagu": matched_info["Song"],
                    "Penyanyi": matched_info["Singer"],
                    "Jumlah Pengguna": jumlah,
                    "ID": final_song_id
                })

            # Buat Excel writer untuk setiap group
            output_file = os.path.join(output_path, f"IDLAGU_Outlet_{group}.xlsx")
            with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                for lang, records in lang_data.items():
                    df_sheet = pd.DataFrame(records)
                    df_sheet.sort_values(by="Jumlah Pengguna", ascending=False, inplace=True)
                    df_sheet.to_excel(writer, sheet_name=lang[:31], index=False)

            self.log_message(f"File berhasil dibuat: {output_file}")
            if progress_update_callback:
                progress_update_callback()


        # 🔽 Simpan file rekap semua grup
        if all_data["ALL"]:
            lang_data_all = defaultdict(list)
            for song_id_raw, jumlah in all_data["ALL"].items():
                # 🔁 Normalisasi
                song_id_clean = re.sub(r"[A-Za-z]", "", song_id_raw).lstrip("0")

                # 🔁 Kategorisasi
                lang = "Lain-Lain"
                if song_id_clean.startswith(("10", "11", "12", "13", "14", "15", "16", "17", "19")):
                    lang = "Indonesia Pop"
                elif song_id_clean.startswith("18"):
                    lang = "Indonesia Daerah"
                elif song_id_clean.startswith("2"):
                    lang = "English"
                elif song_id_clean.startswith("3"):
                    lang = "Mandarin"
                elif song_id_clean.startswith("4"):
                    lang = "Jepang"
                elif song_id_clean.startswith("5"):
                    lang = "Korea"

                # 🔁 Match master
                matched_info = {"Song": "", "Singer": ""}
                final_song_id = song_id_clean
                for master_id, info in song_dict.items():
                    if master_id and song_id_clean in master_id:
                        matched_info = info
                        if master_id != song_id_clean:
                            final_song_id = master_id
                        break

                lang_data_all[lang].append({
                    "Judul Lagu": matched_info["Song"],
                    "Penyanyi": matched_info["Singer"],
                    "Jumlah Pengguna": jumlah,
                    "ID": final_song_id
                })

            # 🔁 Simpan ke file
            output_file_all = os.path.join(output_path, "IDLAGU_ALL.xlsx")
            with pd.ExcelWriter(output_file_all, engine="openpyxl") as writer:
                for lang, records in lang_data_all.items():
                    df_sheet = pd.DataFrame(records)
                    df_sheet = df_sheet.groupby("ID", as_index=False).agg({
                        "Judul Lagu": "first",
                        "Penyanyi": "first",
                        "Jumlah Pengguna": "sum"
                    })

                    df_sheet.sort_values(by="Jumlah Pengguna", ascending=False, inplace=True)
                    df_sheet.to_excel(writer, sheet_name=lang[:31], index=False)

            self.log_message(f"File gabungan berhasil dibuat: {output_file_all}")
            if progress_update_callback:
                progress_update_callback()

    def update_progress(self):
        self.progress["value"] += 1
        self.root.update_idletasks()


if __name__ == "__main__":
    root = tk.Tk()
    app = BatchDecryptGUI(root)
    root.mainloop()
