import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from Crypto.Cipher import AES
from Crypto.Util.Padding import pad, unpad
import hashlib
import threading
import datetime

class BatchDecryptGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Batch Decrypt File")
        self.root.geometry("800x600")

        self.source_folder = tk.StringVar()
        self.destination_folder = tk.StringVar()
        self.encryption_key = tk.StringVar(value="60132323abcd")  # Key diset default dan tidak tampil

        self.total_files = 0
        self.success_count = 0
        self.failure_count = 0

        self.setup_ui()

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Source folder
        ttk.Label(main_frame, text="Folder Sumber:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(main_frame, textvariable=self.source_folder, width=60).grid(row=0, column=1, padx=10)
        ttk.Button(main_frame, text="Pilih", command=self.browse_source).grid(row=0, column=2)

        # Destination folder
        ttk.Label(main_frame, text="Folder Tujuan:").grid(row=1, column=0, sticky=tk.W, pady=(10, 0))
        ttk.Entry(main_frame, textvariable=self.destination_folder, width=60).grid(row=1, column=1, padx=10, pady=(10, 0))
        ttk.Button(main_frame, text="Pilih", command=self.browse_destination).grid(row=1, column=2, pady=(10, 0))

        # Tombol proses
        ttk.Button(main_frame, text="Mulai Proses", command=self.start_decryption).grid(row=2, column=0, columnspan=3, pady=20)

        # Progress bar dan label
        self.progress = ttk.Progressbar(main_frame, orient="horizontal", length=400, mode="determinate")
        self.progress.grid(row=3, column=0, columnspan=3, pady=(0, 10))

        self.progress_label = ttk.Label(main_frame, text="")
        self.progress_label.grid(row=4, column=0, columnspan=3)

        # Kotak log progres
        progress_log_frame = ttk.LabelFrame(main_frame, text="Log Progres File", padding="10")
        progress_log_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        progress_log_frame.columnconfigure(0, weight=1)
        progress_log_frame.rowconfigure(0, weight=1)

        self.progress_text = tk.Text(progress_log_frame, height=10, wrap=tk.WORD)
        self.progress_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        scrollbar = ttk.Scrollbar(progress_log_frame, orient="vertical", command=self.progress_text.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.progress_text.configure(yscrollcommand=scrollbar.set)

        # Kotak log umum
        log_frame = ttk.LabelFrame(main_frame, text="Log Status", padding="10")
        log_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
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
        thread = threading.Thread(target=self.decrypt_all_files)
        thread.start()

    def decrypt_all_files(self):
        src = self.source_folder.get()
        dst = self.destination_folder.get()

        if not os.path.isdir(src):
            messagebox.showerror("Error", "Folder sumber tidak valid.")
            return

        if not os.path.isdir(dst):
            messagebox.showerror("Error", "Folder tujuan tidak valid.")
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

        self.progress["maximum"] = self.total_files

        for idx, file_name in enumerate(files):
            enc_file = os.path.join(src, file_name)
            dec_file = os.path.join(dst, file_name[:-4] + ".xls")  # Ubah ekstensi menjadi .xls
            try:
                self.decrypt_file(enc_file, dec_file, self.encryption_key.get())
                self.success_count += 1
                self.log_message(f"✓ Berhasil: {os.path.basename(enc_file)}")
                self.log_progress(f"✓ {os.path.basename(enc_file)} berhasil didekripsi")
            except Exception as e:
                self.failure_count += 1
                message = f"✗ Gagal: {os.path.basename(enc_file)} - {str(e)}"
                self.log_message(message)
                self.log_progress(f"✗ {os.path.basename(enc_file)} gagal: {message}")

            self.progress["value"] = idx + 1
            self.progress_label.config(text=f"{idx + 1}/{self.total_files} file diproses")
            self.root.update_idletasks()

        summary = f"Selesai! {self.success_count} berhasil, {self.failure_count} gagal dari {self.total_files} file."
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

if __name__ == "__main__":
    root = tk.Tk()
    app = BatchDecryptGUI(root)
    root.mainloop()
