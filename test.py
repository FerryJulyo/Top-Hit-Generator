import hashlib
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
from cryptography.hazmat.backends import default_backend
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter import font as tkFont

class CryptoUtils:
    def __init__(self):
        pass
    
    @staticmethod
    def decrypt_file(key, input_file_path, output_file_path):
        """
        Dekripsi file .enc menggunakan AES dengan key yang di-hash menggunakan SHA-256
        
        Args:
            key (str): Key untuk dekripsi (dalam kasus ini "60132323abcd")
            input_file_path (str): Path ke file .enc yang akan didekripsi
            output_file_path (str): Path untuk menyimpan file hasil dekripsi
        """
        try:
            # Hash key menggunakan SHA-256 (sama seperti kode Java)
            key_bytes = key.encode('utf-8')
            sha256_hash = hashlib.sha256(key_bytes).digest()
            
            # Baca file terenkripsi
            with open(input_file_path, 'rb') as input_file:
                encrypted_data = input_file.read()
            
            # Setup AES cipher dalam mode ECB (sesuai dengan kode Java)
            cipher = Cipher(
                algorithms.AES(sha256_hash),
                modes.ECB(),
                backend=default_backend()
            )
            decryptor = cipher.decryptor()
            
            # Dekripsi data
            decrypted_data = decryptor.update(encrypted_data) + decryptor.finalize()
            
            # Simpan file hasil dekripsi
            with open(output_file_path, 'wb') as output_file:
                output_file.write(decrypted_data)
            
            return True, "Dekripsi berhasil!"
            
        except Exception as e:
            return False, f"Error saat dekripsi: {str(e)}"

def decrypt_enc_file(enc_file_path, output_file_path=None):
    """
    Fungsi utility untuk mendekripsi file .enc
    
    Args:
        enc_file_path (str): Path ke file .enc
        output_file_path (str, optional): Path output. Jika None, akan menggunakan nama file tanpa .enc
    """
    # Key yang digunakan dalam kode Java
    decryption_key = "60132323abcd"
    
    # Jika output path tidak ditentukan, gunakan nama file tanpa ekstensi .enc
    if output_file_path is None:
        if enc_file_path.endswith('.enc'):
            output_file_path = enc_file_path[:-4] + '.xls'  # Hapus .enc dan tambah .xls
        else:
            output_file_path = enc_file_path + '_decrypted.xls'
    
    # Pastikan file input ada
    if not os.path.exists(enc_file_path):
        return False, f"File tidak ditemukan: {enc_file_path}"
    
    # Lakukan dekripsi
    crypto_utils = CryptoUtils()
    return crypto_utils.decrypt_file(decryption_key, enc_file_path, output_file_path)

class DecryptionGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("AES File Decryption Tool")
        self.root.geometry("600x400")
        self.root.resizable(True, True)
        
        # Variabel untuk menyimpan path file
        self.input_file_path = tk.StringVar()
        self.output_file_path = tk.StringVar()
        
        self.setup_ui()
        
    def setup_ui(self):
        # Frame utama
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Konfigurasi grid
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_font = tkFont.Font(family="Arial", size=16, weight="bold")
        title_label = ttk.Label(main_frame, text="AES File Decryption Tool", font=title_font)
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Input file selection
        ttk.Label(main_frame, text="File .enc:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.input_file_path, width=50).grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(10, 10), pady=5)
        ttk.Button(main_frame, text="Browse", command=self.browse_input_file).grid(row=1, column=2, pady=5)
        
        # Output file selection
        ttk.Label(main_frame, text="Save as:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.output_file_path, width=50).grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(10, 10), pady=5)
        ttk.Button(main_frame, text="Browse", command=self.browse_output_file).grid(row=2, column=2, pady=5)
        
        # Info text
        info_text = "Key yang digunakan: 60132323abcd (hardcoded)\nAlgoritma: AES-ECB dengan SHA-256 hash"
        info_label = ttk.Label(main_frame, text=info_text, foreground="gray")
        info_label.grid(row=3, column=0, columnspan=3, pady=(20, 10))
        
        # Decrypt button
        decrypt_btn = ttk.Button(main_frame, text="Decrypt File", command=self.decrypt_file)
        decrypt_btn.grid(row=4, column=0, columnspan=3, pady=20)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        # Status text
        self.status_text = tk.Text(main_frame, height=8, width=70)
        self.status_text.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        
        # Scrollbar untuk status text
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.status_text.yview)
        scrollbar.grid(row=6, column=3, sticky=(tk.N, tk.S), pady=10)
        self.status_text.configure(yscrollcommand=scrollbar.set)
        
        main_frame.rowconfigure(6, weight=1)
        
        self.log_message("Aplikasi siap digunakan. Pilih file .enc untuk dekripsi.")
    
    def browse_input_file(self):
        """Dialog untuk memilih file .enc"""
        file_path = filedialog.askopenfilename(
            title="Pilih file .enc",
            filetypes=[
                ("Encrypted files", "*.enc"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            self.input_file_path.set(file_path)
            # Auto-generate output filename
            if file_path.endswith('.enc'):
                output_path = file_path[:-4] + '.xls'
            else:
                output_path = file_path + '_decrypted.xls'
            self.output_file_path.set(output_path)
            self.log_message(f"File dipilih: {file_path}")
    
    def browse_output_file(self):
        """Dialog untuk memilih lokasi output file"""
        file_path = filedialog.asksaveasfilename(
            title="Simpan file sebagai",
            defaultextension=".xls",
            filetypes=[
                ("Excel files", "*.xls"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            self.output_file_path.set(file_path)
            self.log_message(f"Output akan disimpan di: {file_path}")
    
    def log_message(self, message):
        """Menambahkan pesan ke status text"""
        self.status_text.insert(tk.END, f"{message}\n")
        self.status_text.see(tk.END)
        self.root.update_idletasks()
    
    def decrypt_file(self):
        """Melakukan dekripsi file"""
        input_path = self.input_file_path.get().strip()
        output_path = self.output_file_path.get().strip()
        
        # Validasi input
        if not input_path:
            messagebox.showerror("Error", "Silakan pilih file .enc terlebih dahulu!")
            return
        
        if not output_path:
            messagebox.showerror("Error", "Silakan tentukan lokasi output file!")
            return
        
        if not os.path.exists(input_path):
            messagebox.showerror("Error", f"File tidak ditemukan: {input_path}")
            return
        
        # Mulai progress bar
        self.progress.start()
        self.log_message("Memulai dekripsi...")
        
        try:
            # Lakukan dekripsi
            success, message = decrypt_enc_file(input_path, output_path)
            
            # Stop progress bar
            self.progress.stop()
            
            if success:
                self.log_message(f"✓ Dekripsi berhasil!")
                self.log_message(f"File disimpan di: {output_path}")
                messagebox.showinfo("Sukses", f"Dekripsi berhasil!\nFile disimpan di:\n{output_path}")
            else:
                self.log_message(f"✗ Dekripsi gagal: {message}")
                messagebox.showerror("Error", f"Dekripsi gagal:\n{message}")
                
        except Exception as e:
            self.progress.stop()
            error_msg = f"Terjadi error: {str(e)}"
            self.log_message(f"✗ {error_msg}")
            messagebox.showerror("Error", error_msg)

def main():
    """Fungsi utama untuk menjalankan aplikasi"""
    root = tk.Tk()
    app = DecryptionGUI(root)
    
    # Center window
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f"{width}x{height}+{x}+{y}")
    
    root.mainloop()

if __name__ == "__main__":
    main()