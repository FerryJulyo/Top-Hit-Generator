import hashlib
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
from cryptography.hazmat.backends import default_backend
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter import font as tkFont
import glob

class CryptoUtils:
    def __init__(self):
        pass
    
    @staticmethod
    def decrypt_file(key, input_file_path, output_file_path):
        """
        Dekripsi file .enc menggunakan AES dengan key yang di-hash menggunakan SHA-256
        
        Args:
            key (str): Key untuk dekripsi
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

class BatchDecryptGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Batch AES Decryption Tool")
        self.root.geometry("750x550")
        self.root.resizable(True, True)
        
        # Variabel untuk menyimpan path folder dan key
        self.input_folder_path = tk.StringVar()
        self.output_folder_path = tk.StringVar()
        self.encryption_key = tk.StringVar(value="60132323abcd")  # Default key
        
        # Variabel untuk tracking progress
        self.total_files = 0
        self.processed_files = 0
        self.successful_files = 0
        self.failed_files = 0
        
        self.setup_ui()
        
    def setup_ui(self):
        # Style configuration
        style = ttk.Style()
        style.configure('Title.TLabel', font=('Arial', 16, 'bold'))
        style.configure('Heading.TLabel', font=('Arial', 10, 'bold'))
        style.configure('Success.TLabel', foreground='green')
        style.configure('Warning.TLabel', foreground='orange')
        style.configure('Error.TLabel', foreground='red')
        
        # Frame utama dengan padding
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Konfigurasi grid untuk responsive design
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Batch AES Decryption Tool", style='Title.TLabel')
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Info label
        info_text = "Dekripsi semua file .enc dalam folder secara batch"
        info_label = ttk.Label(main_frame, text=info_text, foreground="gray")
        info_label.grid(row=1, column=0, columnspan=3, pady=(0, 15))
        
        # Key input
        key_frame = ttk.LabelFrame(main_frame, text="Encryption Key", padding="10")
        key_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        key_frame.columnconfigure(1, weight=1)
        
        ttk.Label(key_frame, text="Key:").grid(row=0, column=0, sticky=tk.W, pady=5)
        key_entry = ttk.Entry(key_frame, textvariable=self.encryption_key, width=30, show="*")
        key_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 10), pady=5)
        
        ttk.Button(key_frame, text="Show/Hide", command=self.toggle_key_visibility).grid(row=0, column=2, pady=5)
        self.key_entry = key_entry  # Store reference for show/hide functionality
        
        # Folder selection frame
        folder_frame = ttk.LabelFrame(main_frame, text="Folder Selection", padding="10")
        folder_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        folder_frame.columnconfigure(1, weight=1)
        
        # Input folder selection
        ttk.Label(folder_frame, text="Input Folder (.enc files):").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(folder_frame, textvariable=self.input_folder_path, width=50).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 10), pady=5)
        ttk.Button(folder_frame, text="Browse", command=self.browse_input_folder).grid(row=0, column=2, pady=5)
        
        # Output folder selection
        ttk.Label(folder_frame, text="Output Folder:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(folder_frame, textvariable=self.output_folder_path, width=50).grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(10, 10), pady=5)
        ttk.Button(folder_frame, text="Browse", command=self.browse_output_folder).grid(row=1, column=2, pady=5)
        
        # File count and options frame
        options_frame = ttk.LabelFrame(main_frame, text="Options & Statistics", padding="10")
        options_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        options_frame.columnconfigure(1, weight=1)
        
        # File count display
        self.file_count_label = ttk.Label(options_frame, text="File .enc ditemukan: 0")
        self.file_count_label.grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        # Checkbox untuk preserve structure
        self.preserve_structure = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Pertahankan struktur subfolder", 
                       variable=self.preserve_structure).grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        # Checkbox untuk overwrite existing files
        self.overwrite_existing = tk.BooleanVar(value=False)
        ttk.Checkbutton(options_frame, text="Timpa file yang sudah ada", 
                       variable=self.overwrite_existing).grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        # Action buttons frame
        action_frame = ttk.Frame(main_frame)
        action_frame.grid(row=5, column=0, columnspan=3, pady=20)
        
        ttk.Button(action_frame, text="Scan Folder", command=self.scan_folder).grid(row=0, column=0, padx=5)
        self.decrypt_btn = ttk.Button(action_frame, text="Decrypt All Files", command=self.decrypt_all_files, state="disabled")
        self.decrypt_btn.grid(row=0, column=1, padx=5)
        ttk.Button(action_frame, text="Clear All", command=self.clear_all).grid(row=0, column=2, padx=5)
        
        # Progress frame
        progress_frame = ttk.LabelFrame(main_frame, text="Progress", padding="10")
        progress_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        progress_frame.columnconfigure(0, weight=1)
        
        # Progress bar
        self.progress = ttk.Progressbar(progress_frame, mode='determinate')
        self.progress.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5)
        
        # Progress label
        self.progress_label = ttk.Label(progress_frame, text="Siap memproses...")
        self.progress_label.grid(row=1, column=0, sticky=tk.W, pady=5)
        
        # Statistics labels
        stats_frame = ttk.Frame(progress_frame)
        stats_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=5)
        stats_frame.columnconfigure(0, weight=1)
        stats_frame.columnconfigure(1, weight=1)
        stats_frame.columnconfigure(2, weight=1)
        
        self.success_label = ttk.Label(stats_frame, text="Berhasil: 0", style='Success.TLabel')
        self.success_label.grid(row=0, column=0, sticky=tk.W)
        
        self.failed_label = ttk.Label(stats_frame, text="Gagal: 0", style='Error.TLabel')
        self.failed_label.grid(row=0, column=1, sticky=tk.W)
        
        self.total_label = ttk.Label(stats_frame, text="Total: 0", style='Heading.TLabel')
        self.total_label.grid(row=0, column=2, sticky=tk.W)
        
        # Status frame
        status_frame = ttk.LabelFrame(main_frame, text="Status Log", padding="10")
        status_frame.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        status_frame.columnconfigure(0, weight=1)
        status_frame.rowconfigure(0, weight=1)
        
        # Status text with scrollbar
        text_frame = ttk.Frame(status_frame)
        text_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)
        
        self.status_text = tk.Text(text_frame, height=10, width=80, wrap=tk.WORD)
        self.status_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=self.status_text.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.status_text.configure(yscrollcommand=scrollbar.set)
        
        # Configure main frame row weights for proper resizing
        main_frame.rowconfigure(7, weight=1)
        
        self.log_message("Aplikasi siap digunakan. Pilih folder yang berisi file .enc untuk dekripsi batch.")
        
    def toggle_key_visibility(self):
        """Toggle visibility of encryption key"""
        current_show = self.key_entry.cget('show')
        if current_show == "*":
            self.key_entry.config(show="")
        else:
            self.key_entry.config(show="*")
    
    def browse_input_folder(self):
        """Dialog untuk memilih input folder"""
        folder_path = filedialog.askdirectory(
            title="Pilih folder yang berisi file .enc"
        )
        
        if folder_path:
            self.input_folder_path.set(folder_path)
            # Auto-generate output folder (subfolder dari input)
            output_path = os.path.join(folder_path, "decrypted")
            self.output_folder_path.set(output_path)
            self.log_message(f"Folder input dipilih: {folder_path}")
            # Auto scan folder
            self.scan_folder()
    
    def browse_output_folder(self):
        """Dialog untuk memilih output folder"""
        folder_path = filedialog.askdirectory(
            title="Pilih folder untuk menyimpan file terdekripsi"
        )
        
        if folder_path:
            self.output_folder_path.set(folder_path)
            self.log_message(f"Folder output dipilih: {folder_path}")
    
    def scan_folder(self):
        """Scan folder untuk mencari file .enc"""
        input_folder = self.input_folder_path.get().strip()
        
        if not input_folder:
            messagebox.showerror("Error", "Silakan pilih folder input terlebih dahulu!")
            return
        
        if not os.path.exists(input_folder):
            messagebox.showerror("Error", f"Folder tidak ditemukan: {input_folder}")
            return
        
        # Cari semua file .enc dalam folder dan subfolder
        if self.preserve_structure.get():
            enc_files = glob.glob(os.path.join(input_folder, "**", "*.enc"), recursive=True)
        else:
            enc_files = glob.glob(os.path.join(input_folder, "*.enc"))
        
        self.total_files = len(enc_files)
        self.file_count_label.config(text=f"File .enc ditemukan: {self.total_files}")
        self.total_label.config(text=f"Total: {self.total_files}")
        
        if self.total_files > 0:
            self.decrypt_btn.config(state="normal")
            self.log_message(f"✓ Ditemukan {self.total_files} file .enc untuk diproses")
            
            # Show list of files in log
            self.log_message("File yang ditemukan:")
            for i, file_path in enumerate(enc_files[:10], 1):  # Show first 10 files
                rel_path = os.path.relpath(file_path, input_folder)
                self.log_message(f"  {i}. {rel_path}")
            
            if self.total_files > 10:
                self.log_message(f"  ... dan {self.total_files - 10} file lainnya")
        else:
            self.decrypt_btn.config(state="disabled")
            self.log_message("⚠ Tidak ada file .enc ditemukan dalam folder")
    
    def clear_all(self):
        """Clear semua input fields"""
        self.input_folder_path.set("")
        self.output_folder_path.set("")
        self.total_files = 0
        self.processed_files = 0
        self.successful_files = 0
        self.failed_files = 0
        self.file_count_label.config(text="File .enc ditemukan: 0")
        self.success_label.config(text="Berhasil: 0")
        self.failed_label.config(text="Gagal: 0")
        self.total_label.config(text="Total: 0")
        self.progress_label.config(text="Siap memproses...")
        self.progress['value'] = 0
        self.decrypt_btn.config(state="disabled")
        self.status_text.delete(1.0, tk.END)
        self.log_message("Form telah dikosongkan.")
    
    def log_message(self, message):
        """Menambahkan pesan ke status log"""
        import datetime
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        self.status_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.status_text.see(tk.END)
        self.root.update_idletasks()
    
    def validate_inputs(self):
        """Validasi input sebelum processing"""
        input_folder = self.input_folder_path.get().strip()
        output_folder = self.output_folder_path.get().strip()
        key = self.encryption_key.get().strip()
        
        if not input_folder:
            messagebox.showerror("Error", "Silakan pilih folder input terlebih dahulu!")
            return False
        
        if not output_folder:
            messagebox.showerror("Error", "Silakan tentukan folder output!")
            return False
        
        if not key:
            messagebox.showerror("Error", "Silakan masukkan encryption key!")
            return False
        
        if not os.path.exists(input_folder):
            messagebox.showerror("Error", f"Folder input tidak ditemukan: {input_folder}")
            return False
        
        if self.total_files == 0:
            messagebox.showerror("Error", "Tidak ada file .enc untuk diproses!")
            return False
        
        return True
    
    def decrypt_all_files(self):
        """Proses dekripsi semua file .enc dalam folder"""
        if not self.validate_inputs():
            return
        
        input_folder = self.input_folder_path.get().strip()
        output_folder = self.output_folder_path.get().strip()
        key = self.encryption_key.get().strip()
        
        # Reset counters
        self.processed_files = 0
        self.successful_files = 0
        self.failed_files = 0
        
        # Create output folder if not exists
        try:
            os.makedirs(output_folder, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Error", f"Tidak dapat membuat folder output: {str(e)}")
            return
        
        # Get list of .enc files
        if self.preserve_structure.get():
            enc_files = glob.glob(os.path.join(input_folder, "**", "*.enc"), recursive=True)
        else:
            enc_files = glob.glob(os.path.join(input_folder, "*.enc"))
        
        # Setup progress bar
        self.progress['maximum'] = len(enc_files)
        self.progress['value'] = 0
        
        self.log_message(f"Memulai dekripsi {len(enc_files)} file...")
        
        crypto_utils = CryptoUtils()
        
        for enc_file in enc_files:
            try:
                # Update progress
                self.processed_files += 1
                self.progress['value'] = self.processed_files
                
                # Determine output file path
                if self.preserve_structure.get():
                    rel_path = os.path.relpath(enc_file, input_folder)
                    output_file = os.path.join(output_folder, rel_path[:-4] + '.xls')  # Remove .enc, add .xls
                else:
                    filename = os.path.basename(enc_file)[:-4] + '.xls'  # Remove .enc, add .xls
                    output_file = os.path.join(output_folder, filename)
                
                # Create output subdirectory if needed
                output_dir = os.path.dirname(output_file)
                os.makedirs(output_dir, exist_ok=True)
                
                # Check if output file already exists
                if os.path.exists(output_file) and not self.overwrite_existing.get():
                    self.log_message(f"⚠ Dilewati (sudah ada): {os.path.basename(enc_file)}")
                    continue
                
                # Update progress label
                filename = os.path.basename(enc_file)
                self.progress_label.config(text=f"Memproses: {filename} ({self.processed_files}/{len(enc_files)})")
                self.root.update_idletasks()
                
                # Perform decryption
                success, message = crypto_utils.decrypt_file(key, enc_file, output_file)
                
                if success:
                    self.successful_files += 1
                    self.log_message(f"✓ Berhasil: {os.path.basename(enc_file)}")
                else:
                    self.failed_files += 1
                    self.log_message(f"✗ Gagal: {os.path.basename(enc_file)} - {message}")
                
                # Update statistics
                self.success_label.config(text=f"Berhasil: {self.successful_files}")
                self.failed_label.config(text=f"Gagal: {self.failed_files}")
                
            except Exception as e:
                self.failed_files += 1
                self.log_message(f"✗ Error: {os.path.basename(enc_file)} - {str(e)}")
                self.failed_label.config(text=f"Gagal: {self.failed_files}")
        
        # Completion
        self.progress_label.config(text="Proses selesai!")
        self.log_message(f"🎉 Proses dekripsi selesai!")
        self.log_message(f"   Berhasil: {self.successful_files} file")
        self.log_message(f"   Gagal: {self.failed_files} file")
        self.log_message(f"   Total diproses: {self.processed_files} file")
        
        # Show completion dialog
        if self.failed_files == 0:
            result = messagebox.showinfo("Sukses", 
                f"Semua file berhasil didekripsi!\n\n"
                f"Berhasil: {self.successful_files} file\n"
                f"Lokasi: {output_folder}")
        else:
            result = messagebox.showwarning("Selesai dengan Warning", 
                f"Proses dekripsi selesai dengan beberapa error.\n\n"
                f"Berhasil: {self.successful_files} file\n"
                f"Gagal: {self.failed_files} file\n"
                f"Lokasi: {output_folder}")
        
        # Ask if user wants to open output folder
        if messagebox.askyesno("Buka Folder", "Apakah Anda ingin membuka folder output?"):
            self.open_output_folder(output_folder)
    
    def open_output_folder(self, folder_path):
        """Buka folder output"""
        try:
            import subprocess
            import platform
            
            if platform.system() == "Windows":
                subprocess.Popen(f'explorer "{folder_path}"')
            elif platform.system() == "Darwin":  # macOS
                subprocess.Popen(["open", folder_path])
            else:  # Linux
                subprocess.Popen(["xdg-open", folder_path])
        except Exception as e:
            self.log_message(f"Tidak dapat membuka folder: {str(e)}")

def main():
    """Fungsi utama untuk menjalankan aplikasi"""
    root = tk.Tk()
    
    # Set window icon (jika ada)
    try:
        # root.iconbitmap('icon.ico')  # Uncomment jika ada icon file
        pass
    except:
        pass
    
    app = BatchDecryptGUI(root)
    
    # Center window
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f"{width}x{height}+{x}+{y}")
    
    # Set minimum window size
    root.minsize(700, 500)
    
    root.mainloop()

if __name__ == "__main__":
    main()