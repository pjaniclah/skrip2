# -*- coding: utf-8 -*-
import os
import subprocess
from Tkinter import Tk, Label, StringVar, Canvas
import ttk
import tkFileDialog as filedialog
import tkMessageBox as messagebox

# ========================
# GUI Setup
# ========================
root = Tk()
root.title("GIS Exporter & Updater v2025.RC2")
root.geometry("550x300")
notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True, padx=8, pady=8)

# ========================
# Global Variables
# ========================


# ========================EXCEL
def browse_file(var, ext, label):
    path = filedialog.askopenfilename(title="Pilih {}".format(label), filetypes=[(label, ext)])
    if path:
        var.set(path)

def browse_folder(var):
    folder = filedialog.askdirectory(title="Pilih Folder")
    if folder:
        var.set(folder)

def get_script_path(script_name):
    base_folder = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_folder, script_name)

def run_export_excel():
    excel = excel_path.get()
    out_folder = excel_out_folder.get()
    script = get_script_path("export_excel_arcmap.pyc")

    if not os.path.isfile(excel):
        messagebox.showerror("Error", "File Excel tidak valid.")
        return
    if not os.path.isdir(out_folder):
        messagebox.showerror("Error", "Folder output tidak valid.")
        return
    if not os.path.isfile(script):
        messagebox.showerror("Script Tidak Ditemukan", "export_excel_arcmap.pyc tidak ditemukan.")
        return

    ret = subprocess.call(["C:/Python27/ArcGIS10.8/python.exe", script, excel, out_folder])
    if ret == 0:
        messagebox.showinfo("Selesai", "Ekspor Excel berhasil.")
    else:
        messagebox.showerror("Gagal", "Skrip berakhir dengan kode: {}".format(ret))

# ========================KMZ
def run_export_kmz():
    kmz = kmz_path.get()
    out_folder = kmz_out_folder.get()
    gdb = kmz_target_gdb.get().strip()
    script = get_script_path("export_kmz.pyc")

    if not os.path.isfile(kmz):
        messagebox.showerror("Error", "File KMZ tidak valid.")
        return
    if not os.path.isdir(out_folder):
        messagebox.showerror("Error", "Folder output tidak valid.")
        return
    if not os.path.isfile(script):
        messagebox.showerror("Script Tidak Ditemukan", "export_kmz.pyc tidak ditemukan.")
        return

    cmd = ["C:/Python27/ArcGIS10.8/python.exe", script, kmz, out_folder]
    if gdb:
        cmd.append(gdb)

    ret = subprocess.call(cmd)
    if ret == 0:
        messagebox.showinfo("Selesai", "Ekspor KMZ berhasil.")
    else:
        messagebox.showerror("Gagal", "Skrip berakhir dengan kode: {}".format(ret))

def run_cleanup_temp():
    script = get_script_path("cleanup_temp.pyc")
    if os.path.isfile(script):
        subprocess.call(["C:/Python27/ArcGIS10.8/python.exe", script])
    else:
        messagebox.showerror("Script Tidak Ditemukan", "cleanup_temp.pyc tidak tersedia.")

def open_export_log():
    log_path = os.path.join(kmz_out_folder.get(), "export_log.txt")
    if os.path.exists(log_path):
        os.startfile(log_path)
    else:
        messagebox.showwarning("Log Tidak Ditemukan", "File export_log.txt belum tersedia.")

# ========================SR

# ... kode Tkinter & form lain dipertahankan ...

def load_layers_step1():
    gdb = gdb_path.get()
    try:
        layers = update_sr.load_layers(gdb)
        combo_stltr1['values'] = layers
        combo_pelanggan['values'] = layers
        combo_tiangtr['values'] = layers
        layer_stltr_step1.set("")
        layer_pelanggan.set("")
        layer_tiangtr.set("")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def load_layers_step2():
    gdb = gdb_path.get()
    try:
        layers = update_sr.load_layers(gdb)
        combo_stltr2['values'] = ["<pilih layer STLTR untuk Penomoran Deret SR>"] + layers
        layer_stltr_step2.set("<pilih layer STLTR untuk Penomoran Deret SR>")
    except Exception as e:
        messagebox.showerror("Error", str(e))

#====== SR_GH Subprocess == cek diatas sama gak dengan isi update_sr.pyc
# ... (bagian kode GUI seperti browse_folder, load_layers, dsb TIDAK diubah) ...

def run_update_sr_step1(gdb, stltr, pelanggan, tiangtr):
    python_exe = r"C:\Python27\ArcGIS10.8\python.exe"
    script_path = os.path.join(os.path.dirname(__file__), "update_sr.pyc")
    args = [python_exe, script_path, "step1", gdb, stltr, pelanggan, tiangtr]
    # Jika ingin menunggu proses selesai dan menerima output:
    proc = subprocess.Popen(args, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    out, err = proc.communicate()
    if out:
        messagebox.showinfo("Output Step 1", out.decode("utf-8"))
    if err:
        messagebox.showerror("Error Step 1", err.decode("utf-8"))

def run_update_sr_step2(gdb, stltr, tiangtr):
    python_exe = r"C:\Python27\ArcGIS10.8\python.exe"
    script_path = os.path.join(os.path.dirname(__file__), "update_sr.pyc")
    args = [python_exe, script_path, "step2", gdb, stltr, tiangtr]
    proc = subprocess.Popen(args, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    out, err = proc.communicate()
    if out:
        messagebox.showinfo("Output Step 2", out.decode("utf-8"))
    if err:
        messagebox.showerror("Error Step 2", err.decode("utf-8"))

# ========================SR (modifikasi, TIDAK proses utama di sini, hanya panggil subprocess!)

def step1_prepare_sr():
    try:
        gdb = gdb_path.get()
        stltr = layer_stltr_step1.get()
        pelanggan = layer_pelanggan.get()
        tiangtr = layer_tiangtr.get()

        # Validasi
        if not all([os.path.isdir(gdb), stltr, pelanggan, tiangtr]):
            messagebox.showerror("Error", "Pastikan GDB dan semua layer telah dipilih.")
            return

        run_update_sr_step1(gdb, stltr, pelanggan, tiangtr)
    except Exception as e:
        messagebox.showerror("Error", "Gagal menjalankan Step 1:\n{}".format(e))

def step2_recursive_deret():
    try:
        gdb = gdb_path.get()
        stltr = layer_stltr_step2.get()
        tiangtr = layer_tiangtr.get()

        if not all([os.path.isdir(gdb), stltr, tiangtr]):
            messagebox.showerror("Error", "Pastikan GDB, layer STLTR dan TIANGTR telah dipilih.")
            return

        run_update_sr_step2(gdb, stltr, tiangtr)
    except Exception as e:
        messagebox.showerror("Error", "Gagal menjalankan Step 2:\n{}".format(e))

# ========================SR_end
# Tab: Export Excel
# ========================
tab_excel = ttk.Frame(notebook, borderwidth=1, relief="solid")
notebook.add(tab_excel, text="Export Excel")

excel_path = StringVar()
excel_out_folder = StringVar()

ttk.Label(tab_excel, text="File Excel:").grid(row=0, column=0, sticky="w", padx=5, pady=4)
ttk.Entry(tab_excel, textvariable=excel_path, width=55).grid(row=0, column=1, padx=5, pady=4)
ttk.Button(tab_excel, text="Browse", width=8, command=lambda: browse_file(excel_path, "*.xlsx", "Excel")).grid(row=0, column=2, padx=5)

ttk.Label(tab_excel, text="Folder Output:").grid(row=1, column=0, sticky="w", padx=5, pady=4)
ttk.Entry(tab_excel, textvariable=excel_out_folder, width=55).grid(row=1, column=1, padx=5, pady=4)
ttk.Button(tab_excel, text="Browse", width=8, command=lambda: browse_folder(excel_out_folder)).grid(row=1, column=2, padx=5)

ttk.Button(tab_excel, text="Jalankan Ekspor", width=20, command=run_export_excel).grid(row=2, column=1, pady=(10,4))

# ========================
# Tab: Export KMZ
# ========================
tab_kmz = ttk.Frame(notebook, borderwidth=1, relief="solid")
notebook.add(tab_kmz, text="Export KMZ")

kmz_path = StringVar()
kmz_out_folder = StringVar()
kmz_target_gdb = StringVar()

ttk.Label(tab_kmz, text="File KMZ:").grid(row=0, column=0, sticky="w", padx=5, pady=4)
ttk.Entry(tab_kmz, textvariable=kmz_path, width=51).grid(row=0, column=1, padx=5, pady=4)
ttk.Button(tab_kmz, text="Browse", width=8, command=lambda: browse_file(kmz_path, "*.kmz", "KMZ")).grid(row=0, column=2, padx=5)

ttk.Label(tab_kmz, text="Folder Output:").grid(row=1, column=0, sticky="w", padx=5, pady=4)
ttk.Entry(tab_kmz, textvariable=kmz_out_folder, width=51).grid(row=1, column=1, padx=5, pady=4)
ttk.Button(tab_kmz, text="Browse", width=8, command=lambda: browse_folder(kmz_out_folder)).grid(row=1, column=2, padx=5)

ttk.Label(tab_kmz, text="Target GDB:").grid(row=2, column=0, sticky="w", padx=5, pady=4)
ttk.Entry(tab_kmz, textvariable=kmz_target_gdb, width=51).grid(row=2, column=1, padx=5, pady=4)
#ttk.Button(tab_kmz, text="Browse", width=8, command=lambda: browse_file(kmz_target_gdb, "*.gdb", "Geodatabase")).grid(row=2, column=2, padx=5)
ttk.Button(tab_kmz, text="Browse", width=8, command=lambda: browse_folder(kmz_target_gdb)).grid(row=2, column=2, padx=5)

ttk.Button(tab_kmz, text="Jalankan Ekspor", width=20, command=run_export_kmz).grid(row=3, column=1, pady=(10,4))
ttk.Button(tab_kmz, text="Lihat Log", width=12, command=open_export_log).grid(row=4, column=0, padx=5, pady=(4,6))
ttk.Button(tab_kmz, text="Clear Temp", width=15, command=run_cleanup_temp).grid(row=4, column=2, padx=5, pady=(4,6))

# ========================
# TAB: Update SR with GDB-driven dropdown
# ========================

tab_update_sr = ttk.Frame(notebook, borderwidth=1, relief="solid")
notebook.add(tab_update_sr, text="Update SR")

layer_stltr_step1 = StringVar()
layer_pelanggan = StringVar()
layer_tiangtr = StringVar()
layer_stltr_step2 = StringVar()
gdb_path = StringVar()

ttk.Label(tab_update_sr, text="Folder GDB:").grid(row=0, column=0, sticky="w", padx=5, pady=4)
ttk.Entry(tab_update_sr, textvariable=gdb_path, width=53).grid(row=0, column=1, padx=5, pady=4)
ttk.Button(tab_update_sr, text="Browse", width=8, command=lambda: browse_folder(gdb_path)).grid(row=0, column=2)

# Garis pembatas visual antar proses
ttk.Separator(tab_update_sr, orient="horizontal").grid(row=5, column=0, columnspan=3, sticky="ew", pady=(10,6))

# === Step 1: Input
ttk.Label(tab_update_sr, text="STLTR :").grid(row=1, column=0, sticky="w", padx=5, pady=4)
combo_stltr1 = ttk.Combobox(tab_update_sr, textvariable=layer_stltr_step1, width=50)
combo_stltr1.grid(row=1, column=1, padx=5)
ttk.Button(tab_update_sr, text="Load Layers", command=load_layers_step1).grid(row=1, column=2)

ttk.Label(tab_update_sr, text="PELANGGAN :").grid(row=2, column=0, sticky="w", padx=5, pady=4)
combo_pelanggan = ttk.Combobox(tab_update_sr, textvariable=layer_pelanggan, width=50)
combo_pelanggan.grid(row=2, column=1, padx=5)

ttk.Label(tab_update_sr, text="TIANGTR :").grid(row=3, column=0, sticky="w", padx=5, pady=4)
combo_tiangtr = ttk.Combobox(tab_update_sr, textvariable=layer_tiangtr, width=50)
combo_tiangtr.grid(row=3, column=1, padx=5)

ttk.Button(tab_update_sr, text="Step 1: SplitLines + Delete Identical + Snap + AddField", command=step1_prepare_sr).grid(row=4, column=1, pady=(10,5))
#ttk.Button(tab_update_sr, text="Step 1: SplitLines + Delete Identical + Snap + AddField", command=step1_prepare_sr_handler).grid(row=4, column=1, pady=(10,5))

# === Step 2: Recursive
ttk.Label(tab_update_sr, text="STLTR :").grid(row=6, column=0, sticky="w", padx=5, pady=4)
combo_stltr2 = ttk.Combobox(tab_update_sr, textvariable=layer_stltr_step2, width=50)
combo_stltr2.grid(row=6, column=1, padx=5)
ttk.Button(tab_update_sr, text="Load Layers", command=load_layers_step2).grid(row=6, column=2)

# Penjelasan tambahan sejajar tombol Step 2
ttk.Button(tab_update_sr, text="Step 2: Penomoran Deret SR", command=step2_recursive_deret).grid(row=7, column=1, pady=(4,10))
#ttk.Button(tab_update_sr, text="Step 2: Penomoran Deret SR", command=step2_recursive_deret_handler).grid(row=7, column=1, pady=(4,10))


# ========================
# Tab: Update Modular
for label in ["Update Tiang", "Update JTR", "Update Pelanggan"]:
    tab = ttk.Frame(notebook, borderwidth=2, relief="raised")
    notebook.add(tab, text=label)
    ttk.Label(tab, text="Fungsi '{}' akan segera diintegrasikan.".format(label),
              font=("Segoe UI", 9), foreground="#555").pack(pady=30)

# Tab: README (pindah ke posisi terakhir)
tab_readme = ttk.Frame(notebook, borderwidth=1, relief="raise")
notebook.add(tab_readme, text="README")

readme_text = """=== SYSTEM REQUIREMENTS ===
- Python 2.7 (ArcGIS Python 10.x)
- ArcPy dan akses ke ArcGIS Desktop
- Script: export_excel_arcmap.py, export_kmz.py, update sr, update logic

=== FUNGSI TIAP TAB ===
- Export Excel: Konversi .xlsx ke Geodatabase (GDB)
- Export KMZ: Konversi .kmz ke GDB, dengan tagging otomatis
- Update SR:
    • Step 1: SplitLines + DeleteIdentical + Snap + AddField
    • Step 2: Penomoran recursive field 'deret'

- Update Tiang/JTR/Pelanggan: Untuk update spasial & validasi (akan diaktifkan)

=== TATA CARA PENGISIAN ===
[Export Excel]
- File Excel: pilih file .xlsx yang berisi data spasial
- Folder Output: pilih folder tujuan untuk menyimpan GDB

[Export KMZ]
- File KMZ: pilih file .kmz hasil survei
- Folder Output: tempat log & folder sementara
- Target GDB: pilih GDB hasil ekspor Excel sebagai lokasi hasil akhir

[Update SR]
- Pilih file Geodatabase (GDB)
- Klik tombol Load Layers untuk menampilkan layer (feature class) dalam GDB.
- Pilih layer STLTR, PELANGGAN, dan TIANG untuk memproses beberapa fungsi, sbb:
    - SplitLines (layer yang dibutuhkan: STLTR)
    - DeleteIdentical (layer yang dibutuhkan: STLTR)
    - Snapping (layer yang dibutuhkan: STLTR, PELANGGAN, dan TIANG)
        - Toleransi jarak snapping = @1 meter
- Hasil proses diatas menghasilkan layer SR baru dengan nama suffix *_split.
- Penomoran Deret SR belum berfungsi dengan baik.
    
Catatan:
- Pastikan kesesuaian nama sheet di file excel (*.xlsx) yang akan 
  diekspor (PELANGGAN, TIANG, dan TRAFO) dan memiliki koordinat.
- Penamaan kolom koordinat yang terbaca adalah sbb:
    x_candidates = ['x', 'long', 'longitude', 'koordinatx', 'koordx', 'lon']
    y_candidates = ['y', 'lat', 'latitude', 'koordinaty', 'koordy']
  Sistem tidak akan memproses diluar nama tersebut diatas. 
- Jangan gunakan .gdb sebagai Folder Output
- KMZ tanpa tag STLTR/JTR/SR akan dilewati


=== Program created by: ===

Ridwan Kurniawan | pjalahta@gmail.com | PT PETIRINDO JAYA ABADI | @2025

=== Your feedback would be greatly appreciated.  ===
"""

canvas_readme = Canvas(tab_readme, borderwidth=0)
scrollbar_readme = ttk.Scrollbar(tab_readme, orient="vertical", command=canvas_readme.yview)
scroll_frame_readme = ttk.Frame(canvas_readme)

scroll_frame_readme.bind("<Configure>", lambda e: canvas_readme.configure(scrollregion=canvas_readme.bbox("all")))
canvas_readme.create_window((0, 0), window=scroll_frame_readme, anchor="nw")
canvas_readme.configure(yscrollcommand=scrollbar_readme.set)

canvas_readme.pack(side="left", fill="both", expand=True)
scrollbar_readme.pack(side="right", fill="y")

ttk.Label(scroll_frame_readme, text=readme_text, justify="left", font=("Courier New", 9)).pack(padx=10, pady=10, anchor="w")

# Branding Footer
ttk.Separator(root, orient="horizontal").pack(fill="x", pady=(0,4))
ttk.Label(root, text="Program created by: Ridwan Kurniawan | PT PETIRINDO JAYA ABADI | 2025",
          font=("Segoe UI", 9), foreground="#555").pack()
ttk.Label(root, text="Your feedback would be greatly appreciated.",
          font=("Segoe UI", 9), foreground="#777").pack(pady=(0,8))

root.mainloop()