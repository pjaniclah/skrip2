# -*- coding: utf-8 -*-
import os
import subprocess
import Tkinter as tk
#from Tkinter import tk, label, stringVar, canvas
import ttk
import tkFileDialog as filedialog
import tkMessageBox as messagebox
import tempfile
#=====

from functools import partial

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
#=========Builder untuk Subprocess Command
def build_command(step, gdb, stltr, pelanggan, tiangtr):
    script_path = os.path.join(os.path.dirname(__file__), "update_srx.pyc")
    return ["C:/Python27/ArcGIS10.8/python.exe", script_path,
            "--step", str(step), "--gdb", gdb,
            "--stltr", stltr, "--pelanggan", pelanggan,
            "--tiangtr", tiangtr]

#=========Fungsi Handler: Step 1
def run_update_sr_step1():
    gdb = gdb_path.get()
    stltr = layer_stltr_step1.get()
    pelanggan = layer_pelanggan.get()
    tiangtr = layer_tiangtr.get()

    if not all([gdb, stltr, pelanggan, tiangtr]):
        messagebox.showwarning("Input Belum Lengkap", "Pastikan semua layer dan folder GDB sudah dipilih.")
        return

    try:
        cmd = build_command(1, gdb, stltr, pelanggan, tiangtr)
        output = subprocess.check_output(cmd).decode("utf-8")
        for line in output.splitlines():
            if line.startswith("OUTPUT="):
                layer_stltr_step2.set(line.split("=")[1])
        messagebox.showinfo("Step 1 Berhasil", output)
    except Exception as e:
        messagebox.showerror("Step 1 Gagal", str(e))

#=========
def select_gdb():
    folder = tkFileDialog.askdirectory()
    if folder:
        gdb_path.set(folder)

def select_layer(target_var):
    file = tkFileDialog.askopenfilename()
    if file:
        target_var.set(file)
#=========Browse GDB
def browse_folder(var_gdb):
    path = filedialog.askdirectory(title="Pilih folder GDB")
    if path and os.path.isdir(path): var_gdb.set(path)
    else: messagebox.showerror("Error", "Folder GDB tidak valid.")
#========
def safe_call(cmd):
    try:
        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        stdout, stderr = proc.communicate()
        if proc.returncode != 0:
            raise Exception(stderr)
        return stdout.decode("utf-8").splitlines()
    except Exception as e:
        raise Exception("Subprocess gagal: {}".format(e))

 #=========Load Layers Step 1
def load_layers_step1():
    gdb = gdb_path.get()
    cmd = ["C:/Python27/ArcGIS10.8/python.exe", "update_srx.pyc", "--step", "list", "--gdb", gdb]
    try:
        layers = safe_call(cmd)
        for combo in [combo_stltr1, combo_pelanggan, combo_tiangtr, combo_stltr2]:
            combo["values"] = layers
    except Exception as e:
        messagebox.showerror("Error Load Layers Step 1", str(e))
#========
"""
def load_layers_step1():
    gdb = gdb_path.get()
    try:
        cmd = ["C:/Python27/ArcGIS10.8/python.exe", os.path.join(os.path.dirname(__file__), "update_srx.pyc"),
               "--step", "list", "--gdb", gdb]
        output = subprocess.check_output(cmd)
        layers = output.strip().splitlines()
        combo_stltr1["values"] = layers
        combo_pelanggan["values"] = layers
        combo_tiangtr["values"] = layers
        combo_stltr2["values"] = layers
    except Exception as e:
        messagebox.showerror("Error Load Layers Step 1", str(e))
"""
#=========Load Layers Step 2
def load_layers_step2():
    gdb = gdb_path.get()
    try:
        cmd = ["C:/Python27/ArcGIS10.8/python.exe", os.path.join(os.path.dirname(__file__), "update_srx.pyc"),
               "--step", "list", "--gdb", gdb]
        output = subprocess.check_output(cmd)
        layers = output.strip().splitlines()
        combo_stltr2["values"] = layers
    except Exception as e:
        messagebox.showerror("Error Load Layers Step 2", str(e))
# ================Fungsi Step 2a: Intersect Update
def run_intersect_update():
    gdb = gdb_path.get()
    stltr = layer_stltr_step2.get()
    tiangtr = layer_tiangtr.get()

    if not all([gdb, stltr, tiangtr]):
        messagebox.showwarning("Input Belum Lengkap", "Pastikan GDB, STLTR dan TIANGTR telah dipilih.")
        return

    try:
        cmd = ["C:/Python27/ArcGIS10.8/python.exe", os.path.join(os.path.dirname(__file__), "update_srx.pyc"),
               "--step", "2a", "--gdb", gdb, "--stltr", stltr, "--tiangtr", tiangtr]
        output = subprocess.check_output(cmd).decode("utf-8")
        messagebox.showinfo("Intersect Update Berhasil", output)
    except Exception as e:
        messagebox.showerror("Intersect Gagal", str(e))
# ================Fungsi Step 2b: Penomoran
def run_recursive_deret():
    import subprocess
    gdb = gdb_path.get()
    stltr_layer = layer_stltr_step2.get()

    if not gdb or not stltr_layer:
        messagebox.showwarning("Input kosong", "Pastikan GDB dan Layer STLTR sudah dipilih.")
        return

    try:
        cmd = [
            r"C:\Python27\ArcGIS10.8\python.exe",
            os.path.join(os.path.dirname(__file__), "update_srx.pyc"),
           "--step", "2b",
           "--gdb", gdb,
            "--stltr", stltr_layer
        ]
        output = subprocess.check_output(cmd).decode("utf-8")
        messagebox.showinfo("Deret SR berhasil diupdate.", output)

#    try:
#       hasil = subprocess.check_output(cmd).decode("utf-8")
#        text_log.delete("1.0", tk.END)
#        text_log.insert(tk.END, hasil)
#        messagebox.showinfo("Selesai", "Deret SR berhasil diupdate.")
    except Exception as e:
        messagebox.showerror("Update Deret Gagal", str(e))

# ========================
# GUI Setup
# ========================
root = tk.Tk()
root.title("GIS Exporter & Updater v2025.RC3")
root.geometry("530x320")
notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True, padx=8, pady=8)
# ========================
# Tab: Export Excel
# ========================
tab_excel = ttk.Frame(notebook, borderwidth=1, relief="solid")
notebook.add(tab_excel, text="Export Excel")

excel_path = tk.StringVar()
excel_out_folder = tk.StringVar()

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

kmz_path = tk.StringVar()
kmz_out_folder = tk.StringVar()
kmz_target_gdb = tk.StringVar()

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

# ========================Dibawah ini SR
tab_update_sr = ttk.Frame(notebook, borderwidth=1, relief="solid")
notebook.add(tab_update_sr, text="Update SR")
#=========Variabel Global Tkinter
gdb_path = tk.StringVar()
layer_stltr_step1 = tk.StringVar()
layer_pelanggan = tk.StringVar()
layer_tiangtr = tk.StringVar()
layer_stltr_step2 = tk.StringVar()
#=======
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

ttk.Button(tab_update_sr, text="Step 1: SplitLines + Delete Identical + Snap + AddField", command=run_update_sr_step1).grid(row=4, column=1, pady=(10,5))

# === Step 2: Recursive
ttk.Label(tab_update_sr, text="STLTR :").grid(row=6, column=0, sticky="w", padx=5, pady=4)
combo_stltr2 = ttk.Combobox(tab_update_sr, textvariable=layer_stltr_step2, width=50)
combo_stltr2.grid(row=6, column=1, padx=5)

# Penjelasan tambahan sejajar tombol Step 2
ttk.Button(tab_update_sr, text="Step 2a: Intersect Update", command=run_intersect_update).grid(row=7, column=1, pady=(4,10), padx=4)
ttk.Button(tab_update_sr, text="Step 2b: Penomoran Deret SR", command=run_recursive_deret).grid(row=8, column=1, pady=(4,10))

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
- Python 2.7 (ArcGIS Python 10.8)
- ArcPy dan akses ke ArcGIS Desktop

=== FUNGSI TIAP TAB ===
- Export Excel:
    • Konversi .xlsx ke Geodatabase (GDB)
- Export KMZ
    • Konversi .kmz ke GDB, dengan tagging otomatis
    • Extract Path Photo ke *.xlsx
- Update SR:
    • Step 1: SplitLines + DeleteIdentical + Snap + AddField
    • Step 2a: Intersect Update (merupakan initialisasi proses deret seret)
    • Step 2b: Penomoran Deret SR
    • Next: Panjang SR, IDPELANGGAN, ADMIN    
- Update TIANG
    • Otomasi NO_TIANG
    • Kontruksi
    • Update KODE_TIANG_TR, Hantaran, Description
- Update JTR
    • NO_TIANG_AWAL & TIANG_TR_AKHIR
    • description
    • panjang_hantaran
- Update Pelanggan
    • KODE_TIANG_TR
    • PANJANG_SR
    • DERET SR

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
    • SplitLines (layer yang dibutuhkan: STLTR)
    • DeleteIdentical (layer yang dibutuhkan: STLTR)
    • Snapping (layer yang dibutuhkan: STLTR, PELANGGAN, dan TIANG)
        - Toleransi jarak snapping = @1 meter
    • Add fields = "deret" (untuk penomoran deret sr)        
- Penamaan output layer SR baru dengan nama suffix *_split.
- Proses Penomoran Deret SR:
    - Pilih layer SR
    - Klik tombol "Intersect Update"
        Ini proses penomoran deret kesatu (1) untuk garis SR yang menempel
        ke layer TIANGTR
    - Klik tombol "Penomoran Deret SR"
        Ini proses iterasi untuk penomoran deret sr.
- Untuk update spasial & validasi lainnya belum diaktifkan

Catatan:
- Pastikan kesesuaian nama sheet di file excel (*.xlsx) yang akan 
  diekspor (PELANGGAN, TIANG, dan TRAFO) dan memiliki koordinat.
- Penamaan kolom koordinat yang terbaca adalah sbb:
    x_candidates = ['x', 'long', 'longitude', 'koordinatx', 'koordx', 'lon']
    y_candidates = ['y', 'lat', 'latitude', 'koordinaty', 'koordy']
  Sistem tidak akan memproses diluar nama tersebut diatas. 
- Jangan gunakan .gdb sebagai Folder Output
- KMZ tanpa tag STLTR/JTR/SR di Folder Path akan dilewati


=== Program created by: ===

Ridwan Kurniawan | pjalahta@gmail.com | PT PETIRINDO JAYA ABADI | @2025

=== Your feedback would be greatly appreciated.  ===
"""

tk.canvas_readme = tk.Canvas(tab_readme, borderwidth=0)
scrollbar_readme = ttk.Scrollbar(tab_readme, orient="vertical", command=tk.canvas_readme.yview)
scroll_frame_readme = ttk.Frame(tk.canvas_readme)

scroll_frame_readme.bind("<Configure>", lambda e: tk.canvas_readme.configure(scrollregion=tk.canvas_readme.bbox("all")))
tk.canvas_readme.create_window((0, 0), window=scroll_frame_readme, anchor="nw")
tk.canvas_readme.configure(yscrollcommand=scrollbar_readme.set)

tk.canvas_readme.pack(side="left", fill="both", expand=True)
scrollbar_readme.pack(side="right", fill="y")

ttk.Label(scroll_frame_readme, text=readme_text, justify="left", font=("Courier New", 9)).pack(padx=10, pady=10,
                                                                                               anchor="w")

# Branding Footer
ttk.Separator(root, orient="horizontal").pack(fill="x", pady=(0, 4))
ttk.Label(root, text="Program created by: Ridwan Kurniawan | PT PETIRINDO JAYA ABADI | 2025",
          font=("Segoe UI", 9), foreground="#555").pack()
ttk.Label(root, text="Your feedback would be greatly appreciated.",
          font=("Segoe UI", 9), foreground="#777").pack(pady=(0, 8))

root.mainloop()

