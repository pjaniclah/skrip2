# -*- coding: utf-8 -*-
import arcpy
import os
import sys
import xlrd
import csv
import time

def convert_sheet_to_csv(excel_path, sheet_name, csv_file):
    book = xlrd.open_workbook(excel_path)
    try:
        sheet = book.sheet_by_name(sheet_name)
    except:
        raise Exception("Sheet '{}' tidak ditemukan di file Excel.".format(sheet_name))
    with open(csv_file, "w") as f:
        for row_idx in range(sheet.nrows):
            row = [str(sheet.cell(row_idx, col_idx).value) for col_idx in range(sheet.ncols)]
            f.write(",".join(row) + "\n")

def detect_xy_fields(csv_file):
    """
    Otomatis mendeteksi kolom X/Y dari CSV.
    Mengembalikan tuple: (nama_field_X, nama_field_Y)
    """
    # Daftar kemungkinan nama field X/Y (tanpa spasi, lowercased)
    x_candidates = ['x', 'long', 'longitude', 'koordinatx', 'koordx', 'lon']
    y_candidates = ['y', 'lat', 'latitude', 'koordinaty', 'koordy']

    with open(csv_file, 'r') as f:
        reader = csv.reader(f)
        header = next(reader)
        cols = [h.strip().lower().replace(" ", "") for h in header]

        x_idx = y_idx = None
        for i, col in enumerate(cols):
            if any(col == xc for xc in x_candidates):
                x_idx = i
            if any(col == yc for yc in y_candidates):
                y_idx = i

        if x_idx is not None and y_idx is not None:
            return header[x_idx], header[y_idx]
        else:
            raise Exception("Tidak ditemukan field X/Y (longitude/latitude) yang cocok.\nHeader: {}".format(header))

if len(sys.argv) < 3:
    print("Usage: python export_excel_arcmap.py [excel_file] [output_folder]")
    sys.exit(1)

excel_path = sys.argv[1]
output_folder = sys.argv[2]
timestamp = time.strftime("%Y%m%d_%H%M%S")
gdb_name = "export_{}.gdb".format(timestamp)
gdb_path = os.path.join(output_folder, gdb_name)
spatial_ref = arcpy.SpatialReference(4326)

# Validasi ekstensi
if not (excel_path.lower().endswith(".xlsx") or excel_path.lower().endswith(".xls")):
    print("File harus berekstensi .xls atau .xlsx")
    sys.exit(1)

if not arcpy.Exists(gdb_path):
    arcpy.CreateFileGDB_management(output_folder, gdb_name)

# Mapping sheet ke nama feature class
sheets = {
    "PELANGGAN": "PELANGGAN",
    "TIANG": "TIANGTR",
    "TRAFO": "GARDU"
}

for sheet_name, fc_name in sheets.items():
    csv_file = os.path.join(output_folder, fc_name + "_" + timestamp + ".csv")
    try:
        convert_sheet_to_csv(excel_path, sheet_name, csv_file)
        x_field, y_field = detect_xy_fields(csv_file)
        layer_name = "xy_" + fc_name
        arcpy.MakeXYEventLayer_management(csv_file, x_field, y_field, layer_name, spatial_ref)
        arcpy.CopyFeatures_management(layer_name, os.path.join(gdb_path, fc_name))
        print("✅ Berhasil ekspor sheet '{}' ke feature class '{}'".format(sheet_name, fc_name))
    except Exception as e:
        print("⚠️  Gagal ekspor sheet '{}': {}".format(sheet_name, e))

print("✅ Geodatabase berhasil dibuat:", gdb_path)