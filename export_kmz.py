# -*- coding: utf-8 -*-
import arcpy
import os
import sys
import shutil
import datetime
import time

def detect_tag(value):
    value = value.upper()
    if "STLTR" in value or "SR" in value:
        return "STLTR"
    elif "JTR" in value:
        return "JTR"
    return None

def sanitize_name(text):
    return "".join(c if c.isalnum() else "_" for c in text)[:50]

def find_gdb(folder_path):
    for root, dirs, files in os.walk(folder_path):
        for d in dirs:
            if d.lower().endswith(".gdb"):
                return os.path.join(root, d)
    return None

def log(msg, log_file):
    print msg
    with open(log_file, "a") as f:
        f.write(msg + "\n")

def safe_rmtree(path, retries=5, delay=1):
    for i in range(retries):
        try:
            shutil.rmtree(path)
            return True
        except Exception as e:
            print "⚠️ Gagal hapus folder (percobaan {}): {}".format(i+1, str(e))
            time.sleep(delay)
    return False

def main():
    args = sys.argv

    if len(args) < 3:
        print "❌ Argumen tidak lengkap.\nGunakan: python export_kmz.py <input_kmz> <output_folder> [target_gdb]"
        sys.exit(1)

    input_kmz = args[1]
    output_folder = args[2]
    target_gdb = args[3] if len(args) > 3 else None

    arcpy.env.overwriteOutput = True
    log_file = os.path.join(output_folder, "export_log.txt")
    with open(log_file, "w") as f:
        f.write("📋 Log Ekspor KMZ - {}\n\n".format(datetime.datetime.now()))

    try:
        log("📥 Input KMZ: {}".format(input_kmz), log_file)
        log("📂 Folder Output: {}".format(output_folder), log_file)

        if not input_kmz.lower().endswith(".kmz") or not os.path.isfile(input_kmz):
            log("❌ File KMZ tidak valid: {}".format(input_kmz), log_file)
            sys.exit(1)

        temp_folder = os.path.join(output_folder, "temp_kmz")
        if not os.path.exists(temp_folder):
            os.makedirs(temp_folder)

        arcpy.KMLToLayer_conversion(input_kmz, temp_folder, "kmz_output")

        layer_gdb = find_gdb(temp_folder)
        if not layer_gdb or not arcpy.Exists(layer_gdb):
            log("❌ Gagal menemukan GDB hasil konversi KMZ.", log_file)
            sys.exit(1)

        log("✅ GDB ditemukan: {}".format(layer_gdb), log_file)

        # 💾 Tentukan GDB target
        if target_gdb:
            output_gdb = target_gdb
            if not arcpy.Exists(output_gdb):
                log("❌ GDB target tidak ditemukan: {}".format(output_gdb), log_file)
                sys.exit(1)
            log("📌 Menyimpan hasil ke GDB: {}".format(output_gdb), log_file)
        else:
            output_gdb = os.path.join(output_folder, "kmz_export.gdb")
            if not arcpy.Exists(output_gdb):
                arcpy.CreateFileGDB_management(output_folder, "kmz_export.gdb")
            log("📌 GDB dibuat otomatis: {}".format(output_gdb), log_file)

        arcpy.env.workspace = layer_gdb
        datasets = arcpy.ListDatasets("", "Feature") or [""]
        total_exported = 0

        for ds in datasets:
            arcpy.env.workspace = os.path.join(layer_gdb, ds)
            fc_list = arcpy.ListFeatureClasses()
            if not fc_list:
                continue

            for fc in fc_list:
                log("🔍 Memproses feature class: {}".format(fc), log_file)
                desc = arcpy.Describe(fc)
                if desc.shapeType != "Polyline":
                    continue

                field_names = [f.name for f in arcpy.ListFields(fc)]
                if "FolderPath" not in field_names:
                    log("⚠️ Field 'FolderPath' tidak ditemukan di {}".format(fc), log_file)
                    continue

                unique_values = set()
                with arcpy.da.SearchCursor(fc, ["FolderPath"]) as cursor:
                    for row in cursor:
                        if row[0]:
                            unique_values.add(row[0])

                for val in unique_values:
                    tag = detect_tag(val)
                    if not tag:
                        continue
                    where_clause = "FolderPath = '{}'".format(val.replace("'", "''"))
                    layer_name = sanitize_name(tag + "_" + val)
                    out_fc = os.path.join(output_gdb, layer_name)

                    if arcpy.Exists(out_fc):
                        log("🛡️ Layer sudah ada, dilewati: {}".format(layer_name), log_file)
                        continue

                    arcpy.MakeFeatureLayer_management(fc, "temp_layer", where_clause)
                    arcpy.CopyFeatures_management("temp_layer", out_fc)
                    count = int(arcpy.GetCount_management(out_fc).getOutput(0))
                    log("📤 Menyimpan layer: {} ({} fitur)".format(layer_name, count), log_file)
                    arcpy.Delete_management("temp_layer")
                    total_exported += 1

        if total_exported == 0:
            log("⚠️ Tidak ada layer polyline yang berhasil diekspor.", log_file)
        else:
            log("🎉 Ekspor KMZ selesai. {} layer disimpan.".format(total_exported), log_file)

        arcpy.env.workspace = None
        arcpy.ClearWorkspaceCache_management()

        # 🧹 Bersihkan folder sementara
        if safe_rmtree(temp_folder):
            log("🧹 Folder sementara dihapus: {}".format(temp_folder), log_file)
        else:
            with open(os.path.join(output_folder, "pending_cleanup.txt"), "w") as f:
                f.write(temp_folder)
            log("⚠️ Folder gagal dihapus. Ditandai untuk pembersihan manual: {}".format(temp_folder), log_file)

        sys.exit(0)

    except Exception as e:
        import traceback
        log("❌ Error saat ekspor KMZ: {}".format(str(e)), log_file)
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()