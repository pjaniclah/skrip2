# update_sr.py
import os
import time
import arcpy

def log(gdb_path, text):
    log_path = os.path.join(gdb_path, "update_sr_log.txt")
    with open(log_path, "a") as f:
        f.write("[{}] {}\n".format(time.strftime("%H:%M:%S"), text))

def load_layers(gdb_path):
    """Return a list of feature classes in the given GDB folder."""
    if not os.path.isdir(gdb_path):
        raise ValueError("Path GDB tidak valid.")
    arcpy.env.workspace = gdb_path
    return arcpy.ListFeatureClasses()

def step1_prepare_sr(gdb_path, stltr_layer, pelanggan_layer, tiangtr_layer):
    """Step 1: SplitLine, DeleteIdentical, Snap, AddField 'deret'."""
    arcpy.env.workspace = gdb_path
    stltr_fc = os.path.join(gdb_path, stltr_layer)
    pelanggan_fc = os.path.join(gdb_path, pelanggan_layer)
    tiangtr_fc = os.path.join(gdb_path, tiangtr_layer)
    if not all(map(arcpy.Exists, [stltr_fc, pelanggan_fc, tiangtr_fc])):
        raise RuntimeError("Salah satu layer tidak ditemukan.")

    timestamp = time.strftime("%Y%m%d_%H%M")
    sr_name = "STLTR_{}_split".format(timestamp)
    sr_output = os.path.join(gdb_path, sr_name)

    log(gdb_path, "SplitLine...")
    arcpy.management.SplitLine(stltr_fc, sr_output)

    log(gdb_path, "DeleteIdentical by SHAPE_Length...")
    arcpy.management.DeleteIdentical(sr_output, ["SHAPE_Length"])

    log(gdb_path, "Snap ke PELANGGAN dan TIANGTR...")
    arcpy.edit.Snap(sr_output, [[pelanggan_fc, "VERTEX", "1 Meters"],
                                [tiangtr_fc, "VERTEX", "1 Meters"]])

    if "deret" not in [f.name.lower() for f in arcpy.ListFields(sr_output)]:
        arcpy.management.AddField(sr_output, "deret", "LONG", field_alias="DERET SR")
        log(gdb_path, "Menambahkan field 'deret'")

    log(gdb_path, "Step 1 selesai. Output: {}".format(sr_output))
    return sr_output

def step2_recursive_deret(gdb_path, stltr_layer, tiangtr_layer):
    """Step 2: Penomoran field 'deret' secara rekursif."""
    arcpy.env.workspace = gdb_path
    stltr_fc = os.path.join(gdb_path, stltr_layer)
    tiangtr_fc = os.path.join(gdb_path, tiangtr_layer)

    if not all(map(arcpy.Exists, [stltr_fc, tiangtr_fc])):
        raise RuntimeError("Layer STLTR atau TIANGTR tidak ditemukan.")

    log(gdb_path, "Mulai recursive DERET...")
    lyr = arcpy.management.MakeFeatureLayer(stltr_fc)
    count = 1

    arcpy.management.SelectLayerByLocation(lyr, "INTERSECT", tiangtr_fc, "", "NEW_SELECTION")
    arcpy.management.CalculateField(lyr, "deret", count, "PYTHON")
    arcpy.management.SelectLayerByAttribute(lyr, 'CLEAR_SELECTION')

    def recursive(layer, current, fieldname):
        q = '"{}" = {}'.format(fieldname, current)
        arcpy.management.SelectLayerByAttribute(layer, "NEW_SELECTION", q)
        if int(arcpy.management.GetCount(layer).getOutput(0)) == 0:
            return
        arcpy.management.SelectLayerByLocation(layer, "INTERSECT", layer, "", "NEW_SELECTION")
        arcpy.management.SelectLayerByAttribute(layer, "SUBSET_SELECTION", '"{}" = 0'.format(fieldname))
        arcpy.management.CalculateField(layer, fieldname, current + 1, "PYTHON")
        arcpy.management.SelectLayerByAttribute(layer, "CLEAR_SELECTION")
        recursive(layer, current + 1, fieldname)

    recursive(lyr, count, "deret")
    log(gdb_path, "DERET selesai.")
    
if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Update SR Tools (Step 1/2)")
    parser.add_argument("--gdb", required=True, help="Path ke folder GDB")
    parser.add_argument("--step", required=True, choices=["1", "2"], help="Step 1 (prepare) atau 2 (recursive deret)")
    parser.add_argument("--stltr", required=True, help="Nama layer STLTR")
    parser.add_argument("--pelanggan", help="Nama layer pelanggan (hanya Step 1)")
    parser.add_argument("--tiangtr", required=True, help="Nama layer tiangtr")
    args = parser.parse_args()

    if args.step == "1":
        if not args.pelanggan:
            parser.error("--pelanggan harus diisi untuk Step 1")
        output = step1_prepare_sr(args.gdb, args.stltr, args.pelanggan, args.tiangtr)
        print("Output Step 1:", output)
    elif args.step == "2":
        step2_recursive_deret(args.gdb, args.stltr, args.tiangtr)
        print("Step 2 selesai")