# update_srx.py
# -*- coding: utf-8 -*-
import arcpy, os, time
import traceback
import sys

def log(gdb, msg):
    with open(os.path.join(gdb, "update_srx_log.txt"), "a") as f:
        f.write("[%s] %s\n" % (time.strftime("%H:%M:%S"), msg))

def list_layers(gdb):
    arcpy.env.workspace = gdb
    for fc in arcpy.ListFeatureClasses():
        print(fc)

def step1_process(gdb, stltr, pelanggan, tiangtr):
    """Step 1: Spilit → DeleteIdentical → Snap → AddFields"""
    arcpy.env.workspace = gdb
    ts = time.strftime("%Y%m%d_%H%M")
    out_name = "STLTR_%s_split" % ts
    out_fc = os.path.join(gdb, out_name)

    log(gdb, "SplitLine...")
    arcpy.management.SplitLine(os.path.join(gdb, stltr), out_fc)

    log(gdb, "DeleteIdentical...")
    arcpy.management.DeleteIdentical(out_fc, ["SHAPE_Length"])

    log(gdb, "Snap...")
    arcpy.edit.Snap(out_fc, [
        [os.path.join(gdb, pelanggan), "VERTEX", "1 Meters"],
        [os.path.join(gdb, tiangtr), "VERTEX", "1 Meters"]
    ])

    log(gdb, "AddField deret...")
    arcpy.management.AddField(out_fc, "deret", "LONG", field_alias="DERET STLTR")
    print("OUTPUT=%s" % out_name)

# ==========
def intersect_update_deret(gdb, layer, tiangtr):
    #Step 2a: Intersect STLTR dengan TIANGTR → deret = 1
    lyr = arcpy.MakeFeatureLayer_management(os.path.join(gdb, layer), "temp_lyr")
    src = os.path.join(gdb, tiangtr)

    arcpy.SelectLayerByLocation_management(lyr, "INTERSECT", src, "", "NEW_SELECTION")
    count_selected = int(arcpy.GetCount_management(lyr).getOutput(0))
    if count_selected == 0:
        raise RuntimeError("⚠ Tidak ada fitur STLTR yang berintersect dengan TIANGTR.")

    arcpy.CalculateField_management(lyr, "deret", 1, "PYTHON", "")
    arcpy.SelectLayerByAttribute_management(lyr, "CLEAR_SELECTION")

#    print("Step 2a selesai: Layer %s diberi deret = 1 berdasarkan TIANGTR" % layer)
    print("Step 2a selesai: {} fitur diberi deret = 1".format(count_selected))

# ==========
def recursive_deret(gdb, layer, fieldname="deret"):
    #Step 2b: Penomoran rekursif berdasarkan nilai deret pada layer STLTR hasil Step 2a

    #The input feature class.  The user must have write permissions to these data.
    arcpy.env.workspace = gdb
    in_fc = os.path.join(gdb, layer)
#    in_fc = r'D:\03 LAHAT\01 GDB\xARSIP\SLR.shp'
    #The existing attribute (type must be integer)
    in_field = fieldname

    try:

        counter = 1

        #in_fc1 = arcpy.MakeFeatureLayer_management(in_fc)
        in_fc1 = arcpy.MakeFeatureLayer_management(in_fc)

        results = arcpy.management.GetCount(in_fc1).getOutput(0)
        arcpy.AddMessage('Total objects to attribute =  ' + str(results))

        def recursive_select(in_fc1, counter, in_field):
        
            arcpy.AddMessage('working, please wait...')
        
            #Make the SQL statement
            query = '\"'+ in_field + '\"' +  ' = ' + str(counter)
        
        #Select by attribute that record with the correct counter value
            arcpy.management.SelectLayerByAttribute(in_fc1, 'NEW_SELECTION', query)
        
        #Once all the records are attributed there will be no more features to
        #select.  If so, end this function by returning nothing...
            if arcpy.management.GetCount(in_fc1).getOutput(0) == '0':
                arcpy.AddMessage("All records are calculated...")
                return
        
        #Select the geometries that touch the selected geometry
            arcpy.SelectLayerByLocation_management(in_fc1, "INTERSECT", in_fc1, "", "NEW_SELECTION")

        #Reselect those records with a value of 0
            query0 = '\"{0}\" IS NULL OR \"{0}\" = 0'.format(in_field)
            arcpy.management.SelectLayerByAttribute(in_fc1, 'SUBSET_SELECTION', query0)
        
        #Calculate the fields to those adjacent objects
            arcpy.CalculateField_management(in_fc1, in_field, counter + 1)
        #Clear the selection...
            arcpy.management.SelectLayerByAttribute(in_fc1, 'CLEAR_SELECTION')
        
        #Up the count by one
            counter +=1
        #Make recursive function call...
            recursive_select(in_fc1,  counter, in_field)
    
        recursive_select(in_fc1,  counter, in_field)
    
   
        arcpy.AddMessage("done")
    
    
    except arcpy.ExecuteError:
        msgs = arcpy.GetMessages(2)
        arcpy.AddError(msgs)
        arcpy.AddMessage(msgs)
    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback info:\n" + tbinfo + "\nError Info:\n" + str(sys.exc_info()[1])
        msgs = "ArcPy ERRORS:\n" + arcpy.GetMessages(2) + "\n"
        arcpy.AddError(pymsg)
        arcpy.AddError(msgs)
        arcpy.AddMessage(pymsg + "\n")
        arcpy.AddMessage(msgs)

#============
if __name__ == "__main__":
    import argparse
    p = argparse.ArgumentParser()
    p.add_argument("--step", choices=["list", "1", "2a","2b"], required=True)
    p.add_argument("--gdb", required=True)
    p.add_argument("--mode", choices=["deret"], help="Eksekusi langsung mode tertentu")
    p.add_argument("--stltr_fc", help="Full path ke feature class STLTR (mode deret)")
    p.add_argument("--sr_fc", help="Full path ke feature class seed SR (mode deret)")
    p.add_argument("--debug", action="store_true", help="Aktifkan debug mode (opsional)")
    p.add_argument("--stltr")
    p.add_argument("--pelanggan")
    p.add_argument("--tiangtr")
    args = p.parse_args()

    if args.step == "list":
        list_layers(args.gdb)
    elif args.step == "1":
        step1_process(args.gdb, args.stltr, args.pelanggan, args.tiangtr)
    elif args.step == "2a":
        intersect_update_deret(args.gdb, args.stltr, args.tiangtr)
    elif args.step == "2b":
        recursive_deret(args.gdb, args.stltr)
"""
    if args.mode == "deret":
        if not args.stltr_fc or not args.sr_fc:
            raise ValueError("--stltr_fc dan --sr_fc wajib untuk mode deret")
        recursive_deret(args.stltr_fc, args.sr_fc, debug=args.debug)
"""