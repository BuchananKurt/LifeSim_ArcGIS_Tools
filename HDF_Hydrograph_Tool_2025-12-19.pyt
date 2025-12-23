import arcpy
import h5py
from datetime import datetime, timedelta
import openpyxl
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.chart import ScatterChart, Series
from openpyxl.chart.axis import NumericAxis
from openpyxl.drawing.line import LineProperties
from openpyxl.styles import Font, Alignment
import numpy as np
import os
import re
import math
import textwrap

class Toolbox(object):
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the
        .pyt file)."""
        self.label = "Breach and Nonbreach Hydrographs"
        self.alias = "Breach and Nonbreach Hydrographs from HDF"
        self.description = "ArcGIS python toolbox created to support the USACE MMC Consequences Team."

        # List of tool classes associated with this toolbox
        self.tools = [Hdfhydro, Hdfplannames]

##-------------------------------------------------------------------------------------

class Hdfhydro(object):

    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Breach and Nonbreach Hydrographs"
        self.description = "Does lots of stuff."
        self.canRunInBackground = True

    def getParameterInfo(self):
        """Define parameter definitions"""
        inputpoints = arcpy.Parameter(
            displayName="Input Points Shapefile",
            name="Input Points Shapefile",
            datatype="DEFeatureClass",
            parameterType="Required",
            direction="Input")
        #inputpoints.value = fr'C:\~Kurt\~Projects\Buford_Dam_IES\Shapefiles\Buf_HydroPoints1.shp'
        #inputpoints.value = fr'C:\~Kurt\~Projects\Jennings_Randolph_IES\Shapefiles\Hydro_Points1.shp'
        
        inputpointsname = arcpy.Parameter(
            displayName="Input Points Name Field",
            name="Input Points Shapefile Name Field",
            datatype="GPString",
            parameterType="Optional",
            direction="Input")
        inputpointsname.filter.type = "ValueList"
        inputpointsname.filter.list = []  # Will be populated in updateParameters
  
        inputbreachhdf = arcpy.Parameter(
            displayName="Input Breach HDF File",
            name="Input Breach HDF File",
            datatype="DEFile",
            parameterType="Required",
            direction="Input")
        inputbreachhdf.filter.list = ["hdf"]   # Only show .hdf files
        #inputbreachhdf.value = fr'C:\~Kurt\~Projects\Buford_Dam_IES\2025_RAS\RAS_2025-10\Buford_GA00824.p04.hdf'
        #inputbreachhdf.value = fr'C:\~Kurt\~Projects\Jennings_Randolph_IES\Jennings_RAS_2025\JenningsRandolph_MD.p38.hdf'
        
        breachhdfname = arcpy.Parameter(
            displayName="(info only) Breach ShortID",
            name="Breach HDF Name",
            datatype="GPString",
            parameterType="Optional", 
            direction="Output")
        
        inputnonbreachhdf = arcpy.Parameter(
            displayName="Input NonBreach HDF File",
            name="Input NonBreach HDF File",
            datatype="DEFile",
            parameterType="Required",
            direction="Input")
        inputnonbreachhdf.filter.list = ["hdf"]   # Only show .hdf files
        #inputnonbreachhdf.value = fr'C:\~Kurt\~Projects\Buford_Dam_IES\2025_RAS\RAS_2025-10\Buford_GA00824.p77.hdf'
        #inputnonbreachhdf.value = fr'C:\~Kurt\~Projects\Jennings_Randolph_IES\Jennings_RAS_2025\JenningsRandolph_MD.p36.hdf'
        
        nonbreachhdfname = arcpy.Parameter(
            displayName="(info only) Non-Breach ShortID",
            name="NonBreach HDF Name",
            datatype="GPString",
            parameterType="Optional", 
            direction="Output")
        outputfolder = arcpy.Parameter(
            displayName="Output Folder",
            name="Output Folder",
            datatype="DEFolder",
            parameterType="Required",
            direction="Input")
        #outputfolder.value = fr'C:\~Kurt\~Projects\Buford_Dam_IES\Shapefiles\HydroTest1'
        #outputfolder.value = fr'C:\~Kurt\~Projects\Jennings_Randolph_IES\Results\HDF_Hydrograph1'
        
        hazardtime = arcpy.Parameter(
            displayName="Hazard Time in Hours from Start (optional)",
            name="Hazard Time in Hours from Start",
            datatype="GPString",
            parameterType="Optional", 
            direction="Input")
        #hazardtime.value = "42.75"
        
        deletepoints = arcpy.Parameter(
            displayName="Check to delete interim join files",
            name="Check to delete interim join files",
            datatype="GPBoolean",
            parameterType="Optional",
            direction="Input")
        deletepoints.value = True

        adddatapoints = arcpy.Parameter(
            displayName="Check to add times to input points",
            name="Check to add times",
            datatype="GPBoolean",
            parameterType="Optional",
            direction="Input")
        adddatapoints.value = False
        
        parameters = [inputpoints, inputpointsname, inputbreachhdf, breachhdfname, inputnonbreachhdf, nonbreachhdfname, outputfolder, hazardtime, deletepoints, adddatapoints]
        #                0                  1              2             3               4                 5                  6             7         8
        return parameters

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        #Lookups for Data list selections
        inputpoints = parameters[0].valueAsText
        inputpointsname_param = parameters[1]  # Adjust index if needed
        if inputpoints:
            # List all fields in the input study area
            field_names = [f.name for f in arcpy.ListFields(inputpoints)]
            inputpointsname_param.filter.list = field_names

        ## Lookup HDF plan names
        if parameters[2].value:
            brpath = parameters[2].valueAsText
            hdf_breach = h5py.File(brpath, 'r')
            hdf_brplaninfotable = hdf_breach.get(f'/Plan Data/Plan Information')
            plan_brshortid_value = hdf_brplaninfotable.attrs.get('Plan ShortID')
            plan_brshortid_value = plan_brshortid_value.decode('ascii', errors='ignore').rstrip('\x00').strip()
            parameters[3].value = plan_brshortid_value

         ## Lookup HDF plan names
        if parameters[4].value:
            nbbrpath = parameters[4].valueAsText
            hdf_nonbreach = h5py.File(nbbrpath, 'r')
            hdf_nbrplaninfotable = hdf_nonbreach.get(f'/Plan Data/Plan Information')
            plan_nbrshortid_value = hdf_nbrplaninfotable.attrs.get('Plan ShortID')
            plan_nbrshortid_value = plan_nbrshortid_value.decode('ascii', errors='ignore').rstrip('\x00').strip()
            parameters[5].value = plan_nbrshortid_value

        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        #parameters[2].setWarningMessage(plan_brshortid_value)
        return
       
    def execute(self, parameters, messages):
        """The source code of the tool."""
        #call parameters that were entered and create new variables as needed
        input_point_shapefile = parameters[0].valueAsText
        input_point_name_field = parameters[1].valueAsText
        breach_file_path = parameters[2].valueAsText
        nonbreach_file_path = parameters[4].valueAsText
        output_folder = parameters[6].valueAsText
        hazardtime1 = parameters[7].valueAsText
        #boolean parameters should just be the value, not text, like below
        deletepoints = parameters[8].value
        add_data_to_points = parameters[9].value
        
        if hazardtime1 is not None:
            hazard_time_exists = True
        else:
            hazard_time_exists = False

        ### Input Files
        #input_point_shapefile = fr'C:\~Kurt\~Projects\Arkabutla\RASnew\Calculated Layers\test_points1.shp' # points with a Name field, like CIKR
        #input_point_shapefile = fr'C:\~Kurt\~Projects\Buford_Dam_IES\Shapefiles\Buf_HydroPoints1.shp'
        #input_point_name_field = 'Name'

        #breach_file_path = fr'C:\~Kurt\~Projects\Cherry_Creek_Dam_2025\RAS\CherryCreek_CO0128.p06.hdf' # RAS HDF file
        #breach_file_path = fr'C:\~Kurt\~Projects\Arkabutla\2025_Updates\Model\ArkabutlaBreach_202.p01.hdf' # RAS HDF file
        #breach_file_path = fr'C:\~Kurt\~Projects\Buford_Dam_IES\2025_RAS\RAS_2025-10\Buford_GA00824.p04.hdf' # RAS HDF file

        #nonbreach_file_path = fr'C:\~Kurt\~Projects\Cherry_Creek_Dam_2025\RAS\CherryCreek_CO0128.p05.hdf' # RAS HDF file
        #nonbreach_file_path = fr'C:\~Kurt\~Projects\Arkabutla\2025_Updates\Model\ArkabutlaBreach_202.p04.hdf' # RAS HDF file
        #nonbreach_file_path = fr'C:\~Kurt\~Projects\Buford_Dam_IES\2025_RAS\RAS_2025-10\Buford_GA00824.p77.hdf' # RAS HDF file

        ### Output Location
        #output_folder = fr'C:\~Kurt\GIS_Resources\Toolboxes\2025_HDF_FileReading\junk3' # Existing folder where you want outputs

        ### Blank variables for input geometry
        input_1dsas = ''
        input_2dcells = ''
        input_crosssections = ''

        ### If you already have geometry and don't want it to be created, input here
        #input_2dcells = fr'C:\~Kurt\~Projects\Cherry_Creek_Dam_2025\RAS\Calculated Layers\2D_Cells.shp' # Exported geometry from RAS Mapper
        #input_crosssections = fr'C:\~Kurt\~Projects\Arkabutla\RASnew\Calculated Layers\cross_sections.shp' # Exported geometry from RAS Mapper
        #input_1dsas = fr'C:\~Kurt\~Projects\Cherry_Creek_Dam_2025\RAS\Calculated Layers\1D_StroageAreas.shp' # Exported geometry from RAS Mapper

        # Output Ecel file name, now set near end of simulation so it can use the breach plan ShortID in name
        #output_excel = fr'{output_folder}\testexcel.xlsx'

        ### Script variables
        arrivalthreshold1 = 0.1      #threshold in feet for when a difference in stages is considered an arrival time
        points_join1_2djoin = fr'{output_folder}\Interim_join1_pointto2d.shp'
        points_join2_xsjoin = fr'{output_folder}\Interim_join2_pointtoxs.shp'
        points_join3_1dsajoin = fr'{output_folder}\Interim_join3_pointto1dsa.shp'

        jointargets = input_point_shapefile

        def print(*args, **kwargs):
            try:
                if args:
                    msg = " ".join(str(a) for a in args)
                else:
                    msg = ""
                messages.addMessage(msg)
            except Exception:
                # fallback to a simple representation if formatting fails
                try:
                    messages.addMessage(str(args))
                except Exception:
                    pass

        # printout of time now or start time of script
        script_start_time = datetime.now()
        print(f"Script started at: {(script_start_time).strftime('%Y-%m-%d %H:%M:%S')}")

        # === Inserted: helper to build feature classes from the HDF (uses only arcpy, h5py, numpy) ===

        def dump_hdf_groups(hdf_path):
            """Quick inspector to print groups/datasets to help find geometry paths when troubleshooting."""
            with h5py.File(hdf_path, 'r') as _f:
                def _walk(name, obj):
                    if isinstance(obj, h5py.Dataset):
                        print(f"DS: {name} shape={obj.shape} dtype={obj.dtype}")
                    else:
                        print(f"GR: {name}")
                _f.visititems(_walk)
        #dump_hdf_groups(breach_file_path)

        def _ensure_spatial_ref_from_source(src_path):
            try:
                return arcpy.Describe(src_path).spatialReference
            except Exception:
                return None

        def _split_workspace_and_name(out_fc):
            """
            Return (workspace, name, full_path) suitable for arcpy.CreateFeatureclass_management.
            Ensures the feature class name starts with a letter/underscore and contains only valid chars.
            Handles in_memory, folders (shapefiles) and file geodatabases.
            """
            # allow either 'in_memory/name', 'in_memory\\name', a full path or just a name
            print(f"Splitting output feature class path: {out_fc}")
            if out_fc.lower().startswith('in_memory'):
                workspace = 'in_memory'
                name = out_fc.replace('\\', '/').split('/')[-1]
                full_path = f'in_memory\\{name}'
                #print(f"A Detected in_memory workspace: {full_path}")
            elif os.path.isabs(out_fc):
                workspace = os.path.dirname(out_fc)
                #print(f"Workspace detected: {workspace}")
                name = os.path.basename(out_fc)
                #print(f"Name detected: {name}")
                # if workspace is a file geodatabase, keep name as-is; otherwise assume shapefile and ensure .shp extension
                if not os.path.splitext(name)[1]:
                    if not workspace.lower().endswith('.gdb'):
                        name = f"{name}.shp"
                full_path = os.path.join(workspace, name)
                #print(f"B Detected in_memory workspace: {full_path}")
            else:
                # treat bare name as in_memory
                workspace = 'in_memory'
                name = out_fc.replace('\\', '/').split('/')[-1]
                full_path = f'in_memory\\{name}'
                #print(f"C Detected in_memory workspace: {full_path}")

            # sanitize name: replace invalid chars with underscore (keep extension if present)
            base, ext = os.path.splitext(name)
            base = re.sub(r'[^A-Za-z0-9_]', '_', base)
            if not re.match(r'^[A-Za-z_]', base):
                base = f'hdf_{base}'
            name = base + ext
            # rebuild full_path
            full_path = os.path.join(workspace, name) if workspace != 'in_memory' else f'in_memory\\{name}'
            #print(f"D Detected in_memory workspace: {full_path}")
            return workspace, name, full_path

        def create_2d_cells_from_hdf(hdf_path, out_fc='in_memory/2d_cells', src_sr_path=None):
            """
            Read 2D mesh cells from HDF and build a polygon feature class.
            Tries multiple dataset name patterns used by different HEC‑RAS versions:
            - Preferred: FacePoints Coordinate + Cells FacePoint Indexes
            - Fallbacks: Nodes/Coordinates + Elements/Connectivity, etc.
            Returns the created FC path or None.
            """
            sr = _ensure_spatial_ref_from_source(src_sr_path) if src_sr_path else None

            workspace, fc_name, out_fc_path = _split_workspace_and_name(out_fc)
            if arcpy.Exists(out_fc_path):
                arcpy.Delete_management(out_fc_path)
            res = arcpy.CreateFeatureclass_management(workspace, fc_name, "POLYGON", None, "DISABLED", "DISABLED", sr)
            created_fc = res.getOutput(0)
            arcpy.AddField_management(created_fc, "MeshName", "TEXT", field_length=50)
            arcpy.AddField_management(created_fc, "CellIndex", "LONG")

            created = False
            insert_rows = []

            with h5py.File(hdf_path, 'r') as f:
                # locate the 2D flow areas group (common paths)
                if '/Geometry/2D Flow Areas' in f:
                    base = '/Geometry/2D Flow Areas'
                    candidates = list(f[base].keys())
                elif '/Geometry/2D Flow Area' in f:
                    base = '/Geometry/2D Flow Area'
                    candidates = list(f[base].keys())
                else:
                    print("No 2D Flow Areas group found. Run dump_hdf_groups() to inspect HDF structure.")
                    return None

                print("2D meshes found:", candidates)

                for mesh in candidates:
                    mesh_path = f"{base}/{mesh}"
                    # Try a number of known dataset names (from screenshot / typical RAS HDFs)
                    coord_candidates = [
                        f"{mesh_path}/Nodes/Coordinates",
                        f"{mesh_path}/FacePoints Coordinate",
                        f"{mesh_path}/FacePoints/Coordinate",
                        f"{mesh_path}/FacePoints/Coordinates",
                        f"{mesh_path}/FacePoints/Coordinates/0",
                        f"{mesh_path}/Cells Center Coordinate",          # centroid only (fallback)
                        f"{mesh_path}/Cells/CenterCoordinate",
                        f"{mesh_path}/Cells/FacePoint Coordinates",
                        f"{mesh_path}/Cells/FacePoints/Coordinate"
                    ]
                    conn_candidates = [
                        f"{mesh_path}/Cells FacePoint Indexes",
                        f"{mesh_path}/Cells/FacePoint Indexes",
                        f"{mesh_path}/Cells/FacePointIndexes",
                        f"{mesh_path}/Cells/FacePoint Index Values",
                        f"{mesh_path}/Elements/Connectivity",
                        f"{mesh_path}/Elements/Connectivity/0",
                        f"{mesh_path}/Faces FacePoint Indexes",
                        f"{mesh_path}/Faces/FacePointIndexes",
                    ]

                    coords = None
                    elems = None

                    # find coordinates dataset
                    for p in coord_candidates:
                        if p in f:
                            coords = f[p][:]
                            print(f"Using coords dataset: {p}")
                            break

                    # find connectivity dataset (per-cell index arrays pointing to facepoints or nodes)
                    for p in conn_candidates:
                        if p in f:
                            elems = f[p][:]
                            print(f"Using connectivity dataset: {p}")
                            break

                    if coords is None or elems is None:
                        #print(f"Skipping mesh '{mesh}': couldn't find coords or connectivity at expected locations under {mesh_path}")
                        # if coords is centroid-only but no connectivity, mention it
                        if coords is not None and coords.shape[-1] == 2:
                            print(f"  Found centroid coords for '{mesh}', but no polygon connectivity dataset present.")
                        continue

                    coords = np.asarray(coords)
                    elems = np.asarray(elems)

                    # If coordinates are stored as structured array or extra dims, attempt to coerce to (N,2)
                    if coords.ndim == 1 and coords.dtype.names:
                        # try to extract X and Y fields commonly named 'X','Y' or 'Longitude','Latitude'
                        names = coords.dtype.names
                        if 'X' in names and 'Y' in names:
                            coords = np.vstack([coords['X'], coords['Y']]).T
                        elif 'Longitude' in names and 'Latitude' in names:
                            coords = np.vstack([coords['Longitude'], coords['Latitude']]).T
                        else:
                            # try first two fields
                            coords = np.vstack([coords[names[0]], coords[names[1]]]).T

                    # Some connectivity arrays contain negative padding or 1-based indices
                    try:
                        if elems.size and elems.min() > 0:
                            elems = elems - 1
                    except Exception:
                        pass

                    # elems may be shape (n_cells, max_vertices) or a ragged list; iterate robustly
                    for idx, conn in enumerate(elems):
                        try:
                            conn_arr = np.array(conn, dtype=int)
                        except Exception:
                            # try convert from object
                            conn_arr = np.asarray(conn).astype(int)
                        # remove invalid/padding indices (negative or >= coords length)
                        conn_arr = conn_arr[(conn_arr >= 0) & (conn_arr < len(coords))]
                        if conn_arr.size < 3:
                            continue
                        pts = [arcpy.Point(float(coords[i, 0]), float(coords[i, 1])) for i in conn_arr]
                        # ensure ring closed
                        if not (pts[0].X == pts[-1].X and pts[0].Y == pts[-1].Y):
                            pts.append(pts[0])
                        array = arcpy.Array(pts)
                        poly = arcpy.Polygon(array, sr) if sr else arcpy.Polygon(array)
                        insert_rows.append((poly, mesh, int(idx)))
                    created = True

            if not created:
                print("No 2D meshes created from HDF.")
                return None

            # bulk insert into created feature class
            with arcpy.da.InsertCursor(created_fc, ['SHAPE@', 'MeshName', 'CellIndex']) as icur:
                for rec in insert_rows:
                    icur.insertRow(rec)
            return created_fc

        def create_cross_sections_from_hdf(hdf_path, out_fc='in_memory/cross_sections', src_sr_path=None):
            """
            Build a polyline feature class for cross sections using:
            - /Geometry/Cross Sections/Attributes   (attributes: River, Reach, RiverStat, etc.)
            - /Geometry/Cross Sections/Polyline Info (first column = Point Starting Index)
            - /Geometry/Cross Sections/Polyline Points (Nx2 coordinates concatenated)
            Falls back gracefully for a few variant dataset names and record layouts.
            """
            sr = _ensure_spatial_ref_from_source(src_sr_path) if src_sr_path else None

            workspace, fc_name, out_fc_path = _split_workspace_and_name(out_fc)
            if arcpy.Exists(out_fc_path):
                arcpy.Delete_management(out_fc_path)
            res = arcpy.CreateFeatureclass_management(workspace, fc_name, "POLYLINE", None, "DISABLED", "DISABLED", sr)
            created_fc = res.getOutput(0)
            arcpy.AddField_management(created_fc, "River", "TEXT", field_length=50)
            arcpy.AddField_management(created_fc, "Reach", "TEXT", field_length=50)
            arcpy.AddField_management(created_fc, "RiverStat", "TEXT", field_length=50)

            created = False
            with h5py.File(hdf_path, 'r') as f:
                base = '/Geometry/Cross Sections'
                if base not in f:
                    print("Cross Sections group not found in HDF.")
                    return None

                # dataset candidates
                pts_ds = f.get(f'{base}/Polyline Points') or f.get(f'{base}/PolylinePoints') or f.get(f'{base}/Coordinates') or f.get(f'{base}/Polyline Coordinates') or f.get(f'{base}/Coordinates/0')
                parts_ds = f.get(f'{base}/Polyline Info') or f.get(f'{base}/PolylineInfo') or f.get(f'{base}/Polyline_Parts') or f.get(f'{base}/Polyline Parts')
                attrs_ds = f.get(f'{base}/Attributes') or f.get(f'{base}/Attributes/Table') or f.get(f'{base}/Attributes/Attributes')

                if pts_ds is None:
                    print("Polyline Points dataset not found under Cross Sections. Run dump_hdf_groups() and update paths.")
                    return None
                pts = np.asarray(pts_ds[:])

                # normalize points to Nx2
                if pts.ndim == 1 and pts.dtype.names:
                    names_dt = pts.dtype.names
                    if 'X' in names_dt and 'Y' in names_dt:
                        pts_xy = np.vstack([pts['X'], pts['Y']]).T
                    elif 'Longitude' in names_dt and 'Latitude' in names_dt:
                        pts_xy = np.vstack([pts['Longitude'], pts['Latitude']]).T
                    else:
                        pts_xy = np.vstack([pts[names_dt[0]], pts[names_dt[1]]]).T
                elif pts.ndim == 2 and pts.shape[1] >= 2:
                    pts_xy = pts[:, :2]
                elif pts.ndim == 1 and pts.size % 2 == 0:
                    pts_xy = pts.reshape((-1, 2))
                else:
                    print("Unrecognized Polyline Points array shape:", pts.shape)
                    return None

                # derive starts/counts (same approach as storage areas)
                starts = None
                counts = None

                if parts_ds is not None:
                    try:
                        raw_parts = parts_ds[:]
                        arr = np.asarray(raw_parts)
                        if arr.ndim == 2 and np.issubdtype(arr.dtype, np.integer):
                            starts = arr[:, 0].astype(int).tolist()
                            if arr.shape[1] >= 2 and np.issubdtype(arr[:, 1].dtype, np.integer):
                                counts = arr[:, 1].astype(int).tolist()
                        elif arr.ndim == 1 and np.issubdtype(arr.dtype, np.integer):
                            starts = arr.astype(int).tolist()
                        else:
                            try:
                                starts = [int(row[0]) for row in raw_parts]
                            except Exception:
                                starts = None
                    except Exception:
                        starts = None

                # fallback: try to parse Attributes for start indices if needed (left as-is)
                if starts is not None and len(starts) > 0:
                    smin = min(starts)
                    if smin > 0:
                        starts = [int(s) - 1 for s in starts]
                    starts = sorted(int(s) for s in starts)

                if starts is None:
                    print("Could not derive polyline start indices. Run dump_hdf_groups() and paste outputs.")
                    return None

                poly_ranges = []
                if counts is not None:
                    for s, c in zip(starts, counts):
                        poly_ranges.append((int(s), int(s + c)))
                else:
                    for i, s in enumerate(starts):
                        e = starts[i+1] if i+1 < len(starts) else len(pts_xy)
                        poly_ranges.append((int(s), int(e)))

                # read attribute strings for River/Reach/RiverStat (if present)
                attr_rows = []
                if attrs_ds is not None:
                    raw_attrs = attrs_ds[:]
                    # debug: print attribute field names to confirm available keys
                    #try:
                        #print("Cross-section attribute fields:", getattr(raw_attrs, 'dtype', None).names)
                    #except Exception:
                        #pass

                    if hasattr(raw_attrs, 'dtype') and raw_attrs.dtype.names:
                        for r in raw_attrs:
                            rowd = {}
                            for fn in raw_attrs.dtype.names:
                                v = r[fn]
                                # decode bytes
                                if isinstance(v, (bytes, np.bytes_)):
                                    v = v.decode('utf-8', errors='ignore').rstrip('\x00')
                                # numeric -> string (preserve integer appearance)
                                elif np.isscalar(v) and (np.issubdtype(type(v), np.integer) or np.issubdtype(type(v), np.floating)):
                                    try:
                                        fv = float(v)
                                        if fv.is_integer():
                                            v = str(int(fv))
                                        else:
                                            v = str(fv)
                                    except Exception:
                                        v = str(v)
                                else:
                                    v = str(v)
                                rowd[fn] = v
                            attr_rows.append(rowd)
                    else:
                        # array of simple values (bytes/strings)
                        for r in raw_attrs:
                            if isinstance(r, (bytes, np.bytes_)):
                                attr_rows.append({'Name': r.decode('utf-8', errors='ignore').rstrip('\x00')})
                            else:
                                attr_rows.append({'Name': str(r)})

                # insert polylines
                with arcpy.da.InsertCursor(created_fc, ['SHAPE@', 'River', 'Reach', 'RiverStat']) as icur:
                    for i, (s, e) in enumerate(poly_ranges):
                        if s >= e or s < 0 or e > len(pts_xy):
                            continue
                        pts_list = [arcpy.Point(float(x), float(y)) for x, y in pts_xy[s:e]]
                        if len(pts_list) < 2:
                            continue
                        array = arcpy.Array(pts_list)
                        polyline = arcpy.Polyline(array, sr) if sr else arcpy.Polyline(array)
                        river = reach = RiverStat = ''
                        if i < len(attr_rows):
                            rrow = attr_rows[i]
                            # prefer decoded/string values; handle RS explicitly
                            river = rrow.get('River') or rrow.get('river') or rrow.get('RiverName') or rrow.get('Name') or ''
                            reach = rrow.get('Reach') or rrow.get('reach') or ''
                            # RS sometimes stores numeric or bytes; normalized above to string
                            RiverStat = rrow.get('RS') or rrow.get('RiverStat') or rrow.get('riverstati') or rrow.get('riverstation') or rrow.get('river_sta') or rrow.get('river_station') or ''
                            # ensure simple string (no numpy types)
                            if not isinstance(RiverStat, str):
                                RiverStat = str(RiverStat)
                        icur.insertRow([polyline, river, reach, RiverStat])
                        created = True

            return created_fc if created else None


        def create_storage_areas_from_hdf(hdf_path, out_fc, src_sr_path=None):
            """
            Build polygon feature class for 1D storage areas using alternate HDF layout that
            stores polygons as 'Polygon Points', 'Polygon Parts' and 'Polygon Info'.
            Tries several heuristics for parts/points layouts used by different RAS HDF versions.
            """
            sr = _ensure_spatial_ref_from_source(src_sr_path) if src_sr_path else None

            workspace, fc_name, out_fc_path = _split_workspace_and_name(out_fc)
            if arcpy.Exists(out_fc_path):
                arcpy.Delete_management(out_fc_path)
            print(f"Creating Storage Areas FC at: {workspace}, {fc_name}, {out_fc_path}, SR={sr.name if sr else 'None'}")
            res = arcpy.CreateFeatureclass_management(workspace, fc_name, "POLYGON", None, "DISABLED", "DISABLED", sr)
            created_fc = res.getOutput(0)
            arcpy.AddField_management(created_fc, "Name", "TEXT", field_length=80)

            created = False
            with h5py.File(hdf_path, 'r') as f:
                base = '/Geometry/Storage Areas'
                if base not in f:
                    print("Storage Areas group not found in HDF.")
                    return None

                #tables could be different in different version of RAS, that is the or functions
                #XY Points Table
                pts_ds = f.get(f'{base}/Polygon Points') or f.get(f'{base}/PolygonPoints') or f.get(f'{base}/Polygon_Points')
                #Polygon Info table with start points in first column
                parts_ds = f.get(f'{base}/Polygon Info') or f.get(f'{base}/PolygonParts') or f.get(f'{base}/Polygon_Parts')
                #Attributes table with names, etc.
                info_ds = f.get(f'{base}/Attributes') or f.get(f'{base}/Polygon Info') or f.get(f'{base}/PolygonInfo') or f.get(f'{base}/Polygon_Info')

                ### Get POINTS list from points table
                if pts_ds is None:
                    print("Polygon Points dataset not found under Storage Areas. Run dump_hdf_groups() and update paths.")
                    return None

                pts = np.asarray(pts_ds[:])

                # normalize points to Nx2
                if pts.ndim == 1 and pts.dtype.names:
                    names_dt = pts.dtype.names
                    if 'X' in names_dt and 'Y' in names_dt:
                        pts_xy = np.vstack([pts['X'], pts['Y']]).T
                    elif 'Longitude' in names_dt and 'Latitude' in names_dt:
                        pts_xy = np.vstack([pts['Longitude'], pts['Latitude']]).T
                    else:
                        pts_xy = np.vstack([pts[names_dt[0]], pts[names_dt[1]]]).T
                elif pts.ndim == 2 and pts.shape[1] >= 2:
                    pts_xy = pts[:, :2]
                elif pts.ndim == 1 and pts.size % 2 == 0:
                    pts_xy = pts.reshape((-1, 2))
                else:
                    print("Unrecognized Polygon Points array shape:", pts.shape)
                    return None

                ### Get Starting Point index for each polygon from Polygon Info table
                starts = None
                counts = None

                if parts_ds is not None:
                    print("Attempting to extract polygon starts from parts_ds")
                    try:             
                        raw_info = parts_ds[:]              # whatever shape/format h5py gives
                        # Fast, generic attempt: take the first element of each row (works for 2D numeric and many other layouts)
                        try:
                            starts = [int(row[0]) for row in raw_info]
                            print("generic attempt worked")
                        except Exception:
                            starts = None
                            print("Generic attempt did not work")
                        

                        # If still found, coerce 1-based -> 0-based and sort
                        if starts:
                            smin = min(starts)
                            if smin > 0:
                                starts = [int(s) - 1 for s in starts]
                            starts = sorted(int(s) for s in starts)

                    except Exception:
                        starts = None
                        print("error extracting starts from info_ds")

                    #print(("Starts: ", starts))
                        # 1) If it's a plain numeric 2D array, first column is start index

                # Get NAMES from info_ds provides names and equal-split is possible
                sa_names = []
                if info_ds is not None:
                    try:
                        raw_info = info_ds[:]
                        if hasattr(raw_info, 'dtype') and raw_info.dtype.names:
                            print("G") #prints
                            # try to find a string field
                            for fn in raw_info.dtype.names:
                                if raw_info[fn].dtype.kind in ('S', 'U', 'O'):
                                    sa_names = [ (b.decode('utf-8', errors='ignore').rstrip('\x00') if isinstance(b, (bytes, np.bytes_)) else str(b)) for b in raw_info[fn] ]
                                    #print(sa_names)
                                    break
                            if not sa_names:
                                print("not sa names")
                                sa_names = [str(r) for r in raw_info]
                        else:
                            sa_names = [ (b.decode('utf-8', errors='ignore').rstrip('\x00') if isinstance(b, (bytes, np.bytes_)) else str(b)) for b in raw_info ]
                    except Exception:
                        sa_names = []

                num_polys = len(sa_names) if sa_names else None
                print("Number of polygons:", num_polys)
                if starts is None and num_polys and len(pts_xy) % num_polys == 0:
                    print("H") #doesn't print
                    cnt = len(pts_xy) // num_polys
                    starts = [i * cnt for i in range(num_polys)]
                    counts = [cnt] * num_polys

                if starts is None:
                    print("Could not derive polygon start indices from 'Polygon Parts' or 'Polygon Info'. Try running dump_hdf_groups() and paste samples.")
                    return None

                # compute per-polygon point ranges
                poly_ranges = []
                if counts is not None:
                    print("I") #prints
                    for s, c in zip(starts, counts):
                        e = s + c
                        poly_ranges.append((int(s), int(e)))
                else:
                    for i, s in enumerate(starts):
                        e = starts[i+1] if i+1 < len(starts) else len(pts_xy)
                        poly_ranges.append((int(s), int(e)))

                # insert polygons
                with arcpy.da.InsertCursor(created_fc, ['SHAPE@', 'Name']) as icur:
                    for i, (s, e) in enumerate(poly_ranges):
                        if s >= e or s < 0 or e > len(pts_xy):
                            continue
                        ring_pts = [arcpy.Point(float(x), float(y)) for x, y in pts_xy[s:e]]
                        #print(i)
                        #print(ring_pts)
                        if len(ring_pts) < 3:
                            continue
                        # close ring if needed
                        if not (ring_pts[0].X == ring_pts[-1].X and ring_pts[0].Y == ring_pts[-1].Y):
                            print("closing ring")
                            ring_pts.append(ring_pts[0])
                        array = arcpy.Array(ring_pts)
                        poly = arcpy.Polygon(array, sr) if sr else arcpy.Polygon(array)
                        nm = sa_names[i] if i < len(sa_names) else f"SA_{i}"
                        icur.insertRow([poly, nm])
                        created = True

            return created_fc if created else None
        # === End inserted helpers ===

        # If shapefiles were not supplied, see if they already exist, if not, try to build them from the HDF geometry (uses breach_file_path as geometry source)
        #sa_featureclass = r'in_memory\storage_areas'
        sa_featureclass = fr'{output_folder}\storage_areas.shp'
        if arcpy.Exists(sa_featureclass) and not input_1dsas:
            input_1dsas = sa_featureclass
            print(f"Using existing storage areas: {sa_featureclass}")
            
        if not input_1dsas:
            print(f"No 1D storage areas shapefile provided. Attempting to build storage area polygons from HDF...")
            input_1dsas = create_storage_areas_from_hdf(breach_file_path, out_fc=sa_featureclass, src_sr_path=input_point_shapefile)
            if not input_1dsas:
                print("Failed to create storage areas from HDF. Inspect the HDF with dump_hdf_groups(breach_file_path).")

        #xs_featureclass = r'in_memory\cross_sections'
        xs_featureclass = fr'{output_folder}\cross_sections.shp'
        if arcpy.Exists(xs_featureclass) and not input_crosssections:
            input_crosssections = xs_featureclass
            print(f"Using existing cross sections: {xs_featureclass}")
        if not input_crosssections:
            print("No cross-section shapefile provided. Attempting to build cross-section points from HDF...")
            input_crosssections = create_cross_sections_from_hdf(breach_file_path, out_fc=xs_featureclass, src_sr_path=input_point_shapefile)
            if not input_crosssections:
                print("Failed to create cross-section points from HDF. Inspect the HDF with dump_hdf_groups(breach_file_path).")

        #twod_featureclass = r'in_memory\2d_cells'
        twod_featureclass = fr'{output_folder}\hdf_2d_cells.shp'
        if arcpy.Exists(twod_featureclass) and not input_2dcells:
            input_2dcells = twod_featureclass
            print(f"Using existing storage areas: {twod_featureclass}")
        if not input_2dcells:
            print("No 2D cells shapefile provided. Attempting to build 2D cell polygons from HDF...")
            input_2dcells = create_2d_cells_from_hdf(breach_file_path, out_fc=twod_featureclass, src_sr_path=input_point_shapefile)
            if not input_2dcells:
                print("Failed to build 2D cells from HDF. Run dump_hdf_groups(breach_file_path) to inspect HDF and update helper paths.")

        ########### Geometry made, do spatial joins to pull data ############
        print("Starting spatial joins...")
        def _add_target_fields_to_fm(fm, target_fc):
            for f in arcpy.ListFields(target_fc):
                if f.type in ('OID', 'Geometry'):
                    continue
                fmap = arcpy.FieldMap()
                fmap.addInputField(target_fc, f.name)
                fm.addFieldMap(fmap)

        def _add_join_field_to_fm(fm, join_fc, src_field, out_name=None, length=None):
            fmap = arcpy.FieldMap()
            print(f"Adding join field: joinfc: {join_fc} src_field: {src_field} as {out_name or src_field}")
            fmap.addInputField(join_fc, src_field)
            outfld = fmap.outputField
            outfld.name = out_name or src_field
            outfld.aliasName = outfld.name
            if length and hasattr(outfld, 'length'):
                outfld.length = length
            fmap.outputField = outfld
            fm.addFieldMap(fmap)

        # Join 1 - points to 2D cells
        if input_2dcells:
            # build explicit FieldMappings: keep target attributes and pull CellIndex/MeshName from 2D cells
            fm = arcpy.FieldMappings()
            _add_target_fields_to_fm(fm, jointargets)
            # add 2D fields
            _add_join_field_to_fm(fm, input_2dcells, 'CellIndex', 'CellIndex')
            _add_join_field_to_fm(fm, input_2dcells, 'MeshName', 'MeshName', length=50)

            if arcpy.Exists(points_join1_2djoin):
                arcpy.Delete_management(points_join1_2djoin)
            arcpy.analysis.SpatialJoin(
                target_features=jointargets,
                join_features=input_2dcells,
                out_feature_class=points_join1_2djoin,
                join_operation="JOIN_ONE_TO_ONE",
                join_type="KEEP_ALL",
                field_mapping=fm,
                match_option="INTERSECT"
            )
            jointargets = points_join1_2djoin

        # Join 2 - points to nearest cross section
        if input_crosssections:
            fm = arcpy.FieldMappings()
            _add_target_fields_to_fm(fm, jointargets)
            _add_join_field_to_fm(fm, input_crosssections, 'River', 'River', length=50)
            _add_join_field_to_fm(fm, input_crosssections, 'Reach', 'Reach', length=50)
            _add_join_field_to_fm(fm, input_crosssections, 'RiverStat', 'RiverStat', length=50)  # map RS -> RiverStat

            if arcpy.Exists(points_join2_xsjoin):
                arcpy.Delete_management(points_join2_xsjoin)
            arcpy.analysis.SpatialJoin(
                target_features=jointargets,
                join_features=input_crosssections,
                out_feature_class=points_join2_xsjoin,
                join_operation="JOIN_ONE_TO_ONE",
                join_type="KEEP_ALL",
                field_mapping=fm,
                match_option="CLOSEST"
            )
            jointargets = points_join2_xsjoin

        # Join 3 - points to 1D Storage Areas
        if input_1dsas:
            fm = arcpy.FieldMappings()
            _add_target_fields_to_fm(fm, jointargets)
            # ensure storage area Name maps to Name_1 on the output
            _add_join_field_to_fm(fm, input_1dsas, 'Name', 'Name_1', length=80)

            if arcpy.Exists(points_join3_1dsajoin):
                arcpy.Delete_management(points_join3_1dsajoin)
            arcpy.analysis.SpatialJoin(
                target_features=jointargets,
                join_features=input_1dsas,
                out_feature_class=points_join3_1dsajoin,
                join_operation="JOIN_ONE_TO_ONE",
                join_type="KEEP_ALL",
                field_mapping=fm,
                match_option="INTERSECT"
            )
            jointargets = points_join3_1dsajoin

        # Add field for hydraulic data type
        arcpy.management.AddField(
            in_table=jointargets,
            field_name="HydType",
            field_type="TEXT",
            field_precision=None,
            field_scale=None,
            field_length=10,
            field_alias="",
            field_is_nullable="NULLABLE",
            field_is_required="NON_REQUIRED",
            field_domain=""
        )

        # Calculate field for hydraulic data type. Fields could vary if you only had 2D or only had 1D geometry

        # List of required fields if you had combined 1D/2D
        required_fields = ["MeshName", "RiverStat", "Name_1"]

        # Get the existing fields in the table
        existing_fields = [field.name for field in arcpy.ListFields(jointargets)]

        # Create an expression that dynamically checks for missing fields
        expression_parts = []
        for field in required_fields:
            if field in existing_fields:
                expression_parts.append(f"!{field}!")
            else:
                expression_parts.append("''")  # Substitute missing fields with empty strings

        # Build the final expression string
        expression = f"HydType({', '.join(expression_parts)})"
        code_text = textwrap.dedent("""
        def HydType(MeshName, RiverStat, Name_1):
            if MeshName and str(MeshName).strip():
                return "2D"
            elif Name_1 and str(Name_1).strip():
                return "1D_SA"
            else:
                return "1D_XS"
        """)

        # Execute CalculateField
        arcpy.management.CalculateField(
            in_table=jointargets,
            field="HydType",
            expression=expression,
            expression_type="PYTHON3",
            code_block=code_text,
            field_type="TEXT",
            enforce_domains="NO_ENFORCE_DOMAINS"
        )
        print("Calculated hydraulic data types")

        # Create an Excel workbook
        wb = Workbook()
        ws = wb.create_sheet(title="Summary")
        # Remove the default sheet created at Workbook instantiation
        default_sheet = wb['Sheet']
        wb.remove(default_sheet)

        # Open the input HDF5 file
        hdf_breach = h5py.File(breach_file_path, 'r')
        hdf_nonbreach = h5py.File(nonbreach_file_path, 'r')

        # Define all possible fields
        all_fields = ["FID", input_point_name_field, "HydType", "MeshName", "CellIndex", "River", "Reach", "RiverStat", "Name_1"]

        # Get available fields in the dataset
        existing_fields = {f.name for f in arcpy.ListFields(jointargets)}

        # Filter only available fields
        selected_fields = [field for field in all_fields if field in existing_fields]

        # Print available fields for debugging
        #print(f"Available Fields: {existing_fields}")

        # Count total features
        feature_count = len(list(arcpy.da.SearchCursor(jointargets, ["FID"])))
        print(f"Total features in dataset: {feature_count}")
        if feature_count > 50:
            print("Warning: Large number of features may result in a large Excel file and longer processing time.")

        # Add below here logic to extract the Plan ShortID value from each HDF, maybe Simulation Start Time also
        hdf_file = h5py.File(breach_file_path, 'r')
        hdf_planinfotable = hdf_file.get(f'/Plan Data/Plan Information')
        if hdf_planinfotable is not None:
            plan_filename_value = hdf_planinfotable.attrs.get('Plan Filename')
            plan_filename_value = plan_filename_value.decode('ascii', errors='ignore').rstrip('\x00').strip()
            plan_shortid_value = hdf_planinfotable.attrs.get('Plan ShortID')
            plan_shortid_value = plan_shortid_value.decode('ascii', errors='ignore').rstrip('\x00').strip()
            plan_number = plan_filename_value.split('.')[1]
            print(plan_number)
            #plan_flowtitle_value = hdf_planinfotable.attrs.get('Flow Title')
            #plan_flowtitle_value = plan_flowtitle_value.decode('ascii', errors='ignore').rstrip('\x00').strip()
            plan_starttime_value = hdf_planinfotable.attrs.get('Simulation Start Time')
            plan_starttime_value = plan_starttime_value.decode('ascii', errors='ignore').rstrip('\x00').strip()
            plan_outputinterval_value = hdf_planinfotable.attrs.get('Base Output Interval')
            plan_outputinterval_value = plan_outputinterval_value.decode('ascii', errors='ignore').rstrip('\x00').strip()
            match = re.match(r"(\d+)([A-Za-z]+)", plan_outputinterval_value)
            if match:
                interval_num = match.group(1)   # "15" or "10" or "1"
                interval_unit = match.group(2)  # "MIN" or "HOUR"
            else:
                interval_num = None
                interval_unit = None
            #print(f"RAS Plan Filename: {plan_filename_value}")
            #print(f"RAS Plan ShortID: {plan_shortid_value}")
            #print(f"RAS Plan Flow Name: {plan_flowtitle_value}")
            print(f"RAS Start Time: {plan_starttime_value}")
            print(f"RAS Output Interval: {plan_outputinterval_value}")

            # parse "01Apr2025 10:00:00" -> datetime
            parsed_plan_start = None
            parsed_plan_start = datetime.strptime(plan_starttime_value, "%d%b%Y %H:%M:%S")
            print(f"RAS Start Time (parsed): {parsed_plan_start}")

        nbhdf_file = h5py.File(nonbreach_file_path, 'r')
        nbhdf_planinfotable = nbhdf_file.get(f'/Plan Data/Plan Information')
        nbplan_filename_value = nbhdf_planinfotable.attrs.get('Plan Filename')
        nbplan_filename_value = nbplan_filename_value.decode('ascii', errors='ignore').rstrip('\x00').strip()
        nbplan_shortid_value = nbhdf_planinfotable.attrs.get('Plan ShortID')
        nbplan_shortid_value = nbplan_shortid_value.decode('ascii', errors='ignore').rstrip('\x00').strip()
        nbplan_number = nbplan_filename_value.split('.')[1]

        # Navigate to datasets for location in hdf file using meshname as an f-string variable
        time_data = hdf_breach['/Results/Unsteady/Output/Output Blocks/Base Output/Unsteady Time Series/Time']

        # Convert time_data to human-readable times (if numerical), assuming time_data represents seconds
        time_values = time_data[:]
        #reference_time = datetime(2025, 4, 1, 10)  # Default MMC HEC-RAS reference time, assumes Feb 3, 2099 0:00
        reference_time = parsed_plan_start
        times = [reference_time + timedelta(days=t) for t in time_values] #assumes time table is in days, this converts to 2099-02-03 0:00:00

        #---------- add fields to point shapefile ---------------#
        # Get available fields in the dataset
        input_fields = {f.name for f in arcpy.ListFields(input_point_shapefile)}
        if input_fields is None or f"{plan_number}_ArrTm" not in input_fields:
            arcpy.AddField_management(input_point_shapefile, f"{plan_number}_ArrTm", "TEXT", field_length=50)
            arcpy.AddField_management(input_point_shapefile, f"{plan_number}_PkTim", "TEXT", field_length=50)
            arcpy.AddField_management(input_point_shapefile, f"{plan_number}_PkDif", "TEXT", field_length=50)


        ####### BEGIN Looping through points and extracting data ##########

        with arcpy.da.SearchCursor(jointargets, selected_fields) as cursor1:
            # begin looping through points in the shapefile
            for row1 in cursor1:
                field_values = dict(zip(selected_fields, row1))  # Convert row to dictionary for easier access
                FID = field_values.get("FID")
                point_name = field_values.get(input_point_name_field)
                hydro_type = field_values.get("HydType")
                print(f"FID: {FID}, Name: {point_name}, HydType: {hydro_type}")  # Print hydraulic type
                
                if hydro_type == "2D":
                    meshname = field_values.get("MeshName")
                    cellindex = field_values.get("CellIndex")
                    if cellindex is not None:
                        cellindex = int(cellindex)  # Cast to an integer
                    print(f"MeshName: {meshname}, CellIndex: {cellindex}")

                    # Access HDF5 datasets safely
                    breach_wse_data = hdf_breach.get(f'/Results/Unsteady/Output/Output Blocks/Base Output/Unsteady Time Series/2D Flow Areas/{meshname}/Water Surface')
                    nonbreach_wse_data = hdf_nonbreach.get(f'/Results/Unsteady/Output/Output Blocks/Base Output/Unsteady Time Series/2D Flow Areas/{meshname}/Water Surface')

                    with h5py.File(breach_file_path, 'r') as brhdf:
                        breach_dataset = brhdf[f'/Results/Unsteady/Output/Output Blocks/Base Output/Unsteady Time Series/2D Flow Areas/{meshname}/Water Surface']
                        breach_wse_values = breach_dataset[:, cellindex]
                    
                    with h5py.File(nonbreach_file_path, 'r') as nbbrhdf:
                        nonbreach_dataset = nbbrhdf[f'/Results/Unsteady/Output/Output Blocks/Base Output/Unsteady Time Series/2D Flow Areas/{meshname}/Water Surface']
                        nonbreach_wse_values = nonbreach_dataset[:, cellindex]

                    if breach_wse_values is None or nonbreach_wse_values is None:
                            print(f"Warning: Missing WSE data for MeshName '{meshname}'")
                            continue  # Skip iteration if data is missing

                elif hydro_type == "1D_XS":
                    river = field_values.get("River")
                    reach = field_values.get("Reach")
                    river_station = field_values.get("RiverStat")

                    # Initialize the index variable
                    xs_index = None
                    xs_matchcount = 0

                    # Access the list of storage area names in the HDF5 file
                    xs_station1 = hdf_breach.get(f'/Geometry/Cross Sections/Attributes')
                    #print(f"XS Names1: {xs_station1}")

                    # ...existing code...
                    # Access the list of cross-section attributes in the HDF5 file
                    cs_attrs = hdf_breach.get('/Geometry/Cross Sections/Attributes')
                    #print(f"XS Attributes dataset: {cs_attrs}")

                    xs_station1 = []
                    if cs_attrs is None:
                        print("Warning: Cross Sections Attributes not found in HDF.")
                    else:
                        raw_attrs = cs_attrs[:]
                        # candidate field names that may contain the river station (RS)
                        candidates = ('RS', 'RiverStat', 'RiverStati', 'riverstation', 'River_Station', 'RiverSta', 'Station')
                        # structured array with named fields
                        if hasattr(raw_attrs, 'dtype') and raw_attrs.dtype.names:
                            found_field = None
                            for fn in candidates:
                                if fn in raw_attrs.dtype.names:
                                    found_field = fn
                                    break
                            if not found_field:
                                # fallback: pick first text-like field
                                for fn in raw_attrs.dtype.names:
                                    if raw_attrs[fn].dtype.kind in ('S', 'U', 'O'):
                                        found_field = fn
                                        break
                            if found_field:
                                for v in raw_attrs[found_field]:
                                    if isinstance(v, (bytes, np.bytes_)):
                                        xs_station1.append(v.decode('utf-8', errors='ignore').rstrip('\x00'))
                                    else:
                                        xs_station1.append(str(v))
                            else:
                                # no named text field found: stringify entire rows as fallback
                                for row in raw_attrs:
                                    xs_station1.append(str(row))
                        else:
                            # plain array of values
                            for v in raw_attrs:
                                if isinstance(v, (bytes, np.bytes_)):
                                    xs_station1.append(v.decode('utf-8', errors='ignore').rstrip('\x00'))
                                else:
                                    xs_station1.append(str(v))

                    #print(f"XS RS values sample: {xs_station1[:10]}")

                    # Use SearchCursor to loop through the shapefile
                    #oid_field = arcpy.Describe(input_crosssections).OIDFieldName
                    with arcpy.da.SearchCursor(input_crosssections, ['FID', 'RiverStat']) as cursor2:
                        for row in cursor2:
                            if row[1] == river_station:
                                xs_matchcount += 1
                                xs_index = row[0]  # OID (Object ID) is the index
                                if xs_matchcount > 1:
                                    # If more than one match is found, break the loop and raise an error
                                    print(f"Error: Multiple rows with {river_station} found.")
                                    break

                    xs_index = xs_station1.index(river_station)
                    print(f"Found Cross Section '{river_station}' at index {xs_index}")

                    # Output the result
                    if xs_index is not None:
                        print(f"Found {river_station} at index {xs_index}")
                    else:
                        print(f"{river_station} not found in the shapefile.")

                    if river_station is not None and river is not None:
                        print(f"River: {river}, River Station: {river_station}")

                    with h5py.File(breach_file_path, 'r') as brhdf:
                        breach_dataset = brhdf[f'/Results/Unsteady/Output/Output Blocks/Base Output/Unsteady Time Series/Cross Sections/Water Surface']
                        breach_wse_values = breach_dataset[:, xs_index]
                    
                    with h5py.File(nonbreach_file_path, 'r') as nbbrhdf:
                        nonbreach_dataset = nbbrhdf[f'/Results/Unsteady/Output/Output Blocks/Base Output/Unsteady Time Series/Cross Sections/Water Surface']
                        nonbreach_wse_values = nonbreach_dataset[:, xs_index]

                    if breach_wse_values is None or nonbreach_wse_values is None:
                            print(f"Warning: Missing WSE data for XS '{river_station}'")
                            continue  # Skip iteration if data is missing

                elif hydro_type == "1D_SA":
                    sa_name = field_values.get("Name_1")

                    # Access the list of storage area names in the HDF5 file
                    storage_area_names = hdf_breach.get(f'/Geometry/Storage Areas/Attributes')
                    #print(f"Storage Area Names1: {storage_area_names}")

                    # Convert the structured array to a list of strings, strip null bytes, and isolate the numeric part
                    storage_area_names = [
                        name.tobytes().decode('utf-8', errors='ignore').rstrip('\x00').split('\x00')[0]  # Extract just the numeric part
                        for name in storage_area_names[:]
                    ]
                    #print(f"Storage Area Names2: {storage_area_names}")

                    # Find the index that matches the storage area name
                    if storage_area_names is not None:
                        try:
                            sa_index = storage_area_names.index(sa_name)
                            print(f"Found Storage Area '{sa_name}' at index {sa_index}")
                        except ValueError:
                            print(f"Warning: Storage Area '{sa_name}' not found in the list")
                            continue

                    with h5py.File(breach_file_path, 'r') as brhdf:
                        breach_dataset = brhdf[f'/Results/Unsteady/Output/Output Blocks/Base Output/Unsteady Time Series/Storage Areas/Water Surface']
                        breach_wse_values = breach_dataset[:, sa_index]
                    
                    with h5py.File(nonbreach_file_path, 'r') as nbbrhdf:
                        nonbreach_dataset = nbbrhdf[f'/Results/Unsteady/Output/Output Blocks/Base Output/Unsteady Time Series/Storage Areas/Water Surface']
                        nonbreach_wse_values = nonbreach_dataset[:, sa_index]

                    if breach_wse_values is None or nonbreach_wse_values is None:
                            print(f"Warning: Missing WSE data for XS '{sa_name}'")
                            continue  # Skip iteration if data is missing

                else:
                    print(f"   [Warning] Unknown hydraulic type for FID {FID}")

                ## Exit of the hydraulic data type if-else statement

                # Create an Excel worksheet and add data
                safe_point_name = re.sub(r'[:\\/*?\[\]]', '_', point_name)
                ws = wb.create_sheet(title=safe_point_name)

                # Set column widths
                ws.column_dimensions['A'].width = 20  # Adjust width for Time column
                ws.column_dimensions['D'].width = 20  # Adjust width for Time column
                ws.column_dimensions['E'].width = 15  # Adjust width for Breach_WSE column
                ws.column_dimensions['F'].width = 15  # Adjust width for NonBreach_WSE column
                ws.column_dimensions['G'].width = 15  # Adjust width for Difference column

                # Add headers starting in column C (row 1, column 4)
                ws.cell(row=1, column=1, value="Arrival Time")
                ws.cell(row=5, column=1, value="Peak Time")
                ws.cell(row=9, column=1, value="Peak Difference (ft)")
                ws.cell(row=12, column=1, value="Hydraulics Type")
                ws.cell(row=15, column=1, value="Location")

                start_column = 4
                ws.cell(row=1, column=start_column, value="Time")
                ws.cell(row=1, column=start_column + 1, value="Hours")
                ws.cell(row=1, column=start_column + 2, value=f"{plan_shortid_value} WSE")
                ws.cell(row=1, column=start_column + 3, value=f"{nbplan_shortid_value} WSE")
                ws.cell(row=1, column=start_column + 4, value="Difference")

                # Initialize flag to track the first instance where the difference > 0.1
                first_time_found = False
                max_difference = float('-inf')  # Initialize to negative infinity
                max_difference_time = None

                # Add data rows starting in column C
                start_data_row = 2
                for i, (time, br_elevation, nb_elevation) in enumerate(zip(times, breach_wse_values, nonbreach_wse_values), start=start_data_row):
                    difference = br_elevation - nb_elevation
                    ws.cell(row=i, column=start_column, value=time)
                    hours_since = float((time - reference_time).total_seconds() / 3600.0)
                    ws.cell(row=i, column=start_column + 1, value=hours_since)  # Hours since reference time
                    ws.cell(row=i, column=start_column + 2, value=br_elevation)
                    ws.cell(row=i, column=start_column + 3, value=nb_elevation)
                    ws.cell(row=i, column=start_column + 4, value=difference)

                    # Check if the difference is greater than 0.1 and if it's the first occurrence
                    if abs(difference) > arrivalthreshold1 and not first_time_found:
                        ws.cell(row=2, column=1, value=time)  # Write time to column 1
                        ws.cell(row=3, column=1, value=hours_since)  # Write time to column 1
                        arrival_time_hours = hours_since
                        print(time)
                        first_time_found = True  # Set flag to True so subsequent times are not written

                    # Track the maximum difference and its corresponding time
                    if abs(difference) > max_difference:
                        max_difference = round(abs(difference), 2)
                        max_difference_time = time
                        difference_hours_since = float((time - reference_time).total_seconds() / 3600.0)

                if max_difference_time is not None:
                    ws.cell(row=6, column=1, value=max_difference_time)  # Store max difference time in row 3, column 1
                    ws.cell(row=7, column=1, value=difference_hours_since)
                    ws.cell(row=10, column=1, value=max_difference)  # Store max difference value in row 3, column 2
                    print(f"Maximum difference: {max_difference} at time {max_difference_time}")

                ws.cell(row=13, column=1, value=hydro_type)  # Store hydraulic type in row 4, column 1
                if hydro_type == "2D":
                    ws.cell(row=16, column=1, value=f"Mesh: {meshname}, Cell: {cellindex}")  # Store location info in row 5, column 1
                elif hydro_type == "1D_XS":
                    ws.cell(row=16, column=1, value=f"River: {river}, RS: {river_station}")  # Store location info in row 5, column 1
                elif hydro_type == "1D_SA":
                    ws.cell(row=16, column=1, value=f"Storage Area: {sa_name}")  # Store location info in row 5, column 1

                # --- Compute Y-axis bounds based on WSE data ---
                all_vals = []

                try:
                    # Collect all numeric WSE values
                    all_vals.extend([float(v) for v in breach_wse_values if v is not None])
                    all_vals.extend([float(v) for v in nonbreach_wse_values if v is not None])
                except Exception:
                    all_vals = []

                if all_vals:
                    y_min = min(all_vals)
                    y_max = max(all_vals)

                    if y_min == y_max:
                        # If all values identical, give a buffer
                        buffer = 0.5 if abs(y_min) < 1 else abs(y_min) * 0.01
                        y_min -= buffer
                        y_max += buffer
                    else:
                        # Standard padding (5%)
                        pad = (y_max - y_min) * 0.05
                        y_min -= pad
                        y_max += pad

                    # Round to nearest step for cleaner axis appearance
                    step = 10.0
                    y_min = math.floor(y_min / step) * step
                    y_max = math.ceil(y_max / step) * step

                else:
                    # fallback values
                    y_min, y_max = 0.0, 1.0


                # --- Create Scatter Chart (XY) ---
                chart = ScatterChart()
                #chart = LineChart()
                chart.title = f"Water Surface Elevation at {point_name}"
                chart.style = 2                                # prevents Excel corruption
                #chart.scatterStyle = "lineMarker"                    # safe
                #chart.legend.position = "r"
                #chart.x_axis = NumericAxis()
                #chart.y_axis = NumericAxis()
                chart.x_axis.title = "Time (hours)"
                chart.y_axis.title = "Water Surface Elevation (ft)"

                # --- Apply Y-axis bounds ---
                chart.y_axis.scaling.min = y_min
                chart.y_axis.scaling.max = y_max

                # --- for Scatter: Define X (hours) and Y (WSE) ranges ---
                xvalues = Reference(ws, min_col=5, min_row=2, max_row=ws.max_row)

                y_breach = Reference(ws, min_col=6, min_row=2, max_row=ws.max_row)
                series1 = Series(values=y_breach, xvalues=xvalues, title=f"{plan_shortid_value} WSE")

                y_nonbreach = Reference(ws, min_col=7, min_row=2, max_row=ws.max_row)
                series2 = Series(values=y_nonbreach, xvalues=xvalues, title=f"{nbplan_shortid_value} WSE")

                chart.series.append(series1)
                chart.series.append(series2)

                if hazard_time_exists:
                    hazard_time1 = hazardtime1
                    hazard_time2 = hazard_time1
                    ws.cell(row=18, column=1, value="Hazard Time (hrs)")
                    ws.cell(row=19, column=1, value=hazard_time1)
                    ws.cell(row=20, column=1, value=hazard_time2)
                    ws.cell(row=21, column=1, value=y_min)
                    ws.cell(row=22, column=1, value=y_max)


                    hazard_linex = Reference(ws, min_col=1, min_row=19, max_row=20)
                    hazard_liney = Reference(ws, min_col=1, min_row=21, max_row=22)
                    series3 = Series(values=hazard_liney, xvalues=hazard_linex, title=f"Hazard Time")
                    chart.series.append(series3)

                    chart.series[2].graphicalProperties.line = LineProperties(prstDash='dash')

                # --- For LineChart: Define X (hours) and Y (WSE) ranges ---
                # Define the data range
                #time_col = Reference(ws, min_col=5, min_row=2, max_row=ws.max_row)  # X-axis (Time)
                #wse_cols = Reference(ws, min_col=6, max_col=7, min_row=1, max_row=ws.max_row)  # Y-axis (WSEs)

                # Add data to LineChart chart
                #chart.add_data(wse_cols, titles_from_data=True)
                #chart.set_categories(time_col)


                # --- X-axis formatting: numeric hours ---
                #chart.x_axis.number_format = "0.##"

                # --- Tick spacing ---
                chart.x_axis.majorUnit = 12
                chart.x_axis.minorUnit = 3
                chart.x_axis.majorTickMark = "out"
                chart.x_axis.minorTickMark = "in"
                chart.x_axis.tickLblSkip = 1


                # Set the chart size
                chart.width = 30  # Width in terms of number of Excel columns
                chart.height = 20  # Height in terms of number of Excel rows

                # Add the chart to the sheet at a unique location
                chart_location = "J2"  # You can modify this as needed
                ws.add_chart(chart, chart_location)
                #print(openpyxl.__version__)

                #----------End Chart Creation---------
                #use update cursor to add arrival time to jointargets shapefile
                if add_data_to_points:
                    with arcpy.da.UpdateCursor(input_point_shapefile, [f"{plan_number}_ArrTm", f"{plan_number}_PkTim", f"{plan_number}_PkDif", "FID"]) as cursor2:
                        for row2 in cursor2:
                            fid_value = row2[3]
                            if fid_value == FID:
                                # Update arrival time, peak time, and peak difference fields
                                arrival_time_str = arrival_time_hours
                                peak_time_str = difference_hours_since
                                peak_difference_value = max_difference
                                row2[0] = str(arrival_time_str)
                                row2[1] = str(peak_time_str)
                                row2[2] = str(peak_difference_value)
                                cursor2.updateRow(row2)
                                print(f"Updated FID {FID} with Arrival Time: {arrival_time_str}, Peak Time: {peak_time_str}, Peak Difference: {peak_difference_value}")
                                break  # Exit loop after updating the matching FID



                # END POINT LOOP

        ws = wb["Summary"]
        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 10
        ws.cell(row=1, column=1, value="Total Points Processed")
        ws.cell(row=1, column=2, value=feature_count)
        ws.cell(row=3, column=1, value="Breach HDF File")
        ws.cell(row=4, column=1, value=breach_file_path)
        ws.cell(row=5, column=1, value=plan_shortid_value)
        ws.cell(row=7, column=1, value="NonBreach HDF File")
        ws.cell(row=8, column=1, value=nonbreach_file_path)
        ws.cell(row=9, column=1, value=nbplan_shortid_value)

        output_excel = fr'{output_folder}\{plan_shortid_value}_hydrographs.xlsx'

        # Save the workbook
        wb.save(output_excel)

        # Close the HDF5 file
        hdf_breach.close()
        hdf_nonbreach.close()

        if deletepoints:
            # Clean up interim join files
            if arcpy.Exists(points_join1_2djoin):
                arcpy.Delete_management(points_join1_2djoin)
            if arcpy.Exists(points_join2_xsjoin):    
                arcpy.Delete_management(points_join2_xsjoin)
            if arcpy.Exists(points_join3_1dsajoin):
                arcpy.Delete_management(points_join3_1dsajoin)

        #print end message
        print("Process Complete")

        # printout of time now or end time of script
        script_end_time = datetime.now()
        print(f"Script ended at: {(script_end_time).strftime('%Y-%m-%d %H:%M:%S')}")
        #printout of time it took to run the script
        script_run_time = script_end_time - script_start_time  # this is a datetime.timedelta
        total_seconds = int(script_run_time.total_seconds())
        hours, remainder = divmod(total_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        print(f"Total script run time: {hours}:{minutes:02d}:{seconds:02d}")

        return
    
#_________________________________________________________________________________________________
# ____________________________________________________________________________________________    
    
class Hdfplannames(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Get HDF Plan Information"
        self.description = "Does lots of stuff."
        self.canRunInBackground = True

    def getParameterInfo(self):
        """Define parameter definitions"""
        rasfolder = arcpy.Parameter(
            displayName="RAS or LifeSim Folder",
            name="RAS or LifeSim Folder",
            datatype="DEFolder",
            parameterType="Required",
            direction="Input")
        #rasfolder.value = fr'C:\~Kurt\~Projects\Buford_Dam_IES\2025_RAS\RAS_2025-10'

        lifesimfolder = arcpy.Parameter(
            displayName="Check if this is a LifeSim folder",
            name="Check if this is a LifeSim folder",
            datatype="GPBoolean",
            parameterType="Optional",
            direction="Input")
        #lifesimfolder.value = True

        damsaname = arcpy.Parameter(
            displayName="Dam Storage Area Name",
            name="Dam Storage Area Name",
            datatype="GPString",
            parameterType="Optional",
            direction="Input")
        damsaname.filter.type = "ValueList"
        damsaname.filter.list = []  # Will be populated in updateParameters
        
        parameters = [rasfolder, lifesimfolder, damsaname]
        #                0           1              2             3               4                 5                  6             7
        return parameters

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        if parameters[0].valueAsText:
            #get list of all .hdf files in the rasfolder
            import glob
            rasfolder = parameters[0].valueAsText
            storage_area_list_param = parameters[2]  # Adjust index if needed
            file_extension = "*p**.hdf"
            hdf_files = glob.glob(f"{rasfolder}\{file_extension}")
            if parameters[1].value == True:
                file_extension = "*.hdf"
                hdf_files = glob.glob(f"{rasfolder}\\**\\{file_extension}", recursive=True)

            storage_area_list = []
            storage_area_set = set()  # Set to keep track of unique names

            for file in hdf_files:
                # Open the HDF5 file
                with h5py.File(file, 'r') as hdf_file:
                    sa_attributes_table_key = f'/Geometry/Storage Areas/Attributes'
                    if sa_attributes_table_key in hdf_file:
                        storage_area_names = hdf_file.get(f'/Geometry/Storage Areas/Attributes')
                        saname_field = storage_area_names['Name']  # Access the "Name" field within the compound dataset
                        # Convert the field values to a list (assuming it's a single-column array)
                        saname_list = saname_field[:].tolist()  # Converts the dataset to a list of values
                        # Decode byte strings to regular strings
                        saname_list = [name.decode('utf-8') if isinstance(name, bytes) else name for name in saname_list]

                        # Check if all values are strings (after decoding)
                        if all(isinstance(name, str) for name in saname_list):
                            for name in saname_list:
                                if name not in storage_area_set:
                                    storage_area_list.append(name)
                                    storage_area_set.add(name)  # Add the name to the set to keep track of it

            storage_area_list_param.filter.list = storage_area_list


            # if it is a lifesim folder, do not look for p** and look in subfolders
            #if lifesim:
                #file_extension = "*.hdf"
                #hdf_files = glob.glob(f"{rasfolder}\\**\\{file_extension}", recursive=True)


        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return
       
    def execute(self, parameters, messages):
        """The source code of the tool."""
        #call parameters that were entered and create new variables as needed
        rasfolder = parameters[0].valueAsText
        lifesim = parameters[1].value
        storage_area_name = parameters[2].valueAsText
         
        get_storage_volume = True

        #get list of all .hdf files in the rasfolder
        import glob
        file_extension = "*p**.hdf"
        hdf_files = glob.glob(f"{rasfolder}\{file_extension}")

        # Exclude files that contain '.tmp.' in the filename
        hdf_files = [file for file in hdf_files if '.tmp' not in file]
        # Exclude files that contain 'Backup' in the filename
        hdf_files = [file for file in hdf_files if 'Backup' not in file]
              

        # if it is a lifesim folder, do not look for p** and look in subfolders
        if lifesim:
            file_extension = "*.hdf"
            hdf_files = glob.glob(f"{rasfolder}\\**\\{file_extension}", recursive=True)
    
        # Exclude files that contain 'Terrains' in the filename
        hdf_files = [file for file in hdf_files if 'Terrains' not in file]
        hdf_files = [file for file in hdf_files if 'Terrain' not in file]

        #messages.addMessage(f"HDF file list: {hdf_files}")

        # Create an Excel workbook
        wb = Workbook()
        ws = wb.create_sheet(title="RAS Plans")
        # Remove the default sheet created at Workbook instantiation
        default_sheet = wb['Sheet']
        wb.remove(default_sheet)
        ws = wb["RAS Plans"]

        ws.column_dimensions['A'].width = 10
        ws.column_dimensions['B'].width = 30
        ws.column_dimensions['C'].width = 20
        ws.column_dimensions['D'].width = 20
        ws.column_dimensions['E'].width = 20
        ws.column_dimensions['F'].width = 20
        ws.column_dimensions['G'].width = 20
        ws.column_dimensions['H'].width = 20
        ws.column_dimensions['I'].width = 20
        ws.column_dimensions['J'].width = 20
        ws.column_dimensions['K'].width = 20

        column_mod = 0
        if get_storage_volume and storage_area_name is not None:
            column_mod = 3

        ws.cell(row=1, column=1, value="Plan Number")
        ws.cell(row=1, column=2, value="Plan ShortID")
        ws.cell(row=1, column=column_mod + 3, value="Flow Title")
        ws.cell(row=1, column=column_mod + 4, value="Geometry Title")
        ws.cell(row=1, column=column_mod + 5, value="Terrain Filename")
        ws.cell(row=1, column=column_mod + 6, value="RAS Process Date")
        ws.cell(row=1, column=column_mod + 7, value="RAS Version")
        currentrow = 2

        # for SA name update QC
        storage_area_list = []
        storage_area_set = set()  # Set to keep track of unique names

        for file in hdf_files:
            # reset any variable parameters
            brch_struc_number = 0
            # Open the HDF5 file
            with h5py.File(file, 'r') as hdf_file:
                # Get the Plan Information table
                plan_info_table = hdf_file.get('/Plan Data/Plan Information')
                if plan_info_table is not None:
                    # Extract attributes
                    plan_shortid = plan_info_table.attrs.get('Plan ShortID', 'N/A')
                    plan_filename = plan_info_table.attrs.get('Plan Filename', 'N/A')
                    # If plan_filename is a bytes object, decode it
                    if isinstance(plan_filename, bytes):
                        plan_filename = plan_filename.decode('utf-8')  # Assuming UTF-8 encoding
                    plan_number = plan_filename.split('.')[-1]
                    plan_flow_title = plan_info_table.attrs.get('Flow Title', 'N/A')
                    plan_geometry_title = plan_info_table.attrs.get('Geometry Title', 'N/A')
                    # Print the extracted information
                    ws.cell(row=currentrow, column=1, value=plan_number)
                    ws.cell(row=currentrow, column=2, value=plan_shortid)
                    ws.cell(row=currentrow, column=column_mod + 3, value=plan_flow_title)
                    ws.cell(row=currentrow, column=column_mod + 4, value=plan_geometry_title)
                    messages.addMessage(f"Plan Number: {plan_number} is {plan_shortid}")
                    ishdfvalid = True
                else:
                    messages.addMessage(f"No Plan Information found in HDF file: {file}.")
                    ishdfvalid = False

                geometry_table = hdf_file.get('/Geometry')
                if geometry_table is not None:
                    # Extract attributes
                    terrain_filename = geometry_table.attrs.get('Terrain Filename', 'N/A')
                    # Print the extracted information
                    if ishdfvalid:
                        ws.cell(row=currentrow, column=column_mod + 5, value=terrain_filename)     
                else:
                    messages.addMessage(f"No Geometry table found in HDF file: {file}.")
                
                event_conditions_table = hdf_file.get('/Event Conditions')
                if event_conditions_table is not None:
                    # Extract attributes
                    date_processed = event_conditions_table.attrs.get('Date Processed', 'N/A')
                    # Print the extracted information
                    ws.cell(row=currentrow, column=column_mod + 6, value=date_processed)     
                else:
                    messages.addMessage(f"No Event Conditions table found in HDF file: {file}.")

                # Access main HDF5 file attributes
                ras_fileversion = hdf_file.attrs.get('File Version')
                if isinstance(ras_fileversion, bytes):
                    ras_fileversion = ras_fileversion.decode('ascii', errors='ignore').rstrip('\x00').strip()
                # Print the extracted information
                ws.cell(row=currentrow, column=column_mod + 7, value=ras_fileversion) 
                #messages.addMessage(f"RAS Version: {ras_fileversion}")

                if get_storage_volume and storage_area_name is not None:
                    #storage_area_name = "Mojave River Res"
                    
                    # Check if the sa_volume_table exists in the HDF5 file
                    sa_variable_table_key = f'/Results/Unsteady/Output/Output Blocks/Base Output/Unsteady Time Series/Storage Areas/{storage_area_name}/Storage Area Variables'
                    messages.addMessage(f"SA Variable Table Key: {sa_variable_table_key}")
                    
                    if sa_variable_table_key in hdf_file:
                        stage_col = 4
                        startstage_col = 3
                        vol_col = 5
                        ws.cell(row=1, column=vol_col, value="Max Volume ac-ft")
                        ws.cell(row=1, column=stage_col, value="Max Stage ft")
                        ws.cell(row=1, column=startstage_col, value="Start Stage ft")
                        # If the table exists, proceed with the extraction
                        sa_variable_table = hdf_file[sa_variable_table_key]
                        #messages.addMessage(f"SA Vol Table: {sa_volume_table}")
                        
                        sa_stage_index = 0
                        sa_volume_index = 5
                        if sa_variable_table is not None:
                            sa_volume_values = sa_variable_table[:, sa_volume_index]
                            sa_stage_values = sa_variable_table[:, sa_stage_index]
                            
                            # Select the maximum value from sa_volume_values
                            max_volume = np.max(sa_volume_values)  # This gives the maximum value in the array
                            cell=ws.cell(row=currentrow, column=vol_col, value=max_volume)
                            cell.number_format = '#,##0'  # Set the desired number format
                            # Select the maximum value from sa_stage_values
                            max_stage = np.max(sa_stage_values)  # This gives the maximum value in the array
                            cell=ws.cell(row=currentrow, column=stage_col, value=max_stage)
                            cell.number_format = '#,##0.0'  # Set the desired number format
                            start_stage = sa_stage_values[0]  # This gives the first value in the array
                            cell=ws.cell(row=currentrow, column=startstage_col, value=start_stage)
                            cell.number_format = '#,##0.0'  # Set the desired number format


                    #else:
                        # QC If the table doesn't exist, log a message or take any other action
                        #messages.addMessage(f"SA Volume Table for '{storage_area_name}' not found.")

                    ## END storage area data lookup

                    get_breach_data = True
                    if get_breach_data:
                        # Check if the sa_volume_table exists in the HDF5 file
                        breach_data_table_key = f'/Plan Data/Breach Data/Names'
                        messages.addMessage(f"Breach Data Table Key: {breach_data_table_key}")
                        
                        if breach_data_table_key in hdf_file:
                            breach_status = "Breach"
                            breach_data_name_table = hdf_file[breach_data_table_key]
                            breach_data_name_values = breach_data_name_table[:]
                            #messages.addMessage(f"Breach Data Name Values: {breach_data_name_values}")
                            parsed_names = []
                            start_col = column_mod + 8

                            for item in breach_data_name_values:
                                brch_struc_number += 1
                                # Decode to string if needed
                                text = item.decode('utf-8') if isinstance(item, bytes) else str(item)

                                # Split on vertical bar
                                parts = text.split("|")  
                                parsed_names.append(parts)
                                
                                # Split on the FIRST |, brst=breach structure
                                brst_type, brst_name = text.split("|", 1)
                                # Replace remaining "|" with spaces
                                brst_name = brst_name.replace("|", " ")

                                # rest contains everything after the first bar, with bars removed except those originally in the string
                                messages.addMessage(f"Before first split: {brst_type}")
                                messages.addMessage(f"After first split: {brst_name}")

                                if brst_type == 'Inline Structure': #parameters from lake Drummond
                                    brst_type = 'Inline Structures'
                                    hw_stage_index = 1
                                    total_flow_index = 0
                                    weir_flow_index = 3
                                    bottom_width_index = 0
                                    bottom_elev_index = 1
                                    left_sideslope_index = 2
                                    right_sideslope_index = 3
                                    breach_flow_index = 4
                                    breach_velocity_index = 5
                                if brst_type == 'Connection': #parameters from Buford
                                    brst_type = 'SA 2D Area Conn'
                                    hw_stage_index = 2
                                    total_flow_index = 0
                                    weir_flow_index = 1
                                    bottom_width_index = 2
                                    bottom_elev_index = 3
                                    left_sideslope_index = 4
                                    right_sideslope_index = 5
                                    breach_flow_index = 6
                                    breach_velocity_index = 7

                                
                                breaching_variable_table_key = f'/Results/Unsteady/Output/Output Blocks/DSS Hydrograph Output/Unsteady Time Series/{brst_type}/{brst_name}/Breaching Variables'
                                if breaching_variable_table_key in hdf_file:
                                    breach_variable_table = hdf_file.get(breaching_variable_table_key)
                                    # breaching variables SA 2D Area Conn variables: StageHW[0], StageTW[1], BottomWidth[2], BottomElev[3], LeftSideSlope[4], RightSideSLope[5], BreachFlow[6], BreachVelocity[7]
                                    # breaching variables Inline Structure variables: BottomWidth[0], BottomElev[1], LeftSideSlope[2], RightSideSLope[3], BreachFlow[4], BreachVelocity[5]

                                    breach_time = breach_variable_table.attrs.get('Breach at', 'N/A')
                                    breach_time_days = breach_variable_table.attrs.get('Breach at Time (Days)', 'N/A')
                                    breach_time_hours = breach_time_days * 24
                                    messages.addMessage(f"Breach time: {breach_time}")
                                    messages.addMessage(f"Breach time days: {breach_time_days}")

                                    def _safe_agg(values, mode='max'):
                                        arr = np.asarray(values, dtype=float)
                                        fin = arr[np.isfinite(arr)]
                                        if fin.size == 0:
                                            return None
                                        return float(np.nanmax(fin) if mode == 'max' else np.nanmin(fin))

                                    bottom_width_values = breach_variable_table[:, bottom_width_index]
                                    max_bottom_width = _safe_agg(bottom_width_values, 'max')
                                    messages.addMessage(f"Max Bottom Width: {max_bottom_width if max_bottom_width is not None else 'N/A'}")

                                    bottom_elev_values = breach_variable_table[:, bottom_elev_index]
                                    min_bottom_elev = _safe_agg(bottom_elev_values, 'min')
                                    messages.addMessage(f"Min Bottom Elev: {min_bottom_elev if min_bottom_elev is not None else 'N/A'}")

                                    left_sideslope_values = breach_variable_table[:, left_sideslope_index]
                                    max_left_sideslope = _safe_agg(left_sideslope_values, 'max')
                                    messages.addMessage(f"Max Left SideSlope: {max_left_sideslope if max_left_sideslope is not None else 'N/A'}")

                                    right_sideslope_values = breach_variable_table[:, right_sideslope_index]
                                    max_right_sideslope = _safe_agg(right_sideslope_values, 'max')
                                    messages.addMessage(f"Max Right SideSlope: {max_right_sideslope if max_right_sideslope is not None else 'N/A'}")

                                    breach_flow_values = breach_variable_table[:, breach_flow_index]
                                    max_breach_flow = _safe_agg(breach_flow_values, 'max')
                                    messages.addMessage(f"Max BreachFlow: {max_breach_flow if max_breach_flow is not None else 'N/A'}")

                                    breach_velocity_values = breach_variable_table[:, breach_velocity_index]
                                    max_breach_velocity = _safe_agg(breach_velocity_values, 'max')
                                    messages.addMessage(f"Max Breach Velocity: {max_breach_velocity if max_breach_velocity is not None else 'N/A'}")

                                    breach_structure_variables_table_key = f'/Results/Unsteady/Output/Output Blocks/Base Output/Unsteady Time Series/{brst_type}/{brst_name}/Structure Variables'
                                    breach_structure_variables_table = hdf_file[breach_structure_variables_table_key]                                  
    
                                    hw_stage_values = breach_structure_variables_table[:, hw_stage_index]
                                    max_hw_stage = np.max(hw_stage_values)  # This gives the maximum value in the array
                                    messages.addMessage(f"Max HW Stage: {max_hw_stage}")
                                    
                                    total_flow_values = breach_structure_variables_table[:, total_flow_index]
                                    max_total_flow = np.max(total_flow_values)  # This gives the maximum value in the array
                                    messages.addMessage(f"Max Total Flow: {max_total_flow}")
                                    
                                    weir_flow_values = breach_structure_variables_table[:, weir_flow_index]
                                    max_weir_flow = np.max(weir_flow_values)  # This gives the maximum value in the array
                                    messages.addMessage(f"Max Weir Flow: {max_weir_flow}")
                               
                                    # Note current column would be staring on column 10, comes from start_col
                                    ws.cell(row=1, column=start_col, value=f"Breach Structure {brch_struc_number}")
                                    cell=ws.cell(row=currentrow, column=start_col, value=brst_name)
                                    ws.cell(row=1, column=start_col+1, value=f"Breach Time {brch_struc_number}")
                                    cell=ws.cell(row=currentrow, column=start_col+1, value=breach_time)
                                    ws.cell(row=1, column=start_col+2, value=f"Breach Time Hours {brch_struc_number}")
                                    cell=ws.cell(row=currentrow, column=start_col+2, value=breach_time_hours)
                                    cell.number_format = '#,##0.00'  # Set the desired number format
                                    ws.cell(row=1, column=start_col+3, value=f"Max HW Stage {brch_struc_number}")
                                    cell=ws.cell(row=currentrow, column=start_col+3, value=max_hw_stage)
                                    cell.number_format = '#,##0.00'  # Set the desired number format

                                    ws.cell(row=1, column=start_col+4, value=f"Max Bottom Width {brch_struc_number}")
                                    cell=ws.cell(row=currentrow, column=start_col+4, value=max_bottom_width)
                                    cell.number_format = '#,##0.00'  # Set the desired number format
                                    ws.cell(row=1, column=start_col+5, value=f"Min Bottom Elev {brch_struc_number}")
                                    cell=ws.cell(row=currentrow, column=start_col+5, value=min_bottom_elev)
                                    cell.number_format = '#,##0.00'  # Set the desired number format
                                    ws.cell(row=1, column=start_col+6, value=f"Max Left Side Slope {brch_struc_number}")
                                    cell=ws.cell(row=currentrow, column=start_col+6, value=max_left_sideslope)
                                    cell.number_format = '#,##0.00'  # Set the desired number format
                                    ws.cell(row=1, column=start_col+7, value=f"Max Right Side Slope {brch_struc_number}")
                                    cell=ws.cell(row=currentrow, column=start_col+7, value=max_right_sideslope)
                                    cell.number_format = '#,##0.00'  # Set the desired number format
                                    ws.cell(row=1, column=start_col+8, value=f"Max Breach Flow {brch_struc_number}")
                                    cell=ws.cell(row=currentrow, column=start_col+8, value=max_breach_flow)
                                    cell.number_format = '#,##0'  # Set the desired number format
                                    ws.cell(row=1, column=start_col+9, value=f"Max Breach Velocity {brch_struc_number}")
                                    cell=ws.cell(row=currentrow, column=start_col+9, value=max_breach_velocity)
                                    cell.number_format = '#,##0.00'  # Set the desired number format


                                    ws.cell(row=1, column=start_col+10, value=f"Max Total Flow {brch_struc_number}")
                                    cell=ws.cell(row=currentrow, column=start_col+10, value=max_total_flow)
                                    cell.number_format = '#,##0'  # Set the desired number format
                                    ws.cell(row=1, column=start_col+11, value=f"Max Weir Flow {brch_struc_number}")
                                    cell=ws.cell(row=currentrow, column=start_col+11, value=max_weir_flow)
                                    cell.number_format = '#,##0'  # Set the desired number format

                                    start_col += 12

                            
                            #messages.addMessage(f"Parsed Breach Parts: {parsed_names}")
                            
                        else:
                            if ishdfvalid:
                                breach_status = "No Breach"
                                cell=ws.cell(row=currentrow, column=column_mod + 8, value=breach_status)
                        if ishdfvalid:
                            messages.addMessage(f"Breach or NonBreach: {breach_status}")

                if ishdfvalid:
                    currentrow += 1

            # END for file in hdf file loop
        # set first row of excel table to hieght 30 and wrap text
        ws.row_dimensions[1].height = 30
        for cell in ws[1]:
            cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
        # set cell widths of column L to V at 12
        for col in ("L","M","N","O","P","Q","R","S","T","U","V", "W", "X", "Y", "Z"):
            ws.column_dimensions[col].width = 12
        ws.freeze_panes = "C2"  # Freeze the first row

        messages.addMessage(f"Final SA Set: {storage_area_set}")
        messages.addMessage(f"Final SA List: {storage_area_list}")

        output_excel = f'{rasfolder}/~RAS_Plan_Names_and_Numbers.xlsx'
        # Save the workbook
        try:  #Add a try-except block for save the file at the end
            wb.save(output_excel)
            messages.addMessage(f"Excel file saved to: {output_excel}") #Add success message
        except Exception as e:
            messages.addErrorMessage(f"Error saving Excel file: {e}")
        

        return




        