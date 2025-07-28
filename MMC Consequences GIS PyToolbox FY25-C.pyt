# -*- coding: utf-8 -*-
#-------------------------------------------------------------------------
# Author:     Kurt Buchanan, CELRH
# Created:     May  2023
# ArcGIS Version:   2.9.3
# Requirements: Spatial Analyst extension
# Revisions: 
#FY23-C added arrival time zone tool and raster to polygon tool
#FY23-E removed sysargv, added snap to exclusion polygon tool
#FY23-F added geopackage output to NSI api tool, also breakout by polygon to support FFRD
#FY23-G switched api NSI importer to convert json to geopackage, then to shapefile to support ArcGIS 3.x, added time stamps
#FY24-A 10-16-2023 notes: overwrite projection option on projection tool, default 2021 to 2024 index set at 1.166
#FY24-B 3/1/2024: added ssl ignore to nsi api tool, also switched conversion from geopackage to shapefile method (using Feature Class to Shapefile tool) to account for ArcGIS 3.2 making ObjectID field in 64-bit
#FY24-C 4/15/2024: fixed error on internal name by adding outputlayername, added in -20 foot buffer via clearbuffer
#FY25-A 10/31/2024: added FY25 index of 1.177 to NSI tool, also incorporates some name changes to geopackage split option
#FY25-B 05/09/2025: added new tool to calculate incremental lifesim results
#-------------------------------------------------------------------------
from ast import Return
from tokenize import Number
import arcpy
import os
import urllib.parse, urllib.request
#added below acrpy.sa for reclassify grid tool "con" tool
from arcpy.sa import *
from datetime import datetime
import ssl

class Toolbox(object):
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the
        .pyt file)."""
        self.label = "MMC Toolbox"
        self.alias = "mmctoolbox"
        self.description = "ArcGIS python toolbox created to support the USACE MMC Consequences Team."

        # List of tool classes associated with this toolbox
        self.tools = [Nsicreator, Bufferandsimplifypoly, Polygonstodoublewarning, Nsicreatorapi, Reclassifytozone, Rastertopolygon, Snaptoexclusionpoly, Batchdefineprojection, Incrementallifesimresults]

##-------------------------------------------------------------------------------------

class Nsicreator(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "NSI Creator (from local shapefiles)"
        self.description = "Clips a county shapefile to a study area, then creates a list of all the counties and copies the NSI files of each county from a local folder. Merges all files into a single output."
        self.canRunInBackground = True

    def getParameterInfo(self):
        """Define parameter definitions"""
        inputstudyarea = arcpy.Parameter(
            displayName="Study Area",
            name="study area",
            datatype="GPFeatureLayer",
            parameterType="Required",
            direction="Input")
        workspace = arcpy.Parameter(
            displayName="Output Folder",
            name="output folder",
            datatype="DEFolder",
            parameterType="Required",
            direction="Input")
        inputuscounties = arcpy.Parameter(
            displayName="US Counties Shapefile",
            name="us counties shapefile",
            datatype="GPFeatureLayer",
            parameterType="Required",
            direction="Input")
        inputnsifolder = arcpy.Parameter(
            displayName="Folder with Combined NSI Files",
            name="folder with NSI files",
            datatype="DEFolder",
            parameterType="Required",
            direction="Input")
        priceindex = arcpy.Parameter(
            displayName="Price Index from Base (default: 1)",
            name="price index",
            datatype="GPDouble",
            parameterType="Required",
            direction="Input")
        priceindex.value = 1.177
        outputname = arcpy.Parameter(
            displayName="Output NSI File Name (no extension, include price level)",
            name="output filename",
            datatype="GPString",
            parameterType="Required",
            direction="Output")
        outputname.value = "NSI2022_2025pricelevel"
        addfianame = arcpy.Parameter(
            displayName="Check to add name field for FIA",
            name="check to add fianame",
            datatype="GPBoolean",
            parameterType="Optional",
            direction="Input")
        
        parameters = [inputstudyarea, workspace, inputuscounties, inputnsifolder, priceindex, outputname, addfianame]
        return parameters

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return
       
    def execute(self, parameters, messages):
        """The source code of the tool."""
        #call parameters that were entered and create new variables as needed
        inputstudyarea = parameters[0].valueAsText
        inputuscounties = parameters[2].valueAsText
        inputnsifolder = parameters[3].valueAsText
        workspace = parameters[1].valueAsText
        outputname = parameters[5].valueAsText
        priceindex = parameters[4].valueAsText
        #boolean parameters should just be the value, not text
        addfianame = parameters[6].value
        outputnsifile = fr"{workspace}\{outputname}.shp"

        #Clip US County shapefile to study area
        messages.addMessage("Clipping county shapefile to input study area")
        arcpy.analysis.PairwiseClip(
            in_features=inputuscounties, 
            clip_features=inputstudyarea, 
            out_feature_class=fr"{workspace}\ClippedUSCounties.shp", 
            cluster_tolerance="")
        
        #set clipped counties as a parameter
        ClippedUSCounties=fr"{workspace}\ClippedUSCounties.shp"
        
        #create list of all the FIPS county numbers in the clipped shapefile
        FipsList = [row[0] for row in arcpy.da.SearchCursor(ClippedUSCounties, ["FIPS"])]
        
        #count number of entries and set initial loop number
        cntTotal = len(FipsList)
        messages.addMessage("There are {0} counties in the study area".format(cntTotal))
        loopNumber = 1

        #create output list to append each county NSI chapefile to
        outputshapefilelist=[]
        
        #Iterate over the list of FIPS codes, clipping each NSI county file to the study area
        messages.addMessage("Clipping the NSI file of each FIPS number found in the clipped county shapefile to the Study Area...")
        for FIPS_Value in FipsList:
            messages.addMessage("County {0} of {1}".format(loopNumber, cntTotal))
            loopNumber += 1
            print(FIPS_Value," has a field in the county shapefile")
            FIPS_Value_shps = inputnsifolder + fr"\{FIPS_Value}.shp"
            print(FIPS_Value_shps," has a shapefile in the NSI folder")
            clippednsicounties = fr"{workspace}\{FIPS_Value}.shp"
            outputshapefilelist.append(clippednsicounties)   
            print(clippednsicounties," has been clipped to the study area")
            arcpy.analysis.PairwiseClip(
                in_features=FIPS_Value_shps, 
                clip_features=inputstudyarea, 
                out_feature_class=clippednsicounties, 
                cluster_tolerance="")
            messages.addMessage(str(FIPS_Value)+" has been clipped to the study area")

        #add merge of the list, stringlist was if the merge tool needed a string instead of an array
        #stringlist="'" +','.join(outputshapefilelist).replace(",","','")+"'"
        arcpy.management.Merge(outputshapefilelist,outputnsifile)
        
        #deletes the intermediate files that were just merged into a single result
        arcpy.management.Delete(outputshapefilelist)
        messages.addMessage("Merged county files into final output file, deleted individual county files")
        
        #index value fields using calculate field
        messages.addMessage("Indexing structure value, content value, and vehicle value")
        arcpy.management.CalculateField(in_table=outputnsifile, 
            field="Val_Struct", 
            expression=f"!Val_Struct! * {priceindex}", 
            expression_type="PYTHON_9.3", 
            code_block="", 
            field_type="TEXT", 
            enforce_domains="NO_ENFORCE_DOMAINS")

        arcpy.management.CalculateField(in_table=outputnsifile, 
            field="Val_Cont", 
            expression=f"!Val_Cont! * {priceindex}", 
            expression_type="PYTHON_9.3", 
            code_block="", 
            field_type="TEXT", 
            enforce_domains="NO_ENFORCE_DOMAINS")

        arcpy.management.CalculateField(in_table=outputnsifile, 
            field="Val_Vehic", 
            expression=f"!Val_Vehic! * {priceindex}", 
            expression_type="PYTHON_9.3", 
            code_block="", 
            field_type="TEXT", 
            enforce_domains="NO_ENFORCE_DOMAINS")

        #If boolean is checked ST_Name field will be added and calculated, FIA requires a unique name field
        if addfianame:
            messages.addMessage("Adding St_Name field and setting it equal to OccTyp + field ID")
            arcpy.management.AddField(in_table=outputnsifile, field_name="ST_Name", field_type="TEXT", 
                field_precision=None, field_scale=None, field_length=50, field_alias="", field_is_nullable="NULLABLE", 
                field_is_required="NON_REQUIRED", field_domain="")

            arcpy.management.CalculateField(in_table=outputnsifile, field="ST_Name", expression="!OccType! + ' ' + str(!FID!)", 
                expression_type="PYTHON_9.3", code_block="", field_type="TEXT", enforce_domains="NO_ENFORCE_DOMAINS")

        return

##-------------------------------------------------------------------------------------

class Bufferandsimplifypoly(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Buffer and Simplify a Polygon"
        self.description = ""
        self.canRunInBackground = True

    def getParameterInfo(self):
        """Define parameter definitions"""
        inputpolygon = arcpy.Parameter(
            displayName="Input Polygon",
            name="input_polygon",
            datatype="GPFeatureLayer",
            parameterType="Required",
            direction="Input")
        outputpolygon = arcpy.Parameter(
            displayName="Output Polygon",
            name="output_polygon",
            datatype="GPFeatureLayer",
            parameterType="Required",
            direction="Output")
        bufferdistance = arcpy.Parameter(
            displayName="Buffer Distance in Feet",
            name="buffer_distance",
            datatype="GPLong",
            parameterType="Required",
            direction="Input")
        parameters = [inputpolygon, outputpolygon, bufferdistance]
        return parameters

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return
       
    def execute(self, parameters, messages):
        """The source code of the tool."""
        #call parameters that were entered and create new variables as needed
        inputpolygon = parameters[0].valueAsText
        outputpolygon = parameters[1].valueAsText
        bufferdistance = parameters[2].valueAsText
        interimpolygon = fr"C:\Temp\bufferedpoly.shp"

        #run pairwise buffer that also simplifies using a deviation of half the buffer value
        with arcpy.EnvManager(parallelProcessingFactor="100"):
            arcpy.analysis.PairwiseBuffer(in_features=inputpolygon, 
                out_feature_class=interimpolygon, 
                buffer_distance_or_field=bufferdistance +" Feet", 
                dissolve_option="NONE", 
                dissolve_field=[], 
                method="PLANAR", 
                max_deviation=str(int(bufferdistance) // 2) +" Feet")
        messages.addMessage("Buffered polygon, saved to C:\Temp")
        
        ##below is a simplification after running buffer using 1/4 of the the buffer as tolerance
        arcpy.cartography.SimplifyPolygon(in_features=interimpolygon, 
            out_feature_class=outputpolygon, 
            algorithm="POINT_REMOVE", 
            tolerance=str(int(bufferdistance) // 4) +" Feet", 
            minimum_area="0 SquareFeet", 
            error_option="RESOLVE_ERRORS", 
            collapsed_point_option="NO_KEEP", 
            in_barriers=[inputpolygon])
        messages.addMessage("Simplified polygon")

        # Dissolve to remove holes/gaps
        dissolve = True
        if dissolve:
            dissolved_output = outputpolygon.replace('.shp', '_dissolved.shp')
            arcpy.management.EliminatePolygonPart(
                in_features=outputpolygon,
                out_feature_class=dissolved_output,
                condition="PERCENT",
                part_area="0 SquareFeetUS",
                part_area_percent=25,
                part_option="CONTAINED_ONLY"
            )
            arcpy.management.Delete(outputpolygon)
            arcpy.management.Rename(dissolved_output, outputpolygon)
            messages.addMessage("Ran the eliminate tool to remove interior holes")
            
        arcpy.management.Delete(interimpolygon)
        messages.addMessage("Deleted interim buffered polygon from C:\Temp")

        return

##-------------------------------------------------------------------------------------

class Polygonstodoublewarning(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Polygons to Double Warning EPZ"
        self.description = ""
        self.canRunInBackground = True

    def getParameterInfo(self):
        """Define parameter definitions"""
        inputstudyarea = arcpy.Parameter(
            displayName="Input Study Area Polygon",
            name="study area",
            datatype="GPFeatureLayer",
            parameterType="Required",
            direction="Input")
        inputinpoolpolygon = arcpy.Parameter(
            displayName="Input InPool Area Polygon",
            name="inpool area",
            datatype="GPFeatureLayer",
            parameterType="Required",
            direction="Input")
        inputnonbreach1 = arcpy.Parameter(
            displayName="Input Non-Breach Polygon 1",
            name="input nonbreach 1",
            datatype="GPFeatureLayer",
            parameterType="Required",
            direction="Input")
        epzfolder = arcpy.Parameter(
            displayName="Output Folder",
            name="output folder",
            datatype="DEFolder",
            parameterType="Required",
            direction="Input")
        inputnonbreach2 = arcpy.Parameter(
            displayName="Input Non-Breach Polygon 2 (optional)",
            name="input nonbreach 2",
            datatype="GPFeatureLayer",
            parameterType="Optional",
            direction="Input")
        scenario1name = arcpy.Parameter(
            displayName="Output Name for Scenario 1 (no extension)",
            name="scenario name 1",
            datatype="GPString",
            parameterType="Required",
            direction="Input")
        scenario1name.value = "MHP_DoubleWarn_EPZ"
        scenario2name = arcpy.Parameter(
            displayName="Output Name for Scenario 2 (optional, no extension)",
            name="scenario name 2",
            datatype="GPString",
            parameterType="Optional",
            direction="Input")
        parameters = [inputstudyarea, inputinpoolpolygon, inputnonbreach1, epzfolder, scenario1name, inputnonbreach2, scenario2name]
        return parameters

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return
       
    def execute(self, parameters, messages):
        """The source code of the tool."""
        #call parameters that were entered and create new variables as needed
        inputstudyarea = parameters[0].valueAsText
        inputinpoolpolygon = parameters[1].valueAsText
        inputnonbreach1 = parameters[2].valueAsText
        epzfolder = parameters[3].valueAsText
        scenario1name = parameters[4].valueAsText
        inputnonbreach2 = parameters[5].valueAsText
        scenario2name = parameters[6].valueAsText
        studyareadissolve = fr"{epzfolder}\studyareadissolve.shp"
        nonbreach1dissolve = fr"{epzfolder}\nonbreach1dissolve.shp"
        nonbreach2dissolve = fr"{epzfolder}\nonbreach2dissolve.shp"
        inpooldissolve = fr"{epzfolder}\inpooldissolve.shp"
        scenario1union = fr"{epzfolder}\scenario1union.shp"
        scenario2union = fr"{epzfolder}\scenario2union.shp"

        # Process: Pairwise Dissolve Study Area (Pairwise Dissolve) (analysis)
        arcpy.analysis.PairwiseDissolve(in_features=inputstudyarea, 
            out_feature_class=studyareadissolve, dissolve_field=[], statistics_fields=[], multi_part="MULTI_PART")
        messages.addMessage("Dissolved Study Area")

        # Process: Add Field Fname Study Area (Add Field) (management)
        arcpy.management.AddField(in_table=studyareadissolve, field_name="Fname", 
            field_type="TEXT", field_precision=None, field_scale=None, field_length=15, field_alias="", 
            field_is_nullable="NULLABLE", field_is_required="NON_REQUIRED", field_domain="")[0]
        messages.addMessage("Added Fname field to Study Area")

        # Process: Calculate Field Fname Study Area (Calculate Field) (management)
        arcpy.management.CalculateField(in_table=studyareadissolve, field="Fname", 
            expression="'Breach_EPZ'", expression_type="PYTHON_9.3", code_block="", 
            field_type="TEXT", enforce_domains="NO_ENFORCE_DOMAINS")[0]
        messages.addMessage("Calculated Fname field in Study Area")

        # Process: Pairwise Dissolve NF-1 (Pairwise Dissolve) (analysis)
        arcpy.analysis.PairwiseDissolve(in_features=inputnonbreach1, 
            out_feature_class=nonbreach1dissolve, dissolve_field=[], statistics_fields=[], multi_part="MULTI_PART")
        messages.addMessage("Dissolved NF-1")

        # Process: Add Field NFname NF-1 (Add Field) (management)
        arcpy.management.AddField(in_table=nonbreach1dissolve, field_name="NFname", field_type="TEXT", 
            field_precision=None, field_scale=None, field_length=15, field_alias="", field_is_nullable="NULLABLE", field_is_required="NON_REQUIRED", field_domain="")[0]
        messages.addMessage("Added Fname field to NF-1")

        # Process: Calculate Field NFname NF-1 (Calculate Field) (management)
        arcpy.management.CalculateField(in_table=nonbreach1dissolve, field="NFname", expression="'Non_Breach_EPZ'", 
            expression_type="PYTHON_9.3", code_block="", field_type="TEXT", enforce_domains="NO_ENFORCE_DOMAINS")[0]
        messages.addMessage("Calculated Fname field in NF-1")

        # Process: Pairwise Dissolve InPool (Pairwise Dissolve) (analysis)
        arcpy.analysis.PairwiseDissolve(in_features=inputinpoolpolygon, out_feature_class=inpooldissolve, 
            dissolve_field=[], statistics_fields=[], multi_part="MULTI_PART")
        messages.addMessage("Dissolved InPool Area")

        # Process: Add Field IPname InPool (Add Field) (management)
        arcpy.management.AddField(in_table=inpooldissolve, field_name="IPname", field_type="TEXT", 
            field_precision=None, field_scale=None, field_length=15, field_alias="", field_is_nullable="NULLABLE", field_is_required="NON_REQUIRED", field_domain="")[0]
        messages.addMessage("Added Ipname field to InPool")

        # Process: Calculate Field IPname InPool (Calculate Field) (management)
        arcpy.management.CalculateField(in_table=inpooldissolve, field="IPname", expression="'InPool_EPZ'", 
            expression_type="PYTHON_9.3", code_block="", field_type="TEXT", enforce_domains="NO_ENFORCE_DOMAINS")[0]
        messages.addMessage("Calculated Ipname field to InPool")

        # Process: Union Scenario 1 (Union) (analysis)
        arcpy.analysis.Union(in_features=[[studyareadissolve, ""], [nonbreach1dissolve, ""], [inpooldissolve, ""]], 
            out_feature_class=scenario1union, join_attributes="ALL", cluster_tolerance="", gaps="GAPS")
        messages.addMessage("Unioned Study Area, InPool, and NF-1")

        # Process: Add Field Zone Scenario 1 (Add Field) (management)
        arcpy.management.AddField(in_table=scenario1union, field_name="Zone", field_type="TEXT", field_precision=None, 
            field_scale=None, field_length=15, field_alias="", field_is_nullable="NULLABLE", field_is_required="NON_REQUIRED", field_domain="")[0]
        messages.addMessage("Added Zone field to unioned scenario 1")

        # Process: Calculate Field Zones Scenario 1 (Calculate Field) (management)
        arcpy.management.CalculateField(in_table=scenario1union, field="Zone", 
            expression="Rename(!Fname!, !NFname!, !IPname!)", expression_type="PYTHON_9.3", code_block="""def Rename(fname, nfname, ipname):
            if ipname == 'InPool_EPZ':
                return 'InPool_EPZ'
            elif nfname == 'Non_Breach_EPZ':
                return 'Non_Breach_EPZ'
            else:
                return 'Breach_EPZ'""", field_type="TEXT", enforce_domains="NO_ENFORCE_DOMAINS")[0]
        messages.addMessage("Calculated Zone field in unioned scenario 1")

         # Process: Pairwise Dissolve Scenario 1 (Pairwise Dissolve) (analysis)
        arcpy.analysis.PairwiseDissolve(in_features=scenario1union, out_feature_class=fr"{epzfolder}\{scenario1name}.shp", dissolve_field=["Zone"], 
            statistics_fields=[], multi_part="MULTI_PART")
        messages.addMessage("Dissolved by Zone field in unioned scenario 1 to create scenario 1 EPZ output")

        # Start Optional second scenario
        # Process: Pairwise Dissolve NF-2 (Pairwise Dissolve) (analysis)
        if inputnonbreach2:
            arcpy.analysis.PairwiseDissolve(in_features=inputnonbreach2, out_feature_class=nonbreach2dissolve, dissolve_field=[], 
                statistics_fields=[], multi_part="MULTI_PART")
            messages.addMessage("Dissolved Optional NF-2")

            # Process: Add Field NFname NF-2 (Add Field) (management)
            arcpy.management.AddField(in_table=nonbreach2dissolve, field_name="NFname", field_type="TEXT", field_precision=None, 
                field_scale=None, field_length=15, field_alias="", field_is_nullable="NULLABLE", field_is_required="NON_REQUIRED", field_domain="")[0]
            messages.addMessage("Add NFname field to Optional NF-2")

            # Process: Calculate Field NFname NF-2 (Calculate Field) (management)
            arcpy.management.CalculateField(in_table=nonbreach2dissolve, field="NFname", expression="'Non_Breach_EPZ'", 
                expression_type="PYTHON_9.3", code_block="", field_type="TEXT", enforce_domains="NO_ENFORCE_DOMAINS")[0]
            messages.addMessage("Calculated NFname field in Optional NF-2")

            # Process: Union Scenario 2 (Union) (analysis)
            arcpy.analysis.Union(in_features=[[studyareadissolve, ""], [nonbreach2dissolve, ""], [inpooldissolve, ""]], 
                out_feature_class=scenario2union, join_attributes="ALL", cluster_tolerance="", gaps="GAPS")
            messages.addMessage("Unioned Study Area, InPool, and NF-2")

            # Process: Add Field Zone Scenario 2 (Add Field) (management)
            arcpy.management.AddField(in_table=scenario2union, field_name="Zone", field_type="TEXT", field_precision=None, 
                field_scale=None, field_length=15, field_alias="", field_is_nullable="NULLABLE", field_is_required="NON_REQUIRED", field_domain="")[0]
            messages.addMessage("Added Zone field to unioned scenario 2")

            # Process: Calculate Field Zones Scenario 2 (Calculate Field) (management)
            arcpy.management.CalculateField(in_table=scenario2union, field="Zone", 
            expression="Rename(!Fname!, !NFname!, !IPname!)", expression_type="PYTHON_9.3", code_block="""def Rename(fname, nfname, ipname):
            if ipname == 'InPool_EPZ':
                return 'InPool_EPZ'
            elif nfname == 'Non_Breach_EPZ':
                return 'Non_Breach_EPZ'
            else:
                return 'Breach_EPZ'""", field_type="TEXT", enforce_domains="NO_ENFORCE_DOMAINS")[0]
            messages.addMessage("Calculated Zone field to unioned scenario 2")

            # Process: Pairwise Dissolve Scenario 2 (Pairwise Dissolve) (analysis)
            arcpy.analysis.PairwiseDissolve(in_features=scenario2union, out_feature_class=fr"{epzfolder}\{scenario2name}.shp", dissolve_field=["Zone"], statistics_fields=[], multi_part="MULTI_PART")
            messages.addMessage("Dissolved by Zone field in unioned scenario 2 to create scenario 2 EPZ output")

            # Delete scenario 2 interim files
            arcpy.management.Delete(nonbreach2dissolve)
            arcpy.management.Delete(scenario2union)
            messages.addMessage("Deleted interim scenario 2 files")
    
        # Delete scenario 1 interim files
        arcpy.management.Delete(studyareadissolve)
        arcpy.management.Delete(inpooldissolve)
        arcpy.management.Delete(nonbreach1dissolve)
        arcpy.management.Delete(scenario1union)
        messages.addMessage("Deleted remaining interim files")

        return

##-------------------------------------------------------------------------------------

class Nsicreatorapi(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "NSI Creator (from network api)"
        self.description = ""
        self.canRunInBackground = True

    def getParameterInfo(self):
        """Define parameter definitions"""
        inputstudyarea = arcpy.Parameter(
            displayName="Study Area",
            name="study area",
            datatype="GPFeatureLayer",
            parameterType="Required",
            direction="Input")
        #inputstudyarea.value = fr"D:\~Kurt_2nd\NSI-2022\TESTING\StudyArea_TwoState_TEST.shp"
        workspace = arcpy.Parameter(
            displayName="Output Folder",
            name="output folder",
            datatype="DEFolder",
            parameterType="Required",
            direction="Input")
        inputuscounties = arcpy.Parameter(
            displayName="US Counties Shapefile",
            name="us counties shapefile",
            datatype="GPFeatureLayer",
            parameterType="Required",
            direction="Input")
        #inputuscounties.value = fr"D:\~Kurt_2nd\NSI-2022\TESTING\US_Counties_TwoState_TEST.shp"
        priceindex = arcpy.Parameter(
            displayName="Price Index from Base (default: 1)",
            name="price index",
            datatype="GPDouble",
            parameterType="Required",
            direction="Input")
        priceindex.value = 1.177
        outputname = arcpy.Parameter(
            displayName="Output NSI File Name (no extension, include price level)",
            name="Output NSI filename",
            datatype="GPString",
            parameterType="Required",
            direction="Output")
        outputname.value = "NSI2022_2025pricelevel"
        addfianame = arcpy.Parameter(
            displayName="Check to add name field for FIA",
            name="Check to add name field",
            datatype="GPBoolean",
            parameterType="Optional",
            direction="Input")
        internalnsi = arcpy.Parameter(
            displayName="Check to get USACE-only fields",
            name="Check to get USACE-only fields",
            datatype="GPBoolean",
            parameterType="Optional",
            direction="Input")
        splitbypolygons = arcpy.Parameter(
            displayName="Check to split by study area polygons",
            name="Check to split by study area polygons",
            datatype="GPBoolean",
            parameterType="Optional",
            direction="Input")
        clearbuffer = arcpy.Parameter(
            displayName="Check to remove points within 20 feet of the boundary",
            name="Check to remove points within 20 feet of the boundary",
            datatype="GPBoolean",
            parameterType="Optional",
            direction="Input")
        clearbuffer.value = True
        reprojectoutput = arcpy.Parameter(
            displayName="Check to reproject output gpkg to standard",
            name="Check to reproject output gpkg to standard",
            datatype="GPBoolean",
            parameterType="Optional",
            direction="Input")

        parameters = [inputstudyarea, workspace, inputuscounties, priceindex, outputname, addfianame, internalnsi, splitbypolygons, clearbuffer, reprojectoutput]
        return parameters

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return
       
    def execute(self, parameters, messages):
        """The source code of the tool."""
        #call parameters that were entered and create new variables as needed
        inputstudyarea = parameters[0].valueAsText
        inputuscounties = parameters[2].valueAsText
        workspace = parameters[1].valueAsText
        outputname = parameters[4].valueAsText
        priceindex = parameters[3].valueAsText
        #boolean parameters should just be the value, not text
        addfianame = parameters[5].value
        internalnsi = parameters[6].value
        splitbypolygons = parameters[7].value
        clearbuffer = parameters[8].value
        reproject = parameters[9].value
        #clearbuffer = True

        arcpy.management.CreateFolder(workspace, "interimfiles")
        interimfolder = fr"{workspace}\interimfiles"

        #if statement that checks to make sure that both the clearbuffer and splitbypolygons are not checked at the same time and raises an error if they are
        #if both are selected, points within 20 feet of each individual polygon boundary will be removed, which is bad if you have reaches split out in the study area
        if clearbuffer and splitbypolygons:
            arcpy.AddError("You cannot check both the clear buffer and split by polygons options at the same time. Please uncheck one of them.")
            return
      

        #if then statement below adds the word internal to output filename if the internal api is accessed
        if internalnsi:
            outputnsifile = fr"{workspace}\{outputname}_internal.shp"
            outputnsifileprj = fr"{workspace}\{outputname}_internal_prj.shp"
            outputnamelayer = fr"{outputname}_internal"
            messages.addMessage("USACE-only fields was selected, internal api will be used and filename will have the word internal added at the end.")
        else:
            outputnsifile = fr"{workspace}\{outputname}.shp"
            outputnsifileprj = fr"{workspace}\{outputname}_prj.shp"
            outputnamelayer = fr"{outputname}"
            messages.addMessage("USACE-only fields was not selected, public/external api will be used.")

        #Clip US County shapefile to study area
        messages.addMessage("Clipping county shapefile to input study area...")
        arcpy.analysis.PairwiseClip(
            in_features=inputuscounties, 
            clip_features=inputstudyarea, 
            out_feature_class=fr"{workspace}\ClippedUSCounties.shp", 
            cluster_tolerance="")
        
        #set clipped counties as a parameter
        ClippedUSCounties=fr"{workspace}\ClippedUSCounties.shp"
        
        #create list of all the FIPS county numbers in the clipped shapefile, need error message for when there is no FIPS field
        FipsList = [row[0] for row in arcpy.da.SearchCursor(ClippedUSCounties, ["FIPS"])]
        
        #cont number of entries and set initial loop number
        cntTotal = len(FipsList)
        messages.addMessage("There are {0} counties in the study area".format(cntTotal))
        loopNumber = 1
        
        #create output lists to append each county to
        jsonfilelist=[]
        countygpkgfilelist=[]
        clippedgpkgfilelist=[]
        clippedgpkgtablelist=[]
        clippedshapefilelist=[]
        fullshapefilelist=[]
        timestartloop = datetime.now()
        #Iterate over the list of FIPS codes, clipping each NSI county file to the study area
        messages.addMessage("Downloading JSON, converting to gpkg, and clipping the NSI file of each FIPS number found in the clipped county shapefile to the Study Area...")
        
        #ctx method for ignoring ssl certificate error when hitting the api url
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE

        for FIPS_Value in FipsList:
            messages.addMessage("County {0} of {1}".format(loopNumber, cntTotal))
            loopNumber += 1
            #download the JSON string for the FIPS code from api and write it to a JSON file, based on article at https://support.esri.com/en/technical-article/000019645
            messages.addMessage(str(FIPS_Value)+" county JSON file is being downloaded from api...")
            #if elese statement sets url to either the internal or extrenal api address
            if internalnsi:
                url = fr"https://nsi.sec.usace.army.mil/internal/nsiapi/structures?fips={FIPS_Value}&fmt=fc"
            else:
                url = fr"https://nsi.sec.usace.army.mil/nsiapi/structures?fips={FIPS_Value}&fmt=fc"
            response = urllib.request.urlopen(url, context=ctx)
            json = response.read()
            with open(fr"{interimfolder}\{FIPS_Value}.json", "wb") as ms_json:
                ms_json.write(json)
            
            #create geopackage for full county and convert json to geopackage
            countygpkg = fr"{interimfolder}\{FIPS_Value}.gpkg"
            countygpkgtable = os.path.join(countygpkg, "nsi")
            arcpy.management.CreateSQLiteDatabase(countygpkg, "GEOPACKAGE")
            arcpy.conversion.JSONToFeatures(fr"{interimfolder}\{FIPS_Value}.json", countygpkgtable)
            messages.addMessage(str(FIPS_Value)+" JSON file has been downloaded and converted to gpkg")

            #create geopackage for clipped county
            clippedcountygpkg = fr"{interimfolder}\{FIPS_Value}_clipped.gpkg"
            clippedcountygpkgtable = os.path.join(clippedcountygpkg, "nsi")
            arcpy.management.CreateSQLiteDatabase(clippedcountygpkg, "GEOPACKAGE")

            #append county json, gkpg, clipped gpkg and clipped tables to lists, then clip to study area
            jsonfilelist.append(fr"{interimfolder}\{FIPS_Value}.json")
            countygpkgfilelist.append(countygpkg)
            clippedgpkgfilelist.append(clippedcountygpkg)
            clippedgpkgtablelist.append(clippedcountygpkgtable)
            arcpy.analysis.PairwiseClip(
                in_features=countygpkgtable, 
                clip_features=inputstudyarea, 
                out_feature_class=clippedcountygpkgtable, 
                cluster_tolerance="")
            messages.addMessage(str(FIPS_Value)+" gpkg has been clipped to the study area")

            #end of county loop

        timenow0 = datetime.now()
        looptime = timenow0 - timestartloop
        #timenow0format = timenow0.time().strftime('%H:%M:%S')
        messages.addMessage("Downloads took {0}...".format(looptime))

        #create interim geopackage to merge clipped county geopackages table list into one interim geopackage
        messages.addMessage("Merging county gpkg tables into single output shapefile...")
        mergedgpkg = fr"{interimfolder}\InterimMergedGP.gpkg"
        arcpy.management.CreateSQLiteDatabase(mergedgpkg, "GEOPACKAGE")
        #arcpy.management.CreateTable(mergedgpkg, "nsimerge", oid_type="32_BIT")
        mergedgpkgtable = os.path.join(mergedgpkg, outputnamelayer)
        arcpy.management.Merge(clippedgpkgtablelist, mergedgpkgtable)
        ###export interim geopackage to shapefile, Pro 3.2 has issues here because it create 64-bit ObjectID fields, feature class to shapefile seems to work
        #arcpy.conversion.ExportFeatures(mergedgpkgtable, outputnsifile)
        #arcpy.conversion.FeatureClassToFeatureClass(mergedgpkgtable, workspace, fr"{outputname}.shp")
        arcpy.conversion.FeatureClassToShapefile(mergedgpkgtable, workspace)
     
        # --- Check shapefile size after conversion ---
        import glob

        shapefile_base = os.path.splitext(os.path.basename(outputnsifile))[0]
        shapefile_pattern = os.path.join(workspace, shapefile_base + ".*")
        shapefile_parts = glob.glob(shapefile_pattern)
        total_size = sum(os.path.getsize(f) for f in shapefile_parts)

        # 2 GB = 2 * 1024 * 1024 * 1024 bytes
        if total_size > 2 * 1024 * 1024 * 1024:
            messages.addWarningMessage(
                f"WARNING: The output shapefile '{shapefile_base}' is larger than 2 GB ({total_size/1024/1024/1024:.2f} GB). "
                "Shapefiles have a 2 GB size limit and may be missing features or attributes. "
                "See output folder /interimfiles/InterimMergedGP.gpkg for the full output with no indexing."
            )
        # --- End of shapefile size check ---

        timenow2 = datetime.now()
        mergetime = timenow2 - timenow0
        messages.addMessage("Merge time was {0}".format(mergetime))

        #index value fields using calculate field
        messages.addMessage("Indexing structure values, content values, and vehicle values...")
        arcpy.management.CalculateField(in_table=outputnsifile, 
            field="Val_Struct", 
            expression=f"!Val_Struct! * {priceindex}", 
            expression_type="PYTHON_9.3", 
            code_block="", 
            field_type="TEXT", 
            enforce_domains="NO_ENFORCE_DOMAINS")

        arcpy.management.CalculateField(in_table=outputnsifile, 
            field="Val_Cont", 
            expression=f"!Val_Cont! * {priceindex}", 
            expression_type="PYTHON_9.3", 
            code_block="", 
            field_type="TEXT", 
            enforce_domains="NO_ENFORCE_DOMAINS")

        arcpy.management.CalculateField(in_table=outputnsifile, 
            field="Val_Vehic", 
            expression=f"!Val_Vehic! * {priceindex}", 
            expression_type="PYTHON_9.3", 
            code_block="", 
            field_type="TEXT", 
            enforce_domains="NO_ENFORCE_DOMAINS")

        timenow3 = datetime.now()
        indextime = timenow3 - timenow2
        messages.addMessage("Index time was {0}".format(indextime))

        #If boolean is checked ST_Name field will be added and calculated, FIA requires a unique name field
        #Shapefile needs FID field, geopackage needs OBJECTID field
        if addfianame:
            messages.addMessage("Adding St_Name field and setting it equal to OccTyp + FID...")
            arcpy.management.AddField(in_table=outputnsifile, field_name="ST_Name", field_type="TEXT", 
                field_precision=None, field_scale=None, field_length=50, field_alias="", field_is_nullable="NULLABLE", 
                field_is_required="NON_REQUIRED", field_domain="")

            arcpy.management.CalculateField(in_table=outputnsifile, field="ST_Name", expression="!OccType! + ' ' + str(!FID!)", 
                expression_type="PYTHON_9.3", code_block="", field_type="TEXT", enforce_domains="NO_ENFORCE_DOMAINS")

        if clearbuffer:
            messages.addMessage("Clearing points within 20 feet of boundary")
            clearbufferpoly = fr"{workspace}\studybuffer.shp"
            arcpy.analysis.PairwiseDissolve(inputstudyarea, clearbufferpoly, "", "", "MULTI_PART")
            deleteselect = arcpy.management.SelectLayerByLocation(outputnsifile, "INTERSECT", clearbufferpoly, "-50 Feet", "NEW_SELECTION", "INVERT")
            arcpy.management.DeleteFeatures(deleteselect)

        if reproject:
            # Define the target spatial reference using the provided projection
            usa_albers_spatial_ref = arcpy.SpatialReference()
            usa_albers_spatial_ref.loadFromString("""PROJCS["USA_Contiguous_Albers_Equal_Area_Conic_USGS_version",
                GEOGCS["GCS_North_American_1983",
                DATUM["D_North_American_1983",
                SPHEROID["GRS_1980",6378137.0,298.257222101]],
                PRIMEM["Greenwich",0.0],
                UNIT["Degree",0.0174532925199433]],
                PROJECTION["Albers"],
                PARAMETER["False_Easting",0.0],
                PARAMETER["False_Northing",0.0],
                PARAMETER["Central_Meridian",-96.0],
                PARAMETER["Standard_Parallel_1",29.5],
                PARAMETER["Standard_Parallel_2",45.5],
                PARAMETER["Latitude_Of_Origin",23.0],
                UNIT["Foot_US",0.3048006096012192]]""")

            # Project the shapefile to the USA Contiguous Albers projection
            messages.addMessage("Reprojecting the output file...")
            arcpy.management.Project(outputnsifile, outputnsifileprj, usa_albers_spatial_ref)
            #delete original outputNSIfile
            arcpy.management.Delete(outputnsifile)
            outputnsifile = outputnsifileprj

        #create output gpkg file, merge the clipped county table list
        messages.addMessage("Creating geopackage (.gpkg) file output...")
        outputgpkg = fr"{workspace}\{outputnamelayer}.gpkg"
        arcpy.management.CreateSQLiteDatabase(outputgpkg, "GEOPACKAGE")
        arcpy.conversion.FeatureClassToFeatureClass(outputnsifile, outputgpkg, "nsi")

        if splitbypolygons:
            # make sure fieldname is correct
            fieldname = "Name"
            messages.addMessage("Splitting by input study area field named: {0}".format(f'{fieldname}'))
            arcpy.management.CreateFolder(workspace, "SplitOuts")
            splitFeatures = inputstudyarea
            splitfolder = fr"{workspace}\SplitOuts"
            try:
                arcpy.analysis.Split(outputnsifile, splitFeatures, fieldname, splitfolder)
            except Exception:
                messages.addMessage("Error, likely did not find study area field named: {0}, so skipping the split".format(f'{fieldname}'))
            arcpy.env.workspace = splitfolder
            shps = arcpy.ListFeatureClasses()
            for shp in shps:
                filename = os.path.basename(shp)
                filenamenoext = "{}".format((os.path.splitext(filename)[0]))
                messages.addMessage("_Splitting Filename: {0}".format(f'{filename}'))
                
                ffrd=True
                if ffrd:
                    outputsplitgpkg = fr"{splitfolder}\{filenamenoext}_unadjusted.gpkg"
                else:
                    outputsplitgpkg = fr"{splitfolder}\{filenamenoext}.gpkg"
                
                
                arcpy.management.CreateSQLiteDatabase(outputsplitgpkg, "GEOPACKAGE")
                arcpy.conversion.FeatureClassToFeatureClass(shp, outputsplitgpkg, "nsi")

        #messages.addMessage("Successful completion, deleting individual json files...")
        #timenow4 = datetime.now()
        #deleting the interim folder only gets rid of JSONs in the folder, deleting all of the other files adds lots of time
        #arcpy.management.Delete(interimfolder)
        #arcpy.management.Delete(jsonfilelist)
        #arcpy.management.Delete(countygpkgfilelist)
        #arcpy.management.Delete(clippedgpkgfilelist)
        #timenow5 = datetime.now()
        #deletetime = timenow5 - timenow4
        #messages.addMessage("Delete time was {0}".format(deletetime))
        messages.addMessage("All done!! Yay!!")

        return

##-------------------------------------------------------------------------------------

class Reclassifytozone(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Reclassify Grid into a Zone Polygon"
        self.description = ""
        self.canRunInBackground = True

    def getParameterInfo(self):
        """Define parameter definitions"""
        inputgrid = arcpy.Parameter(
            displayName="Input Grid",
            name="Input Grid",
            datatype="DERasterDataset",
            parameterType="Required",
            direction="Input")
        Base_zero_value = arcpy.Parameter(
            displayName="Base or Zero Value for Zones (ex. breach time)",
            name="Base or Zero value",
            datatype="GPDouble",
            parameterType="Required",
            direction="Input")
        Max_for_Range_1 = arcpy.Parameter(
            displayName="Max Value for First Range (1)",
            name="range1",
            datatype="GPDouble",
            parameterType="Required",
            direction="Input")
        Max_for_Range_1.value = 1
        Max_for_Range_2 = arcpy.Parameter(
            displayName="Max Value for Second Range (2)",
            name="range2",
            datatype="GPDouble",
            parameterType="Required",
            direction="Input")
        Max_for_Range_2.value = 2
        Max_for_Range_3 = arcpy.Parameter(
            displayName="Max Value for Third Range (4)",
            name="range3",
            datatype="GPDouble",
            parameterType="Required",
            direction="Input")
        Max_for_Range_3.value = 4
        Max_for_Range_4 = arcpy.Parameter(
            displayName="Max Value for Fourth Range (8)",
            name="range4",
            datatype="GPDouble",
            parameterType="Required",
            direction="Input")
        Max_for_Range_4.value = 8
        Max_for_Range_5 = arcpy.Parameter(
            displayName="Max Value for Fifth Range (24)",
            name="range5",
            datatype="GPDouble",
            parameterType="Required",
            direction="Input")
        Max_for_Range_5.value = 24
        units = arcpy.Parameter(
            displayName="Unit Name (default: hrs)",
            name="units",
            datatype="GPString",
            parameterType="Required",
            direction="Input")
        outputfolder2 = arcpy.Parameter(
            displayName="Output Folder",
            name="Output Folder",
            datatype="DEFolder",
            parameterType="Required",
            direction="Input")
        outputfilename = arcpy.Parameter(
            displayName="Output Filename with no extension",
            name="Output Filename",
            datatype="GPString",
            parameterType="Required",
            direction="Input")
        outputfilename.value = "NAMEHERE_ArrivalRanges"
        simplifypolygons = arcpy.Parameter(
            displayName="Check to Simplify Polygons",
            name="Check to simplify polygons",
            datatype="GPBoolean",
            parameterType="Optional",
            direction="Input")
        
        # Add a read-only textbox parameter
        info_text = arcpy.Parameter(
            displayName="Information / Help",
            name="info_text",
            datatype="GPString",
            parameterType="Optional",  # Derived makes it read-only
            direction="Output",)  # Output ensures it is displayed as a textbox
        #info_text.multiValue = True  # \n adds a return, but with multivalue you can use a semicolon to make a whole separate text box
        info_text.value = "Typically used to create a polygon with arrival time zones. " \
        "\nNeeds an arrival time grid from RAS Mapper. " \
        "\nCould be used to create depth zones also (use Max Depth grid, ft instead of hrs). " 

        parameters = [inputgrid, Base_zero_value, Max_for_Range_1, Max_for_Range_2, Max_for_Range_3, Max_for_Range_4, Max_for_Range_5, units, outputfolder2, outputfilename, simplifypolygons, info_text]
        return parameters

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return
       
    def execute(self, parameters, messages):
        """The source code of the tool."""
        #call parameters that were entered and create new variables as needed
        Input_Grid = parameters[0].valueAsText
        Base_or_Zero_Value = parameters[1].valueAsText
        Max_for_range_1 = parameters[2].valueAsText
        Max_for_range_2 = parameters[3].valueAsText
        Max_for_range_3 = parameters[4].valueAsText
        Max_for_range_4 = parameters[5].valueAsText
        Max_for_range_5 = parameters[6].valueAsText
        Units = parameters[7].valueAsText
        Output_Folder = parameters[8].valueAsText
        Output_Polygon_Name = parameters[9].valueAsText
        #boolean parameters should just be the value, not text
        Simplify_polygons_ = parameters[10].value
    
        Output_Polygon_Shapefile = fr"{Output_Folder}\{Output_Polygon_Name}.shp"

        arcpy.env.overwriteOutput = True

        Range1 = float(Base_or_Zero_Value) + float(Max_for_range_1)
        Range2 = float(Base_or_Zero_Value) + float(Max_for_range_2)
        Range3 = float(Base_or_Zero_Value) + float(Max_for_range_3)
        Range4 = float(Base_or_Zero_Value) + float(Max_for_range_4)
        Range5 = float(Base_or_Zero_Value) + float(Max_for_range_5)
        zero = float(Base_or_Zero_Value)
        # Check out any necessary licenses.
        arcpy.CheckOutExtension("Spatial")

        # Process: Calculate ranges (Raster Calculator) (sa)
        messages.addMessage("Reclassifying grid into zones...")
        Recalculated_grid = fr"{Output_Folder}\recalcgrid.tif"
        Calculate_ranges = Recalculated_grid
        Recalculated_grid = Con(Float(Input_Grid) <= float(zero),0,Con(Float(Input_Grid) <= float(Range1),1,Con(Float(Input_Grid) <= float(Range2),2,Con(Float(Input_Grid) <= float(Range3),3,Con(Float(Input_Grid) <= float(Range4),4,Con(Float(Input_Grid) <= float(Range5),5,Con(Float(Input_Grid) > float(Range5),6)))))))
        Recalculated_grid.save(Calculate_ranges)
        messages.addMessage("Finished reclassifying grid into zones")

        # Process: recalculated grid to polygon (Raster to Polygon) (conversion)
        messages.addMessage("Converting reclassified raster to polygon...")
        output_polygon_1 = fr"{Output_Folder}\RecalcPolygon1.shp"
        arcpy.conversion.RasterToPolygon(in_raster=Recalculated_grid, out_polygon_features=output_polygon_1, simplify=Simplify_polygons_, raster_field="", create_multipart_features="SINGLE_OUTER_PART", max_vertices_per_feature=None)
        messages.addMessage("Finished converting reclassified raster to polygon")

        # Process: Pairwise Dissolve Ranges (Pairwise Dissolve) (analysis)
        messages.addMessage("Dissolving polygon into aggregated zones...")
        arcpy.analysis.PairwiseDissolve(in_features=output_polygon_1, out_feature_class=Output_Polygon_Shapefile, dissolve_field="GRIDCODE", statistics_fields=[], multi_part="MULTI_PART")
        messages.addMessage("Finished dissolving polygon into aggregated zones")

        # Process: Add range field (Add Field) (management)
        messages.addMessage("Adding a range text field...")
        Range_field_added = arcpy.management.AddField(in_table=Output_Polygon_Shapefile, field_name="Range", field_type="TEXT", field_precision=None, field_scale=None, field_length=None, field_alias="", field_is_nullable="NULLABLE", field_is_required="NON_REQUIRED", field_domain="")[0]

        messages.addMessage("Calculating names for the range text field...")
        fieldName = "Range"
        renameExpression = fr"Rangerename(str(!GRIDCODE!))"
        codeblock = """def Rangerename(value):
            if value == "0":
                return "Negative Values"

            elif value == "1":
                return "0 - " + fr"{0} {5}"

            elif value == "2":
                return fr"{0} - " + fr"{1} {5}"

            elif value == "3":
                return fr"{1} - " + fr"{2} {5}"

            elif value == "4":
                return fr"{2} - " + fr"{3} {5}"

            elif value == "5":
                return fr"{3} - " + fr"{4} {5}"
            
            elif value == "6":
                return "Greater than " + fr"{4} {5}"

            else:
                return value""".format(Max_for_range_1, Max_for_range_2, Max_for_range_3, Max_for_range_4, Max_for_range_5, Units)
        
        # Process: Calculate range field (Calculate Field) (management)
        arcpy.management.CalculateField(Range_field_added, fieldName, renameExpression, "PYTHON3", codeblock)

        #Delete interim grid and polygon
        arcpy.management.Delete(fr"{Output_Folder}\recalcgrid.tif")
        arcpy.management.Delete(fr"{Output_Folder}\RecalcPolygon1.shp")
        return

        ##-------------------------------------------------------------------------------------

class Rastertopolygon(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Convert Raster to Polygon"
        self.description = ""
        self.canRunInBackground = True

    def getParameterInfo(self):
        """Define parameter definitions"""
        inputraster = arcpy.Parameter(
            displayName="Input Raster Dataset",
            name="Input Raster Dataset",
            datatype="DERasterDataset",
            parameterType="Required",
            direction="Input")
        outputpolygon = arcpy.Parameter(
            displayName="Output Polygon",
            name="Output Polygon",
            datatype="GPFeatureLayer",
            parameterType="Required",
            direction="Output")
        simplifypolygon = arcpy.Parameter(
            displayName="Check to simplify output polygon",
            name="Check to simplify output polygon",
            datatype="GPBoolean",
            parameterType="Optional",
            direction="Input")

        parameters = [inputraster, outputpolygon, simplifypolygon]
        return parameters

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return
       
    def execute(self, parameters, messages):
        """The source code of the tool."""
        #call parameters that were entered and create new variables as needed
        inputraster = parameters[0].valueAsText
        outputpolygon = parameters[1].valueAsText
        simplifypolygon = parameters[2].value

        interim_raster = fr"C:\Temp\reclassgrid.tif"
        interim_polygon = fr"C:\Temp\temppolygon.shp"

        arcpy.CheckOutExtension("Spatial")
        
        if simplifypolygon:
            simp = "SIMPLIFY"
            messages.addMessage("Simplify option selected.")
        else: 
            simp = "NO_SIMPLIFY"
            messages.addMessage("Simplify option was not selected.")

        messages.addMessage("Reclassifying raster to integer...")
        # Process: Reclassify (Reclassify) (sa)
        Reclass_tif1 = interim_raster
        myRemapRange = RemapRange([[-100000, 0, 1],[0, 100000,2]])
        Reclass_tif1 = Reclassify(inputraster, "VALUE", myRemapRange)
        Reclass_tif1.save(interim_raster)

        # Process: Raster to Polygon (Raster to Polygon) (conversion)
        messages.addMessage("Converting interim reclassified raster to interim polygon...")
        with arcpy.EnvManager(outputMFlag="Disabled", outputZFlag="Disabled"):
            arcpy.conversion.RasterToPolygon(in_raster=Reclass_tif1, out_polygon_features=interim_polygon, simplify=simp, raster_field="", create_multipart_features="SINGLE_OUTER_PART", max_vertices_per_feature=None)

        # Process: Dissolve (Dissolve) (management)
        messages.addMessage("Dissolving interim polygon...")
        #arcpy.management.Dissolve(in_features=interim_polygon, out_feature_class=outputpolygon, dissolve_field=[], statistics_fields=[], multi_part="MULTI_PART", unsplit_lines="DISSOLVE_LINES")
        arcpy.analysis.PairwiseDissolve(in_features=interim_polygon, out_feature_class=outputpolygon, dissolve_field=[], statistics_fields=[], multi_part="MULTI_PART")

        #Delete interim grid and polygon
        messages.addMessage("Deleting interim grid and interim polygon...")
        arcpy.management.Delete(fr"C:\Temp\reclassgrid.tif")
        arcpy.management.Delete(fr"C:\Temp\temppolygon.shp")
        
        return
    
class Snaptoexclusionpoly(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Snap Structures to Exclusion Polygon"
        self.description = ""
        self.canRunInBackground = False

    def getParameterInfo(self):
        """Define parameter definitions"""
        sourcensifile = arcpy.Parameter(
            displayName="NSI Inventory Shapefile",
            name="NSI inventory shapefile",
            datatype="DEShapefile",
            parameterType="Required",
            direction="Input")
        #sourcensifile.value = fr"D:\~Kurt_2nd\~2nd_Projects\Little_Goose_MMC\Shapefiles\SnapTesting1\DataPackage2\Structures1.shp"

        inputpolygon = arcpy.Parameter(
            displayName="Exclusion Polygon",
            name="Input Polygon",
            datatype="DEFeatureClass",
            parameterType="Required",
            direction="Input")
        #inputpolygon.value = fr"D:\~Kurt_2nd\~2nd_Projects\Little_Goose_MMC\Shapefiles\SnapTesting1\DataPackage2\In_Pool.shp"

        bufferdistance = arcpy.Parameter(
            displayName="Distance to move structures out of polygon (ft)",
            name="buffer_distance",
            datatype="GPLong",
            parameterType="Required",
            direction="Input")
        bufferdistance.value = fr"25"

        outputname = arcpy.Parameter(
            displayName="Output NSI File Name (no extension)",
            name="output filename",
            datatype="GPString",
            parameterType="Required",
            direction="Output")
        outputname.value = "NSI2022_2023pricelevel_snapped"

        parameters = [sourcensifile, inputpolygon, bufferdistance, outputname]
        return parameters

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return

    def execute(self, parameters, messages):
        """The source code of the tool."""
      
        sourcensifile = parameters[0].valueAsText
        inputpolygon = parameters[1].valueAsText
        bufferdistance = parameters[2].valueAsText
        outputname = parameters[3].valueAsText
         
        arcpy.env.overwriteOutput = True

        outputfolder1 = os.path.dirname(os.path.abspath(sourcensifile))
        outputnsifile = fr"{outputfolder1}\{outputname}.shp"
        
        interimpolybuf = "interimpoly"

        messages.addMessage("Copying inventory...")
        arcpy.management.CopyFeatures(sourcensifile, outputnsifile)
        structurecount = arcpy.management.GetCount(outputnsifile)

        messages.addMessage("Adding SnapMod field...")
        arcpy.management.AddField(in_table=outputnsifile, field_name="SnapMod", field_type="TEXT", 
            field_precision=None, field_scale=None, field_length=10, field_alias="", field_is_nullable="NULLABLE", 
            field_is_required="NON_REQUIRED", field_domain="")

        messages.addMessage("Calculating SnapMod field as no for all structures...")
        arcpy.management.CalculateField(in_table=outputnsifile, field="SnapMod", expression="'no'", 
            expression_type="PYTHON3", code_block="", field_type="TEXT", enforce_domains="NO_ENFORCE_DOMAINS")

        messages.addMessage("Buffering polygon by {} feet...".format(bufferdistance))
        #run pairwise buffer that also simplifies a little by allowing for a deviation of half the buffer value
        with arcpy.EnvManager(parallelProcessingFactor="100"):
            arcpy.analysis.PairwiseBuffer(in_features=inputpolygon, 
                out_feature_class=interimpolybuf, 
                buffer_distance_or_field=bufferdistance +" Feet", 
                dissolve_option="NONE", 
                dissolve_field=[], 
                method="PLANAR", 
                max_deviation=str(int(bufferdistance) // 2) +" Feet")
        
        messages.addMessage("Selecting structures in original polygon (longest part of the process)...")
        floodedstrucs = arcpy.management.SelectLayerByLocation(outputnsifile, "INTERSECT", inputpolygon, None, "NEW_SELECTION", "NOT_INVERT")
        snapcount = arcpy.management.GetCount(floodedstrucs)

        messages.addMessage("Calculating SnapMod field as yes for selected structures...")
        arcpy.management.CalculateField(in_table=floodedstrucs, field="SnapMod", expression="'yes'", 
            expression_type="PYTHON3", code_block="", field_type="TEXT", enforce_domains="NO_ENFORCE_DOMAINS")

        #the snap tool will only snap things within a specified distance of the edge boundary, set at 100k ft to snap everything selected
        #this could be reduced if you wanted to keep things that are in the center of the polygon for manual deletion
        snapdistance = 100000

        messages.addMessage("Snapping selected structures ({0} of {1} structures) to edge of buffered polygon...".format(snapcount, structurecount))
        snapEnv = "{0} EDGE '{1} Feet'".format(interimpolybuf, snapdistance)
        arcpy.edit.Snap(floodedstrucs, snapEnv)

        messages.addMessage("Analysis Complete. Output saved in input file folder as {}".format(outputnsifile))
        messages.addMessage("Thanos won. You should have went for the head.")
        return
    
class Batchdefineprojection(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Define Unknown .tif Projections in folder"
        self.description = ""
        self.canRunInBackground = True

    def getParameterInfo(self):
        """Define parameter definitions"""
        workspace = arcpy.Parameter(
            displayName="Workspace folder",
            name="workspace",
            datatype="DEFolder",
            parameterType="Required",
            direction="Input")
        projectedfile = arcpy.Parameter(
            displayName="Raster or shp file with correct projection",
            name="projectedfile",
            datatype=["DERasterDataset","DEShapeFile"],
            parameterType="Required",
            direction="Input")
        filetype = arcpy.Parameter(
            displayName="Raster File Types",
            name="filetype",
            datatype="GPString",
            parameterType="Required",
            direction="Input")
        filetype.filter.type = "ValueList"
        filetype.filter.list = [".tif", ".vrt"]
        filetype.value = ".tif"

        overwriteprojection = arcpy.Parameter(
            displayName="Check to overwrite any existing projection",
            name="overwrite projection",
            datatype="GPBoolean",
            parameterType="Optional",
            direction="Input")

        parameters = [workspace, projectedfile, filetype, overwriteprojection]
        return parameters

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return
       
    def execute(self, parameters, messages):
        """The source code of the tool."""
        #call parameters that were entered and create new variables as needed
        workspace = parameters[0].valueAsText
        projection_file = parameters[1].valueAsText
        filetype = parameters[2].valueAsText
        overwrite = parameters[3].value

        # Get the spatial reference object from the projection file
        prj1 = arcpy.Describe(projection_file)
        prj2 = prj1.spatialReference.Name
        messages.addMessage(f"Projection to define: {prj2}")
        #projection_sr = arcpy.SpatialReference(projection_file)
        filenumber = 1
        # Recursively search the workspace for files without a projection
        for root, dirs, files in os.walk(workspace):
            for file in files:
                if file.endswith(filetype):
                    # Check if the file has a projection already defined
                    desc = arcpy.Describe(os.path.join(root, file))
                    messages.addMessage(f"File {filenumber}- FileName: {file}")
                    desc2 = desc.spatialReference.Name
                    messages.addMessage(f"File {filenumber}- Original projection: {desc2}")
                    if desc.spatialReference.Name == "Unknown":
                        # Define the projection of the file
                        #arcpy.DefineProjection_management(os.path.join(root, file), prj1.spatialReference)
                        arcpy.management.DefineProjection(os.path.join(root, file), prj1.spatialReference)
                        messages.addMessage(f"File {filenumber}- Defined new projection.")
                    else:
                        messages.addMessage(f"File {filenumber}- Already has a projection defined.")
                        if overwrite:
                            messages.addMessage(f"File {filenumber}- Overwriting existing projection with new projection.")
                            arcpy.management.DefineProjection(os.path.join(root, file), prj1.spatialReference)
                            
                    filenumber += 1

        return

##-----------------------------------------------

class Incrementallifesimresults(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Calculate Incremental LifeSim Results"
        self.description = ""
        self.canRunInBackground = False

    def getParameterInfo(self):
        """Define parameter definitions"""

        higherscenario = arcpy.Parameter(
            displayName = "Higher Scenario Shapefile (exported from LS)",
            name = "Higher Scenario",
            datatype = "DEFeatureClass",
            parameterType = "Required",
            direction = "Input",)

        lowerscenario = arcpy.Parameter(
            displayName = "Lower Scenario Shapefile (exported from LS)",
            name = "Lower Scenario",
            datatype = "DEFeatureClass",
            parameterType = "Required",
            direction = "Input",)

        outputfile = arcpy.Parameter(
            displayName = "Output GeoPackage Name, no extension",
            name = "Output GeoPackage",
            datatype = "DEFile",
            parameterType = "Required",
            direction = "Output",)
        
        deletenonflooded = arcpy.Parameter(
            displayName="Check to delete structures not flooded in either scenario from incremental results",
            name="Check to delete nonflooded",
            datatype="GPBoolean",
            parameterType="Optional",
            direction="Input")
        deletenonflooded.value = True

        # Add a read-only textbox parameter
        info_text = arcpy.Parameter(
            displayName="Information / Help",
            name="info_text",
            datatype="GPString",
            parameterType="Optional",  # Derived makes it read-only
            direction="Output",)  # Output ensures it is displayed as a textbox
        #info_text.multiValue = True  # \n adds a return, but with multivalue you can use a semicolon to make a whole separate text box
        info_text.value = "In LifeSim, right click on a Simulation and hit View Results Maps. " \
        "\nCheck the two Structure Summary results you want. " \
        "\nIn map window, right click, go to Tools, and Export to New Shape."

        parameters = [higherscenario, lowerscenario, outputfile, deletenonflooded, info_text]
        return parameters

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return

    def execute(self, parameters, messages):
        """The source code of the tool."""
        higher_scenario_input = parameters[0].valueAsText
        lower_scenario_input = parameters[1].valueAsText
        output_gpkg = parameters[2].valueAsText
        delete_zero_depths = parameters[3].value

        arcpy.env.overwriteOutput = True

        # Check if both input shapefiles have the same number of rows
        higher_scenario_count = int(arcpy.management.GetCount(higher_scenario_input)[0])
        lower_scenario_count = int(arcpy.management.GetCount(lower_scenario_input)[0])

        if higher_scenario_count != lower_scenario_count:
            messages.addErrorMessage(
                f"Row count mismatch: Higher Scenario has {higher_scenario_count} rows, "
                f"while Lower Scenario has {lower_scenario_count} rows."
            )
            raise arcpy.ExecuteError("Input shapefiles must have the same number of rows.")
        else:
            messages.addMessage("Row counts are equal in both files, check passed.")

        # Ensure the output file has the .gpkg extension
        if not output_gpkg.lower().endswith(".gpkg"):
            output_gpkg += ".gpkg"
            messages.addMessage(f"Output file extension corrected to: {output_gpkg}")

        # Get geopackage directory and name, not actually used anymore though
        output_directory = os.path.dirname(output_gpkg)  # Get the directory of the output GeoPackage
        output_gpkg_name = os.path.splitext(os.path.basename(output_gpkg))[0]
        
        # Check if the GeoPackage already exists
        if arcpy.Exists(output_gpkg):
            messages.addMessage(f"GeoPackage already exists at: {output_gpkg}. Skipping creation.")
        else:
            # Create the GeoPackage if it doesn't exist
            arcpy.management.CreateSQLiteDatabase(out_database_name=output_gpkg, spatial_type="GEOPACKAGE")
            messages.addMessage(f"GeoPackage created at: {output_gpkg}")

        # Set the workspace to the GeoPackage
        arcpy.env.workspace = output_gpkg        

        # Define the layer names for the export
        higher_scenario_namenoext = os.path.splitext(os.path.basename(higher_scenario_input))[0]
        lower_scenario_namenoext = os.path.splitext(os.path.basename(lower_scenario_input))[0]

        # Try to copy features over using the shapefile names, use generic high-low-incremental if that causes an error
        try:
            higher_scenario = f"{output_gpkg}\\{higher_scenario_namenoext}"
            lower_scenario = f"{output_gpkg}\\{lower_scenario_namenoext}"
            incremental_scenario = f"{output_gpkg}\\Incremental_1_{higher_scenario_namenoext}_2_{lower_scenario_namenoext}"
            arcpy.management.CopyFeatures(higher_scenario_input, higher_scenario)
            arcpy.management.CopyFeatures(lower_scenario_input, lower_scenario)
            arcpy.management.CopyFeatures(higher_scenario_input, incremental_scenario)
            messages.addMessage("Exported data to GeoPackage with scenario names.")
        except:
            higher_scenario = f"{output_gpkg}\\Higher_Scenario"
            lower_scenario = f"{output_gpkg}\\Lower_Scenario"
            incremental_scenario = f"{output_gpkg}\\Incremental"
            arcpy.management.CopyFeatures(higher_scenario_input, higher_scenario)
            arcpy.management.CopyFeatures(lower_scenario_input, lower_scenario)
            arcpy.management.CopyFeatures(higher_scenario_input, incremental_scenario)
            messages.addMessage("Exported data to GeoPackage with generic names due to error.")

        # Combine all fields into a single AddFields call
        messages.addMessage("Adding all necessary fields to the output feature classes...")

        all_fields = [
            # Fields from the higher scenario
            ["EPZ1", "TEXT", "", "100", "", ""],
            ["EPZ2", "TEXT", "", "100", "", ""],
            ["OccTyp1", "TEXT", "", "15", "", ""],
            ["OccTyp2", "TEXT", "", "15", "", ""],  
            ["NumStry1", "DOUBLE", "10", "3", "", ""], 
            ["NumStry2", "DOUBLE", "10", "3", "", ""], 
            ["MaxDepth1", "DOUBLE", "10", "3", "", ""],  
            ["MaxDepth2", "DOUBLE", "10", "3", "", ""],
            ["ArrTime1", "DOUBLE", "10", "3", "", ""],
            ["ArrTime2", "DOUBLE", "10", "3", "", ""],
            ["EvacTime1", "DOUBLE", "10", "3", "", ""],
            ["EvacTime2", "DOUBLE", "10", "3", "", ""],
            ["MaxVelo1", "DOUBLE", "10", "3", "", ""],
            ["MaxVelo2", "DOUBLE", "10", "3", "", ""],
            ["PAR_1", "DOUBLE", "10", "3", "", ""],
            ["PAR_2", "DOUBLE", "10", "3", "", ""],
            ["LL_mean_1", "DOUBLE", "10", "3", "", ""],
            ["LL_mean_2", "DOUBLE", "10", "3", "", ""],
            ["Tot_Dam_1", "DOUBLE", "15", "3", "", ""],
            ["Tot_Dam_2", "DOUBLE", "15", "3", "", ""],
            ["IncAvDepth", "DOUBLE", "10", "3", "", ""],  # BEGIN incremental fields
            ["IncAvArriv", "DOUBLE", "10", "3", "", ""],
            ["IncEvacTi", "DOUBLE", "10", "3", "", ""],
            ["IncPAR", "DOUBLE", "10", "3", "", ""],
            ["IncAvLL", "DOUBLE", "10", "3", "", ""],
            ["IncAvDamg", "DOUBLE", "15", "3", "", ""]
        ]

        # List of fields added in the script, all other fields get deleted
        fields_to_keep = [
            "OBJECTID", "Shape", "Summary_Set_EPZ", "Emergency_Zone", #fields that you must keep
            "OccTyp1", "MaxDepth1", "NumStry1", "LL_mean_1", "Tot_Dam_1", "ArrTime1", "MaxVelo1", "EPZ1", "EvacTime1", #high scenario
            "OccTyp2", "MaxDepth2", "NumStry2", "LL_mean_2", "Tot_Dam_2", "ArrTime2", "MaxVelo2", "EPZ2", "EvacTime2", #low scenario
            "IncAvDepth","IncAvArriv", "IncAvLL", "IncPAR", "IncAvDamg", "IncEvacTi" #incremental
        ]

        # Add fields to the higher scenario feature class
        arcpy.management.AddFields(in_table=incremental_scenario, field_description=all_fields)
        messages.addMessage("Finished adding all fields to the incremental table.")

        # Define the field mapping for the higher and lower scenarios
        field_mapping = {
            "OccTyp1": "Occupancy_",  # Output field : Input field (Higher Scenario)
            "MaxDepth1": "Max_Depth",
            "NumStry1": "Number_of_", #Number_of_Stories
            "LL_mean_1": "Life_Loss9",  # Sum of two fields
            "Tot_Dam_1": ["Structure1", "Content__1", "Vehicle__1"],  # Sum of three fields
            "PAR_1": ["Pop_Under7", "PAR_Over65"],  # Sum of two fields
            "ArrTime1": "Time_To_Fi",
            "EvacTime1": "Time_To_No",  # Evacuation time field
            "MaxVelo1": "Max_Veloci",
            "EPZ1": "Emergency_",
            "OccTyp2": "Occupancy_",  # Output field : Input field (Lower Scenario)
            "MaxDepth2": "Max_Depth",
            "NumStry2": "Number_of_",
            "LL_mean_2": "Life_Loss9",  # Sum of two fields
            "Tot_Dam_2": ["Structure1", "Content__1", "Vehicle__1"],  # Sum of three fields
            "PAR_2": ["Pop_Under7", "PAR_Over65"],  # Sum of two fields
            "ArrTime2": "Time_To_Fi",
            "EvacTime2": "Time_To_No",  # Evacuation time field
            "MaxVelo2": "Max_Veloci",
            "EPZ2": "Emergency_",
        }
        
        # Define incremental fields
        incremental_fields = {
            "IncAvDepth": ("MaxDepth1", "MaxDepth2"),  # Output field : (Field1, Field2)
            "IncAvArriv": ("ArrTime1", "ArrTime2"),
            "IncEvacTi": ("EvacTime1", "EvacTime2"),
            "IncAvLL": ("LL_mean_1", "LL_mean_2"),
            "IncPAR": ("PAR_1", "PAR_2"),
            "IncAvDamg": ("Tot_Dam_1", "Tot_Dam_2"),
        }
        
        # Use cursors to calculate the fields
        def calculate_field_values(update_row, lower_row, higher_row, field_mapping, incremental_fields):
            """Helper function to calculate field values."""
            # Map fields from the input to the output using the field mapping
            for output_field, input_field in field_mapping.items():
                if output_field.endswith("1"):
                    # Fields ending in '1' are calculated from the higher_scenario
                    if isinstance(input_field, list):  # If the input is a list of fields, sum their values
                        update_row[output_field] = sum(
                            higher_row.get(f, 0) if f not in ["Time_To_Fi", "Time_To_No"] or higher_row.get(f, 0) <= 1_000_000 else 0
                            for f in input_field
                        )
                    else:
                        value = higher_row.get(input_field, 0)
                        # Apply the check for both Time_To_Fi and Time_To_No fields and convert to hours
                        if input_field in ["Time_To_Fi", "Time_To_No"]:
                            value = value / 60 if value <= 1_000_000 else 0
                        update_row[output_field] = value
                elif output_field.endswith("2"):
                    # Fields ending in '2' are calculated from the lower_scenario
                    if isinstance(input_field, list):  # If the input is a list of fields, sum their values
                        update_row[output_field] = sum(
                            lower_row.get(f, 0) if f not in ["Time_To_Fi", "Time_To_No"] or lower_row.get(f, 0) <= 1_000_000 else 0
                            for f in input_field
                        )
                    else:
                        value = lower_row.get(input_field, 0)
                        # Apply the check for both Time_To_Fi and Time_To_No fields and convert to hours
                        if input_field in ["Time_To_Fi", "Time_To_No"]:
                            value = value / 60 if value <= 1_000_000 else 0
                        update_row[output_field] = value

            # Calculate incremental fields
            for output_field, (field1, field2) in incremental_fields.items():
                # Handle NoneType values by using 0 as the default
                value1 = update_row.get(field1, 0) or 0
                value2 = update_row.get(field2, 0) or 0
                update_row[output_field] = value1 - value2

            return update_row


        # Precompute field lists
        output_fields = list(field_mapping.keys()) + list(incremental_fields.keys())
        input_fields = list(set(sum([v if isinstance(v, list) else [v] for v in field_mapping.values()], [])))

        # Use cursors to calculate the fields
        with arcpy.da.UpdateCursor(incremental_scenario, output_fields) as update_cursor, \
            arcpy.da.SearchCursor(lower_scenario, input_fields) as lower_cursor, \
            arcpy.da.SearchCursor(higher_scenario, input_fields) as higher_cursor:

            for update_row, lower_row, higher_row in zip(update_cursor, lower_cursor, higher_cursor):
                # Convert lower_row and higher_row to dictionaries for easier access
                lower_row_dict = {field: value for field, value in zip(input_fields, lower_row)}
                higher_row_dict = {field: value for field, value in zip(input_fields, higher_row)}

                # Calculate field values
                update_row_dict = {field: value for field, value in zip(output_fields, update_row)}
                update_row_dict = calculate_field_values(update_row_dict, lower_row_dict, higher_row_dict, field_mapping, incremental_fields)

                # Update the row
                update_cursor.updateRow([update_row_dict[field] for field in output_fields])

        # Clean up the cursors
        del update_cursor, lower_cursor, higher_cursor
        messages.addMessage("Finished combining attributes and calculating incremental fields.")

        def delete_zero_depth_rows(incremental_scenario, messages):
            """Delete rows where both MaxDepth1 and MaxDepth2 are 0."""
            fields = ["MaxDepth1", "MaxDepth2"]
            with arcpy.da.UpdateCursor(incremental_scenario, fields) as cursor:
                rows_deleted = 0
                for row in cursor:
                    if row[0] == 0 and row[1] == 0:  # Check if both MaxDepth1 and MaxDepth2 are 0
                        cursor.deleteRow()
                        rows_deleted += 1
                messages.addMessage(f"Deleted {rows_deleted} rows where both MaxDepth1 and MaxDepth2 are 0.")

        # Call the function to delete rows with zero depth if delete_zero_depths is True
        if delete_zero_depths:
            delete_zero_depth_rows(incremental_scenario, messages)

        # Get all fields in the feature class
        all_fields = arcpy.ListFields(incremental_scenario)

        # Identify fields to delete (those not in fields_to_keep)
        fields_to_delete = [field.name for field in all_fields if field.name not in fields_to_keep]

        # Delete the fields
        #fields_to_delete = []  # Uncomment this line to test without deleting fields
        if fields_to_delete:
            #messages.addMessage(f"Deleting fields: {fields_to_delete}")
            messages.addMessage(f"Deleting fields.....")
            arcpy.management.DeleteField(in_table=incremental_scenario, drop_field=fields_to_delete)
        else:
            messages.addMessage("No fields to delete.")

        messages.addMessage("Success. Note in the incremental file the times have been converted from minutes to HOURS.")
        
        return
#---------------------------------------------------------------