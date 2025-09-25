# -*- coding: utf-8 -*-
#-------------------------------------------------------------------------
# Author:     Kurt Buchanan, CELRH
# Created:     Sep 2024
# ArcGIS Version:   2.9.3
# Requirements: ?
# Revisions: 
    #2024-09-17 Initial Draft
    #2025-04-29 Added arrival, depth, velocity ranges, fixed export to gpkg
    #2025-07-27 added structure inventory checks
    #2025-08-27 fixed queries on structure inventory and area names to allow for a - or ' or whatever wierd character in them
    #2025-09-22 added capability to include life loss on road networks, analyzes road networks for vertical offsets
#-------------------------------------------------------------------------
import arcpy
import os, math
from arcpy.sa import *
from datetime import datetime
import sqlite3
import re
from openpyxl import Workbook
from openpyxl.styles import Font
import shutil


class Toolbox(object):
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the
        .pyt file)."""
        self.label = "MMC Toolbox"
        self.alias = "mmctoolbox"
        self.description = "ArcGIS python toolbox created to support the USACE MMC Consequences Team."

        # List of tool classes associated with this toolbox
        self.tools = [Lifesim1]

##-------------------------------------------------------------------------------------

class Lifesim1(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "LifeSim Results Summary"
        self.description = "Does lots of stuff."
        self.canRunInBackground = True

    def getParameterInfo(self):
        """Define parameter definitions"""
        inputlifesimfile = arcpy.Parameter(
            displayName="LifeSim File",
            name="study area",
            datatype="DEFile",
            parameterType="Required",
            direction="Input")
        simulation1 = arcpy.Parameter(
            displayName="Simulation",
            name="Simulation",
            datatype="GPString",
            parameterType="Required",
            direction="Input")

        arrival_option = arcpy.Parameter(
            displayName="Arrival Time Field",
            name="Arrival_Time_Field",
            datatype="GPString",
            parameterType="Required",
            direction="Input")
        arrival_option.filter.type = "ValueList"
        arrival_option.filter.list = ["Time_To_First_Wet", "Time_To_No_Evac"]
        arrival_option.value = "Time_To_First_Wet"  # default value
        range_option = arcpy.Parameter(
            displayName="Range Percentiles",
            name="Range_Percentiles",
            datatype="GPString",
            parameterType="Required",
            direction="Input")
        range_option.filter.type = "ValueList"
        range_option.filter.list = ["25th and 75th", "15th and 85th"]
        range_option.value = "25th and 75th"  # default value

        output_excel_file = arcpy.Parameter(
            displayName="Output Excel File",
            name="Output_Excel_File",
            datatype="DEFile",
            parameterType="Required",
            direction="Output")
        
        alternative1 = arcpy.Parameter(
            displayName="Alternative (optional)",
            name="Alternative",
            datatype="GPString",
            parameterType="Optional",
            direction="Input")
        areas1 = arcpy.Parameter(
            displayName="Summary Areas (optional)",
            name="Area",
            datatype="GPString",
            parameterType="Optional",
            direction="Input")
        mmc_sop = arcpy.Parameter(
            displayName="Check to flag MMC SOP violations",
            name="Check to flag MMC SOP violations",
            datatype="GPBoolean",
            parameterType="Optional",
            direction="Input")
        mmc_sop.value = True
        exportresults = arcpy.Parameter(
            displayName="Check to export alternative results to geopackage (experimental, no decimals)",
            name="Check to export results",
            datatype="GPBoolean",
            parameterType="Optional",
            direction="Input")

        parameters = [inputlifesimfile, simulation1, arrival_option, range_option, output_excel_file, alternative1, areas1, mmc_sop, exportresults]
        #                0                  1              2             3               4                 5           6          7       8
        return parameters

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        #Lookups for Data list selections
        import os
        if parameters[0].value:
            con = sqlite3.connect(fr"{parameters[0].value}")
            
            ### Lookup for Simulation List
            cursor2 = con.cursor()
            # Query the "Name" field from the Simulations_Lookup_Table
            cursor2.execute("SELECT Name FROM Simulations_Lookup_Table;")
            # Flatten the list of tuples to get the "Name" field values
            nameList2 = [row[0] for row in cursor2.fetchall()]
            # Assign the list of "Name" values to the parameter filter
            parameters[1].filter.list = nameList2

            ### Lookup for Alternatives
            cursor = con.cursor()
            # Query the "Name" field from the Alternatives_Lookup_Table
            cursor.execute("SELECT Name FROM Alternatives_Lookup_Table;")
            # Flatten the list of tuples to get the "Name" field values
            nameList = [row[0] for row in cursor.fetchall()]
            # Assign the list of "Name" values to the parameter filter
            parameters[5].filter.list = nameList

            if parameters[0].valueAsText:  # Use .valueAsText to get the file path as a string
                lifesim_file_path = parameters[0].valueAsText  # Get the LifeSim file path as a string
                lifesim_file_name = os.path.basename(lifesim_file_path)  # Get the LifeSim file name
                lifesim_base_name = os.path.splitext(lifesim_file_name)[0]  # Remove the extension

            # Generate default Excel file name if a simulation is selected
                if parameters[1].value:  # Check if a simulation is selected and output hasn't altered
                    simulation_name = parameters[1].value  # Get the selected simulation name
                    default_excel_name = f"{lifesim_base_name}_{simulation_name}_Results.xlsx"
                    # Update the output Excel file name if it hasn't been manually altered
                if not parameters[4].altered or parameters[1].altered:
                    parameters[4].value = os.path.join(os.path.dirname(lifesim_file_path), default_excel_name)

        if parameters[1].value:
            con = sqlite3.connect(fr"{parameters[0].value}")
            
            ### Lookup for Simulation List
            cursor3 = con.cursor()
            # Query the "Name" field from the Simulations_Lookup_Table
            cursor3.execute("SELECT Summary_Name_Fields FROM Simulations_Lookup_Table WHERE Name = ?", (parameters[1].value,))
            results = cursor3.fetchall()
            # Flatten the list of tuples to get the "Summary Name Fields" values
            options = str([row[0] for row in results])
            pattern = re.compile(r'\[.*?\]')
            matches = pattern.findall(options)
            # Assign the list of "Summary Name" values to the parameter filter
            list3 = []
            for match in matches:
                list3.append(match.strip())
            chars_to_remove = ["'", "[", "]"]
            cleaned_list = [element.translate(str.maketrans('', '', ''.join(chars_to_remove))) for element in list3]
            parameters[6].filter.list = cleaned_list


        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return
       
    def execute(self, parameters, messages):
        """The source code of the tool."""
        #call parameters that were entered and create new variables as needed
        inputlifesimfile = parameters[0].valueAsText
        simulation1 = parameters[1].valueAsText
        arrival_column = parameters[2].valueAsText
        range_option = parameters[3].valueAsText
        output_excel_file = parameters[4].valueAsText
        single_alternative = parameters[5].valueAsText
        areafile = parameters[6].valueAsText
        #boolean parameters should just be the value, not text, like below
        mmc_sop = parameters[7].value
        export = parameters[8].value
        
        # Exports structure summary results to a new output geopackage if true
        # Allows for overwriting the fia gpkg and/or output gpkg if set to true, otherwise it looks for those and skips if they exist
        if export:
            overwrite_tempFIAgpkg = True
            overwrite_outputgpkg = True
        
        #EPZ Filter allows the arrival queries to filter to only one EPZ if you need that data, replace 'North' with whatever epz name
        epzfilter = ""
        #epzfilter = " AND Emergency_Zone = 'North'"

        # Set the percentile ranges, some cursors executes with a different percentile, like 0.15 or 0.85
        if range_option == "25th and 75th":
            depthlowpercentile = 0.25
            depthhighpercentile = 0.75
            arrivallowpercentile = 0.25
            arrivalhighpercentile = 0.75
            velocitylowpercentile = 0.25
            velocityhighpercentile = 0.75
        if range_option == "15th and 85th":
            depthlowpercentile = 0.15
            depthhighpercentile = 0.85
            arrivallowpercentile = 0.15
            arrivalhighpercentile = 0.85
            velocitylowpercentile = 0.15
            velocityhighpercentile = 0.85


        #Create the excel workbook
        wb = Workbook()
        ws = wb.create_sheet(title="Summary")
        # Remove the default sheet created at Workbook instantiation
        default_sheet = wb['Sheet']
        wb.remove(default_sheet)

        # helper to create safe Excel sheet names (max 31 chars, cannot contain: : \\ / ? * [ ] )
        def sanitize_sheet_name(name):
            # replace forbidden chars with underscore
            sanitized = re.sub(r'[:\\/\?\*\[\]]', '_', str(name))
            # Excel limits sheet names to 31 characters
            sanitized = sanitized[:31]
            # If the name becomes empty, use a fallback
            if not sanitized:
                sanitized = 'Sheet'
            return sanitized

        # Lists to store all the message values, list1 is the SUmmary tab, list2 is in alternatives
        message_list1 = []
        message_list2 = []
        terrainlist = []
        inventorylist = []
        # mapping of original alternative name -> sanitized sheet name
        alt_to_sheet_map = []
        # function to add messages as both normal messages and append them to the lists
        def add_message(msg, lst):
            messages.addmessage(msg)
            if lst == 1:
                message_list1.append(msg)
            else:
                message_list2.append(msg)

        #Add date created/ran
        date = datetime.now()
        formatted_date = date.strftime("%Y-%m-%d %H:%M:%S")
        add_message("Date created: {}".format(formatted_date), lst=1)

        # Parse output folder and file name from input file
        outputfolder1 = os.path.dirname(os.path.abspath(inputlifesimfile))
        lifesimfilename = os.path.splitext(os.path.basename(inputlifesimfile))[0]
        add_message("Folder Name: {}".format(outputfolder1), lst=1)
        add_message("LifeSim File Name: {}.fia".format(lifesimfilename), lst=1)
        add_message("Arrival Time Column: {}".format(arrival_column), lst=1)
        add_message("Range Percentiles: {}".format(range_option), lst=1)
        add_message("Simulation Name: {}".format(simulation1), lst=2)

        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 60
        ws.column_dimensions['C'].width = 20
        ws.column_dimensions['D'].width = 15
        ws.column_dimensions['E'].width = 15
        ws.column_dimensions['F'].width = 15
        ws.column_dimensions['G'].width = 15
        ws.column_dimensions['H'].width = 15
        ws.column_dimensions['I'].width = 15
        ws.column_dimensions['J'].width = 15
        ws.column_dimensions['K'].width = 15
        ws.column_dimensions['L'].width = 15
        ws.column_dimensions['M'].width = 15
        ws.column_dimensions['N'].width = 15
        ws.column_dimensions['O'].width = 15
        ws.column_dimensions['P'].width = 15

        # excel For total table (first table)
        start_row_summary = 1
        start_col_summary = 3

        # excel Define where the tables should start for totals (e.g., row 1, column 3)
        row = start_row_summary
        col = start_col_summary
        # Set headers for columns in total table
        ws.cell(row=row, column=col, value="Alternative")
        ws.cell(row=row, column=col+1, value="Structure #")
        ws.cell(row=row, column=col+2, value="PAR Day")
        ws.cell(row=row, column=col+3, value="PAR Night")
        ws.cell(row=row, column=col+4, value="LL Day")
        ws.cell(row=row, column=col+5, value="LL Night")
        ws.cell(row=row, column=col+6, value="Damage Total")
        ws.cell(row=row, column=col+8, value="LL Evac Day")
        ws.cell(row=row, column=col+9, value="LL Evac Night")
        # excel Increment row for EPZ after the column headings
        current_row_summary = start_row_summary + 1

        # create name variables for exporting to geopackage
        tempFIAgpkg=fr"{outputfolder1}\{lifesimfilename}_gpkgtemp.gpkg"
        tempFIAgpkg_nameonly=os.path.splitext(os.path.basename(tempFIAgpkg))[0]
        outputgpkg=fr"{outputfolder1}\Sim_{simulation1}_results.gpkg"
        outputgpkg_nameonly=os.path.splitext(os.path.basename(outputgpkg))[0]

        if export:

            # Check if tempFIAgpkg already exists
            if not os.path.exists(tempFIAgpkg) or overwrite_tempFIAgpkg:
                # Directly copy the .fia file to a new file with a .gpkg extension
                shutil.copy(inputlifesimfile, tempFIAgpkg)

                if overwrite_tempFIAgpkg:
                    add_message(f"{tempFIAgpkg_nameonly} overwritten.", lst=1)
                else:
                    add_message(f"Created temp gpkg file: {tempFIAgpkg_nameonly}", lst=1)


            else:
                add_message(f"{tempFIAgpkg_nameonly} already exists, skipping creation.", lst=1)

            # Check if output gpkg already exists
            if not os.path.exists(outputgpkg) or overwrite_outputgpkg:
                with arcpy.EnvManager(outputCoordinateSystem='PROJCS["WGS_1984_Web_Mercator_Auxiliary_Sphere",GEOGCS["GCS_WGS_1984",DATUM["D_WGS_1984",SPHEROID["WGS_1984",6378137.0,298.257223563]],PRIMEM["Greenwich",0.0],UNIT["Degree",0.0174532925199433]],PROJECTION["Mercator_Auxiliary_Sphere"],PARAMETER["False_Easting",0.0],PARAMETER["False_Northing",0.0],PARAMETER["Central_Meridian",0.0],PARAMETER["Standard_Parallel_1",0.0],PARAMETER["Auxiliary_Sphere_Type",0.0],UNIT["Meter",1.0]]'):
                    arcpy.management.CreateSQLiteDatabase(
                        out_database_name=outputgpkg, spatial_type="GEOPACKAGE")
                if overwrite_outputgpkg:
                    add_message(f"{outputgpkg_nameonly} overwritten.", lst=1)
                else:
                    add_message(f"Created output gpkg file: {outputgpkg_nameonly}", lst=1)
            else:
                add_message(f"{outputgpkg_nameonly} already exists, skipping creation.", lst=1)

            ## Finished file creation

        #Set sql cursor to read .fia file
        con = sqlite3.connect(fr"{inputlifesimfile}")
        cursor = con.cursor()

        # Lookup number of iterations from simulation
        cursor.execute("SELECT Iterations FROM Simulations_Lookup_Table WHERE Name = ?", (parameters[1].value,))
        numiterations = cursor.fetchall()[0][0]  # Access the first element in the tuple
        add_message("Number of Iterations: {}".format(numiterations), lst=1)
        if numiterations <1000 and mmc_sop == True:
            add_message("MMC SOP Violation: Need at least 1000 iterations.", lst=1)

        # Get Alternative List from simulation
        alternative_names_query = f"SELECT Alternatives FROM Simulations_Lookup_Table WHERE Name = ?;"
        cursor.execute(alternative_names_query, (simulation1,))
        alternative_names_string = cursor.fetchone()[0]
        alternative_names_list = alternative_names_string.split(',')
        alternative_names_list = [name.strip() for name in alternative_names_list]
        # if a single alternative is selected, this sets the list to just be one alternative
        if single_alternative:
            alternative_names_list = [single_alternative]
        add_message(f"List of alternatives in this simulation:", lst=1)
        for alternative in alternative_names_list:
            add_message(f"   {alternative}", lst=1)

        # Query unique Road_Network names from Alternatives_Lookup_Table, check for vertical offsets and bridges
        cursor.execute("SELECT DISTINCT Road_Network FROM Alternatives_Lookup_Table WHERE Road_Network IS NOT NULL;")
        #unique_road_networks = [row[0] for row in cursor.fetchall()]
        raw_networks = [row[0] for row in cursor.fetchall()]
        # filter out None, empty strings and strings containing only whitespace; trim names
        unique_road_networks = [r.strip() for r in raw_networks if r and str(r).strip()]
        # ensure we always have a list (empty if nothing valid found)
        if not unique_road_networks:
            unique_road_networks = []
        add_message(f"Uniqure road network list:  {unique_road_networks}", lst=1)
        add_message(f"Uniqure road network list length:  {len(unique_road_networks)}", lst=1)
        if len(unique_road_networks) > 0:
            add_message(f"At least one Alternative has a road network selected.", lst=1)
            for road_network in unique_road_networks:
                add_message(f"Road Network in alternative list: {road_network}", lst=1)
                vert_offset_query = f"SELECT COUNT(*) FROM '{road_network}' WHERE VertOffset > 0"
                cursor.execute(vert_offset_query)
                vertical_offset_count = cursor.fetchone()[0]
                add_message(f"   Number of vertical offsets in this road network: {vertical_offset_count}", lst=1)
                if vertical_offset_count == 0:
                    add_message(f"POSSIBLE ISSUE: No vertical offsets found in road network {road_network}.", lst=1)
                number_bridges_query = f"SELECT COUNT(*) FROM '{road_network}' WHERE bridge = 'yes'"
                cursor.execute(number_bridges_query)
                number_bridges_count = cursor.fetchone()[0]
                add_message(f"   Number of bridges in this road network: {number_bridges_count}", lst=1)
                bridges_with_no_offset_query = f"SELECT COUNT(*) FROM '{road_network}' WHERE bridge = 'yes' AND VertOffset = 0"
                cursor.execute(bridges_with_no_offset_query)
                bridges_with_no_offset_count = cursor.fetchone()[0]
                add_message(f"   Number of bridges without a vertical offset: {bridges_with_no_offset_count}", lst=1)
                if bridges_with_no_offset_count > 0:
                    add_message(f"POSSIBLE ISSUE: There are {bridges_with_no_offset_count} bridges with no vertical offsets in {road_network}.", lst=1)
                
        

        ## Get Summary Polygon List (similar to the logic in updateparameters)
        cursor.execute("SELECT Summary_Name_Fields FROM Simulations_Lookup_Table WHERE Name = ?", (parameters[1].value,))
        results = cursor.fetchall()
        # Flatten the list of tuples to get the "Name" field values
        options = str([row[0] for row in results])
        pattern = re.compile(r'\[.*?\]')
        matches = pattern.findall(options)
        # Assign the list of "Name" values to the parameter filter
        list3 = []
        for match in matches:
            list3.append(match.strip())
        chars_to_remove = ["'", "[", "]"]
        cleaned_list = [element.translate(str.maketrans('', '', ''.join(chars_to_remove))) for element in list3]
        
        #option below removes some summary areas from the list if needed, indices are the rows to remove and need to be changed as needed
        remove_problem_areas = False #change to true and adjust if you need to remove some areas
        if remove_problem_areas:
            add_message(f"NOTE: Some summary polygons are being removed from the query, only ones listed below are used", lst=1)
            if not areafile:
                indices_to_remove = [0, 2, 3]
                cleaned_list = [item for i, item in enumerate(cleaned_list) if i not in indices_to_remove]

        # if a single area is selected, cleaned list gets set to just that one selection
        if areafile:
            cleaned_list = [areafile]
        add_message(f"List of summary polygons and name fields in this simulation:", lst=1)
        for polygonset in cleaned_list:
            add_message(f"   {polygonset}", lst=1)

        # Initialize the progress bar
        total_alternatives = len(alternative_names_list)
        arcpy.SetProgressor("step", "Processing alternatives...", 0, total_alternatives, 1)
        alternative_number = 1

        # LOOP A startlooking at alternative data in a loop
        for alternative1 in alternative_names_list:

            # Update the progress bar with the alternative name and its position
            arcpy.SetProgressorLabel(f"Processing alternative: {alternative1} ({alternative_number} of {total_alternatives})")
            arcpy.SetProgressorPosition()
            alternative_number += 1

            #reset message_list2 to blank
            message_list2 = []
            add_message(f".........................", lst=2)
            add_message("Alternative: {}".format(alternative1), lst=2)
            add_message("Simulation: {}".format(simulation1), lst=2)
            
            # checks for anything that makes this look like a non-fail scenario
            if any(s in alternative1.lower() for s in ("nb", "nonbreach", "non-breach", "nofail", "nonfail", "non-fail")):
                breach_condition = "NonFail"
            else:
                breach_condition = "Fail"
            add_message("Estimated alternative type: {}".format(breach_condition), lst=2)
            
            # create variables for Names of Day and Night structure summary tables
            results1_day_input = fr'"{simulation1}>Structure_Summary>{alternative1}>14"'
            results1_night_input = fr'"{simulation1}>Structure_Summary>{alternative1}>2"'

            # create variables for Names of Day and Night roads summary tables
            results1_day_roads = fr'"{simulation1}>Roads_Summary>{alternative1}>14"'
            results1_night_roads = fr'"{simulation1}>Roads_Summary>{alternative1}>2"'

            #sets up logic to be able to copy and export GIS data
            if export:
                results1_day_input_clean = results1_day_input.replace('"', '').replace("'", "")
                tablename_clean_day=fr"{tempFIAgpkg}\main.{results1_day_input_clean}"
                results1_night_input_clean = results1_night_input.replace('"', '').replace("'", "")
                tablename_clean_night=fr"{tempFIAgpkg}\main.{results1_night_input_clean}"
                
                #sanitize function ensures no wierd characters are in the output name
                def sanitize_table_name(name):
                    # Replace any character that is not alphanumeric or underscore with underscore
                    sanitized_name = re.sub(r'[^a-zA-Z0-9_]', '_', name)
                    # Strip leading and trailing underscores
                    sanitized_name = sanitized_name.strip('_')
                    return sanitized_name

                sanitized_output_day = sanitize_table_name(results1_day_input)
                sanitized_output_night = sanitize_table_name(results1_night_input)
                
            # startlooking at alternative data
            # create a sanitized Excel worksheet name for the alternative and set column widths
            sanitized_sheet = sanitize_sheet_name(alternative1)
            # store mapping for later summary
            alt_to_sheet_map.append((alternative1, sanitized_sheet))
            # Avoid duplicate sheet names by appending a counter if needed
            base_name = sanitized_sheet
            counter = 1
            while sanitized_sheet in wb.sheetnames:
                # leave room for counter at end
                suffix = f"_{counter}"
                allowed_len = 31 - len(suffix)
                sanitized_sheet = base_name[:allowed_len] + suffix
                counter += 1
            ws = wb.create_sheet(title=sanitized_sheet)
            ws.column_dimensions['A'].width = 30
            ws.column_dimensions['B'].width = 80
            ws.column_dimensions['C'].width = 20
            ws.column_dimensions['D'].width = 15
            ws.column_dimensions['E'].width = 15
            ws.column_dimensions['F'].width = 15
            ws.column_dimensions['G'].width = 15
            ws.column_dimensions['H'].width = 15
            ws.column_dimensions['I'].width = 15
            ws.column_dimensions['J'].width = 15
            ws.column_dimensions['K'].width = 15
            ws.column_dimensions['L'].width = 15
            ws.column_dimensions['M'].width = 15
            ws.column_dimensions['N'].width = 15
            ws.column_dimensions['O'].width = 15
            ws.column_dimensions['P'].width = 15
            ws.column_dimensions['Q'].width = 15
            ws.column_dimensions['R'].width = 15
            ws.column_dimensions['S'].width = 15
            ws.column_dimensions['T'].width = 15

            
            # query hydraulic data to get scenario, start time, and hazard time
            hydscenario_query = "SELECT Hydraulic_Scenario FROM Alternatives_Lookup_Table WHERE Name = ?;"
            cursor.execute(hydscenario_query, (alternative1,))
            hydscenario = cursor.fetchone()[0]
            add_message("Hydraulic scenario: {}".format(hydscenario), lst=2)

            # Check if hydscenario is part of alternative name, normalize both strings for comparison
            def normalize_name(name):
                # Lowercase, replace underscores and hyphens with spaces, collapse multiple spaces
                return ' '.join(name.lower().replace('_', ' ').replace('-', ' ').split())

            # After you get hydscenario and alternative1:
            normalized_hydscenario = normalize_name(hydscenario)
            normalized_alternative = normalize_name(alternative1)

            if normalized_hydscenario not in normalized_alternative:
                add_message(f"POSSIBLE ISSUE: Hydraulic scenario name '{hydscenario}' is not part of the Alternative name '{alternative1}'.", lst=2)
            else:
                add_message(f"Hydraulic scenario name '{hydscenario}' is included in Alternative name '{alternative1}'.", lst=2)


            hydstart_query = "SELECT Start_Time FROM Hydraulic_Data_Lookup_Table WHERE Name = ?;"
            cursor.execute(hydstart_query, (hydscenario,))
            hydstart = cursor.fetchone()[0]

            haztime_query = "SELECT Imminent_Hazard_Time FROM Hydraulic_Data_Lookup_Table WHERE Name = ?;"
            cursor.execute(haztime_query, (hydscenario,))
            haztime = cursor.fetchone()[0]

            #logic for identifying time formats, LifeSim 2.1.3 stores different than 2.1.4 due to a timing bug fix
            def parse_date(date_string):
                # Remove the 'Z' if it exists at the end
                if date_string.endswith('Z'):
                    date_string = date_string[:-1]
                # If there are more than 6 digits in the fractional seconds, truncate to 6 digits
                if '.' in date_string:
                    main_part, fractional_part = date_string.split('.')
                    if len(fractional_part) > 6:
                        # Only keep the first 6 digits of the fractional seconds
                        fractional_part = fractional_part[:6]
                    date_string = main_part + '.' + fractional_part
                formats = [
                    "%Y-%m-%dT%H:%M:%S.%f",  # Format with fractional seconds
                    "%Y-%m-%dT%H:%M:%S",     # Format without fractional seconds
                    "%Y-%m-%d %H:%M:%S",     # Format without 'T'
                    "%m/%d/%Y %I:%M %p"      # New format with AM/PM notation for LifeSim 2.0
                ]
                for timeformat in formats:
                    try:
                        return datetime.strptime(date_string, timeformat)
                    except ValueError:
                        continue
        
                raise ValueError(f"Time data '{date_string}' does not match any of the expected formats")

            #format time from start to hazard
            hydstart_formatted = parse_date(hydstart)
            haztime_formatted = parse_date(haztime)
            BreachTime = parse_date(haztime) - parse_date(hydstart)
            BreachTime_in_hours = BreachTime.total_seconds() / 3600
            BreachTime_in_minutes = BreachTime.total_seconds() / 60
        
            add_message("Hydraulic Start Time: {}".format(hydstart_formatted), lst=2)
            add_message("Imminent Hazard Time: {}".format(haztime_formatted), lst=2)
            add_message("Hazard Time from start: {}".format(BreachTime), lst=2)
            add_message("Hazard Time from start in hours: {}".format(round(BreachTime_in_hours,3)), lst=2)
            add_message("Hazard Time from start in minutes: {}".format(BreachTime_in_minutes), lst=2)

            #query terrain source
            terrainpath_query = "SELECT Terrain_Path FROM Hydraulic_Data_Lookup_Table WHERE Name = ?;"
            cursor.execute(terrainpath_query, (hydscenario,))
            terrainpath = cursor.fetchone()[0]
            add_message("Terrain path: {}".format(terrainpath), lst=2)
            if terrainpath not in terrainlist:
                terrainlist.append(terrainpath)
            
            #query structure source selection
            structure_source_query = f"SELECT Structure_Inventory FROM Alternatives_Lookup_Table WHERE Name = ?;"
            cursor.execute(structure_source_query, (alternative1,))
            structure_source = cursor.fetchone()[0]
            add_message("Structure source: {}".format(structure_source), lst=2)
            if structure_source not in inventorylist:
                inventorylist.append(structure_source)         

            #query total results, use try-except on first one in case a selected alternative-sim combo doesn't exist
            # If the table is missing, skip this alternative and continue the loop so the workbook still gets written
            total_structurecount_day_query = f"SELECT COUNT(*) FROM {results1_day_input} WHERE Max_Depth > 0"
            skip_alternative = False
            try:
                cursor.execute(total_structurecount_day_query,)
                total_structurecount_day = cursor.fetchone()[0]
            except sqlite3.OperationalError as e:
                # SQLite reports missing table as OperationalError containing 'no such table'
                if 'no such table' in str(e).lower():
                    add_message(f"WARNING: Results table not found for alternative '{alternative1}': {results1_day_input}. Skipping this alternative.", lst=2)
                    skip_alternative = True
                else:
                    # re-raise unexpected OperationalErrors
                    raise
            except Exception as e:
                # For any other unexpected exceptions, re-raise as before
                raise
            # If initial query succeeded, double-check that both day and night tables exist before running many further queries
            if not skip_alternative:
                # remove quotes from the table identifiers
                table_day_clean = results1_day_input.replace('"', '').replace("'", "")
                table_night_clean = results1_night_input.replace('"', '').replace("'", "")
                def table_exists(tbl_name):
                    try:
                        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name = ?;", (tbl_name,))
                        return cursor.fetchone() is not None
                    except Exception:
                        return False

                if not table_exists(table_day_clean) or not table_exists(table_night_clean):
                    add_message(f"WARNING: One or both results tables are missing for alternative '{alternative1}': {table_day_clean}, {table_night_clean}. Skipping this alternative.", lst=2)
                    skip_alternative = True
            if skip_alternative:
                try:
                    if sanitized_sheet in wb.sheetnames:
                        wb.remove(wb[sanitized_sheet])
                except Exception:
                    pass
                continue
                if skip_alternative:
                    # remove the worksheet created for this alternative since we're skipping it
                    try:
                        if sanitized_sheet in wb.sheetnames:
                            wb.remove(wb[sanitized_sheet])
                    except Exception:
                        pass
                    # continue to the next alternative in the outer loop
                    continue
            #add_message("Total structures wet (day): {}".format(total_structurecount_day), lst=2)

            # query to see if evacuation was modeled, if yes then check road network
            evacuation_query = "SELECT Simulate_Evacuation FROM Alternatives_Lookup_Table WHERE Name = ?;"
            cursor.execute(evacuation_query, (alternative1,))
            evacuation_simulated = cursor.fetchone()[0]
            if evacuation_simulated == 1:
                evacuation_simulated = True
            if evacuation_simulated:
                road_network_query = "SELECT Road_Network FROM Alternatives_Lookup_Table WHERE Name = ?;"
                cursor.execute(road_network_query, (alternative1,))
                road_network = cursor.fetchone()[0]
                add_message(f"Evacuation simulated with road network: {road_network}", lst=2)
                # create variables for Names of Day and Night roads summary tables
                results1_day_roads = fr'"{simulation1}>Roads_Summary>{alternative1}>14"'
                results1_night_roads = fr'"{simulation1}>Roads_Summary>{alternative1}>2"'
                # number of bridges that get flooded higher than offset
                bridges_still_flooded_day_query = f"SELECT COUNT(*) FROM {results1_day_roads} WHERE Vertical_Offset > 0 AND Max_Depth_ft > 0"
                cursor.execute(bridges_still_flooded_day_query)
                bridges_still_flooded_day = cursor.fetchone()[0]
                add_message(f" Roads with vertical offset where flood is higher than offset: {bridges_still_flooded_day}", lst=2)
                # life loss on roads with offsets
                bridges_still_flooded_wLL_day_query = f"SELECT COUNT(*) FROM {results1_day_roads} WHERE Vertical_Offset > 0 AND Max_Depth_ft > 0 AND Life_Loss_Mean > 0"
                cursor.execute(bridges_still_flooded_wLL_day_query)
                bridges_still_flooded_wLL_day = cursor.fetchone()[0]
                add_message(f" Roads with vertical offset where flood is higher and there is life loss: {bridges_still_flooded_wLL_day}", lst=2)
                if bridges_still_flooded_wLL_day > 0:
                    bridges_still_flooded_LLsum_day_query = f"SELECT SUM(Life_Loss_Mean) FROM {results1_day_roads} WHERE Vertical_Offset > 0 AND Max_Depth_ft > 0 AND Life_Loss_Mean > 0"
                    cursor.execute(bridges_still_flooded_LLsum_day_query)
                    bridges_still_flooded_LLsum_day = cursor.fetchone()[0]
                    add_message(f" Total Life Loss on Roads with vertical offset where flood is higher: {bridges_still_flooded_LLsum_day}", lst=2)

            # query total results
            total_paru65_day_query = f"SELECT SUM(Pop_Under65_Mean) FROM {results1_day_input} WHERE Max_Depth > 0"
            cursor.execute(total_paru65_day_query,)
            total_paru65_day = cursor.fetchone()[0] or 0
            #add_message("Total PAR under65 (day): {}".format(total_paru65_day), lst=2)

            total_paro65_day_query = f"SELECT SUM(PAR_Over65_Mean) FROM {results1_day_input} WHERE Max_Depth > 0"
            cursor.execute(total_paro65_day_query,)
            total_paro65_day = cursor.fetchone()[0] or 0
            #add_message("Total PAR over65 (day): {}".format(total_paro65_day), lst=2)

            total_par_day = total_paru65_day + total_paro65_day
            #add_message("Total PAR (day): {}".format(total_par_day), lst=2)

            total_lifeloss_st_day_query = f"SELECT SUM(Life_Loss_Total_Mean) FROM {results1_day_input}"
            cursor.execute(total_lifeloss_st_day_query,)
            total_lifeloss_st_day = cursor.fetchone()[0] or 0
            #add_message("Total life loss st (day): {}".format(total_lifeloss_st_day), lst=2)

            total_lifeloss_evac_day_query = f"SELECT SUM(Life_Loss_Evacuating_Mean) FROM {results1_day_input}"
            cursor.execute(total_lifeloss_evac_day_query,)
            total_lifeloss_evac_day = cursor.fetchone()[0] or 0
            #add_message("Total life loss evac (day): {}".format(total_lifeloss_evac_day), lst=2)
            
            total_lifeloss_day = total_lifeloss_st_day + total_lifeloss_evac_day
            #add_message("Total life loss (day): {}".format(total_lifeloss_day), lst=2)

            total_paru65_night_query = f"SELECT SUM(Pop_Under65_Mean) FROM {results1_night_input} WHERE Max_Depth > 0"
            cursor.execute(total_paru65_night_query,)
            total_paru65_night = cursor.fetchone()[0] or 0
            #add_message("Total PAR under65 (night): {}".format(total_paru65_night), lst=2)

            total_paro65_night_query = f"SELECT SUM(PAR_Over65_Mean) FROM {results1_night_input} WHERE Max_Depth > 0"
            cursor.execute(total_paro65_night_query,)
            total_paro65_night = cursor.fetchone()[0] or 0
            #add_message("Total PAR over65 (night): {}".format(total_paro65_night), lst=2)

            total_par_night = total_paru65_night + total_paro65_night
            #add_message("Total PAR (night): {}".format(total_par_night), lst=2)

            total_lifeloss_st_night_query = f"SELECT SUM(Life_Loss_Total_Mean) FROM {results1_night_input}"
            cursor.execute(total_lifeloss_st_night_query,)
            total_lifeloss_st_night = cursor.fetchone()[0] or 0
            #add_message("Total life loss st (night): {}".format(total_lifeloss_st_night), lst=2)

            total_lifeloss_evac_night_query = f"SELECT SUM(Life_Loss_Evacuating_Mean) FROM {results1_night_input}"
            cursor.execute(total_lifeloss_evac_night_query,)
            total_lifeloss_evac_night = cursor.fetchone()[0] or 0
            #add_message("Total life loss evac (night): {}".format(total_lifeloss_evac_night), lst=2)

            total_lifeloss_night = total_lifeloss_st_night + total_lifeloss_evac_night
            #add_message("Total life loss  (night): {}".format(total_lifeloss_night), lst=2)

            total_strucdamage_day_query = f"SELECT SUM(Structure_Damage_Mean) FROM {results1_day_input} WHERE Max_Depth > 0"
            cursor.execute(total_strucdamage_day_query,)
            total_strucdamage_day = cursor.fetchone()[0] or 0
            
            total_contdamage_day_query = f"SELECT SUM(Content_Damage_Mean) FROM {results1_day_input} WHERE Max_Depth > 0"
            cursor.execute(total_contdamage_day_query,)
            total_contdamage_day = cursor.fetchone()[0] or 0

            total_vehicdamage_day_query = f"SELECT SUM(Vehicle_Damage_Mean) FROM {results1_day_input} WHERE Max_Depth > 0"
            cursor.execute(total_vehicdamage_day_query,)
            total_vehicdamage_day = cursor.fetchone()[0] or 0

            total_damage_day = total_strucdamage_day + total_contdamage_day + total_vehicdamage_day

            # arrival zone queries, breaks and data in minutes, reports out in hours
            arrival_time_break2 = 30 #.5 hours
            arrival_time_break3 = 120 #2 hours
            arrival_time_break4 = 240 #4 hours
            arrival_time_break5 = 480 #8 hours
            arrival_time_break6 = 1440 #24 hours
            arrival_zone2_start = BreachTime_in_minutes
            arrival_zone2_end = BreachTime_in_minutes + arrival_time_break2
            arrival_zone3_start = BreachTime_in_minutes + arrival_time_break2
            arrival_zone3_end = BreachTime_in_minutes + arrival_time_break3
            arrival_zone4_start = BreachTime_in_minutes + arrival_time_break3
            arrival_zone4_end = BreachTime_in_minutes + arrival_time_break4
            arrival_zone5_start = BreachTime_in_minutes + arrival_time_break4
            arrival_zone5_end = BreachTime_in_minutes + arrival_time_break5
            arrival_zone6_start = BreachTime_in_minutes + arrival_time_break5
            arrival_zone6_end = BreachTime_in_minutes + arrival_time_break6
            arrival_zone7_start = BreachTime_in_minutes + arrival_time_break6
            arrival_zone7_end = BreachTime_in_minutes + 1000000 #arbitrary large number to represent get to end of time window
            arrival_zone1 = "PreHazard"
            arrival_zone2 = "0 to {:,.1f} hrs".format(arrival_time_break2/60) # 0 to .5
            arrival_zone3 = "{:,.1f} to {:,.1f} hrs".format(arrival_time_break2/60, arrival_time_break3/60) # 0.5 to 2
            arrival_zone4 = "{:,.0f} to {:,.0f} hrs".format(arrival_time_break3/60, arrival_time_break4/60) # 2 to 4
            arrival_zone5 = "{:,.0f} to {:,.0f} hrs".format(arrival_time_break4/60, arrival_time_break5/60) # 4 to 8
            arrival_zone6 = "{:,.0f} to {:,.0f} hrs".format(arrival_time_break5/60, arrival_time_break6/60) # 8 to 24
            arrival_zone7 = "Over {:,.0f} hrs".format(arrival_time_break6/60) # over 24
            arrival_zone8 = "Evaced, Not Wet"
            
            if epzfilter:
                add_message("Arrival time table is filtered to an EPZ using: {}".format(epzfilter), lst=2)

            #arrival_start_list = [0, arrival_zone2_start, arrival_zone3_start, arrival_zone4_start, arrival_zone5_start, arrival_zone6_start, arrival_zone7_start]
            #arrival_end_list = [BreachTime_in_minutes, arrival_zone2_end, arrival_zone3_end, arrival_zone4_end, arrival_zone5_end, arrival_zone6_end, arrival_zone7_end]


            # structure count by arrival zone
            struc_count_day_arzone1_query = f"SELECT COUNT(*) FROM {results1_day_input} WHERE Max_Depth > 0 AND {arrival_column} < ?{epzfilter}"
            cursor.execute(struc_count_day_arzone1_query, (BreachTime_in_minutes,))
            struc_count_day_arzone1_sum = cursor.fetchone()[0]

            struc_count_day_arzone2_query = f"SELECT COUNT(*) FROM {results1_day_input} WHERE Max_Depth > 0 AND {arrival_column} > ? AND {arrival_column} < ?{epzfilter}"
            cursor.execute(struc_count_day_arzone2_query, (arrival_zone2_start, arrival_zone2_end))
            struc_count_day_arzone2_sum = cursor.fetchone()[0]

            struc_count_day_arzone3_query = f"SELECT COUNT(*) FROM {results1_day_input} WHERE Max_Depth > 0 AND {arrival_column} > ? AND {arrival_column} < ?{epzfilter}"
            cursor.execute(struc_count_day_arzone3_query, (arrival_zone3_start, arrival_zone3_end))
            struc_count_day_arzone3_sum = cursor.fetchone()[0]

            struc_count_day_arzone4_query = f"SELECT COUNT(*) FROM {results1_day_input} WHERE Max_Depth > 0 AND {arrival_column} > ? AND {arrival_column} < ?{epzfilter}"
            cursor.execute(struc_count_day_arzone4_query, (arrival_zone4_start, arrival_zone4_end))
            struc_count_day_arzone4_sum = cursor.fetchone()[0]

            struc_count_day_arzone5_query = f"SELECT COUNT(*) FROM {results1_day_input} WHERE Max_Depth > 0 AND {arrival_column} > ? AND {arrival_column} < ?{epzfilter}"
            cursor.execute(struc_count_day_arzone5_query, (arrival_zone5_start, arrival_zone5_end))
            struc_count_day_arzone5_sum = cursor.fetchone()[0]

            struc_count_day_arzone6_query = f"SELECT COUNT(*) FROM {results1_day_input} WHERE Max_Depth > 0 AND {arrival_column} > ? AND {arrival_column} < ?{epzfilter}"
            cursor.execute(struc_count_day_arzone6_query, (arrival_zone6_start, arrival_zone6_end))
            struc_count_day_arzone6_sum = cursor.fetchone()[0]

            struc_count_day_arzone7_query = f"SELECT COUNT(*) FROM {results1_day_input} WHERE Max_Depth > 0 AND {arrival_column} > ?{epzfilter}"
            cursor.execute(struc_count_day_arzone7_query, (arrival_zone7_start,))
            struc_count_day_arzone7_sum = cursor.fetchone()[0]

            struc_count_day_arzone8_query = f"SELECT COUNT(*) FROM {results1_day_input} WHERE Max_Depth = 0 AND Life_Loss_Evacuating_Mean > 0"
            cursor.execute(struc_count_day_arzone8_query, )
            struc_count_day_arzone8_sum = cursor.fetchone()[0]

            # PAR DAY by arrival zone
            par_day_arzone1_query = f"SELECT SUM(Pop_Under65_Mean + PAR_Over65_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND {arrival_column} < ?{epzfilter}"
            cursor.execute(par_day_arzone1_query, (BreachTime_in_minutes,))
            par_day_arzone1_sum = cursor.fetchone()[0] or 0

            par_day_arzone2_query = f"SELECT SUM(Pop_Under65_Mean + PAR_Over65_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND {arrival_column} > ? AND {arrival_column} < ?{epzfilter}"
            cursor.execute(par_day_arzone2_query, (BreachTime_in_minutes, (BreachTime_in_minutes + arrival_time_break2)))
            par_day_arzone2_sum = cursor.fetchone()[0] or 0

            par_day_arzone3_query = f"SELECT SUM(Pop_Under65_Mean + PAR_Over65_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND {arrival_column} > ? AND {arrival_column} < ?{epzfilter}"
            cursor.execute(par_day_arzone3_query, ((BreachTime_in_minutes + arrival_time_break2), (BreachTime_in_minutes + arrival_time_break3)))
            par_day_arzone3_sum = cursor.fetchone()[0] or 0

            par_day_arzone4_query = f"SELECT SUM(Pop_Under65_Mean + PAR_Over65_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND {arrival_column} > ? AND {arrival_column} < ?{epzfilter}"
            cursor.execute(par_day_arzone4_query, ((BreachTime_in_minutes + arrival_time_break3), (BreachTime_in_minutes + arrival_time_break4)))
            par_day_arzone4_sum = cursor.fetchone()[0] or 0

            par_day_arzone5_query = f"SELECT SUM(Pop_Under65_Mean + PAR_Over65_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND {arrival_column} > ? AND {arrival_column} < ?{epzfilter}"
            cursor.execute(par_day_arzone5_query, ((BreachTime_in_minutes + arrival_time_break4), (BreachTime_in_minutes + arrival_time_break5)))
            par_day_arzone5_sum = cursor.fetchone()[0] or 0

            par_day_arzone6_query = f"SELECT SUM(Pop_Under65_Mean + PAR_Over65_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND {arrival_column} > ? AND {arrival_column} < ?{epzfilter}"
            cursor.execute(par_day_arzone6_query, ((BreachTime_in_minutes + arrival_time_break5), (BreachTime_in_minutes + arrival_time_break6)))
            par_day_arzone6_sum = cursor.fetchone()[0] or 0

            par_day_arzone7_query = f"SELECT SUM(Pop_Under65_Mean + PAR_Over65_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND {arrival_column} > ?{epzfilter}"
            cursor.execute(par_day_arzone7_query, ((BreachTime_in_minutes + arrival_time_break6),))
            par_day_arzone7_sum = cursor.fetchone()[0] or 0

            par_day_arzone8_query = f"SELECT SUM(Pop_Under65_Mean + PAR_Over65_Mean) FROM {results1_day_input} WHERE Max_Depth = 0 AND Life_Loss_Evacuating_Mean > 0"
            cursor.execute(par_day_arzone8_query, )
            par_day_arzone8_sum = cursor.fetchone()[0] or 0


            # PAR Night by arrival zone
            par_night_arzone1_query = f"SELECT SUM(Pop_Under65_Mean + PAR_Over65_Mean) FROM {results1_night_input} WHERE Max_Depth > 0 AND {arrival_column} < ?{epzfilter}"
            cursor.execute(par_night_arzone1_query, (BreachTime_in_minutes,))
            par_night_arzone1_sum = cursor.fetchone()[0] or 0

            par_night_arzone2_query = f"SELECT SUM(Pop_Under65_Mean + PAR_Over65_Mean) FROM {results1_night_input} WHERE Max_Depth > 0 AND {arrival_column} > ? AND {arrival_column} < ?{epzfilter}"
            cursor.execute(par_night_arzone2_query, (BreachTime_in_minutes, (BreachTime_in_minutes + arrival_time_break2)))
            par_night_arzone2_sum = cursor.fetchone()[0] or 0

            par_night_arzone3_query = f"SELECT SUM(Pop_Under65_Mean + PAR_Over65_Mean) FROM {results1_night_input} WHERE Max_Depth > 0 AND {arrival_column} > ? AND {arrival_column} < ?{epzfilter}"
            cursor.execute(par_night_arzone3_query, ((BreachTime_in_minutes + arrival_time_break2), (BreachTime_in_minutes + arrival_time_break3)))
            par_night_arzone3_sum = cursor.fetchone()[0] or 0

            par_night_arzone4_query = f"SELECT SUM(Pop_Under65_Mean + PAR_Over65_Mean) FROM {results1_night_input} WHERE Max_Depth > 0 AND {arrival_column} > ? AND {arrival_column} < ?{epzfilter}"
            cursor.execute(par_night_arzone4_query, ((BreachTime_in_minutes + arrival_time_break3), (BreachTime_in_minutes + arrival_time_break4)))
            par_night_arzone4_sum = cursor.fetchone()[0] or 0

            par_night_arzone5_query = f"SELECT SUM(Pop_Under65_Mean + PAR_Over65_Mean) FROM {results1_night_input} WHERE Max_Depth > 0 AND {arrival_column} > ? AND {arrival_column} < ?{epzfilter}"
            cursor.execute(par_night_arzone5_query, ((BreachTime_in_minutes + arrival_time_break4), (BreachTime_in_minutes + arrival_time_break5)))
            par_night_arzone5_sum = cursor.fetchone()[0] or 0

            par_night_arzone6_query = f"SELECT SUM(Pop_Under65_Mean + PAR_Over65_Mean) FROM {results1_night_input} WHERE Max_Depth > 0 AND {arrival_column} > ? AND {arrival_column} < ?{epzfilter}"
            cursor.execute(par_night_arzone6_query, ((BreachTime_in_minutes + arrival_time_break5), (BreachTime_in_minutes + arrival_time_break6)))
            par_night_arzone6_sum = cursor.fetchone()[0] or 0

            par_night_arzone7_query = f"SELECT SUM(Pop_Under65_Mean + PAR_Over65_Mean) FROM {results1_night_input} WHERE Max_Depth > 0 AND {arrival_column} > ?{epzfilter}"
            cursor.execute(par_night_arzone7_query, ((BreachTime_in_minutes + arrival_time_break6),))
            par_night_arzone7_sum = cursor.fetchone()[0] or 0

            par_night_arzone8_query = f"SELECT SUM(Pop_Under65_Mean + PAR_Over65_Mean) FROM {results1_night_input} WHERE Max_Depth = 0 AND Life_Loss_Evacuating_Mean > 0"
            cursor.execute(par_night_arzone8_query, )
            par_night_arzone8_sum = cursor.fetchone()[0] or 0


            # Life loss day by arrival zone
            total_lifeloss_day_arzone1_query = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND {arrival_column} < ?{epzfilter}"
            cursor.execute(total_lifeloss_day_arzone1_query, (BreachTime_in_minutes,))
            total_lifeloss_day_arzone1_sum = cursor.fetchone()[0]

            total_lifeloss_day_arzone2_query = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND {arrival_column} > ? AND {arrival_column} < ?{epzfilter}"
            cursor.execute(total_lifeloss_day_arzone2_query, (BreachTime_in_minutes, (BreachTime_in_minutes + arrival_time_break2)))
            total_lifeloss_day_arzone2_sum = cursor.fetchone()[0]

            total_lifeloss_day_arzone3_query = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND {arrival_column} > ? AND {arrival_column} < ?{epzfilter}"
            cursor.execute(total_lifeloss_day_arzone3_query, ((BreachTime_in_minutes + arrival_time_break2), (BreachTime_in_minutes + arrival_time_break3)))
            total_lifeloss_day_arzone3_sum = cursor.fetchone()[0]

            total_lifeloss_day_arzone4_query = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND {arrival_column} > ? AND {arrival_column} < ?{epzfilter}"
            cursor.execute(total_lifeloss_day_arzone4_query, ((BreachTime_in_minutes + arrival_time_break3), (BreachTime_in_minutes + arrival_time_break4)))
            total_lifeloss_day_arzone4_sum = cursor.fetchone()[0]

            total_lifeloss_day_arzone5_query = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND {arrival_column} > ? AND {arrival_column} < ?{epzfilter}"
            cursor.execute(total_lifeloss_day_arzone5_query, ((BreachTime_in_minutes + arrival_time_break4), (BreachTime_in_minutes + arrival_time_break5)))
            total_lifeloss_day_arzone5_sum = cursor.fetchone()[0]

            total_lifeloss_day_arzone6_query = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND {arrival_column} > ? AND {arrival_column} < ?{epzfilter}"
            cursor.execute(total_lifeloss_day_arzone6_query, ((BreachTime_in_minutes + arrival_time_break5), (BreachTime_in_minutes + arrival_time_break6)))
            total_lifeloss_day_arzone6_sum = cursor.fetchone()[0]

            total_lifeloss_day_arzone7_query = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND {arrival_column} > ?{epzfilter}"
            cursor.execute(total_lifeloss_day_arzone7_query, ((BreachTime_in_minutes + arrival_time_break6),))
            total_lifeloss_day_arzone7_sum = cursor.fetchone()[0]

            total_lifeloss_day_arzone8_query = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_day_input} WHERE Max_Depth = 0 AND Life_Loss_Evacuating_Mean > 0"
            cursor.execute(total_lifeloss_day_arzone8_query, )
            total_lifeloss_day_arzone8_sum = cursor.fetchone()[0]

            # Life loss night by arrival zone
            total_lifeloss_night_arzone1_query = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_night_input} WHERE Max_Depth > 0 AND {arrival_column} < ?{epzfilter}"
            cursor.execute(total_lifeloss_night_arzone1_query, (BreachTime_in_minutes,))
            total_lifeloss_night_arzone1_sum = cursor.fetchone()[0]

            total_lifeloss_night_arzone2_query = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_night_input} WHERE Max_Depth > 0 AND {arrival_column} > ? AND {arrival_column} < ?{epzfilter}"
            cursor.execute(total_lifeloss_night_arzone2_query, (BreachTime_in_minutes, (BreachTime_in_minutes + arrival_time_break2)))
            total_lifeloss_night_arzone2_sum = cursor.fetchone()[0]

            total_lifeloss_night_arzone3_query = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_night_input} WHERE Max_Depth > 0 AND {arrival_column} > ? AND {arrival_column} < ?{epzfilter}"
            cursor.execute(total_lifeloss_night_arzone3_query, ((BreachTime_in_minutes + arrival_time_break2), (BreachTime_in_minutes + arrival_time_break3)))
            total_lifeloss_night_arzone3_sum = cursor.fetchone()[0]

            total_lifeloss_night_arzone4_query = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_night_input} WHERE Max_Depth > 0 AND {arrival_column} > ? AND {arrival_column} < ?{epzfilter}"
            cursor.execute(total_lifeloss_night_arzone4_query, ((BreachTime_in_minutes + arrival_time_break3), (BreachTime_in_minutes + arrival_time_break4)))
            total_lifeloss_night_arzone4_sum = cursor.fetchone()[0]

            total_lifeloss_night_arzone5_query = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_night_input} WHERE Max_Depth > 0 AND {arrival_column} > ? AND {arrival_column} < ?{epzfilter}"
            cursor.execute(total_lifeloss_night_arzone5_query, ((BreachTime_in_minutes + arrival_time_break4), (BreachTime_in_minutes + arrival_time_break5)))
            total_lifeloss_night_arzone5_sum = cursor.fetchone()[0]

            total_lifeloss_night_arzone6_query = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_night_input} WHERE Max_Depth > 0 AND {arrival_column} > ? AND {arrival_column} < ?{epzfilter}"
            cursor.execute(total_lifeloss_night_arzone6_query, ((BreachTime_in_minutes + arrival_time_break5), (BreachTime_in_minutes + arrival_time_break6)))
            total_lifeloss_night_arzone6_sum = cursor.fetchone()[0]

            total_lifeloss_night_arzone7_query = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_night_input} WHERE Max_Depth > 0 AND {arrival_column} > ?{epzfilter}"
            cursor.execute(total_lifeloss_night_arzone7_query, ((BreachTime_in_minutes + arrival_time_break6),))
            total_lifeloss_night_arzone7_sum = cursor.fetchone()[0]

            total_lifeloss_night_arzone8_query = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_night_input} WHERE Max_Depth = 0 AND Life_Loss_Evacuating_Mean > 0"
            cursor.execute(total_lifeloss_night_arzone8_query, )
            total_lifeloss_night_arzone8_sum = cursor.fetchone()[0]


            # SQL query to calculate the low and high percentile depths
                    #the count times 0.15 identifies the 15th percentile row, change for other percentiles
            arzone_percentile_depth_query = f"""
                    SELECT Max_Depth FROM {results1_day_input} 
                    WHERE {arrival_column} > ? AND {arrival_column} < ? 
                    AND Max_Depth > 0{epzfilter} 
                    ORDER BY Max_Depth
                    LIMIT 1 
                    OFFSET CAST((SELECT COUNT(*) FROM {results1_day_input} 
                    WHERE {arrival_column} > ? AND {arrival_column} < ?  
                    AND Max_Depth > 0{epzfilter}) * ? AS INTEGER)
                    """
                    
            #zone 1 prebreach
            cursor.execute(arzone_percentile_depth_query, (0, BreachTime_in_minutes, 0, BreachTime_in_minutes, depthlowpercentile))
            arzone1_depthlowpercent = cursor.fetchone()
            arzone1_depthlowpercent = round(arzone1_depthlowpercent[0], 1) if arzone1_depthlowpercent else 0
            cursor.execute(arzone_percentile_depth_query, (0, BreachTime_in_minutes, 0, BreachTime_in_minutes, depthhighpercentile))
            arzone1_depthhighpercent = cursor.fetchone()
            arzone1_depthhighpercent = round(arzone1_depthhighpercent[0], 1) if arzone1_depthhighpercent else 0
            arzone1_depth_range = f"{arzone1_depthlowpercent} - {arzone1_depthhighpercent}"
            #add_message("Depth Percentile Range arzone1: {}".format(arzone1_depth_range), lst=2)

            #zone 2 0-.5 hours
            cursor.execute(arzone_percentile_depth_query, (arrival_zone2_start, arrival_zone2_end, arrival_zone2_start, arrival_zone2_end, depthlowpercentile))
            arzone2_depthlowpercent = cursor.fetchone()
            arzone2_depthlowpercent = round(arzone2_depthlowpercent[0], 1) if arzone2_depthlowpercent else 0
            cursor.execute(arzone_percentile_depth_query, (arrival_zone2_start, arrival_zone2_end, arrival_zone2_start, arrival_zone2_end, depthhighpercentile))
            arzone2_depthhighpercent = cursor.fetchone()
            arzone2_depthhighpercent = round(arzone2_depthhighpercent[0], 1) if arzone2_depthhighpercent else 0
            arzone2_depth_range = f"{arzone2_depthlowpercent} - {arzone2_depthhighpercent}"
            #add_message("Depth Percentile Range arzone2: {}".format(arzone2_depth_range), lst=2)

            #zone 3 .5-2 hours
            cursor.execute(arzone_percentile_depth_query, (arrival_zone3_start, arrival_zone3_end, arrival_zone3_start, arrival_zone3_end, depthlowpercentile))
            arzone3_depthlowpercent = cursor.fetchone()
            arzone3_depthlowpercent = round(arzone3_depthlowpercent[0], 1) if arzone3_depthlowpercent else 0
            cursor.execute(arzone_percentile_depth_query, (arrival_zone3_start, arrival_zone3_end, arrival_zone3_start, arrival_zone3_end, depthhighpercentile))
            arzone3_depthhighpercent = cursor.fetchone()
            arzone3_depthhighpercent = round(arzone3_depthhighpercent[0], 1) if arzone3_depthhighpercent else 0
            arzone3_depth_range = f"{arzone3_depthlowpercent} - {arzone3_depthhighpercent}"
            #add_message("Depth Percentile Range arzone3: {}".format(arzone3_depth_range), lst=2)

            #zone 4 2-4 hours
            cursor.execute(arzone_percentile_depth_query, (arrival_zone4_start, arrival_zone4_end, arrival_zone4_start, arrival_zone4_end, depthlowpercentile))
            arzone4_depthlowpercent = cursor.fetchone()
            arzone4_depthlowpercent = round(arzone4_depthlowpercent[0], 1) if arzone4_depthlowpercent else 0
            cursor.execute(arzone_percentile_depth_query, (arrival_zone4_start, arrival_zone4_end, arrival_zone4_start, arrival_zone4_end, depthhighpercentile))
            arzone4_depthhighpercent = cursor.fetchone()
            arzone4_depthhighpercent = round(arzone4_depthhighpercent[0], 1) if arzone4_depthhighpercent else 0
            arzone4_depth_range = f"{arzone4_depthlowpercent} - {arzone4_depthhighpercent}"
            #add_message("Depth Percentile Range arzone4: {}".format(arzone4_depth_range), lst=2)

            #zone 5 4-8 hours
            cursor.execute(arzone_percentile_depth_query, (arrival_zone5_start, arrival_zone5_end, arrival_zone5_start, arrival_zone5_end, depthlowpercentile))
            arzone5_depthlowpercent = cursor.fetchone()
            arzone5_depthlowpercent = round(arzone5_depthlowpercent[0], 1) if arzone5_depthlowpercent else 0
            cursor.execute(arzone_percentile_depth_query, (arrival_zone5_start, arrival_zone5_end, arrival_zone5_start, arrival_zone5_end, depthhighpercentile))
            arzone5_depthhighpercent = cursor.fetchone()
            arzone5_depthhighpercent = round(arzone5_depthhighpercent[0], 1) if arzone5_depthhighpercent else 0
            arzone5_depth_range = f"{arzone5_depthlowpercent} - {arzone5_depthhighpercent}"
            #add_message("Depth Percentile Range arzone5: {}".format(arzone5_depth_range), lst=2)

            #zone 6 8-24 hours
            cursor.execute(arzone_percentile_depth_query, (arrival_zone6_start, arrival_zone6_end, arrival_zone6_start, arrival_zone6_end, depthlowpercentile))
            arzone6_depthlowpercent = cursor.fetchone()
            arzone6_depthlowpercent = round(arzone6_depthlowpercent[0], 1) if arzone6_depthlowpercent else 0
            cursor.execute(arzone_percentile_depth_query, (arrival_zone6_start, arrival_zone6_end, arrival_zone6_start, arrival_zone6_end, depthhighpercentile))
            arzone6_depthhighpercent = cursor.fetchone()
            arzone6_depthhighpercent = round(arzone6_depthhighpercent[0], 1) if arzone6_depthhighpercent else 0
            arzone6_depth_range = f"{arzone6_depthlowpercent} - {arzone6_depthhighpercent}"
            #add_message("Depth Percentile Range arzone6: {}".format(arzone6_depth_range), lst=2)

            #zone 7 over 24 hours
            cursor.execute(arzone_percentile_depth_query, (arrival_zone7_start, arrival_zone7_end, arrival_zone7_start, arrival_zone7_end, depthlowpercentile))
            arzone7_depthlowpercent = cursor.fetchone()
            arzone7_depthlowpercent = round(arzone7_depthlowpercent[0], 1) if arzone7_depthlowpercent else 0
            cursor.execute(arzone_percentile_depth_query, (arrival_zone7_start, arrival_zone7_end, arrival_zone7_start, arrival_zone7_end, depthhighpercentile))
            arzone7_depthhighpercent = cursor.fetchone()
            arzone7_depthhighpercent = round(arzone7_depthhighpercent[0], 1) if arzone7_depthhighpercent else 0
            arzone7_depth_range = f"{arzone7_depthlowpercent} - {arzone7_depthhighpercent}"
            #add_message("Depth Percentile Range arzone7: {}".format(arzone7_depth_range), lst=2)


            # SQL query to calculate the low and high percentile depths
                    #the count times 0.15 identifies the 15th percentile row, change for other percentiles
            arzone_percentile_velocity_query = f"""
                    SELECT Max_Velocity FROM {results1_day_input} 
                    WHERE {arrival_column} > ? AND {arrival_column} < ? 
                    AND Max_Depth > 0{epzfilter} 
                    ORDER BY Max_Velocity
                    LIMIT 1 
                    OFFSET CAST((SELECT COUNT(*) FROM {results1_day_input} 
                    WHERE {arrival_column} > ? AND {arrival_column} < ?  
                    AND Max_Depth > 0{epzfilter}) * ? AS INTEGER)
                    """
            
            #zone 1 prebreach
            cursor.execute(arzone_percentile_velocity_query, (0, BreachTime_in_minutes, 0, BreachTime_in_minutes, velocitylowpercentile))
            arzone1_velocitylowpercent = cursor.fetchone()
            arzone1_velocitylowpercent = round(arzone1_velocitylowpercent[0], 1) if arzone1_velocitylowpercent else 0
            cursor.execute(arzone_percentile_velocity_query, (0, BreachTime_in_minutes, 0, BreachTime_in_minutes, velocityhighpercentile))
            arzone1_velocityhighpercent = cursor.fetchone()
            arzone1_velocityhighpercent = round(arzone1_velocityhighpercent[0], 1) if arzone1_velocityhighpercent else 0
            arzone1_velocity_range = f"{arzone1_velocitylowpercent} - {arzone1_velocityhighpercent}"
            #add_message("Velocity Percentile Range arzone1: {}".format(arzone1_velocity_range), lst=2)
            #zone 2 0-.5 hours
            cursor.execute(arzone_percentile_velocity_query, (arrival_zone2_start, arrival_zone2_end, arrival_zone2_start, arrival_zone2_end, velocitylowpercentile))
            arzone2_velocitylowpercent = cursor.fetchone()
            arzone2_velocitylowpercent = round(arzone2_velocitylowpercent[0], 1) if arzone2_velocitylowpercent else 0
            cursor.execute(arzone_percentile_velocity_query, (arrival_zone2_start, arrival_zone2_end, arrival_zone2_start, arrival_zone2_end, velocityhighpercentile))
            arzone2_velocityhighpercent = cursor.fetchone()
            arzone2_velocityhighpercent = round(arzone2_velocityhighpercent[0], 1) if arzone2_velocityhighpercent else 0
            arzone2_velocity_range = f"{arzone2_velocitylowpercent} - {arzone2_velocityhighpercent}"
            #add_message("Velocity Percentile Range arzone2: {}".format(arzone2_velocity_range), lst=2)
            #zone 3 .5-2 hours
            cursor.execute(arzone_percentile_velocity_query, (arrival_zone3_start, arrival_zone3_end, arrival_zone3_start, arrival_zone3_end, velocitylowpercentile))
            arzone3_velocitylowpercent = cursor.fetchone()
            arzone3_velocitylowpercent = round(arzone3_velocitylowpercent[0], 1) if arzone3_velocitylowpercent else 0
            cursor.execute(arzone_percentile_velocity_query, (arrival_zone3_start, arrival_zone3_end, arrival_zone3_start, arrival_zone3_end, velocityhighpercentile))
            arzone3_velocityhighpercent = cursor.fetchone()
            arzone3_velocityhighpercent = round(arzone3_velocityhighpercent[0], 1) if arzone3_velocityhighpercent else 0
            arzone3_velocity_range = f"{arzone3_velocitylowpercent} - {arzone3_velocityhighpercent}"
            #add_message("Velocity Percentile Range arzone3: {}".format(arzone3_velocity_range), lst=2)
            #zone 4 2-4 hours
            cursor.execute(arzone_percentile_velocity_query, (arrival_zone4_start, arrival_zone4_end, arrival_zone4_start, arrival_zone4_end, velocitylowpercentile))
            arzone4_velocitylowpercent = cursor.fetchone()
            arzone4_velocitylowpercent = round(arzone4_velocitylowpercent[0], 1) if arzone4_velocitylowpercent else 0
            cursor.execute(arzone_percentile_velocity_query, (arrival_zone4_start, arrival_zone4_end, arrival_zone4_start, arrival_zone4_end, velocityhighpercentile))
            arzone4_velocityhighpercent = cursor.fetchone()
            arzone4_velocityhighpercent = round(arzone4_velocityhighpercent[0], 1) if arzone4_velocityhighpercent else 0
            arzone4_velocity_range = f"{arzone4_velocitylowpercent} - {arzone4_velocityhighpercent}"
            #add_message("Velocity Percentile Range arzone4: {}".format(arzone4_velocity_range), lst=2)
            #zone 5 4-8 hours
            cursor.execute(arzone_percentile_velocity_query, (arrival_zone5_start, arrival_zone5_end, arrival_zone5_start, arrival_zone5_end, velocitylowpercentile))
            arzone5_velocitylowpercent = cursor.fetchone()
            arzone5_velocitylowpercent = round(arzone5_velocitylowpercent[0], 1) if arzone5_velocitylowpercent else 0
            cursor.execute(arzone_percentile_velocity_query, (arrival_zone5_start, arrival_zone5_end, arrival_zone5_start, arrival_zone5_end, velocityhighpercentile))
            arzone5_velocityhighpercent = cursor.fetchone()
            arzone5_velocityhighpercent = round(arzone5_velocityhighpercent[0], 1) if arzone5_velocityhighpercent else 0
            arzone5_velocity_range = f"{arzone5_velocitylowpercent} - {arzone5_velocityhighpercent}"
            #add_message("Velocity Percentile Range arzone5: {}".format(arzone5_velocity_range), lst=2)
            #zone 6 8-24 hours
            cursor.execute(arzone_percentile_velocity_query, (arrival_zone6_start, arrival_zone6_end, arrival_zone6_start, arrival_zone6_end, velocitylowpercentile))
            arzone6_velocitylowpercent = cursor.fetchone()
            arzone6_velocitylowpercent = round(arzone6_velocitylowpercent[0], 1) if arzone6_velocitylowpercent else 0
            cursor.execute(arzone_percentile_velocity_query, (arrival_zone6_start, arrival_zone6_end, arrival_zone6_start, arrival_zone6_end, velocityhighpercentile))
            arzone6_velocityhighpercent = cursor.fetchone()
            arzone6_velocityhighpercent = round(arzone6_velocityhighpercent[0], 1) if arzone6_velocityhighpercent else 0
            arzone6_velocity_range = f"{arzone6_velocitylowpercent} - {arzone6_velocityhighpercent}"
            #add_message("Velocity Percentile Range arzone6: {}".format(arzone6_velocity_range), lst=2)
            #zone 7 over 24 hours
            cursor.execute(arzone_percentile_velocity_query, (arrival_zone7_start, arrival_zone7_end, arrival_zone7_start, arrival_zone7_end, velocitylowpercentile))
            arzone7_velocitylowpercent = cursor.fetchone()
            arzone7_velocitylowpercent = round(arzone7_velocitylowpercent[0], 1) if arzone7_velocitylowpercent else 0
            cursor.execute(arzone_percentile_velocity_query, (arrival_zone7_start, arrival_zone7_end, arrival_zone7_start, arrival_zone7_end, velocityhighpercentile))
            arzone7_velocityhighpercent = cursor.fetchone()
            arzone7_velocityhighpercent = round(arzone7_velocityhighpercent[0], 1) if arzone7_velocityhighpercent else 0
            arzone7_velocity_range = f"{arzone7_velocitylowpercent} - {arzone7_velocityhighpercent}"
            #add_message("Velocity Percentile Range arzone7: {}".format(arzone7_velocity_range), lst=2)

            #SQL query to calculate PAR_Warned_Mean day in each arrival zone
            #zone 1 prebreach
            par_warned_day_arzone1_query = f"SELECT SUM(PAR_Warned_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND {arrival_column} < ?{epzfilter}"
            cursor.execute(par_warned_day_arzone1_query, (BreachTime_in_minutes,))
            par_warned_day_arzone1_sum = cursor.fetchone()[0] or 0
            par_warned_day_arzone1_percentage = (par_warned_day_arzone1_sum / par_day_arzone1_sum) if par_day_arzone1_sum > 0 else 0

            #zone 2 0-.5 hours
            par_warned_day_arzone2_query = f"SELECT SUM(PAR_Warned_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND {arrival_column} > ? AND {arrival_column} < ?{epzfilter}"
            cursor.execute(par_warned_day_arzone2_query, (arrival_zone2_start, arrival_zone2_end))
            par_warned_day_arzone2_sum = cursor.fetchone()[0] or 0
            par_warned_day_arzone2_percentage = (par_warned_day_arzone2_sum / par_day_arzone2_sum) if par_day_arzone2_sum > 0 else 0

            #zone 3 .5-2 hours
            par_warned_day_arzone3_query = f"SELECT SUM(PAR_Warned_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND {arrival_column} > ? AND {arrival_column} < ?{epzfilter}"
            cursor.execute(par_warned_day_arzone3_query, (arrival_zone3_start, arrival_zone3_end))
            par_warned_day_arzone3_sum = cursor.fetchone()[0] or 0
            par_warned_day_arzone3_percentage = (par_warned_day_arzone3_sum / par_day_arzone3_sum) if par_day_arzone3_sum > 0 else 0

            #zone 4 2-4 hours
            par_warned_day_arzone4_query = f"SELECT SUM(PAR_Warned_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND {arrival_column} > ? AND {arrival_column} < ?{epzfilter}"
            cursor.execute(par_warned_day_arzone4_query, (arrival_zone4_start, arrival_zone4_end))
            par_warned_day_arzone4_sum = cursor.fetchone()[0] or 0
            par_warned_day_arzone4_percentage = (par_warned_day_arzone4_sum / par_day_arzone4_sum) if par_day_arzone4_sum > 0 else 0

            #zone 5 4-8 hours
            par_warned_day_arzone5_query = f"SELECT SUM(PAR_Warned_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND {arrival_column} > ? AND {arrival_column} < ?{epzfilter}"
            cursor.execute(par_warned_day_arzone5_query, (arrival_zone5_start, arrival_zone5_end))
            par_warned_day_arzone5_sum = cursor.fetchone()[0] or 0
            par_warned_day_arzone5_percentage = (par_warned_day_arzone5_sum / par_day_arzone5_sum) if par_day_arzone5_sum > 0 else 0

            #zone 6 8-24 hours
            par_warned_day_arzone6_query = f"SELECT SUM(PAR_Warned_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND {arrival_column} > ? AND {arrival_column} < ?{epzfilter}"
            cursor.execute(par_warned_day_arzone6_query, (arrival_zone6_start, arrival_zone6_end))
            par_warned_day_arzone6_sum = cursor.fetchone()[0] or 0
            par_warned_day_arzone6_percentage = (par_warned_day_arzone6_sum / par_day_arzone6_sum) if par_day_arzone6_sum > 0 else 0

            #zone 7 over 24 hours
            par_warned_day_arzone7_query = f"SELECT SUM(PAR_Warned_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND {arrival_column} > ?{epzfilter}"
            cursor.execute(par_warned_day_arzone7_query, (arrival_zone7_start,))
            par_warned_day_arzone7_sum = cursor.fetchone()[0] or 0
            par_warned_day_arzone7_percentage = (par_warned_day_arzone7_sum / par_day_arzone7_sum) if par_day_arzone7_sum > 0 else 0

            #SQL query to calculate PAR_Mobilized_Mean day in each arrival zone
            #zone 1 prebreach
            par_mobilized_day_arzone1_query = f"SELECT SUM(PAR_Mobilized_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND {arrival_column} < ?{epzfilter}"
            cursor.execute(par_mobilized_day_arzone1_query, (BreachTime_in_minutes,))
            par_mobilized_day_arzone1_sum = cursor.fetchone()[0] or 0
            par_mobilized_day_arzone1_percentage = (par_mobilized_day_arzone1_sum / par_day_arzone1_sum) if par_day_arzone1_sum > 0 else 0

            #zone 2 0-.5 hours
            par_mobilized_day_arzone2_query = f"SELECT SUM(PAR_Mobilized_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND {arrival_column} > ? AND {arrival_column} < ?{epzfilter}"
            cursor.execute(par_mobilized_day_arzone2_query, (arrival_zone2_start, arrival_zone2_end))
            par_mobilized_day_arzone2_sum = cursor.fetchone()[0] or 0
            par_mobilized_day_arzone2_percentage = (par_mobilized_day_arzone2_sum / par_day_arzone2_sum) if par_day_arzone2_sum > 0 else 0

            #zone 3 .5-2 hours
            par_mobilized_day_arzone3_query = f"SELECT SUM(PAR_Mobilized_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND {arrival_column} > ? AND {arrival_column} < ?{epzfilter}"
            cursor.execute(par_mobilized_day_arzone3_query, (arrival_zone3_start, arrival_zone3_end))
            par_mobilized_day_arzone3_sum = cursor.fetchone()[0] or 0
            par_mobilized_day_arzone3_percentage = (par_mobilized_day_arzone3_sum / par_day_arzone3_sum) if par_day_arzone3_sum > 0 else 0

            #zone 4 2-4 hours
            par_mobilized_day_arzone4_query = f"SELECT SUM(PAR_Mobilized_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND {arrival_column} > ? AND {arrival_column} < ?{epzfilter}"
            cursor.execute(par_mobilized_day_arzone4_query, (arrival_zone4_start, arrival_zone4_end))
            par_mobilized_day_arzone4_sum = cursor.fetchone()[0] or 0
            par_mobilized_day_arzone4_percentage = (par_mobilized_day_arzone4_sum / par_day_arzone4_sum) if par_day_arzone4_sum > 0 else 0

            #zone 5 4-8 hours
            par_mobilized_day_arzone5_query = f"SELECT SUM(PAR_Mobilized_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND {arrival_column} > ? AND {arrival_column} < ?{epzfilter}"
            cursor.execute(par_mobilized_day_arzone5_query, (arrival_zone5_start, arrival_zone5_end))
            par_mobilized_day_arzone5_sum = cursor.fetchone()[0] or 0
            par_mobilized_day_arzone5_percentage = (par_mobilized_day_arzone5_sum / par_day_arzone5_sum) if par_day_arzone5_sum > 0 else 0

            #zone 6 8-24 hours
            par_mobilized_day_arzone6_query = f"SELECT SUM(PAR_Mobilized_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND {arrival_column} > ? AND {arrival_column} < ?{epzfilter}"
            cursor.execute(par_mobilized_day_arzone6_query, (arrival_zone6_start, arrival_zone6_end))
            par_mobilized_day_arzone6_sum = cursor.fetchone()[0] or 0
            par_mobilized_day_arzone6_percentage = (par_mobilized_day_arzone6_sum / par_day_arzone6_sum) if par_day_arzone6_sum > 0 else 0

            #zone 7 over 24 hours
            par_mobilized_day_arzone7_query = f"SELECT SUM(PAR_Mobilized_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND {arrival_column} > ?{epzfilter}"
            cursor.execute(par_mobilized_day_arzone7_query, (arrival_zone7_start,))
            par_mobilized_day_arzone7_sum = cursor.fetchone()[0] or 0
            par_mobilized_day_arzone7_percentage = (par_mobilized_day_arzone7_sum / par_day_arzone7_sum) if par_day_arzone7_sum > 0 else 0

            #calculate day fatality rate for arrival zones
            fatality_arzone1 = (total_lifeloss_day_arzone1_sum / par_day_arzone1_sum) if par_day_arzone1_sum > 0 else 0
            fatality_arzone2 = (total_lifeloss_day_arzone2_sum / par_day_arzone2_sum) if par_day_arzone2_sum > 0 else 0
            fatality_arzone3 = (total_lifeloss_day_arzone3_sum / par_day_arzone3_sum) if par_day_arzone3_sum > 0 else 0
            fatality_arzone4 = (total_lifeloss_day_arzone4_sum / par_day_arzone4_sum) if par_day_arzone4_sum > 0 else 0
            fatality_arzone5 = (total_lifeloss_day_arzone5_sum / par_day_arzone5_sum) if par_day_arzone5_sum > 0 else 0
            fatality_arzone6 = (total_lifeloss_day_arzone6_sum / par_day_arzone6_sum) if par_day_arzone6_sum > 0 else 0
            fatality_arzone7 = (total_lifeloss_day_arzone7_sum / par_day_arzone7_sum) if par_day_arzone7_sum > 0 else 0


            # SQL query to count features with depth greater than zero and arrival time pre and post hazard the numeric field above the threshold
            posthazardarrival_day_query = f"SELECT COUNT(*) FROM {results1_day_input} WHERE {arrival_column} > ? AND Max_Depth > 0"
            cursor.execute(posthazardarrival_day_query, (BreachTime_in_minutes,))
            posthazardarrival_day_count = cursor.fetchone()[0]
            add_message("Structures with post hazard arrival (day): {}".format(posthazardarrival_day_count), lst=2)

            prehazardarrival_day_query = f"SELECT COUNT(*) FROM {results1_day_input} WHERE {arrival_column} < ? AND Max_Depth > 0"
            cursor.execute(prehazardarrival_day_query, (BreachTime_in_minutes,))
            prehazardarrival_day_count = cursor.fetchone()[0]
            add_message("Structures with pre hazard arrival (day): {}".format(prehazardarrival_day_count), lst=2)

            # SQL query to sum mean life loss based on pre or post hazard arrival time
            posthazardarrival_LLday_query = f"SELECT SUM(Life_Loss_Total_Mean) FROM {results1_day_input} WHERE {arrival_column} > ? AND Max_Depth > 0"
            cursor.execute(posthazardarrival_LLday_query, (BreachTime_in_minutes,))
            posthazardarrival_LLday_sum = cursor.fetchone()[0] or 0
            add_message("LL in Structures with post hazard arrival (day): {}".format(round(posthazardarrival_LLday_sum,2)), lst=2)

            prehazardarrival_LLday_query = f"SELECT SUM(Life_Loss_Total_Mean) FROM {results1_day_input} WHERE {arrival_column} < ? AND Max_Depth > 0"
            cursor.execute(prehazardarrival_LLday_query, (BreachTime_in_minutes,))
            prehazardarrival_LLday_sum = cursor.fetchone()[0] or 0
            add_message("LL in Structures with pre hazard arrival (day): {}".format(round(prehazardarrival_LLday_sum,2)), lst=2)

            ##Structure Inventory Checks
            #count flooded buildings over 10 stories
            query_tallbuildings1 = f"SELECT COUNT(*) FROM {results1_day_input} WHERE Max_Depth > 0 AND Number_of_Stories > 10"
            cursor.execute(query_tallbuildings1,)
            count_tallbuildings1 = cursor.fetchone()[0]
            add_message("Structures over 10 stories flooded: {}".format(count_tallbuildings1), lst=2) 

            #count flooded buildings over 30 stories
            query_tallbuildings2 = f"SELECT COUNT(*) FROM {results1_day_input} WHERE Max_Depth > 0 AND Number_of_Stories > 30"
            cursor.execute(query_tallbuildings2,)
            count_tallbuildings2 = cursor.fetchone()[0]
            add_message("Structures over 30 stories flooded: {}".format(count_tallbuildings2), lst=2) 

            ##Life loss by occupancy type
            add_message("LL mean total (day): {:,.1f} ...(night): {:,.1f}".format(total_lifeloss_day, total_lifeloss_night), lst=2)
            # sum day life loss in schools
            query_total_LLday_inschools = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND Occupancy_Type LIKE 'EDU%'"
            cursor.execute(query_total_LLday_inschools,)
            total_LLday_inschools = cursor.fetchone()[0] or 0
            query_total_LLnight_inschools = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_night_input} WHERE Max_Depth > 0 AND Occupancy_Type LIKE 'EDU%'"
            cursor.execute(query_total_LLnight_inschools,)
            total_LLnight_inschools = cursor.fetchone()[0] or 0
            add_message("LL mean in EDU (day): {:,.1f} ...(night): {:,.1f}".format(total_LLday_inschools, total_LLnight_inschools), lst=2)

            # sum day life loss in RES1
            query_total_LLday_res1 = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND Occupancy_Type LIKE 'RES1%'"
            cursor.execute(query_total_LLday_res1,)
            total_LLday_res1 = cursor.fetchone()[0] or 0
            query_total_LLnight_res1 = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_night_input} WHERE Max_Depth > 0 AND Occupancy_Type LIKE 'RES1%'"
            cursor.execute(query_total_LLnight_res1,)
            total_LLnight_res1 = cursor.fetchone()[0] or 0
            add_message("LL mean in RES1 (day): {:,.1f} ...(night): {:,.1f}".format(total_LLday_res1, total_LLnight_res1), lst=2)

            # sum day life loss in RES2
            query_total_LLday_res2 = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND Occupancy_Type LIKE 'RES2%'"
            cursor.execute(query_total_LLday_res2,)
            total_LLday_res2 = cursor.fetchone()[0] or 0
            query_total_LLnight_res2 = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_night_input} WHERE Max_Depth > 0 AND Occupancy_Type LIKE 'RES2%'"
            cursor.execute(query_total_LLnight_res2,)
            total_LLnight_res2 = cursor.fetchone()[0] or 0
            add_message("LL mean in RES2 (day): {:,.1f} ...(night): {:,.1f}".format(total_LLday_res2, total_LLnight_res2), lst=2)

            # sum day life loss in RES3
            query_total_LLday_res3 = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND Occupancy_Type LIKE 'RES3%'"
            cursor.execute(query_total_LLday_res3,)
            total_LLday_res3 = cursor.fetchone()[0] or 0
            query_total_LLnight_res3 = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_night_input} WHERE Max_Depth > 0 AND Occupancy_Type LIKE 'RES3%'"
            cursor.execute(query_total_LLnight_res3,)
            total_LLnight_res3 = cursor.fetchone()[0] or 0
            add_message("LL mean in RES3 (day): {:,.1f} ...(night): {:,.1f}".format(total_LLday_res3, total_LLnight_res3), lst=2)

            # sum day life loss in RES4, RES5, or RES6
            query_total_LLday_res456 = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND Occupancy_Type IN ('RES4', 'RES5', 'RES6')"
            cursor.execute(query_total_LLday_res456,)
            total_LLday_res456 = cursor.fetchone()[0] or 0
            query_total_LLnight_res456 = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_night_input} WHERE Max_Depth > 0 AND Occupancy_Type IN ('RES4', 'RES5', 'RES6')"
            cursor.execute(query_total_LLnight_res456,)
            total_LLnight_res456 = cursor.fetchone()[0] or 0
            add_message("LL mean in RES4-6 (day): {:,.1f} ...(night): {:,.1f}".format(total_LLday_res456, total_LLnight_res456), lst=2)

            # sum day life loss in COM
            query_total_LLday_com = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND Occupancy_Type LIKE 'COM%'"
            cursor.execute(query_total_LLday_com,)
            total_LLday_com = cursor.fetchone()[0] or 0
            query_total_LLnight_com = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_night_input} WHERE Max_Depth > 0 AND Occupancy_Type LIKE 'COM%'"
            cursor.execute(query_total_LLnight_com,)
            total_LLnight_com = cursor.fetchone()[0] or 0
            add_message("LL mean in COM (day): {:,.1f} ...(night): {:,.1f}".format(total_LLday_com, total_LLnight_com), lst=2)

            # sum day life loss in IND
            query_total_LLday_ind = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND Occupancy_Type LIKE 'IND%'"
            cursor.execute(query_total_LLday_ind,)
            total_LLday_ind = cursor.fetchone()[0] or 0
            query_total_LLnight_ind = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_night_input} WHERE Max_Depth > 0 AND Occupancy_Type LIKE 'IND%'"
            cursor.execute(query_total_LLnight_ind,)
            total_LLnight_ind = cursor.fetchone()[0] or 0
            add_message("LL mean in IND (day): {:,.1f} ...(night): {:,.1f}".format(total_LLday_ind, total_LLnight_ind), lst=2)

            # sum day life loss in GOV
            query_total_LLday_gov = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND Occupancy_Type LIKE 'GOV%'"
            cursor.execute(query_total_LLday_gov,)
            total_LLday_gov = cursor.fetchone()[0] or 0
            query_total_LLnight_gov = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_night_input} WHERE Max_Depth > 0 AND Occupancy_Type LIKE 'GOV%'"
            cursor.execute(query_total_LLnight_gov,)
            total_LLnight_gov = cursor.fetchone()[0] or 0
            add_message("LL mean in GOV (day): {:,.1f} ...(night): {:,.1f}".format(total_LLday_gov, total_LLnight_gov), lst=2)

            # sum day life loss in REL
            query_total_LLday_rel = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND Occupancy_Type LIKE 'REL%'"
            cursor.execute(query_total_LLday_rel,)
            total_LLday_rel = cursor.fetchone()[0] or 0
            query_total_LLnight_rel = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_night_input} WHERE Max_Depth > 0 AND Occupancy_Type LIKE 'REL%'"
            cursor.execute(query_total_LLnight_rel,)
            total_LLnight_rel = cursor.fetchone()[0] or 0
            add_message("LL mean in REL (day): {:,.1f} ...(night): {:,.1f}".format(total_LLday_rel, total_LLnight_rel), lst=2)

            # sum day life loss in AGR
            query_total_LLday_agr = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND Occupancy_Type LIKE 'AGR%'"
            cursor.execute(query_total_LLday_agr,)
            total_LLday_agr = cursor.fetchone()[0] or 0
            query_total_LLnight_agr = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_night_input} WHERE Max_Depth > 0 AND Occupancy_Type LIKE 'AGR%'"
            cursor.execute(query_total_LLnight_agr,)
            total_LLnight_agr = cursor.fetchone()[0] or 0
            add_message("LL mean in AGR (day): {:,.1f} ...(night): {:,.1f}".format(total_LLday_agr, total_LLnight_agr), lst=2)

            # structures collapsed once
            # Initialize boolean to track if the "Collapsed" column exists
            collapsed_column_exists = True
            try:
                # Attempt to execute the query that uses the "Collapsed" column
                query_total_collapsedonce = f"SELECT COUNT(*) FROM {results1_day_input} WHERE Max_Depth > 0 AND Collapsed > 0"
                cursor.execute(query_total_collapsedonce)
                total_collapsedonce = cursor.fetchone()[0] or 0
                add_message("Num Structures that collapsed in at least one iteration: {:,.0f} of {:,.0f}".format(total_collapsedonce, total_structurecount_day), lst=2)
            except sqlite3.OperationalError as e:
                if "no such column: Collapsed" in str(e):
                    # If the error is due to the missing "Collapsed" column, set the flag to False
                    collapsed_column_exists = False
                    # Optionally, log a message indicating the column was not found
                    add_message("The 'Collapsed' column does not exist. Skipping collapsed structure count.", lst=2)
                else:
                    # Re-raise the error if it's something other than a missing column
                    raise

            # structures collapsed half of time
            # Only proceed with further queries if the "Collapsed" column exists boolean is true
            if collapsed_column_exists:
                fiftypercentiterations = numiterations / 2
                query_total_collapsed50p = f"SELECT COUNT(*) FROM {results1_day_input} WHERE Max_Depth > 0 AND Collapsed > {fiftypercentiterations}"
                cursor.execute(query_total_collapsed50p,)
                total_collapsed50p = cursor.fetchone()[0] or 0
                add_message("Num Structures that collapsed in more than half of iterations: {:,.0f} of {:,.0f}".format(total_collapsed50p, total_structurecount_day), lst=2)

                # structures collapsed all of time
                query_total_collapsedall = f"SELECT COUNT(*) FROM {results1_day_input} WHERE Max_Depth > 0 AND Collapsed = {numiterations}"
                cursor.execute(query_total_collapsedall,)
                total_collapsedall = cursor.fetchone()[0] or 0
                add_message("Num Structures that collapsed in all iterations: {:,.0f} of {:,.0f}".format(total_collapsedall, total_structurecount_day), lst=2)
            
            # ORIGINAL max fatality rate day
            # query_total_maxfatality_day = f"SELECT MAX(Life_Loss_Total_Mean / NULLIF((Pop_Under65_Mean + PAR_Over65_Mean), 0)) FROM {results1_day_input} WHERE Max_Depth > 0"
            # cursor.execute(query_total_maxfatality_day,)
            # total_maxfatality_day = cursor.fetchone()[0] or 0
            # add_message("Highest mean fatality rate (day): {:,.3f}".format(total_maxfatality_day), lst=2)

            # NEW max fatality rate day that also returns the FID value of the structure with max fatality rate, threshold PAR value avoids things with 0.002 PAR due to interpolation
            #query_total_maxfatality_day = f"SELECT FID, (Life_Loss_Total_Mean / NULLIF((Pop_Under65_Mean + PAR_Over65_Mean), 0)) AS Fatality_Rate FROM {results1_day_input} WHERE Max_Depth > 0 ORDER BY Fatality_Rate DESC LIMIT 1"
            ir_par_threshold=0.5
            add_message(f"PAR threshold for fatality rates: {ir_par_threshold}", lst=2)
            query_total_maxfatality_day = f"SELECT FID, (Life_Loss_Total_Mean / NULLIF((Pop_Under65_Mean + PAR_Over65_Mean), 0)) AS Fatality_Rate FROM {results1_day_input} WHERE Max_Depth > 0 AND (Pop_Under65_Mean + PAR_Over65_Mean) > ? ORDER BY Fatality_Rate DESC LIMIT 1"
            cursor.execute(query_total_maxfatality_day, (ir_par_threshold,))
            max_fatality_day_row = cursor.fetchone()
            # Check if a result was returned
            if max_fatality_day_row is not None:
                total_maxfatality_day = max_fatality_day_row[1] or 0

                # Report the highest mean fatality rate
                add_message(f"Highest mean fatality rate (day): {total_maxfatality_day:.3f}", lst=2)

                query_total_countmaxfatalityday = f"SELECT COUNT(*) FROM {results1_day_input} WHERE Max_Depth > 0 AND (Pop_Under65_Mean + PAR_Over65_Mean) > ? AND (Life_Loss_Total_Mean / NULLIF((Pop_Under65_Mean + PAR_Over65_Mean), 0)) = ?"
                cursor.execute(query_total_countmaxfatalityday, (ir_par_threshold, total_maxfatality_day,))
                total_countmaxfatalityday = cursor.fetchone()[0] or 0
                add_message("Count of structures with the max fatality rate (day): {:,.0f}".format(total_countmaxfatalityday), lst=2)
                # Create a SQL expression to select all rows with the maximum fatality rate
                maxfatalitydayfid = f"(Life_Loss_Total_Mean / NULLIF((Pop_Under65_Mean + PAR_Over65_Mean), 0)) = {total_maxfatality_day} AND (Pop_Under65_Mean + PAR_Over65_Mean) > {ir_par_threshold}"

                # If exporting GIS data, this exports the structures with the highest fatality rate
                if export:
                    try:
                        ir_output_day=fr"{outputgpkg}\IR_day_{sanitized_output_day}"
                        arcpy.conversion.ExportFeatures(
                            in_features=tablename_clean_day,
                            out_features=ir_output_day,
                            where_clause=maxfatalitydayfid,
                            use_field_alias_as_name="NOT_USE_ALIAS",
                            sort_field=None)

                        # Add a new field for fatality rates
                        fatality_rate_gis_field = "Fatality_Rate"
                        arcpy.management.AddField(
                            in_table=ir_output_day,
                            field_name=fatality_rate_gis_field,
                            field_type="DOUBLE",  # Choose an appropriate data type
                            field_precision=None,
                            field_scale=None,
                            field_length=None,
                            field_alias=None,
                            field_is_nullable="NULLABLE",
                            field_is_required="NON_REQUIRED"
                        )
                        
                        # Calculate the fatality rates, something wrong here
                        fatality_calc_expression = (
                            f"( !Life_Loss_Total_Mean! / ( !Pop_Under65_Mean! + !PAR_Over65_Mean! ) )"
                            f"if ( !Pop_Under65_Mean! + !PAR_Over65_Mean! ) > 0 else 0" 
                            )
                        arcpy.management.CalculateField(
                            in_table=ir_output_day,
                            field=fatality_rate_gis_field,
                            expression=fatality_calc_expression,
                            expression_type="PYTHON3"
                        )
                        
                        add_message("Structures with highest fatality rate exported for individual risk - day.", lst=2)
                        
                    except Exception as e:
                        arcpy.AddWarning(f"Export of IR failed - day. Error: {str(e)}")

            # NEW max fatality rate night that also returns the FID value of the structure with max fatality rate
            #query_total_maxfatality_night = f"SELECT FID, (Life_Loss_Total_Mean / NULLIF((Pop_Under65_Mean + PAR_Over65_Mean), 0)) AS Fatality_Rate FROM {results1_night_input} WHERE Max_Depth > 0 ORDER BY Fatality_Rate DESC LIMIT 1"
            query_total_maxfatality_night = f"SELECT FID, (Life_Loss_Total_Mean / NULLIF((Pop_Under65_Mean + PAR_Over65_Mean), 0)) AS Fatality_Rate FROM {results1_night_input} WHERE Max_Depth > 0 AND (Pop_Under65_Mean + PAR_Over65_Mean) > ? ORDER BY Fatality_Rate DESC LIMIT 1"
            cursor.execute(query_total_maxfatality_night, (ir_par_threshold,))
            max_fatality_night_row = cursor.fetchone()
            # Check if a result was returned
            if max_fatality_night_row is not None:
                total_maxfatality_night = max_fatality_night_row[1] or 0

                # Report the highest mean fatality rate
                add_message(f"Highest mean fatality rate (night): {total_maxfatality_night:.3f}", lst=2)

                query_total_countmaxfatalitynight = f"SELECT COUNT(*) FROM {results1_night_input} WHERE Max_Depth > 0 AND (Pop_Under65_Mean + PAR_Over65_Mean) > ? AND (Life_Loss_Total_Mean / NULLIF((Pop_Under65_Mean + PAR_Over65_Mean), 0)) = ?"
                cursor.execute(query_total_countmaxfatalitynight, (ir_par_threshold, total_maxfatality_night,))
                total_countmaxfatalitynight = cursor.fetchone()[0] or 0
                add_message("Count of structures with the max fatality rate (night): {:,.0f}".format(total_countmaxfatalitynight), lst=2)
                # Create a SQL expression to select all rows with the maximum fatality rate
                maxfatalitynightfid = f"(Life_Loss_Total_Mean / NULLIF((Pop_Under65_Mean + PAR_Over65_Mean), 0)) = {total_maxfatality_night} AND (Pop_Under65_Mean + PAR_Over65_Mean) > {ir_par_threshold}"
                
                # If exporting GIS data, this exports the structure with the highest fatality rate, allows this to be turned off in the code

                if export:
                    try:
                        ir_output_night=fr"{outputgpkg}\IR_night_{sanitized_output_night}"
                        arcpy.conversion.ExportFeatures(
                            in_features=tablename_clean_night,
                            out_features=ir_output_night,
                            where_clause=maxfatalitynightfid,
                            use_field_alias_as_name="NOT_USE_ALIAS",
                            sort_field=None)
                        
                        # Add a new field for fatality rates
                        fatality_rate_gis_field = "Fatality_Rate"
                        arcpy.management.AddField(
                            in_table=ir_output_night,
                            field_name=fatality_rate_gis_field,
                            field_type="DOUBLE",  # Choose an appropriate data type
                            field_precision=None,
                            field_scale=None,
                            field_length=None,
                            field_alias=None,
                            field_is_nullable="NULLABLE",
                            field_is_required="NON_REQUIRED"
                        )
                        
                        # Calculate the fatality rates
                        fatality_calc_expression = (
                            f"( !Life_Loss_Total_Mean! / ( !Pop_Under65_Mean! + !PAR_Over65_Mean! ) )"
                            f"if ( !Pop_Under65_Mean! + !PAR_Over65_Mean! ) > 0 else 0" 
                            )
                        arcpy.management.CalculateField(
                            in_table=ir_output_night,
                            field=fatality_rate_gis_field,
                            expression=fatality_calc_expression,
                            expression_type="PYTHON3"
                        )
                        
                        add_message("Structures with highest fatality rate exported for individual risk - night.", lst=2)
                        
                        
                    except Exception as e:
                        arcpy.AddWarning(f"Export of IR failed - night. Error: {str(e)}")


            # # of structs with over 90% fatality rate
            query_total_over90fatality_day = f"SELECT COUNT(*) FROM {results1_day_input} WHERE Max_Depth > 0 AND (Pop_Under65_Mean + PAR_Over65_Mean) > ? AND (Life_Loss_Total_Mean / NULLIF((Pop_Under65_Mean + PAR_Over65_Mean), 0)) > 0.9"
            cursor.execute(query_total_over90fatality_day,(ir_par_threshold,))
            total_over90fatalitycount_day = cursor.fetchone()[0] or 0
            add_message("Count of structures with over 90 percent fatality rate (day): {:,.0f}".format(total_over90fatalitycount_day), lst=2)

            # # of structs with over 90% fatality rate
            query_total_over90fatality_night = f"SELECT COUNT(*) FROM {results1_night_input} WHERE Max_Depth > 0 AND (Pop_Under65_Mean + PAR_Over65_Mean) > ? AND (Life_Loss_Total_Mean / NULLIF((Pop_Under65_Mean + PAR_Over65_Mean), 0)) > 0.9"
            cursor.execute(query_total_over90fatality_night,(ir_par_threshold,))
            total_over90fatalitycount_night = cursor.fetchone()[0] or 0
            add_message("Count of structures with over 90 percent fatality rate (night): {:,.0f}".format(total_over90fatalitycount_night), lst=2)

            # excel For total table (first table)
            start_row_total = 1
            start_col_total = 3

            # excel Define where the tables should start for totals (e.g., row 1, column 3)
            row = start_row_total
            col = start_col_total
            # Set headers for columns in total table
            ws.cell(row=row, column=col, value="Alternative")
            ws.cell(row=row, column=col+1, value="Structure #")
            ws.cell(row=row, column=col+2, value="PAR Day")
            ws.cell(row=row, column=col+3, value="PAR Night")
            ws.cell(row=row, column=col+4, value="LL Day")
            ws.cell(row=row, column=col+5, value="LL Night")
            ws.cell(row=row, column=col+6, value="Damage Total")
            if total_lifeloss_evac_day > 0 or total_lifeloss_evac_night > 0:
                ws.cell(row=row, column=col+7, value="LL Evac Day")
                ws.cell(row=row, column=col+8, value="LL Evac Night")
            # excel Increment row for EPZ after the column headings
            current_row_total = start_row_total + 1

            # excel Add relevant data starting from row 2
            ws.cell(row=current_row_total, column=start_col_total, value=alternative1)
            ws.cell(row=current_row_total, column=start_col_total+1, value=total_structurecount_day)
            ws.cell(row=current_row_total, column=start_col_total+2, value=total_par_day)
            ws.cell(row=current_row_total, column=start_col_total+3, value=total_par_night)
            ws.cell(row=current_row_total, column=start_col_total+4, value=total_lifeloss_day)
            ws.cell(row=current_row_total, column=start_col_total+5, value=total_lifeloss_night)
            ws.cell(row=current_row_total, column=start_col_total+6, value=total_damage_day)
            if total_lifeloss_evac_day > 0 or total_lifeloss_evac_night > 0:
                ws.cell(row=current_row_total, column=start_col_total+7, value=total_lifeloss_evac_day)
                ws.cell(row=current_row_total, column=start_col_total+8, value=total_lifeloss_evac_night)

            # Apply number formatting to the relevant cells
            for col_total in range(start_col_total, start_col_total + 9):  # Adjust range as necessary
                cell = ws.cell(row=current_row_total, column=col_total)
                cell.number_format = '#,##0.0'  # Set the desired number format

            # excel Move a few rows down to start the arrival tables
            current_row_total += 3
            
            # Set headers for columns in arrival table
            ws.cell(row=current_row_total, column=col, value="Arrival Zone")
            ws.cell(row=current_row_total, column=col+1, value="Structure #")
            ws.cell(row=current_row_total, column=col+2, value="PAR Day")
            ws.cell(row=current_row_total, column=col+3, value="PAR Night")
            ws.cell(row=current_row_total, column=col+4, value="LL Day")
            ws.cell(row=current_row_total, column=col+5, value="LL Night")
            ws.cell(row=current_row_total, column=col+6, value="Depth Range")
            ws.cell(row=current_row_total, column=col+7, value="Velocity Range")
            ws.cell(row=current_row_total, column=col+8, value="% Warned (Day)")
            ws.cell(row=current_row_total, column=col+9, value="% Mobilized (Day)")
            ws.cell(row=current_row_total, column=col+10, value="Fatality Rate (Day)")

            # excel Increment row for arrival data after the column headings
            current_row_total += 1

            # excel Add relevant arrival data starting from row 5

            ws.cell(row=current_row_total, column=start_col_total, value=arrival_zone1)
            ws.cell(row=current_row_total, column=start_col_total+1, value=struc_count_day_arzone1_sum)
            ws.cell(row=current_row_total, column=start_col_total+2, value=round(par_day_arzone1_sum if par_day_arzone1_sum is not None else 0, 1))
            ws.cell(row=current_row_total, column=start_col_total+3, value=round(par_night_arzone1_sum if par_night_arzone1_sum is not None else 0, 1))
            ws.cell(row=current_row_total, column=start_col_total+4, value=round(total_lifeloss_day_arzone1_sum if total_lifeloss_day_arzone1_sum is not None else 0, 1))
            ws.cell(row=current_row_total, column=start_col_total+5, value=round(total_lifeloss_night_arzone1_sum if total_lifeloss_night_arzone1_sum is not None else 0, 1))
            ws.cell(row=current_row_total, column=start_col_total+6, value=arzone1_depth_range)
            ws.cell(row=current_row_total, column=start_col_total+7, value=arzone1_velocity_range)
            ws.cell(row=current_row_total, column=start_col_total+8, value=par_warned_day_arzone1_percentage if par_warned_day_arzone1_percentage is not None else 0)
            ws.cell(row=current_row_total, column=start_col_total+9, value=par_mobilized_day_arzone1_percentage if par_mobilized_day_arzone1_percentage is not None else 0)
            ws.cell(row=current_row_total, column=start_col_total+10, value=round(fatality_arzone1 if fatality_arzone1 is not None else 0, 3))
            #set Percent Warned and percent mobilized to 1 decimal place %
            cell = ws.cell(row=current_row_total, column=start_col_total+8)
            cell.number_format = '0.0%'  # Set the desired number format    
            cell = ws.cell(row=current_row_total, column=start_col_total+9)
            cell.number_format = '0.0%'  # Set the desired number format
            current_row_total += 1

            ws.cell(row=current_row_total, column=start_col_total, value=arrival_zone2)
            ws.cell(row=current_row_total, column=start_col_total+1, value=struc_count_day_arzone2_sum)
            ws.cell(row=current_row_total, column=start_col_total+2, value=round(par_day_arzone2_sum if par_day_arzone2_sum is not None else 0, 1))
            ws.cell(row=current_row_total, column=start_col_total+3, value=round(par_night_arzone2_sum if par_night_arzone2_sum is not None else 0, 1))
            ws.cell(row=current_row_total, column=start_col_total+4, value=round(total_lifeloss_day_arzone2_sum if total_lifeloss_day_arzone2_sum is not None else 0, 1))
            ws.cell(row=current_row_total, column=start_col_total+5, value=round(total_lifeloss_night_arzone2_sum if total_lifeloss_night_arzone2_sum is not None else 0, 1))
            ws.cell(row=current_row_total, column=start_col_total+6, value=arzone2_depth_range)
            ws.cell(row=current_row_total, column=start_col_total+7, value=arzone2_velocity_range)
            ws.cell(row=current_row_total, column=start_col_total+8, value=par_warned_day_arzone2_percentage if par_warned_day_arzone2_percentage is not None else 0)
            ws.cell(row=current_row_total, column=start_col_total+9, value=par_mobilized_day_arzone2_percentage if par_mobilized_day_arzone2_percentage is not None else 0)
            ws.cell(row=current_row_total, column=start_col_total+10, value=round(fatality_arzone2 if fatality_arzone2 is not None else 0, 3))
            #set Percent Warned and percent mobilized to 1 decimal place %
            cell = ws.cell(row=current_row_total, column=start_col_total+8)
            cell.number_format = '0.0%'  # Set the desired number format    
            cell = ws.cell(row=current_row_total, column=start_col_total+9)
            cell.number_format = '0.0%'  # Set the desired number format
            current_row_total += 1
            ws.cell(row=current_row_total, column=start_col_total, value=arrival_zone3)
            ws.cell(row=current_row_total, column=start_col_total+1, value=struc_count_day_arzone3_sum)
            ws.cell(row=current_row_total, column=start_col_total+2, value=round(par_day_arzone3_sum if par_day_arzone3_sum is not None else 0, 1))
            ws.cell(row=current_row_total, column=start_col_total+3, value=round(par_night_arzone3_sum if par_night_arzone3_sum is not None else 0, 1))
            ws.cell(row=current_row_total, column=start_col_total+4, value=round(total_lifeloss_day_arzone3_sum if total_lifeloss_day_arzone3_sum is not None else 0, 1))
            ws.cell(row=current_row_total, column=start_col_total+5, value=round(total_lifeloss_night_arzone3_sum if total_lifeloss_night_arzone3_sum is not None else 0, 1))
            ws.cell(row=current_row_total, column=start_col_total+6, value=arzone3_depth_range)
            ws.cell(row=current_row_total, column=start_col_total+7, value=arzone3_velocity_range)
            ws.cell(row=current_row_total, column=start_col_total+8, value=par_warned_day_arzone3_percentage if par_warned_day_arzone3_percentage is not None else 0)
            ws.cell(row=current_row_total, column=start_col_total+9, value=par_mobilized_day_arzone3_percentage if par_mobilized_day_arzone3_percentage is not None else 0)
            ws.cell(row=current_row_total, column=start_col_total+10, value=round(fatality_arzone3 if fatality_arzone3 is not None else 0, 3))
            #set Percent Warned and percent mobilized to 1 decimal place %
            cell = ws.cell(row=current_row_total, column=start_col_total+8)
            cell.number_format = '0.0%'  # Set the desired number format    
            cell = ws.cell(row=current_row_total, column=start_col_total+9)
            cell.number_format = '0.0%'  # Set the desired number format
            current_row_total += 1
            ws.cell(row=current_row_total, column=start_col_total, value=arrival_zone4)
            ws.cell(row=current_row_total, column=start_col_total+1, value=struc_count_day_arzone4_sum)
            ws.cell(row=current_row_total, column=start_col_total+2, value=round(par_day_arzone4_sum if par_day_arzone4_sum is not None else 0, 1))
            ws.cell(row=current_row_total, column=start_col_total+3, value=round(par_night_arzone4_sum if par_night_arzone4_sum is not None else 0, 1))
            ws.cell(row=current_row_total, column=start_col_total+4, value=round(total_lifeloss_day_arzone4_sum if total_lifeloss_day_arzone4_sum is not None else 0, 1))
            ws.cell(row=current_row_total, column=start_col_total+5, value=round(total_lifeloss_night_arzone4_sum if total_lifeloss_night_arzone4_sum is not None else 0, 1))
            ws.cell(row=current_row_total, column=start_col_total+6, value=arzone4_depth_range)
            ws.cell(row=current_row_total, column=start_col_total+7, value=arzone4_velocity_range)
            ws.cell(row=current_row_total, column=start_col_total+8, value=par_warned_day_arzone4_percentage if par_warned_day_arzone4_percentage is not None else 0)
            ws.cell(row=current_row_total, column=start_col_total+9, value=par_mobilized_day_arzone4_percentage if par_mobilized_day_arzone4_percentage is not None else 0)
            ws.cell(row=current_row_total, column=start_col_total+10, value=round(fatality_arzone4 if fatality_arzone4 is not None else 0, 3))
            #set Percent Warned and percent mobilized to 1 decimal place %
            cell = ws.cell(row=current_row_total, column=start_col_total+8)
            cell.number_format = '0.0%'  # Set the desired number format    
            cell = ws.cell(row=current_row_total, column=start_col_total+9)
            cell.number_format = '0.0%'  # Set the desired number format
            current_row_total += 1
            ws.cell(row=current_row_total, column=start_col_total, value=arrival_zone5)
            ws.cell(row=current_row_total, column=start_col_total+1, value=struc_count_day_arzone5_sum)
            ws.cell(row=current_row_total, column=start_col_total+2, value=round(par_day_arzone5_sum if par_day_arzone5_sum is not None else 0, 1))
            ws.cell(row=current_row_total, column=start_col_total+3, value=round(par_night_arzone5_sum if par_night_arzone5_sum is not None else 0, 1))
            ws.cell(row=current_row_total, column=start_col_total+4, value=round(total_lifeloss_day_arzone5_sum if total_lifeloss_day_arzone5_sum is not None else 0, 1))
            ws.cell(row=current_row_total, column=start_col_total+5, value=round(total_lifeloss_night_arzone5_sum if total_lifeloss_night_arzone5_sum is not None else 0, 1))
            ws.cell(row=current_row_total, column=start_col_total+6, value=arzone5_depth_range)
            ws.cell(row=current_row_total, column=start_col_total+7, value=arzone5_velocity_range)
            ws.cell(row=current_row_total, column=start_col_total+8, value=par_warned_day_arzone5_percentage if par_warned_day_arzone5_percentage is not None else 0)
            ws.cell(row=current_row_total, column=start_col_total+9, value=par_mobilized_day_arzone5_percentage if par_mobilized_day_arzone5_percentage is not None else 0)
            ws.cell(row=current_row_total, column=start_col_total+10, value=round(fatality_arzone5 if fatality_arzone5 is not None else 0, 3))
            #set Percent Warned and percent mobilized to 1 decimal place %
            cell = ws.cell(row=current_row_total, column=start_col_total+8)
            cell.number_format = '0.0%'  # Set the desired number format    
            cell = ws.cell(row=current_row_total, column=start_col_total+9)
            cell.number_format = '0.0%'  # Set the desired number format
            current_row_total += 1
            ws.cell(row=current_row_total, column=start_col_total, value=arrival_zone6)
            ws.cell(row=current_row_total, column=start_col_total+1, value=struc_count_day_arzone6_sum)
            ws.cell(row=current_row_total, column=start_col_total+2, value=round(par_day_arzone6_sum if par_day_arzone6_sum is not None else 0, 1))
            ws.cell(row=current_row_total, column=start_col_total+3, value=round(par_night_arzone6_sum if par_night_arzone6_sum is not None else 0, 1))
            ws.cell(row=current_row_total, column=start_col_total+4, value=round(total_lifeloss_day_arzone6_sum if total_lifeloss_day_arzone6_sum is not None else 0, 1))
            ws.cell(row=current_row_total, column=start_col_total+5, value=round(total_lifeloss_night_arzone6_sum if total_lifeloss_night_arzone6_sum is not None else 0, 1))
            ws.cell(row=current_row_total, column=start_col_total+6, value=arzone6_depth_range)
            ws.cell(row=current_row_total, column=start_col_total+7, value=arzone6_velocity_range)
            ws.cell(row=current_row_total, column=start_col_total+8, value=par_warned_day_arzone6_percentage  if par_warned_day_arzone6_percentage is not None else 0)
            ws.cell(row=current_row_total, column=start_col_total+9, value=par_mobilized_day_arzone6_percentage  if par_mobilized_day_arzone6_percentage is not None else 0)
            ws.cell(row=current_row_total, column=start_col_total+10, value=round(fatality_arzone6 if fatality_arzone6 is not None else 0, 3))
            #set Percent Warned and percent mobilized to 1 decimal place %
            cell = ws.cell(row=current_row_total, column=start_col_total+8)
            cell.number_format = '0.0%'  # Set the desired number format    
            cell = ws.cell(row=current_row_total, column=start_col_total+9)
            cell.number_format = '0.0%'  # Set the desired number format
            current_row_total += 1
            ws.cell(row=current_row_total, column=start_col_total, value=arrival_zone7)
            ws.cell(row=current_row_total, column=start_col_total+1, value=struc_count_day_arzone7_sum)
            ws.cell(row=current_row_total, column=start_col_total+2, value=round(par_day_arzone7_sum if par_day_arzone7_sum is not None else 0, 1))
            ws.cell(row=current_row_total, column=start_col_total+3, value=round(par_night_arzone7_sum if par_night_arzone7_sum is not None else 0, 1))
            ws.cell(row=current_row_total, column=start_col_total+4, value=round(total_lifeloss_day_arzone7_sum if total_lifeloss_day_arzone7_sum is not None else 0, 1))
            ws.cell(row=current_row_total, column=start_col_total+5, value=round(total_lifeloss_night_arzone7_sum if total_lifeloss_night_arzone7_sum is not None else 0, 1))
            ws.cell(row=current_row_total, column=start_col_total+6, value=arzone7_depth_range)
            ws.cell(row=current_row_total, column=start_col_total+7, value=arzone7_velocity_range)
            ws.cell(row=current_row_total, column=start_col_total+8, value=par_warned_day_arzone7_percentage if par_warned_day_arzone7_percentage is not None else 0)
            ws.cell(row=current_row_total, column=start_col_total+9, value=par_mobilized_day_arzone7_percentage  if par_mobilized_day_arzone7_percentage is not None else 0)
            ws.cell(row=current_row_total, column=start_col_total+10, value=round(fatality_arzone7 if fatality_arzone7 is not None else 0, 3))
            #set Percent Warned and percent mobilized to 1 decimal place %
            cell = ws.cell(row=current_row_total, column=start_col_total+8)
            cell.number_format = '0.0%'  # Set the desired number format    
            cell = ws.cell(row=current_row_total, column=start_col_total+9)
            cell.number_format = '0.0%'  # Set the desired number format
            current_row_total += 1
            ws.cell(row=current_row_total, column=start_col_total, value=arrival_zone8)
            ws.cell(row=current_row_total, column=start_col_total+1, value=struc_count_day_arzone8_sum)
            ws.cell(row=current_row_total, column=start_col_total+2, value=round(par_day_arzone8_sum if par_day_arzone8_sum is not None else 0, 1))
            ws.cell(row=current_row_total, column=start_col_total+3, value=round(par_night_arzone8_sum if par_night_arzone8_sum is not None else 0, 1))
            ws.cell(row=current_row_total, column=start_col_total+4, value=round(total_lifeloss_day_arzone8_sum if total_lifeloss_day_arzone8_sum is not None else 0, 1))
            ws.cell(row=current_row_total, column=start_col_total+5, value=round(total_lifeloss_night_arzone8_sum if total_lifeloss_night_arzone8_sum is not None else 0, 1))

            # excel Move a few rows down to start the EPZ tables
            current_row_total += 3

            #Write data above to the summary worksheet also
            ws = wb["Summary"]
            ws.cell(row=current_row_summary, column=start_col_summary, value=alternative1)
            ws.cell(row=current_row_summary, column=start_col_summary+1, value=total_structurecount_day)
            ws.cell(row=current_row_summary, column=start_col_summary+2, value=total_par_day)
            ws.cell(row=current_row_summary, column=start_col_summary+3, value=total_par_night)
            ws.cell(row=current_row_summary, column=start_col_summary+4, value=total_lifeloss_day)
            ws.cell(row=current_row_summary, column=start_col_summary+5, value=total_lifeloss_night)
            ws.cell(row=current_row_summary, column=start_col_summary+6, value=total_damage_day)
            ws.cell(row=current_row_summary, column=start_col_summary+8, value=total_lifeloss_evac_day)
            ws.cell(row=current_row_summary, column=start_col_summary+9, value=total_lifeloss_evac_night)

            # Apply number formatting to the relevant cells
            for col_summary in range(start_col_summary + 1, start_col_summary + 10):  # Adjust range as necessary
                cellsummary = ws.cell(row=current_row_summary, column=col_summary)
                cellsummary.number_format = '#,##0'  # Set the desired number format
            #add one to move down to next row in the next loop
            current_row_summary += 1
            #switch active worksheet back to the alternative sheet (use sanitized sheet name)
            ws = wb[sanitized_sheet]
            
            #get EPZ set name and field name for individual epz polygons
            epz_source_query = f"SELECT Emergency_Planning_Zones FROM Alternatives_Lookup_Table WHERE Name = ?;"
            cursor.execute(epz_source_query, (alternative1,))
            epz_source = cursor.fetchone()[0]

            epz_namefield_query = f"SELECT Zone_Name_Attribute FROM Emergency_Planning_Zone_Lookup_Table WHERE Name = ?;"
            cursor.execute(epz_namefield_query, (epz_source,))
            epz_namefield = cursor.fetchone()[0]
            add_message("EPZ Source: {0}    EPZ name field: {1}".format(epz_source, epz_namefield), lst=2)

            #get list of epzs from the epz table
            epztable = fr'"{epz_source}"'
            epzlistquery = f"SELECT {epz_namefield} FROM {epztable}"
            cursor.execute(epzlistquery)
                # Fetch all results
            epzrows = cursor.fetchall()
                # Extract values from the rows and create a list
            epz_list = [row[0] for row in epzrows]
            #epz_count = len(epz_list)

             # excel For EPZ table
            start_row_epz = current_row_total
            start_col_epz = 3

            # excel Define where the table should start for epzs (e.g., row 10, column 2)
            row = start_row_epz
            col = start_col_epz
            # Set column headers
            ws.cell(row=row, column=col, value="EPZ Name")
            ws.cell(row=row, column=col+1, value="Struc Total")
            ws.cell(row=row, column=col+2, value="PAR Day")
            ws.cell(row=row, column=col+3, value="PAR Night")
            ws.cell(row=row, column=col+4, value="LL Day")
            ws.cell(row=row, column=col+5, value="LL Night")
            ws.cell(row=row, column=col+6, value="Struc PreHaz")
            ws.cell(row=row, column=col+7, value="Struc PostHaz")
            ws.cell(row=row, column=col+8, value="PAR PreHaz (day)")
            ws.cell(row=row, column=col+9, value="PAR PostHaz (day)")
            ws.cell(row=row, column=col+10, value="LL PreHaz (day)")
            ws.cell(row=row, column=col+11, value="LL PostHaz (day)")
            ws.cell(row=row, column=col+12, value="% Warned")
            ws.cell(row=row, column=col+13, value="% Mobilized")
            ws.cell(row=row, column=col+14, value="Fatality Rate (day)")
            # excel Increment row for EPZ after the header
            current_row_epz = start_row_epz + 1
            current_row_total += 1

            ## LOOP B Begin looping through EPZ list, DistributionData table doesn't have named EPZ's just an order 0 to X
            epz_listorder = 0
            for epz in epz_list:
                # EPZ Name
                add_message("{} EPZ".format(epz), lst=2)
                ## Pull EPZ Parameters for issuance delay, first alert, and PAI
                epz_issuancedelay_query = f"SELECT Issuance_Delay FROM {epztable} WHERE {epz_namefield} = ?;"
                cursor.execute(epz_issuancedelay_query, (epz,))
                epz_issuancedelay = cursor.fetchone()[0]
                add_message("...Issuance Delay = {0}".format(epz_issuancedelay), lst=2)
                if epz_issuancedelay != "Preparedness Unknown" and mmc_sop == True:
                        add_message("...MMC SOP Violation - Warning Issuance Delay is not Preparedness Unknown", lst=2)

                epz_firstalert_query = f"SELECT First_Alert_Diffusion FROM {epztable} WHERE {epz_namefield} = ?;"
                cursor.execute(epz_firstalert_query, (epz,))
                epz_firstalert = cursor.fetchone()[0]
                add_message("...First Alert Diffusion = {0}".format(epz_firstalert), lst=2)
                if epz_firstalert != "Unknown" and mmc_sop == True:
                        add_message("...MMC SOP Violation - First Alert Diffusion curve is not Unknown", lst=2)

                epz_pai_query = f"SELECT PAI_Diffusion FROM {epztable} WHERE {epz_namefield} = ?;"
                cursor.execute(epz_pai_query, (epz,))
                epz_pai = cursor.fetchone()[0]
                add_message("...PAI = {0}".format(epz_pai), lst=2)
                if epz_pai != "Perception: Unknown / Preparedness: Unknown" and mmc_sop == True:
                        add_message("...MMC SOP Violation - PAI curve is not Perception: Unknown / Preparedness: Unknown", lst=2)

                ## Query EPZ Warning Time Distribution and values
                epz_distribution_hazname = f"Hazard_Identified_{epz_listorder}"
                #add_message("...EPZ Distribution Data Name: {}".format(epz_distribution_hazname))
                
                epz_warndistribution_query = f"SELECT Distribution FROM '{alternative1}>DistributionData' WHERE Name = ?;"
                cursor.execute(epz_warndistribution_query, (epz_distribution_hazname,))
                epz_warndistribution = cursor.fetchone()[0]           
                #add_message("...EPZ Warn Distribution: {}".format(epz_warndistribution), lst=2)
                
                if epz_warndistribution == "Uniform":
                    epz_warndistribution_min_query = f"SELECT Minimum FROM '{alternative1}>DistributionData' WHERE Name = ?;"
                    cursor.execute(epz_warndistribution_min_query, (epz_distribution_hazname,))
                    epz_warndistribution_min = cursor.fetchone()[0]

                    epz_warndistribution_max_query = f"SELECT Maximum FROM '{alternative1}>DistributionData' WHERE Name = ?;"
                    cursor.execute(epz_warndistribution_max_query, (epz_distribution_hazname,))
                    epz_warndistribution_max = cursor.fetchone()[0]

                    epz_earliestwarning_minutes = (BreachTime_in_minutes + (epz_warndistribution_min * 60))
                    add_message("...Warning Hazard ID Time, Uniform Distribution, min: {0}  max: {1}".format(epz_warndistribution_min, epz_warndistribution_max), lst=2)
                    
                    if mmc_sop == True:
                        if epz_warndistribution_max > 0.5:
                            add_message("...MMC SOP Violation: Max warning time is greater than +0.5 hours after the hazard", lst=2)
                        if epz_warndistribution_max < -2:
                            add_message("...MMC SOP Violation: Max warning time is earlier than -2 hours before the hazard", lst=2)
                        if epz_warndistribution_min < -6:
                            add_message("...MMC SOP Violation: Min warning time is earlier than -6 hours before the hazard", lst=2)
                        if epz_warndistribution_min > -2:
                            add_message("...MMC SOP Violation: Min warning time is later than -2 hours", lst=2)
                        
                        if "min" in alternative1.lower(): #checks against both levee and dam minimal warning times
                            add_message("...This looks like a minimal warning scenario, it has min in the name", lst=2)
                            if epz_warndistribution_max != 0 and epz_warndistribution_max != 0.5:
                                add_message("...MMC SOP Violation: Max warning time for a min alternative should be 0 for dams or 0.5 for levees", lst=2)
                            if epz_warndistribution_min not in (-3, -2):
                                add_message("...MMC SOP Violation: Min warning time for a min alternative should be -2 for dams or -3 for levees", lst=2)

                        if "amp" in alternative1.lower(): #for levee ample warning, there shouldn't be an uniform distribution
                            add_message("...This looks like an ample warning scenario, it has amp in the name", lst=2)
                            if epz_warndistribution_max != -2:
                                add_message("...MMC SOP Violation: Max warning time for an amp alternative should be -2 for dams", lst=2)
                            if epz_warndistribution_min != -6:
                                add_message("...MMC SOP Violation: Min warning time for an amp alternative should be -6 for dams", lst=2)
                        
                        # checks for anything that makes this look like a non-fail scenario, except levees that might have a -3 to 0.5 nonfail
                        if breach_condition == "NonFail" and epz_warndistribution_min != -3:
                            add_message("...MMC SOP Violation: This looks like a non-fail dam scenario, it shouldn't need a uniform distribution warning", lst=2)

                elif epz_warndistribution == "None":
                    epz_warndistribution_mean_query = f"SELECT Mean FROM '{alternative1}>DistributionData' WHERE Name = ?;"
                    cursor.execute(epz_warndistribution_mean_query, (epz_distribution_hazname,))
                    epz_warndistribution_mean = cursor.fetchone()[0]
                    
                    epz_earliestwarning_minutes = (BreachTime_in_minutes + (epz_warndistribution_mean * 60))
                    add_message("...Warning Hazard ID Time, no distribution, mean: {}".format(epz_warndistribution_mean), lst=2)

                    if epz_warndistribution_mean > 0:
                        add_message("...POSSIBLE ISSUE Warning time with no uncertainty is after the hazard", lst=2)
                    if mmc_sop == True:
                        if epz_warndistribution_mean > -24 and epz_warndistribution_mean < 0:
                            add_message("...MMC SOP Violation: Warning time with no uncertainty is within 24 hours of the hazard", lst=2)
                        if epz_warndistribution_mean > -(BreachTime_in_hours) and epz_warndistribution_mean != -24: #-24 is probably a levee
                            add_message("...MMC SOP Violation: Warning time with no uncertainty is not before the start of the simulation", lst=2)
                        # Check if mean is a multiple of 24
                        if not math.isclose(epz_warndistribution_mean % 24, 0, abs_tol=1e-6):
                            add_message("...MMC SOP Violation: Warning time with no uncertainty is not a multiple of 24 hours", lst=2)

                elif epz_warndistribution == "Triangular":
                    epz_warndistribution_min_query = f"SELECT Minimum FROM '{alternative1}>DistributionData' WHERE Name = ?;"
                    cursor.execute(epz_warndistribution_min_query, (epz_distribution_hazname,))
                    epz_warndistribution_min = cursor.fetchone()[0]

                    epz_warndistribution_max_query = f"SELECT Maximum FROM '{alternative1}>DistributionData' WHERE Name = ?;"
                    cursor.execute(epz_warndistribution_max_query, (epz_distribution_hazname,))
                    epz_warndistribution_max = cursor.fetchone()[0]

                    epz_warndistribution_mostlikely_query = f"SELECT Most_Likely FROM '{alternative1}>DistributionData' WHERE Name = ?;"
                    cursor.execute(epz_warndistribution_mostlikely_query, (epz_distribution_hazname,))
                    epz_warndistribution_mostlikely = cursor.fetchone()[0]

                    epz_earliestwarning_minutes = (BreachTime_in_minutes + (epz_warndistribution_min * 60))
                    add_message("...Warning Hazard ID Time, Triangular Distribution, min: {0}  most likely: {1}  max: {2}".format(epz_warndistribution_min, epz_warndistribution_mostlikely, epz_warndistribution_max), lst=2)

                    if mmc_sop == True:
                        add_message("...MMC SOP Violation: Warning Distribution is Triangle, not uniform or none", lst=2)
                
                else:
                    error_message = f"Error: Unexpected EPZ Warning Distribution: {epz_warndistribution}. Should be Uniform, Triangular, or None."
                    add_message(error_message, lst=2)
                    raise ValueError(error_message)

                ## Query EPZ EMA notification communication delay distribution and values
                epz_distribution_notifyema_name = f"Notify_EMA_{epz_listorder}"
                
                epz_notifyema_distribution_query = f"SELECT Distribution FROM '{alternative1}>DistributionData' WHERE Name = ?;"
                cursor.execute(epz_notifyema_distribution_query, (epz_distribution_notifyema_name,))
                epz_notifyema_distribution = cursor.fetchone()[0] 

                if epz_notifyema_distribution == "Uniform":
                    epz_notifyema_min_query = f"SELECT Minimum FROM '{alternative1}>DistributionData' WHERE Name = ?;"
                    cursor.execute(epz_notifyema_min_query, (epz_distribution_notifyema_name,))
                    epz_notifyema_min = cursor.fetchone()[0] 

                    if epz_notifyema_min < 0.01 and mmc_sop == True:
                        add_message("...MMC SOP Violation: Notify EMA min is less than 0.01", lst=2)

                    epz_notifyema_max_query = f"SELECT Maximum FROM '{alternative1}>DistributionData' WHERE Name = ?;"
                    cursor.execute(epz_notifyema_max_query, (epz_distribution_notifyema_name,))
                    epz_notifyema_max = cursor.fetchone()[0] 
                    add_message("...Notify EMA (Com Delay), Uniform Distribution, min: {0}  max: {1}".format(epz_notifyema_min, epz_notifyema_max), lst=2)

                    if epz_notifyema_max > 0.5 and mmc_sop == True:
                        add_message("...MMC SOP Violation: Notify EMA max is greater than 0.5 hours", lst=2)

                elif epz_notifyema_distribution == "None":
                    epz_notifyema_mean_query = f"SELECT Mean FROM '{alternative1}>DistributionData' WHERE Name = ?;"
                    cursor.execute(epz_notifyema_mean_query, (epz_distribution_notifyema_name,))
                    epz_notifyema_mean = cursor.fetchone()[0] 
                    add_message("...Notify EMA (Com Delay), no distribution, mean: {}".format(epz_notifyema_mean), lst=2)

                    if epz_notifyema_mean != 0 and mmc_sop == True:
                        add_message("...MMC SOP Violation: Notify EMA is distribution is none, but it's not zero", lst=2)

                    if epz_warndistribution == "Uniform":
                        add_message("...MMC SOP Violation: Notify EMA is distribution is none, but the warning hazard ID time is uniform, looks like a breach zone", lst=2)

                else:
                    error_message1 = f"Error: Unexpected Communication Delay Distribution: {epz_notifyema_distribution}. Should be Uniform or None."
                    add_message(error_message1, lst=2)
                    raise ValueError(error_message1)

                epz_listorder += 1

                # Count structures in each EPZ
                # SQL query to count features with depth greater than zero and arrival time pre and post hazard the numeric field above the threshold
                epz_structurecount_day_query = f"SELECT COUNT(*) FROM {results1_day_input} WHERE Emergency_Zone = ? AND Max_Depth > 0"
                try:
                    cursor.execute(epz_structurecount_day_query, (epz,))
                    epz_structurecount_day = cursor.fetchone()[0]
                except sqlite3.OperationalError as e:
                    # If EPZ queries fail because the results table is missing, skip the rest of this alternative
                    if 'no such table' in str(e).lower():
                        add_message(f"WARNING: Results table not found during EPZ processing for alternative '{alternative1}': {results1_day_input}. Skipping this alternative.", lst=2)
                        skip_alternative = True
                        # break out of EPZ loop
                        break
                    else:
                        raise
                except Exception as e:
                    raise
                #add_message("...Structure count w/ depth>0 (day): {}".format(epz_structurecount_day), lst=2)

                epz_PAR_u65_day_query = f"SELECT SUM(Pop_Under65_Mean) FROM {results1_day_input} WHERE Emergency_Zone = ? AND Max_Depth > 0"
                cursor.execute(epz_PAR_u65_day_query, (epz,))
                epz_PAR_u65_day = cursor.fetchone()[0] or 0

                epz_PAR_o65_day_query = f"SELECT SUM(PAR_Over65_Mean) FROM {results1_day_input} WHERE Emergency_Zone = ? AND Max_Depth > 0"
                cursor.execute(epz_PAR_o65_day_query, (epz,))
                epz_PAR_o65_day = cursor.fetchone()[0] or 0
                epz_PAR_day = epz_PAR_u65_day + epz_PAR_o65_day

                epz_PAR_u65_night_query = f"SELECT SUM(Pop_Under65_Mean) FROM {results1_night_input} WHERE Emergency_Zone = ? AND Max_Depth > 0"
                cursor.execute(epz_PAR_u65_night_query, (epz,))
                epz_PAR_u65_night = cursor.fetchone()[0] or 0

                epz_PAR_o65_night_query = f"SELECT SUM(PAR_Over65_Mean) FROM {results1_night_input} WHERE Emergency_Zone = ? AND Max_Depth > 0"
                cursor.execute(epz_PAR_o65_night_query, (epz,))
                epz_PAR_o65_night = cursor.fetchone()[0] or 0
                epz_PAR_night = epz_PAR_u65_night + epz_PAR_o65_night

                epz_LL_day_query = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_day_input} WHERE Emergency_Zone = ?"
                cursor.execute(epz_LL_day_query, (epz,))
                epz_LL_day = cursor.fetchone()[0] or 0

                epz_LL_night_query = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_night_input} WHERE Emergency_Zone = ?"
                cursor.execute(epz_LL_night_query, (epz,))
                epz_LL_night = cursor.fetchone()[0] or 0

                # count structures flooded in the EPZ before the earliest warning
                epz_floodedbeforewarning_query = f"SELECT COUNT(*) FROM {results1_day_input} WHERE Emergency_Zone = ? AND Time_To_First_Wet < {epz_earliestwarning_minutes} AND Max_Depth > 0"
                cursor.execute(epz_floodedbeforewarning_query, (epz,))
                epz_floodedbeforewarning = cursor.fetchone()[0]
                if epz_floodedbeforewarning:
                    # sum life loss day of those structures
                    epz_floodedbeforewarning_LLday_query = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_day_input} WHERE Emergency_Zone = ? AND Time_To_First_Wet < {epz_earliestwarning_minutes} AND Max_Depth > 0"
                    cursor.execute(epz_floodedbeforewarning_LLday_query, (epz,))
                    epz_floodedbeforewarning_LLday_sum = cursor.fetchone()[0] or 0

                    epz_floodedbeforewarning_LLnight_query = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_night_input} WHERE Emergency_Zone = ? AND Time_To_First_Wet < {epz_earliestwarning_minutes} AND Max_Depth > 0"
                    cursor.execute(epz_floodedbeforewarning_LLnight_query, (epz,))
                    epz_floodedbeforewarning_LLnight_sum = cursor.fetchone()[0] or 0

                    add_message("... POSSIBLE ISSUE: {0} structures with Life Loss of {1} (day) and {2} (night) are flooded in the EPZ before the earliest possible warning".format(epz_floodedbeforewarning, round(epz_floodedbeforewarning_LLday_sum,2), round(epz_floodedbeforewarning_LLnight_sum,2)), lst=2)

                epz_posthazardarrival_day_query = f"SELECT COUNT(*) FROM {results1_day_input} WHERE Emergency_Zone = ? AND Time_To_First_Wet > ? AND Max_Depth > 0"
                cursor.execute(epz_posthazardarrival_day_query, (epz,BreachTime_in_minutes))
                epz_posthazardarrival_day = cursor.fetchone()[0]
                #add_message("...Structure count w/ post hazard arrival time and depth>0 (day): {}".format(epz_posthazardarrival_day))

                epz_prehazardarrival_day_query = f"SELECT COUNT(*) FROM {results1_day_input} WHERE Emergency_Zone = ? AND Time_To_First_Wet < ? AND Max_Depth > 0"
                cursor.execute(epz_prehazardarrival_day_query, (epz,BreachTime_in_minutes))
                epz_prehazardarrival_day = cursor.fetchone()[0]
                #add_message("...Structure count w/ pre hazard arrival time and depth>0 (day): {}".format(epz_prehazardarrival_day))

                # SQL query to sum mean life loss based on pre or post hazard arrival time
                epz_posthazardarrival_LLday_query = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_day_input} WHERE Emergency_Zone = ? AND Time_To_First_Wet > ? AND Max_Depth > 0"
                cursor.execute(epz_posthazardarrival_LLday_query, (epz,BreachTime_in_minutes))
                epz_posthazardarrival_LLday_sum = cursor.fetchone()[0] or 0
                #add_message("...LL in Structures with post hazard arrival (day): {}".format(epz_posthazardarrival_LLday_sum))

                epz_prehazardarrival_LLday_query = f"SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_day_input} WHERE Emergency_Zone = ? AND Time_To_First_Wet < ? AND Max_Depth > 0"
                cursor.execute(epz_prehazardarrival_LLday_query, (epz,BreachTime_in_minutes))
                epz_prehazardarrival_LLday_sum = cursor.fetchone()[0] or 0
                #add_message("...LL in Structures with pre hazard arrival (day): {}".format(epz_prehazardarrival_LLday_sum))

                #sql query to sum mean PAR based on pre or post hazard arrival time
                epz_posthazardarrival_PARu65day_query = f"SELECT SUM(Pop_Under65_Mean) FROM {results1_day_input} WHERE Emergency_Zone = ? AND Time_To_First_Wet > ? AND Max_Depth > 0"
                cursor.execute(epz_posthazardarrival_PARu65day_query, (epz,BreachTime_in_minutes))
                epz_posthazardarrival_PARu65day_sum = cursor.fetchone()[0] or 0
                epz_posthazardarrival_PARo65day_query = f"SELECT SUM(PAR_Over65_Mean) FROM {results1_day_input} WHERE Emergency_Zone = ? AND Time_To_First_Wet > ? AND Max_Depth > 0"
                cursor.execute(epz_posthazardarrival_PARo65day_query, (epz,BreachTime_in_minutes))
                epz_posthazardarrival_PARo65day_sum = cursor.fetchone()[0] or 0
                epz_posthazardarrival_PARday_sum = epz_posthazardarrival_PARu65day_sum + epz_posthazardarrival_PARo65day_sum

                epz_prehazardarrival_PARu65day_query = f"SELECT SUM(Pop_Under65_Mean) FROM {results1_day_input} WHERE Emergency_Zone = ? AND Time_To_First_Wet < ? AND Max_Depth > 0"
                cursor.execute(epz_prehazardarrival_PARu65day_query, (epz,BreachTime_in_minutes))
                epz_prehazardarrival_PARu65day_sum = cursor.fetchone()[0] or 0
                epz_prehazardarrival_PARo65day_query = f"SELECT SUM(PAR_Over65_Mean) FROM {results1_day_input} WHERE Emergency_Zone = ? AND Time_To_First_Wet < ? AND Max_Depth > 0"
                cursor.execute(epz_prehazardarrival_PARo65day_query, (epz,BreachTime_in_minutes))
                epz_prehazardarrival_PARo65day_sum = cursor.fetchone()[0] or 0
                epz_prehazardarrival_PARday_sum = epz_prehazardarrival_PARu65day_sum + epz_prehazardarrival_PARo65day_sum
                
                #query to calculate PAR_Warned_Mean day in each epz
                epz_PARwarnedday_query = f"SELECT SUM(PAR_Warned_Mean) FROM {results1_day_input} WHERE Emergency_Zone = ? AND Max_Depth > 0"
                cursor.execute(epz_PARwarnedday_query, (epz,))
                epz_PARwarnedday_sum = cursor.fetchone()[0] or 0
                epz_par_warned_day_percentage = (epz_PARwarnedday_sum / (epz_posthazardarrival_PARday_sum + epz_prehazardarrival_PARday_sum)) if (epz_posthazardarrival_PARday_sum + epz_prehazardarrival_PARday_sum) > 0 else 0

                #query to calculate PAR_Mobilized_Mean day in each epz
                epz_PARmobilizedday_query = f"SELECT SUM(PAR_Mobilized_Mean) FROM {results1_day_input} WHERE Emergency_Zone = ? AND Max_Depth > 0"
                cursor.execute(epz_PARmobilizedday_query, (epz,))
                epz_PARmobilizedday_sum = cursor.fetchone()[0] or 0
                epz_par_mobilized_day_percentage = (epz_PARmobilizedday_sum / (epz_posthazardarrival_PARday_sum + epz_prehazardarrival_PARday_sum)) if (epz_posthazardarrival_PARday_sum + epz_prehazardarrival_PARday_sum) > 0 else 0

                #query to calculate mean fatality rate in each epz
                epz_fatalityrate_day = (epz_posthazardarrival_LLday_sum + epz_prehazardarrival_LLday_sum) / (epz_posthazardarrival_PARday_sum + epz_prehazardarrival_PARday_sum) if (epz_posthazardarrival_PARday_sum + epz_prehazardarrival_PARday_sum) > 0 else 0

                #add_message("...Pre hazard arrival structure count / LL (day): {0} / {1}".format(epz_prehazardarrival_day, epz_prehazardarrival_LLday_sum), lst=2)
                #add_message("...Post hazard arrival structure count / LL (day): {0} / {1}".format(epz_posthazardarrival_day, epz_posthazardarrival_LLday_sum), lst=2)

                # excel Add relevant data starting from row 11
                ws.cell(row=current_row_epz, column=start_col_epz, value=epz)
                ws.cell(row=current_row_epz, column=start_col_epz+1, value=epz_structurecount_day)
                ws.cell(row=current_row_epz, column=start_col_epz+2, value=epz_PAR_day)
                ws.cell(row=current_row_epz, column=start_col_epz+3, value=epz_PAR_night)
                ws.cell(row=current_row_epz, column=start_col_epz+4, value=epz_LL_day or 0)
                ws.cell(row=current_row_epz, column=start_col_epz+5, value=epz_LL_night or 0)
                ws.cell(row=current_row_epz, column=start_col_epz+6, value=epz_prehazardarrival_day)
                ws.cell(row=current_row_epz, column=start_col_epz+7, value=epz_posthazardarrival_day)
                ws.cell(row=current_row_epz, column=start_col_epz+8, value=epz_prehazardarrival_PARday_sum)
                ws.cell(row=current_row_epz, column=start_col_epz+9, value=epz_posthazardarrival_PARday_sum)
                ws.cell(row=current_row_epz, column=start_col_epz+10, value=epz_prehazardarrival_LLday_sum or 0)
                ws.cell(row=current_row_epz, column=start_col_epz+11, value=epz_posthazardarrival_LLday_sum or 0)
                ws.cell(row=current_row_epz, column=start_col_epz+12, value=epz_par_warned_day_percentage)
                ws.cell(row=current_row_epz, column=start_col_epz+13, value=epz_par_mobilized_day_percentage)
                ws.cell(row=current_row_epz, column=start_col_epz+14, value=round(epz_fatalityrate_day, 3))

                # Apply number formatting to the relevant cells
                for col_epz in range(start_col_epz+1, start_col_epz+12):  # Adjust range as necessary
                    cell = ws.cell(row=current_row_epz, column=col_epz)
                    cell.number_format = '#,##0.0'  # Set the desired number format

                #set Percent Warned and percent mobilized to 1 decimal place %
                cell = ws.cell(row=current_row_epz, column=start_col_epz+12)
                cell.number_format = '0.0%'  # Set the desired number format    
                cell = ws.cell(row=current_row_epz, column=start_col_epz+13)
                cell.number_format = '0.0%'  # Set the desired number format

                # excel Move to the next row for data
                current_row_epz += 1

                # excel keep track of total rows
                current_row_total += 1

                ##END EPZ LOOP B

            ## Begin looking at summary areas in the summary results
            ## LOOP C loops through summary polygon sets
            for summarypolygonset in cleaned_list:
                # Split each pair into file and field
                areapolygon, areanamefield = summarypolygonset.split(", ")
                add_message("Looping through Summary Polygon File: {}...".format(areapolygon), lst=2)
                #add_message("Summary Polygon NameField: {}".format(areanamefield), lst=2)

                areatable = fr'"{simulation1}>Summary_Polygon_Set>{areapolygon}"'
                
                #get list of areas from the area table
                areaquery = f"SELECT {areanamefield} FROM {areatable}"
                cursor.execute(areaquery)
                    # Fetch all results
                rows = cursor.fetchall()
                    # Extract values from the rows and create a list
                area_list = [row[0] for row in rows]
                
                #Check to see if an Unassigned area needs to be added to the list
                area_unassigned_query = f'SELECT COUNT(*) FROM {results1_day_input} WHERE Max_Depth > 0 AND "Summary_Set_{areapolygon}" = \'Unassigned\''
                #add_message("area unassigned query: {}".format(area_unassigned_query), lst=2)

                cursor.execute(area_unassigned_query)
                area_unassigned = cursor.fetchone()[0]
                if area_unassigned > 0:  #Check if count is greater than 0
                    area_list.append('Unassigned')

                current_row_total += 1
                # excel For areas loop worksheet
                start_row_area = current_row_total
                start_col_area = 3

                # excel Define where the table should start for areas (e.g., row 10, column 2)
                row = start_row_area
                col = start_col_area
                # Set headers at a different row
                ws.cell(row=row, column=col, value="Area Name")
                ws.cell(row=row, column=col+1, value="Struc Total")
                ws.cell(row=row, column=col+2, value="PAR Day")
                ws.cell(row=row, column=col+3, value="PAR Night")
                ws.cell(row=row, column=col+4, value="LL Day")
                ws.cell(row=row, column=col+5, value="LL Night")
                ws.cell(row=row, column=col+6, value="Arrival Range")
                ws.cell(row=row, column=col+7, value="Depth Range")
                ws.cell(row=row, column=col+8, value="Velocity Range")
                ws.cell(row=row, column=col+9, value="DxV Range")
                ws.cell(row=row, column=col+10, value="Collapsed >50pct")
                ws.cell(row=row, column=col+11, value="Fatality Rate Day")
                ws.cell(row=row, column=col+12, value="% Warned Day")
                ws.cell(row=row, column=col+13, value="% Mobilized Day")
                ws.cell(row=row, column=col+14, value="Struc PostHaz")
                ws.cell(row=row, column=col+15, value="Avg Arrival PostHaz (hrs)")
                ws.cell(row=row, column=col+16, value="Avg Flood Depth (ft)")
                ws.cell(row=row, column=col+17, value="Arrival at first flooded PostHaz (hrs)")

                # excel Increment row for EPZ after the header
                current_row_area = start_row_area + 1
                current_row_total += 1

                #LOOP C-1 (subloop inside loop C)
                for area in area_list:
                    # SQL query to count features with depth greater than zero and arrival time pre and psot hazard the numeric field above the threshold
                    area_structurecount_day_query = f'SELECT COUNT(*) FROM {results1_day_input} WHERE Max_Depth > 0 AND "Summary_Set_{areapolygon}" = ?'
                    cursor.execute(area_structurecount_day_query, (area,))
                    area_structurecount_day = cursor.fetchone()[0] or 0

                    # SQL query to count features with depth greater than zero and arrival time pre and psot hazard the numeric field above the threshold
                    area_posthazstructurecount_day_query = f'SELECT COUNT(*) FROM {results1_day_input} WHERE Time_to_First_Wet > ? AND Max_Depth > 0 AND "Summary_Set_{areapolygon}" = ?'
                    cursor.execute(area_posthazstructurecount_day_query, (BreachTime_in_minutes, area))
                    area_posthazstructurecount_day = cursor.fetchone()[0] or 0
                    
                    # SQL query to calculate average post-hazard arrival time in each area, filters for any crazy numbers over 10million minutes
                    area_avgposthazardarrival_day_query = f'SELECT AVG(Time_To_First_Wet) FROM {results1_day_input} WHERE Time_To_First_Wet > ? AND Time_To_First_Wet < 10000000 AND Max_Depth > 0 AND "Summary_Set_{areapolygon}" = ?'
                    cursor.execute(area_avgposthazardarrival_day_query, (BreachTime_in_minutes, area))
                    area_avgposthazardarrival_day_minutes = cursor.fetchone()[0] or 0
                    #add_message("...raw PostHaz arrival (min) in: {0} = {1}".format(area, area_avgposthazardarrival_day_minutes), lst=2)
                    if area_avgposthazardarrival_day_minutes == 0:
                        area_avgposthazardarrival_day_hrs = 0
                    else:
                        area_avgposthazardarrival_day_hrs = round((area_avgposthazardarrival_day_minutes - BreachTime_in_minutes) / 60, 1)
                    #add_message("...raw PostHaz arrival (hrs) in: {0} = {1}".format(area, area_avgposthazardarrival_day_hrs), lst=2)
                    #add_message("...PostHaz arrival (hrs) in: {0} = {1}".format(area, round(area_avgposthazardarrival_day_hrs,2)), lst=2)

                    # SQL query to calculate the 15th percentile post-hazard arrival time in each area, filters for any crazy numbers over 10 million minutes
                    #the count times 0.15 identifies the 15th percentile row, change for other percentiles
                    area_percentile_posthazardarrival_day_query = f"""
                        SELECT Time_To_First_Wet FROM {results1_day_input} 
                        WHERE Time_To_First_Wet > ? AND Time_To_First_Wet < 10000000 
                        AND Max_Depth > 0 AND "Summary_Set_{areapolygon}" = ? 
                        ORDER BY Time_To_First_Wet
                        LIMIT 1 
                        OFFSET CAST((SELECT COUNT(*) FROM {results1_day_input} 
                        WHERE Time_To_First_Wet > ? AND Time_To_First_Wet < 10000000 
                        AND Max_Depth > 0 AND "Summary_Set_{areapolygon}" = ?) * ? AS INTEGER)
                        """

                    #cursor executes with a different percentile, like 0.15 or 0.85
                    cursor.execute(area_percentile_posthazardarrival_day_query, (BreachTime_in_minutes, area, BreachTime_in_minutes, area, arrivallowpercentile))
                    result15th = cursor.fetchone()
                    # Check if a result was returned, if it is none it is set to 0
                    area_15thpercentile_posthazardarrival_day_minutes = result15th[0] if result15th else 0
                    area_15thpercentile_posthazardarrival_day_hrs = round((area_15thpercentile_posthazardarrival_day_minutes - BreachTime_in_minutes) / 60, 1) if area_15thpercentile_posthazardarrival_day_minutes else 0
                    
                    cursor.execute(area_percentile_posthazardarrival_day_query, (BreachTime_in_minutes, area, BreachTime_in_minutes, area, arrivalhighpercentile))
                    result85th = cursor.fetchone()
                    # Check if a result was returned, if it is none it is set to 0
                    area_85thpercentile_posthazardarrival_day_minutes = result85th[0] if result85th else 0
                    area_85thpercentile_posthazardarrival_day_hrs = round((area_85thpercentile_posthazardarrival_day_minutes - BreachTime_in_minutes) / 60, 1) if area_85thpercentile_posthazardarrival_day_minutes else 0
                    
                    area_arrival_day_range = f"{area_15thpercentile_posthazardarrival_day_hrs} - {area_85thpercentile_posthazardarrival_day_hrs}"

                    #depth high and low percentiles
                    area_percentile_depth_query = f"""
                        SELECT Max_Depth FROM {results1_day_input} 
                        WHERE "Summary_Set_{areapolygon}" = ? 
                        AND Max_Depth > 0{epzfilter} 
                        ORDER BY Max_Depth
                        LIMIT 1 
                        OFFSET CAST((SELECT COUNT(*) FROM {results1_day_input} 
                        WHERE "Summary_Set_{areapolygon}" = ?  
                        AND Max_Depth > 0{epzfilter}) * ? AS INTEGER)
                        """

                    cursor.execute(area_percentile_depth_query, (area, area, depthlowpercentile))
                    area_depthlowpercent = cursor.fetchone()
                    area_depthlowpercent = round(area_depthlowpercent[0], 1) if area_depthlowpercent else 0
                    cursor.execute(area_percentile_depth_query, (area, area, depthhighpercentile))
                    area_depthhighpercent = cursor.fetchone()
                    area_depthhighpercent = round(area_depthhighpercent[0], 1) if area_depthhighpercent else 0
                    area_depth_range = f"{area_depthlowpercent} - {area_depthhighpercent}"
                    
                    #velocity high and low percentiles
                    area_percentile_velocity_query = f"""
                        SELECT Max_Velocity FROM {results1_day_input} 
                        WHERE "Summary_Set_{areapolygon}" = ? 
                        AND Max_Depth > 0{epzfilter} 
                        ORDER BY Max_Velocity
                        LIMIT 1 
                        OFFSET CAST((SELECT COUNT(*) FROM {results1_day_input} 
                        WHERE "Summary_Set_{areapolygon}" = ?  
                        AND Max_Depth > 0{epzfilter}) * ? AS INTEGER)
                        """

                    cursor.execute(area_percentile_velocity_query, (area, area, depthlowpercentile))
                    area_velocitylowpercent = cursor.fetchone()
                    area_velocitylowpercent = round(area_velocitylowpercent[0], 1) if area_velocitylowpercent else 0
                    cursor.execute(area_percentile_velocity_query, (area, area, depthhighpercentile))
                    area_velocityhighpercent = cursor.fetchone()
                    area_velocityhighpercent = round(area_velocityhighpercent[0], 1) if area_velocityhighpercent else 0
                    area_velocity_range = f"{area_velocitylowpercent} - {area_velocityhighpercent}"
                    
                    #DxV high and low percentiles
                    area_percentile_dxv_query = f"""
                        SELECT Max_DxV FROM {results1_day_input} 
                        WHERE "Summary_Set_{areapolygon}" = ? 
                        AND Max_Depth > 0{epzfilter} 
                        ORDER BY Max_DxV
                        LIMIT 1 
                        OFFSET CAST((SELECT COUNT(*) FROM {results1_day_input} 
                        WHERE "Summary_Set_{areapolygon}" = ?  
                        AND Max_Depth > 0{epzfilter}) * ? AS INTEGER)
                        """

                    cursor.execute(area_percentile_dxv_query, (area, area, velocitylowpercentile))
                    area_dxvlowpercent = cursor.fetchone()
                    area_dxvlowpercent = round(area_dxvlowpercent[0], 1) if area_dxvlowpercent else 0
                    cursor.execute(area_percentile_dxv_query, (area, area, velocityhighpercentile))
                    area_dxvhighpercent = cursor.fetchone()
                    area_dxvhighpercent = round(area_dxvhighpercent[0], 1) if area_dxvhighpercent else 0
                    area_dxv_range = f"{area_dxvlowpercent} - {area_dxvhighpercent}"

                    fiftypercentiterations = numiterations / 2
                    query_area_collapsed50p = f'SELECT COUNT(*) FROM {results1_day_input} WHERE Max_Depth > 0 AND "Summary_Set_{areapolygon}" = ? AND Collapsed > {fiftypercentiterations}'
                    cursor.execute(query_area_collapsed50p, (area,))
                    area_collapsed50p = cursor.fetchone()[0] or 0

                    # SQL query to get the first structure flooded after breach
                    area_firstflooded_posthazardarrival_day_query = f"""
                    SELECT MIN(Time_To_First_Wet) FROM {results1_day_input} 
                    WHERE Time_To_First_Wet > ? AND Time_To_First_Wet < 10000000 
                    AND Max_Depth > 0 AND "Summary_Set_{areapolygon}" = ?
                    """
                    cursor.execute(area_firstflooded_posthazardarrival_day_query, (BreachTime_in_minutes, area))
                    area_first_flooded_time = cursor.fetchone()
                    # Check if a result was returned, if it is none it is set to 0
                    area_first_flooded_posthazardarrival_day_minutes = area_first_flooded_time[0] if area_first_flooded_time else 0
                    area_first_flooded_posthazardarrival_day_hrs = round((area_first_flooded_posthazardarrival_day_minutes - BreachTime_in_minutes) / 60, 1) if area_first_flooded_posthazardarrival_day_minutes else 0

                    # SQL query to calculate average post-hazard depth in each area
                    area_avgdepth_day_query = f'SELECT AVG(Max_Depth) FROM {results1_day_input} WHERE Max_Depth > 0 AND "Summary_Set_{areapolygon}" = ?'
                    cursor.execute(area_avgdepth_day_query, (area,))
                    area_avgdepth_day = cursor.fetchone()[0] or 0
                    #add_message("...Avg Depth in: {0} = {1}".format(area, round(area_avgdepth_day,2)), lst=2)

                    # SQL query to calculate average day life loss in each area
                    area_avgLL_day_query = f'SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_day_input} WHERE "Summary_Set_{areapolygon}" = ?'
                    cursor.execute(area_avgLL_day_query, (area,))
                    area_avgLL_day = cursor.fetchone()[0] or 0
                    #add_message("...Avg LL (day) in: {0} = {1}".format(area, round(area_avgLL_day,2)), lst=2)

                    # SQL query to calculate average day PARu65 in each area
                    area_paru65_day_query = f'SELECT SUM(Pop_Under65_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND "Summary_Set_{areapolygon}" = ?'
                    cursor.execute(area_paru65_day_query, (area,))
                    area_paru65_day = cursor.fetchone()[0] or 0
                    #add_message("...Avg paru65 (day) in: {0} = {1}".format(area, round(area_paru65_day,2)), lst=2)

                    area_paro65_day_query = f'SELECT SUM(PAR_Over65_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND "Summary_Set_{areapolygon}" = ?'
                    cursor.execute(area_paro65_day_query, (area,))
                    area_paro65_day = cursor.fetchone()[0] or 0
                    #add_message("area PAR over65 (night): {}".format(area_paro65_night), lst=2)

                    area_par_day = round(area_paru65_day + area_paro65_day, 0)

                    # SQL query to calculate average night life loss in each area
                    area_avgLL_night_query = f'SELECT SUM(Life_Loss_Total_Mean) + SUM(Life_Loss_Evacuating_Mean) FROM {results1_night_input} WHERE "Summary_Set_{areapolygon}" = ?'
                    cursor.execute(area_avgLL_night_query, (area,))
                    area_avgLL_night = cursor.fetchone()[0] or 0

                    # SQL query to calculate average night PARu65 in each area
                    area_paru65_night_query = f'SELECT SUM(Pop_Under65_Mean) FROM {results1_night_input} WHERE Max_Depth > 0 AND "Summary_Set_{areapolygon}" = ?'
                    cursor.execute(area_paru65_night_query, (area,))
                    area_paru65_night = cursor.fetchone()[0] or 0
                    #add_message("...Avg paru65 (night) in: {0} = {1}".format(area, round(area_paru65_night,2)), lst=2)
                    area_paro65_night_query = f'SELECT SUM(PAR_Over65_Mean) FROM {results1_night_input} WHERE Max_Depth > 0 AND "Summary_Set_{areapolygon}" = ?'
                    cursor.execute(area_paro65_night_query, (area,))
                    area_paro65_night = cursor.fetchone()[0] or 0
                    #add_message("area PAR over65 (night): {}".format(area_paro65_night), lst=2)
                    area_par_night = round(area_paru65_night + area_paro65_night, 0)

                    # SQL query to calculate Fatality_Rate_Mean
                    #area_meanfatalityrate_query = f"SELECT AVG(Fatality_Rate_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND Summary_Set_{areapolygon} = '{area}'"
                    #cursor.execute(area_meanfatalityrate_query,)
                    #area_meanfatalityrate_day = cursor.fetchone()[0] or 0

                    area_meanfatalityrate_day = area_avgLL_day / area_par_day if area_par_day > 0 else 0

                    # SQL query to calculate PAR_Warned_Mean day
                    area_par_warned_day_query = f'SELECT SUM(PAR_Warned_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND "Summary_Set_{areapolygon}" = ?'
                    cursor.execute(area_par_warned_day_query, (area,))
                    area_par_warned_day = cursor.fetchone()[0] or 0
                    area_par_warned_day_percentage = (area_par_warned_day / area_par_day) if area_par_day > 0 else 0

                    # SQL query to calculate PAR_Mobilized_Mean day
                    area_par_mobilized_day_query = f'SELECT SUM(PAR_Mobilized_Mean) FROM {results1_day_input} WHERE Max_Depth > 0 AND "Summary_Set_{areapolygon}" = ?'
                    cursor.execute(area_par_mobilized_day_query, (area,))
                    area_par_mobilized_day = cursor.fetchone()[0] or 0
                    area_par_mobilized_day_percentage = (area_par_mobilized_day / area_par_day) if area_par_day > 0 else 0
                
                    # excel Add relevant data starting from row 11
                    ws.cell(row=current_row_area, column=start_col_area, value=area)
                    ws.cell(row=current_row_area, column=start_col_area+1, value=area_structurecount_day)
                    ws.cell(row=current_row_area, column=start_col_area+2, value=area_par_day)
                    ws.cell(row=current_row_area, column=start_col_area+3, value=area_par_night)
                    ws.cell(row=current_row_area, column=start_col_area+4, value=area_avgLL_day)
                    ws.cell(row=current_row_area, column=start_col_area+5, value=area_avgLL_night)
                    ws.cell(row=current_row_area, column=start_col_area+6, value=area_arrival_day_range)
                    ws.cell(row=current_row_area, column=start_col_area+7, value=area_depth_range)
                    ws.cell(row=current_row_area, column=start_col_area+8, value=area_velocity_range)
                    ws.cell(row=current_row_area, column=start_col_area+9, value=area_dxv_range)
                    ws.cell(row=current_row_area, column=start_col_area+10, value=area_collapsed50p)
                    ws.cell(row=current_row_area, column=start_col_area+11, value=area_meanfatalityrate_day)
                    ws.cell(row=current_row_area, column=start_col_area+12, value=area_par_warned_day_percentage)
                    ws.cell(row=current_row_area, column=start_col_area+13, value=area_par_mobilized_day_percentage)
                    ws.cell(row=current_row_area, column=start_col_area+14, value=area_posthazstructurecount_day)
                    ws.cell(row=current_row_area, column=start_col_area+15, value=area_avgposthazardarrival_day_hrs)
                    ws.cell(row=current_row_area, column=start_col_area+16, value=round(area_avgdepth_day, 1))
                    ws.cell(row=current_row_area, column=start_col_area+17, value=area_first_flooded_posthazardarrival_day_hrs)

                    # Apply number formatting to the relevant cells
                    for col_area in range(start_col_area+3, start_col_area+18):  # Adjust range as necessary, needs to be one greater than above table for some reason
                        cell = ws.cell(row=current_row_area, column=col_area)
                        cell.number_format = '#,##0.0'  # Set the desired number format

                    #set mean fatality to 3 decimal places
                    cell = ws.cell(row=current_row_area, column=start_col_area+11)
                    cell.number_format = '0.000'  # Set the desired number format

                    #set Percent Warned to 2 decimal places
                    cell = ws.cell(row=current_row_area, column=start_col_area+12)
                    cell.number_format = '0.0%'  # Set the desired number format
                    
                    #set Percent Mobilized to 2 decimal places
                    cell = ws.cell(row=current_row_area, column=start_col_area+13)
                    cell.number_format = '0.0%'  # Set the desired number format

                    # excel Move to the next row for data
                    current_row_area += 1
                    current_row_total += 1

                    ## End LOOP C-A area list loop
                
                ## End LOOP C area polygon loops

            # Begin exporting alternative structure summary results to output geopackage
            if export:
                add_message("Copying summary results GIS files into simulation results geopackage...", lst=2)
                results1_output_day=fr"{outputgpkg}\{sanitized_output_day}"
                results1_output_night=fr"{outputgpkg}\{sanitized_output_night}"
                try:                  
                    arcpy.management.CopyFeatures(
                        in_features=tablename_clean_day,
                        out_feature_class=results1_output_day,
                        )
                    arcpy.management.CopyFeatures(
                        in_features=tablename_clean_night,
                        out_feature_class=results1_output_night,
                        )
                    add_message("Features copied successfully with CopyFeatures.", lst=2)

                    #attempted logic to add and calculate total PAR and fatality rate fields, it doesn't error but for some reason all PAR ends up as 0
                    #since it doesn't work, calulcate fields is set to false for now. Also it takes a long time.
                    calculate_fields=False
                    if calculate_fields:
                        try:
                            add_message("Adding and calculating Total_PAR and Fatality_Rate fields using SQL.", lst=2)

                            #add_message(f"Fields added and calculated for table: {table}", lst=2)

                            add_message("Field calculations completed successfully.", lst=2)
                        except Exception as e:
                            arcpy.AddWarning(f"Field Calculation Error: {str(e)}")

                        # add_message("Adding and calculating Total_PAR and Fatality_Rate fields.", lst=2)
                        # # Add a new field for PAR
                        # total_par_gis_field = "Total_PAR"
                        # arcpy.management.AddField(
                        #     in_table=results1_output_day,
                        #     field_name=total_par_gis_field,
                        #     field_type="DOUBLE",  # Choose an appropriate data type
                        #     field_precision=None,
                        #     field_scale=None,
                        #     field_length=None,
                        #     field_alias=None,
                        #     field_is_nullable="NULLABLE",
                        #     field_is_required="NON_REQUIRED"
                        #     )
                        # arcpy.management.AddField(
                        #     in_table=results1_output_night,
                        #     field_name=total_par_gis_field,
                        #     field_type="DOUBLE",  # Choose an appropriate data type
                        #     field_precision=None,
                        #     field_scale=None,
                        #     field_length=None,
                        #     field_alias=None,
                        #     field_is_nullable="NULLABLE",
                        #     field_is_required="NON_REQUIRED"
                        #     )
                            
                        # # Calculate the PAR
                        # total_par_calc_expression = "!Pop_Under65_Mean! + !PAR_Over65_Mean!"
                        # arcpy.management.CalculateField(
                        #     in_table=results1_output_day,
                        #     field=total_par_gis_field,
                        #     expression=total_par_calc_expression,
                        #     expression_type="PYTHON3"
                        #     )
                        # arcpy.management.CalculateField(
                        #     in_table=results1_output_night,
                        #     field=total_par_gis_field,
                        #     expression=total_par_calc_expression,
                        #     expression_type="PYTHON3"
                        #     )
                        # add_message("PAR field added and calculated.", lst=2)
                        # # Add a new field for fatality rate, field name is established in individual risk export
                        # try:
                        #     arcpy.management.AddField(
                        #         in_table=results1_output_day,
                        #         field_name=fatality_rate_gis_field,
                        #         field_type="DOUBLE",  # Choose an appropriate data type
                        #         field_precision=None,
                        #         field_scale=None,
                        #         field_length=None,
                        #         field_alias=None,
                        #         field_is_nullable="NULLABLE",
                        #         field_is_required="NON_REQUIRED"
                        #         )
                        #     arcpy.management.AddField(
                        #         in_table=results1_output_night,
                        #         field_name=fatality_rate_gis_field,
                        #         field_type="DOUBLE",  # Choose an appropriate data type
                        #         field_precision=None,
                        #         field_scale=None,
                        #         field_length=None,
                        #         field_alias=None,
                        #         field_is_nullable="NULLABLE",
                        #         field_is_required="NON_REQUIRED"
                        #         )
                                
                        #     # Calculate the fatality rate
                        #     fatality_calc_expression2 = "(!Life_Loss_Total_Mean! / !Total_PAR!) if !Total_PAR! > 0 else 0"                       
                    
                        #     arcpy.management.CalculateField(
                        #         in_table=results1_output_day,
                        #         field=total_par_gis_field,
                        #         expression=fatality_calc_expression2,
                        #         expression_type="PYTHON3"
                        #         )
                        #     arcpy.management.CalculateField(
                        #         in_table=results1_output_night,
                        #         field=total_par_gis_field,
                        #         expression=fatality_calc_expression2,
                        #         expression_type="PYTHON3"
                        #         )
                        #     add_message("Fatality rate field added and calculated.", lst=2)
                        #except Exception as e:
                            #arcpy.AddWarning(f"Fatality rate calculation broke, skipping ahead. Error: {str(e)}")
                    ##END OF CALCULATE FIELDS

                except Exception as e:
                    arcpy.AddWarning(f"Copy features failed. Error: {str(e)}")
                ##END of GIS Export

            # Define the orange font style
            orange_font = Font(color="FFA500")  # FFA500 is the hex code for orange
            red_font = Font(color="C80000")  # 8B0000 is the hex code for red
            
            # Write the messages to the alternative Excel sheet
            for idx, message in enumerate(message_list2, start=1):
                idxcell = ws.cell(row=idx, column=1, value=message)
            
                # Ensure we remove extra spaces and check for "POSSIBLE ISSUE" in any case (case-insensitive)
                if "POSSIBLE ISSUE" in message.upper().strip():
                    idxcell.font = orange_font  # Apply the orange font
                    add_message("See POSSIBLE ISSUE in Alt {0}: {1}".format(alternative1, message), lst=1)
                
                # Ensure we remove extra spaces and check for "MMC SOP Violation" in any case (case-insensitive)
                if "MMC SOP VIOLATION" in message.upper().strip():
                    idxcell.font = red_font  # Apply the red font
                    add_message("See MMC SOP VIOLATION in Alt {0}: {1}".format(alternative1, message), lst=1)

        

            #clear loop message list (not necessary, it reinitializes at the start of the loop)
            #message_list2.clear

            #End LOOP A - alternative loop

        arcpy.ResetProgressor() #reset progressor to default

        ### BEGIN structure inventory check
        #add_message("Checking structure inventory for all alternatives...", lst=1)

        def summarize_inventory_fields(cursor, inventory, fields):
            summary = {}
            # Quote the table name to handle special characters
            inventory_quoted = f'"{inventory}"'
            for field in fields:
                summary[field] = {}
                for agg in ['SUM', 'MIN', 'MAX', 'AVG']:
                    query = f"SELECT {agg}({field}) FROM {inventory_quoted}"
                    cursor.execute(query)
                    value = cursor.fetchone()[0]
                    summary[field][agg] = value if value is not None else 0
            return summary
        
        # list of fields to summarize
        fields_to_summarize = [
            "DayU65Population",
            "DayO65Population",
            "NightU65Population",
            "NightO65Population",
            "Value_Structure",
            "Value_Content",
            "Value_Vehicle",
            "Stories_Number",
            "Height_Foundation",
            # Add more fields as needed
        ]

        for inventory in inventorylist:
            add_message("Checking inventory: {0}".format(inventory), lst=1)
            inventory_summary = summarize_inventory_fields(cursor, inventory, fields_to_summarize)
            # Example: print or log the results
            for field, aggs in inventory_summary.items():
                add_message(f"{inventory} - {field}: " +
                    ", ".join([
                        f"{agg}={f'{aggs[agg]:,.2f}' if isinstance(aggs[agg], (int, float)) else aggs[agg]}"
                        for agg in aggs
                    ]),
                    lst=1
                )

            # Check if any SUMs are equal to each other for this inventory
            sums = {field: aggs['SUM'] for field, aggs in inventory_summary.items()}
            checked_pairs = set()
            for field1, sum1 in sums.items():
                for field2, sum2 in sums.items():
                    if field1 != field2 and (field2, field1) not in checked_pairs:
                        if sum1 == sum2:
                            add_message(f"POSSIBLE ISSUE: SUM of '{field1}' equals SUM of '{field2}' in inventory '{inventory}' ({sum1})", lst=1)
                        checked_pairs.add((field1, field2))

        ### END structure inventory check

        # Close the database connections
        cursor.close()
        con.close()

        #if there are no messages in message_list1 with "MMC SOP Violation", add message stating "MMC SOP Check Cleared Successfully"
        if mmc_sop:
            if not any("MMC SOP Violation" in message for message in message_list1):
                add_message("MMC SOP Check Cleared Successfully", lst=1)
            else:
                add_message("MMC SOP Violations found, please review messages above", lst=1)
        #if there are no messages in message_list1 with "POSSIBLE ISSUE", add message stating "No Issues Found"
        if not any("POSSIBLE ISSUE" in message for message in message_list1):
            add_message("No Potential Issues Found", lst=1)

        add_message("Terrains used in simulation: {}".format(terrainlist), lst=1)
        if len(terrainlist) > 1:
            add_message("POSSIBLE ISSUE... the same terrain was NOT used for all alternatives", lst=1)
        
        add_message(f"Excel workbook created at: {output_excel_file}", lst=1)

        #Calculate compute time
        date_end = datetime.now()
        compute_time = date_end - date
        # Custom formatting for timedelta
        hours, remainder = divmod(compute_time.total_seconds(), 3600)
        minutes, seconds = divmod(remainder, 60)
        formatted_compute_time = f"{int(hours):02}:{int(minutes):02}:{int(seconds):02}"
        add_message("Compute Time: {}".format(formatted_compute_time), lst=1)

        green_font = Font(color="008000")

        # Write the summary messages for list 1 to the Excel summary sheet
        ws = wb["Summary"]
        for idx1, message1 in enumerate(message_list1, start=1):
            cell1 = ws.cell(row=idx1, column=1, value=message1)
            if "MMC SOP Violation" in message1:
                cell1.font = red_font
            if "POSSIBLE ISSUE" in message1:
                cell1.font = orange_font
            if "MMC SOP Check Cleared" in message1:
                cell1.font = green_font
        # After writing messages, append the alternative -> sheet name mapping at the bottom of the summary
        mapping_start_row = len(message_list1) + 2
        if alt_to_sheet_map:
            ws.cell(row=mapping_start_row, column=1, value="Alternative Name")
            ws.cell(row=mapping_start_row, column=2, value="Sanitized Sheet Name")
            for i, (altname, sheetname) in enumerate(alt_to_sheet_map, start=1):
                ws.cell(row=mapping_start_row + i, column=1, value=altname)
                ws.cell(row=mapping_start_row + i, column=2, value=sheetname)
        # Save the workbook in the output folder
        wb.save(output_excel_file)

        add_message(f"Excel workbook saved, all done!", lst=1)

        return



##-------------------------------------------------------------------------------------

