# LifeSim Supplementary ArcGIS Python Toolboxes

The LifeSim Supplementary ArcGIS Python Toolboxes have been developed by LifeSim users to assist in pre-processing and post-processing data from LifeSim. For more information on LifeSim, see https://github.com/USACE-RMC/LifeSim. The tools are comprised of two separate python toolbox files:

### [LifeSim Results Python Toolbox](#lifesim-results-python-toolbox)
### [LifeSim GIS Preprocessing Python Toolbox](#lifesim-gis-preprocessing-python-toolbox)



## Python Toolbox Installation

1. **Download the repository as a .pyt file**
    - Click on each file in the list above.
    - In the upper right corner, a dropdown allows you to download the raw file.

3. **Add the toolbox in ArcGIS Pro:**
   - Open the Catalog pane.
   - Right-click on **Toolboxes** > **Add Toolbox**.
   - Select the `.pyt` file from this repository.

## LifeSim GIS Preprocessing Python Toolbox - Tool List

**[Buffer and Simplify a Polygon](#buffer-and-simplify-a-polygon)**

**[Calculate Incremental LifeSim Results](#calculate-incremental-lifesim-results)**

**[Convert Raster to Polygon](#convert-raster-to-polygon)** 

**[Define Unknown .tif Projections in a Folder](#define-unknown-tif-projections-in-a-folder)** 

**[NSI Creator (from local shapefiles)](#nsi-creator-from-local-shapefiles)**  

**[NSI Creator (from network API)](#nsi-creator-from-network-api)**  

**[Polygons to Double Warning EPZ](#polygons-to-double-warning-epz)**  

**[Reclassify Grid into a Zone Polygon](#reclassify-grid-into-a-zone-polygon)**

**[Snap Structures to Exclusion Polygon](#snap-structures-to-exclusion-polygon)** 

## LifeSim GIS Preprocessing Python Toolbox - Tool Parameters

### Buffer and Simplify a Polygon
Buffers an input polygon by a user-specified distance and simplifies the result. Useful for creating simplified study area polygons.

| Parameter              | Explanation                                               | Data Type      |
|------------------------|----------------------------------------------------------|----------------|
| Input Polygon          | Polygon shapefile to buffer and simplify                 | Feature Layer  |
| Output Polygon         | Output polygon location and filename                     | Feature Layer  |
| Buffer Distance in Feet| Buffer distance (in feet)                                | Long           |
___
[^Back to top](#lifesim-supplementary-arcgis-python-toolboxes)
### Calculate Incremental LifeSim Results
Takes two scenarios exported as shapefiles from the LifeSim map window and calculates the incremental depths, arrival times, velocities, Population At Risk (PAR), life loss, and damages. Outputs layers for both scenarios and the incremental values into a new geopackage. Negative incremental values mean that the bigger scenario actually wasn't bigger for the specific structure.

| Parameter      | Explanation                  | Data Type      |
|----------------|-----------------------------|----------------|
| Higher Scenario Shapefile (exported from LS)   | Input shapefile of the higher/bigger scenario | Shapefile |
| Lower Scenario Shapefile (exported from LS)   | Input shapefile of the lower/smaller scenario | Shapefile |
| Output Geopackage Name, no extension   | Output location and name for the geopackage, the gpkg extension is added upon creation | File |
| Check to delete structures not flooded in either scenario from incremental results  | Checked by default, this deletes structures points from the incremental output that have 0 depth in both scenarios. | Boolean |
___
[^Back to top](#lifesim-supplementary-arcgis-python-toolboxes)
### Convert Raster to Polygon
Converts a raster in .tif or .vrt format to a single polygon shapefile. Used to turn depth grids into a polygon for further analysis or to get a non-fail boundary for use in creating a double warning Emergency Planning Zone (EPZ).

| Parameter              | Explanation                                               | Data Type      |
|------------------------|----------------------------------------------------------|----------------|
| Input Raster Dataset          | Input raster, depth or water surface elevation in .vrt or .tif                 | Raster  |
| Output Polygon         | Output polygon location and filename                     | Shapefile  |
| Check to simplify output polygon| Option to simplify polygon, not recommended                               | Boolean           |
___
[^Back to top](#lifesim-supplementary-arcgis-python-toolboxes)
### Define Unknown tif Projections in a Folder
When given a folder, script creates a list of all the raster files of specified type and defines their projection to a defined projection. Designed for a few unique cases of the Rapid Inundation Mapping program where large sets of depth grids were created without a defined projection in RAS Mapper.

| Parameter              | Explanation                                               | Data Type      |
|------------------------|----------------------------------------------------------|----------------|
| Workspace folder         | Folder containing all of the raster grid files. Can have subfolders.                 | Folder  |
| Raster or shp file with correct projection        | Some GIS file, raster or shapefile, that has the correct projection already defined.                     | File  |
| Raster File Types| Tool uses a filter to either find .vrt files or .tif files, not both. If you have tiled grids where one .vrt has multiple .tifs, you must use the .vrt option.                                | List           |
___
[^Back to top](#lifesim-supplementary-arcgis-python-toolboxes)
### NSI Creator (from local shapefiles)
Clips and merges county-level NSI shapefiles to a study area, indexes dollar values, and can create unique name fields for HEC-FIA. This tool requires a folder with all of the individual county NSI shapefiles that are needed for the study area, it DOES NOT access the online NSI API to get data. For that, use the NSI Creatr (from network api).

| Parameter                    | Explanation                                                                 | Data Type      |
|------------------------------|-----------------------------------------------------------------------------|----------------|
| Study Area                   | Input study area polygon shapefile. Inundation boundaries are not recommended for study area use. Instead, use the buffer and simplify a polygon tool or another method to avoid an overly complex study area.                                          | Feature Layer  |
| Output Folder                | Output folder for interim and final data                                    | Folder         |
| US Counties Shapefile        | US county boundaries shapefile with FIPS code field                         | Feature Layer  |
| Folder with Combined NSI Files| Folder with all necessary NSI county shapefiles                            | Folder         |
| Price Index from Base        | Decimal index for value adjustment (default: 1)                             | Double         |
| Output NSI Filename          | Output shapefile name (no extension)                                        | String         |
| Check to add name field for FIA| Adds descriptive St_Name field for HEC-FIA                                | Boolean        |
___
[^Back to top](#lifesim-supplementary-arcgis-python-toolboxes)
### NSI Creator (from network API)
Downloads and processes NSI data from the network API, clips each county json file to the study area, then merges them all together into one output shapefile and geopackage. Supports both public and internal USACE APIs (internal requires USACE network access). See the NSI Technical References for more info at https://www.hec.usace.army.mil/confluence/nsi.

| Parameter                    | Explanation                                                                 | Data Type      |
|------------------------------|-----------------------------------------------------------------------------|----------------|
| Study Area                   | Input study area polygon shapefile. Inundation boundaries are not recommended for study area use. Instead, use the buffer and simplify a polygon tool or another method to avoid an overly complex study area.                                         | Feature Layer  |
| Output Folder                | Output folder for interim and final data                                    | Folder         |
| US Counties Shapefile        | US county boundaries shapefile with FIPS code field                         | Feature Layer  |
| Price Index from Base        | Decimal index for dollar value adjustments. Value of 1 means no index. Current default is 1.177 which is the ENR CCI index from 2021 to 2025.                          | Double         |
| Output NSI Filename          | Output shapefile name (no extension)                                        | String         |
| Check to add name field for FIA| Adds descriptive St_Name field for HEC-FIA                                | Boolean        |
| Check to get USACE-only fields| Use internal API for extra fields (not releasable)                         | Boolean        |
| Check to split by study area polygons| This option is used if the study area is split into multiple subpolygons and an individual NSI file is needed for each subpolygon. Use case is for FEMA FFRD projects where there are model area sub-units                                | Boolean        |
| Check to remove points within 20 feet of the boundary| Removes points within 20 feet out the study area polygon boundary. Sometimes points very close to the boundary can cause errors in LifeSim that it is outside of the EPZ boundary due to very minor reprojection differences. Ensure that your study area polygon is a single polygon if this is used, and make sure it is buffered and not an exact inundation boundary.                                | Boolean        |
| Check to reproject output to standard| Reprojects the output to the spatial projection used by msot USACE hydraulic models (USGS Albers Equal Area Conic)                                | Boolean        |
___
[^Back to top](#lifesim-supplementary-arcgis-python-toolboxes)
### Polygons to Double Warning EPZ
Creates a double warning EPZ polygon with in-pool, non-breach, and remaining study area zones. Double warning EPZs are used in LifeSim breach analysis to allow any areas that would get flooded by predictable operation (such as spillway flow in a non-breach condition) to be given a warning in the model, while other areas that would only be flooded if a breach occurs can get a different warning relative to the breach or other hazard.

| Parameter                    | Explanation                                                                 | Data Type      |
|------------------------------|-----------------------------------------------------------------------------|----------------|
| Input Study Area Polygon     | Study area polygon shapefile                                                | Feature Layer  |
| Input In-Pool Area Polygon   | Shapefile for in-pool area                                                  | Feature Layer  |
| Input Non-Breach Polygon 1   | Non-breach scenario polygon                                                 | Feature Layer  |
| Output Folder                | Output folder                                                               | Folder         |
| Output Name for Scenario 1   | Output name for scenario 1 (no extension)                                   | String         |
| Input Non-Breach Polygon 2   | (Optional) Second non-breach scenario polygon                               | Feature Layer  |
| Output Name for Scenario 2   | (Optional) Output name for scenario 2 (no extension)                        | String         |
___
[^Back to top](#lifesim-supplementary-arcgis-python-toolboxes)
### Reclassify Grid into a Zone Polygon
Reclassifies a raster into value ranges and creates polygons for each range. This can be used to create arrival time zones (this is the intention of the default settings) or depth range zones as a polygon shapefile.

| Parameter                    | Explanation                                                                 | Data Type      |
|------------------------------|-----------------------------------------------------------------------------|----------------|
| Input Grid                   | Input `.tif` or `.vrt` file (e.g., arrival time grid)                       | Raster Dataset |
| Base or Zero Value for Zones | Value to subtract from grid for zone calculation                            | Double         |
| Max Value for 1st Range      | Upper bound for first range                                                 | Double         |
| Max Value for 2nd Range      | Upper bound for second range                                                | Double         |
| Max Value for 3rd Range      | Upper bound for third range                                                 | Double         |
| Max Value for 4th Range      | Upper bound for fourth range                                                | Double         |
| Max Value for 5th Range      | Upper bound for fifth range                                                 | Double         |
| Unit Name                    | Unit for zone names (e.g., hours, feet)                                     | String         |
| Output Folder                | Output folder                                                               | Folder         |
| Output Filename              | Output shapefile name (no extension)                                        | String         |
| Check to Simplify Polygons   | Option to simplify polygons (not recommended)                               | Boolean        |
___
[^Back to top](#lifesim-supplementary-arcgis-python-toolboxes)
### Snap Structures to Exclusion Polygon
This tool takes a point shapefile structure inventory, selects the points that are inside a given polygon boundary, and snaps them to a specified distance outside of the polygon. An example use case would be when there are houses on floating docks along a river, and to simplify a life loss calculation the points need to be moved to the bank so they are not flooded in the first time step of the simulation (in this case the exclusion polygon would be a normal flow scenario). 

| Parameter                    | Explanation                                                                 | Data Type      |
|------------------------------|-----------------------------------------------------------------------------|----------------|
| NSI Inventory Shapefile      | Input inventory as a point shapefile                                                | Feature Layer  |
| Exclusion Polygon            | Shapefile for in-pool area                                                  | Feature Layer  |
| Distance to move structures out of polygon (ft)   | Distance in feet that points will be relocated to outside of the exclusion zone                                                 | Feature Layer  |
| Output NSI File Name (no extension)               | Output location and file name                                                              | Shapefile        |
___
[^Back to top](#lifesim-supplementary-arcgis-python-toolboxes)
## LifeSim Results Python Toolbox
The LifeSim Results Python Toolbox uses SQL queries to read data from the LifeSim .fia file, summarizes the data in meaningful ways, and writes the summary outputs to an Excel file. 

| Parameter              | Explanation                                               | Data Type      |
|------------------------|----------------------------------------------------------|----------------|
| LifeSim File           | The .fia file saved by the LifeSim software              | Feature Layer  |
| Simulation             | Once the .fia file is selected, a list is populated of the simulations that have been ran in the model. <br>Select the simulation to analyze.                    | Dynamic List  |
| Arrival Time Field     | The output generates arrival time ranges based on either the Time to First Wet or the Time to No Evac (the default is 2 feet in LifeSim)                               | Static List           |
| Range Percentiles      | The output generates ranges of depth and arrival times for each summary output polygon. <br>These can be based on the 25th and 25th percentlie values, or on the 15th and 85th percentlie values.                               | Static List           |
| Output Excel File      | Autopopulates to the directory of the .fia file, using the LifeSim file name and simulation name. Can be changed or left as default.             | xlsx file  |
| Alternative (optional) | Autopopulates with a list of alternatives in the model. If one is selected, the script only runs on that alternative instead of all alternatives in the simulation. Note that it does not filter by simulation, so the selected alternative must have been run in the selected simulation, otherwise there will be an error.             | Dynamic List  |
| Summary Areas (optional)| Autopopulates with a list of summary area polygons included in the simulation run. If one is selected, the script only runs on that summary area set instead of all summary area sets in the simulation.             | Dynamic List  |
| Check to Flag MMC SOP Violations          | Selected by default. If selected, the script will flag possible violations where parameters such as warning times and curves are not set to the standard MMC SOP parameters. Violations will be in red text.              | Boolean  |
| Check to export alternative results to geopackage (experimental)          | Exports day and night alternative results into a new geopackage in the LifeSim directory for use in GIS software. Also exports the structure in each scenario with the highest fatality rate for identifying the drivers of individual risk (will have IR in the layer name). Process may error and needs work, also it takes longer. Recommend selecting an optional alternative if you only need this for one or two alternatives.            | Boolean  |
___
[^Back to top](#lifesim-supplementary-arcgis-python-toolboxes)
