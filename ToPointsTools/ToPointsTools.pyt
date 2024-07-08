import arcpy
import pandas as pd
import os
from arcpy.sa import *
from netCDF4 import Dataset
import xarray as xr
import pandas as pd
from openpyxl import Workbook



class Toolbox(object):
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the .pyt file)."""
        self.label = "Excel To Points Toolbox"
        self.alias = "excel_to_points"

        # List of tool classes associated with this toolbox
        self.tools = [ExcelToPointsTool, NetCDFToExcelTool]


class ExcelToPointsTool(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Excel To Points"
        self.description = "Convert specified Excel sheet coordinates to points in a workspace"
        self.canRunInBackground = False

    def getParameterInfo(self):
        """Define parameter definitions"""
        # First parameter: Excel workbook
        param0 = arcpy.Parameter(
            displayName="Excel Workbook",
            name="excel_workbook",
            datatype="DEFile",
            parameterType="Required",
            direction="Input"
        )
        param0.filter.list = ["xls", "xlsx"]

        # Second parameter: Multi-input parameter of the sheets
        param1 = arcpy.Parameter(
            displayName="Excel Sheets",
            name="excel_sheets",
            datatype="GPString",
            parameterType="Required",
            direction="Input",
            multiValue=True
        )

        # Longitude Field Name parameter
        param2 = arcpy.Parameter(
            displayName="Longitude Field Name",
            name="longitude_field",
            datatype="GPString",
            parameterType="Required",
            direction="Input"
        )

        # Latitude Field Name parameter
        param3 = arcpy.Parameter(
            displayName="Latitude Field Name",
            name="latitude_field",
            datatype="GPString",
            parameterType="Required",
            direction="Input"
        )

        # Depth Field Name parameter
        param4 = arcpy.Parameter(
            displayName="Depth Field Name",
            name="depth_field",
            datatype="GPString",
            parameterType="Required",
            direction="Input"
        )

        # Vertical Exaggeration parameter
        param5 = arcpy.Parameter(
            displayName="Vertical Exaggeration",
            name="user_input_vertical_exaggeration",
            datatype="GPLong",
            parameterType="Optional",
            direction="Input"
        )

        # Workspace parameter
        param6 = arcpy.Parameter(
            displayName="Workspace",
            name="workspace",
            datatype="DEWorkspace",
            parameterType="Required",
            direction="Input"
        )
        
        # Associated Raster parameter
        param7 = arcpy.Parameter(
            displayName = "Associated Raster",
            name = "display_raster",
            datatype = "DERasterDataset",
            parameterType = "Optional",
            direction = "Input"
        )
        
        # Invert Depth Values parameter
        param8 = arcpy.Parameter(
            displayName = "Invert Depth Values?",
            name = "depth_values",
            datatype = "GPBoolean",
            parameterType = "Optional",
            direction = "Input"
        )
        
        # Output TIF File parameter
        param9 = arcpy.Parameter(
            displayName="Output TIF File",
            name="output_tif",
            datatype="DERasterDataset",
            parameterType="Optional",
            direction="Output"
        )

        return [param0, param1, param2, param3, param4, param5, param6, param7, param8, param9]

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal validation is performed."""
        if parameters[0].altered and not parameters[1].altered:
            # Read the sheet names from the Excel workbook
            excel_path = parameters[0].valueAsText
            if excel_path:
                sheet_names = pd.ExcelFile(excel_path).sheet_names
                parameters[1].filter.list = sheet_names
                parameters[1].value = sheet_names

        # Read the column names from the first sheet to populate the field name parameters
        if parameters[1].altered:
            excel_path = parameters[0].valueAsText
            selected_sheets = parameters[1].values
            if excel_path and selected_sheets:
                first_sheet_name = selected_sheets[0]
                df = pd.read_excel(excel_path, sheet_name=first_sheet_name)
                column_names = df.columns.tolist()
                parameters[2].filter.list = column_names
                parameters[3].filter.list = column_names
                parameters[4].filter.list = column_names

        # Set the Output TIF parameter as required if Associated Raster is provided
        if parameters[7].value:
            parameters[9].parameterType = "Required"
        else:
            parameters[9].parameterType = "Optional"

        return
    
    def execute(self, parameters, messages):
        """The source code of the tool."""
        # Changes the Z value of how the point file will be drawn by a value that the user inputs.
        def edit_z_values(shapefile, exaggeration_factor):
            # Check if the input shapefile exists
            if not arcpy.Exists(shapefile):
                print(f"Error: Shapefile '{shapefile}' does not exist.")
                return

            # Check if the shapefile is a point feature class
            desc = arcpy.Describe(shapefile)
            if desc.shapeType != "Point":
                print(f"Error: '{shapefile}' is not a point feature class.")
                return

            # Open an update cursor to edit the shapefile
            with arcpy.da.UpdateCursor(shapefile, ['SHAPE@']) as cursor:
                for row in cursor:
                    point = row[0]
                    if point:
                        # Get the Z value of the point
                        z = point.firstPoint.Z
                        if z is not None:
                            # Multiply the Z value by the exaggeration factor
                            new_z = z * exaggeration_factor
                            # Create a new point with the updated Z value
                            new_point = arcpy.Point(point.firstPoint.X, point.firstPoint.Y, new_z)
                            row[0] = new_point
                            cursor.updateRow(row)
            arcpy.AddMessage("Complete")
            
        excel_path = parameters[0].valueAsText
        sheet_names = parameters[1].values
        longitude_field = parameters[2].valueAsText
        latitude_field = parameters[3].valueAsText
        depth_field = parameters[4].valueAsText
        vertical_exaggeration = parameters[5].value
        workspace = parameters[6].valueAsText
        input_tif = parameters[7].valueAsText
        invert_depth = parameters[8].value
        output_tif = parameters[9].valueAsText
        
        # Creating the default workspace for functions to work with.
        arcpy.env.workspace = workspace
        
        if vertical_exaggeration is None:
            messages.addMessage("No user input number provided. Proceeding with default logic.")
        else:
            messages.addMessage(f"User input number provided: {vertical_exaggeration}")
            
        if input_tif is None:
            arcpy.AddMessage("No TIF found")
        else:
            arcpy.AddMessage("TIF has been found!")
            # Perform the raster calculation
            input_raster = Raster(input_tif)
            result_raster = input_raster * vertical_exaggeration

            # Save the result to the output TIF file
            result_raster.save(output_tif)
            arcpy.AddMessage(f"Successfully multiplied {input_tif} by {vertical_exaggeration} and saved to {output_tif}")
            
        """
        For the selected Excel Workbook, this function will loop through all the selected sheets.
        For each sheet, it will invert the value of the depth field if desired.
        It will search for the fields of Long(DecDeg), Lat(DecDeg), and Depth(m) to be used with 
        the arcpy function XYTableToPoint.
        If the user has added a vertical exaggeration, it will then change the drawing location of the points
        to reflect the input of vertical exaggeration. 
        The name that will be given to each of the points is "Points_" followed by the name of the sheet and then "_VE_" and the amount
        of vertical exaggeration given to the points. 
        ArcGIS does not natively support '-' in their naming schemes, so for all '-' found in the sheet name they will be replaced with "_"
        These points will be stored in the desired workspace with final output name schemes like "Points_V15A_01_VE_5"
        When an error occurs during processing, "Something went wrong" message will display and will move onto the next sheet.
        """
        
        for sheet in sheet_names:
            df = pd.read_excel(excel_path, sheet_name=sheet)
            try:
                if invert_depth and depth_field in df.columns:
                    df[depth_field] = df[depth_field] * -1
                
                temp_csv = os.path.join(arcpy.env.scratchFolder, f"{sheet}.csv")
                df.to_csv(temp_csv, index=False)
                sheet = sheet.replace('-', '_')

                if vertical_exaggeration is None:
                    out_name = f"Points_{sheet}_VE_NONE"
                    arcpy.management.XYTableToPoint(temp_csv, out_name, longitude_field, latitude_field, depth_field)

                    messages.addMessage(f"Processed sheet: {sheet} new file {out_name}")
                else: 
                    out_name = f"Points_{sheet}_VE_{vertical_exaggeration}"
                    arcpy.management.XYTableToPoint(temp_csv, out_name, longitude_field, latitude_field, depth_field)

                    messages.addMessage(f"Processed sheet: {sheet} new file {out_name}")

                    new_points_path = f"{workspace}\\{out_name}"
                    arcpy.AddMessage(new_points_path)

                    edit_z_values(new_points_path, vertical_exaggeration)
            except:
                arcpy.AddMessage("Something went wrong")

        return        
        
class NetCDFToExcelTool(object):
    def __init__(self):
        self.label = "NetCDF to Excel"
        self.description = "Extracts data from a NetCDF file and outputs an Excel file with data along selected dimensions."
        self.canRunInBackground = False

    def getParameterInfo(self):
        params = []

        # Parameter 1: Input netCDF File
        param1 = arcpy.Parameter(
            displayName="Input netCDF File",
            name="in_netCDF_file",
            datatype="DEFile",
            parameterType="Required",
            direction="Input"
        )
        params.append(param1)

        # Parameter 2: Dimensions from netCDF File
        param2 = arcpy.Parameter(
            displayName="Select Dimensions",
            name="dimensions",
            datatype="String",
            parameterType="Required",
            direction="Input",
            multiValue=True
        )
        params.append(param2)

        # Parameter 3: Variables from the first dimension
        param3 = arcpy.Parameter(
            displayName="Select Variable",
            name="variable",
            datatype="String",
            parameterType="Required",
            direction="Input"
        )
        params.append(param3)

        # Parameter 4: Slices for Dimensions
        param4 = arcpy.Parameter(
            displayName="Slices for Dimensions (format: dim1,start1,end1;dim2,start2,end2;...)",
            name="slices",
            datatype="String",
            parameterType="Optional",
            direction="Input"
        )
        params.append(param4)

        # Parameter 5: Output Excel File Name
        param5 = arcpy.Parameter(
            displayName="Output Excel File Name",
            name="excel_file_name",
            datatype="GPString",
            parameterType="Required",
            direction="Output"
        )
        params.append(param5)

        # Parameter 6: Output Folder (DEFolder)
        param6 = arcpy.Parameter(
            displayName="Output Folder",
            name="out_folder",
            datatype="DEFolder",
            parameterType="Required",
            direction="Output"
        )
        params.append(param6)

        return params

    def isLicensed(self):
        return True

    def updateParameters(self, parameters):
        # Populate dimensions from the netCDF file
        if parameters[0].value and not parameters[0].hasBeenValidated:
            netcdf_file = parameters[0].valueAsText
            dimensions = self.getDimensions(netcdf_file)
            parameters[1].filter.list = dimensions
            parameters[1].value = None  # Reset value to force user to select

        # Populate variables from the first selected dimension
        if parameters[1].value and not parameters[1].hasBeenValidated:
            netcdf_file = parameters[0].valueAsText
            first_dimension = parameters[1].values[0]  # Assume first selected dimension
            variables = self.getVariables(netcdf_file, first_dimension)
            parameters[2].filter.list = variables
            parameters[2].value = None  # Reset value to force user to select

        return

    def updateMessages(self, parameters):
        return

    def execute(self, parameters, messages):
        netcdf_file = parameters[0].valueAsText
        dimensions = [str(dim) for dim in parameters[1].values]
        variable = parameters[2].valueAsText
        slices_input = parameters[3].valueAsText if parameters[3].valueAsText else ""
        excel_file_name = parameters[4].valueAsText
        output_folder = parameters[5].valueAsText
        
        arcpy.AddMessage(variable)
        # Ensure the output folder exists
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        # Ensure the Excel file name has a valid extension
        if not excel_file_name.lower().endswith('.xlsx'):
            excel_file_name += '.xlsx'

        try:
            # Load the NetCDF file using xarray
            ds = xr.open_dataset(netcdf_file)

            # Parse slices input and create slices for each dimension
            dim_slices = {dim: slice(None) for dim in dimensions}
            if slices_input:
                for slice_str in slices_input.split(';'):
                    dim, start, end = slice_str.split(',')
                    dim_slices[dim] = slice(int(start), int(end))

            # Select the data along specified dimensions and slices
            data_subset = ds.sel(**dim_slices)

            # Create a new Excel workbook
            wb = Workbook()
            wb.remove(wb.active)  # Remove the default sheet created by openpyxl

            # First dimension for sheet names (e.g., N_STATIONS)
            first_dim = dimensions[0]
            # Second dimension for data filling (e.g., N_SAMPLES)
            second_dim = dimensions[1]

            # Iterate over unique values in the first selected dimension
            for dim_value in data_subset[first_dim].values:
                # Filter data for the current first dimension value
                current_data_subset = data_subset.sel({first_dim: dim_value})

                # Convert the data to a DataFrame for Excel output
                df = current_data_subset.to_dataframe().reset_index()

                # Filter columns to include only those related to the second dimension
                columns_to_include = [col for col in df.columns]# if second_dim in ds[col].dims]
                df_filtered = df[columns_to_include]

                # Identify the columns that typically have more rows
                extra_columns = ['N_SAMPLES', 'Cruise', 'Station', 'Type', 'Longitude', 'Latitude', 'Bot. Depth']

                # Determine the second maximum valid row count among the other columns
                valid_row_counts = [len(df_filtered[col].dropna()) for col in df_filtered.columns]
                unique_counts = sorted(set(valid_row_counts), reverse=True)
                if len(unique_counts) > 1:
                    second_max_valid_rows = unique_counts[1]
                else:
                    second_max_valid_rows = unique_counts[0]
                    
                arcpy.AddMessage(second_max_valid_rows)

                # Truncate the excess rows in the specified extra columns
                for col in extra_columns:
                    if col in df.columns:
                        df = df.iloc[:second_max_valid_rows]

                # Create a new worksheet for the current station name
                long_name_for_sheet = data_subset[variable].values[0]
                arcpy.AddMessage(long_name_for_sheet)
                ws = wb.create_sheet(title=long_name_for_sheet)
                #ws = wb.create_sheet(title=str(dim_value))

                # Write the headers to the worksheet using the variables' long_name
                for col_idx, var_name in enumerate(df.columns, start=1):
                    long_name = ds[var_name].attrs.get('long_name', var_name)
                    ws.cell(row=1, column=col_idx, value=long_name)

                # Write data to the worksheet, focusing on the second dimension
                for row_idx, row in enumerate(df.itertuples(index=False), start=2):
                    for col_idx, value in enumerate(row, start=1):
                        ws.cell(row=row_idx, column=col_idx, value=value)

            # Save the Excel workbook
            excel_output_path = os.path.join(output_folder, excel_file_name)
            wb.save(excel_output_path)

            arcpy.AddMessage(f"Data successfully extracted and saved to {excel_output_path}")

        except Exception as e:
            arcpy.AddError(f"Error processing netCDF file: {str(e)}")


    def getDimensions(self, netcdf_file):
        try:
            dataset = xr.open_dataset(netcdf_file)
            dimensions = list(dataset.dims.keys())
            dataset.close()
            return dimensions
        except Exception as e:
            arcpy.AddError(f"Error reading dimensions from netCDF file: {str(e)}")
            return []

    def getVariables(self, netcdf_file, dimension):
        try:
            dataset = xr.open_dataset(netcdf_file)
            variables = [var for var in dataset.variables if dimension in dataset[var].dims]
            dataset.close()
            return variables
        except Exception as e:
            arcpy.AddError(f"Error reading variables from netCDF file: {str(e)}")
            return []
