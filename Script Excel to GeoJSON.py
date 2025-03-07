# %% [markdown]
# # Automation Excel to CSV and GeoJSON

# %% [markdown]
# ## 1. Import Library

# %%
import pandas as pd
from openpyxl import load_workbook
import geopandas as gpd
from shapely.geometry import Point
import json
import glob
import os

# %% [markdown]
# ## 2. Application to Export Excel into GeoJson

# %% [markdown]
# ### 2.1. Function Codes

# %%
def unique_column_names(columns): # Ensure column names are unique by appending suffix.
    seen = {}
    new_columns = []
    for col in columns:
        if col in seen:
            seen[col] += 1
            new_columns.append(f"{col}_{seen[col]}")
        else:
            seen[col] = 0
            new_columns.append(col)
    return new_columns

def clean_column_names(columns): #Standardize column names by capitalizing each word properly.
    cleaned_columns = []
    seen = {}
    for col in columns:
        col = str(col).strip()
        col = " ".join(word.capitalize() for word in col.split())
        if col in seen:
            seen[col] += 1
            col = f"{col} {seen[col]}"
        else:
            seen[col] = 0
        cleaned_columns.append(col)
    return cleaned_columns

def fix_coordinates(row, lat_col, lon_col): #Fix latitude and longitude values that may be in the wrong format.
    lat, lon = row[lat_col], row[lon_col]
    if pd.notna(lat) and abs(lat) > 90:
        lat /= 1_000_000
    if pd.notna(lon) and abs(lon) > 180:
        lon /= 1_000_000
    return pd.Series([lat, lon])

def clean_geojson(gdf, output_path): #Save GeoDataFrame in a clean format GeoJSON file.
    temp_path = output_path.replace(".geojson", "_temp.geojson")
    gdf.to_file(temp_path, driver="GeoJSON")

    with open(temp_path, "r", encoding="utf-8") as file:
        geojson_data = json.load(file)

    with open(output_path, "w", encoding="utf-8") as file:
        json.dump(geojson_data, file, indent=4)

    print(f"✅ Saved: {output_path}")

def flatten_excel_to_geojson(file_path, output_folder): 
#Convert all sheets from an Excel file to GeoJSON, ensuring clean column names, valid geometries, and clean GeoJSON format
    
    # Load workbook
    wb = load_workbook(file_path, data_only=True)

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]

        # Unmerge cells and fill values
        for merge in list(ws.merged_cells):
            ws.unmerge_cells(str(merge))
            top_left = ws.cell(merge.min_row, merge.min_col).value
            for row in range(merge.min_row, merge.max_row + 1):
                for col in range(merge.min_col, merge.max_col + 1):
                    ws.cell(row, col, top_left)

        # Convert to DataFrame
        data = list(ws.values)
        df = pd.DataFrame(data)

        # Identify the header row
        header_index = df[df.apply(lambda x: x.astype(str).str.contains("NO", case=False, na=False)).any(axis=1)].index[0]

        # Use row 3 (index 2) as the header
        df.columns = df.iloc[2].astype(str).str.strip()

        # Remove empty columns
        df = df.dropna(axis=1, how="all")

        # Drop "REKAP" section if present
        df = df.loc[:, ~df.columns.str.contains("REKAP", case=False, na=False)]
        df = df.drop(index=[0, 1, 4]).reset_index(drop=True)

        # Merge first two rows if needed
        merged_header = [a if a == b else f"{a} {b}" for a, b in zip(df.iloc[0], df.iloc[1])]

        # Ensure column names are unique
        df.columns = unique_column_names(merged_header)

        # Remove the first two rows used for headers
        df = df.drop(index=[0, 1]).reset_index(drop=True)

        # Normalize column names for consistent detection
        df.columns = df.columns.str.lower().str.strip()

        # Find Latitude & Longitude columns dynamically
        lat_col = next((col for col in df.columns if "latitude" in col or "lat" in col), None)
        lon_col = next((col for col in df.columns if "longitude" in col or "lon" in col), None)

        if not lat_col or not lon_col:
            print(f"⚠️ Skipping '{sheet_name}' (No Latitude/Longitude columns)")
            continue  # Skip this sheet if Lat/Lon are missing

        # Convert Lat/Lon to numeric first (forcing errors to NaN)
        df[lat_col] = pd.to_numeric(df[lat_col], errors='coerce')
        df[lon_col] = pd.to_numeric(df[lon_col], errors='coerce')

        # Apply the fix function
        df[[lat_col, lon_col]] = df.apply(fix_coordinates, axis=1, lat_col=lat_col, lon_col=lon_col)

        # Remove rows where Lat/Lon are still missing
        df = df.dropna(subset=[lat_col, lon_col]).reset_index(drop=True)

        # Ensure geometry column exists before creating GeoDataFrame
        df["geometry"] = df.apply(
            lambda row: Point(row[lon_col], row[lat_col]) if pd.notna(row[lon_col]) and pd.notna(row[lat_col]) else None,
            axis=1
        )

        # Drop rows where geometry is None
        df = df.dropna(subset=["geometry"]).reset_index(drop=True)

        # Only create GeoDataFrame if there are valid geometries
        if not df["geometry"].isnull().all():
            properties_cols = [col for col in df.columns if col not in [lat_col, lon_col, "geometry"]]
            gdf = gpd.GeoDataFrame(df[properties_cols + ["geometry"]], crs="EPSG:4326")
        else:
            print(f"⚠️ Skipping '{sheet_name}' (No valid geometry found)")
            continue  # Skip processing this sheet if no valid geometries exist

        # Apply column renaming after creating the GeoDataFrame
        gdf.columns = clean_column_names(gdf.columns)

        # Remove unwanted "None_" and "None" columns
        gdf = gdf.loc[:, ~gdf.columns.str.match(r"^None$|None_", na=False)]

        # Remove " None" from remaining column names
        gdf.columns = gdf.columns.str.replace(r"\sNone\b", "", regex=True).str.strip()

        # Create output folder if it doesn't exist
        os.makedirs(output_folder, exist_ok=True)

        # Define output file path
        output_path = os.path.join(output_folder, f"{sheet_name}.geojson")

        # Save the GeoJSON in a clean format
        clean_geojson(gdf, output_path)

    # Delete temporary files
    for temp_file in glob.glob(os.path.join(output_folder, "*_temp.geojson")):
        os.remove(temp_file)

    print("🎉 All sheets processed successfully!")

# %% [markdown]
# ### 2.2. Run Function

# %%
excel_file = r"C:\Users\kanzi\Documents\Part Time Job\Automation Codes\02. CILEUNGSI - CIBINONG (CITEUREUP).xlsx"  # Fill with the path file of excel
export_folder = r"C:\Users\kanzi\Documents\Part Time Job\Automation Codes\check2"  # Fill with the path folder of export result
flatten_excel_to_geojson(excel_file, export_folder) # Run the function!


