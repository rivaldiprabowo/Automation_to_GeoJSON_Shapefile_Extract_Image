import os
import pandas as pd
import geopandas as gpd
import numpy as np
import openpyxl
from openpyxl_image_loader import SheetImageLoader
import json

class ExcelConverter:
    def __init__(self, output_folder="converted_output") -> None:
        self.output_folder = output_folder
        os.makedirs(self.output_folder, exist_ok=True)
        self.list_df = {}

    def load_excel_file(self, file_path):
        """
        Loads the Excel file and extracts each sheet into each DataFrames.
        """
        xl = pd.ExcelFile(file_path)

        for sheet in xl.sheet_names:
            df_temp = xl.parse(sheet, header=None)
            table_name = df_temp.iloc[0].dropna().values[0]

            df = xl.parse(sheet, header=[2, 3, 4])
            df.columns = ["_".join([str(c) for c in col if "Unnamed" not in str(c)]).strip() for col in df.columns.values]
            df = df.loc[:, ~df.columns.str.contains("Rekap", case=False, na=False)]

            df["DETAIL LOKASI"] = df["DETAIL LOKASI"].fillna(table_name)

            self.list_df[sheet] = df

    def clean_dataframes(self, file_path):
        """
        Cleans and preprocessed the DataFrames.
        """
        for sheet, df in self.list_df.items():
            # from this on df below didn't recognized as dataframes, why?
            binary_columns = [col for col in df.columns if col.startswith(("JENIS RAMBU", "LOKASI PEMASANGAN"))]
            df[binary_columns] = df[binary_columns].fillna("No").replace({1.0: "Yes"})

            if "DOKUMENTASI" not in df.columns:
                df["DOKUMENTASI"] = ""
            df["DOKUMENTASI"] = df["DOKUMENTASI"].astype(str)

            self.extract_images(file_path, df, sheet)

            # Drop completely empty columns (after extracting images)
            self.list_df[sheet] = df.dropna(how="all", subset=df.columns[1:])

    def extract_images(self, file_path, df, sheet_name):
        """
        Extracts images from the Excel file and updates the DataFrame.
        """
        pxl_doc = openpyxl.load_workbook(file_path, data_only=True)
        sheet = pxl_doc[sheet_name]  
        image_loader = SheetImageLoader(sheet)  

        output_folder = os.path.join(self.output_folder, "extracted_images")
        os.makedirs(output_folder, exist_ok=True)

        dokumentasi_column = "C"  # Replace with actual column letter
        for row in range(6, sheet.max_row + 1):
            cell_address = f"{dokumentasi_column}{row}"
            image_name = f"{sheet_name}_{row}.jpg"
            image_path = os.path.join(output_folder, image_name)

            # ✅ Skip if image already exists
            if os.path.exists(image_path):
                self.list_df[sheet_name].loc[row - 6, "DOKUMENTASI"] = image_path
                continue

            if image_loader.image_in(cell_address):
                try:
                    image = image_loader.get(cell_address)

                    # Ensure it's a valid image
                    if image:
                        image = image.convert("RGB")
                        image.save(image_path)

                        # Store the image path in the DataFrame
                        self.list_df[sheet_name].loc[row - 6, "DOKUMENTASI"] = image_path  
                except Exception as e:
                    print(f"⚠️ Warning: Could not process image at {cell_address} in sheet {sheet_name}. Error: {e}")
                    self.list_df[sheet_name].loc[row - 6, "DOKUMENTASI"] = "Image extraction failed"

        # Close the workbook **after** all processing is done
        pxl_doc.close()

    
    def convert_to_geojson(self, output_path):
        """
        Converts each cleaned DataFrame to a GeoJSON format.
        """
        for table_name, df in self.list_df.items():
            features = []
            for _, row in df.iterrows():
                properties = row.drop(["TITIK KORDINAT_Latitude", "TITIK KORDINAT_Longitude"], errors="ignore").to_dict()
                feature = {
                    "type": "Feature",
                    "geometry": {
                        "type": "Point",
                        "coordinates": [row["TITIK KORDINAT_Longitude"], row["TITIK KORDINAT_Latitude"]]
                    },
                    "properties": properties
                }
                features.append(feature)

            geojson_data = {"type": "FeatureCollection", "features": features}
            file_name = f"{table_name}.geojson"
            geojson_path = os.path.join(output_path, file_name)

            with open(geojson_path, "w", encoding="utf-8") as f:
                json.dump(geojson_data, f, ensure_ascii=False, indent=4)

    def convert_to_shapefile(self, output_path):
        """
        Converts DataFrames to Shapefile format using GeoPandas.
        """
        for table_name, df in self.list_df.items():
            if "TITIK KORDINAT_Latitude" in df.columns and "TITIK KORDINAT_Longitude" in df.columns:
                gdf = gpd.GeoDataFrame(
                    df,
                    geometry=gpd.points_from_xy(df["TITIK KORDINAT_Longitude"], df["TITIK KORDINAT_Latitude"]),
                    crs="EPSG:4326"
                )

                shapefile_folder = os.path.join(output_path, f"{table_name}_shapefile")
                os.makedirs(shapefile_folder, exist_ok=True)
                shapefile_path = os.path.join(shapefile_folder, table_name)

                gdf.to_file(shapefile_path, driver="ESRI Shapefile")

    def process_single_file(self, file_path):
        """
        Processes a single Excel file and converts it to GeoJSON and Shapefile.
        """
        print(f"Processing file: {file_path}")
        self.load_excel_file(file_path)
        self.clean_dataframes(file_path)

        output_path = self.output_folder
        self.convert_to_geojson(output_path)
        self.convert_to_shapefile(output_path)

        print(f"Conversion completed for {file_path}")

    def process_folder(self, input_folder):
        """
        Processes all Excel files in a folder.
        """
        print(f"Processing folder: {input_folder}")
        for file_name in os.listdir(input_folder):
            if file_name.endswith(".xlsx") or file_name.endswith(".xls"):
                file_path = os.path.join(input_folder, file_name)
                self.process_single_file(file_path)

        print("Batch processing completed.")

# Example Usage:
converter = ExcelConverter()
converter.process_single_file("Data/eksisting Jalan Siliwangi.xlsx")
# converter.process_folder("Data/Excel_Files")

