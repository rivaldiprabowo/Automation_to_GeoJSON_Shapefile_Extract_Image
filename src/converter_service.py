import os
from src.converter_worker import ExcelConverter

class Process:
    def __init__(self, output_folder, progress_callback=None) -> None:
        self.output_folder = output_folder
        self.progress_callback = progress_callback
        os.makedirs(self.output_folder, exist_ok=True)
    
    def process_single_file(self, file_path, log_callback=None, qml_folder=None, batas_wilayah_path=None):
        """
        Processes a single Excel file and converts it to GeoJSON, Shapefile, and Extracts Images.
        """
        if log_callback:
            log_callback(f"Processing file: {file_path}")
        
        converter = ExcelConverter(self.output_folder, log_callback, self.progress_callback)
        
        # Process for Shapefile output
        converter.process_single_excel_file_shapefile(
            file_path, 
            self.output_folder, 
            qml_folder=qml_folder, 
            batas_wilayah_path=batas_wilayah_path
        )
        
        # Extract images from Excel
        converter.process_single_excel_file_images(
            file_path, 
            self.output_folder
        )

        # Process for GeoJSON output
        converter.process_single_excel_file_geojson(
            file_path, 
            self.output_folder, 
            batas_wilayah_path=batas_wilayah_path
        )
        
        if log_callback:
            log_callback(f"Finished processing {file_path}")
    
    def process_folder(self, input_folder, log_callback=None, qml_folder=None, batas_wilayah_path=None):
        """
        Processes all Excel files in a folder and converts it to GeoJSON, Shapefile, and Extracts Images.
        """
        if log_callback:
            log_callback(f"Processing folder: {input_folder}")
        
        converter = ExcelConverter(self.output_folder, log_callback, self.progress_callback)
        
        # Process for Shapefile output
        converter.process_excel_folder_shapefile(
            input_folder, 
            self.output_folder, 
            qml_folder=qml_folder, 
            batas_wilayah_path=batas_wilayah_path
        )
        
        # Extract images from Excel files
        converter.process_excel_folder_images(
            input_folder, 
            self.output_folder
        )

        # Process for GeoJSON output
        converter.process_excel_folder_geojson(
            input_folder, 
            self.output_folder, 
            batas_wilayah_path=batas_wilayah_path
        )
        
        if log_callback:
            log_callback("Batch processing completed successfully")
