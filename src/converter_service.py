import os
from src.converter_worker import ExcelConverter

class Process:
    def __init__(self, output_folder, progress_callback = None) -> None:
        self.output_folder = output_folder + "/converted_output"
        self.progress_callback = progress_callback
        os.makedirs(self.output_folder, exist_ok=True)

    def process_single_file(self, file_path, log_callback = None):
        """
        Processes a single Excel file and converts it to GeoJSON and Shapefile.
        """
        if log_callback:
            log_callback(f"Processing file: {file_path}")

        converter = ExcelConverter(self.output_folder, log_callback, self.progress_callback)

        converter.load_excel_file(file_path)
        converter.clean_dataframes(file_path)

        converter.convert_to_geojson(self.output_folder)
        converter.convert_to_shapefile(self.output_folder)

        if log_callback:
            log_callback(f"Finished processing {file_path}")

    def process_folder(self, input_folder, log_callback = None):
        """
        Processes all Excel files in a folder.
        """
        files = [f for f in os.listdir(input_folder) if f.endswith((".xlsx", ".xls"))]
        total_files = len(files)

        for idx, file_name in enumerate(files, 1):
            file_path = os.path.join(input_folder, file_name)
            if log_callback:
                log_callback(f"Processing file: {file_name} ({idx}/{total_files})")
            converter = ExcelConverter(self.output_folder, log_callback, lambda p: self.progress_callback(int(((idx - 1 + p/100) / total_files) * 100)))
            converter.load_excel_file(file_path)
            converter.clean_dataframes(file_path)
            converter.convert_to_geojson(self.output_folder)
            converter.convert_to_shapefile(self.output_folder)

        if log_callback:
            log_callback("Batch processing completed.")

