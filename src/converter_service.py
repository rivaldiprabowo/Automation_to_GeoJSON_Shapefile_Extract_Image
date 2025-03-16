import os
from src.converter_worker import ExcelConverter

class Process:
    # What happens if the output dir is from conversion thread to _init_?
    def __init__(self, output_folder) -> None:
        self.output_folder = output_folder + "/converted_output"
        os.makedirs(self.output_folder, exist_ok=True)
        self.list_df = {}

    # Need to add, output path
    def process_single_file(self, file_path, log_callback = None):
        """
        Processes a single Excel file and converts it to GeoJSON and Shapefile.
        """
        if log_callback:
            log_callback(f"Processing file: {file_path}")

        ExcelConverter.load_excel_file(file_path)
        ExcelConverter.clean_dataframes(file_path)

        ExcelConverter.convert_to_geojson(self.output_folder)
        ExcelConverter.convert_to_shapefile(self.output_folder)

        if log_callback:
            log_callback(f"Finished processing {file_path}")

    # Need to add, output path
    def process_folder(self, input_folder, log_callback = None):
        """
        Processes all Excel files in a folder.
        """
        if log_callback:
            log_callback(f"Processing folder: {input_folder}")

        for file_name in os.listdir(input_folder):
            if file_name.endswith(".xlsx") or file_name.endswith(".xls"):
                file_path = os.path.join(input_folder, file_name)
                self.process_single_file(file_path)

        if log_callback:
            log_callback("Batch processing completed.")

