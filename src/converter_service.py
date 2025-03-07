import os

class Process:
    def __init__(self, output_folder="converted_output") -> None:
        self.output_folder = output_folder
        os.makedirs(self.output_folder, exist_ok=True)
        self.list_df = {}

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

