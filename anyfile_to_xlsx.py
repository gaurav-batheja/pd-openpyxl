import pandas as pd
import os

def convert_to_xlsx(input_file):
    # Determine the file extension
    file_extension = os.path.splitext(input_file)[1].lower()
    output_file = os.path.splitext(input_file)[0] + '.xlsx'

    # Read the input file based on its extension
    if file_extension == '.csv':
        df = pd.read_csv(input_file)
        df.to_excel(output_file, index=False, engine='openpyxl')

    elif file_extension in ['.xls', '.xlsx']:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            xls = pd.read_excel(input_file, sheet_name=None)
            for sheet_name, sheet_df in xls.items():
                sheet_df.to_excel(writer, index=False, sheet_name=sheet_name)

    elif file_extension == '.txt':
        df = pd.read_csv(input_file, delimiter='\t')  # Assuming tab-delimited text file
        df.to_excel(output_file, index=False, engine='openpyxl')
    else:
        raise ValueError(f"Unsupported file format: {file_extension}")
    

    print(f"File converted to {output_file}")

# Example usage
convert_to_xlsx('dataset_5_439_csv.csv')
convert_to_xlsx('dataset_5_439_xls.xls')
convert_to_xlsx('dataset_5_439_txt.txt')