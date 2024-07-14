import win32com.client
import os
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def extract_vba_code(file_path):
    try:
        logging.info(f"Opening Excel file: {file_path}")
        excel = win32com.client.Dispatch("Excel.Application")
        workbook = excel.Workbooks.Open(file_path)
        vba_project = workbook.VBProject
        vba_code = {}

        for component in vba_project.VBComponents:
            if component.Type == 1:  # vbext_ct_StdModule
                module = component.CodeModule
                vba_code[component.Name] = module.Lines(1, module.CountOfLines)

        workbook.Close()
        excel.Quit()
        logging.info(f"Successfully extracted VBA code from: {file_path}")
        return vba_code

    except Exception as e:
        logging.error(f"Failed to extract VBA code: {e}")
        raise

def save_vba_code(vba_code, output_dir):
    try:
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        for module_name, code in vba_code.items():
            with open(os.path.join(output_dir, f"{module_name}.vba"), "w") as f:
                f.write(code)
        logging.info(f"Successfully saved VBA code to: {output_dir}")

    except Exception as e:
        logging.error(f"Failed to save VBA code: {e}")
        raise

def generate_documentation(vba_code, output_file):
    try:
        with open(output_file, "w") as f:
            f.write("# VBA Macro Documentation\n\n")
            for module_name, code in vba_code.items():
                f.write(f"## Module: {module_name}\n\n")
                f.write("```\n")
                f.write(code)
                f.write("\n```\n\n")
        logging.info(f"Successfully generated documentation: {output_file}")

    except Exception as e:
        logging.error(f"Failed to generate documentation: {e}")
        raise

def transform_vba_code(vba_code):
    try:
        transformed_code = {}
        for module_name, code in vba_code.items():
            transformed_code[module_name] = "' Transformed VBA Code\n" + code
        logging.info("Successfully transformed VBA code")
        return transformed_code

    except Exception as e:
        logging.error(f"Failed to transform VBA code: {e}")
        raise

# Paths
file_path = "path_to_your_excel_file.xlsm"  # Update this path to your actual Excel file
extracted_vba_dir = "extracted_vba_code"
transformed_vba_dir = "transformed_vba_code"
output_file = "vba_documentation.md"

try:
    # Execution
    vba_code = extract_vba_code(file_path)
    save_vba_code(vba_code, extracted_vba_dir)
    generate_documentation(vba_code, output_file)

    transformed_code = transform_vba_code(vba_code)
    save_vba_code(transformed_code, transformed_vba_dir)
    logging.info("Script execution completed successfully")

except Exception as e:
    logging.error(f"Script execution failed: {e}")
