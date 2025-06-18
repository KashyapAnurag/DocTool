import os
from pdf_extractor.core import run_processing_batch

input_pdfs_folder =   "  path_to_your_pdf_folder"  # Replace with your actual folder path
output_excel_folder = "  path_to_excel_folder   "  # Replace with your actual output folder path

print(f"Starting PDF processing from '{input_pdfs_folder}' to '{output_excel_folder}'")

os.makedirs(output_excel_folder, exist_ok=True)

failed_files = run_processing_batch(input_pdfs_folder, output_excel_folder)

if failed_files:
    print("\n--- Processing Summary ---")
    print(f"Number of PDFs that failed to process: {len(failed_files)}")
    print("Failed files:")
    for f in failed_files:
        print(f"- {f}")
else:
    print("\nAll PDFs in the input folder processed successfully!")

print("Processing complete.")