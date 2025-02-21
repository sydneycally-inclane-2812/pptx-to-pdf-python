import os
import comtypes.client

def convert_pptx_to_pdf(folder_path):
    # Check if folder exists
    if not os.path.exists(folder_path):
        print(f"Error: The folder '{folder_path}' does not exist.")
        return

    # Get all .pptx files in the folder
    pptx_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pptx')]
    if not pptx_files:
        print(f"No .pptx files found in the folder '{folder_path}'.")
        return

    # Initialize PowerPoint application
    try:
        powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
        powerpoint.Visible = 1  # Make PowerPoint visible (optional)
    except Exception as e:
        print(f"Failed to initialize PowerPoint: {e}")
        return

    for pptx_file in pptx_files:
        try:
            input_path = os.path.join(folder_path, pptx_file)
            output_path = os.path.join(folder_path, os.path.splitext(pptx_file)[0] + ".pdf")

            print(f"Converting '{pptx_file}' to PDF...")

            # Open the presentation
            presentation = powerpoint.Presentations.Open(input_path)
            # Save as PDF
            presentation.SaveAs(output_path, 32)  # 32 = ppSaveAsPDF
            presentation.Close()

            print(f"'{pptx_file}' converted successfully.")
        except Exception as e:
            print(f"Failed to convert '{pptx_file}': {e}")
        finally:
            # Ensure the presentation is closed
            if 'presentation' in locals() and presentation:
                presentation.Close()

    # Quit PowerPoint application
    try:
        powerpoint.Quit()
        print("All files processed. PowerPoint application closed.")
    except Exception as e:
        print(f"Error closing PowerPoint: {e}")

if __name__ == "__main__":
    while True:
        folder_path = input("Enter the folder path containing .pptx files (or type 'exit' to quit): ")
        if folder_path.lower() == 'exit':
            print("Exiting the script. Goodbye!")
            break
        convert_pptx_to_pdf(folder_path)
