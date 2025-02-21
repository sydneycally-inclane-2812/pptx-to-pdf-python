import os
import win32com.client

def convert_pptx_to_pdf(folder_path):
    if not os.path.exists(folder_path):
        print(f"Error: The folder '{folder_path}' does not exist.")
        return

    pptx_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pptx')]
    if not pptx_files:
        print(f"No .pptx files found in the folder '{folder_path}'.")
        return

    try:
        # Initialize PowerPoint application via win32com.client
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = True  # Optional, set to False to run in the background
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
            if presentation is None:
                print(f"Failed to open '{pptx_file}'. Skipping this file.")
                continue

            # Save the presentation as PDF
            presentation.SaveAs(output_path, 32)  # 32 = ppSaveAsPDF
            presentation.Close()

            print(f"'{pptx_file}' converted successfully.")
        except Exception as e:
            print(f"Failed to convert '{pptx_file}': {e}")
            # Ensure the presentation is closed if an error occurs
            try:
                if 'presentation' in locals() and presentation:
                    presentation.Close()
            except Exception as close_error:
                print(f"Error closing presentation '{pptx_file}': {close_error}")
        finally:
            # Ensure cleanup happens for each presentation
            if 'presentation' in locals() and presentation:
                try:
                    presentation.Close()
                except Exception as close_error:
                    print(f"Error closing presentation '{pptx_file}': {close_error}")

    try:
        powerpoint.Quit()  # Quit the PowerPoint application
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
