import os
import win32com.client

def convert_pptx_to_pdf(input_folder, output_folder):
    # Ensure output folder exists
    os.makedirs(output_folder, exist_ok=True)

    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = 1

    for filename in os.listdir(input_folder):
        if filename.lower().endswith(".pptx"):
            input_path = os.path.abspath(os.path.join(input_folder, filename))
            output_path = os.path.abspath(
                os.path.join(output_folder, filename.replace(".pptx", ".pdf"))
            )

            try:
                presentation = powerpoint.Presentations.Open(input_path, WithWindow=False)
                presentation.SaveAs(output_path, 32)  # 32 = PDF format
                presentation.Close()
                print(f"✅ Converted: {filename}")
            except Exception as e:
                print(f"❌ Error converting {filename}: {e}")

    powerpoint.Quit()


if __name__ == "__main__":
    input_folder = r"C:\path\to\your\pptx_files"
    output_folder = r"C:\path\to\save\pdfs"

    convert_pptx_to_pdf(input_folder, output_folder)