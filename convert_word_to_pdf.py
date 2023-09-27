import os
import sys
import win32com.client

def word_to_pdf(input_docx, output_pdf):
    try:
        # Create a word application instance
        word = win32com.client.Dispatch("Word.Application")

        # Open the word document
        doc = word.Documents.Open(input_docx)

        # Get the base name of the input file (excluding the extension)
        input_base_name = os.path.splitext(os.path.basename(input_docx))[0]
        
        # Save the document as a PDF
        doc.SaveAs(output_pdf, FileFormat=17)

        # Close the word document
        doc.Close()

        # Quit the word application
        word.Quit()

        print(f"Conversion successful: {input_base_name}.docx -> {input_base_name}.pdf")

    except Exception as e:
        print(f"Error: {e}")

def convert_word_to_pdf(input_folder, output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for filename in os.listdir(input_folder):
        if filename.endswith(".docx"):
            input_docx = os.path.join(input_folder, filename)
            output_pdf = os.path.join(output_folder, os.path.splitext(filename)[0] + ".pdf")

            # Check if the PDF already exists inside the output folder, and if so, overwrite it
            if os.path.exists(output_pdf):
                os.remove(output_pdf)

            word_to_pdf(input_docx, output_pdf)

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python batch_to_convert_word_to_pdf.py input_folder output_folder")
    else:
        print("Starting conversion:")
        input_folder = sys.argv[1]
        output_folder = sys.argv[2]
        convert_word_to_pdf(input_folder, output_folder)
        print("Finished converting all files.")