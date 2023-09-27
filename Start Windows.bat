@echo off

:: Get the directory where the batch file is located
set "batch_dir=%~dp0"

:: Set the input folder as a subfolder of the batch file directory
set "input_folder=%batch_dir%Word_Documents"

:: Set the output folder where the PDFs will be saved
set "output_folder=%batch_dir%PDF_Output"

:: Ensure the output folder exists
if not exist "%output_folder%" (
    mkdir "%output_folder%"
)

:: Ensure the input folder exists or create it
if not exist "%input_folder%" (
    mkdir "%input_folder%"
)

:: Run the Python script to perform the conversion
python convert_word_to_pdf.py "%input_folder%" "%output_folder%"

:: Pause to see any error messages (optional)
pause
