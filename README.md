# Covert-Word-To-PDFs
 Converts all word documents inside an inputed folder to pdf documents and outputs them into a folder

The code is ran with 2 simple functions named batch_convert_word_to_pdf and word_to_pdf respectively
The first function batch_conver_word_to_pdf checks for an existing output_folder and makes one if there is not
It will then loop through each file until there is a .docx file and send it to the word_to_pdf function

This second function is carrying out the application opening, saving as a pdf formatted file and closing the file
These will all be available in the output folder with the same name
