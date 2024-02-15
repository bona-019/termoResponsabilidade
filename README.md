# termoResponsabilidade
Automatic fill of .docx file (templates)

The script reads an Excel file with all the information needed and automatic fill up the template file. The template file contains some tags (example: {NAME}), where name is filled up with information from the NAME column in the Excel file.

# Libraries
- Pandas: used to read/write the Excel file;
- Docx: used to fill up the template file;
- SMTPLIB: used to send and e-mail when the script finished running;
- Datetime- used to get the current date.
