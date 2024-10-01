
<img width="482" alt="image" src="https://github.com/user-attachments/assets/c7164fe6-546f-47ce-94b6-37342dd07479">

# Overview
 I originally developed this app for my internship at the Illinois Prison Project in the Summer of 2024 after seeing the hell attorneys go through to put together hundreds of exhibits. The app is designed to manage PDF exhibits for legal cases. It allows users to select a main PDF document, add exhibit PDFs, and merge them into a single file. This app automates the cleaning, organizing, and merging of legal exhibits, providing an efficient way to compile court documents.

# Key Features:
- Drag-and-drop support for adding exhibits.
- Easily add, remove, and reorder exhibits.
- Automatic PDF cleaning to optimize file size and format.
- Combines main documents with corresponding exhibits.
- User-friendly interface for handling multiple exhibits at once.

# Installation
To install and run the app, follow these steps:

1. Ensure you have Python 3.x installed on your machine.
2. Install the required Python libraries:
`pip install tk pypdf pdfplumber fitz sv_ttk tkinterdnd2`
3. Run the script:
`python IPPExhibit.py`

Alternatively, you can compile the app into a standalone executable using PyInstaller:

`pyinstaller --onefile --noconsole IPPExhibit.py`

# How to Use
1. Select Main PDF
Click the Browse button next to "Main PDF" to select the main PDF document that you will be working with. The main PDF should be very simple and straightfoward - simply going to include pages that have "Exhibit A", "Exhibit B", etc. and there is supposed to be one of these per page. For example, page 1 should include "Exhibit A", page 2 should include "Exhibit B", and so on. This is important because the program then inserts the PDF files after the pages that correspond to the exhibit letters.

2. Add Exhibits
- In the "Exhibit Letter" field, enter the exhibit's identifier (e.g., "A", "B", etc.).
- Click Add Exhibit, and select the corresponding PDF files. You can add multiple files for the same exhibit.
- The selected exhibits will appear in the list below, along with their file names.

3. Manage Exhibits
- Double-click an exhibit in the list to manage files (reorder or remove files).
- You can drag and drop files into the list or reorder them by dragging.
- Use the Up and Down buttons to adjust the order of the files.

4. Select Output Path
Click the Browse button next to "Output Path" to specify where the merged PDF will be saved.

5. Merge PDFs
After selecting the main PDF and exhibits, and setting the output path, click Merge PDFs to create the final compiled document. A confirmation message will appear when the process is complete.

# Troubleshooting
## PDF Cleaning Issues
If a PDF fails to load or merge, the app attempts to repair it automatically. If this fails, ensure the PDF is not corrupted by opening it in a viewer. What I found fixes it (have not encountered a case yet where the program fails to repair a bad PDF, but just in case) is compressing and re-encoding the file through any online compression tool.

# Credits
All the awesome people at IPP who supported me and gave endless feedback throughout the development of the app!
