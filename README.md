# APKG to Excel Converter

This project is a Flask-based web application that allows users to upload `.apkg` (Anki package) files, convert them to Excel format (`.xlsx`), and download the converted file. The application has been packaged into a standalone executable using PyInstaller, so it can be run as a desktop application.

## Features
- Drag-and-drop `.apkg` file upload interface.
- Converts Anki notes into a structured Excel format.
- Automatically cleans and formats the data from `.apkg` files.
- Download the resulting `.xlsx` file after processing.

## Requirements

To run this project as a Python application, you need the following:

- Python 3.x
- The following Python libraries:
  - Flask
  - pandas
  - openpyxl
  - sqlite3
  - re
  - html
  - random
  - datetime
  - zipfile

### Install Dependencies

```bash
pip install flask pandas openpyxl
```

If you are using this as an executable (.exe), no dependencies are required for the user.

### Running the Application as a Python Script
1.Clone this repository to your local machine.
git clone https://github.com/your-username/apkg-to-excel-converter.git
cd apkg-to-excel-converter
2.Run the Flask app.
python app.py
3.Open your browser and navigate to http://127.0.0.1:8000/ to use the web interface.

### Building the Executable with PyInstaller
If you'd like to distribute the application as a standalone executable (.exe), follow these steps:

1.Install PyInstaller.
pip install pyinstaller
2.Run PyInstaller to package the app:
pyinstaller --onefile --add-data "templates;templates" --add-data "static;static" app.py
3.After building, the executable file will be located in the dist/ directory.

### Usage
Using the Python Script
1.Launch the Flask server by running the app with:
python app.py
2.In your browser, visit http://127.0.0.1:8000/ to access the web interface.
3.Drag and drop or select an .apkg file to upload.
4.Once the conversion is done, an Excel file (.xlsx) will be available for download.


### Using the Executable
Run the executable (app.exe) that you created or downloaded from the dist/ folder.

1.The Flask server will start in the background.
2.Open your browser and visit http://127.0.0.1:8000/.
3.Use the drag-and-drop interface to upload .apkg files and download the converted .xlsx file.
### Folder Structure
apkg-to-excel-converter/
│
├── app.py                    # Main application file (Flask server)
├── templates/                # HTML templates for the web interface
│   └── index.html
├── static/                   # Static files (CSS, JavaScript, etc.)
│   └── style.css
├── dist/                     # Directory where the executable (.exe) is generated
├── build/                    # PyInstaller temporary build files
├── uploads/                  # Directory to store uploaded files
├── Created_Files/            # Directory where generated Excel files are stored
└── README.md                 # Project documentation

### Troubleshooting
### Common Errors
File Not Found (500 Error): Ensure the Created_Files directory exists and that the app has permission to write to it.
Executable Fails to Create File: Make sure paths are correctly handled inside the executable by using relative paths and current_app.root_path for Flask.

### License
This project is licensed under the MIT License. See the LICENSE file for details


### Key Sections in the README:

1. **Introduction**: Describes what the project does.
2. **Requirements**: Lists dependencies or prerequisites.
3. **Running the Application**: Instructions for running the app as a Python script or executable.
4. **Usage**: Explains how to use the app.
5. **Folder Structure**: Shows the project layout for easier navigation.
6. **Troubleshooting**: Common issues and solutions.
7. **License**: Details of the project's licensing.

You can modify any part of this to suit your project structure or specific needs.
