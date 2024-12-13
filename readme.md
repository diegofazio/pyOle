# pyOle

pyOle is a demonstration project that shows how to create an OLE server in Python. It also provides an example of packaging the project into an executable using PyInstaller. The demo includes the PyPDF2 library to showcase the use of third-party libraries from other languages through OLE.

---

## Features
- Implements an OLE server with Python.
- Demonstrates how to package the OLE server into a standalone `.exe` file using PyInstaller.
- Provides an example function to extract text from PDF files using PyPDF2.

---

## Requirements
Install the required libraries using the provided `requirements.txt` file:

```bash
pip install -r requirements.txt
python Scripts\pywin32_postinstall.py -install  # Run from python.exe folder
```

---

## Files
### Main Files
- **pyOle.py**: The main script that implements the OLE server.
- **requirements.txt**: Lists the Python dependencies.

### Packaging Files
- **installer.nsi**: NSIS script to create an installer for the project.
- **pyOle_file.spec** and **pyOle_dir.spec**: PyInstaller specification files for building the `.exe`.
---

## Usage

### Running pyOle.py
To register the OLE server, run:
```bash
python pyOle.py --register
```

To unregister the OLE server, run:
```bash
python pyOle.py --unregister
```

### Extracting Text from PDFs
Use the `extract` method exposed via OLE to extract text from PDF files. This method demonstrates integration with third-party libraries like PyPDF2.

---

## Packaging into Executable

### Using PyInstaller
1. Install PyInstaller if not already installed:
  - pyOle_file.spec: One file package
  - pyOle_dir.spec: One directory package
  - Usage: 
    ```bash
    pyinstaller pyOle_file.spec    # Create one file package
    pyinstaller pyOle_dir.spec     # Create one directory package
    ```
This will generate a standalone executable in 'dist' folder to install into any computer( no python needed ).
```bash
pyOle.exe --register
```

To unregister the OLE server, run:
```bash
pyOle.exe --unregister
```


### Creating an Installer with NSIS
1. Download and install NSIS from [NSIS Download Page](https://nsis.sourceforge.io/Download).
2. Open `installer.nsi` in the NSIS compiler.
3. Build the installer to create a distributable `.exe`.

---

## Notes
- Ensure all required dependencies are installed before building or running the project.
- For more information on PyInstaller, refer to [PyInstaller Documentation](https://pyinstaller.org).
- For more information on NSIS, refer to [NSIS Documentation](https://nsis.sourceforge.io/Docs).
- Tested with Python 3.12

---

## License
This project is licensed under the MIT License.

---

Feel free to clone, modify, and share this project to suit your needs.

🔗 **[Click here to donate.](https://paypal.me/DiegoHernanFazio)** 
