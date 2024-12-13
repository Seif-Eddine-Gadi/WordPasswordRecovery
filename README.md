<p align="center">
  <img src="https://i.imgur.com/UerzXyh.png" width="100" />
</p>

# Word Password Recovery

Word Password Recovery is a Python application that helps in recovering passwords from Word documents. This tool works by extracting and modifying settings in `.docx` files. It’s an easy-to-use tool for anyone who needs to recover a Word document password or simply modify a document in batch.
   
## Features
- Select `.docx` files for processing.
- Automatically creates a "Result" folder to store the recovered files.
- Easy-to-use graphical user interface (GUI) using PyQt5.
- Open-source and free to use under the MIT license.
# Installation

### Requirements:
- Python 3.x
- PyQt5 library

### Option 1: Download the Executable (No Python Setup Needed)

If you don’t want to mess with Python and dependencies, you can download the precompiled executable file (.exe) from the Releases section of this repository. Just download the file and run it directly on your Windows machine.

### Option 2: Run Using Python (For Developers)
1. **Clone the repository:**
   ```bash
   git clone https://github.com/Seif-Eddine-Gadi/WordPasswordRecovery.git

2. **Navigate to the project directory:**
   ```bash
   cd WordPasswordRecovery

3. **Navigate to the Source Code folder (if applicable):**
   ```bash
   cd "Source Code"

3. **Create and activate a virtual environment (optional but recommended):**
   ```bash
   python -m venv venv
   venv\Scripts\activate

4. **Install the required dependencies:**
   ```bash
   pip install -r requirements.txt

5. **Run the application:**
   ```bash
   python script.py

# Usage
- **Open the application:** Once you have either installed or downloaded the executable, launch the application.
- **Select the .docx file** The app will prompt you to choose a folder with .docx files that need to be processed.
- **Click "Start":**  After selecting the file, click the Start button to begin processing the files.
- **Check the "Result" folder:**  The app will automatically create a folder called Result in the directory of the selected files and place the processed files there.
- **Once the process is complete:**  A success message will appear, and you can check the Result folder for the recovered documents.

## License

[MIT](This project is licensed under the MIT License - see the LICENSE file for details.)


## Acknowledgements

 - [PyQt5](https://pypi.org/project/PyQt5/) Used for building the graphical user interface (GUI).
 - [Python](https://www.python.org/) Used for scripting and automation.


## Contributing

Contributions are always welcome!

1. **Fork this repository:**

   Click the "Fork" button at the top of the repository page.

2. **Create your feature branch:**
   ```bash
   git checkout -b feature-name

3. **Commit your changes:**
   ```bash
   git commit -m 'Add feature'

4. **Push to the branch:**
   ```bash
   git push origin feature-name

5. **Open a pull request:**

   On GitHub, navigate to the original repository and click "New Pull Request".   

Feel free to submit issues or pull requests to improve this project!
## Important Notes
- This tool only works with .docx files. For other formats like .doc, you may need to convert the file to .docx before processing.

- If you encounter issues while using the app, check for any error messages in the output and consult the repository's issues page for solutions.

- This tool is provided as-is, and you are free to modify, distribute, and use it as you wish, under the terms of the MIT License.
