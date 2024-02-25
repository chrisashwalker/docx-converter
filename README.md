
# .doc to .docx file converter
This Python utility uses the Microsoft Word COM API to convert the .doc files in a directory and its subdirectories into .docx files.

## How to use it:
Open Command Prompt in the program directory and run the following commands.

### 1. Create the Python virtual environment
```
py -m venv .venv
```

### 2. Activate the Python virtual environment
```
.venv\Scripts\activate.bat
```

### 3. Install the program dependencies
```
py -m pip install -r requirements.txt
```

### 4. Run the program
```
py -m docx-converter
```

### 5. (Optional) Deactivate the virtual environment
```
deactivate
```
