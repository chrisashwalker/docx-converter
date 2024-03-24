
# .doc to .docx file converter
This Python utility uses the Microsoft Word COM API to convert the .doc files in a directory and its subdirectories into .docx files.

## Things to be aware of:
- I highly recommend taking a backup of the directory that you will run the conversion on.
- The program will ask you if you want to delete the original .doc files. 
- Don't use Word whilst the program is running.
- It's not fast. The Word COM API is quite slow to perform file I/O operations, so depending on how many files you need to convert, it could take a while. 

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
