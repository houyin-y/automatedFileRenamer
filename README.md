# automatedFileRenamer
This project is used to rename files based on the data in an excel file (.xlsx).

It is a simple project for me to explore on python-based GUI, specifically Custom TKinter by [Tom Schimansky](https://github.com/tomschimansky/customtkinter)\
Thanks to him for creating Custom TKinter and making such an amazing and readable documentation. 🌠✨

## Building for dev
Steps required to run it locally:
1. Create a virtual venv
```
py -m venv .venv
```

2. Activate venv (Windows)
```
.venv\Scripts\activate.bat
```

3. Install the requirements
```
pip install -r requirements.txt
```

4. Run rename automator
```
py "rename automator.py"
```

## Compile into exe
After performing the following command, you can find your .exe file in /dist folder
```
pyinstaller --onefile --noconsole --add-data "resources;resources" "rename automator.py"
```
