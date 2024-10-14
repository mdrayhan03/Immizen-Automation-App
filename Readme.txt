This app can create Initial Consultation, Service Agreement, Invoice and Money Receipt
This app is already installed so no need of installation.
This app can't make desktop shortcut. We will integrate this feature in the next update

Instruction:-
First run app with Registry.exe there will create a new folder DocFile. From the assets file copy MilestoneSheet.xlsx and paste(replace) it in DocFiie/Sheet. Your app is ready to work.

Important Folder:
assets, Helper, Screen 
DocFile will provide all save doc file and excel sheet


pyinstaller --onefile --noconsole --hidden-import=babel.numbers --icon=assets/icon.ico --name=Registry MainApplication.py

pyinstaller --onefile --noconsole --icon=assets/icon.ico --add-data "assets:assets" --add-data "DocFile:DocFile" --hidden-import=babel.numbers --name="Registry" MainApplication.py