

#convert png to icon
https://icoconvert.com/

#first run to create spec file
pyinstaller --icon tkpost.ico --nowindowed --noconsole --noconfirm tkpost.py

two directories are going to be created: build, dist 

in order to add necessary files and directories this step must be taken:
insert this block in front of tkpost.spec
added_files = [
         ( './test_config.json', '.' ),
         ( './utils.py', '.' ),
         ( './img/*png', 'img' ),
         ( './db/*.*', 'db' ),
         ( './export/*.*', 'export' ),
         ( './logs/*.*', 'logs' ),
         ( './responses/*.*', 'responses' ),
         ( './save/*.*', 'save' ),
         ]

and modify this row from:
datas=[],
to:
datas=added_files,           
  
#second run from spec file
pyinstaller --icon tkpost.ico --nowindowed --noconsole --noconfirm tkpost.spec




