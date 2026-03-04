# Minutes Editor

Purpose: This program was developed to help prepare the minutes for a Knights of Columbus meeting. These meetings follow a specficit format.

## The Original Program

The program "minutes_editor.py" is deprecated. It was vibe coded as a prototype and as an experiment. The program served its purpose and is no longer being used.  

## Changes to Minutes Editor

The program was broken into the main program, four classes, and two csv files. Each class handles a specific function.  

### main.py

Main.py is used to start the program. On Linux machines, main is called by a bash command. On Windows and Linux, main.py can be started from a command line.  

### editor.py

Editor.py is used to create the graphical interface. The user can select any of the menubar options or tabs.  

Editor.py stores the active minutes in the dictionary editor_data. This is the single source for the minutes being created or edited.  

#### Data Dictionary

The data from the tabs is stored in a data dictionary - editor_data. The dictionary structure is defined in file_mgr.py.  

The data in the Entry, ScrolledText, and Radio Button fields are stored as soon as the field loses focus. The financial report is saved by use of a botton or when the tab loses focus. At this time, there is no means to save data based on time.  Unless the data is saved to a file, it will be lost when the program is closed.  

#### Window formatting

The widgets will shift to fit the available space in the window.  

### file_mgr.py

File_mgr.py creates the data dictionary for Minutes Editor. This dictionary contains the keys with null values.  

File_mgr.py opens and saves the JSON files used to during the process of recording and editing the minutes. JSON was chosen because of the small size of the file and the ease of converting the data between the dictionary and the file.  

File_mgr.py converts the minutes data into Word file as its final output.  Minutes Editor cannot import the Word file.  

### assembly.py

Assembly.py provides the Assembly Name and Number to editor.py. Because each entiy has a unique name and number, assemnly.py gets the information from assembly.py and creates the dictionary assembly_info. Editor.py and file_mgr.py read this data.  

At present, the user must modify the CSV directly.  

### officers.py

Officers.py provides the names of the officers to editor.py through the dictionary officers.  

At present, the user must modify the CSV directly.  
