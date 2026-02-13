# Knights of Columbus - Minutes
A purpose-built app to record minutes of my Assembly's meetings.

The Assembly information is contained in the Assembly CSV. The CSV can be changed to reflect the user's Assembly.
The names of the officers are contained in the Officers CSV. The CSV can be changed to reflect the user's Assembly.

The meeting is broken down into various tabs. The tabs are organized by topic. There is some hopping around as the meeting progresses. 
Meeting data is stored in a dictionary that is defined in the file manager class (file_mgr.py). The data is not saved by the program. The user must save the data to a JSON file. If no file is identified, the program will prompt the user for a file name. 

The user can reload a saved file and edit the data. Selecting save will overwrite the existing data. A good practice is to save the data to a new file. 

The user can exxport the data to a Word Document. The document's default file name is based on the date it was saved. The user may enter a different name.
The user can enter changes into the Minutes Editor or the Word document. 
