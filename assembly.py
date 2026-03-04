""" Processes Assmembly data."""
import csv
import os
from tkinter import messagebox

class AssemblyInfo:
    """ 
    Records and revises the Assembly information. 
    Input: assembly.csv
    Output: Assembly data as a dictionary
    """
    def __init__(self):
        self.assembly_info = {}
        self.load_assembly()

    def load_assembly(self):
        """Load assembly data from CSV file"""
        csv_file = 'assembly.csv'
        if os.path.exists(csv_file):
            with open(csv_file, 'r') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    self.assembly_info[row['Field']] = row['Value']
        else:
            # Create default CSV if it doesn't exist
            self.create_default_assembly_csv()
            self.load_assembly()
    
    def create_default_assembly_csv(self):
        """Create default assembly.csv"""
        with open('assembly.csv', 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['Field', 'Value'])
            writer.writerow(['Assembly Name', 'Add Name'])
            writer.writerow(['Assembly Number', 'xxxx'])
    
    def save_assembly(self):
        """Save current assembly data back to CSV"""
        with open('assembly.csv', 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['Field', 'Value'])
            for field, value in self.assembly_info.items():
                writer.writerow([field, value])
    
    def revise_assembly(self):
        """Reload assembly info from CSV"""
        messagebox.showinfo("Revise Assembly Info", "Revise the assembly information in the assembly.csv file and click OK. Close and reload Minutes Editor.")
        self.load_assembly()
        print("Assembly info reloaded from CSV")

