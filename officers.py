""" Processes the Assembly officer data."""

import csv
import os
from tkinter import messagebox

class OfficerDatabase:
    """ Lists the Assemblies Officers and their attendance data"""
    # TODO: This class hard codes the officer data. It needs a means to 
    # revise the listing.
    def __init__ (self):
        self.officers = {}
        self.load_officers()

    def load_officers(self):
        """Load officer data from CSV file"""
        csv_file = 'officers.csv'
        if os.path.exists(csv_file):
            with open(csv_file, 'r') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    self.officers[row['Office']] = {
                        "name": row['Name'],
                        "attendance": ""
                    }
        else:
            # Create default CSV if it doesn't exist
            self.create_default_officers_csv()
            self.load_officers()
    
    def create_default_officers_csv(self):
        """Create default officers.csv with current data"""
        default_officers = {
            "Faithful Navigator": "Change Name",
            "Faithful Friar": "Change Name",
            "Faithful Admiral": "Change Name",
            "Faithful Captain": "Change Name",
            "Faithful Pilot" : "Change Name",
            "Faithful Comptroller" : "Change Name",
            "Faithful Scribe" : "Change Name",
            "Faithful Purser" : "Change Name",
            "Faithful Inner Sentinel" : "Change Name",
            "Faithful Outer Sentinel" : "Change Name",
            "Faithful Trustee (1 Yr)" : "Change Name",
            "Faithful Trustee (2 Yr)" : "Change Name",
            "Faithful Trustee (3 Yr)" : "Change Name",
        }
        with open('officers.csv', 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['Office', 'Name'])
            for office, name in default_officers.items():
                writer.writerow([office, name])
    
    def save_officers(self):
        """Save current officer data back to CSV"""
        with open('officers.csv', 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['Office', 'Name'])
            for office, info in self.officers.items():
                writer.writerow([office, info['name']])

    def revise_officer(self):
        """Opens a dialog to edit officer information"""
        # This would open a simple GUI dialog to edit officers
        # For now, just reload from CSV
        messagebox.showinfo("Revise Officers", "Revise the officer information in the officers.csv file and click OK. Close and reload Minutes Editor.")
        self.load_officers()
        print("Officers reloaded from CSV")

    def init_officer_data(self):
        """ Initializes the officers attendance data."""
        for office in self.officers:
            self.officers[office]['attendance'] = ""

