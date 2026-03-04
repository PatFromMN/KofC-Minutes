import tkinter as tk
from tkinter import ttk

class AssemblyInfoDialog:
    """Dialog for editing Assembly information"""
    # This class was developed by Claude AI.
    
    def __init__(self, parent, assembly_info_obj):
        self.assembly_info_obj = assembly_info_obj
        self.result = None  # Will be True if saved, False if cancelled
        
        # Create dialog window
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Edit Assembly Information")
        self.dialog.geometry("400x200")
        self.dialog.resizable(False, False)
        
        # Make dialog modal
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Create widgets
        self.create_widgets()
        
        # Center dialog on parent window
        self.dialog.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() // 2) - (self.dialog.winfo_width() // 2)
        y = parent.winfo_y() + (parent.winfo_height() // 2) - (self.dialog.winfo_height() // 2)
        self.dialog.geometry(f"+{x}+{y}")
        
    def create_widgets(self):
        """Create and layout dialog widgets"""
        # Main frame with padding
        main_frame = ttk.Frame(self.dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Assembly Nameself.officer_info.officer_info['Faithful Trustee (1 Yr)'] = self.trustee1_entry.get().strip()
        ttk.Label(main_frame, text="Assembly Name:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.name_entry = ttk.Entry(main_frame, width=30)
        self.name_entry.grid(row=0, column=1, sticky=tk.EW, pady=5, padx=(10, 0))
        self.name_entry.insert(0, self.assembly_info_obj.assembly_info.get('Assembly Name', ''))
        
        # Assembly Number
        ttk.Label(main_frame, text="Assembly Number:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.number_entry = ttk.Entry(main_frame, width=30)
        self.number_entry.grid(row=1, column=1, sticky=tk.EW, pady=5, padx=(10, 0))
        self.number_entry.insert(0, self.assembly_info_obj.assembly_info.get('Assembly Number', ''))
        
        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=(20, 0))
        
        # Save button
        save_btn = ttk.Button(button_frame, text="Save", command=self.save)
        save_btn.pack(side=tk.LEFT, padx=5)
        
        # Cancel button
        cancel_btn = ttk.Button(button_frame, text="Cancel", command=self.cancel)
        cancel_btn.pack(side=tk.LEFT, padx=5)
        
        # Configure grid weights
        main_frame.columnconfigure(1, weight=1)
        
        # Bind Enter and Escape keys
        self.dialog.bind('<Return>', lambda e: self.save())
        self.dialog.bind('<Escape>', lambda e: self.cancel())
        
        # Focus on first entry
        self.name_entry.focus()
        
    def save(self):
        """Save changes to assembly info"""
        # Update the assembly_info dictionary
        self.assembly_info_obj.assembly_info['Assembly Name'] = self.name_entry.get().strip()
        self.assembly_info_obj.assembly_info['Assembly Number'] = self.number_entry.get().strip()
        
        # Save to CSV
        self.assembly_info_obj.save_assembly()
        
        self.result = True
        self.dialog.destroy()
        
    def cancel(self):
        """Close dialog without saving"""
        self.result = False
        self.dialog.destroy()
        
    def show(self):
        """Show dialog and wait for it to close"""
        self.dialog.wait_window()
        return self.result

class OfficerInfoDialog:
    """This class creates a dialog for editing officer metadata."""
    # This dialog was created using the model presented by Claude AI.
    # Written by Patrick Norris.

    def __init__(self, parent, officer_info_obj):
        self.officer_info = officer_info_obj
        self.result = None # Will be True if saved, False if cancelled

        # Create the dialog window
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Edit Officer Informaton")
        self.dialog.geometry("400x520")
        self.dialog.resizable(False, False)

        # Make dialog modal
        self.dialog.transient(parent)
        self.dialog.grab_set()

        # Create widgets
        self.create_widgets()

        # Center dialog on the parent window
        self.dialog.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() // 2) - (self.dialog.winfo_width() // 2)
        y = parent.winfo_y() + (parent.winfo_height() // 2) - (self.dialog.winfo_height() // 2)
        self.dialog.geometry(f"+{x}+{y}")

    def create_widgets(self):
        """Create the widgets and layout for the dialog."""
        # Main frame with padding
        main_frame = ttk.Frame(self.dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Faithful Navigator 
        ttk.Label(main_frame, text="Faithful Navigator:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.navigator_entry = ttk.Entry(main_frame, width=30)
        self.navigator_entry.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=5)
        self.navigator_entry.insert(0, self.officer_info.officers['Faithful Navigator'].get('name', ''))

        # Faithful Friar
        ttk.Label(main_frame, text="Faithful Friar:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.friar_entry = ttk.Entry(main_frame, width=30)
        self.friar_entry.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=5)
        self.friar_entry.insert(0, self.officer_info.officers['Faithful Friar'].get('name', ''))

        # Faithful Admiral
        ttk.Label(main_frame, text="Faithful Admiral:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.admiral_entry = ttk.Entry(main_frame, width=30)
        self.admiral_entry.grid(row=2, column=1, sticky=tk.EW, padx=5, pady=5)
        self.admiral_entry.insert(0, self.officer_info.officers['Faithful Admiral'].get('name', ''))

        # Faithful Captain
        ttk.Label(main_frame, text="Faithful Captain:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.captain_entry = ttk.Entry(main_frame, width=30)
        self.captain_entry.grid(row=3, column=1, sticky=tk.EW, padx=5, pady=5)
        self.captain_entry.insert(0, self.officer_info.officers['Faithful Captain'].get('name', ''))

        # Faithful Pilot
        ttk.Label(main_frame, text="Faithful Pilot:").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.pilot_entry = ttk.Entry(main_frame, width=30)
        self.pilot_entry.grid(row=4, column=1, sticky=tk.EW, padx=5, pady=5)
        self.pilot_entry.insert(0, self.officer_info.officers['Faithful Pilot'].get('name', ''))

        # Faithful Comptroller
        ttk.Label(main_frame, text="Faithful Comptroller:").grid(row=5, column=0, sticky=tk.W, pady=5)
        self.comptroller_entry = ttk.Entry(main_frame, width=30)
        self.comptroller_entry.grid(row=5, column=1, sticky=tk.EW, padx=5, pady=5)
        self.comptroller_entry.insert(0, self.officer_info.officers['Faithful Comptroller'].get('name', ''))

        # Failthful Scribe
        ttk.Label(main_frame, text="Faithful Scribe:").grid(row=6, column=0, sticky=tk.W, pady=5)
        self.scribe_entry = ttk.Entry(main_frame, width=30)
        self.scribe_entry.grid(row=6, column=1, sticky=tk.EW, padx=5, pady=5)
        self.scribe_entry.insert(0, self.officer_info.officers['Faithful Scribe'].get('name', ''))

        # Faithful Purser
        ttk.Label(main_frame, text="Faithful Purser:").grid(row=7, column=0, sticky=tk.W, pady=5)
        self.purser_entry = ttk.Entry(main_frame, width=30)
        self.purser_entry.grid(row=7, column=1, sticky=tk.EW, padx=5, pady=5)
        self.purser_entry.insert(0, self.officer_info.officers['Faithful Purser'].get('name', ''))

        # Faithful Inner Sentinel
        ttk.Label(main_frame, text="Faithful Inner Sentinel:").grid(row=8, column=0, sticky=tk.W, pady=5)
        self.inner_entry = ttk.Entry(main_frame, width=30)
        self.inner_entry.grid(row=8, column=1, sticky=tk.EW, padx=5, pady=5)
        self.inner_entry.insert(0, self.officer_info.officers['Faithful Inner Sentinel'].get('name', ''))

        # Faithful Outer Sentinel
        ttk.Label(main_frame, text="Faithful Outer Sentinel:").grid(row=9, column=0, sticky=tk.W, pady=5)
        self.outer_entry = ttk.Entry(main_frame, width=30)
        self.outer_entry.grid(row=9, column=1, sticky=tk.EW, padx=5, pady=5)
        self.outer_entry.insert(0, self.officer_info.officers['Faithful Outer Sentinel'].get('name', ''))

        # Faithful Trustee (1 Yr)
        ttk.Label(main_frame, text="Faithful Trustee (1 yr):").grid(row=10, column=0, sticky=tk.W, pady=5)
        self.trustee1_entry = ttk.Entry(main_frame, width=30)
        self.trustee1_entry.grid(row=10, column=1, sticky=tk.EW, padx=5, pady=5)
        self.trustee1_entry.insert(0, self.officer_info.officers['Faithful Trustee (1 Yr)'].get('name', ''))

        # Faithful Trustee (2 Yr)
        ttk.Label(main_frame, text="Faithful Trustee (2 Yr):").grid(row=11, column=0, sticky=tk.W, pady=5)
        self.trustee2_entry = ttk.Entry(main_frame, width=30)
        self.trustee2_entry.grid(row=11, column=1, sticky=tk.EW, padx=5, pady=5)
        self.trustee2_entry.insert(0, self.officer_info.officers['Faithful Trustee (2 Yr)'].get('name', ''))

        # Faithful Trustee (3 Yr)
        ttk.Label(main_frame, text="Faithful Trustee (3 Yr):").grid(row=12, column=0, sticky=tk.W, pady=5)
        self.trustee3_entry = ttk.Entry(main_frame, width=30)
        self.trustee3_entry.grid(row=12, column=1, sticky=tk.EW, padx=5, pady=5)
        self.trustee3_entry.insert(0, self.officer_info.officers['Faithful Trustee (3 Yr)'].get('name', ''))

        # Button Frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=13, column=0, columnspan=2, pady=(20, 0))

        # Save Button
        save_button = ttk.Button(button_frame, text='Save', command=self.save)
        save_button.pack(side=tk.LEFT, padx=5)

        # Cancel Button
        cancel_button = ttk.Button(button_frame, text='Cancel', command=self.cancel)
        cancel_button.pack(side=tk.LEFT, padx=5)

        # Configure grid weights
        main_frame.columnconfigure(1, weight=1)

        # Bind Enter and escape keys
        self.dialog.bind('<Return>', lambda e: self.save())
        self.dialog.bind('<Escape>', lambda e: self.cancel())

        # Focus on first entry
        self.navigator_entry.focus()

    def save(self):
        """Save changes to officer info."""
        # Update the officer_info dictionary
        self.officer_info.officers['Faithful Navigator']['name'] = self.navigator_entry.get().strip()
        self.officer_info.officers['Faithful Friar']['name'] = self.friar_entry.get().strip()
        self.officer_info.officers['Faithful Admiral']['name'] = self.admiral_entry.get().strip()
        self.officer_info.officers['Faithful Captain']['name'] = self.captain_entry.get().strip()
        self.officer_info.officers['Faithful Pilot']['name'] = self.pilot_entry.get().strip()
        self.officer_info.officers['Faithful Comptroller']['name'] = self.comptroller_entry.get().strip()
        self.officer_info.officers['Faithful Scribe']['name'] = self.scribe_entry.get().strip()
        self.officer_info.officers['Faithful Purser']['name'] = self.purser_entry.get().strip()
        self.officer_info.officers['Faithful Inner Sentinel']['name'] = self.inner_entry.get().strip()
        self.officer_info.officers['Faithful Outer Sentinel']['name'] = self.outer_entry.get().strip()
        self.officer_info.officers['Faithful Trustee (1 Yr)']['name'] = self.trustee1_entry.get().strip()
        self.officer_info.officers['Faithful Trustee (2 Yr)']['name'] = self.trustee2_entry.get().strip()
        self.officer_info.officers['Faithful Trustee (3 Yr)']['name'] = self.trustee3_entry.get().strip()

        # Save to CSV
        self.officer_info.save_officers()
        self.result = True
        self.dialog.destroy()

    def cancel(self):
        """Close dialog without saving."""
        self.result = False
        self.dialog.destroy()

    def show(self):
        """Show dialog and wait for it to close."""
        self.dialog.wait_window()
        return self.result
    