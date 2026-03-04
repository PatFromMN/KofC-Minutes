"""
    This function allows the user to enter and edit
    the minutes for an Assembly meeting. It will also
    call the functions that allow the user to edit the 
    Assembly information and the Officers' list.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
from datetime import datetime
import os
from numpy import var
import assembly
import dialogs
import officers
import file_mgr

DISPLAY_FONT = ('Arial', 12)

class EditorGui:
    def __init__(self):
        self.editor_data = file_mgr.FileManager.default_editor_data()
        self.file_mgr = file_mgr.FileManager()
        self.assembly_data = assembly.AssemblyInfo()
        self.officer_data = officers.OfficerDatabase()
        self._inject_officers()

        self.root = tk.Tk()
        self.root.title(f"Minutes Editor for the {self.assembly_data.assembly_info['Assembly Name']}, {self.assembly_data.assembly_info['Assembly Number']}")

        # Get the screen width and height
        self.screen_width = self.root.winfo_screenwidth()
        self.screen_height = self.root.winfo_screenheight()

        # Set the window size to match the screen size
        # root.geometry(f"{screen_width}x{screen_height}")
        self.root.geometry("1000x600")

        self.create_editor_gui()

        self.root.mainloop()

    def _inject_officers(self):
        for office, info in self.officer_data.officers.items():
            self.editor_data['officers'][office] = {
                'name': info['name'],
                'attendance': '',
            }

    def create_editor_gui(self):
        self._create_menubar()
        self._create_notebook()
        self._create_meeting_info_tab()
        self._create_opening_ceremony_tab()
        self._create_roll_call_tab()
        self._create_minutes_tab()
        self._create_reports_tab()
        self._create_business_tab()
        self._create_council_tab()
        self._create_financial_tab()
        self._create_closing_ceremony_tab()

    def new_meeting(self):
        self.file_mgr.new_json()

    def open_working(self):
        loaded_data = self.file_mgr.open_json()

        if loaded_data:
            # Populated the GUI fields loaded data
            self.editor_data = loaded_data
            self.populate_gui_fields(self.editor_data)

    def save_working(self):
        self.file_mgr.save_to_file(self.editor_data)

    def save_as(self):
        self.file_mgr.save_file_as_json(self.editor_data)

    def export_to_word(self):
        self.file_mgr.convert_to_word(self.editor_data)

    def update_assembly_info(self):
        """Update the assemmly information."""
        assembly_dialog = dialogs.AssemblyInfoDialog(self.root, self.assembly_data)
        saved = assembly_dialog.show()

        if saved:
            self.assembly_data.save_assembly()
            messagebox.showinfo("Assembly Info Saved", "The assembly metadata was successfully saved.")
        else:
            messagebox.showinfo("Assembly Info Not Saved", "The assembly metadata was not saved. The inforamtion was not changed.")

    def update_officer_info(self):
        officer_dialog = dialogs.OfficerInfoDialog(self.root, self.officer_data)
        saved = officer_dialog.show()

        print(saved)

    # Save meeting information
    def save_meeting_date(self, event=None):
        date_value = self.date_entry.get().strip()
        self.editor_data['Meeting Info']['Meeting Date'] = date_value
        print(f"Saved meeting date: {date_value}")
    
    def save_start_time(self, event=None):
        time_value = self.start_time_entry.get().strip()
        self.editor_data['Meeting Info']['Start Time'] = time_value
        print(f"Saved start time: {time_value}")

    # Save the opening ceremony
    def save_opening_intentions(self, event=None):
        intentions_string = self.intentions_text.get(1.0, tk.END).strip()
        self.editor_data['Opening Ceremony']['Intentions'] = intentions_string
        print(f"Saved the prayer intetions: {intentions_string}")

    def save_opening_prayer(self, event=None):
        """Save the opening prayer."""
        prayer_string = self.opening_prayer_entry.get().strip()
        self.editor_data['Opening Ceremony']['Prayer'] = prayer_string
        print(f"Saved the Opening Prayer: {prayer_string}")

    def save_opening_prayer_leader(self, event=None):
        """Save the the prayer leader's name."""
        opening_leader_string = self.led_by_entry.get().strip()
        self.editor_data['Opening Ceremony']['Leader'] = opening_leader_string
        print(f"Saved the Prayer Leader: {opening_leader_string}")

    def save_pledge_leader(self, event=None):
        """Save the the Pledge leader's name."""
        pledge_leader_string = self.pledge_entry.get().strip()
        self.editor_data['Opening Ceremony']['Pledge'] = pledge_leader_string
        print(f"Saved the Pledge Leader: {pledge_leader_string}")

    # Save the meeting attendance
    def save_attendance(self):
        """Save the attendance data for officers."""
        for office, var in self.attendance_vars.items():
            self.editor_data['officers'][office]['attendance'] = var.get()
        print('Saved  attendance data.')
    
    def save_other_attendees(self, event=None):
        """Save the other attendees."""
        attendees_string = self.other_attendees_text.get(1.0, tk.END).strip()
        self.editor_data['attendees'] = attendees_string
        print(f"Saved other attendees: {attendees_string}")

    # Save the previous meeting minutes
    def save_corrections(self, event=None):
        """Save the corrections to the minutes."""
        corrections_value = self.corrections_entry.get(1.0, tk.END).strip()
        self.editor_data['Minutes']['Corrections'] = corrections_value
        print(f"Saved corrections to the minutes: {corrections_value}")

    def save_motion(self, event=None):
        motion_string = self.motion_entry.get().strip()
        self.editor_data['Minutes']['Motion to Approve'] = motion_string
        print(f"Saved the motion: {motion_string}")

    def save_second(self, event=None):
        second_string = self.seconded_entry.get().strip()
        self.editor_data['Minutes']['Seconded by'] = second_string
        print(f"Saved Seconded by: {second_string}")

    def save_approval(self, event=None):
        approval_value = self.approved_combobox.get().strip()
        self.editor_data['Minutes']['Approval'] = approval_value
        print(f"Saved the approval vote: {approval_value}")

    # Save the reports 
    def save_friars_report(self, event=None):
        friars_text = self.friars_report_text.get(1.0, tk.END).strip()
        self.editor_data['Reports']['Friar'] = friars_text
        print(f"Saved the Friars report: {friars_text}")

    def save_bills(self, event=None):
        bills_text = self.bills_report_text.get(1.0, tk.END).strip()
        self.editor_data['Reports']['Bills'] = bills_text
        print(f"Saved the Bills and Communications: {bills_text}")
    
    def save_comptrollers(self, event=None):
        comptrollers_text = self.comptrollers_report_text.get(1.0, tk.END).strip()
        self.editor_data['Reports']['Comptroller'] = comptrollers_text
        print(f"Saved the Comptroller's report: {comptrollers_text}")

    def save_standing_committees(self, event=None):
        committees_text = self.committee_report_text.get(1.0, tk.END).strip()
        self.editor_data['Reports']['Standing Committees'] = committees_text
        print(f"Saved the Bills and Communications: {committees_text}")

    def save_applications(self, event=None):
        applications_text = self.applications_report_text.get(1.0, tk.END).strip()
        self.editor_data['Reports']['Applications'] = applications_text
        print(f"Saved the applications report: {applications_text}")

    def save_trustees(self, event=None):
        trustees_text = self.trustees_report_text.get(1.0, tk.END).strip()
        self.editor_data['Reports']['Trustees'] = trustees_text
        print(f"Saved the Trustees report: {trustees_text}")

    def save_pursers(self, event=None):
        pursers_text = self.pursers_report_text.get(1.0, tk.END).strip()
        self.editor_data['Reports']['Purser'] = pursers_text
        print(f"Saved the pursers report: {pursers_text}")

    # Handle the Business tab
    def save_unfinished_business(self, event=None):
        unfinished_business_text = self.unfinished_business_text.get(1.0, tk.END).strip()
        self.editor_data['Business']['Unfinished Business'] = unfinished_business_text
        print(f"Saved the Unfinished Business: {unfinished_business_text}")

    def save_new_business(self, event=None):
        new_business_text = self.new_business_text.get(1.0, tk.END).strip()
        self.editor_data['Business']['New Business'] = new_business_text
        print(f"Saved the New Business: {new_business_text}")

    # Councils reports
    def save_riverside(self, event=None):
        riverside_text = self.riverside_report_text.get(1.0, tk.END).strip()
        self.editor_data['Council Reports']['Riverside'] = riverside_text
        print(f"Saved the Riverside report: {riverside_text}")

    def save_st_thomas(self, event=None):
        st_thomas_text = self.st_thomas_report_text.get(1.0, tk.END).strip()
        self.editor_data['Council Reports']['St Thomas'] = st_thomas_text
        print(f"Saved the St Thomas reports: {st_thomas_text}")

    def save_sacred_heart(self, event=None):
        sacred_heart_text = self.sacred_heart_report_text.get(1.0, tk.END).strip()
        self.editor_data['Council Reports']['Sacred Heart'] = sacred_heart_text
        print(f"Saved the Sacred Heart report: {sacred_heart_text}")

    def save_olph(self, event=None):
        olph_text = self.olph_report_text.get(1.0, tk.END).strip()
        self.editor_data['Council Reports']['OLPH'] = olph_text
        print(f"Saved the OLPH report: {olph_text}")

    # Closing ceremonies
    def save_closing_prayer(self, event=None):
        prayer_text = self.closing_prayer_entry.get().strip()
        self.editor_data['Closing Ceremony']['Closing Prayer'] = prayer_text
        print(f"Saved Closing prayer: {prayer_text}")

    def save_closing_intentions(self, event=None):
        intentions_text = self.closing_prayer_intentions_text.get(1.0, tk.END)
        self.editor_data['Closing Ceremony']['Intentions'] = intentions_text
        print(f"Saved the closing intentions: {intentions_text}")

    def save_closing_leader(self, event=None):
        prayer_leader_text = self.closing_led_by_entry.get().strip()
        self.editor_data['Closing Ceremony']['Leader'] = prayer_leader_text
        print(f"Saved the closing intentions: {prayer_leader_text}")

    def save_adjourned_time(self, event=None):
        adjourned_text = self.adjourned_time_entry.get().strip()
        self.editor_data['Closing Ceremony']['Adjurned At'] = adjourned_text
        print(f"Saved the closing intentions: {adjourned_text}")

    def save_next_officers_mtg(self, event=None):
        next_officers_mtg_text = self.next_officers_meeting_entry.get().strip()
        self.editor_data['Closing Ceremony']['Next Officers Meeting'] = next_officers_mtg_text
        print(f"Saved the closing intentions: {next_officers_mtg_text}")

    def save_next_bus_mtg(self, event=None):
        next_business_mtg_text = self.next_business_meeting_entry.get().strip()
        self.editor_data['Closing Ceremony']['Next Business Meeting'] = next_business_mtg_text
        print(f"Saved the closing intentions: {next_business_mtg_text}")

    def save_financials(self, event=None):
        """
        Save the financial report tab to working JSON
        
        :param self: Description
        :param event: The function is called when the "Save" button is clicked 
        or the tab is changed.
        """
        # Save the data from the general account
        self.editor_data['Financials']['General']['Start Balance'] = self.general_start_entry.get().strip()
        self.editor_data['Financials']['General']['Receipts'] = self.general_receipts_entry.get().strip()
        self.editor_data['Financials']['General']['Deposits'] = self.general_deposit_entry.get().strip()
        self.editor_data['Financials']['General']['Disbursements'] = self.general_spent_entry.get().strip()
        self.editor_data['Financials']['General']['End Balance'] = self.general_end_entry.get().strip()

        # Save the data from the Chalice account
        self.editor_data['Financials']['Chalice']['Start Balance'] = self.chalice_start_entry.get().strip()
        self.editor_data['Financials']['Chalice']['Receipts'] = self.chalice_receipts_entry.get().strip()
        self.editor_data['Financials']['Chalice']['Deposits'] = self.chalice_deposit_entry.get().strip()
        self.editor_data['Financials']['Chalice']['Disbursements'] = self.chalice_spent_entry.get().strip()
        self.editor_data['Financials']['Chalice']['End Balance'] = self.chalice_end_entry.get().strip()

        # Sve the data from the Flag account 
        self.editor_data['Financials']['Flag']['Start Balance'] = self.flag_start_entry.get().strip()
        self.editor_data['Financials']['Flag']['Receipts'] = self.flag_receipts_entry.get().strip()
        self.editor_data['Financials']['Flag']['Deposits'] = self.flag_deposit_entry.get().strip()
        self.editor_data['Financials']['Flag']['Disbursements'] = self.flag_spent_entry.get().strip()
        self.editor_data['Financials']['Flag']['End Balance'] = self.flag_end_entry.get().strip()

        # Save the deposits and withdrawals.
        self.editor_data['Financials']['Withdrawals'] = self.withdrawal_text.get(1.0, tk.END).strip()
        self.editor_data['Financials']['Deposits'] = self.deposit_text.get(1.0, tk.END).strip()

        print(f"Saving finincials: {self.editor_data['Financials']}")

    def populate_gui_fields(self, data):
        # load Meeting Info tab
        if 'Meeting Info' in data:
            self.date_entry.delete(0, tk.END)
            self.date_entry.insert(0, data['Meeting Info'].get('Meeting Date', ''))

            self.start_time_entry.delete(0, tk.END)
            self.start_time_entry.insert(0, data['Meeting Info'].get('Start Time', ''))
        # load Opening Ceremonies tab
        if 'Opening Ceremony' in data:
            self.intentions_text.delete(1.0, tk.END)
            self.intentions_text.insert(1.0, data['Opening Ceremony'].get('Intentions', ''))

            self.opening_prayer_entry.delete(0, tk.END)
            self.opening_prayer_entry.insert(0, data['Opening Ceremony'].get('Prayer', ''))

            self.led_by_entry.delete(0, tk.END)
            self.led_by_entry.insert(0, data['Opening Ceremony'].get('Leader'))

            self.pledge_entry.delete(0, tk.END)
            self.pledge_entry.insert(0, data['Opening Ceremony'].get('Pledge', ''))
        # load Roll Call tab
        if 'officers' in data:
            for office, info in data['officers'].items():
                if office in self.attendance_vars:
                    self.attendance_vars[office].set(info.get('attendance', 'Present'))
        
        if 'attendees' in data:
            self.other_attendees_text.delete(1.0, tk.END)
            self.other_attendees_text.insert(1.0, data.get('attendees', ''))
        # load Minutes tab
        if 'Minutes' in data:
            self.corrections_entry.delete(1.0, tk.END)
            self.corrections_entry.insert(1.0, data['Minutes'].get('Corrections', ''))
            
            self.motion_entry.delete(0, tk.END)
            self.motion_entry.insert(0, data['Minutes'].get('Motion to Approve', ''))
            
            self.seconded_entry.delete(0, tk.END)
            self.seconded_entry.insert(0, data['Minutes'].get('Seconded by', ''))
            
            # Set combobox value
            approval = data['Minutes'].get('Approval', 'Approved')
            self.approved_combobox.set(approval)
        # load Reports tab
        if 'Reports' in data:
            self.friars_report_text.delete(1.0, tk.END)
            self.friars_report_text.insert(1.0, data['Reports'].get('Friar', ''))

            self.bills_report_text.delete(1.0, tk.END)
            self.bills_report_text.insert(1.0, data['Reports'].get('Bills', ''))

            self.comptrollers_report_text.delete(1.0, tk.END)
            self.comptrollers_report_text.insert(1.0, data['Reports'].get('Comptroller', ''))

            self.pursers_report_text.delete(1.0, tk.END)
            self.pursers_report_text.insert(1.0, data['Reports'].get('Purser', ''))

            self.committee_report_text.delete(1.0, tk.END)
            self.committee_report_text.insert(1.0, data['Reports'].get('Standing Committees'))

            self.applications_report_text.delete(1.0, tk.END)
            self.applications_report_text.insert(1.0, data['Reports'].get('Applications', ''))

            self.trustees_report_text.delete(1.0, tk.END)
            self.trustees_report_text.insert(1.0, data['Reports'].get('Trustees', ''))
        # load Business tab
        if 'Business' in data:
            self.unfinished_business_text.delete(1.0, tk.END)
            self.unfinished_business_text.insert(1.0, data['Business'].get('Unfinished Business'), '')

            self.new_business_text.delete(1.0, tk.END)
            self.new_business_text.insert(1.0, data['Business'].get('New Business', ''))
        # load Council tab
        if 'Council Reports' in data:
            self.riverside_report_text.delete(1.0, tk.END)
            self.riverside_report_text.insert(1.0, data['Council Reports'].get('Riverside'), '')

            self.st_thomas_report_text.delete(1.0, tk.END)
            self.st_thomas_report_text.insert(1.0, data['Council Reports'].get('St Thomas'), '')
            
            self.sacred_heart_report_text.delete(1.0, tk.END)
            self.sacred_heart_report_text.insert(1.0, data['Council Reports'].get('Sacred Heart'), '')

            self.olph_report_text.delete(1.0, tk.END)
            self.olph_report_text.insert(1.0, data['Council Reports'].get('OLPH'), '')

        # load Closing Ceremony tab
        if 'Closing Ceremony' in data:
            self.closing_prayer_entry.delete(0, tk.END)
            self.closing_prayer_entry.insert(0, data['Closing Ceremony'].get('Closing Prayer', ''))

            self.closing_prayer_intentions_text.delete(1.0, tk.END)
            self.closing_prayer_intentions_text.insert(1.0, data['Closing Ceremony'].get('Intentions'), '')

            self.closing_led_by_entry.delete(0, tk.END)
            self.closing_led_by_entry.insert(0, data['Closing Ceremony'].get('Leader', ''))

            self.adjourned_time_entry.delete(0, tk.END)
            self.adjourned_time_entry.insert(0, data['Closing Ceremony'].get('Adjurned At', ''))

            self.next_officers_meeting_entry.delete(0, tk.END)
            self.next_officers_meeting_entry.insert(0, data['Closing Ceremony'].get('Next Officers Meeting', ''))

            self.next_business_meeting_entry.delete(0, tk.END)
            self.next_business_meeting_entry.insert(0, data['Closing Ceremony'].get('Next Business Meeting', ''))
        
        # TODO: load Financial Report
        if 'Financials' in data:
            # Load the general account data
            self.general_start_entry.delete(0, tk.END)
            self.general_start_entry.insert(0, data['Financials']['General'].get('Start Balance', ''))  
            self.general_receipts_entry.delete(0, tk.END)
            self.general_receipts_entry.insert(0, data['Financials']['General'].get('Receipts', ''))
            self.general_deposit_entry.delete(0, tk.END)
            self.general_deposit_entry.insert(0, data['Financials']['General'].get('Deposits', ''))
            self.general_spent_entry.delete(0, tk.END)
            self.general_spent_entry.insert(0, data['Financials']['General'].get('Disbursements', ''))
            self.general_end_entry.delete(0, tk.END)
            self.general_end_entry.insert(0, data['Financials']['General'].get('End Balance', ''))

            # Load the Chalice acccount data
            self.chalice_start_entry.delete(0, tk.END)
            self.chalice_start_entry.insert(0, data['Financials']['Chalice'].get('Start Balance', ''))  
            self.chalice_receipts_entry.delete(0, tk.END)
            self.chalice_receipts_entry.insert(0, data['Financials']['Chalice'].get('Receipts', ''))
            self.chalice_deposit_entry.delete(0, tk.END)
            self.chalice_deposit_entry.insert(0, data['Financials']['Chalice'].get('Deposits', ''))
            self.chalice_spent_entry.delete(0, tk.END)
            self.chalice_spent_entry.insert(0, data['Financials']['Chalice'].get('Disbursements', ''))
            self.chalice_end_entry.delete(0, tk.END)
            self.chalice_end_entry.insert(0, data['Financials']['Chalice'].get('End Balance', ''))

            # Load the Flag account data
            self.flag_start_entry.delete(0, tk.END)
            self.flag_start_entry.insert(0, data['Financials']['Flag'].get('Start Balance', ''))  
            self.flag_receipts_entry.delete(0, tk.END)
            self.flag_receipts_entry.insert(0, data['Financials']['Flag'].get('Receipts', ''))
            self.flag_deposit_entry.delete(0, tk.END)
            self.flag_deposit_entry.insert(0, data['Financials']['Flag'].get('Deposits', ''))
            self.flag_spent_entry.delete(0, tk.END)
            self.flag_spent_entry.insert(0, data['Financials']['Flag'].get('Disbursements', ''))
            self.flag_end_entry.delete(0, tk.END)
            self.flag_end_entry.insert(0, data['Financials']['Flag'].get('End Balance', ''))

            # Load the deposits and withdrawals.
            self.withdrawal_text.delete(1.0, tk.END)
            self.withdrawal_text.insert(1.0, data['Financials'].get('Withdrawals', ''))
            self.deposit_text.delete(1.0, tk.END)
            self.deposit_text.insert(1.0, data['Financials'].get('Deposits', ''))

    # Create the GUI
    def _create_menubar(self):
        # Create a menu bar
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label='New Meeting', command=self.new_meeting)
        file_menu.add_command(label='Open Working File', command=self.open_working)
        file_menu.add_command(label='Save Working File', command=self.save_working)
        file_menu.add_command(label='Save As', command=self.save_as)
        file_menu.add_separator()
        file_menu.add_command(label='Export to Word', command=self.export_to_word)

        # Add the Info menu to change Assembly and Officer data
        info_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Info", menu=info_menu)
        info_menu.add_command(label='Update Assembly Information', command=self.update_assembly_info)
        info_menu.add_command(label="Update Officer Information", command=self.update_officer_info)

        menubar.add_command(label='Quit', command=self.root.quit)

    def _create_notebook(self):
        # Create a notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    def _create_meeting_info_tab(self):
        #create the Meeting Info tab
        info_tab = ttk.Frame(self.notebook)
        self.notebook.add(info_tab, text='Meeting Info')
        ttk.Label(info_tab, text="Meeting Date:", font=DISPLAY_FONT).grid(column=0, row=0, sticky=tk.E, padx=5, pady=5)
        self.date_entry = ttk.Entry(info_tab, width=30)
        self.date_entry.grid(column=1, row=0, sticky=tk.W, padx=5, pady=5)
        self.date_entry.bind('<FocusOut>', self.save_meeting_date)

        ttk.Label(info_tab, text="Starting Time:", font=DISPLAY_FONT).grid(column=0, row=1, sticky=tk.E, padx=5, pady=5)
        self.start_time_entry = ttk.Entry(info_tab, width=30)
        self.start_time_entry.grid(column=1, row=1, sticky=tk.W, padx=5, pady=5)
        self.start_time_entry.bind('<FocusOut>', self.save_start_time)

    def _create_opening_ceremony_tab(self):
        # create the Opening Ceremony tab
        # TODO: make the scrolled text box match window resizing
        opening_tab = ttk.Frame(self.notebook)
        self.notebook.add(opening_tab, text='Opening Ceremony')
        item_row = 0
        ttk.Label(opening_tab, text='Intentions:', font=DISPLAY_FONT, anchor='w').grid(column=0, row=item_row, sticky=tk.W, padx=5, pady=5)
        self.intentions_text = scrolledtext.ScrolledText(opening_tab, width=80, height=5)
        self.intentions_text.grid(column=1, row=item_row, sticky=(tk.W, tk.E), padx=5, pady=5)
        self.intentions_text.bind('<FocusOut>', self.save_opening_intentions)
        item_row += 1
        ttk.Label(opening_tab, text='Opening Prayer:', font=DISPLAY_FONT, anchor='w').grid(column=0, row=item_row, sticky=tk.W, padx=5, pady=5)
        self.opening_prayer_entry = ttk.Entry(opening_tab, width=50, font=DISPLAY_FONT)
        self.opening_prayer_entry.grid(column=1, row=item_row, sticky=(tk.W, tk.E), padx=5, pady=5)
        self.opening_prayer_entry.bind('<FocusOut>', self.save_opening_prayer)
        item_row += 1
        ttk.Label(opening_tab, text='Led by:', font=DISPLAY_FONT, anchor='w').grid(column=0, row=item_row, sticky=tk.W, padx=5, pady=5)
        self.led_by_entry = ttk.Entry(opening_tab, width=50, font=DISPLAY_FONT)
        self.led_by_entry.grid(column=1, row=item_row, sticky=(tk.W, tk.E), padx=5, pady=5)
        self.led_by_entry.bind('<FocusOut>', self.save_opening_prayer_leader)
        item_row += 1
        ttk.Label(opening_tab, text='Pledge led by', font=DISPLAY_FONT, anchor='w').grid(column=0, row=item_row, sticky=tk.W, padx=5, pady=5)
        self.pledge_entry = ttk.Entry(opening_tab, width=50, font=DISPLAY_FONT)
        self.pledge_entry.grid(column=1, row=item_row, sticky=(tk.W, tk.E), padx=5, pady=5)
        self.pledge_entry.bind('<FocusOut>', self.save_pledge_leader)

    def _create_roll_call_tab(self):
        # create the Roll Call tab
        roll_call_tab = ttk.Frame(self.notebook)
        self.notebook.add(roll_call_tab, text='Roll Call')
        roll_call_tab.rowconfigure(0, weight=1)
        roll_call_tab.rowconfigure(1, weight=1)
        roll_call_tab.rowconfigure(2, weight=1)
        roll_call_tab.rowconfigure(3, weight=1)
        roll_call_tab.rowconfigure(4, weight=1)
        roll_call_tab.rowconfigure(5, weight=1)
        roll_call_tab.rowconfigure(6, weight=1)
        roll_call_tab.rowconfigure(7, weight=1)
        roll_call_tab.rowconfigure(8, weight=1)
        roll_call_tab.rowconfigure(9, weight=1)
        roll_call_tab.rowconfigure(10, weight=1)
        roll_call_tab.rowconfigure(11, weight=1)
        roll_call_tab.rowconfigure(12, weight=1)
        roll_call_tab.rowconfigure(13, weight=1)
        roll_call_tab.rowconfigure(14, weight=1)
        roll_call_tab.columnconfigure(0, weight=1)
        roll_call_tab.columnconfigure(1, weight=1)
        roll_call_tab.columnconfigure(2, weight=1)
        roll_call_tab.columnconfigure(3, weight=1)
        roll_call_tab.columnconfigure(4, weight=1)
        # 
        ttk.Label(roll_call_tab, text='Office', font=DISPLAY_FONT, anchor='w').grid(column=0, row=0, padx=5, pady=5, sticky='w')
        ttk.Label(roll_call_tab, text='Name', font=DISPLAY_FONT, anchor='w').grid(column=1, row=0, padx=5, pady=5, sticky='w')
        ttk.Label(roll_call_tab, text='Attendance', font=DISPLAY_FONT, anchor='w').grid(column=2, columnspan=3, row=0, padx=5, pady=5, sticky='w')

        # officer_data = officers.OfficerDatabase().officers
        row_index = 1
        self.attendance_vars = {}

        for office, info in self.officer_data.officers.items():
            var = tk.StringVar(value='Present')
            self.attendance_vars[office] = var

            ttk.Label(roll_call_tab, text=office, font=DISPLAY_FONT, anchor='w').grid(column=0, row=row_index, padx=5, pady=5, sticky='w')
            ttk.Label(roll_call_tab, text=info["name"], font=DISPLAY_FONT, anchor='w').grid(column=1, row=row_index, padx=5, pady=5, sticky='w')
            if officers.OfficerDatabase().officers[office]['name'].lower() != "vacant":
                ttk.Radiobutton(roll_call_tab, text='Present', value='Present', variable=var, command=self.save_attendance).grid(column=2, row=row_index, padx=5, pady=5, sticky='w')
                ttk.Radiobutton(roll_call_tab, text='Absent', value='Absent', variable=var, command=self.save_attendance).grid(column=3, row=row_index, padx=5, pady=5, sticky='w')
                ttk.Radiobutton(roll_call_tab, text='Excused', value='Excused', variable=var, command=self.save_attendance).grid(column=4, row=row_index, padx=5, pady=5, sticky='w')
            
            row_index += 1

        ttk.Label(roll_call_tab, text='Other Attendees:', font=DISPLAY_FONT).grid(column=0, row=14, sticky=tk.EW, padx=5, pady=5)
        self.other_attendees_text = scrolledtext.ScrolledText(roll_call_tab, width=80, height=5)
        self.other_attendees_text.grid(column=1, row=14, columnspan=4, sticky=tk.NSEW, padx=5, pady=5)
        self.other_attendees_text.bind('<FocusOut>', self.save_other_attendees)

    def _create_minutes_tab(self):
        # create a tab for the Minutes
        minutes_tab = ttk.Frame(self.notebook)
        self.notebook.add(minutes_tab, text='Minutes')
        minutes_tab.columnconfigure(0, weight=1)
        minutes_tab.columnconfigure(1, weight=4)
        ttk.Label(minutes_tab, text='Reading of Minutes:', font=DISPLAY_FONT).grid(column=0, row=0, columnspan=5, sticky=tk.N, padx=5, pady=5)
        ttk.Label(minutes_tab, text='Corrections:', font=DISPLAY_FONT).grid(column=0, row=1, sticky=tk.EW, padx=5, pady=5)
        self.corrections_entry = scrolledtext.ScrolledText(minutes_tab, width=80, height=5)
        self.corrections_entry.grid(column=2, row=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        self.corrections_entry.bind('<FocusOut>', self.save_corrections)
        ttk.Label(minutes_tab, text='Motion to Approve:', font=DISPLAY_FONT).grid(column=0, row=2, sticky=tk.EW, padx=5, pady=5)
        self.motion_entry = ttk.Entry(minutes_tab, width=50, font=DISPLAY_FONT)
        self.motion_entry.grid(column=2, row=2, sticky=tk.W, padx=5, pady=5)
        self.motion_entry.bind('<FocusOut>', self.save_motion)
        ttk.Label(minutes_tab, text='Seconded by:', font=DISPLAY_FONT).grid(column=0, row=3, sticky=tk.EW, padx=5, pady=5)
        self.seconded_entry = ttk.Entry(minutes_tab, width=50, font=DISPLAY_FONT)
        self.seconded_entry.grid(column=2, row=3, sticky=tk.W, padx=5, pady=5)
        self.seconded_entry.bind('<FocusOut>', self.save_second)
        ttk.Label(minutes_tab, text='Approval:', font=DISPLAY_FONT).grid(column=0, row=4, sticky=tk.EW, padx=5, pady=5)
        self.approved_combobox = ttk.Combobox(minutes_tab, width=20, values=('Approved', 'Corrected', 'Tabled'), state='readonly')
        self.approved_combobox.grid(column=2, row=4, sticky=tk.W, padx=5, pady=5)
        self.approved_combobox.current(0)
        self.approved_combobox.bind('<<ComboboxSelected>>', self.save_approval)

    def _create_reports_tab(self):
        # create the Reports tab
        reports_tab = ttk.Frame(self.notebook)
        self.notebook.add(reports_tab, text='Reports')
        reports_tab.rowconfigure(0, weight=1)
        reports_tab.rowconfigure(1, weight=1)
        reports_tab.rowconfigure(2, weight=1)
        reports_tab.rowconfigure(3, weight=1)
        reports_tab.rowconfigure(4, weight=1)
        reports_tab.rowconfigure(5, weight=1)
        reports_tab.rowconfigure(6, weight=1)
        reports_tab.columnconfigure(0, weight=1)
        reports_tab.columnconfigure(1, weight=3)
        ttk.Label(reports_tab, text="Faithful Friars Report:", font=DISPLAY_FONT).grid(column=0, row=0, sticky=tk.NW, padx=5, pady=5)
        self.friars_report_text = scrolledtext.ScrolledText(reports_tab, wrap=tk.WORD)
        self.friars_report_text.grid(column=1, row=0, sticky='nsew', padx=5, pady=5)
        self.friars_report_text.bind('<FocusOut>', self.save_friars_report)
        ttk.Label(reports_tab, text="Bills and Communications:", font=DISPLAY_FONT).grid(column=0, row=1, sticky=tk.W, padx=5, pady=5)
        self.bills_report_text = scrolledtext.ScrolledText(reports_tab, wrap=tk.WORD)
        self.bills_report_text.grid(column=1, row=1, sticky='nsew', padx=5, pady=5)
        self.bills_report_text.bind('<FocusOut>', self.save_bills)
        ttk.Label(reports_tab, text="Faithful Comptroller Report:", font=DISPLAY_FONT).grid(column=0, row=2, sticky=tk.W, padx=5, pady=5)
        self.comptrollers_report_text = scrolledtext.ScrolledText(reports_tab, wrap=tk.WORD)
        self.comptrollers_report_text.grid(column=1, row=2, sticky='nsew', padx=5, pady=5)
        self.comptrollers_report_text.bind('<FocusOut>', self.save_comptrollers)
        ttk.Label(reports_tab, text="Faithful Purser's Report:", font=DISPLAY_FONT).grid(column=0, row=3, sticky=tk.W, padx=5, pady=5)
        self.pursers_report_text = scrolledtext.ScrolledText(reports_tab, wrap=tk.WORD)
        self.pursers_report_text.grid(column=1, row=3, sticky='nsew', padx=5, pady=5)
        self.pursers_report_text.bind('<FocusOut>', self.save_pursers)
        ttk.Label(reports_tab, text="Standing Committees Report:", font=DISPLAY_FONT).grid(column=0, row=4, sticky=tk.W, padx=5, pady=5)
        self.committee_report_text = scrolledtext.ScrolledText(reports_tab, wrap=tk.WORD)
        self.committee_report_text.grid(column=1, row=4, sticky='nsew', padx=5, pady=5)
        self.committee_report_text.bind('<FocusOut>', self.save_standing_committees)
        ttk.Label(reports_tab, text="Reading of Applications:", font=DISPLAY_FONT).grid(column=0, row=5, sticky=tk.W, padx=5, pady=5)
        self.applications_report_text = scrolledtext.ScrolledText(reports_tab, wrap=tk.WORD)
        self.applications_report_text.grid(column=1, row=5, sticky='nsew', padx=5, pady=5)
        self.applications_report_text.bind('<FocusOut>', self.save_applications)
        ttk.Label(reports_tab, text="Trustees Report:", font=DISPLAY_FONT).grid(column=0, row=6, sticky=tk.W, padx=5, pady=5)
        self.trustees_report_text = scrolledtext.ScrolledText(reports_tab, wrap=tk.WORD)
        self.trustees_report_text.grid(column=1, row=6, sticky='nsew', padx=5, pady=5)
        self.trustees_report_text.bind('<FocusOut>', self.save_trustees)

    def _create_business_tab (self):
        # create the Business tab
        business_tab = ttk.Frame(self.notebook)
        self.notebook.add(business_tab, text='Business')
        business_tab.rowconfigure(0, weight=1)
        business_tab.rowconfigure(1, weight=1)
        business_tab.columnconfigure(0, weight=1)
        business_tab.columnconfigure(1, weight=3)

        ttk.Label(business_tab, text='Unfinished Business:', font=DISPLAY_FONT).grid(column=0, row=0, sticky=tk.NW, padx=5, pady=5)
        self.unfinished_business_text = scrolledtext.ScrolledText(business_tab, wrap=tk.WORD)
        self.unfinished_business_text.grid(column=1, row=0, sticky='nsew', padx=5, pady=5)
        self.unfinished_business_text.bind('<FocusOut>', self.save_unfinished_business)
        ttk.Label(business_tab, text='New Business:', font=DISPLAY_FONT).grid(column=0, row=1, sticky=tk.NW, padx=5, pady=5)
        self.new_business_text = scrolledtext.ScrolledText(business_tab, wrap=tk.WORD)
        self.new_business_text.grid(column=1, row=1, sticky='nsew', padx=5, pady=5)
        self.new_business_text.bind('<FocusOut>', self.save_new_business)

    def _create_council_tab(self): 
        # create a tab for the Council reports
        council_tab = ttk.Frame(self.notebook)
        self.notebook.add(council_tab, text='Council Reports')
        council_tab.rowconfigure(0, weight=1)
        council_tab.rowconfigure(1, weight=1)
        council_tab.rowconfigure(2, weight=1)
        council_tab.rowconfigure(3, weight=1)
        council_tab.columnconfigure(0, weight=1)
        council_tab.columnconfigure(1, weight=3)

        ttk.Label(council_tab, text='Riverside:', font=DISPLAY_FONT).grid(column=0, row=0, sticky=tk.NW, padx=5, pady=5)
        self.riverside_report_text = scrolledtext.ScrolledText(council_tab, wrap=tk.WORD)
        self.riverside_report_text.grid(column=1, row=0, sticky='nsew', padx=5, pady=5)
        self.riverside_report_text.bind('<FocusOut>', self.save_riverside)
        ttk.Label(council_tab, text='St. Thomas:', font=DISPLAY_FONT).grid(column=0, row=1, sticky=tk.NW, padx=5, pady=5)
        self.st_thomas_report_text = scrolledtext.ScrolledText(council_tab, wrap=tk.WORD)
        self.st_thomas_report_text.grid(column=1, row=1, sticky='nsew', padx=5, pady=5)
        self.st_thomas_report_text.bind('<FocusOut>', self.save_st_thomas)
        ttk.Label(council_tab, text='Sacred Heart:', font=DISPLAY_FONT).grid(column=0, row=2, sticky=tk.NW, padx=5, pady=5)
        self.sacred_heart_report_text = scrolledtext.ScrolledText(council_tab, wrap=tk.WORD)
        self.sacred_heart_report_text.grid(column=1, row=2, sticky='nsew', padx=5, pady=5)
        self.sacred_heart_report_text.bind('<FocusOut>', self.save_sacred_heart)
        ttk.Label(council_tab, text='Our Lady of Perpetual Help:', font=DISPLAY_FONT).grid(column=0, row=3, sticky=tk.NW, padx=5, pady=5)
        self.olph_report_text = scrolledtext.ScrolledText(council_tab, wrap=tk.WORD)
        self.olph_report_text.grid(column=1, row=3, sticky='nsew', padx=5, pady=5)
        self.olph_report_text.bind('<FocusOut>', self.save_olph)

    def _create_financial_tab(self):
        """
        Docstring for create_financial_tab
        
        :param self: Collect the financial information for the minutes.
        """
        finacial_tab = ttk.Frame(self.notebook)
        self.notebook.add(finacial_tab, text="Financial Statement")
        self.notebook.bind('<FocusOut>', self.save_financials)
        finacial_tab.rowconfigure(6, weight=1)
        finacial_tab.rowconfigure(7, weight=1)
        finacial_tab.columnconfigure(0, weight=1)
        finacial_tab.columnconfigure(1, weight=1)
        finacial_tab.columnconfigure(2, weight=1)
        finacial_tab.columnconfigure(3, weight=1)
        finacial_tab.columnconfigure(4, weight=1)

        ttk.Label(finacial_tab, text='General Fund Account', font=DISPLAY_FONT).grid(column=1, row=0, sticky=tk.N, padx=5, pady=5)
        ttk.Label(finacial_tab, text='Chalice Special Fund', font=DISPLAY_FONT).grid(column=2, row=0, sticky=tk.N, padx=5,pady=5)
        ttk.Label(finacial_tab, text='Flag Special Fund', font=DISPLAY_FONT).grid(column=3, row=0, sticky=tk.N, padx=5, pady=5)
        ttk.Label(finacial_tab, text='Starting Balance', font=DISPLAY_FONT).grid(column=0, row=1, sticky=tk.W, padx=5, pady=5)
        self.general_start_entry = ttk.Entry(finacial_tab,width=25, font=DISPLAY_FONT)
        self.general_start_entry.grid(column=1,row=1,sticky=tk.EW, padx=5, pady=5)
        self.chalice_start_entry = ttk.Entry(finacial_tab,width=25, font=DISPLAY_FONT)
        self.chalice_start_entry.grid(column=2,row=1,sticky=tk.EW, padx=5, pady=5)
        self.flag_start_entry = ttk.Entry(finacial_tab,width=25, font=DISPLAY_FONT)
        self.flag_start_entry.grid(column=3,row=1,sticky=tk.EW, padx=5, pady=5)
        ttk.Label(finacial_tab, text='Total Receipts', font=DISPLAY_FONT).grid(column=0, row=2, sticky=tk.W, padx=5, pady=5)
        self.general_receipts_entry = ttk.Entry(finacial_tab,width=25, font=DISPLAY_FONT)
        self.general_receipts_entry.grid(column=1,row=2,sticky=tk.EW, padx=5, pady=5)
        self.chalice_receipts_entry = ttk.Entry(finacial_tab,width=25, font=DISPLAY_FONT)
        self.chalice_receipts_entry.grid(column=2,row=2,sticky=tk.EW, padx=5, pady=5)
        self.flag_receipts_entry = ttk.Entry(finacial_tab,width=25, font=DISPLAY_FONT)
        self.flag_receipts_entry.grid(column=3,row=2,sticky=tk.EW, padx=5, pady=5)
        ttk.Label(finacial_tab, text='Funds Deposited', font=DISPLAY_FONT).grid(column=0, row=3, sticky=tk.W, padx=5, pady=5)
        self.general_deposit_entry = ttk.Entry(finacial_tab,width=25, font=DISPLAY_FONT)
        self.general_deposit_entry.grid(column=1,row=3,sticky=tk.EW, padx=5, pady=5)
        self.chalice_deposit_entry = ttk.Entry(finacial_tab,width=25, font=DISPLAY_FONT)
        self.chalice_deposit_entry.grid(column=2,row=3,sticky=tk.EW, padx=5, pady=5)
        self.flag_deposit_entry = ttk.Entry(finacial_tab,width=25, font=DISPLAY_FONT)
        self.flag_deposit_entry.grid(column=3,row=3,sticky=tk.EW, padx=5, pady=5)
        ttk.Label(finacial_tab, text='Total Disbusements', font=DISPLAY_FONT).grid(column=0, row=4, sticky=tk.W, padx=5, pady=5)
        self.general_spent_entry = ttk.Entry(finacial_tab,width=25, font=DISPLAY_FONT)
        self.general_spent_entry.grid(column=1,row=4,sticky=tk.EW, padx=5, pady=5)
        self.chalice_spent_entry = ttk.Entry(finacial_tab,width=25, font=DISPLAY_FONT)
        self.chalice_spent_entry.grid(column=2,row=4,sticky=tk.EW, padx=5, pady=5)
        self.flag_spent_entry = ttk.Entry(finacial_tab,width=25, font=DISPLAY_FONT)
        self.flag_spent_entry.grid(column=3,row=4,sticky=tk.EW, padx=5, pady=5)
        ttk.Label(finacial_tab, text='Ending Balance', font=DISPLAY_FONT).grid(column=0, row=5, sticky=tk.W, padx=5, pady=5)
        self.general_end_entry = ttk.Entry(finacial_tab,width=25, font=DISPLAY_FONT)
        self.general_end_entry.grid(column=1,row=5,sticky=tk.EW, padx=5, pady=5)
        self.chalice_end_entry = ttk.Entry(finacial_tab,width=25, font=DISPLAY_FONT)
        self.chalice_end_entry.grid(column=2,row=5,sticky=tk.EW, padx=5, pady=5)
        self.flag_end_entry = ttk.Entry(finacial_tab,width=25, font=DISPLAY_FONT)
        self.flag_end_entry.grid(column=3,row=5,sticky=tk.EW, padx=5, pady=5)

        ttk.Label(finacial_tab, text='Withdrawals', font=DISPLAY_FONT).grid(column=0, row=6, sticky=tk.W, padx=5, pady=5)
        self.withdrawal_text = scrolledtext.ScrolledText(finacial_tab, wrap=tk.WORD, height=5)
        self.withdrawal_text.grid(column=1, row=6,columnspan=3, sticky=tk.NSEW, padx=5, pady=5)

        ttk.Label(finacial_tab, text="Deposits", font=DISPLAY_FONT).grid(column=0, row=7, sticky=tk.W, padx=5, pady=5)
        self.deposit_text = scrolledtext.ScrolledText(finacial_tab, wrap=tk.WORD, height=5)
        self.deposit_text.grid(column=1, row=7,columnspan=3, sticky=tk.NSEW, padx=5, pady=5)

        # This tab is not used during meetings. 
        ttk.Button(finacial_tab, text="Save Page", command=self.save_financials).grid(row=8, column=0, padx=5, pady=5)

    def _create_closing_ceremony_tab(self):
        closing_tab = ttk.Frame(self.notebook)
        self.notebook.add(closing_tab, text='Closing Ceremony')
        closing_tab.rowconfigure(1, weight=1)
        closing_tab.columnconfigure(0, weight=1)
        closing_tab.columnconfigure(1, weight=3)

        ttk.Label(closing_tab, text='Closing Prayer:', font=DISPLAY_FONT).grid(column=0, row=0, sticky=tk.W, padx=5, pady=5)
        self.closing_prayer_entry = ttk.Entry(closing_tab, width=50, font=DISPLAY_FONT)
        self.closing_prayer_entry.grid(column=1, row=0, sticky=tk.W, padx=5, pady=5)
        self.closing_prayer_entry.bind('<FocusOut>', self.save_closing_prayer)
        ttk.Label(closing_tab, text='Intentions:', font=DISPLAY_FONT).grid(column=0, row=1, sticky=tk.NSEW, padx=5, pady=5)
        self.closing_prayer_intentions_text = scrolledtext.ScrolledText(closing_tab, wrap=tk.WORD, height=5)
        self.closing_prayer_intentions_text.grid(column=1, row=1, sticky=tk.NSEW, padx=5, pady=5)
        self.closing_prayer_intentions_text.bind('<FocusOut>', self.save_closing_intentions)
        ttk.Label(closing_tab, text='Led by:', font=DISPLAY_FONT).grid(column=0, row=2, sticky=tk.W, padx=5, pady=5)
        self.closing_led_by_entry = ttk.Entry(closing_tab, width=50, font=DISPLAY_FONT)
        self.closing_led_by_entry.grid(column=1, row=2, sticky=tk.W, padx=5, pady=5)
        self.closing_led_by_entry.bind('<FocusOut>', self.save_closing_leader)
        ttk.Label(closing_tab, text='Meeting Adjourned at:', font=DISPLAY_FONT).grid(column=0, row=3, sticky=tk.W, padx=5, pady=5)
        self.adjourned_time_entry = ttk.Entry(closing_tab, width=30)
        self.adjourned_time_entry.grid(column=1, row=3, sticky=tk.W, padx=5, pady=5)
        self.adjourned_time_entry.bind('<FocusOut>', self.save_adjourned_time)
        ttk.Label(closing_tab, text='Next Officers Meeting:', font=DISPLAY_FONT).grid(column=0, row=4, sticky=tk.W, padx=5, pady=5)
        self.next_officers_meeting_entry = ttk.Entry(closing_tab, width=30)
        self.next_officers_meeting_entry.grid(column=1, row=4, sticky=tk.W, padx=5, pady=5)
        self.next_officers_meeting_entry.bind('<FocusOut>', self.save_next_officers_mtg)
        ttk.Label(closing_tab, text='Next Business Meeting:', font=DISPLAY_FONT).grid(column=0, row=5, sticky=tk.W, padx=5, pady=5)
        self.next_business_meeting_entry = ttk.Entry(closing_tab, width=30)
        self.next_business_meeting_entry.grid(column=1, row=5, sticky=tk.W, padx=5, pady=5)
        self.next_business_meeting_entry.bind('<FocusOut>', self.save_next_bus_mtg)
        
