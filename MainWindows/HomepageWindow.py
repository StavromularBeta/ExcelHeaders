import tkinter as Tk
from tkinter import font as tkFont
from tkinter import filedialog
import os, sys, inspect
import datetime
import re
import xlsxwriter
currentdir = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))
parentdir = os.path.dirname(currentdir)
sys.path.insert(0, parentdir)
parentdir = os.path.dirname(parentdir)
sys.path.insert(0, parentdir)
parentdir = os.path.dirname(parentdir)
sys.path.insert(0, parentdir)
sys.path.insert(0, currentdir)


class HomepageWindow(Tk.Frame):
    def __init__(self, parent, **kwargs):
        Tk.Frame.__init__(self, parent, **kwargs)
        self.parent = parent
        self.header_name_entry = Tk.Entry(self)
        self.select_button_font = tkFont.Font(size=18, weight='bold')
        self.date_dict = {1: 'JAN',
                          2: 'FEB',
                          3: 'MAR',
                          4: 'APR',
                          5: 'MAY',
                          6: 'JUN',
                          7: 'JUL',
                          8: 'AUG',
                          9: 'SEP',
                          10: 'OCT',
                          11: 'NOV',
                          12: 'DEC'}
        self.current_month_directory = ''
        self.last_month_directory = ''
        self.parsed_header_dictionary = {}
        self.header_information = []

    def homepage(self):
        select_button = Tk.Button(self, text="Go!", command=self.header_parsing_controller, font=self.select_button_font)
        Tk.Label(self, text="Jobnumber: ", font=self.select_button_font).grid(row=0, column=0, sticky=Tk.W, padx=10, pady=10)
        self.header_name_entry.grid(row=0, column=1, sticky=Tk.W, padx=10, pady=10)
        select_button.grid(row=0, column=2, sticky=Tk.W, padx=10, pady=10)

    def header_parsing_controller(self):
        self.get_current_and_last_month_directory()
        self.get_header_information()
        self.add_data_to_excel_file()

    def get_current_and_last_month_directory(self):
        self.current_month_directory = 'U:\\TXT-' + self.date_dict[datetime.datetime.now().month] + "\\"
        self.last_month_directory = 'U:\\TXT-' + self.date_dict[int(datetime.datetime.now().month)-1] + "\\"

    def get_header_information(self):
        item = self.header_name_entry.get()
        current_month_file_path = self.current_month_directory + 'W' + item + '.TXT'
        last_month_file_path = self.last_month_directory + 'W' + item + '.TXT'
        try:
            header = open(current_month_file_path, 'r')
            header_contents = header.read()
        except FileNotFoundError:
            try:
                header = open(last_month_file_path, 'r')
                header_contents = header.read()
            except FileNotFoundError:
                print("ERROR: header not found. Exiting.")
                sys.exit()
        self.header_parser_V2(header_contents)

    def header_parser_V2(self, header_contents):
        if header_contents == "Header not found":
            parsed_header_contents = ['no', 'no', 'no', 'no', 'no', 'no', 'no', 'no', 'no', 'no', 'no', 'no', 'no',
                                      'no', 'no', 'no']
            return parsed_header_contents
        name1 = header_contents[0:55].strip()
        date = header_contents[55:66].strip()
        time = header_contents[66:84].strip()
        w_number = header_contents[84:98].strip()
        remainder_of_header = header_contents[98:].strip().split("   ")
        remainder_of_header = [word for word in remainder_of_header if len(word) >= 1]
        if remainder_of_header[0][0] == "*":
            attn = remainder_of_header[0]
            sample_type = re.sub('[\n]', '', remainder_of_header[1])
            address_1 = remainder_of_header[2]
            sample_subtype = re.sub('[\n]', '', remainder_of_header[3])
            address_2 = remainder_of_header[4]
            number_of_samples = re.sub('[\n]', '', remainder_of_header[5])
            postal_code = re.sub('[\n]', '', remainder_of_header[6])
        else:
            attn = "*"
            address_1 = remainder_of_header[0]
            sample_type = re.sub('[\n]', '', remainder_of_header[1])
            address_2 = remainder_of_header[2]
            sample_subtype = re.sub('[\n]', '', remainder_of_header[3])
            postal_code = re.sub('[\n]', '', remainder_of_header[4])
            number_of_samples = re.sub('[\n]', '', remainder_of_header[5])
        email = "can't find email"
        payment_information = "can't find payment information"
        arrival_temp = "can't find arrival temperature"
        sampler = "can't find sampler information"
        phone_number = "can't find phone number"
        for item in remainder_of_header:
            if "TEL:" in item:
                phone_number = re.sub('[\n]', '', item)
            elif "@" in item:
                email = re.sub('[\n]', '', item)
            elif "group" in item:
                email = re.sub('[\n]', '', item)
            elif "Arrival temp" in item:
                arrival_temp = re.sub('[\n]', '', item)
            elif "Sampler" in item:
                sampler = re.sub('[\n]', '', item)
            elif "Pd" in item:
                payment_information = re.sub('[\n]', '', item)
        self.parsed_header_dictionary = {"primary name": name1.strip(),
                                         "date": date.strip(),
                                         "time": time.strip(),
                                         "jobnumber": w_number.strip(),
                                         "sample type": sample_type.strip(),
                                         "sample subtype": sample_subtype.strip(),
                                         "attention line": attn.strip(),
                                         "address line 1": address_1.strip(),
                                         "address line 2": address_2.strip(),
                                         "postal code": postal_code.strip(),
                                         "no. of samples": number_of_samples.strip(),
                                         "email": email.strip(),
                                         "phone number": phone_number.strip(),
                                         "arrival temperature": arrival_temp.strip(),
                                         "sampler (optional)": sampler.strip(),
                                         "payment information": payment_information.strip()
                                         }
        headerrow = 2
        for key, value in self.parsed_header_dictionary.items():
            labelstring = key + ": " + value
            Tk.Label(self, text=labelstring).grid(sticky=Tk.W, row=headerrow, column=0, columnspan=3, padx=10)
            headerrow += 1

    def add_data_to_excel_file(self):
        jobnumber = self.header_name_entry.get()
        target = r'T:\ANALYST WORK FILES\Peter\ExcelHeaders\ExcelHeaders\ '
        target = target[0:-1] + jobnumber + '.xlsx'
        Tk.Label(self, text="EXCEL FILE: " + target).grid(row=1, column=0, columnspan=3, sticky=Tk.W, padx=10)
        workbook = xlsxwriter.Workbook(target)
        worksheet = workbook.add_worksheet()
        row = 0
        for key, value in self.parsed_header_dictionary.items():
            worksheet.write(row, 0, key)
            worksheet.write(row, 1, value)
            row += 1
        workbook.close()



