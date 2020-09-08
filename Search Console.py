# Importing Libraries
import tkinter as tk
import time
import pickle
import os
import pandas as pd
from datetime import datetime, timedelta
from google_auth_oauthlib.flow import InstalledAppFlow
# from apiclient.discovery import build
from googleapiclient.discovery import build
from tkinter import ttk
from tkinter import messagebox
from googleapiclient.errors import HttpError as GoogleHttpError


class MyGUI(tk.Tk):

    def __init__(self):
        tk.Tk.__init__(self)
        self.initialize_tkinter_window()
        self.url_values = ['Custom URL',
                           'https://www.samedelman.com/']
        self.create_upper_frame()
        self.create_lower_frame()
        self.connect_to_search_console()
        self.folder = 'P:\Clients\Sam Edelman\Analytics\Data\Search Console'
        self.previous_count = 0

    def connect_to_search_console(self):
        """
        Function to connect to Search Console.
        Function looks for the client_secret in the folder, if it's not present then goes through the authentication process
        """
        self.update_output('Connecting to Search Console....')
        self.client_secret_location = 'C:\My Files\Python Projects\Client Secrets\Search Console Config\credentials.pickle'
        self.start_progress_bar()
        self.oauth_scope = ('https://www.googleapis.com/auth/webmasters.readonly', 'https://www.googleapis.com/auth/webmasters')

        # Building Connection
        try:
            self.credentials = pickle.load(open(self.client_secret_location, "rb"))
        except (OSError, IOError) as e:
            self.flow = InstalledAppFlow.from_client_secrets_file("client_secrets.json", scopes=self.oauth_scope)
            self.credentials = self.flow.run_console()
            pickle.dump(self.credentials, open("config/credentials.pickle", "wb"))

        # Connect to Search Console Service using the credentials
        self.webmasters_service = build('webmasters', 'v3', credentials=self.credentials)
        self.update_output('Connected to Search Console')

    def initialize_tkinter_window(self):
        """
        Function to initialize the Window and it's properties.
        """
        self.width, self.height = 500, 600  # Screen size of the initial window
        self.ws = self.winfo_screenwidth()  # width of the screen
        self.hs = self.winfo_screenheight()  # height of the screen

        # calculate x and y coordinates for the Tk root window
        self.x = (self.ws / 2) - (self.width / 2)
        self.y = (self.hs / 2) - (self.height / 2)
        self.geometry('%dx%d+%d+%d' % (self.width, self.height, self.x, self.y))

        self.title("Sam Edelman Search Console Data Extract")
        self.canvas = tk.Canvas(self, bg='#D6EAF8')
        self.canvas.place(relwidth=1, relheight=1)
        
        self.get_files_button = tk.Button(self, text='Get All Files', command=self.get_files)
        self.get_files_button.place(relx=0.18, rely=0.95, relwidth=0.15, relheight=0.05, anchor='w')

        # self.get_details_button = tk.Button(self, text='Get Details', command=self.get_details)
        # self.get_details_button.place(relx=0.43, rely=0.95, relwidth=0.15, relheight=0.05, anchor='w')

        self.clear_output_button = tk.Button(self, text='Clear', command=self.clear_output)
        self.clear_output_button.place(relx=0.43, rely=0.95, relwidth=0.15, relheight=0.05, anchor='w')

        self.exit_button = tk.Button(self, text='Exit', command=self.exit_application)
        self.exit_button.place(relx=0.68, rely=0.95, relwidth=0.15, relheight=0.05, anchor='w')

        self.progress_bar = ttk.Progressbar(self, orient='horizontal', length=286, mode='determinate')
        self.progress_bar.place(relx=0.5, rely=0.44, relwidth=0.75, relheight=0.05, anchor='n')

    def create_upper_frame(self):
        """
        Function to define the top frame and it's widgets
        """
        self.frame_1 = tk.Frame(self, bg='#5DADE2', bd=5)
        self.frame_1.place(relx=0.5, rely=0.1, relwidth=0.75, relheight=0.25, anchor='n')

        self.url_label = tk.Label(self.frame_1, text='Custom URL:', justify='left', anchor='w')
        self.url_label.place(relx=0, rely=0, relwidth=0.43, relheight=0.22)
        self.url_entry = tk.StringVar()
        # self.url_entry = tk.Entry(self.frame_1, bd=1, textvariable=self.url_entry)
        self.url_entry = ttk.Combobox(self.frame_1, values=self.url_values)
        self.url_entry.place(relx=0.45, rely=0, relwidth=0.54, relheight=0.22)

        self.start_date_label = tk.Label(self.frame_1, text='Start Date (YYYY-MM-DD):', justify='left', anchor='w')
        self.start_date_label.place(relx=0, rely=0.25, relwidth=0.43, relheight=0.22)
        self.start_date_entry = tk.StringVar()
        self.start_entry = tk.Entry(self.frame_1, bd=1, textvariable=self.start_date_entry)
        self.start_entry.place(relx=0.45, rely=0.25, relwidth=0.54, relheight=0.22)

        self.end_date_label = tk.Label(self.frame_1, text='End Date (YYYY-MM-DD):', justify='left', anchor='w')
        self.end_date_label.place(relx=0, rely=0.50, relwidth=0.43, relheight=0.22)
        self.end_date_entry = tk.StringVar()
        self.end_entry = tk.Entry(self.frame_1, bd=1, textvariable=self.end_date_entry)
        self.end_entry.place(relx=0.45, rely=0.50, relwidth=0.54, relheight=0.22)

        self.submit_button = tk.Button(self.frame_1, text='Submit', command=self.get_data)
        self.submit_button.place(relx=0.5, rely=0.80, relwidth=0.30, relheight=0.20, anchor='n')

    def create_lower_frame(self):
        """
        Function to define lower frame and it's widgets
        """
        self.frame_2 = tk.Frame(self, bg='#5DADE2', bd=5)
        self.frame_2.place(relx=0.5, rely=0.45, relwidth=0.75, relheight=0.45, anchor='n')
        
        # self.output_text = tk.StringVar()
        self.output_label = tk.Listbox(self.frame_2, justify='left')
        self.output_label.place(relwidth=1, relheight=1)
        
        self.scrollbar = tk.Scrollbar(self.frame_2)
        self.scrollbar.place(relx=1, rely=0, relwidth=0.03, relheight=1, anchor='ne')
        
        self.output_label.config(yscrollcommand=self.scrollbar.set)
        self.scrollbar.config(command=self.output_label.yview)

    def clear_output(self):
        self.output_label.delete(0, tk.END)

    def get_files(self):
        self.clear_output()
        try:
            self.all_files = os.listdir(self.folder)
            for file in self.all_files:
                self.update_output(file)
        except FileNotFoundError:
            messagebox.showwarning('File Not Found', 'Can\'t connect to the folder in order to get all the files.')
            self.update_output('Can\'t connect to the folder')

    def get_details(self):
        # self.selected_file = self.output_label.get(tk.ACTIVE)
        # self.selected_file_path = os.path.join(self.folder, self.selected_file)
        # self.file_df = pd.read_excel(self.selected_file_path)
        # self.count = pd.DataFrame.groupby.ag
        pass

    def date_range(self, start_date, end_date, delta=timedelta(days=1)):
        """
            The range is inclusive, so both start_date and end_date will be returned
            Args:
                start_date: The datetime object representing the first day in the range.
                end_date: The datetime object representing the second day in the range.
                delta: A datetime.timedelta instance, specifying the step interval. Defaults to one day.
            Yields:
                Each datetime object in the range.
            """
        self.current_date = start_date
        while self.current_date <= end_date:
            yield self.current_date
            self.current_date += delta

    def get_data(self):
        """
        Fetches the user provided data in the text labels and exports data from Search Console
        :return:
        """
        self.update_output(' ')
        self.max_rows = 25000
        self.i = 0
        self.output_rows = []
        self.start_date_text = self.start_date_entry.get()
        self.end_date_text = self.end_date_entry.get()
        self.site_url = self.url_entry.get()
        self.proceed = self.check_file()
        if self.proceed:
            try:
                self.request_data()
            except GoogleHttpError:
                messagebox.showerror('URL Invalid', message='The URL you entered in Invalid. Please try again.')
                self.update_output('URL Invalid')
                return
        else:
            pass

    def request_data(self):
        """
        Requests data from Search Console
        """
        self.max_rows = 25000
        self.output_rows = []
        try:
            self.start_date = datetime.strptime(self.start_date_text, "%Y-%m-%d")
            self.end_date = datetime.strptime(self.end_date_text, "%Y-%m-%d")
        except ValueError:
            messagebox.showerror('Error', 'Date value should be YYYY-MM-DD')
            return
        if self.start_date > self.end_date:
            messagebox.showerror('Error', 'Start Date can not be greater than the End Date')
            return
        else:
            pass
        for self.date in self.date_range(self.start_date, self.end_date):
            self.i = 0
            self.date = self.date.strftime("%Y-%m-%d")
            print(self.date)
            self.update_output('Fetching Data for: %s' % self.date)
            while True:

                self.start_progress_bar()
                self.request = {
                    'startDate': self.date,
                    'endDate': self.date,
                    'dimensions': ["query", "page"],
                    "searchType": "Web",
                    'rowLimit': self.max_rows,
                    'startRow': self.i * self.max_rows
                }
                self.response = self.webmasters_service.searchanalytics().query(siteUrl=self.site_url, body=self.request).execute()

                print()
                if self.response is None:
                    messagebox.showerror('Error', 'No Response from Search Console')
                    break
                if 'rows' not in self.response:
                    print("row not in response")
                    if self.i == 0:
                        messagebox.showwarning('Data unavailable', 'Data not available for %s.' % self.date)
                        self.update_output('0 Rows for %s' % self.date)
                    break
                else:
                    print('Fetching data for: %s......' % self.date)
                    for self.row in self.response['rows']:
                        self.keyword = self.row['keys'][0]
                        self.page = self.row['keys'][1]
                        self.output_row = [self.date, self.keyword, self.page, self.row['clicks'], self.row['impressions'], self.row['ctr'], self.row['position']]
                        self.output_rows.append(self.output_row)
                    self.i = self.i + 1
                self.rows_per_date = len(self.output_rows) - self.previous_count
                self.update_output('%s Rows for %s' % (self.rows_per_date, self.date))
                self.previous_count = len(self.output_rows)
            self.update_output(' ')
        self.start_progress_bar()
        self.df = pd.DataFrame(self.output_rows, columns=['date', 'query', 'page', 'clicks', 'impressions', 'ctr', 'avg_position'])
        self.export_df()

    def check_file(self):
        """
        Checks whether the file is already present in the desired folder
        :return:
            True -> File is not present
            False -> File is present
        """
        self.file = 'SE_Search_%s_%s.xlsx' % (self.start_date_text, self.end_date_text)
        self.file_name = os.path.join(self.folder, self.file)
        self.all_files = os.listdir(self.folder)
        if self.file not in self.all_files:
            return True
        else:
            self.duplicate_box = messagebox.askquestion('Data already present', message='Would you like to replace the data?')
            if self.duplicate_box == 'yes':
                return True
            else:
                return False

    def export_df(self):
        """
        Exports the report into an excel file in a desired format
        """
        self.folder = 'P:\Clients\Sam Edelman\Analytics\Data\Search Console'
        self.file = 'SE_Search_%s_%s.xlsx' % (self.start_date_text, self.end_date_text)
        self.rows, self.columns = self.df.shape
        self.file_name = os.path.join(self.folder, self.file)
        self.all_files = os.listdir(self.folder)
        if self.rows == 0:
            self.update_output('No Data for the specified Date Range')
        else:
            # if self.file not in self.all_files:
            self.update_output('Exporting to an Excel File')
            self.start_progress_bar()
            self.df.to_excel(self.file_name, sheet_name='Data', index=False)
            self.update_output('File: %s was exported.' % self.file)
            self.update_output('Folder: %s' % self.folder)
            self.update_output('Total no. of Rows and Columns: %s, %s' % (self.rows, self.columns))

    def exit_application(self):
        """
        Question box for exiting the application
        """
        self.exit_box = messagebox.askquestion('Exit Application', 'Are you sure you want to exit the application', icon='warning')
        if self.exit_box == 'yes':
            self.destroy()
        else:
            pass

    def update_output(self, text: str):
        """
        Updates the output label and adds in a blank line
        :param text: String text to add in to the output label
        """
        self.output_label.insert(tk.END, text)

    def start_progress_bar(self):
        """
        Starts the progress bar
        """
        self.progress_bar['maximum'] = 10
        for i in range(11):
            time.sleep(0.05)
            self.progress_bar['value'] = i
            self.progress_bar.update()
        self.progress_bar['value'] = 0

    def stop_progress_bar(self):
        """
        Stops the progress bar
        """
        self.progress_bar.stop()


app = MyGUI()
app.mainloop()
