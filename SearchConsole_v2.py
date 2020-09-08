# Importing Libraries
import tkinter as tk
import time
import pickle
import os
import pandas as pd
from datetime import datetime, timedelta
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from tkinter import ttk
from tkinter import messagebox
from googleapiclient.errors import HttpError as GoogleHttpError

# Defining Colors
LIGHT_BLUE = '#D6EAF8'

class MyGUI(tk.Tk):

    def __init__(self, width, height):
        tk.Tk.__init__(self)
        self.url_values = ['Custom URL',
                           'https://www.samedelman.com/']
        self.end_date_entry = ''
        # client_secret_location = 'C:\My Files\Python Projects\Client Secrets\Search Console Config\credentials.pickle'
        # oauth_scope = ('https://www.googleapis.com/auth/webmasters.readonly', 'https://www.googleapis.com/auth/webmasters')
        
        # # Building Connection
        # try:
        #     credentials = pickle.load(open(client_secret_location, "rb"))
        # except (OSError, IOError) as e:
        #     flow = InstalledAppFlow.from_client_secrets_file("client_secrets.json", scopes=self.oauth_scope)
        #     credentials = flow.run_console()
        #     pickle.dump(credentials, open("config/credentials.pickle", "wb"))

        # # Connect to Search Console Service using the credentials
        # webmasters_service = build('webmasters', 'v3', credentials=credentials)

        ws = self.winfo_screenwidth()
        hs = self.winfo_screenheight()

        # Calculating X and Y coordinates for the Tk root window
        x = (ws/2) - (width/2)
        y = (hs/2) - (height/2)
        self.geometry(f'{width}x{height}+{round(x)}+{round(y)}')

        self.title('Google Search Console Reporting Tool')
        tk.Canvas(self, bg=LIGHT_BLUE).place(relwidth=1, relheight=1)
        
        self._frame = None
        self.switch_frame(StartPage)

    def switch_frame(self, frame_class):
        new_frame = frame_class(self)
        if self._frame is not None:
            self._frame.destroy()
        self._frame = new_frame
        self._frame.place(relwidth=1, relheight=1)


class StartPage(tk.Frame):

    def __init__(self, master):

        tk.Frame.__init__(self, master, bg=LIGHT_BLUE)
        welcome_label = tk.Label(self, text='Welcome to Search Console Reporting Tool', bg=LIGHT_BLUE)
        welcome_label.place(relwidth=1, relheight=0.10)
        
        explanation = """
        This is where you can explain about your program.
        """
        tk.Label(self, text=explanation, bg=LIGHT_BLUE).place(relx=0.50, rely=0.20, relwidth=0.80, relheight=0.30, anchor='n')

        # preset_button = tk.Button(self, text='Pre-set Report', command=lambda: master.switch_frame(PresetReport))
        # preset_button.place(relx=0.20, rely=0.80, relwidth=0.20, relheight=0.05, anchor='nw')

        # custom_dim_met_button = tk.Button(self, text='Custom Report', command=lambda: master.switch_frame(CustomDimensionMetrics))
        # custom_dim_met_button.place(relx=0.60, rely=0.80, relwidth=0.200, relheight=0.05, anchor='nw')

        start_report = tk.Button(self, text="Let's start Reporting!!!", command=lambda: master.switch_frame(Reporting))
        start_report.place(relx=0.50, rely=0.80, relwidth=0.30, relheight=0.05, anchor='n')
        
        start_page_exit_button = tk.Button(self, text='Exit', command=exit_application)
        start_page_exit_button.place(relx=0.50, rely=0.90, relwidth=0.20, relheight=0.05, anchor='n')

class Reporting(tk.Frame):

    def __init__(self, master):

        self.folder = 'P:\Clients\Sam Edelman\Analytics\Data\Search Console'
        
        tk.Frame.__init__(self, master, bg=LIGHT_BLUE)
        self.init_window(master)

    def init_window(self, master):
        
        """ Upper Frame """

        frame_1 = tk.Frame(self, bg='#5DADE2', bd=5)
        frame_1.place(relx=0.5, rely=0.05, relwidth=0.75, relheight=0.25, anchor='n')
        url_label = tk.Label(frame_1, text='Custom URL:', justify='left', anchor='w')
        url_label.place(relx=0, rely=0, relwidth=0.43, relheight=0.22)
        url_entry = tk.StringVar()
        url_entry = ttk.Combobox(frame_1, values=master.url_values)
        url_entry.place(relx=0.45, rely=0, relwidth=0.54, relheight=0.22)

        start_date_label = tk.Label(frame_1, text='Start Date (YYYY-MM-DD):', justify='left', anchor='w')
        start_date_label.place(relx=0, rely=0.25, relwidth=0.43, relheight=0.22)
        start_date_entry = tk.StringVar()
        start_entry = tk.Entry(frame_1, bd=1, textvariable=start_date_entry)
        start_entry.place(relx=0.45, rely=0.25, relwidth=0.54, relheight=0.22)

        end_date_label = tk.Label(frame_1, text='End Date (YYYY-MM-DD):', justify='left', anchor='w')
        end_date_label.place(relx=0, rely=0.50, relwidth=0.43, relheight=0.22)
        end_date_entry = tk.StringVar()
        end_entry = tk.Entry(frame_1, bd=1, textvariable=end_date_entry)
        end_entry.place(relx=0.45, rely=0.50, relwidth=0.54, relheight=0.22)

        custom_button = tk.Button(frame_1, text='Custom Report', command=self.custom_report)
        custom_button.place(relx=0.21, rely=0.80, relwidth=0.30, relheight=0.20, anchor='n')
        
        preset_button = tk.Button(frame_1, text='Preset Report', command=self.preset_report)
        preset_button.place(relx=0.75, rely=0.80, relwidth=0.30, relheight=0.20, anchor='n')

        # submit_button = tk.Button(frame_1, text='Submit', command=lambda: self.check_values(url_entry.get(), start_date_entry.get(), end_date_entry.get()))
        # submit_button.place(relx=0.5, rely=0.80, relwidth=0.30, relheight=0.20, anchor='n')

        """ Lower Frame """
        frame_2 = tk.Frame(self, bg='#5DADE2', bd=0)
        # frame_2.place(relx=0.5, rely=0.45, relwidth=0.75, relheight=0.45, anchor='n')
        frame_2.place(relx=0.5, rely=0.60, relwidth=0.75, relheight=0.25, anchor='n')
        
        progress_bar = ttk.Progressbar(frame_2, orient='horizontal', length=286, mode='determinate')
        progress_bar.place(relwidth=10, relheight=0.03, anchor='n')

        self.output_label = tk.Listbox(frame_2, justify='left')
        self.output_label.place(rely=0.03, relwidth=1, relheight=1)

        scrollbar = tk.Scrollbar(frame_2)
        scrollbar.place(relx=1, rely=0.03, relwidth=0.03, relheight=1, anchor='ne')
        
        self.output_label.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.output_label.yview)

        """ Buttons """
        home_button = tk.Button(self, text='Home', command=lambda: master.switch_frame(StartPage))
        home_button.place(relx=0.00, rely=0.00, relwidth=0.10, relheight=0.05)

        clear_output_button = tk.Button(self, text='Clear', command=self.clear_output)
        clear_output_button.place(relx=0.43, rely=0.95, relwidth=0.15, relheight=0.05, anchor='w')

        exit_button = tk.Button(self, text='Exit', command=exit_application)
        exit_button.place(relx=0.68, rely=0.95, relwidth=0.15, relheight=0.05, anchor='w')

        

        self.connect_to_search_console()

    def connect_to_search_console(self):
        """
        Function to connect to Search Console.
        Function looks for the client_secret in the folder, if it's not present then goes through the authentication process
        """
        self.update_output('Connecting to Search Console....')
        client_secret_location = 'C:\\My Files\\Python Projects\\Client Secrets\\Search Console Config\\credentials.pickle'
        # self.start_progress_bar()
        oauth_scope = ('https://www.googleapis.com/auth/webmasters.readonly', 'https://www.googleapis.com/auth/webmasters')

        # Building Connection
        try:
            credentials = pickle.load(open(client_secret_location, "rb"))
        except (OSError, IOError):
            flow = InstalledAppFlow.from_client_secrets_file("client_secrets.json", scopes=oauth_scope)
            credentials = flow.run_console()
            pickle.dump(credentials, open("config/credentials.pickle", "wb"))

        # Connect to Search Console Service using the credentials
        webmasters_service = build('webmasters', 'v3', credentials=credentials)
        self.update_output('Connected to Search Console')

    def custom_report(self):
        dimension_list = ['Date', 'Page', 'Query']
        metric_list = ['Clicks', 'Impressions', 'CTR', 'Avg. Position']
        frame_3 = tk.Frame(self, bg='#5DADE2', bd=5)
        frame_3.place(relx=0.5, rely=0.30, relwidth=0.75, relheight=0.25, anchor='n')
        # test_label = tk.Label(frame_3, text='Custom Dimensions and Metrics')
        # test_label.pack()
        tk.Label(frame_3, text='Dimensions', bg='#5DADE2').place(relx=0.01, relwidth=0.20, relheight=0.10)
        dimension_listbox = tk.Listbox(frame_3, selectmode=tk.EXTENDED)
        dimension_listbox.place(relx=0.01, rely=0.12, relwidth=0.20, relheight=0.85)
        for dimension in dimension_list:
            dimension_listbox.insert(tk.END, dimension)
        
        tk.Label(frame_3, text='Metrics', bg='#5DADE2').place(relx=0.22, relwidth=0.20, relheight=0.10)
        metric_listbox = tk.Listbox(frame_3, selectmode=tk.EXTENDED)
        metric_listbox.place(relx=0.22, rely=0.12, relwidth=0.20, relheight=0.85)
        for metric in metric_list:
            metric_listbox.insert(tk.END, metric)

        tk.Button(frame_3).place(relx=0.45, rely=0.12, relwidth=0.03, relheigh=0.85)

    def preset_report(self):
        frame_3 = tk.Frame(self, bg='#5DADE2', bd=5)
        frame_3.place(relx=0.5, rely=0.30, relwidth=0.75, relheight=0.25, anchor='n')
        test_label = tk.Label(frame_3, text='Preset Report')
        test_label.pack()

    def clear_output(self):
        self.output_label.delete(0, tk.END)
    
    def update_output(self, text: str):
        self.output_label.insert(tk.END, text)

    def check_values(self, url, start_date_text, end_date_text):
        try:
            start_date = datetime.strptime(start_date_text, "%Y-%m-%d")
            end_date = datetime.strptime(end_date_text, "%Y-%m-%d")
            
        except ValueError:
            messagebox.showerror('Error', 'Date value should be YYYY-MM-DD')
            return
        if start_date > end_date:
            messagebox.showerror('Error', 'Start Date can not be greater than the End Date')
            return
        else:
            pass

class CustomDimensionMetrics(tk.Frame):

    def __init__(self, master):
        tk.Frame.__init__(self, master, bg=LIGHT_BLUE)
        welcome_label = tk.Label(self, text='Please choose the Dimensions and Metrics', bg=LIGHT_BLUE)
        welcome_label.place(relwidth=1, relheight=0.10)

        start_report_button = tk.Button(self, text="Let's start Reporting!!!", command=lambda: master.switch_frame(PresetReport))
        start_report_button.place(relx=0.50, rely=0.90, relwidth=0.25, relheight=0.05, anchor='n')

        dimension_listbox = tk.Listbox(self)
        dimension_listbox.place(relx=0.10, rely=0.10, relwidth=0.40, relheight=0.70)

def exit_application():
    """
    Question box for exiting the application
    """
    exit_box = messagebox.askquestion('Exit Application', 'Are you sure you want to exit the application', icon='warning')
    if exit_box == 'yes':
        # self.destroy()
        quit()
    else:
        pass

search_console = MyGUI(600, 600)
search_console.mainloop()