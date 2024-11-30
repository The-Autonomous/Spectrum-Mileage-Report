#pip install openpyxl
#pip install xlrd
#pip install tkcalendar
# Required imports
import pandas as pd
from pandas import option_context
import tkinter.filedialog as load_file
import tkinter as tk
from tkcalendar import Calendar
from tkinter import simpledialog
from geopy.geocoders import Nominatim
from geopy.distance import geodesic
from tkinter import scrolledtext
from tkinter import messagebox
from threading import Thread
from time import sleep
import json
import os

class Files:
    def __init__(self):
        # Set cachePath to a 'cache' folder in the script's directory
        script_dir = os.path.dirname(os.path.abspath(__file__))
        self.cachePath = os.path.join(script_dir, "cache")
        self.quickSave = {}

        # Create the cache folder if it doesn't exist
        if not os.path.exists(self.cachePath):
            os.makedirs(self.cachePath)

    def loadData(self, fileName):
        filePath = os.path.join(self.cachePath, fileName)
        if not os.path.exists(filePath):
            # If the file doesn't exist, create a blank JSON file
            with open(filePath, "w") as file:
                json.dump({}, file)  # Create an empty JSON object
            print(f"File '{fileName}' did not exist. A blank JSON file has been created at '{filePath}'.")
            return {}

        try:
            with open(filePath, "r") as file:
                data = json.load(file)  # Load JSON data
            return data  # Return the loaded data
        except json.JSONDecodeError:
            print(f"Error: File '{fileName}' is not a valid JSON file.")
            return None

    def saveData(self, fileName, data):
        filePath = os.path.join(self.cachePath, fileName)
        try:
            with open(filePath, "w") as file:
                json.dump(data, file, indent=4)  # Save data as JSON
            print(f"Data successfully saved to '{filePath}'.")
        except Exception as e:
            print(f"Error: Could not save data to '{filePath}'. {e}")
            
    def setQuickSave(self, fileName, functionName, dataRetrieval):
        def forwardSave():
            self.saveData(fileName, dataRetrieval())
        self.quickSave[functionName] = forwardSave

class Utils:
    
    ### BILL GATES ###
    
    def __init__(self):
        # Initialize the Tkinter root window (necessary for file dialogs and other widgets)
        self.root = tk.Tk()
        self.root.withdraw()  # Hide the root window initially
        self.completedScan = True # Initialize Variable
        
        # Init Cache
        userDataFileName = "userData.json"
        self.fileCache = Files()
        self.loadedCache = self.fileCache.loadData(userDataFileName)
        self.fileCache.setQuickSave(userDataFileName, "loadedCache", lambda: self.loadedCache)

        # Open a file dialog to choose an Excel file
        self.filePath = load_file.askopenfilename(title="Select The Mileage Excel File", filetypes=[("Excel files", "*.xlsx;*.xls")])

        if self.filePath:
            try:
                self.file = pd.read_excel(io=self.filePath, engine='openpyxl', na_filter=False)
                self.file.columns = ['D2D Rep', 'Sales ID', 'Employee ID', 'Form ID', 'Form Name', 'FormInstanceID', 'Date Submitted', 
                  'Time Submitted', 'Address1', 'Address2', 'City', 'State', 'Zip', 'Distance from Entity']
                self.file['Date Submitted'] = pd.to_datetime(self.file['Date Submitted'], errors='coerce')
                self.promptUser() # Get Name And Confirm Its Located In File
                self.formatForUser() # Format File List To User
                self.file['Time Submitted'] = pd.to_datetime(self.file['Time Submitted'], format='%H:%M:%S', errors='coerce').dt.time
                self.file = self.file.sort_values(by=['Date Submitted', 'Time Submitted'], ascending=[True, True]) # Sort Dates
                self.availableDates = self.getDatesAvailable()
                self.previouslySelectedDate = self.availableDates[0]
            except Exception as ErrorCode:
                messagebox.showerror("Error In Reading File", ErrorCode)
                self.__init__()
        else:
            messagebox.showwarning("Wrong File", "Please Provide A Valid Mileage Report!")
            self.__init__()
    
    def promptUser(self):
        # Prompt the user for their first and last name
        try:
            self.firstName = self.loadedCache["firstName"] or simpledialog.askstring("Input", "First Name:").capitalize().strip()
            self.lastName = self.loadedCache["lastName"] or simpledialog.askstring("Input", "Last Name:").capitalize().strip()
            self.d2d_rep = f"{self.firstName} {self.lastName}"
            if self.d2d_rep in self.file["D2D Rep"].values:
                self.loadedCache["firstName"], self.loadedCache["lastName"] = self.firstName, self.lastName
                self.fileCache.quickSave["loadedCache"]()
                return 
            else:
                messagebox.showwarning("Wrong User", "The Name You Provided Was Incorrect! Please Try Again!")
                self.loadedCache["firstName"], self.loadedCache["lastName"] = "", ""
                return self.promptUser()
        except Exception as E:
            print(f"Failure In Prompting Users Name. Resetting Saved Values And Running Again. {E}")
            self.loadedCache["firstName"], self.loadedCache["lastName"] = "", ""
            return self.promptUser()
            
    def formatForUser(self):
        self.file = self.file[self.file["D2D Rep"] == self.d2d_rep]
    
    ### UI GATES ###
    
    def selectDay(self):
        calendar_window = tk.Toplevel(self.root)  # New window for calendar
        calendar_window.title("Select Work Day")
        selected_date = None  # Initialize selected date as None

        # Create calendar widget
        cal = Calendar(calendar_window, selectmode='day', date_pattern='yyyy-mm-dd')
        cal.pack(pady=20, padx=50)

        # Highlight available dates in the calendar
        availableDates_set = set(self.availableDates)  # Convert to a set for fast lookup
        for date in availableDates_set:
            cal.calevent_create(date, "Available", "highlight")
        cal.tag_config('highlight', background='lightblue')

        # Function to retrieve selected date
        def get_date():
            nonlocal selected_date
            selected_date = cal.get_date()
            selected_date = pd.Timestamp(selected_date).date()
            self.previouslySelectedDate = selected_date
            calendar_window.destroy()
            
        # Button to confirm date selection
        select_button = tk.Button(calendar_window, text="Select Date", command=get_date)
        select_button.pack(pady=10)
        cal.selection_set(self.previouslySelectedDate)

        # Wait for the calendar window to close
        calendar_window.wait_window()
        if selected_date == None:
            raise SystemExit
        if not self.arrayContains(self.availableDates, selected_date):
            formatted_dates = str(self.availableDates)
            messagebox.showwarning("Wrong Date", "Please Provide A Valid Date!")
            print(f"Not A Valid Day. Availability Includes:{formatted_dates}")
            return self.selectDay()
        print(f"Exiting Selector With {selected_date}")
        return selected_date
    
    def displayOutput(self):
        if self.completedScan == False:
            print("Attempted Display With Already Existing Display In Action")
            return
        newWindow = tk.Tk()
        newWindow.withdraw()  # Hide the root window initially
        output_window = tk.Toplevel(newWindow)  # New window for output
        output_window.title("Address List")
        self.completedScan = False
        self.dataNeedingProcessed = []
        self.dataProcessed = []
        
        # Create a scrolled text widget
        self.displayed_text_area = scrolledtext.ScrolledText(output_window, wrap=tk.WORD, width=160, height=50)
        self.displayed_text_area.pack(padx=10, pady=10)
        self.displayed_text_area.config(state=tk.DISABLED)
        output_window.after(100, self.addOutput, output_window)
        output_window.wait_window()
        output_window.destroy()
        self.completedScan = True
    
    def addOutput(self, originalWindowObject):
        if self.completedScan != True:
            for data in self.dataNeedingProcessed:
                if not self.dataProcessed.__contains__(data):
                    self.displayed_text_area.config(state=tk.NORMAL)
                    self.displayed_text_area.insert(tk.INSERT, f"\n{data}")
                    self.displayed_text_area.config(state=tk.DISABLED)
                    self.dataProcessed.append(data)
                    print(f"Displaying Data {data}")
            if originalWindowObject:
                originalWindowObject.after(100, self.addOutput, originalWindowObject)
            
    def insertNewData(self, newData: str):
        self.dataNeedingProcessed.append(newData)
    
    ### LOGiC GATES ###
    
    def loadDay(self, date: str):
        # Ensure 'Date Submitted' is of type datetime
        self.file["Date Submitted"] = pd.to_datetime(self.file["Date Submitted"], errors='coerce')
        
        # Convert input date to datetime
        date = pd.to_datetime(date, errors='coerce')

        if date is pd.NaT:
            print(f"Invalid date format: {date}")
            return []

        # Filter rows by the provided date
        day_content = self.file[self.file["Date Submitted"] == date]

        # Collect rows into a list
        all_rows = []
        for _, current_row in day_content.iterrows():
            all_rows.append(current_row.to_dict())

        return all_rows
    
    def getDatesAvailable(self):
        availableDates = []
        for date in self.file["Date Submitted"]:
            if not availableDates.__contains__(date):
                availableDates.append(date)
        return availableDates
    
    def arrayContains(self, array: list, key: str):
        try:
            return array.index(key) > -1
        except ValueError as VE:
            return False
        except Exception as E:
            print(E)
            return False
        
    def getAddress(self, current_address: list):
        try:
            return f'{current_address["Address1"]} {current_address["Address2"]}, {current_address["City"]}, {current_address["State"]}, {current_address["Zip"]}'
        except Exception as E:
            print(f"{E}; The Address Listed Is: {current_address}")
            
    def isSameRoadAddress(self, address_start, address_end):
        try:
            return ''.join([i for i in address_start["Address1"] if not i.isdigit()]) == ''.join([i for i in address_end["Address1"] if not i.isdigit()])
        except Exception as E:
            print(f"{E}; The Address's Listed Are: {address_start}\n{address_end}")

    def waitForCompletion(self, desiredAchievment: bool):
        SecondsSpent, WaitTime = 0, 0.1
        while self.completedScan != desiredAchievment:
            sleep(WaitTime)
            SecondsSpent += WaitTime
        print(f"Completion Completed {desiredAchievment} in {WaitTime}s")

class Geography:
    def getCoordinates(self, address):
        geolocator = Nominatim(user_agent="address_locator", timeout=10)
        location = geolocator.geocode(address)
        if not location:
            return None
        return (location.latitude, location.longitude)

    def getDistance(self, address1, address2):
        try:
            coords_1 = self.getCoordinates(address1)
            coords_2 = self.getCoordinates(address2)
            
            if not coords_1 or not coords_2:
                return 0
            
            return geodesic(coords_1, coords_2).miles
        except:
            return 0

Session = Utils()
GPS = Geography()

if __name__ == "__main__":
    while True:
        Session.waitForCompletion(True)
        CurrentDay = Session.loadDay(Session.selectDay())
        PreviousAddress = ["", ""]
        TotalDaysMiles, TravelDistance = 0, 0
        Thread(target=Session.displayOutput).start()
        Session.waitForCompletion(False)
        for current_address in CurrentDay:
            if Session.completedScan == True:
                continue
            FormattedAddress = Session.getAddress(current_address)
            if PreviousAddress[0] != "" and not Session.isSameRoadAddress(PreviousAddress[0], current_address):
                TravelDistance = GPS.getDistance(PreviousAddress[1], FormattedAddress)
                TotalDaysMiles += TravelDistance
            PreviousAddress = [current_address, FormattedAddress]
            Session.insertNewData(f"{FormattedAddress} | {TravelDistance:.1f}mi")
        Session.insertNewData(f"{TotalDaysMiles:.1f}mi Traveled")
        