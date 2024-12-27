import os, sys, subprocess

# List of external packages to ensure are installed
packages = ["pandas", "tkcalendar", "geopy", "requests"]

for pkg in packages:
    try: __import__(pkg)
    except: subprocess.check_call([sys.executable, "-m", "pip", "install", pkg])

# Import all required modules
import pandas as pd
import tkinter.filedialog as load_file
import tkinter as tk
from tkcalendar import Calendar
from tkinter import simpledialog, messagebox, ttk
from geopy.geocoders import Nominatim
from geopy.distance import geodesic
from threading import Thread
from time import sleep
import json, os, requests, sys, re

class SelfInstall:
    
    def __init__(self, CurrentSession):
        self.Session = CurrentSession
        self.ProductVersion = "1.0.0"
        self.VersionSearchTerm = "Version Number "
        self.BackupName = "Backup_Main.py"
        self.LoadURL = "https://raw.githubusercontent.com/The-Autonomous/Spectrum-Mileage-Report/refs/heads/main/README.md"
        self.DownloadURL = "https://raw.githubusercontent.com/The-Autonomous/Spectrum-Mileage-Report/refs/heads/main/main.py"
    
    def noInstall(self):
        print("Skipping Auto-Update...")
    
    def installUpdate(self):
        root = tk.Tk()
        root.title("Installing Update")
        root.attributes("-topmost", True)
        root.geometry("300x100")

        root.protocol("WM_DELETE_WINDOW", lambda: None)  # Ignore close events
        
        # Progress bar widget
        progress_label = tk.Label(root, text="Downloading update...")
        progress_label.pack(pady=10)
        progress = ttk.Progressbar(root, orient="horizontal", length=250, mode="determinate")
        progress.pack(pady=10)
        
        # Function to download and update the script
        skipToTheEnd = False
        def perform_update():
            global skipToTheEnd
            try:
                current_file_path = os.path.realpath(sys.argv[0])
                backup_file_path = os.path.join(os.path.dirname(current_file_path), self.BackupName)

                # Read the current script's contents
                with open(current_file_path, "r") as current_file:
                    content = current_file.read()

                # Write the contents to the backup file
                with open(backup_file_path, "w") as backup_file:
                    backup_file.write(content)
            except:
                messagebox.showerror("Backup Failed!", ErrorCode)
                skipToTheEnd = True
                
            try:
                # Fetch the script from the URL
                response = requests.get(self.DownloadURL, stream=True)
                total_size = int(response.headers.get('content-length', 0))
                downloaded_size = 0

                # Read and write in chunks while updating the progress bar
                with open(current_file_path, "wb") as file:
                    for chunk in response.iter_content(chunk_size=1024):
                        if chunk:
                            file.write(chunk)
                            downloaded_size += len(chunk)
                            progress["value"] = (downloaded_size / total_size) * 100
                            root.update_idletasks()
                
                progress_label.config(text="Update Complete!")
                progress["value"] = 100

                # Close the window after a short delay
                root.after(100, lambda: restart_script(current_file_path))

            except Exception as ErrorCode:
                messagebox.showerror("Error In Downloading Update", ErrorCode)
                root.destroy()
                
        def restart_script(script_path):
            root.destroy()
            os.execv(sys.executable, ['python'] + [script_path])  # Restart the script
        
        if skipToTheEnd == False:
            # Run the update in the Tkinter main loop
            root.after(10, perform_update)
            root.mainloop()
    
    def checkPublicRecord(self):
        LatestCode = requests.get(url=self.LoadURL).text
        for line in LatestCode.splitlines():
            if line.__contains__(self.VersionSearchTerm):
                LatestVersion = line.replace(self.VersionSearchTerm, "").strip()
                if not LatestVersion == self.ProductVersion:
                    Session.quickPrompt("Update Required!", [f"Keep Using Version {self.ProductVersion}", self.noInstall], [f"Update To Version {LatestVersion}", self.installUpdate])
    
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
        self.dateToShow = ""
        
        self.defaultOffice = "3315 N Ridge Rd E, Ashtabula, Oh, 44004"
        
        # Init Cache
        userDataFileName = "userData.json"
        self.fileCache = Files()
        self.loadedCache = self.fileCache.loadData(userDataFileName)
        self.fileCache.setQuickSave(userDataFileName, "loadedCache", lambda: self.loadedCache)

    def loadMileageDocument(self):
        while True:
            self.filePath = load_file.askopenfilename(
                title="Select The Mileage Excel File",
                filetypes=[("Excel files", "*.xlsx;*.xls")]
            )
            if not self.filePath:
                messagebox.showwarning("Wrong File", "Please Provide A Valid Mileage Report!")
                continue
            try:
                # Load the Excel file into a DataFrame
                self.file = pd.read_excel(io=self.filePath, engine='openpyxl', na_filter=False)
                
                # Define the required columns
                required_columns = [
                    'D2D Rep', 'Sales ID', 'Employee ID', 'Form ID', 'Form Name', 
                    'FormInstanceID', 'Date Submitted', 'Time Submitted', 
                    'Address1', 'Address2', 'City', 'State', 'Zip', 'Distance from Entity'
                ]
                
                # Ensure all required columns exist in the DataFrame
                for column in required_columns:
                    if column not in self.file.columns:
                        self.file[column] = ""  # Add missing column with blank values
                
                # Convert 'Date Submitted' to datetime
                self.file['Date Submitted'] = pd.to_datetime(self.file['Date Submitted'], errors='coerce')
                
                self.promptUser()  # Get Name and Confirm It's Located in File
                self.formatForUser()  # Format File List to User
                
                # Convert 'Time Submitted' to time
                self.file['Time Submitted'] = pd.to_datetime(
                    self.file['Time Submitted'], format='%H:%M:%S', errors='coerce'
                ).dt.time
                
                # Sort DataFrame by 'Date Submitted' and 'Time Submitted'
                self.file = self.file.sort_values(by=['Date Submitted', 'Time Submitted'], ascending=[True, True])
                
                # Get available dates
                self.availableDates = self.getDatesAvailable()
                
                if self.availableDates:
                    self.previouslySelectedDate = self.availableDates[0]
                else:
                    raise ValueError("No valid dates available in the file.")
                break  # Exit loop on successful processing
            except Exception as e:
                messagebox.showerror("Error In Reading File", str(e))
    
    def promptUser(self):
        # Prompt the user for their first and last name
        try:
            self.firstName = self.loadedCache["firstName"] or simpledialog.askstring("Input", "First Name:").capitalize().strip()
            self.lastName = self.loadedCache["lastName"] or simpledialog.askstring("Input", "Last Name:").capitalize().strip()
            self.homeAddress = self.loadedCache["homeAddress"] or simpledialog.askstring("Input", "Home Address (Street, City, Zip):").strip()
            self.homeZipCode = re.search(r'\b\d{5}\b', self.homeAddress)
            self.officeAddress = self.loadedCache["officeAddress"] or self.defaultOffice or simpledialog.askstring("Input", "Office Address (Street, City, State, Zip):").strip()
            self.d2d_rep = f"{self.firstName} {self.lastName}"
            
            if not Geography().getCoordinates(self.homeAddress):
                self.loadedCache["homeAddress"] = ""
                messagebox.showwarning("Wrong Address", "The Address You Provided Did Not Work! Please Try Another One! Remember To Use The Format Of Street, City, State, Zip!")
                return self.promptUser()
            
            if self.d2d_rep in self.file["D2D Rep"].values:
                self.loadedCache["firstName"], self.loadedCache["lastName"], self.loadedCache["homeAddress"], self.loadedCache["officeAddress"] = self.firstName, self.lastName, self.homeAddress, self.officeAddress
                self.fileCache.quickSave["loadedCache"]()
                return 
            else:
                messagebox.showwarning("Wrong User", "The Name You Provided Was Incorrect! Please Try Again!")
                self.loadedCache["firstName"], self.loadedCache["lastName"] = "", ""
                return self.promptUser()
                
        except Exception as E:
            print(f"Failure In Prompting Users Name. Resetting Saved Values And Running Again. {E}")
            self.loadedCache["firstName"], self.loadedCache["lastName"], self.loadedCache["homeAddress"], self.loadedCache["officeAddress"] = "", "", "", ""
            return self.promptUser()
            
    def formatForUser(self):
        self.file = self.file[self.file["D2D Rep"] == self.d2d_rep]
    
    ### UI GATES ###
    
    def quickPrompt(self, Header: str, Option1: list = None, Option2: list = None):
        """Options should be lists containing first the text to display, then the function to be called"""
        
        newWindow = tk.Tk()
        newWindow.attributes("-topmost", True)
        newWindow.withdraw()  # Hide the root window initially
        output_window = tk.Toplevel(newWindow)  # New window for output
        output_window.title(Header)
        def handleOption1():
            newWindow.destroy()
            Option1[1]()
            
        def handleOption2():
            newWindow.destroy()
            Option2[1]()
        
        button1 = tk.Button(output_window, text=Option1[0], command=handleOption1)
        button2 = tk.Button(output_window, text=Option2[0], command=handleOption2)

        # Position the buttons next to each other
        button1.pack(side=tk.LEFT, padx=10, pady=10)
        button2.pack(side=tk.LEFT, padx=10, pady=10)
        output_window.wait_window()
        try:
            newWindow.destroy()
        except:
            pass
    
    def selectDay(self):
        calendar_window = tk.Toplevel(self.root)  # New window for calendar
        calendar_window.title("Select Work Day")
        calendar_window.attributes("-topmost", True)
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
    
    
    def displayMouseWheelFunction(self, event):
        # Scroll by the amount of the wheel scroll, in this case 3 units
        self.displayed_text_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    def displayOutput(self):
        if self.completedScan == False:
            print("Attempted Display With Already Existing Display In Action")
            return
        self.newWindow = tk.Tk()
        self.newWindow.withdraw()  # Hide the root window initially
        output_window = tk.Toplevel(self.newWindow)  # New window for output
        output_window.attributes("-topmost", True)
        output_window.title(f"Address List For {str(self.dateToShow)}")
        self.completedScan = False
        self.dataNeedingProcessed = []
        self.dataProcessed = []
        
        # Create canvas and scrollbar
        self.displayed_text_canvas = tk.Canvas(output_window)
        scrollbar = tk.Scrollbar(output_window, orient="vertical", command=self.displayed_text_canvas.yview)
        self.displayed_text_canvas.configure(yscrollcommand=scrollbar.set)
        self.displayed_text_frame = tk.Frame(self.displayed_text_canvas)
        
        # Create a window on the canvas to place the frame
        self.displayed_text_canvas.create_window((0, 0), window=self.displayed_text_frame, anchor="nw")
        
        self.displayed_text_canvas.bind_all("<MouseWheel>", lambda event: self.displayMouseWheelFunction(event))
        
        # Pack scrollbar and canvas
        scrollbar.pack(side="right", fill="y")
        self.displayed_text_canvas.pack(side="left", fill="both", expand=True)
        
        output_window.after(100, self.addOutput, output_window)
        output_window.wait_window()
        output_window.destroy()
        self.completedScan = True
    
    def addOutput(self, originalWindowObject):
        if self.completedScan != True:
            for idx, data in enumerate(self.dataNeedingProcessed):
                if not self.dataProcessed.__contains__(data):
                    
                    button = tk.Button(self.displayed_text_frame, text=data, command=lambda d=data: self.copyToClipboard(str(d).split("|")[0].strip()))
                    button.grid(row=idx, column=0, sticky="w", pady=5)
                    
                    self.dataProcessed.append(data)
                    print(f"Displaying Data {data}")
                    
            # Update the scrollregion after adding buttons to the frame
            self.displayed_text_frame.update_idletasks()  # Make sure all buttons are laid out
            self.displayed_text_canvas.config(scrollregion=self.displayed_text_canvas.bbox("all"))  # Update scroll region

            if originalWindowObject:
                originalWindowObject.after(100, self.addOutput, originalWindowObject)
            
    def copyToClipboard(self, text):
        self.newWindow.clipboard_clear()  # Clear the clipboard
        self.newWindow.clipboard_append(text)  # Append the text to the clipboard
        self.newWindow.update()  # Update the clipboard
            
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
            return f'{current_address["Address1"]}, {current_address["City"]}, {current_address["State"]}, {current_address["Zip"]}'
        except Exception as E:
            print(f"{E}; The Address Listed Is: {current_address}")
    
    def normalizeAddress(self, address: dict):
            normalized = re.sub(r'^\d+\s+', '', str(address["Address1"]).lower().strip())
            return str(normalized)
        
    def isSameRoadAddress(self, address_start, address_end):
        try:
            return self.normalizeAddress(address_start) + str(address_start["Zip"]) == self.normalizeAddress(address_end) + str(address_end["Zip"])
        except Exception as E:
            print(f"{E}; The Address's Listed Are: {address_start}\n{address_end}")

    def waitForCompletion(self, desiredAchievment: bool):
        SecondsSpent, WaitTime = 0, 0.1
        while self.completedScan != desiredAchievment:
            sleep(WaitTime)
            SecondsSpent += WaitTime
        print(f"Completion Completed {desiredAchievment} in {WaitTime}s")
        
    def manualMainLoop(self):
        while True:
            self.waitForCompletion(True)
            
            self.dateToShow = self.selectDay()
            CurrentDay = self.loadDay(self.dateToShow)
            PreviousAddress = [{"Address1":self.homeAddress, "Zip":self.homeZipCode}, self.homeAddress]
            TotalDaysMiles, TravelDistance = 0, 0
            
            Thread(target=self.displayOutput).start()
            self.waitForCompletion(False)
            
            self.insertNewData(f"{self.homeAddress} | Home")
            self.insertNewData(f"{self.officeAddress} | Office\n\n")
            
            for current_address in CurrentDay:
                if self.completedScan == True:
                    continue
                FormattedAddress = self.getAddress(current_address)
                if PreviousAddress[0] != "" and not self.isSameRoadAddress(PreviousAddress[0], current_address):
                    TravelDistance = GPS.getDistance(PreviousAddress[1], FormattedAddress)
                    TotalDaysMiles += TravelDistance
                else:
                    TravelDistance = 0
                PreviousAddress = [current_address, FormattedAddress]
                self.insertNewData(f"{FormattedAddress} | {TravelDistance:.1f}mi")
            self.insertNewData(f"{TotalDaysMiles:.1f}mi Traveled")
            self.insertNewData(f"{self.homeAddress} | Home")
    
    def automaticMainLoop(self):
        while True:
            sleep(1)

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
        except Exception as E:
            print(E)
            return 0

if __name__ == "__main__":

    #########################
    #####Initialize Data#####
    #########################
    
    Session = Utils()
    GPS = Geography()
    
    #########################
    ##Check Version On Load##
    #########################

    InstallChecker = SelfInstall(CurrentSession=Session)
    InstallChecker.checkPublicRecord()
    
    #########################
    ##Start Mileage Process##
    #########################
    
    Session.loadMileageDocument()
    Session.quickPrompt("Select Below", ["Manually Do Mileage", Session.manualMainLoop], ["Automatically Do Mileage", Session.automaticMainLoop])