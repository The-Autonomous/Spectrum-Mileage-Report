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

class Utils:
    
    def __init__(self):
        # Initialize the Tkinter root window (necessary for file dialogs and other widgets)
        self.root = tk.Tk()
        self.root.withdraw()  # Hide the root window initially

        # Prompt the user for their first and last name
        self.first_name = simpledialog.askstring("Input", "Enter your first name:").capitalize().strip()
        self.last_name = simpledialog.askstring("Input", "Enter your last name:").capitalize().strip()

        self.d2d_rep = f"{self.first_name} {self.last_name}"

        # Open a file dialog to choose an Excel file
        self.file_path = load_file.askopenfilename(title="Select The Mileage Excel File", filetypes=[("Excel files", "*.xlsx;*.xls")])

        if self.file_path:
            try:
                self.file = pd.read_excel(io=self.file_path, engine='openpyxl', na_filter=False)
                self.file.columns = ['D2D Rep', 'Sales ID', 'Employee ID', 'Form ID', 'Form Name', 'FormInstanceID', 'Date Submitted', 
                  'Time Submitted', 'Address1', 'Address2', 'City', 'State', 'Zip', 'Distance from Entity']
                self.file['Date Submitted'] = pd.to_datetime(self.file['Date Submitted'], errors='coerce')
                self.file = self.file[self.file["D2D Rep"] == self.d2d_rep]
                self.available_dates = self.getDatesAvailable()
                self.file['Time Submitted'] = pd.to_datetime(
                    self.file['Time Submitted'], format='%H:%M:%S', errors='coerce'
                ).dt.time
                self.file = self.file.sort_values(by=['Date Submitted', 'Time Submitted'], ascending=[True, True])
            except Exception as Err_Code:
                print(Err_Code)
                self.__init__()
        else:
            print("No File Selected")
            self.__init__()
    
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
        available_dates = []
        for date in self.file["Date Submitted"]:
            if not available_dates.__contains__(date):
                available_dates.append(date)
        return available_dates
    
    def arrayContains(self, array: list, key: str):
        try:
            return array.index(key) > -1
        except ValueError as VE:
            return False
        except Exception as E:
            print(E)
            return False
    
    def selectDay(self):
        calendar_window = tk.Toplevel(self.root)  # New window for calendar
        calendar_window.title("Select Work Day")
        selected_date = None  # Initialize selected date as None

        # Create calendar widget
        cal = Calendar(calendar_window, selectmode='day', date_pattern='yyyy-mm-dd')
        cal.pack(pady=20)

        # Highlight available dates in the calendar
        available_dates_set = set(self.available_dates)  # Convert to a set for fast lookup
        for date in available_dates_set:
            cal.calevent_create(date, "Available", "highlight")
        cal.tag_config('highlight', background='lightblue')

        # Function to retrieve selected date
        def get_date():
            nonlocal selected_date
            selected_date = cal.get_date()
            selected_date = pd.Timestamp(selected_date).date()
            calendar_window.destroy()
            
        # Button to confirm date selection
        select_button = tk.Button(calendar_window, text="Select Date", command=get_date)
        select_button.pack(pady=10)

        # Wait for the calendar window to close
        calendar_window.wait_window()
        if not self.arrayContains(self.available_dates, selected_date):
            formatted_dates = str(self.available_dates)
            print(f"Not A Valid Day. Availability Includes:{formatted_dates}")
            return self.selectDay()
        print(f"Exiting Selector With {selected_date}")
        return selected_date
    
    def getAddress(self, current_address: list):
        try:
            return f'{current_address["Address1"]} {current_address["Address2"]}, {current_address["City"]}, {current_address["State"]}, {current_address["Zip"]}'
        except Exception as E:
            print(f"{E}; The Address List Is: {current_address}")

            
Session = Utils()

while True:
    CurrentDay = Session.loadDay(Session.selectDay())
    for x in CurrentDay:
        print(Session.getAddress(x))

raise

            
#pip install osmnx
from geopy.geocoders import Nominatim
import networkx as nx
import osmnx as ox
class Geography:
    def __init__(self):
        self.geocoder = Nominatim(user_agent="address_locator")
    
    def get_coordinates(self, address: str):
        # Sanitize the address
        sanitized_address = " ".join(address.split()).strip()
        location = self.geocoder.geocode(sanitized_address)
        if location:
            return location.latitude, location.longitude
        else:
            print(f"Unable to geocode address: {sanitized_address}")
            return None

    def calculate_distance(self, start_address: str, end_address: str):
        # Get coordinates for both addresses
        start_coords = self.get_coordinates(start_address)
        end_coords = self.get_coordinates(end_address)

        if not start_coords or not end_coords:
            print(f"Failed to calculate distance. Check the addresses:\n"
                  f"Start Address: {start_address}, End Address: {end_address}")
            return None

        try:
            # Create a graph around the start address with a specified distance
            G = ox.graph_from_point(start_coords, dist=10000, network_type="drive")
            
            # Get the nearest nodes for start and end points
            start_node = ox.distance.nearest_nodes(G, start_coords[1], start_coords[0])
            end_node = ox.distance.nearest_nodes(G, end_coords[1], end_coords[0])

            # Calculate the shortest path distance
            distance = nx.shortest_path_length(G, start_node, end_node, weight="length")
            return distance
        except Exception as e:
            print(f"Error calculating distance: {e}")
            return None
GPS = Geography()
    #print(GPS.calculate_distance(Session.getAddress(CurrentDay, 0), Session.getAddress(CurrentDay, 1))) -- Too Slow For It To Be Viable