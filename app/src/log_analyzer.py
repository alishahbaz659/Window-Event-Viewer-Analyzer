import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog, ttk
from Evtx.Evtx import Evtx
import xml.etree.ElementTree as ET

# Function to read EVTX files and extract relevant data
def read_evtx(files):
    data = []
    namespace = {
        'ns': 'http://schemas.microsoft.com/win/2004/08/events/event'
    }
    
    for file in files:
        print(f"Reading file: {file}")
        with Evtx(file) as log:
            for record in log.records():
                xml_str = record.xml()
                root = ET.fromstring(xml_str)

                try:
                    system = root.find('.//ns:System', namespaces=namespace)
                    level = system.find('.//ns:Level', namespaces=namespace)
                    date_time = system.find('.//ns:TimeCreated', namespaces=namespace)
                    source = system.find('.//ns:Provider', namespaces=namespace)
                    event_id = system.find('.//ns:EventID', namespaces=namespace)
                    task_category = system.find('.//ns:Task', namespaces=namespace)

                    data.append({
                        "Level": level.text if level is not None else None,
                        "Date and Time": date_time.attrib['SystemTime'] if date_time is not None and 'SystemTime' in date_time.attrib else None,
                        "Source": source.attrib['Name'] if source is not None and 'Name' in source.attrib else None,
                        "Event ID": event_id.text if event_id is not None else None,
                        "Task Category": task_category.text if task_category is not None else None
                    })
                except Exception as e:
                    print(f"Error processing record: {e}")
                    continue

    df = pd.DataFrame(data)
    return df


# Function to preprocess logs
def preprocess_logs(df):
    # Drop rows with any NaN values
    df.dropna(inplace=True)
    
    # Convert 'Date and Time' to datetime format
    df['Date and Time'] = pd.to_datetime(df['Date and Time'], format='%Y-%m-%d %H:%M:%S.%f', errors='coerce')
    
    # Drop rows where 'Date and Time' could not be converted (invalid format)
    df.dropna(subset=['Date and Time'], inplace=True)
    
    # Convert 'Event ID' to integer
    df['Event ID'] = df['Event ID'].astype(int)
    
    return df

# Function to identify relevant Event IDs based on Source
def identify_event_ids(df):
    relevant_sources = ['Outlook', 'Teams', 'WinWord', 'PowerPoint', 'MicrosoftEdge', 'Excel']
    relevant_event_ids = df[df['Source'].isin(relevant_sources)]['Event ID'].unique()
    return relevant_event_ids

# Function to categorize event
def categorize_event(source, event_id):
    if event_id == 4624:
        return 'Logon'
    elif event_id == 4634:
        return 'Logoff'
    elif event_id == 4800:
        return 'Lock'
    elif event_id == 4801:
        return 'Unlock'
    elif event_id == 4778:
        return 'Session Connect'
    elif event_id == 4779:
        return 'Session Disconnect'
    elif source in ['Outlook', 'Teams', 'WinWord', 'PowerPoint', 'MicrosoftEdge', 'Excel']:
        return f'{source} Event'
    return 'Other'

# Function to process logs
def process_logs(df):
    activity_data = {}
    for _, row in df.iterrows():
        timestamp = row['Date and Time']
        event_id = row['Event ID']
        source = row['Source']
        category = categorize_event(source, event_id)
        user = 'Andrea'  # Adjust based on your data if user info is available
        
        if user not in activity_data:
            activity_data[user] = []
        
        if category in ['Logon', 'Logoff', 'Lock', 'Unlock', 'Session Connect', 'Session Disconnect', 
                        'Outlook Event', 'Teams Event', 'WinWord Event', 'PowerPoint Event', 
                        'MicrosoftEdge Event', 'Excel Event']:
            activity_data[user].append((timestamp, category))
    return activity_data

# Function to calculate activity
def calculate_activity(activity_data):
    activity_summary = {}
    for user, events in activity_data.items():
        daily_activity = {}
        login_time = None
        lock_time = None
        for timestamp, category in events:
            date = timestamp.date()
            if date not in daily_activity:
                daily_activity[date] = timedelta()
            
            if category == 'Logon' or category == 'Session Connect':
                login_time = timestamp
            elif category == 'Logoff' or category == 'Session Disconnect':
                if login_time:
                    daily_activity[date] += timestamp - login_time
                    login_time = None
            elif category == 'Lock':
                lock_time = timestamp
            elif category == 'Unlock':
                if lock_time:
                    daily_activity[date] -= timestamp - lock_time
                    lock_time = None
        
        activity_summary[user] = {date: min(activity, timedelta(hours=7)) for date, activity in daily_activity.items()}
    
    return activity_summary

# Function to display activity
def display_activity(activity_summary):
    for user, daily_activity in activity_summary.items():
        dates = list(daily_activity.keys())
        hours = [activity.total_seconds() / 3600 for activity in daily_activity.values()]
        
        plt.figure(figsize=(10, 5))
        plt.bar(dates, hours)
        plt.axhline(y=7, color='r', linestyle='-')
        plt.xlabel('Date')
        plt.ylabel('Active Hours')
        plt.title(f'User Activity for {user}')
        plt.xticks(rotation=45)
        plt.show()

# Function to open files and process logs
def open_files():
    files = filedialog.askopenfilenames(filetypes=[("EVTX files", "*.evtx")])
    if files:
        try:
            dataframes = [read_evtx([file]) for file in files]
            combined_df = pd.concat(dataframes, ignore_index=True)
            df = preprocess_logs(combined_df)
            
            # Identify relevant Event IDs
            relevant_event_ids = identify_event_ids(df)
            print(f"Relevant Event IDs: {relevant_event_ids}")
            
            activity_data = process_logs(df)
            activity_summary = calculate_activity(activity_data)
            display_activity(activity_summary)
        except Exception as e:
            print(f"Error processing files: {e}")

# Setting up the GUI
root = tk.Tk()
root.title("Activity Tracker")

frame = ttk.Frame(root, padding="10")
frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

open_button = ttk.Button(frame, text="Open Log Files", command=open_files)
open_button.grid(row=0, column=0, pady=10)

root.mainloop()
