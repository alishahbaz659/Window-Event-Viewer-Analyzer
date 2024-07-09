import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog, filedialog, Canvas, PhotoImage, messagebox
from Evtx.Evtx import Evtx
import xml.etree.ElementTree as ET
import customtkinter as ctk
import os
import threading
import ttkbootstrap as ttk
from ttkbootstrap.dialogs import Messagebox

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
    
    # Filter events between 7 AM and 7 PM
    df = df[(df['Date and Time'].dt.time >= datetime.strptime('07:00', '%H:%M').time()) &
            (df['Date and Time'].dt.time <= datetime.strptime('19:00', '%H:%M').time())]
    
    return df

# Function to identify relevant Event IDs based on Source
def identify_event_ids(df):
    relevant_sources = ['Outlook', 'Teams', 'WinWord', 'PowerPoint', 'MicrosoftEdge', 'Excel', 'Chrome', 'ESENT']
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
    elif source in ['Outlook', 'Teams', 'WinWord', 'PowerPoint', 'MicrosoftEdge', 'Excel', 'Chrome', 'ESENT']:
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
                        'MicrosoftEdge Event', 'Excel Event', 'Chrome Event', 'ESENT Event']:
            activity_data[user].append((timestamp, category))
    return activity_data

# Function to calculate activity summary
def calculate_activity(activity_data):
    activity_summary = {}
    for user, events in activity_data.items():
        daily_activity = {}
        activity_timers = {  # Dictionary to store timers for specific activities
            'Logon': None, 'Logoff': None, 'Lock': None, 'Unlock': None,
            'Session Connect': None, 'Session Disconnect': None,
            'Outlook Event': None, 'Teams Event': None, 'WinWord Event': None,
            'PowerPoint Event': None, 'MicrosoftEdge Event': None, 'Excel Event': None, 'Chrome Event': None, 'ESENT Event': None,
        }
        
        for timestamp, category in events:
            date = timestamp.date()
            
            # Initialize daily activity if not already initialized
            if date not in daily_activity:
                daily_activity[date] = {cat: timedelta() for cat in activity_timers}
            
            # Handle specific activities
            if category in ['Logon', 'Session Connect']:
                activity_timers['Logon'] = timestamp
            elif category in ['Logoff', 'Session Disconnect']:
                if activity_timers['Logon']:
                    duration = timestamp - activity_timers['Logon']
                    daily_activity[date]['Logon'] += duration
                    activity_timers['Logon'] = None
            elif category == 'Lock':
                activity_timers['Lock'] = timestamp
            elif category == 'Unlock':
                if activity_timers['Lock']:
                    duration = timestamp - activity_timers['Lock']
                    daily_activity[date]['Lock'] += duration
                    activity_timers['Lock'] = None
            elif category in ['Outlook Event', 'Teams Event', 'WinWord Event', 'PowerPoint Event', 
                              'MicrosoftEdge Event', 'Excel Event', 'Chrome Event', 'ESENT Event']:
                if activity_timers[category] is None:
                    activity_timers[category] = timestamp
                else:
                    duration = timestamp - activity_timers[category]
                    daily_activity[date][category] += duration
                    activity_timers[category] = None
        
        # Cap total hours to 7 hours per day and exclude idle time
        activity_hours = {}
        for date, activities in daily_activity.items():
            total_active_time = sum(activity.total_seconds() for activity in activities.values()) / 3600
            if total_active_time > 7:
                factor = 7 / total_active_time
                for category in activities:
                    activities[category] *= factor
            activity_hours[date] = {cat: min(activity.total_seconds() / 3600, 7) for cat, activity in activities.items()}
        
        activity_summary[user] = activity_hours
    
    return activity_summary

# Function to display activity
def display_activity(activity_summary):
    for user, daily_activity in activity_summary.items():
        dates = list(daily_activity.keys())
        categories = list(daily_activity[dates[0]].keys())  # Get categories from the first date
        
        # Prepare data for plotting
        data = {cat: [] for cat in categories}
        for date in dates:
            for cat in categories:
                data[cat].append(daily_activity[date][cat] if cat in daily_activity[date] else 0)
        
        # Plotting the stacked bar chart
        plt.figure(figsize=(12, 6))
        x_indexes = range(len(dates))
        
        bottom = [0] * len(dates)
        for cat in categories:
            plt.bar(x_indexes, [data[cat][i] for i in range(len(dates))], label=cat, bottom=bottom)
            bottom = [bottom[i] + data[cat][i] for i in range(len(dates))]
        
        plt.xlabel('Date')
        plt.ylabel('Hours')
        plt.title(f'Stacked User Activity for {user}')
        plt.xticks(x_indexes, dates, rotation=45)
        plt.legend()
        plt.tight_layout()
        plt.show()

# Function to show loading dialog
def show_loading_dialog():
    loading_dialog = tk.Toplevel(root)
    loading_dialog.geometry("300x100")
    loading_dialog.title("Loading")
    label = ctk.CTkLabel(loading_dialog, text="Loading, please wait...", font=("Helvetica", 14))
    label.pack(expand=True)

    return loading_dialog

# Function to process files in a background thread
def process_files_in_background(files):
    global loading_dialog

    # Show loading dialog
    loading_dialog = show_loading_dialog()

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
    finally:
        # Hide loading label after processing
        loading_dialog.destroy()

# Function to open files and process logs
def open_files():
    files = filedialog.askopenfilenames(filetypes=[("EVTX files", "*.evtx")])

    if len(files) > 3:
        messagebox.showerror("Error", "Please select a maximum of 3 files.")
        return

    if files:
        filenames = [os.path.basename(file) for file in files]  # Extract file names
        file_label.configure(text="Selected files: " + "  ,  ".join(filenames))
        root.update_idletasks()  # Update the GUI immediately
        # Start background processing
        threading.Thread(target=process_files_in_background, args=(files,)).start()


# Setting up the GUI
def cancel_action():
    root.destroy()

# Initialize the main window
root = ctk.CTk()
root.geometry("800x500")
ctk.set_appearance_mode("light")
root.title("Activity Analyser")

# Create and place widgets
title_label = ctk.CTkLabel(root, text="User Activity Analyser", font=("Helvetica", 20))
title_label.pack(pady=50)

# Create a canvas for the dashed frame
canvas_frame = ctk.CTkFrame(root, width=500, height=150)
canvas_frame.pack(pady=10)
canvas = Canvas(canvas_frame, width=500, height=270, highlightthickness=1)
canvas.pack()
canvas.create_rectangle(5, 5, 495, 265, outline="#A9A9A9", dash=(5, 5))  # Dashed border

# Upload icon
upload_icon = PhotoImage(file="./icons/upload_file_icon.png")  # Ensure this file is in the same directory
upload_icon = upload_icon.subsample(6, 6)

# Place the upload icon and text in the canvas
canvas.create_image(250, 70, image=upload_icon)

canvas.create_text(250, 150, text="Choose files to analyse\n\nSupported formats: EVTx", fill="gray", font=("Helvetica", 12), justify="center")

# Use a window in the canvas to place the button
upload_button = ctk.CTkButton(canvas, text="Choose file", command=open_files)
canvas.create_window(250, 230, window=upload_button, anchor="center")

file_label = ctk.CTkLabel(root, text="")
file_label.pack(pady=10)

root.mainloop()
