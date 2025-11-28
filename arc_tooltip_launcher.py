import tkinter as tk
from tkinter import PhotoImage, font, ttk
import requests
import zipfile
import os
import sys
import shutil
import threading

# Global variables
root = None
progress_bar = None
progress_label = None
download_thread = None

def center_window(window):
    window.update_idletasks()
    width = window.winfo_width()
    height = window.winfo_height()
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry(f'{width}x{height}+{x}+{y}')

def check_for_update(current_version):
    try:
        response = requests.get("https://ghostworld073.pythonanywhere.com/latest_arc_companion")
        response.raise_for_status()  # Check for HTTP errors
        latest_version = response.json()[0]  # Assuming JSON response instead of text eval
        if latest_version != current_version:
            print(f"New version available: {latest_version}")
            return True
        else:
            print("No new version available.")
            return False
    except requests.RequestException as e:
        print(f"Error checking for updates: {e}")
        return False

def download_update():
    global download_thread
    download_thread = threading.Thread(target=download_update_thread)
    download_thread.start()
    # Start a periodic UI update
    root.after(100, check_download_thread)

def download_update_thread():
    try:
        response = requests.get("https://ghostworld073.pythonanywhere.com/download_latest_arc_companion", stream=True)
        response.raise_for_status()  # Check for HTTP errors
        total_size = int(response.headers.get('content-length', 0))
        block_size = 1024
        downloaded_size = 0

        # Prepare progress bar for the download
        progress_bar['maximum'] = total_size

        with open("arc_companion_update.zip", "wb") as file:
            for data in response.iter_content(block_size):
                file.write(data)
                downloaded_size += len(data)
                # Update progress variables
                mb_downloaded = downloaded_size / (1024 * 1024)
                mb_total = total_size / (1024 * 1024)
                # Schedule UI update
                root.after(0, update_progress_ui, downloaded_size, total_size, mb_downloaded, mb_total)
        print("File downloaded and saved as arc_companion_update.zip")
        # Start extraction in a separate thread
        root.after(0, apply_update)
    except requests.RequestException as e:
        print(f"Error downloading update: {e}")
        root.after(0, launch_application)

def update_progress_ui(downloaded_size, total_size, mb_downloaded, mb_total):
    progress_bar['value'] = downloaded_size
    progress_label.config(text=f"Updating: {mb_downloaded:.2f} MB / {mb_total:.2f} MB")

def check_download_thread():
    if download_thread.is_alive():
        # Reschedule this check
        root.after(100, check_download_thread)
    else:
        # Download thread has finished
        pass  # Do nothing here; apply_update will be called from the download thread

def apply_update():
    extract_thread = threading.Thread(target=apply_update_thread)
    extract_thread.start()
    # Start a periodic check for the extraction thread
    root.after(100, check_extract_thread, extract_thread)

def apply_update_thread():
    try:
        with zipfile.ZipFile('arc_companion_update.zip', 'r') as zip_ref:
            zip_ref.extractall('.')
        print("Update extracted successfully.")
    except zipfile.BadZipFile:
        print("Error extracting update.")
    finally:
        # Proceed to launch the application
        root.after(0, launch_application)

def check_extract_thread(extract_thread):
    if extract_thread.is_alive():
        # Reschedule this check
        root.after(100, check_extract_thread, extract_thread)
    else:
        # Extraction thread has finished
        pass  # Do nothing here; launch_application will be called from the extraction thread

def launch_application():
    try:
        if os.path.isfile('arc_companion.exe'):
            os.system('start arc_companion.exe')  # Use 'start' command for Windows
        else:
            print("Executable not found.")
    except Exception as e:
        print(f"Error launching application: {e}")

    # Exit the updater
    if root:
        root.quit()
        root.destroy()
    sys.exit()

def update_app():
    try:
        with open('arc_companion_version.txt', 'r') as version_file:
            current_version = version_file.read().strip()
        if check_for_update(current_version):
            global root, progress_bar, progress_label
            root = tk.Tk()
            root.title("Updating ARC Companion")
            root.configure(bg='black')
            root.minsize(400, 200)

            frame = tk.Frame(root, bg='black')
            frame.pack(fill='both', expand=True)

            custom_font = font.Font(family="Helvetica", size=14)
            progress_label = tk.Label(frame, text="Updating...", bg='black', fg='white', font=custom_font)
            progress_label.pack(padx=20, pady=30)

            progress_bar = ttk.Progressbar(frame, orient='horizontal', length=300, mode='determinate')
            progress_bar.pack(pady=10)

            center_window(root)

            try:
                icon = PhotoImage(file='Companion.png')
                root.iconphoto(False, icon)
            except tk.TclError:
                print("Icon file not found.")

            # Start the download in a separate thread
            download_update()
            root.mainloop()
        else:
            launch_application()
    except FileNotFoundError:
        print("Version file not found.")
        launch_application()

update_app()
