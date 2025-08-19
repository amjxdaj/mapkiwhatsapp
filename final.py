import tkinter as tk
from tkinter import filedialog, Label, Text, ttk, messagebox, scrolledtext
import pandas as pd
import pywhatkit as kit
import pyautogui as pag
import time
import threading
import os
import logging
from datetime import datetime

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Define window size
WINDOW_SIZE = "500x700"

# Optimized timing constants
WAIT_TIMES = {
    'image_load': 15,      # Reduced from 25
    'text_only': 10,       # Reduced from 20
    'between_messages': 8,  # Reduced from 15
    'tab_close': 3,        # Reduced from 10
    'error_recovery': 5    # Reduced from 10
}

# Function to validate access key
def validate_access_key(key):
    return key == "mapki"

# Function to log messages to console widget
def log_to_console(message, level="INFO"):
    try:
        if console_text and root:
            timestamp = datetime.now().strftime("%H:%M:%S")
            log_message = f"[{timestamp}] {level}: {message}\n"
            console_text.insert(tk.END, log_message)
            console_text.see(tk.END)
            root.update_idletasks()
        logger.info(message)
    except Exception as e:
        logger.error(f"Console logging error: {e}")
        print(f"[{datetime.now().strftime('%H:%M:%S')}] {level}: {message}")

# Function to validate phone number
def validate_phone_number(number):
    try:
        # Remove any non-digit characters except +
        clean_number = ''.join(filter(str.isdigit, str(number)))
        if len(clean_number) == 10:
            return '+91' + clean_number
        elif len(clean_number) == 12 and clean_number.startswith('91'):
            return '+' + clean_number
        else:
            return None
    except:
        return None

# Function to check if image file exists and is valid
def validate_image_path(path):
    if not path or not os.path.exists(path):
        return False
    valid_extensions = ['.jpg', '.jpeg', '.png', '.gif', '.bmp']
    return any(path.lower().endswith(ext) for ext in valid_extensions)

# Global variables
root = None
file_path = None
image_path = None
stop_sending = False
console_text = None
file_label = None
image_label = None
message_text = None
progress_bar = None
progress_label = None
status_label = None

# Function to open the main GUI
def open_main_gui():
    global root, file_path, image_path, stop_sending, console_text
    global file_label, image_label, message_text, progress_bar, progress_label, status_label
    
    root = tk.Tk()
    root.title("Mapki WhatsApp Tool - Optimized")
    root.geometry(WINDOW_SIZE)

    file_path = None
    image_path = None
    stop_sending = False

    # Title Label
    title_label = Label(root, text="Mapki WhatsApp Tool", font=("Arial", 16, "bold"))
    title_label.pack(pady=10)

    # Select Excel File Section
    select_button = tk.Button(root, text="Select Excel File", command=select_file)
    select_button.pack(pady=(10, 2))

    file_label = Label(root, text="No file selected", fg="gray")
    file_label.pack(pady=(2, 5))

    # Select Image File Section
    select_image_button = tk.Button(root, text="Select Image File (Optional)", command=select_image)
    select_image_button.pack(pady=(5, 2))

    image_label = Label(root, text="No image selected", fg="gray")
    image_label.pack(pady=(2, 5))

    # Custom Message Section
    message_label = Label(root, text="Enter your message: (Use {name} for personalization)")
    message_label.pack(pady=(5, 2))

    message_text = Text(root, height=4, width=50)
    message_text.pack(pady=(2, 5))

    # Progress Section
    progress_frame = tk.Frame(root)
    progress_frame.pack(pady=5)

    progress_bar = ttk.Progressbar(progress_frame, orient="horizontal", length=300, mode="determinate")
    progress_bar.pack()
    
    progress_label = Label(progress_frame, text="0/0 messages sent", font=("Arial", 9))
    progress_label.pack(pady=2)

    # Status Label
    status_label = Label(root, text="Ready", font=("Arial", 10), fg="green")
    status_label.pack(pady=5)

    # Control Buttons Section
    button_frame = tk.Frame(root)
    button_frame.pack(pady=10)

    send_button = tk.Button(button_frame, text="Send Messages", width=15, bg="green", fg="white",
                           command=lambda: threading.Thread(target=send_messages, daemon=True).start())
    send_button.grid(row=0, column=0, padx=5)

    stop_button = tk.Button(button_frame, text="Stop Sending", width=15, bg="red", fg="white",
                           command=stop_messages)
    stop_button.grid(row=0, column=1, padx=5)

    reset_button = tk.Button(button_frame, text="Reset", width=15, command=reset_fields)
    reset_button.grid(row=0, column=2, padx=5)

    # Console Log Section
    console_frame = tk.Frame(root)
    console_frame.pack(pady=(10, 5), fill='both', expand=True)

    console_label = Label(console_frame, text="Console Log:", font=("Arial", 10, "bold"))
    console_label.pack(anchor='w')

    console_text = scrolledtext.ScrolledText(console_frame, height=8, width=60, 
                                           font=("Consolas", 9), bg="black", fg="green")
    console_text.pack(fill='both', expand=True)

    # Clear console button
    clear_console_button = tk.Button(root, text="Clear Console", command=clear_console)
    clear_console_button.pack(pady=5)

    log_to_console("Application started successfully")
    root.mainloop()

# Function to clear console
def clear_console():
    console_text.delete(1.0, tk.END)

# Function to create access key window
def access_key_window():
    key_window = tk.Tk()
    key_window.title("Access Key")
    key_window.geometry("400x300")
    key_window.resizable(False, False)

    frame = tk.Frame(key_window)
    frame.pack(expand=True)

    title_label = Label(frame, text="Mapki WhatsApp Tool", font=("Arial", 18, "bold"))
    title_label.pack(pady=(20, 30))

    label = Label(frame, text="Enter Access Key:", font=("Arial", 12))
    label.pack(pady=5)

    access_key_entry = tk.Entry(frame, width=30, font=("Arial", 11), show="*")
    access_key_entry.pack(pady=10)
    access_key_entry.focus()

    def submit_key():
        access_key = access_key_entry.get().strip()
        if validate_access_key(access_key):
            key_window.destroy()
            open_main_gui()
        else:
            messagebox.showerror("Error", "Invalid Access Key. Please try again.")
            access_key_entry.delete(0, tk.END)

    access_key_entry.bind("<Return>", lambda event: submit_key())

    submit_button = tk.Button(frame, text="Submit", command=submit_key, 
                             bg="blue", fg="white", width=15)
    submit_button.pack(pady=20)

    key_window.mainloop()

# Function to select Excel file
def select_file():
    global file_path
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        try:
            # Validate the Excel file
            df = pd.read_excel(file_path)
            required_columns = ['Number']
            if not all(col in df.columns for col in required_columns):
                messagebox.showerror("Error", "Excel file must contain 'Number' column")
                file_path = None
                file_label.config(text="Invalid file format", fg="red")
                return
            
            file_name = os.path.basename(file_path)
            file_label.config(text=f"✓ {file_name} ({len(df)} contacts)", fg="green")
            log_to_console(f"Excel file loaded: {len(df)} contacts found")
        except Exception as e:
            messagebox.showerror("Error", f"Error reading Excel file: {str(e)}")
            file_path = None
            file_label.config(text="Error reading file", fg="red")
    else:
        file_label.config(text="No file selected", fg="gray")

# Function to select an image file
def select_image():
    global image_path
    image_path = filedialog.askopenfilename(filetypes=[("Image files", "*.jpg *.jpeg *.png *.gif *.bmp")])
    if image_path:
        if validate_image_path(image_path):
            image_name = os.path.basename(image_path)
            image_label.config(text=f"✓ {image_name}", fg="green")
            log_to_console(f"Image selected: {image_name}")
        else:
            messagebox.showerror("Error", "Invalid image file")
            image_path = None
            image_label.config(text="Invalid image file", fg="red")
    else:
        image_label.config(text="No image selected", fg="gray")

# Function to handle stop messages
def stop_messages():
    global stop_sending
    stop_sending = True
    status_label.config(text="Stopping...", fg="orange")
    log_to_console("Stop signal sent", "WARNING")

# Function to reset all fields
def reset_fields():
    global file_path, image_path, stop_sending
    stop_sending = True
    file_path = None
    image_path = None
    file_label.config(text="No file selected", fg="gray")
    image_label.config(text="No image selected", fg="gray")
    message_text.delete("1.0", tk.END)
    progress_bar['value'] = 0
    progress_label.config(text="0/0 messages sent")
    status_label.config(text="Ready", fg="green")
    log_to_console("All fields reset")

# Optimized function to send messages
def send_messages():
    global stop_sending, file_path, image_path
    
    stop_sending = False

    if not file_path:
        status_label.config(text="No file selected", fg="red")
        log_to_console("Error: No Excel file selected", "ERROR")
        return

    custom_message = message_text.get("1.0", tk.END).strip()
    if not custom_message:
        messagebox.showerror("Error", "Please enter a message")
        return

    try:
        df = pd.read_excel(file_path)
        total_contacts = len(df)
        
        log_to_console(f"Starting bulk message sending to {total_contacts} contacts")
        
        progress_bar['value'] = 0
        progress_bar['maximum'] = total_contacts
        
        sent_count = 0
        failed_count = 0
        
        status_label.config(text="Sending messages...", fg="blue")

        # Pre-validate image if selected
        current_image_path = image_path  # Store current image path
        if current_image_path and not validate_image_path(current_image_path):
            log_to_console("Error: Invalid image file, switching to text-only mode", "WARNING")
            current_image_path = None

        for index, row in df.iterrows():
            if stop_sending:
                log_to_console("Message sending stopped by user", "WARNING")
                status_label.config(text="Stopped", fg="orange")
                break

            # Validate and format phone number
            raw_number = row.get('Number', '')
            phone_number = validate_phone_number(raw_number)
            
            if not phone_number:
                log_to_console(f"Invalid phone number: {raw_number}", "ERROR")
                failed_count += 1
                continue

            name = row.get('Name', 'Customer')
            personalized_message = custom_message.replace("{name}", name)

            try:
                log_to_console(f"Sending to {phone_number} ({name})")
                
                if current_image_path:
                    # Send image with message
                    kit.sendwhats_image(phone_number, current_image_path, personalized_message)
                    time.sleep(WAIT_TIMES['image_load'])
                    
                    # Try to send with error handling
                    try:
                        pag.press('enter')
                        time.sleep(WAIT_TIMES['tab_close'])
                        pag.hotkey('ctrl', 'w')
                    except Exception as e:
                        log_to_console(f"UI automation error: {str(e)}", "WARNING")
                else:
                    # Send text only with optimized settings
                    kit.sendwhatmsg_instantly(
                        phone_number, 
                        personalized_message, 
                        wait_time=WAIT_TIMES['text_only'], 
                        tab_close=True, 
                        close_time=WAIT_TIMES['tab_close']
                    )

                sent_count += 1
                log_to_console(f"✓ Message sent successfully to {phone_number}")
                
            except Exception as e:
                failed_count += 1
                log_to_console(f"✗ Failed to send to {phone_number}: {str(e)}", "ERROR")
                time.sleep(WAIT_TIMES['error_recovery'])

            # Update progress
            progress_bar['value'] = sent_count + failed_count
            progress_label.config(text=f"{sent_count}/{total_contacts} messages sent ({failed_count} failed)")
            root.update_idletasks()
            
            # Wait between messages (only if not stopped)
            if not stop_sending and (sent_count + failed_count) < total_contacts:
                time.sleep(WAIT_TIMES['between_messages'])

        # Final status update
        if stop_sending:
            status_label.config(text=f"Stopped: {sent_count} sent, {failed_count} failed", fg="orange")
        else:
            status_label.config(text=f"Complete: {sent_count} sent, {failed_count} failed", fg="green")
            
        log_to_console(f"Bulk sending completed. Sent: {sent_count}, Failed: {failed_count}")

    except Exception as e:
        error_msg = f"Error processing file: {str(e)}"
        log_to_console(error_msg, "ERROR")
        status_label.config(text="Error occurred", fg="red")
        messagebox.showerror("Error", error_msg)

# Start the application
if __name__ == "__main__":
    access_key_window()