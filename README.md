Mapki WhatsApp Tool
A GUI-based Python application for sending bulk WhatsApp messages (with optional image attachments) using contact data from Excel files.
Prerequisites

Python 3.6 or higher
Tkinter (usually included with Python, but may require python3-tk on some Linux systems, e.g., sudo apt install python3-tk on Debian-based systems)

Required Packages

pandas - For reading Excel files
pywhatkit - For sending WhatsApp messages
pyautogui - For UI automation
openpyxl - For reading .xlsx Excel files
xlrd - For reading .xls Excel files

Installation

Clone or download this repository.
Install the required packages using pip:pip install pandas pywhatkit pyautogui openpyxl xlrd

Alternatively, Install packages via requirements.txt file:pip install -r requirements.txt


How to Run

Ensure all dependencies are installed.
Run the script:python final.py

Use the GUI to:
Select an Excel file (.xlsx or .xls) with a Number column (and optional Name column).
Optionally select an image file (.jpg, .jpeg, .png, .gif, or .bmp).
Enter a message (use {name} for personalization).
Click "Send Messages" to start sending.



Notes

Ensure WhatsApp Web is accessible in your browser.
Make the browser default and Whatsapp Web Logged In.
The script uses pyautogui for browser interactions, so avoid using your mouse/keyboard during message sending.
Stop sending or reset fields using the respective buttons in the GUI.
