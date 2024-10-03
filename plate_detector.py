import tkinter as tk
from tkinter import messagebox, simpledialog
import pandas as pd
import os
import platform
import cv2
import pytesseract
import tkinter.font as tkFont
import threading

# Uncomment and set the path if Tesseract-OCR is not in your PATH
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Function to create or update an Excel file
def create_or_update_excel():
    file_path = "data.xlsx"

    if os.path.exists(file_path):
        df = pd.read_excel(file_path, engine='openpyxl')
    else:
        df = pd.DataFrame(columns=['Name', 'Number Plate', 'City'])

    name = simpledialog.askstring("Input", "Enter Name:")
    number_plate = simpledialog.askstring("Input", "Enter Number Plate:")
    city = simpledialog.askstring("Input", "Enter City:")

    if name and number_plate and city:
        new_data = pd.DataFrame({'Name': [name], 'Number Plate': [number_plate], 'City': [city]})
        df = pd.concat([df, new_data], ignore_index=True)

        try:
            df.to_excel(file_path, index=False, engine='openpyxl')
            messagebox.showinfo("Success", f"Excel file '{file_path}' updated successfully!")
            open_file(file_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update file: {e}")
    else:
        messagebox.showwarning("Input Error", "All fields are required!")

def open_file(file_path):
    try:
        if platform.system() == 'Windows':
            os.startfile(file_path)
        elif platform.system() == 'Darwin':
            os.system(f"open {file_path}")
        else:
            os.system(f"xdg-open {file_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to open file: {e}")

# Function to remove data (open Excel file)
def Open_Excel_File():
    open_file("data.xlsx")

# Function to detect vehicles and recognize text
saved_letter = "0123456789ZXCVBNMASDFGHJKLQWERTYUIOP"

def detect_plate(frame, n_plate_detector, df):
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    detections = n_plate_detector.detectMultiScale(gray, scaleFactor=1.05, minNeighbors=7)

    if len(detections) == 0:
        return  # Exit if no plates are detected

    for (x, y, w, h) in detections:
        number_plate = gray[y:y + h, x:x + w]
        number_plate = cv2.threshold(number_plate, 100, 255, cv2.THRESH_BINARY)[1]

        try:
            extracted_text = pytesseract.image_to_string(number_plate, config='--psm 8')
        except Exception as e:
            messagebox.showerror("OCR Error", f"Error during text extraction: {e}")
            continue

        plate = ''.join([char for char in extracted_text if char in saved_letter]).strip()

        if plate:
            cv2.rectangle(frame, (x, y), (x + w, y + h), 2)
            cv2.putText(frame, plate, (x, y - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (255, 0, 0), 2)

            match = df[df['Number Plate'].str.contains(plate, na=False, case=False)]
            if not match.empty:
                for _, row in match.iterrows():
                    messagebox.showinfo("Match Found", f"Name: {row['Name']}\nNumber Plate: {row['Number Plate']}")
                    print(f"Match Found: Name: {row['Name']}, Number Plate: {row['Number Plate']}")
            else:
                cv2.putText(frame, "No Match Found", (x, y + h + 20), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 0, 255), 2)

def capture_frames(n_plate_detector, df):
    camera = cv2.VideoCapture(0)
    if not camera.isOpened():
        messagebox.showerror("Camera Error", "Could not open camera.")
        return

    try:
        while True:
            ret, frame = camera.read()
            if not ret:
                messagebox.showerror("Camera Error", "Error reading frame.")
                break

            frame = cv2.resize(frame, (640, 480))
            detect_plate(frame, n_plate_detector, df)

            cv2.imshow("Number Plate Detection", frame)

            if cv2.waitKey(1) & 0xFF == ord('q'):
                break
    finally:
        camera.release()
        cv2.destroyAllWindows()

def scan_vehicle_and_text():
    n_plate_detector_path = "haarcascade_russian_plate_number.xml"
    if not os.path.exists(n_plate_detector_path):
        messagebox.showerror("File Error", "Haar Cascade file not found.")
        return

    n_plate_detector = cv2.CascadeClassifier(n_plate_detector_path)
    if n_plate_detector.empty():
        messagebox.showerror("Error", "Could not load Haar Cascade classifier.")
        return

    excel_file_path = "data.xlsx"
    if os.path.exists(excel_file_path):
        df = pd.read_excel(excel_file_path, engine='openpyxl')
    else:
        messagebox.showerror("File Error", "Excel file not found.")
        return

    # Start frame capture in a separate thread
    threading.Thread(target=capture_frames, args=(n_plate_detector, df), daemon=True).start()

# Create the main window
root = tk.Tk()
root.title("Vehicle Detection and Text Recognition")
root.geometry("600x400")
root.configure(bg="#f0f0f0")  # Light gray background

# Create a frame for the heading
heading_frame = tk.Frame(root, bg="#4a90e2")
heading_frame.pack(pady=20)

# Define font for the heading (bold)
heading_font = tkFont.Font(size=18, weight="bold")

# Add a label for the heading
heading_label = tk.Label(heading_frame, text="Vehicle Detection", font=heading_font, fg="black")
heading_label.pack()

# Create a frame for buttons
button_frame = tk.Frame(root, bg="#f0f0f0")
button_frame.pack(pady=20)

# Create buttons with enhanced styles
button_style = {'bg': '#4caf50', 'fg': 'white', 'font': ('Arial', 12), 'activebackground': '#388e3c'}

scan_button = tk.Button(button_frame, text="Scan Vehicle", command=scan_vehicle_and_text, **button_style)
scan_button.pack(side=tk.LEFT, padx=10)

excel_button = tk.Button(button_frame, text="Register New Vehicle", command=create_or_update_excel, **button_style)
excel_button.pack(side=tk.LEFT, padx=10)

remove_data_button = tk.Button(button_frame, text="Open File", command=Open_Excel_File, **button_style)
remove_data_button.pack(side=tk.LEFT, padx=10)

# Add a footer frame
footer_frame = tk.Frame(root, bg="#4a90e2")
footer_frame.pack(side=tk.BOTTOM, fill=tk.X)

# Add footer text
footer_label = tk.Label(footer_frame, text="© 2024 Vehicle Detection System", bg="#4a90e2", fg="white")
footer_label.pack(pady=5)

def on_closing():
    root.destroy()


root.protocol("WM_DELETE_WINDOW", on_closing)

# Run the application
root.mainloop()
