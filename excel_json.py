import tkinter as tk
from tkinter import filedialog
import pandas as pd
import json


def load_excel_and_convert_to_json():
    # Open a file dialog to select the Excel file
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx")])
    if not file_path:  # If the user cancels, do nothing
        return

    # Read the Excel file
    df = pd.read_excel(file_path)

    # Convert the DataFrame to JSON (adjust this part to fit your specific structure)
    json_data = df.to_json(orient="records")

    # Parse the JSON string so we can pretty-print it
    parsed_json = json.loads(json_data)

    # Define the JSON output file name
    output_file_name = "output.json"

    # Save the JSON to a file
    with open(output_file_name, 'w', encoding='utf-8') as f:
        json.dump(parsed_json, f, ensure_ascii=False, indent=4)

    print(f"JSON saved to {output_file_name}")


def create_gui():
    # Create the main window
    root = tk.Tk()
    root.title("Excel to JSON Converter")

    # Create a frame for the button
    frame = tk.Frame(root)
    frame.pack(pady=20)

    # Add a button to the frame
    convert_button = tk.Button(
        frame, text="Load Excel & Convert to JSON", command=load_excel_and_convert_to_json)
    convert_button.pack()

    root.mainloop()


if __name__ == "__main__":
    create_gui()
