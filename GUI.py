import tkinter as tk
from tkinter import filedialog
import pandas as pd
from main import main


def process_excel():
    file_path = filedialog.askopenfilename()
    if file_path:
        try:
            # df = pd.read_excel(file_path)  # Read the Excel file
            # Your processing code here using the data in 'df'
            main(file_path)

            succeeded = tk.Label(root, text="succeeded")
            succeeded.pack()


        except Exception as e:
            output_text.delete('1.0', tk.END)
            output_text.insert(tk.END, f"Error: {e}")



# GUI setup
root = tk.Tk()
root.title("Excel Processor")

load_button = tk.Button(root, text="Load Excel File", command=process_excel)
load_button.pack()

output_text = tk.Text(root, height=10, width=50)
output_text.pack()

root.mainloop()
