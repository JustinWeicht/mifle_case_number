import os
import threading
import subprocess
import tkinter as tk
from ttkthemes import ThemedTk # pip install ttkthemes
from tkinter import ttk, Menu
from tkinter.filedialog import askopenfilename
from main import main

class GUI:
    def __init__(self, master):
        self.master = master
        self.master.title("Mifile Case Number")

        # Handle window close event
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Create a menu bar
        menu_bar = Menu(self.master)
        self.master.config(menu=menu_bar)

        # Create a Help menu
        help_menu = Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="Help", menu=help_menu)

        # Add an Instructions option to the menu
        help_menu.add_command(label="Instructions", command=self.show_instructions)

         # Create labels and entry widgets
        self.input_file_label = tk.Label(master, text="Input File:")
        self.input_file_label.grid(row=0, column=0, sticky=tk.E, padx=5, pady=(5,0))

        self.input_file_var = tk.StringVar()
        self.input_file_entry = ttk.Entry(master, textvariable=self.input_file_var, width=30)
        self.input_file_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=(5,0))

        self.browse_input_button = ttk.Button(master, text="Browse", command=self.browse_input)
        self.browse_input_button.grid(row=0, column=2, padx=10, pady=(5,0), sticky=tk.W)

        # Create label for processing status
        self.processing_status_var = tk.StringVar()
        self.processing_status_label = tk.Label(master, textvariable=self.processing_status_var, fg="grey")
        self.processing_status_label.grid(row=2, columnspan=2, padx=10, pady=10)

        # Create submit button
        self.generate_button = ttk.Button(master, text="Submit", command=self.generate_excel)
        self.generate_button.grid(row=2, columnspan=3, padx=10, pady=10, sticky=tk.E)

        # Initialize thread variable
        self.thread = None

    def set_processing_status(self, status):
        self.processing_status_var.set(status)  

    def browse_input(self):
        file_path = askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.input_file_var.set(file_path)   

    def open_file_explorer(self, folder_path):
        os.startfile(folder_path)

    def show_instructions(self):
        # Specify the path to your premade instructions document
        instructions_path = r''

        if os.path.exists(instructions_path):
            self.open_file_explorer(instructions_path)
        else:
            self.set_processing_status("Instructions document not found.")

    def generate_excel(self):
        input_file = self.input_file_var.get()

        if not input_file:
            self.set_processing_status("Please select an input file.")
            return

        try:
            # Run the main function in a separate thread
            self.thread = threading.Thread(target=self.main_threaded, args=(input_file,))
            self.thread.daemon = True
            self.thread.start()

            # Disable the Submit button during processing
            self.generate_button.config(state=tk.DISABLED)

            # Set processing status
            self.set_processing_status("Processing...")

        except Exception as e:
            # Provide user-friendly error message
            print("Error:", e)
            self.set_processing_status("Error while creating Excel file.")
            # Re-enable the Submit button on error
            self.generate_button.config(state=tk.NORMAL)

    def main_threaded(self, input_file):
        try:
            main(input_file)
            # Provide user feedback upon completion
            self.set_processing_status("Finished processing.")
            # Open the output folder
            os.startfile(input_file)
        except Exception as e:
            # Provide user-friendly message when Chrome window is closed
            print("Process stopped:", e)
            self.set_processing_status("Process stopped")
        finally:
            # Re-enable the Submit button after processing
            self.generate_button.config(state=tk.NORMAL)

    def on_closing(self):
        if self.thread and self.thread.is_alive():
            self.thread.join()
        self.master.destroy()

if __name__ == "__main__":
    themed = ThemedTk(theme='plastik')
    app = GUI(themed)
    themed.mainloop()
