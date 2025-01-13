import sys
import os
import comtypes.client
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from docx2pdf import convert
import threading
import webbrowser
import logging
import psutil  # Import the psutil library


# Setup logging
logging.basicConfig(filename='docx2pdf_converter.log', level=logging.INFO, format='%(asctime)s:%(levelname)s:%(message)s')

# Define the StdoutRedirector class
class StdoutRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, string):
        self.text_widget.configure(state='normal')
        self.text_widget.insert(tk.END, string)
        self.text_widget.see(tk.END)
        self.text_widget.configure(state='disabled')

    def flush(self):
        pass

# Function to convert .doc to .docx
def doc_to_docx(doc_path):
    # Normalize the path
    doc_path = os.path.normpath(doc_path)
    
    file_name, file_extension = os.path.splitext(doc_path)
    docx_path = file_name + ".docx"

    if not os.path.exists(docx_path):  # Check if .docx version already exists
        # Check if the .doc file exists
        if not os.path.exists(doc_path):
            logging.error(f"File not found: {doc_path}")
            return None

        try:
            word = comtypes.client.CreateObject('Word.Application')
            word.Visible = False  # Hide the Word application
            doc = word.Documents.Open(doc_path)
            doc.SaveAs(docx_path, FileFormat=16)  # FileFormat=16 for docx
            doc.Close()
            word.Quit()
        except Exception as e:
            logging.error(f"Failed to convert {doc_path} to .docx: {e}")
            return None
    
    return docx_path

# Function to convert all .docx to .pdf in a given directory
def convert_all_docx_to_pdf(input_directory, status_text_widget):
    pdf_directory = os.path.join(input_directory, 'PDFs')
    
    # Check if the input_directory exists
    if not os.path.exists(input_directory):
        update_status(status_text_widget, f"Error: Input directory does not exist: {input_directory}")
        return
    
    # Create a 'PDFs' subfolder if it doesn't exist
    if not os.path.exists(pdf_directory):
        try:
            os.mkdir(pdf_directory)
        except Exception as e:
            update_status(status_text_widget, f"Error creating PDF directory: {str(e)}")
            logging.error(f"Error creating PDF directory: {str(e)}")
            return

    # Process .doc files first, convert to .docx
    doc_files = [f for f in os.listdir(input_directory) if f.lower().endswith('.doc')]
    for doc_file in doc_files:
        doc_path = os.path.join(input_directory, doc_file)
        docx_path = doc_to_docx(doc_path)
        if docx_path is None:
            continue

    # Convert .docx to .pdf
    docx_files = [f for f in os.listdir(input_directory) if f.lower().endswith('.docx')]
    for docx_file in docx_files:
        docx_file_path = os.path.join(input_directory, docx_file)
        pdf_file_path = os.path.join(pdf_directory, os.path.splitext(docx_file)[0] + '.pdf')
        try:
            update_status(status_text_widget, f"Converting: {docx_file}")
            convert(docx_file_path, pdf_file_path)
            update_status(status_text_widget, f"Converted: {docx_file}")
        except Exception as e:
            error_msg = f"Error converting {docx_file}: {str(e)}"
            update_status(status_text_widget, error_msg)
            logging.error(error_msg)

    update_status(status_text_widget, "All files converted successfully!")
    webbrowser.open(pdf_directory)

    # Clean up temporary .docx files that were created from .doc files
    for doc_file in doc_files:
        docx_path = os.path.join(input_directory, os.path.splitext(doc_file)[0] + '.docx')
        if os.path.exists(docx_path):
            os.remove(docx_path)
            update_status(status_text_widget, f"Removed temporary file: {os.path.basename(docx_path)}")

def is_word_running():
    """
    Check if Microsoft Word (winword.exe) is running.
    """
    for process in psutil.process_iter(['pid', 'name']):
        if process.info['name'] == 'WINWORD.EXE':
            return True
    return False

def wait_for_word_to_close(status_text_widget):
    """
    Wait for Microsoft Word to close before proceeding.
    """
    while is_word_running():
        update_status(status_text_widget, "Waiting for Microsoft Word to close...")
        time.sleep(1)  # Check every 1 second
    update_status(status_text_widget, "Microsoft Word closed. Proceeding with conversion...")

def select_folder_and_convert(status_text_widget):
    if is_word_running():
        messagebox.showwarning("Microsoft Word Running", "Please close Microsoft Word before proceeding.")
        threading.Thread(target=wait_for_word_to_close, args=(status_text_widget,), daemon=True).start()
        return
    
# Function to handle folder selection and start the conversion
def select_folder_and_convert(status_text_widget):
    input_directory = filedialog.askdirectory()
    if input_directory:
        update_status(status_text_widget, "Conversion started...")
        threading.Thread(target=convert_all_docx_to_pdf, args=(input_directory, status_text_widget), daemon=True).start()
    else:
        update_status(status_text_widget, "No folder was selected.")

# Function to update the status text widget
def update_status(status_text_widget, message):
    status_text_widget.configure(state='normal')
    status_text_widget.insert(tk.END, message + "\n")
    status_text_widget.see(tk.END)
    status_text_widget.configure(state='disabled')

def main():
    # Create the main window
    app = tk.Tk()
    app.title("Bulk DOCX to PDF Converter")
    app.geometry('500x300')

    status_text_widget = scrolledtext.ScrolledText(app, wrap=tk.WORD, state='disabled', height=10)
    status_text_widget.pack(pady=10)

    button_convert_folder = tk.Button(app, text="Select Folder and Convert DOCX Files", command=lambda: select_folder_and_convert(status_text_widget))
    button_convert_folder.pack(pady=5)

    explanation_text = tk.Label(app, text="The converted PDFs will be saved in a subfolder named 'PDFs' within the selected folder.\nAfter conversion, the 'PDFs' folder will be opened automatically.")
    explanation_text.pack(pady=10)

    # Redirect stdout and stderr to the status_text_widget
    sys.stdout = StdoutRedirector(status_text_widget)
    sys.stderr = StdoutRedirector(status_text_widget)

    app.mainloop()

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        logging.error(f"Unhandled exception: {e}")
        messagebox.showerror("Error", f"An unexpected error occurred: {e}")
