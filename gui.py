#create an GUI for docs_handler_v2.py using Tkinter library
import tkinter as tk
import docs_handler_v2 as dh2
from tkinter import messagebox
from tkinter import filedialog

class DocsHandlerGUI:
    #create window for the GUI with title "document handler", size 800x600 and resizable, 
    #and create a frame for the GUI, and set the frame to fill the window
    #create a label for the title of the GUI, and set the font size to 20
    #use the base_dir and pages defined in docs_handler_v2.py to create the GUI
    #create a button to create and combine the document, and set the command to create the document
    
    def __init__(self, root):
        self.root = root
        self.root.title("Document Handler")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        self.frame = tk.Frame(self.root)
        self.frame.pack(fill=tk.BOTH, expand=True)
        
        self.title_label = tk.Label(self.frame, text="Document Handler", font=("Arial", 20))
        self.title_label.pack()
        
        self.create_button = tk.Button(self.frame, text="Create and Combine Document", command=self.create_document)
        self.create_button.pack()

    def create_document(self):
        base_dir = filedialog.askdirectory()
        if base_dir:
            try:
                dh2.create_doc(base_dir)
                messagebox.showinfo("Success", "Document created and combined successfully")

            except TypeError as e:
                messagebox.showerror("Error", f"Type error occurred: {e}. Please ensure all inputs are correct.")
                
            except Exception as e:
                messagebox.showerror("Error", f"An unexpected error occurred: {e}")

    
if __name__ == "__main__":
    root = tk.Tk()
    app = DocsHandlerGUI(root)
    root.mainloop()