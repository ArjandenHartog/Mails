import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from nietwebshop import process_orders
import os

class MailProcessorGUI:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("LampenTotaal Mail Verwerker")
        self.window.geometry("600x400")
        self.window.configure(bg="#f0f0f0")
        
        # Style configuration
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('Custom.TFrame', background='#f0f0f0')
        style.configure('Custom.TButton', 
                       padding=10, 
                       font=('Helvetica', 10),
                       background='#4a90e2')
        style.configure('Title.TLabel', 
                       font=('Helvetica', 16, 'bold'),
                       background='#f0f0f0')
        style.configure('Info.TLabel',
                       font=('Helvetica', 10),
                       background='#f0f0f0')
        
        self.create_widgets()
        
    def create_widgets(self):
        # Main container
        main_frame = ttk.Frame(self.window, style='Custom.TFrame', padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title = ttk.Label(main_frame, 
                         text="LampenTotaal Mail Verwerker",
                         style='Title.TLabel')
        title.pack(pady=(0, 20))
        
        # Description
        description = ttk.Label(main_frame,
                              text="Selecteer het type orders dat u wilt verwerken",
                              style='Info.TLabel')
        description.pack(pady=(0, 20))
        
        # Buttons container
        buttons_frame = ttk.Frame(main_frame, style='Custom.TFrame')
        buttons_frame.pack(pady=20)
        
        # Niet Webshop Button
        niet_webshop_btn = ttk.Button(
            buttons_frame,
            text="Niet Webshop Orders",
            style='Custom.TButton',
            command=lambda: self.process_file('niet_webshop')
        )
        niet_webshop_btn.pack(pady=10, ipadx=20, ipady=10)
        
        # NML Button
        nml_btn = ttk.Button(
            buttons_frame,
            text="NML Orders",
            style='Custom.TButton',
            command=lambda: self.process_file('nml')
        )
        nml_btn.pack(pady=10, ipadx=20, ipady=10)
        
        # Status label
        self.status_label = ttk.Label(main_frame,
                                    text="",
                                    style='Info.TLabel')
        self.status_label.pack(pady=20)
        
    def process_file(self, type_order):
        file_path = filedialog.askopenfilename(
            title=f"Selecteer Excel bestand voor {type_order}",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        
        if file_path:
            try:
                if type_order == 'niet_webshop':
                    process_orders(file_path)
                else:
                    # Here you would call the NML process function
                    # process_nml_orders(file_path)
                    messagebox.showinfo("Info", "NML verwerking nog niet geïmplementeerd")
                    return
                    
                self.status_label.config(
                    text=f"✅ Bestand succesvol verwerkt!\nLaatst verwerkt: {os.path.basename(file_path)}"
                )
                messagebox.showinfo("Succes", "Bestand is verwerkt!")
                
            except Exception as e:
                self.status_label.config(
                    text="❌ Er is een fout opgetreden"
                )
                messagebox.showerror("Fout", f"Er is een fout opgetreden:\n{str(e)}")

    def run(self):
        self.window.mainloop()

if __name__ == "__main__":
    app = MailProcessorGUI()
    app.run()
