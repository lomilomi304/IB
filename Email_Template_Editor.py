import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import csv
from pathlib import Path

def main():
    root = tk.Tk()
    root.title("Library Email Generator")
    root.geometry("800x600")
    app = LibraryEmailGenerator(root)
    root.mainloop()

class LibraryEmailGenerator:
    def __init__(self, root):
        self.root = root
        self.setup_gui()
        self.profiles = {}
        self.profiles_case_map = {}
        self.csv_filename = ""
        self.current_profile_index = 0
        self.profile_list = []
        self.email_template = """UPEI Off-Campus Student Credentials - IB {csv_filename} class of 2026

Hi {First Name},

Welcome to the Robertson Library! Your UPEI Library profile is now active, which means that you now have access to all of our online resources even while you're outside of our campus. 

You will be receiving instructions on how to use these credentials and access our physical and online resources during your Robertson Library Orientation Session on Monday Nov 4th.

You'll also receive your Robertson Library library card at the Library Orientation Session, which allows you to borrow many physical resources that we offer at the Library. Keep in mind that you must have this card with you in order for us to loan anything to you.

Below are your username and password to access the library resources from off campus.
username: {Username}
password: {Password}

You'll have to login through the following link, whenever you are trying to access our online resources:
https://proxy.library.upei.ca/public/proxylogin.htm

If you have any questions, please feel free to ask myself, your IB coordinator, or email your librarian, Katelyn Browne, at krbrowne@upei.ca. You can also always speak to someone at the Robertson Library Service Desk.

Thank you,
Spencer"""

    def setup_gui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # File selection
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(file_frame, text="Select CSV File", command=self.browse_file).pack(side=tk.LEFT, padx=5)
        self.file_label = ttk.Label(file_frame, text="No file selected")
        self.file_label.pack(side=tk.LEFT, padx=5)

        # Navigation and search frame
        nav_search_frame = ttk.Frame(main_frame)
        nav_search_frame.pack(fill=tk.X, pady=5)

        # Navigation buttons
        nav_frame = ttk.Frame(nav_search_frame)
        nav_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Button(nav_frame, text="Next Profile", command=self.next_profile).pack(side=tk.LEFT, padx=5)
        self.profile_counter = ttk.Label(nav_frame, text="Profile: 0/0")
        self.profile_counter.pack(side=tk.LEFT, padx=5)

        # Search
        search_frame = ttk.Frame(nav_search_frame)
        search_frame.pack(side=tk.RIGHT, fill=tk.X)
        
        ttk.Label(search_frame, text="Search Last Name:").pack(side=tk.LEFT, padx=5)
        self.last_name_var = tk.StringVar()
        self.name_entry = ttk.Entry(search_frame, textvariable=self.last_name_var)
        self.name_entry.pack(side=tk.LEFT, padx=5)
        self.name_entry.bind('<Return>', lambda e: self.search_profile())
        
        ttk.Button(search_frame, text="Search", command=self.search_profile).pack(side=tk.LEFT, padx=5)

        # Email display
        self.email_text = tk.Text(main_frame, wrap=tk.WORD, height=20)
        self.email_text.pack(fill=tk.BOTH, expand=True, pady=5)

        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(button_frame, text="Copy to Clipboard", command=self.copy_to_clipboard).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Clear Email", command=self.clear_email).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Edit Template", command=self.edit_template).pack(side=tk.LEFT, padx=5)

    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Select CSV File",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if filename:
            self.load_profiles(filename)
            self.file_label.config(text=Path(filename).name)
            self.generate_email_for_current_profile()

    def load_profiles(self, filename):
        try:
            with open(filename, 'r') as file:
                reader = csv.DictReader(file)
                self.profiles = {}
                self.profiles_case_map = {}
                self.profile_list = []
                for row in reader:
                    self.profiles[row['Last Name']] = {
                        'First Name': row['First Name'],
                        'Last Name': row['Last Name'],
                        'Email': row['Email'],
                        'Username': row['Username'],
                        'Password': row['Password']
                    }
                    self.profiles_case_map[row['Last Name'].lower()] = row['Last Name']
                    self.profile_list.append(row['Last Name'])
                self.csv_filename = Path(filename).stem
                self.current_profile_index = 0
                self.update_profile_counter()
                messagebox.showinfo("Success", f"Loaded {len(self.profiles)} profiles")
        except Exception as e:
            messagebox.showerror("Error", f"Error loading file: {str(e)}")

    def update_profile_counter(self):
        total = len(self.profile_list)
        current = self.current_profile_index + 1 if self.profile_list else 0
        self.profile_counter.config(text=f"Profile: {current}/{total}")

    def next_profile(self):
        if not self.profile_list:
            messagebox.showwarning("Warning", "Please load a CSV file first")
            return
        
        self.current_profile_index = (self.current_profile_index + 1) % len(self.profile_list)
        self.update_profile_counter()
        self.generate_email_for_current_profile()

    def generate_email_for_current_profile(self):
        if not self.profile_list:
            return
        
        current_last_name = self.profile_list[self.current_profile_index]
        profile = self.profiles[current_last_name]
        self.generate_email_content(profile)

    def search_profile(self):
        if not self.profiles:
            messagebox.showwarning("Warning", "Please load a CSV file first")
            return
        
        last_name = self.last_name_var.get().lower()
        if last_name in self.profiles_case_map:
            original_name = self.profiles_case_map[last_name]
            profile = self.profiles[original_name]
            self.generate_email_content(profile)
        else:
            messagebox.showwarning("Not Found", 
                f"No profile found for: {self.last_name_var.get()}\nAvailable names: {', '.join(sorted(self.profiles.keys()))}")

    def generate_email_content(self, profile):
        email_text = profile['Email'] + "\n\n" + self.email_template.format(
            csv_filename=self.csv_filename,
            **profile
        )
        self.email_text.delete(1.0, tk.END)
        self.email_text.insert(1.0, email_text)

    def copy_to_clipboard(self):
        email_content = self.email_text.get(1.0, tk.END).strip()
        if email_content:
            self.root.clipboard_clear()
            self.root.clipboard_append(email_content)
            messagebox.showinfo("Success", "Copied to clipboard!")

    def clear_email(self):
        self.email_text.delete(1.0, tk.END)
        self.last_name_var.set("")

    def edit_template(self):
        template_editor = TemplateEditor(self.root, self.email_template)
        self.root.wait_window(template_editor.top)
        if template_editor.result:
            self.email_template = template_editor.result
            messagebox.showinfo("Success", "Email template updated successfully!")
            if self.profile_list:
                self.generate_email_for_current_profile()

class TemplateEditor:
    def __init__(self, parent, template):
        self.top = tk.Toplevel(parent)
        self.top.title("Edit Email Template")
        self.top.geometry("600x400")
        self.result = None

        self.text_widget = tk.Text(self.top, wrap=tk.WORD)
        self.text_widget.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)
        self.text_widget.insert(tk.END, template)

        button_frame = ttk.Frame(self.top)
        button_frame.pack(pady=10)

        ttk.Button(button_frame, text="Save", command=self.save).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.cancel).pack(side=tk.LEFT, padx=5)

    def save(self):
        self.result = self.text_widget.get(1.0, tk.END).strip()
        self.top.destroy()

    def cancel(self):
        self.top.destroy()

if __name__ == "__main__":
    main()