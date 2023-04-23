# Imports
import json
import logging
import os
import sys
import tkinter as tk
import tkinter.filedialog as filedialog
import tkinter.messagebox as messagebox

from pyunpack import Archive


# Driver class
class MidnightSunsMM:
    def __init__(self, master):
        self.master = master
        master.title("Simple Midnight Suns Mod Manager")

        self.setup_logging()

        # Create and position the setup folders button
        self.setup_folders_button = tk.Button(master, text="Setup Folders", command=self.setup_folders)
        self.setup_folders_button.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)

        # Create and position the source folder selection button
        self.source_folder_button = tk.Button(master, text="Select Local 'mods' Folder",
                                              command=self.choose_source_folder)
        self.source_folder_button.grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)

        # Create and position the destination folder selection button
        self.destination_folder_button = tk.Button(master, text="Select '~mods' Folder",
                                                   command=self.choose_destination_folder)
        self.destination_folder_button.grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)

        # Create and position the file listbox label
        self.file_listbox_label = tk.Label(master, text="Mod Zips")
        self.file_listbox_label.grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)

        # Create and position the file listbox
        self.file_listbox = tk.Listbox(master, selectmode=tk.MULTIPLE)
        self.file_listbox.grid(row=4, column=0, padx=5, pady=5, sticky=tk.W)

        # Create and position the unzip button
        self.unzip_button = tk.Button(master, text="Unzip Selected Mod/s", command=self.unzip_mods)
        self.unzip_button.grid(row=5, column=0, padx=5, pady=5, sticky=tk.W)

        # Create and position the source folder label
        self.source_folder_label = tk.Label(master, text="No mod folder selected.")
        self.source_folder_label.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)

        # Create and position the destination folder label
        self.destination_folder_label = tk.Label(master, text="No ~mods folder selected.")
        self.destination_folder_label.grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)

        # Create and position the mod listbox label
        self.mod_listbox_label = tk.Label(self.master, text="Installed Mods")
        self.mod_listbox_label.grid(row=3, column=1, padx=5, pady=5, sticky=tk.W)

        # Create and position the mod listbox
        self.mod_listbox = tk.Listbox(self.master, selectmode=tk.MULTIPLE)
        self.mod_listbox.grid(row=4, column=1, padx=5, pady=5, sticky=tk.W)

        # Create and position the delete mods button
        self.delete_mods_button = tk.Button(master, text="Delete Selected Mods", command=self.delete_selected_mods)
        self.delete_mods_button.grid(row=5, column=1, padx=5, pady=5, sticky=tk.W)

        # Create the source and destination folder variables
        self.source_folder = ""
        self.destination_folder = ""

        # Create the selected mods variable
        self.selected_mods = []

        # Load the configuration from the JSON file
        self.config = self.load_config()

        # Set the source and destination folders from the configuration
        self.source_folder = self.config.get("source_folder", "")
        self.destination_folder = self.config.get("destination_folder", "")

        # Update the source and destination folder labels
        self.source_folder_label.config(text=self.source_folder)
        self.destination_folder_label.config(text=self.destination_folder)

        # If the source and destination folders have been previously set, populate the file listbox
        if self.source_folder and self.destination_folder:
            self.populate_file_listbox()
            self.populate_mod_list_box()

        # Bind the close event to the save_config method
        master.protocol("WM_DELETE_WINDOW", self.on_closing)

    def setup_folders(self):
        if hasattr(sys, 'frozen'):
            local_directory = os.path.dirname(sys.executable)
        else:
            local_directory = os.path.dirname(os.path.realpath(__file__))

        mods_folder = os.path.join(local_directory, "mods")

        if not os.path.exists(mods_folder):
            os.makedirs(mods_folder)

        self.source_folder = mods_folder
        self.source_folder_label.config(text=self.source_folder)
        self.populate_file_listbox()

        # Populate the file listbox with the .zip files in the source folder
        self.populate_file_listbox()

        # Prompt the user to select their "Midnight Suns" install directory
        install_directory = filedialog.askdirectory(title="Select your Midnight Suns install directory")

        if not install_directory:
            return

        # Create a "~mods" folder in the "MidnightSuns\Content\Paks" subdirectory
        paks_directory = os.path.join(install_directory, "MidnightSuns", "Content", "Paks")
        mods_directory = os.path.join(paks_directory, "~mods")

        if not os.path.exists(mods_directory):
            os.makedirs(mods_directory)

        # Set the destination folder to the newly created "~mods" folder
        self.destination_folder = mods_directory

        # Update the destination folder label
        self.destination_folder_label.config(text=self.destination_folder)

        # Refresh the installed mods list
        self.populate_mod_list_box()

        # Save the updated configuration
        self.save_config()

    def populate_file_listbox(self):
        # Clear the file listbox
        self.clear_file_listbox()

        # Get a list of all the files in the source folder
        file_list = os.listdir(self.source_folder)

        # Loop through each file in the source folder
        for file_name in file_list:
            if file_name.endswith('.zip') or file_name.endswith('.rar') or file_name.endswith('.7z'):
                # Add the file to the file listbox
                self.file_listbox.insert(tk.END, file_name)

        # Adjust the width of the file listbox
        self.adjust_listbox_width(self.file_listbox)

    def populate_mod_list_box(self):
        # Get a list of all the .pak files in the destination folder
        mod_files = [f for f in os.listdir(self.destination_folder) if f.endswith('.pak')]

        # Add each mod file to the mod listbox
        for mod_file in mod_files:
            self.mod_listbox.insert(tk.END, mod_file)

        # Adjust the width of the mod listbox
        self.adjust_listbox_width(self.mod_listbox)

    def choose_source_folder(self):
        # Open a file dialog to choose the source folder
        chosen_folder = filedialog.askdirectory()

        # If the user has canceled the file dialog, return without making changes
        if not chosen_folder:
            # print("No mod folder selected.")
            return

        self.source_folder = chosen_folder

        # Update the source folder label
        self.source_folder_label.config(text=self.source_folder)

        # Clear the file listbox
        self.file_listbox.delete(0, tk.END)

        # Populate the file listbox with the .zip files in the source folder
        self.populate_file_listbox()

        # Save the updated configuration
        self.save_config()

    def choose_destination_folder(self):
        # Open a file dialog to choose the destination folder
        chosen_folder = filedialog.askdirectory()

        # If the user has canceled the file dialog, return without making changes
        if not chosen_folder:
            # print("No ~mods folder selected.")
            return

        self.destination_folder = chosen_folder

        # Update the destination folder label
        self.destination_folder_label.config(text=self.destination_folder)

        # Refresh the installed mods list
        self.populate_mod_list_box()

        # Save the updated configuration
        self.save_config()

    def unzip_mods(self):
        # Make sure both a source folder and a destination folder have been selected
        if not self.source_folder or not self.destination_folder:
            messagebox.showerror("Error", "Please select both a source folder and a destination folder.")
            return

        # # Check if the destination folder contains any mods
        # if self.selected_mods:
        #     self.delete_selected_mods()

        # Get the list of selected files in the file listbox
        selected_mods = [self.file_listbox.get(i) for i in self.file_listbox.curselection()]

        # Loop through each selected file
        for file_name in selected_mods:
            # Get the full path of the archive file
            archive_file_path = os.path.join(self.source_folder, file_name)

            try:
                # Extract all the contents of the archive file into the destination folder
                Archive(archive_file_path).extractall(self.destination_folder)
            except Exception as e:
                # print(f"An error occurred while extracting {file_name}: {e}")
                logging.error(f"An error occurred while extracting {file_name}: {e}")
                messagebox.showerror("Error", f"An error occurred while extracting {file_name}: {str(e)}")
                continue

        # Show a success message
        messagebox.showinfo("Success", "Selected mods have been unzipped to the destination folder.")

        # Refresh the installed mods list
        self.populate_mod_list_box()

    def delete_selected_mods(self):
        # Get the list of selected mods in the mod listbox
        self.selected_mods = [self.mod_listbox.get(i) for i in self.mod_listbox.curselection()]

        # Loop through each selected mod
        for mod_file in self.selected_mods:
            # Get the full path of the mod file
            mod_file_path = os.path.join(self.destination_folder, mod_file)

            try:
                os.remove(mod_file_path)
            except IOError as e:
                # print(f"Error deleting {mod_file}: {e}")
                logging.error(f"Error deleting {mod_file}: {e}")
                messagebox.showerror("Error", f"Error deleting {mod_file}.")

        # Delete the selected items from the mod listbox
        for index in reversed(self.mod_listbox.curselection()):
            self.mod_listbox.delete(index)

        # Show a success message
        messagebox.showinfo("Success", "Selected mods have been deleted.")

        # Refresh the installed mods list
        self.populate_mod_list_box()

    def adjust_listbox_width(self, listbox, min_width=25, max_width=50):
        max_length = min_width
        for item in listbox.get(0, tk.END):
            item_length = len(item)
            if item_length > max_length:
                max_length = item_length

        if max_length > max_width:
            max_length = max_width

        listbox.config(width=max_length)

    def center_window(self):
        self.master.update_idletasks()
        width = self.master.winfo_width()
        height = self.master.winfo_height()
        x = (self.master.winfo_screenwidth() // 2) - (width // 2)
        y = (self.master.winfo_screenheight() // 2) - (height // 2)
        self.master.geometry('{}x{}+{}+{}'.format(width, height, x, y))

    def clear_file_listbox(self):
        self.file_listbox.delete(0, tk.END)

    def load_config(self):
        """
        Load the configuration from the JSON file. If the file does not exist, return an empty dictionary.
        """
        if hasattr(sys, 'frozen'):
            local_directory = os.path.dirname(sys.executable)
        else:
            local_directory = os.path.dirname(os.path.realpath(__file__))

        config_file_path = os.path.join(local_directory, 'config.json')
        try:
            if os.path.exists(config_file_path):
                with open(config_file_path, 'r') as f:
                    return json.load(f)
            else:
                return {}
        except (IOError, json.JSONDecodeError) as e:
            # print(f"Error loading config file: {e}")
            logging.error(f"Error loading config file: {e}")
            return {}

    def save_config(self):
        """
        Save the current source and destination folders to the configuration JSON file.
        """
        if hasattr(sys, 'frozen'):
            local_directory = os.path.dirname(sys.executable)
        else:
            local_directory = os.path.dirname(os.path.realpath(__file__))

        config = {'source_folder': self.source_folder, 'destination_folder': self.destination_folder}
        config_file_path = os.path.join(local_directory, 'config.json')
        try:
            with open(config_file_path, 'w') as f:
                json.dump(config, f)
        except IOError as e:
            # print(f"Error saving config file: {e}")
            logging.error(f"Error saving config file: {e}")
            messagebox.showerror("Error", "Error saving config file.")

    def on_closing(self):
        """
        This function is called when the window is closed. It saves the current source and destination folders to the
        configuration JSON file.
        """
        self.save_config()
        self.master.destroy()

    def setup_logging(self):
        logging.basicConfig(filename='error.txt', level=logging.ERROR,
                            format='%(asctime)s - %(levelname)s - %(message)s')


if __name__ == '__main__':
    # Create the GUI window
    root = tk.Tk()
    root.geometry("700x400")  # Set the default window size here
    app = MidnightSunsMM(root)
    app.center_window()  # Center the window on the screen
    root.mainloop()