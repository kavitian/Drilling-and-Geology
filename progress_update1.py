from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from pymongo import MongoClient
from bson import ObjectId
from bson.errors import InvalidId
import csv
from tkinter import filedialog



tree = None

# MongoDB connection
client_value = "mongodb+srv://pushpcmpd:pushpcmpd@pushcmpd.thcbgri.mongodb.net/test"
client = MongoClient(client_value)
db = client["drilling_data"]
collection = db["data"]

# Tkinter GUI
root = Tk()
root.title("Drilling Data GUI")

# Global variable to store selected entry ID
selected_entry_id = None


def import_csv_data():
    # Open file selection prompt
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    
    if file_path:
        # Open the selected CSV file for reading
        with open(file_path, 'r') as file:
            reader = csv.DictReader(file)

            # Iterate over each row in the CSV file
            for row in reader:
                # Insert the row data into the MongoDB collection
                collection.insert_one(row)

        status_label.config(text="CSV data imported successfully")

# Function to insert data into MongoDB
def add_entry():
    data = {
        "Date": date_entry.get(),
        "Rig Name": rig_name_entry.get(),
        "First Shift": first_shift_entry.get(),
        "Second Shift": second_shift_entry.get(),
        "BH Name": bh_name_entry.get(),
        "Bit No": bit_no_entry.get(),
        "Remarks": remarks_entry.get()
    }
    collection.insert_one(data)
    status_label.config(text="Entry added successfully")

def view_last_entry():
    data = collection.find().sort("_id", -1).limit(1)
    for entry in data:
        result_text.delete(1.0, END)
        result_text.insert(END, f"Date: {entry['Date']}\n")
        result_text.insert(END, f"Rig Name: {entry['Rig Name']}\n")
        result_text.insert(END, f"First Shift: {entry['First Shift']}\n")
        result_text.insert(END, f"Second Shift: {entry['Second Shift']}\n")
        result_text.insert(END, f"BH Name: {entry['BH Name']}\n")
        result_text.insert(END, f"Bit No: {entry['Bit No']}\n")
        result_text.insert(END, f"Remarks: {entry['Remarks']}\n")

def edit_entry():
    global tree
    # Retrieve the selected entry's ID from the Treeview widget
    selected_item = tree.focus()
    if selected_item:
        selected_entry_id = tree.item(selected_item)["text"]
        # Close the previously opened 
        tree.destroy()
        
        try:            
            # Query the MongoDB collection to fetch the entry data
            entry = collection.find_one({"_id": ObjectId(selected_entry_id)})
            if entry:
                # Create a new window for editing the entry
                edit_window = Toplevel(root)
                edit_window.title("Edit Entry")

                # Date
                date_label = Label(edit_window, text="Date")
                date_label.grid(row=0, column=0)
                date_entry = Entry(edit_window)
                date_entry.grid(row=0, column=1, columnspan=3)
                date_entry.insert(END, entry["Date"])

                # Rig Name
                rig_name_label = Label(edit_window, text="Rig Name")
                rig_name_label.grid(row=1, column=0)
                rig_name_entry = Entry(edit_window)
                rig_name_entry.grid(row=1, column=1, columnspan=3)
                rig_name_entry.insert(END, entry["Rig Name"])

                # First Shift
                first_shift_label = Label(edit_window, text="First Shift")
                first_shift_label.grid(row=2, column=0)
                first_shift_entry = Entry(edit_window)
                first_shift_entry.grid(row=2, column=1, columnspan=3)
                first_shift_entry.insert(END, entry["First Shift"])

                # Second Shift
                second_shift_label = Label(edit_window, text="Second Shift")
                second_shift_label.grid(row=3, column=0)
                second_shift_entry = Entry(edit_window)
                second_shift_entry.grid(row=3, column=1, columnspan=3)
                second_shift_entry.insert(END, entry["Second Shift"])

                # BH Name
                bh_name_label = Label(edit_window, text="BH Name")
                bh_name_label.grid(row=4, column=0)
                bh_name_entry = Entry(edit_window)
                bh_name_entry.grid(row=4, column=1, columnspan=3)
                bh_name_entry.insert(END, entry["BH Name"])

                # Bit No
                bit_no_label = Label(edit_window, text="Bit No")
                bit_no_label.grid(row=5, column=0)
                bit_no_entry = Entry(edit_window)
                bit_no_entry.grid(row=5, column=1, columnspan=3)
                bit_no_entry.insert(END, entry["Bit No"])

                # Remarks
                remarks_label = Label(edit_window, text="Remarks")
                remarks_label.grid(row=6, column=0)
                remarks_entry = Entry(edit_window)
                remarks_entry.grid(row=6, column=1, columnspan=3)
                remarks_entry.insert(END, entry["Remarks"])

                def save_changes():
                # Retrieve the edited data from the entry form fields
                    edited_data = {
                                    "Date": date_entry.get(),
                                    "Rig Name": rig_name_entry.get(),
                                    "First Shift": first_shift_entry.get(),
                                    "Second Shift": second_shift_entry.get(),
                                    "BH Name": bh_name_entry.get(),
                                    "Bit No": bit_no_entry.get(),
                                    "Remarks": remarks_entry.get()
                                    }

                    # Update the entry in the MongoDB collection
                    collection.update_one({"_id": ObjectId(selected_entry_id)}, {"$set": edited_data})
                    messagebox.showinfo("Success", "Entry updated successfully")

                    # Close the edit window
                    edit_window.destroy()

                    # Refresh the Treeview widget to reflect the updated data
                    view_all_entries()
                    
                # Save button
                commit_button = ttk.Button(edit_window, text="Commit Changes", command=save_changes)
                commit_button.grid(row = 9,column = 2, pady=10)
        except InvalidId:
            messagebox.showerror("Error", "Please delete the Entry First and re-enter the data")


def delete_selected_entry():
    global tree
    # Retrieve the selected entry's ID from the Treeview widget
    selected_item = tree.focus()
    if selected_item:
        selected_entry_id = tree.item(selected_item)["text"]

        # Confirm deletion with the user
        confirm = messagebox.askyesno("Confirm Deletion", "Are you sure you want to delete this entry?")

        if confirm:
            tree.destroy()         
            try:
                # Delete the entry from the MongoDB collection
                collection.delete_one({"_id": ObjectId(selected_entry_id)})
                messagebox.showinfo("Success", "Entry deleted successfully")

                # Refresh the Treeview widget to reflect the updated data
                view_all_entries()

            except InvalidId:
                messagebox.showerror("Error", "Invalid ID selected")
                
            view_all_entries()


def view_all_entries():
    global tree
    all_entries_window = Toplevel(root)
    all_entries_window.title("All Entries")

    # Create Treeview widget
    tree = ttk.Treeview(all_entries_window)
    tree["columns"] = ("date", "rig_name", "first_shift", "second_shift", "bh_name", "bit_no", "remarks")
    tree.heading("#0", text="Entry ID")
    tree.heading("date", text="Date")
    tree.heading("rig_name", text="Rig Name")
    tree.heading("first_shift", text="First Shift")
    tree.heading("second_shift", text="Second Shift")
    tree.heading("bh_name", text="BH Name")
    tree.heading("bit_no", text="Bit No")
    tree.heading("remarks", text="Remarks")
    tree.pack(fill=BOTH, expand=True)

    scrollbar = ttk.Scrollbar(all_entries_window, orient="vertical", command=tree.yview)
    scrollbar.pack(side=RIGHT, fill=Y)
    tree.configure(yscrollcommand=scrollbar.set)

    # Fetch all entries from MongoDB
    data = collection.find()

    for entry in data:
        entry_id = entry["_id"]
        date = entry["Date"]
        rig_name = entry["Rig Name"]
        first_shift = entry["First Shift"]
        second_shift = entry["Second Shift"]
        bh_name = entry["BH Name"]
        bit_no = entry["Bit No"]
        remarks = entry["Remarks"]

        tree.insert("", "end", text=entry_id, values=(date, rig_name, first_shift, second_shift, bh_name, bit_no, remarks))
                    

    edit_button = ttk.Button(all_entries_window, text="Edit", command=edit_entry)
    edit_button.pack(side=LEFT, padx=10, pady=10)

    delete_button = ttk.Button(all_entries_window, text="Delete", command=delete_selected_entry)
    delete_button.pack(side=LEFT, padx=10, pady=10)

# Date
date_label = Label(root, text="Date")
date_label.grid(row=0, column=0)
date_entry = Entry(root)
date_entry.grid(row=0, column=1, columnspan=3)

# Rig Name
rig_name_label = Label(root, text="Rig Name")
rig_name_label.grid(row=1, column=0)
rig_name_entry = Entry(root)
rig_name_entry.grid(row=1, column=1, columnspan=3)

# First Shift
first_shift_label = Label(root, text="First Shift")
first_shift_label.grid(row=2, column=0)
first_shift_entry = Entry(root)
first_shift_entry.grid(row=2, column=1, columnspan=3)

# Second Shift
second_shift_label = Label(root, text="Second Shift")
second_shift_label.grid(row=3, column=0)
second_shift_entry = Entry(root)
second_shift_entry.grid(row=3, column=1, columnspan=3)

# BH Name
bh_name_label = Label(root, text="BH Name")
bh_name_label.grid(row=4, column=0)
bh_name_entry = Entry(root)
bh_name_entry.grid(row=4, column=1, columnspan=3)

# Bit No
bit_no_label = Label(root, text="Bit No")
bit_no_label.grid(row=5, column=0)
bit_no_entry = Entry(root)
bit_no_entry.grid(row=5, column=1, columnspan=3)

# Remarks
remarks_label = Label(root, text="Remarks")
remarks_label.grid(row=6, column=0)
remarks_entry = Entry(root)
remarks_entry.grid(row=6, column=1, columnspan=3)

# Add Button
add_button = Button(root, text="Add Entry", command=add_entry)
add_button.grid(row=7, column=0)

# View Last Entry Button
view_button = Button(root, text="View Last Entry", command=view_last_entry)
view_button.grid(row=7, column=1)

# View All Entries Button
view_all_button = Button(root, text="View All Entries", command=view_all_entries)
view_all_button.grid(row=7, column=3)

import_button = Button(root, text="Import CSV", command=import_csv_data)
import_button.grid(row=7, column=4)


# Result Text
result_text = Text(root, height=10, width=20)
result_text.grid(row=8, columnspan=8)

# Status Label
status_label = Label(root, text="")
status_label.grid(row=12, columnspan=4)


root.mainloop()

