from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from pymongo import MongoClient
from bson import ObjectId
import csv
from tkinter import filedialog
from bson.errors import InvalidId



# MongoDB connection
client_value = "mongodb+srv://pushpcmpd:pushpcmpd@pushcmpd.thcbgri.mongodb.net/"
client = MongoClient(client_value)
db = client["drilling_data"]
collection = db["data"]

#======================================CMPDI VARIABLE DEFINATIONS=========================================================================================

regional_institutes = ["RI-I-Asansol","RI-II-Dhanbad","RI-III-Ranchi","RI-IV-Nagpur", "RI-V-Bilaspur", "RI-VI-Singrauli", "RI-VII-Bhubneshwar"]

regional_institute_I = ["Dharma", "Gourangdih", "Mallarpur"]

regional_institute_III = ["Barkakana","Hazribagh","Orla"]

regional_institute_IV = ["Anandwan", "Durgapur", "Murpar"]

regional_institute_V = ["Kudumkela", "Kusmuda", "Korba", "Rajnagar","Singhpur"]

regional_institute_II = []

regional_institute_VI = ["Singrauli"]

regional_institute_VII = ["Gopalpur","Kosala","Talcher"]

camp_kudumkela = ["WAIIIC-1","WAIIIC-16","WAIIIC-19"]

camp_kusmunda = ["WAIIIC-6","CT-1","CT-2","CT-3","CT-4"]

camp_korba = ["CMME-01","CMDM-14","CMKME-05","WAIIIC-10"]

camp_rajnagar= ["CMDM-04","CMDM-22","WAIIIC-11","WAIIIC-12"]

camp_singhpur = ["CMDM-13","CMDM-15","CMKME-04","WAIIIC-07","WAIIIC-08"]


# OTHER VARIABLES
my_display_table = NONE


#==================================================CMPDI VARAIABLE DEFINATIONS ENDING=================================================================
#
#
#========================================================FUNCTION DEFINATIONS===============================================================

def view_last_entry():
    data = collection.find().sort("_id", -1).limit(1)
    for entry in data:
        result_text.delete(1.0, END)
        result_text.insert(END, f"id    : {entry['_id']}\n")
        result_text.insert(END, f"Date              : {entry['Date']}\n")
        result_text.insert(END, f"Regional Institute: {entry['Regional Institute']}\n")
        result_text.insert(END, f"Drilling Camp     : {entry['Drilling Camp']}\n")
        result_text.insert(END, f"Rig Name          : {entry['Rig Name']}\n")
        result_text.insert(END, f"First Shift       : {entry['First Shift']}\n")
        result_text.insert(END, f"Second Shift      : {entry['Second Shift']}\n")
        result_text.insert(END, f"BH Name           : {entry['BH Name']}\n")
        result_text.insert(END, f"Drilling Type     : {entry['Drilling Type']}\n")
        result_text.insert(END, f"Bit No            : {entry['Bit No']}\n")
        result_text.insert(END, f"Remarks           : {entry['Remarks']}\n")
        
# Function to insert data into MongoDB
def add_entry():
        data = {
            "Date": date_entry.get(),
            "Regional Institute": ri_combo.get(),
            "Drilling Camp": camp_combo.get(),
            "Rig Name": rig_combo.get(),
            "First Shift": first_shift_entry.get(),
            "Second Shift": second_shift_entry.get(),
            "BH Name": bh_name_entry.get(),
            "Drilling Type": bh_drilling_type.get(),
            "Bit No": bit_no_entry.get(),
            "Remarks": remarks_entry.get()
        }
        
        if "" in [data["Date"],data["Regional Institute"],data["Drilling Camp"],data["Rig Name"],data["Drilling Type"],data["Bit No"]]:
            messagebox.showinfo("Failure", "Please enter valid values in the form")
        else:
            collection.insert_one(data)
            messagebox.showinfo("Success", "Entry Added successfully")
            
        view_last_entry()
        view_database()
    
def clear():
    date_entry.delete(0,END)
    rig_combo.delete(0,END)
    ri_combo.delete(0,END)
    camp_combo.delete(0,END)
    first_shift_entry.delete(0,END)
    second_shift_entry.delete(0,END)
    bh_name_entry.delete(0,END)
    bh_drilling_type.delete(0,END)
    bit_no_entry.delete(0,END)
    remarks_entry.delete(0,END)
    result_text.delete(1.0, END)
    
    
def select_camp(event):
    if ri_combo.get()==regional_institutes[0]:
        camp_combo.config(values=regional_institute_I)
    elif ri_combo.get()==regional_institutes[1]:
        camp_combo.config(values=regional_institute_II)
    elif ri_combo.get()==regional_institutes[2]:
        camp_combo.config(values=regional_institute_III)
    elif ri_combo.get()==regional_institutes[3]:
        camp_combo.config(values=regional_institute_IV)
    elif ri_combo.get()==regional_institutes[4]:
        camp_combo.config(values=regional_institute_V)
    elif ri_combo.get()==regional_institutes[5]:
        camp_combo.config(values=regional_institute_VI)
    else:
        camp_combo.config(values=regional_institute_VII)
        
def select_rig(event):
    if camp_combo.get()==regional_institute_V[0]:
        rig_combo.config(values=camp_kudumkela)
    elif camp_combo.get()==regional_institute_V[1]:
        rig_combo.config(values=camp_kusmunda)
    elif camp_combo.get()==regional_institute_V[2]:
        rig_combo.config(values=camp_korba)
    elif camp_combo.get()==regional_institute_V[3]:
        rig_combo.config(values=camp_rajnagar)
    else:
        ri_combo.config(values=camp_singhpur)

def view_database():
    global drilling_treeview
    global my_display_table
    #scollable treeview for displaying database values inside editble window
    my_display_table = ttk.Treeview(drilling_treeview)

    my_display_table["columns"] = ("Date","Regional Institute", "Drilling Camp","Rig Name", "1st Shift Progress", 
                                "2nd Shift Progress","BH Name", "Drilling Type","Bit No","Remarks")
    #formatting the columns in treeview
    my_display_table.column("#0",width = 80,minwidth=30)
    my_display_table.column("Date",anchor=CENTER,width = 50,minwidth=30)
    my_display_table.column("Regional Institute",anchor=CENTER,width = 80,minwidth=30)
    my_display_table.column("Drilling Camp",anchor=CENTER,width = 50,minwidth=30)
    my_display_table.column("Rig Name",anchor=CENTER,width = 50,minwidth=30)
    my_display_table.column("1st Shift Progress",anchor=CENTER,width = 40,minwidth=10)
    my_display_table.column("2nd Shift Progress",anchor=CENTER,width = 40,minwidth=10)
    my_display_table.column("BH Name",anchor=CENTER,width = 80,minwidth=10)
    my_display_table.column("Drilling Type",anchor=CENTER,width = 40,minwidth=10)
    my_display_table.column("Bit No",anchor=CENTER,width = 50,minwidth=30)
    my_display_table.column("Remarks",anchor=CENTER,width = 150,minwidth=30)

    #Headings of columns in treeview
    my_display_table.heading("#0",text = "Object_ID", anchor=CENTER)
    my_display_table.heading("Date",text = "Date", anchor=CENTER)
    my_display_table.heading("Regional Institute",text = "RI", anchor=CENTER)
    my_display_table.heading("Drilling Camp",text = "Camp", anchor=CENTER)
    my_display_table.heading("Rig Name",text = "Rig", anchor=CENTER,)
    my_display_table.heading("1st Shift Progress",text = "1st Shift", anchor=CENTER,)
    my_display_table.heading("2nd Shift Progress",text = "2nd Shift", anchor=CENTER,)
    my_display_table.heading("BH Name",text = "BH Name", anchor=CENTER)
    my_display_table.heading("Drilling Type",text = "Type", anchor=CENTER)
    my_display_table.heading("Bit No",text = "Bit", anchor=CENTER)
    my_display_table.heading("Remarks",text = "Remarks", anchor=CENTER)
    my_display_table.place(x=5,y=5,width= 1170,height=400)

    scrollbar = ttk.Scrollbar(my_display_table, orient="vertical", command=my_display_table.yview)
    scrollbar.pack(side=RIGHT, fill=Y)
    my_display_table.configure(yscrollcommand=scrollbar.set)

    # Fetch all entries from MongoDB
    my_table = collection.find()

    for entry in my_table:
        entry_id = entry["_id"]
        date = entry["Date"]
        ri_val = entry["Regional Institute"]
        camp_val = entry["Drilling Camp"]
        rig_name = entry["Rig Name"]
        first_shift = entry["First Shift"]
        second_shift = entry["Second Shift"]
        bh_name = entry["BH Name"]
        bh_drilling_type =entry["Drilling Type"]
        bit_no = entry["Bit No"]
        remarks = entry["Remarks"]

        my_display_table.insert("", "end", text=entry_id, values=(date, ri_val,camp_val, rig_name, first_shift, second_shift, bh_name,bh_drilling_type, bit_no, remarks))

#Edit the record by loading values in the form
def edit_record():
    global my_display_table
    # global date_entry, ri_combo, camp_combo, rig_combo
    # global first_shift_entry, second_shift_entry, bh_drilling_type, bh_name_label, bit_no_entry, remarks_entry
    
    # Retrieve the selected entry's ID from the Treeview widget
    selected_item = my_display_table.focus()
    selected_entry_value = my_display_table.item(selected_item)['values']  
    selected_entry_id = my_display_table.item(selected_item)['text']
    
    #inserting to form selected values 
    clear()
    if selected_item:
        date_entry.insert(0,selected_entry_value[0])
        ri_combo.insert(0,selected_entry_value[1])
        camp_combo.insert(0,selected_entry_value[2])
        rig_combo.insert(0,selected_entry_value[3])
        first_shift_entry.insert(0,selected_entry_value[4])
        second_shift_entry.insert(0,selected_entry_value[5])
        bh_name_entry.insert(0,selected_entry_value[6])
        bh_drilling_type.insert(0,selected_entry_value[7])
        bit_no_entry.insert(0,selected_entry_value[8])
        remarks_entry.insert(0,selected_entry_value[9])
    else:
        messagebox.showinfo("Failure", "Select a record to edit")
    
def update_record():
    global my_display_table   
    selected_item = my_display_table.focus()
    
    if selected_item:
        selected_entry_id = my_display_table.item(selected_item)['text']    
        
        data = {
                "Date": date_entry.get(),
                "Regional Institute": ri_combo.get(),
                "Drilling Camp": camp_combo.get(),
                "Rig Name": rig_combo.get(),
                "First Shift": first_shift_entry.get(),
                "Second Shift": second_shift_entry.get(),
                "BH Name": bh_name_entry.get(),
                "Drilling Type": bh_drilling_type.get(),
                "Bit No": bit_no_entry.get(),
                "Remarks": remarks_entry.get()
        }
        if "" in [data["Date"],data["Regional Institute"],data["Drilling Camp"],data["Rig Name"],data["Drilling Type"],data["Bit No"]]:
            messagebox.showinfo("Failure", "Please update with legitimate values")
        else:
            collection.update_one({"_id": ObjectId(selected_entry_id)}, {"$set": data})
            messagebox.showinfo("Success", "Entry updated successfully")
        
    else:
        messagebox.showinfo("Failure","Please select a record for editing first")
    
    view_database()

      
    
#Delete the record from database by selecting from the treeview

def delete_record():
    global my_display_table
    # Retrieve the selected entry's ID from the Treeview widget
    selected_item = my_display_table.focus()
    if selected_item:
        selected_entry_id = my_display_table.item(selected_item)["text"]

        # Confirm deletion with the user
        confirm = messagebox.askyesno("Confirm Deletion", "Are you sure you want to delete this entry?")

        if confirm:
            my_display_table.destroy()         
            try:
                # Delete the entry from the MongoDB collection
                collection.delete_one({"_id": ObjectId(selected_entry_id)})
                messagebox.showinfo("Success", "Entry deleted successfully")

            except InvalidId:
                messagebox.showerror("Error", "Invalid ID selected")
                
    view_database()
            
            
#=========================================================GUI CREATIONS======================================================================
# Create the GUI
drilling_gui = Tk()
drilling_gui.title("Drilling Data GUI")
drilling_gui.geometry("1200x800")
style = ttk.Style()
style.theme_use('clam')
#========================================================PROGRESS INFORMATION================================================================
#Create Frame for form entry of drilling and view of last entry with buttons Add data and clearing of fields 
drilling_entry = LabelFrame(drilling_gui, text = "Progress information")
drilling_entry.place(x=5,y=5,width = 1150,height=280)

#-------------------------------------------Drilling progress entry----------------------------------------------------

#++++++++++++++++++++++++++CMPDI STRUCTURE Selection Frame for entry of Data+++++++++++++++++++++++++++++++++++++++++++++
rig_address = LabelFrame(drilling_entry, text="Basic Information")
rig_address.place(x=5, y=5, width = 400,height=280)

# Date
date_label = Label(rig_address, text="Date(dd-mm-yyyy)")
date_label.grid(row=0, column=0,padx=2,pady=10)
date_entry = Entry(rig_address)
date_entry.grid(row=0, column=1, padx=2,pady=2)

ri_label = Label(rig_address, text="Regional Institure")
ri_label.grid(row=1, column=0,padx=2,pady=10)

ri_combo = ttk.Combobox(rig_address,value = regional_institutes)
ri_combo.grid(row=1,column=1, padx=2,pady=10)
ri_combo.bind("<<ComboboxSelected>>",select_camp)


camp_label = Label(rig_address, text="Drilling Camp")
camp_label.grid(row=2, column=0,padx=2,pady=10)

camp_combo = ttk.Combobox(rig_address,value = [""])
camp_combo.grid(row=2,column=1, padx=2,pady=10)
camp_combo.bind("<<ComboboxSelected>>",select_rig)

#+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

# Drilling Entry Form Frame
drilling_form = LabelFrame(drilling_entry, text="Drilling Progress Form")
drilling_form.place(x=405, y=5, width = 385,height=300)

#Creating various form entries inside drilling entry form frame

# Rig Name
rig_name_label = Label(drilling_form, text="Rig Name")
rig_name_label.grid(row=1, column=0,padx=2,pady=2)
rig_combo = ttk.Combobox(drilling_form,value = [""])
rig_combo.grid(row=1,column=1, padx=2,pady=2)

# First Shift
first_shift_label = Label(drilling_form, text="First Shift")
first_shift_label.grid(row=2, column=0,padx=2,pady=2)
first_shift_entry = Entry(drilling_form)
first_shift_entry.grid(row=2, column=1,columnspan=5,padx=2,pady=2)

# Second Shift
second_shift_label = Label(drilling_form, text="Second Shift")
second_shift_label.grid(row=3, column=0,padx=2,pady=2)
second_shift_entry = Entry(drilling_form)
second_shift_entry.grid(row=3, column=1,columnspan=5,padx=2,pady=2)

# BH Name
bh_name_label = Label(drilling_form, text="BH Name")
bh_name_label.grid(row=4, column=0,padx=2,pady=2)
bh_name_entry = Entry(drilling_form)
bh_name_entry.grid(row=4, column=1,columnspan=5,padx=2,pady=2)

# Drilling Type
bh_drilling_type = Label(drilling_form, text="Drilling Type")
bh_drilling_type.grid(row=5, column=0,padx=2,pady=2)
bh_drilling_type = Entry(drilling_form)
bh_drilling_type.grid(row=5, column=1,columnspan=5,padx=2,pady=2)

# Bit No
bit_no_label = Label(drilling_form, text="Bit No")
bit_no_label.grid(row=6, column=0,padx=2,pady=2)
bit_no_entry = Entry(drilling_form)
bit_no_entry.grid(row=6, column=1,columnspan=5,padx=2,pady=2)

# Remarks
remarks_label = Label(drilling_form, text="Remarks")
remarks_label.grid(row=7, column=0,padx=2,pady=2)
remarks_entry = Entry(drilling_form)
remarks_entry.grid(row=7, column=1,columnspan=5,padx=2,pady=2)

# Add Button
add_button = Button(drilling_form, text="Add Entry", command=add_entry)
add_button.grid(row=8, column=0,padx=5,pady=5)

# View Last Entry Button
update_button = Button(drilling_form, text="Update", command=update_record)
update_button.grid(row=8, column=1,padx=5,pady=5)

# Clear field  Button
clear_button = Button(drilling_form, text="Clear", command=clear)
clear_button.grid(row=8, column=2,padx=5,pady=5)

#====================VIEW LAST ADDED INFORMATION======================================================
# Last Drillng Data entered viewing
drilling_view = LabelFrame(drilling_entry, text="Last Entered Data")
drilling_view.place(x=800, y=5, width = 350,height=280)

# Result Text
result_text = Text(drilling_view, height=280, width=280)
result_text.grid(row=0, column=0)

#========================================EDITING DATABASE/VIEWING DATABASE IN TREE VIEW=========================================
#Create Frame for viewing/editing/deleting database entry of drilling data
drilling_editing= LabelFrame(drilling_gui, text = "Viewing and editing of database")
drilling_editing.place(x=5,y=300,width = 1190,height =500)


#Treeview Frame
drilling_treeview= LabelFrame(drilling_editing, text = "Check all database records")
drilling_treeview.place(x=5, y= 5, width = 1180, height = 350)
view_database()

#Buttons Frame for editing deleting and searching and any further function
edit_button_frame = LabelFrame(drilling_editing, text = "Carry editing from here")
edit_button_frame.place(x=5, y= 365, width = 1190, height = 100)

#Creating Edit button for database
edit_button = Button(edit_button_frame, text="Edit", command=edit_record)
edit_button.grid(row = 0, column=0, padx=50, pady=25)

#Creating Delete button fr database
delete_button = Button(edit_button_frame, text="Delete", command=delete_record)
delete_button.grid(row = 0, column=2, padx=50, pady=25)


#========================================Mainloop close=======================================
drilling_gui.mainloop()
