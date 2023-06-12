import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkcalendar import Calendar
import pandas as pd

class ExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel App")
        self.current_place = None
        self.current_index = 0

        # Load Excel data
        self.data = pd.read_excel("phonebook.xlsx")

        # Add 'Count' column if not present
        if 'Count' not in self.data.columns:
            self.data['Count'] = 0

        # GUI elements
        self.place_dropdown = ttk.Combobox(root, state="readonly")
        self.place_dropdown.grid(row=0, column=0, padx=10, pady=10)
        self.place_dropdown.bind("<<ComboboxSelected>>", self.select_place)

        self.phone_label = tk.Label(root, text="Phone Number:")
        self.phone_label.grid(row=1, column=0, padx=10, pady=10)

        self.name_label = tk.Label(root, text="Name:")
        self.name_label.grid(row=1, column=1, padx=10, pady=10)

        self.date_label = tk.Label(root, text="Selected Date:")
        self.date_label.grid(row=2, column=0, padx=10, pady=10)

        self.date_button = tk.Button(root, text="Select Date", command=self.select_date)
        self.date_button.grid(row=0, column=1, padx=10, pady=10)

        self.message_button = tk.Button(root, text="Message", command=self.send_message)
        self.message_button.grid(row=4, column=1, padx=10, pady=10)

        self.ignore_button = tk.Button(root, text="Next", command=self.ignore_number)
        self.ignore_button.grid(row=4, column=0, padx=10, pady=10)

    def select_place(self, event):
        self.current_place = self.place_dropdown.get()
        self.current_index = 0
        self.update_info()

    def update_info(self):
        filtered_data = self.data[self.data['Place'] == self.current_place]
        if not filtered_data.empty:
            current_row = filtered_data.iloc[self.current_index]
            phone_number = current_row['Phone']
            name = current_row['Name']
            self.phone_label.config(text="Phone Number: " + str(phone_number))
            self.name_label.config(text="Name: " + str(name))
        else:
            self.phone_label.config(text="Phone Number:")
            self.name_label.config(text="Name:")

    def select_date(self):
        top = tk.Toplevel(self.root)

        def on_date_selected():
            selected_date = cal.selection_get()
            self.date_label.config(text="Selected Date: " + str(selected_date))
            top.destroy()

        cal = Calendar(top, selectmode="day")
        cal.pack()
        confirm_button = tk.Button(top, text="Confirm", command=on_date_selected)
        confirm_button.pack()

    # ...

    def send_message(self):
        if self.current_place is None:
            messagebox.showinfo("Error", "Please select a place")
        elif self.date_label.cget("text") == "Selected Date:":
            messagebox.showinfo("Error", "Please select a date")
        else:
            selected_date = self.date_label.cget("text")[15:]
            phone_number_text = self.phone_label.cget("text")
            name_text = self.name_label.cget("text")

            phone_number = ""
            if phone_number_text.startswith("Phone Number:"):
                phone_number = phone_number_text.split(":")[1].strip()

            name = ""
            if name_text.startswith("Name:"):
                name = name_text.split(":")[1].strip()

            if phone_number == "":
                messagebox.showinfo("Error", "Phone number not found")
            elif not phone_number.isdigit():
                messagebox.showinfo("Error", "Invalid phone number")
            else:
                # Customized message
                message = f" Sairam {name} ({phone_number}),You have been placed as an incharge for Shirdi Sai Program on {selected_date}."
                self.display_message(message)

                # Update count in the 4th column
                filtered_indices = self.data[
                    (self.data['Place'] == self.current_place) & (self.data['Phone'] == int(phone_number))].index
                if not filtered_indices.empty:
                    index = filtered_indices[0]
                    count_value = self.data.at[index, 'Count']
                    if pd.isnull(count_value):
                        count_value = 0
                    self.data.at[index, 'Count'] = count_value + 1

                    # Save the updated DataFrame to the Excel file
                    self.data.to_excel("phonebook.xlsx", index=False)

                self.current_index = (self.current_index + 1) % len(self.data[self.data['Place'] == self.current_place])
                self.update_info()

    # ...

    def display_message(self, message):
        top = tk.Toplevel(self.root)
        top.title("Message")

        message_label = tk.Label(top, text="Message:")
        message_label.pack(padx=10, pady=10)

        message_text = tk.Text(top, height=5, width=40)
        message_text.insert(tk.END, message)
        message_text.pack(padx=10, pady=10)

        copy_button = tk.Button(top, text="Copy", command=lambda: self.copy_message(message))
        copy_button.pack(padx=10, pady=10)

    def copy_message(self, message):
        self.root.clipboard_clear()
        self.root.clipboard_append(message)


    def ignore_number(self):
        if self.current_place is None:
            messagebox.showinfo("Error", "Please select a place")
        else:
            self.current_index = (self.current_index + 1) % len(self.data[self.data['Place'] == self.current_place])
            self.update_info()


if __name__ == '__main__':
    root = tk.Tk()
    app = ExcelApp(root)

    # Populate the place dropdown with unique places from the Excel sheet
    places = app.data['Place'].unique()
    app.place_dropdown['values'] = sorted(list(set(app.data['Place'].dropna())))

    root.mainloop()
