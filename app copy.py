# Program to automate WDT mailing with outlook by K. Krysa
import os
import tkinter as tk
import tkinter.messagebox as messagebox
import tkinter.ttk as ttk
import win32com.client as win32
from webbrowser import open
from shutil import move
from model import Database
from settings import *
import csv

db = Database(Settings.DATABASE)


class InvoicesMailing:
    def __init__(self):
        rows = db.select_clients()
        for row in rows:
            name = (row[1].lower()).strip()
            mail = (row[2].lower()).strip()
            if name == "poczta":
                continue
            dir_path = os.path.join(Settings.DIRECTORY, name)
            if not os.path.exists(dir_path):
                continue

            files2 = []
            for file in os.listdir(dir_path):
                file_path = os.path.join(dir_path, file)
                if os.path.isfile(file_path):
                    files = []
                    files.append(file_path)
                    self.emailer(mail, files)
                    if not os.path.exists(dir_path + "\\archiwum"):
                        os.makedirs(dir_path + "\\archiwum")
                    move(
                        file_path, dir_path + "\\archiwum\\" + Settings.date_now + file
                    )

                if os.path.isdir(
                    file_path
                ):  # search for files in subfolders except archiwum
                    for file in os.listdir(file_path):
                        file2_path = os.path.join(file_path, file)

                        if not ("archiwum" or "Archiwum") in file2_path:
                            files2.append(file2_path)
                    if files2:
                        self.emailer(mail, files2)
                        if not os.path.exists(dir_path + "\\archiwum"):
                            os.makedirs(dir_path + "\\archiwum")
                        for entry in files2:
                            file_name = os.path.basename(entry)
                            move(
                                entry,
                                dir_path
                                + "\\archiwum\\"
                                + Settings.date_now
                                + file_name,
                            )

        self.report()

    def report(self):
        myapp.count_items()
        my_report = "Raport z mailingu faktur - jeżeli jakieś pliki pozostaną niewysłane, zostaną wylistowane niżej.<br>"
        report = ""
        for k, v in myapp.items.items():
            if v != 0:
                report += f"{k} - {str(v)} <br>"
        my_report += report
        report_mail = db.select_client_by_id(1)
        self.emailer(report_mail[0][2], "", my_report)

    def file_list(self, name, dir_path):
        dict = {}
        files = []
        for path in os.listdir(dir_path):
            if ("archiwum" or "Archiwum") not in path:
                file_path = os.path.join(dir_path, path)
                files.append(file_path)
        dict[name] = files

        return dict

    def emailer(self, recipient, attachment="", body=Settings.body):
        outlook = win32.Dispatch("outlook.application")
        mail = outlook.CreateItem(0)
        mail.BCC = recipient
        mail.Subject = "HMT FAKTURA"
        mail.HtmlBody = body
        for i in range(len(attachment)):
            mail.Attachments.Add(attachment[i])
        mail.send  # Send mails
        # mail.Display(True) # Display False if you want to send email without seeing outlook window


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self["bg"] = "#f1f3f4"
        self.geometry("650x520")
        # self.geometry('650x495')
        self.title("Dystrybutor faktur")
        self.resizable(width=False, height=False)

        style = Style(self)
        self.data = {
            "name": tk.StringVar(),
            "mail": tk.StringVar(),
            "pk": 0,
        }
        self.items = {}
        self.create_dirs()

        # frames
        f_tree = ttk.Frame(self)
        f_tree.pack()

        f_control = ttk.Frame(self)
        f_control.pack(pady=20)

        self.f_footer = ttk.Frame(self)
        self.f_footer.pack(pady=5)

        # control
        l_name = ttk.Label(f_control, text="Nazwa firmy")
        l_name.grid(row=0, column=0)
        self.e_name = tk.Entry(f_control, width=35, textvariable=self.data["name"])
        self.e_name.grid(row=0, column=1)

        l_email = ttk.Label(f_control, text="Adres email")
        l_email.grid(row=1, column=0)
        self.e_email = tk.Entry(f_control, width=35, textvariable=self.data["mail"])
        self.e_email.grid(row=1, column=1)

        b_add = ttk.Button(
            f_control,
            text="Dodaj",
            width=13,
            command=lambda: self.add(
                self.data["name"].get().strip(), self.data["mail"].get().strip()
            ),
        )
        b_add.grid(row=0, column=2, padx=(40, 5))

        b_delete = ttk.Button(
            f_control,
            text="Usuń",
            width=13,
            command=lambda: self.delete((self.data["name"].get()).strip()),
        )
        b_delete.grid(row=1, column=2, padx=(40, 5))

        b_update = ttk.Button(
            f_control,
            text="Aktualizuj",
            width=13,
            command=lambda: self.update(
                self.data["name"].get().strip(),
                self.data["mail"].get().strip(),
                self.data["pk"],
            ),
        )
        b_update.grid(row=0, column=3, padx=(0, 5))

        b_open = ttk.Button(
            f_control,
            text="Otwórz folder",
            width=13,
            command=lambda: self.open_folder((self.data["name"].get()).strip()),
        )
        b_open.grid(row=1, column=3, padx=(0, 5))

        b_send = ttk.Button(
            f_control, text="\nWyślij\n", width=13, command=lambda: InvoicesMailing()
        )
        b_send.grid(row=0, column=4, rowspan=2, padx=(10, 5))
        b_test = ttk.Button(
            f_control, text="Odśwież", width=13, command=self.count_items
        )
        b_test.grid(row=2, column=4, padx=(10, 5))

        # tree
        self.tree = ttk.Treeview(
            f_tree, columns=(1, 2, 3), height=16, show="headings", selectmode="browse"
        )
        self.tree.pack(side="left")

        # tree scrollbar
        scroll = ttk.Scrollbar(f_tree, orient="vertical", command=self.tree.yview)
        scroll.pack(side="right", fill="y")

        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.heading(1, text="Kontrahent")
        self.tree.heading(2, text="Mail")
        self.tree.heading(3, text="Pliki")
        self.tree.column(1, width=300)
        self.tree.column(2, width=280)
        self.tree.column(3, width=50)
        self.tree.bind("<ButtonRelease-1>", self.display)

        # functions
        self.count_items()

    def count_items(self):
        rows = db.select_clients()
        if not rows:
            return
        for row in rows:
            count = 0
            name = row[1].lower()
            mail = row[2].lower()
            if mail == "poczta":
                continue
            path = os.path.join(Settings.DIRECTORY, name)
            if os.path.exists(path):
                for file in os.listdir(path):
                    if not ("Thumbs.db") in file and os.path.isfile(
                        os.path.join(path, file)
                    ):
                        print(os.path.join(path, file))
                        count += 1
                    if not ("archiwum" or "Archiwum") in file and os.path.isdir(
                        os.path.join(path, file)
                    ):
                        for file2 in os.listdir(os.path.join(path, file)):
                            if os.path.isfile(os.path.join(path, file, file2)):
                                count += 1
                self.items[name] = count
        for i in self.tree.get_children():
            self.tree.delete(i)
        self.view()

    def view(self):
        for i in self.tree.get_children():
            self.tree.delete(i)

        self.tree.tag_configure("orange", background="orange")
        rows = db.select_clients()
        for row in rows:
            new_row = ""
            value = 0
            for k, v in self.items.items():
                if k.lower() == row[1].lower():
                    value = v
                new_row = (row[1], row[2], value)
            if value != 0:
                self.tree.insert("", "end", values=new_row, tags=("orange",))
            else:
                self.tree.insert("", "end", values=new_row)
        self.after(600000, lambda: self.view())

    def add(self, name, email):
        if len(name) != 0 and len(email) != 0:
            rows = db.select_client_by_name(name)
            if rows:
                self.show_warning("Kontrahent " + name + " już istnieje!")
            else:
                db.insert_clients(name, email)
                self.show_info("Dodano nowy wpis do bazy danych.")

                if not os.path.exists(Settings.DIRECTORY + "\\" + name):
                    os.makedirs(Settings.DIRECTORY + "\\" + name)
            self.view()
        else:
            self.show_warning("Pola nie mogą być puste.")

    def update(self, name, email, pk):
        if len(name) != 0 and len(email) != 0:
            rows = db.select_client_by_id(pk)
            old_name = list(rows)[0][1]

            self.rename_folder(old_name, name)
            db.update_client(name, email, pk)
            self.show_info("Zmieniono dane.")
        else:
            self.show_warning("Wszystkie pola muszą być wypełnione!")
        self.view()

    def rename_folder(self, old_name, new_name):
        if not os.path.exists(Settings.DIRECTORY + "\\" + new_name):
            os.rename(
                Settings.DIRECTORY + "\\" + old_name,
                Settings.DIRECTORY + "\\" + new_name,
            )

    def open_folder(self, name):
        if len(self.e_name.get()) != 0:
            if not os.path.exists(Settings.DIRECTORY + "\\" + name):
                self.show_warning("Nie znaleziono takiego folderu!")
            else:
                open(Settings.DIRECTORY + "\\" + name)

        else:
            self.show_warning(
                "Wybierz Kontrahenta z listy i następnie spróbuj ponownie!"
            )

    def delete(self, name):
        if len(self.e_name.get()) != 0:
            confirm = messagebox.askokcancel(
                "Potwierdzenie", "Czy na pewno usunąć " + name + "?"
            )
            if confirm:
                db.delete_client(name)
                self.show_info("Kontrahent " + name + " został usunięty!")

                self.view()

    def display(self, event):
        curItem = self.tree.focus()
        id = self.tree.item(curItem)["values"]
        row = db.select_client_by_name(id[0])
        self.data["pk"] = row[0][0]

        self.e_name.delete(0, "end")
        self.e_email.delete(0, "end")
        self.e_name.insert("end", id[0])
        self.e_email.insert("end", id[1])

    def show_warning(self, msg):
        l_warning = ttk.Label(self.f_footer, text=msg, style="warning.TLabel")
        l_warning.pack(side="bottom", fill="x", pady=10)
        l_warning.after(3000, l_warning.destroy)

    def show_info(self, msg):
        l_warning = ttk.Label(self.f_footer, text=msg, style="info.TLabel")
        l_warning.pack(side="bottom", fill="x", pady=10)
        l_warning.after(3000, l_warning.destroy)

    def create_dirs(self):
        rows = db.select_clients()
        for row in rows:
            name = row[1].strip()
            if not os.path.exists(Settings.DIRECTORY + "\\" + name):
                os.makedirs(Settings.DIRECTORY + "\\" + name)


if __name__ == "__main__":
    myapp = App()
    myapp.mainloop()
