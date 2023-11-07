import os
import win32com.client as win32
from shutil import move
from model import Database
from settings import *

db = Database(Settings.DATABASE)


class Emailer:
    def __init__(self):
        self.clients_info = self.get_clients_info()
        self.send_mail(self.clients_info)

    def get_clients_info(self):
        rows = db.select_clients()
        clients_list = []
        for row in rows:
            name = (row[1].lower()).strip()
            mail = (row[2].lower()).strip()
            if name == "poczta":
                continue
            client_dir = os.path.join(Settings.DIRECTORY, name)
            if not os.path.exists(client_dir):
                continue
            files, subfiles = self.get_files_and_subfiles(client_dir)
            client_info = [name, mail, client_dir, files, subfiles]
            clients_list.append(client_info)
        return clients_list

    def get_files_and_subfiles(self, client_dir):
        files = []
        subfiles = []
        for path in os.listdir(client_dir):
            if path:
                if "." in path:
                    file_path = os.path.join(client_dir, path)
                    files.append(file_path)
                else:
                    if not path.lower() == "archiwum":
                        subfolder_path = os.path.join(client_dir, path)
                        for i in os.listdir(subfolder_path):
                            subfile_path = os.path.join(subfolder_path, i)
                            subfiles.append(subfile_path)
        return files, subfiles

    def move_to_archiwum(self, file):
        file_path = os.path.dirname(file)
        file_name = os.path.basename(file)
        new_file_name = Settings.date_now_formated + file_name
        archiwum_path = self.create_archiwum_path(file_path)
        if os.path.isfile(file):
            move(file, archiwum_path + "\\" + new_file_name)

    def create_archiwum_path(self, file_path):
        archiwum_path = os.path.join(file_path, "archiwum")
        if not os.path.exists(archiwum_path):
            os.makedirs(archiwum_path)
        return archiwum_path

    def send_mail(self, clients_info):
        for client in clients_info:
            if client[3]:
                for item in client[3]:
                    self.prepare_mail(client[1], client[0], item)
                    self.move_to_archiwum(item)

            if client[4]:
                self.prepare_mail(client[1], client[0], client[4])
                for item in client[4]:
                    pass
                    # self.move_to_archiwum(item)

    def prepare_mail(self, mail_address, client_name="", attach="", body=Settings.body):
        outlook = win32.Dispatch("outlook.application")
        mail = outlook.CreateItem(0)
        mail.BCC = mail_address
        mail.HtmlBody = body
        file_name_list = []
        if type(attach) == list:
            for i in range(len(attach)):
                if os.path.isdir(attach[i]):
                    continue
                mail.Attachments.Add(attach[i])
                file_name = self.file_name_from_path(attach[i])
                file_name_list.append(file_name)
                db.add_report(file_name, Settings.date_now, client_name)
        else:
            mail.Attachments.Add(attach)
            file_name = self.file_name_from_path(attach)
            file_name_list.append(file_name)
            db.add_report(file_name, Settings.date_now, client_name)
        file_names = ", ".join(file_name_list)

        subject = f"Faktura HMT nr: {file_names}"
        if "porsche" in mail_address.lower():
            subject = f"Faktura HMT nr: {file_names}, numer klienta: 641990"
        mail.Subject = subject
        # mail.send
        # Display False if you want to send email without seeing outlook window
        mail.Display(True)

    def file_name_from_path(self, item):
        file_name = os.path.basename(item)
        file_name = os.path.splitext(file_name)[0]
        return file_name
