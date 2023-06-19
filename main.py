from tkinter import *
import tkinter as tk
from tkinter import messagebox, filedialog
import pandas as p
import smtplib as sm
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os


class Mail:
    def __init__(self, root):
        self.root = root
        self.root.title("Mass Mail System - by Sarthak Jain")
        self.root.geometry("1000x500+0+0")

        title = Label(
            self.root,
            text="Mass Mail System",
            bd=10,
            relief=GROOVE,
            font=("times new roman", 35, "bold"),
            bg="#49c5b6",
            fg="white",
        )
        title.pack(side=TOP, fill=X)

        # ==== Variables ====

        self.sender_email = StringVar()
        self.sender_pass = StringVar()
        self.excel_file = StringVar()
        self.excel_file.set("*Upload Excel File*")
        self.excel_file_path = StringVar()
        self.attachment_file = StringVar()
        self.attachment_file.set("*Upload File*")
        self.attachment_filepath = StringVar()

        # ====Sender Frame====

        Sender_Frame = Frame(self.root, bd=4, relief=GROOVE, bg="#49c5b6")
        Sender_Frame.place(y=75, width=400, height=425)

        m_title = Label(
            Sender_Frame,
            text="Sender Info",
            bg="#49c5b6",
            fg="white",
            font=("times new roman", 30, "bold"),
        )
        m_title.grid(row=0, columnspan=2, pady=10)

        lbl_email = Label(
            Sender_Frame,
            text=" Email: ",
            bd=10,
            bg="#49c5b6",
            relief=GROOVE,
            fg="white",
            font=("times new roman", 20, "bold"),
        )
        lbl_email.grid(row=1, column=0, pady=10, padx=10, sticky="w")

        txt_email = Entry(
            Sender_Frame,
            textvariable=self.sender_email,
            font=("times new roman", 15, "bold"),
            bd=10,
            relief=GROOVE,
        )
        txt_email.grid(row=1, column=1, pady=10, padx=5, sticky="w")

        lbl_password = Label(
            Sender_Frame,
            text=" Passw: ",
            bd=10,
            bg="#49c5b6",
            relief=GROOVE,
            fg="white",
            font=("times new roman", 20, "bold"),
        )
        lbl_password.grid(row=2, column=0, pady=10, padx=10, sticky="w")

        txt_password = Entry(
            Sender_Frame,
            font=("times new roman", 15, "bold"),
            textvariable=self.sender_pass,
            bd=10,
            relief=GROOVE,
            show="*",
        )
        txt_password.grid(row=2, column=1, pady=10, padx=5, sticky="w")

        # ==== File Upload Button ====

        lbl_str = Label(Sender_Frame, textvariable=self.excel_file, fg="red")
        lbl_str.grid(row=3, columnspan=3, padx=5, pady=10)

        file_upload_btn = Button(
            Sender_Frame,
            command=self.browse_file,
            text="Upload",
            width=10,
            fg="#49c5b6",
            bd=5,
            relief=SUNKEN,
        ).grid(row=4, column=0, padx=5, pady=5, sticky="e")

        delete_upload_btn = Button(
            Sender_Frame,
            command=self.clear_file,
            text="Delete",
            width=10,
            fg="#49c5b6",
            bd=5,
            relief=SUNKEN,
        ).grid(row=4, column=1, padx=0, pady=5)

        # ==== Button Frame ====

        btn_frame = Frame(Sender_Frame, bd=4, relief=GROOVE, bg="#49c5b6")
        btn_frame.place(x=15, y=310, width=350, height=100)

        Sendbtn = Button(
            btn_frame,
            bd=10,
            relief=GROOVE,
            text="Send",
            command=self.send_email,
            font=("times new roman", 10, "bold"),
            fg="#49c5b6",
            width=12,
            height=3,
        ).grid(row=0, column=0, padx=30, pady=10)
        Clearbtn = Button(
            btn_frame,
            bd=10,
            relief=GROOVE,
            command=self.clear,
            text="Clear",
            font=("times new roman", 10, "bold"),
            fg="#49c5b6",
            width=12,
            height=3,
        ).grid(row=0, column=1, padx=10, pady=10)

        # ==== Detail Frame ====

        Detail_Frame = Frame(self.root, bd=4, relief=GROOVE, bg="#49c5b6")
        Detail_Frame.place(x=400, y=75, width=600, height=425)

        lbl_subject = Label(
            Detail_Frame,
            text=" Subject: ",
            bd=10,
            bg="#49c5b6",
            relief=GROOVE,
            fg="white",
            font=("times new roman", 20, "bold"),
        )
        lbl_subject.grid(row=0, column=0, padx=10, pady=10)

        self.txt_subject = Text(
            Detail_Frame, bd=10, relief=GROOVE, width=55, height=4, font=("", 10)
        )
        self.txt_subject.grid(row=0, column=1, pady=10, padx=20, sticky="w")

        lbl_content = Label(
            Detail_Frame,
            text=" Content: ",
            bd=10,
            bg="#49c5b6",
            relief=GROOVE,
            fg="white",
            font=("times new roman", 20, "bold"),
        )
        lbl_content.grid(row=1, column=0, padx=10, pady=10)

        self.txt_content = Text(
            Detail_Frame, bd=10, relief=GROOVE, width=55, height=15, font=("", 10)
        )
        self.txt_content.grid(row=1, rowspan=6, column=1, pady=10, padx=20, sticky="w")

        # ==== Upload File ====

        lbl_att = Label(Detail_Frame, textvariable=self.attachment_file, fg="red")
        lbl_att.grid(row=2, column=0, padx=5, pady=10)

        file_attach_btn = Button(
            Detail_Frame,
            command=self.attach_file,
            text="Upload",
            width=10,
            fg="#49c5b6",
            bd=5,
            relief=SUNKEN,
        ).grid(row=3, column=0, padx=5, pady=5)

        file_delete_btn = Button(
            Detail_Frame,
            command=self.delete_file,
            text="Delete",
            width=10,
            fg="#49c5b6",
            bd=5,
            relief=SUNKEN,
        ).grid(row=4, column=0, padx=5, pady=5)

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            initialdir="/",
            title="Select xlsx file",
            filetypes=(("Excel Files", "*.xlsx"), ("CSV Files", "*.csv")),
        )
        if file_path:
            self.excel_file_path.set(file_path)
            file_name = os.path.basename(file_path)
            self.excel_file.set(file_name)

    def attach_file(self):
        file_path = filedialog.askopenfilename(
            initialdir="/", title="Select File", filetypes=(("All files", "*.*"),)
        )
        if file_path:
            self.attachment_filepath.set(file_path)
            file_name = os.path.basename(file_path)
            self.attachment_file.set(file_name)

    def delete_file(self):
        self.attachment_file.set("*Upload File*")
        self.attachment_filepath.set("")

    def clear_file(self):
        self.excel_file.set("*Upload Excel File*")
        self.excel_file_path.set("")

    def clear(self):
        self.sender_email.set("")
        self.sender_pass.set("")
        self.excel_file_path.set("")
        self.excel_file.set("*Upload Excel File*")
        self.attachment_filepath.set("")
        self.attachment_file.set("*Upload Excel File*")
        self.txt_subject.delete("1.0", END)
        self.txt_content.delete("1.0", END)
        messagebox.showinfo("Success", "Record has been cleared")

    def send_email(self):
        if self.sender_email.get() == "" or self.sender_pass.get() == "":
            messagebox.showerror("Error", "Fill Sender Information Carefully!!!")
        elif self.txt_content.get("1.0", "end-1c").strip() == "":
            messagebox.showerror("Error", "Content can't be Empty!!!")
        else:
            file_name = self.excel_file.get()
            file_path = self.excel_file_path.get()
            if file_name != "*Upload Excel File*":
                data = p.read_excel(file_path)
                if "email" in data.columns:
                    email_col = data.get("email")
                    list_email = list(email_col)
                    try:
                        server = sm.SMTP("smtp.gmail.com", 587)
                        server.starttls()
                        server.login(self.sender_email.get(), self.sender_pass.get())
                        from_ = self.sender_email.get()
                        to_ = list_email
                        message = MIMEMultipart("alternative")
                        message["Subject"] = self.txt_subject.get("1.0", END)
                        message["From"] = self.sender_email.get()
                        body = self.txt_content.get("1.0", END)
                        message.attach(MIMEText(body, "plain"))
                        if self.attachment_filepath.get() != "":
                            filename = self.attachment_filepath.get()
                            basename = os.path.basename(filename)
                            with open(filename, "rb") as attachment:
                                part = MIMEBase("application", "octet-stream")
                                part.set_payload(attachment.read())
                            encoders.encode_base64(part)
                            part.add_header(
                                "Content-Disposition",
                                f"attachment; filename= {basename}",
                            )
                            message.attach(part)
                        server.sendmail(from_, to_, msg=message.as_string())
                        messagebox.showinfo(
                            "Success", "Your Email has been Send Successfully"
                        )

                    except Exception as e:
                        messagebox.showerror("Error", e)

                else:
                    messagebox.showerror("Error", "email Column not found!!!")
            else:
                messagebox.showerror("Error", "File Path is Empty!!!")


root = Tk()
root.resizable(False, False)
ob = Mail(root)
root.mainloop()
