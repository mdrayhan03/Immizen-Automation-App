from customtkinter import *
from tkinter import messagebox
from numpy import random
from Helper.Supabase import Excel_Book

class Second:
    def __init__(self, window, current_path):
        self.window = window
        self.bgcolor = "#FCFBFA"
        self.headlinecolor = "#B08A6C"
        self.textcolor = "#151515"
        self.entrycolor = "#A79E87"
        self.entrytextcolor = "#151515"
        self.current_path = current_path
        self.vcmd = (self.window.register(self.validate_input), '%P')
        self.mainFrame = CTkFrame(self.window, fg_color=self.bgcolor)
        self.client_part()
        self.first_part()
        self.second_part()
        self.mainFrame.pack()

    def font(self, n):
        return ("Arial", n)
    
    def validate_input(self, value_if_allowed):
        if value_if_allowed == "" or value_if_allowed.isdigit():
            return True
        else:
            return False
    
    def fetch_client(self):
        self.excel = Excel_Book()        
        return self.excel.select_info()

    def client_part(self) :
        self.client_body = CTkFrame(self.mainFrame, fg_color=self.bgcolor)

        CTkLabel(self.client_body, text="Select Client", text_color=self.textcolor).grid(column=0,row=0)

        self.client = CTkComboBox(self.client_body, width=300, values=self.fetch_client(), state="readonly", fg_color=self.entrycolor, text_color=self.entrytextcolor)
        self.client.grid(column=1,row=0)

        for child in self.client_body.winfo_children():
            child.configure(font=self.font(18))
            child.grid_configure(padx=5, pady=5, sticky="w")

        self.client_body.pack()
    
    def first_part(self):
        CTkLabel(self.mainFrame, text="Invoice", font=self.font(30), text_color=self.headlinecolor).pack(pady=(40,20))
        self.firstBody = CTkFrame(self.mainFrame, fg_color=self.bgcolor)

        # first part label
        
        CTkLabel(self.firstBody, text="Professional Fee's", text_color="#151515").grid(column=0,row=1)
        CTkLabel(self.firstBody, text="Govt.Application Fee's", text_color="#151515").grid(column=0,row=2)
        

        # first part combobox        
        self.professional = CTkEntry(self.firstBody, width=300, validate="key", validatecommand=self.vcmd, fg_color=self.entrycolor, text_color=self.entrytextcolor)
        self.professional.grid(column=1,row=1)
        self.govt = CTkEntry(self.firstBody, width=300, validate="key", validatecommand=self.vcmd, fg_color=self.entrycolor, text_color=self.entrytextcolor)
        self.govt.grid(column=1,row=2)        

        for child in self.firstBody.winfo_children():
            child.configure(font=self.font(18))
            child.grid_configure(padx=5, pady=5, sticky="w")

        self.firstBody.pack()

        button_frame = CTkFrame(self.mainFrame, fg_color=self.bgcolor)
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)
        CTkButton(button_frame, text="Create Invoice Docx", command=self.create_invoice).grid(row=0, column=0, sticky="news", padx=(0,3))
        self.invoicepdf = CTkButton(button_frame, text="Create Invoice Pdf", command=self.create_invoicepdf, state="disabled")
        self.invoicepdf.grid(row=0, column=1, sticky="news", padx=(0,3))
        button_frame.pack(fill="x", pady=(10,0))
    
    def second_part(self) :
        CTkLabel(self.mainFrame, text="Money Receipt", font=self.font(30), text_color=self.headlinecolor).pack(pady=(40,20))
        self.second_body = CTkFrame(self.mainFrame, fg_color=self.bgcolor)

        CTkLabel(self.second_body, text="Consultation Fee's", text_color="#151515").grid(column=0,row=0)
        CTkLabel(self.second_body, text="Application Fee's", text_color="#151515").grid(column=0,row=1)
        CTkLabel(self.second_body, text="Payment Method", text_color="#151515").grid(column=0,row=2)

        self.consultation = CTkEntry(self.second_body, width=300, validate="key", validatecommand=self.vcmd, fg_color=self.entrycolor, text_color=self.entrytextcolor)
        self.consultation.grid(column=1,row=0)
        self.application = CTkEntry(self.second_body, width=300, validate="key", validatecommand=self.vcmd, fg_color=self.entrycolor, text_color=self.entrytextcolor)
        self.application.grid(column=1,row=1)
        method = ["E-transfer", "Bkash", "Bank transfer", "Credit Card", "Others"]
        self.payment = CTkComboBox(self.second_body, width=300, values=method, state="readonly", fg_color=self.entrycolor, text_color=self.entrytextcolor)
        self.payment.grid(column=1,row=2)

        for child in self.second_body.winfo_children():
            child.configure(font=self.font(18))
            child.grid_configure(padx=5, pady=5, sticky="w")

        self.second_body.pack()

        button_frame = CTkFrame(self.mainFrame, fg_color=self.bgcolor)
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)
        CTkButton(button_frame, text="Create Money Receipt Docx", command=self.create_receipt).grid(row=0, column=0, sticky="news", padx=(0,3))
        self.receiptpdf = CTkButton(button_frame, text="Create Money Receipt Pdf", command=self.create_receiptpdf, state="disabled")
        self.receiptpdf.grid(row=0, column=1, sticky="news", padx=(3,0))
        button_frame.pack(fill="x", pady=(10, 0))

        CTkButton(self.mainFrame, text="First page", command=self.firstScreen).pack(pady=10, fill="x")
    
    def check_exist(self, arr, rand) :
        if rand in arr :
            return True
        else :
            return False        

    def serial_no(self, type):
        r = random.randint(100000, 999999)        
        if type == "invoice" :
            if self.check_exist(self.excel.select_invoice(), r) :
                return self.serial_no("invoice")
            else :
                return r
        
        elif type == "receipt" :
            if self.check_exist(self.excel.select_receipt(), r) :
                return self.serial_no("receipt")
            else :
                return r

    def create_invoice(self):
        self.arr = [self.serial_no("invoice"), self.client.get(), self.professional.get(), self.govt.get(), self.consultation.get(), self.application.get(), self.payment.get()]
        try:
            from Helper.Invoice import Invoice
            from datetime import datetime
            arr = [self.arr[0], self.arr[1], self.arr[2], self.arr[3], datetime.today()]
            self.excel.add_invoice(arr)
            self.invoiceclass = Invoice(self.arr, self.current_path)
            messagebox.showinfo("Information", f"Invoice docx for client file no: {self.arr[1]} is done and file no: {self.arr[0]}.")
            self.invoicepdf.configure(state = "normal")
        except Exception as e:
            messagebox.showerror("Error", f"Some error happened, try again. Error: {e}")
            print(e)

    def create_invoicepdf(self):
        try :
            self.invoiceclass.save_pdf()
            messagebox.showinfo("Information", f"Invoice pdf for client file no: {self.arr[1]} is done and file no: {self.arr[0]}.")
        except Exception as e :
            messagebox.showerror("Error", f"Some error happened, try again. Error: {e}")
            print(e)

    def create_receipt(self):
        self.arr = [self.serial_no("receipt"), self.client.get(), self.professional.get(), self.govt.get(), self.consultation.get(), self.application.get(), self.payment.get()]
        try:
            from Helper.Money_Recipt import Money_Receipt
            from datetime import datetime
            arr = [self.arr[0], self.arr[1], self.arr[4], self.arr[5], self.arr[6], datetime.today()]
            self.excel.add_receipt(arr)
            self.receiptclass = Money_Receipt(self.arr, self.current_path)
            messagebox.showinfo("Information", f"Money Receipt docx for client file no: {self.arr[1]} is done and file no: {self.arr[0]}.")
            self.receiptpdf.configure(state="normal")
        except Exception as e:
            messagebox.showerror("Error", f"Some error happened, try again. Error: {e}")
            print(e)

    def create_receiptpdf(self) :
        try :
            self.receiptclass.save_pdf()
            messagebox.showinfo("Information", f"Money Receipt pdf for client file no: {self.arr[1]} is done and file no: {self.arr[0]}.")
        except Exception as e :
            messagebox.showerror("Error", f"Some error happened, try again. Error: {e}")
            print(e)

    def firstScreen(self):
        self.mainFrame.destroy()
        from Screen import FirstScreen 
        FirstScreen.First(self.window, self.current_path)