from customtkinter import CTkFrame, CTkButton, CTkLabel, CTkEntry, CTkComboBox, CTkCheckBox, CTkTextbox, BooleanVar, StringVar, IntVar, END
from tkinter import messagebox
from tkcalendar import DateEntry
from datetime import datetime

class First:
    def __init__(self, window, current_dir) :
        self.window = window        
        self.current_path = current_dir
        self.bgcolor = "#FCFBFA"
        self.headlinecolor = "#B08A6C"
        self.textcolor = "#151515"
        self.entrycolor = "#A79E87"
        self.entrytextcolor = "#151515"
        self.new_client = True
        self.mainFrame = CTkFrame(self.window, fg_color=self.bgcolor)
        self.vcmd = (self.window.register(self.validate_input), '%P')
        self.data_fetch()
        self.personalInfo()
        self.initial_consultation()
        self.milestoneInfo()
        # self.clientResponsibilities()
        self.billing()        
        self.mainFrame.pack(fill="x")

    def font(self, n) :
        return ("Arial", n)
    
    def validate_input(self, value_if_allowed):
        if value_if_allowed == "" or value_if_allowed.isdigit():
            return True
        else:
            return False
        
    def data_fetch(self) :
        frame = CTkFrame(self.mainFrame, fg_color=self.bgcolor)
        CTkLabel(frame, text="Client Number: ", font=self.font(18), text_color=self.textcolor).grid(row=0,column=0)
        from Helper.Supabase import Excel_Book
        excel = Excel_Book()
        self.client_info = excel.select_client_info()
        value = [""]+ [lst[0] for lst in self.client_info]
        self.client_combobox = CTkComboBox(frame, values=value, width=300, state="readonly", command=self.client_command, fg_color=self.entrycolor, text_color=self.entrytextcolor, font=self.font(18))
        self.client_combobox.grid(row=0, column=1)
        frame.pack(pady=(0,20), expand=True)

    def client_command(self, value) :
        self.name.delete(0, END)
        self.fname.delete(0, END)
        self.nationality.delete(0, END)
        self.dob.delete(0, END)
        self.pN.delete(0, END)
        self.email.delete(0, END)
        self.passNo.delete(0, END)
        self.address.delete(0, END)

        
        if value == "" :
            self.new_client = True
            return
                
        for data in self.client_info :
            if data[0] == value :
                self.name.insert(0, data[1])
                self.fname.insert(0, data[2])
                self.nationality.insert(0, data[3])
                self.dob.insert(0, data[4])
                self.pN.insert(0, data[5])
                self.email.insert(0, data[6])
                self.passNo.insert(0, data[7])
                self.address.insert(0, data[8])
                self.type.set(data[9])
                self.default_milestone(self.type.get())
                self.new_client = False
                return
    
    def personalInfo(self) :
        CTkButton(self.mainFrame, text="Invoice & Money Receipt", command=self.secondScreen).pack(pady=(0,20), fill="x")
        CTkLabel(self.mainFrame, text="Personal Information", font=self.font(30), text_color=self.headlinecolor).pack(pady=(40,20))

        self.body = CTkFrame(self.mainFrame, fg_color=self.bgcolor)
        
        # label for personal info
        CTkLabel(self.body, text="Name", text_color=self.textcolor).grid(column=0,row=1)
        CTkLabel(self.body, text="Family Name", text_color=self.textcolor).grid(column=0,row=2)
        CTkLabel(self.body, text="Nationality", text_color=self.textcolor).grid(column=0,row=3)
        CTkLabel(self.body, text="Date of Birth", text_color=self.textcolor).grid(column=0,row=4)
        CTkLabel(self.body, text="Phone No", text_color=self.textcolor).grid(column=0,row=5)
        CTkLabel(self.body, text="Email", text_color=self.textcolor).grid(column=0,row=6)
        CTkLabel(self.body, text="Passport No", text_color=self.textcolor).grid(column=0,row=7)
        CTkLabel(self.body, text="Address", text_color=self.textcolor).grid(column=0,row=8)

        # entry for personal info
        self.name = CTkEntry(self.body, width=300, fg_color=self.entrycolor, text_color=self.entrytextcolor)
        self.name.grid(column=1,row=1)
        self.fname = CTkEntry(self.body, width=300, fg_color=self.entrycolor, text_color=self.entrytextcolor)
        self.fname.grid(column=1,row=2)
        self.nationality = CTkEntry(self.body, width=300, fg_color=self.entrycolor, text_color=self.entrytextcolor)
        self.nationality.grid(column=1,row=3)
        self.dob = DateEntry(self.body, width=32, font=self.font(20), state="readonly", fg_color=self.entrycolor, text_color=self.entrytextcolor)
        self.dob.grid(column=1,row=4)
        self.pN = CTkEntry(self.body, width=300, fg_color=self.entrycolor, text_color=self.entrytextcolor)
        self.pN.grid(column=1,row=5)
        self.email = CTkEntry(self.body, width=300, fg_color=self.entrycolor, text_color=self.entrytextcolor)
        self.email.grid(column=1,row=6)
        self.passNo = CTkEntry(self.body, width=300, fg_color=self.entrycolor, text_color=self.entrytextcolor)
        self.passNo.grid(column=1,row=7)
        self.address = CTkEntry(self.body, width=300, fg_color=self.entrycolor, text_color=self.entrytextcolor)
        self.address.grid(column=1,row=8)

        for child in self.body.winfo_children():
            child.configure(font=self.font(18))
            child.grid_configure(padx=5 , pady=5 , sticky="w")

        self.body.pack()
    
    def application_value(self):
        from Helper.Excel_control import Milestone_Book
        self.ml = Milestone_Book(self.current_path)
        self.info = self.ml.select()
        name = list()
        for ele in self.info :            
            name.append(ele[0])
        return name
    
    def initial_consultation(self) :
        CTkLabel(self.mainFrame, text="Initial Consultation", font=self.font(30), text_color=self.headlinecolor).pack(pady=(40,20))
        self.consultationBody = CTkFrame(self.mainFrame, fg_color=self.bgcolor)

        # bill label
        CTkLabel(self.consultationBody, text="Time", text_color=self.textcolor).grid(column=0, row=0)
        CTkLabel(self.consultationBody, text="Payment", text_color=self.textcolor).grid(column=0, row=1)
        CTkLabel(self.consultationBody, text="Application Type", text_color=self.textcolor).grid(column=0, row=2)

        # bill entry
        self.time = CTkEntry(self.consultationBody, width=300, validate="key", validatecommand=self.vcmd, fg_color=self.entrycolor, text_color=self.entrytextcolor)
        self.time.grid(column=1, row=0)
        self.pay = CTkEntry(self.consultationBody, width=300, validate="key", validatecommand=self.vcmd, fg_color=self.entrycolor, text_color=self.entrytextcolor)
        self.pay.grid(column=1, row=1)
        self.type = CTkComboBox(self.consultationBody, values=self.application_value(), width=300, state="readonly", command=self.default_milestone, fg_color=self.entrycolor, text_color=self.entrytextcolor)
        self.type.grid(column=1,row=2)
        self.type.set("--Select--")

        for child in self.consultationBody.winfo_children():
            child.configure(font=self.font(18))
            child.grid_configure(padx=5 , pady=5 , sticky="w")  

        self.consultationBody.pack()

        button_frame = CTkFrame(self.mainFrame, fg_color=self.bgcolor)
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)
        CTkButton(button_frame, text="Create Initial Consultation Docx", command=self.create_initialconsultation).grid(row=0, column=0, sticky="news", padx=(0,3))
        self.initalpdf = CTkButton(button_frame, text="Create Initial Consultation Pdf", command=self.create_initialconsultationpdf, state="disabled")
        self.initalpdf.grid(row=0, column=1, sticky="news", padx=(3,0))
        button_frame.pack(pady=20, fill="x")

        CTkLabel(self.mainFrame, text="RCIC Responsibilities and Commitments", font=self.font(30), text_color=self.headlinecolor).pack(pady=(40,20))

        self.dmilestoneBody_info = CTkFrame(self.mainFrame, fg_color=self.bgcolor, height=0)
        self.dmilestoneBody_info.pack()
        self.dmilestoneBody = CTkFrame(self.dmilestoneBody_info, fg_color=self.bgcolor, height=0)
        self.dmilestoneBody.pack()

        self.milestoneBody = CTkFrame(self.mainFrame, fg_color=self.bgcolor)
        self.milestoneBody.pack()
        
        CTkLabel(self.mainFrame, text="Client Responsibilities and Commitments", font=self.font(30), text_color=self.headlinecolor).pack(pady=(40,20))

        self.responsibeBody_info = CTkFrame(self.mainFrame, fg_color=self.bgcolor)
        self.responsibeBody_info.pack(fill="x")

        self.responsibeBody = CTkFrame(self.responsibeBody_info, fg_color=self.bgcolor, height=0)
        self.responsibeBody.pack()
    
    def default_milestone(self, type) :
        self.dmilestoneBody.destroy()
        self.dmilestoneBody = CTkFrame(self.dmilestoneBody_info, fg_color=self.bgcolor, height=0)
        self.dmilestoneBody.pack()         
        self.dmchecklst = list()
        self.dmdatelst = list()
        self.dmprofessionallst = list()
        self.dmadminlst = list()

        self.row = list()
        for ele in self.info :
            if ele[0] == type :
                self.row = ele

        for i in range(2, len(self.row)) :
            check = BooleanVar()
            check.set(True)
            CTkCheckBox(self.dmilestoneBody, text=f"Milestone-{i-1}", variable=check, font=self.font(16), text_color=self.textcolor).pack(anchor="w")
            self.dmchecklst.append(check)
            CTkLabel(self.dmilestoneBody, text=self.row[i], text_color=self.textcolor, justify="left").pack(anchor="w")

            CTkLabel(self.dmilestoneBody, text="Date : ", text_color=self.textcolor).pack(anchor="w")
            date = DateEntry(self.dmilestoneBody, width=27, font=self.font(20), state="readonly", fg_color=self.entrycolor, text_color=self.entrytextcolor)
            date.pack()
            self.dmdatelst.append(date)

            CTkLabel(self.dmilestoneBody, text="Professional Fee : ", text_color=self.textcolor).pack(anchor="w")
            professional = CTkEntry(self.dmilestoneBody, width=300, font=self.font(18), validate="key", validatecommand=self.vcmd, fg_color=self.entrycolor, text_color=self.entrytextcolor)
            professional.pack()
            self.dmprofessionallst.append(professional)
            
            CTkLabel(self.dmilestoneBody, text="Govt. Fee : ", text_color=self.textcolor).pack(anchor="w")
            admin = CTkEntry(self.dmilestoneBody, width=300, font=self.font(18), validate="key", validatecommand=self.vcmd, fg_color=self.entrycolor, text_color=self.entrytextcolor)
            admin.pack(pady=(0,10))
            self.dmadminlst.append(admin)
        
            self.clientResponsibilities()
    
    def default_milestoneinfo(self) :
        self.dmlst = list()
        for i in range(2, len(self.row)) :
            if self.dmchecklst[i-2].get() == True :
                lst = list()
                lst.append(self.row[i])
                lst += [self.dmdatelst[i-2].get_date().strftime("%d/%m/%Y"), self.dmprofessionallst[i-2].get(), self.dmadminlst[i-2].get()]                
                self.dmlst.append(lst)

    def milestoneInfo(self) :   
        # milestone label
        CTkLabel(self.milestoneBody, text="Extra Milestones", text_color=self.textcolor).grid(column=0, row=0)    
        CTkLabel(self.milestoneBody, text="Milestone", text_color=self.textcolor).grid(column=0, row=1)
        CTkLabel(self.milestoneBody, text="Date", text_color=self.textcolor).grid(column=0, row=2)
        CTkLabel(self.milestoneBody, text="Professional Fees", text_color=self.textcolor).grid(column=0, row=3)
        CTkLabel(self.milestoneBody, text="Govt. Fees", text_color=self.textcolor).grid(column=0, row=4)

        # milestone entry        
        self.milestone = CTkTextbox(self.milestoneBody, width=300, border_width=1, fg_color=self.entrycolor, text_color=self.entrytextcolor)
        self.milestone.grid(column=1,row=1)
        self.date = DateEntry(self.milestoneBody, width=32, font=self.font(20), state="readonly", fg_color=self.entrycolor, text_color=self.entrytextcolor)
        self.date.grid(column=1,row=2)
        self.professional_mile = CTkEntry(self.milestoneBody, width=300, validate="key", validatecommand=self.vcmd, fg_color=self.entrycolor, text_color=self.entrytextcolor)
        self.professional_mile.grid(column=1, row=3)
        self.adminstrative_mile = CTkEntry(self.milestoneBody, width=300, validate="key", validatecommand=self.vcmd, fg_color=self.entrycolor, text_color=self.entrytextcolor)
        self.adminstrative_mile.grid(column=1, row=4)

        for child in self.milestoneBody.winfo_children():
            child.configure(font=self.font(18))
            child.grid_configure(padx=5 , pady=5 , sticky="wn")

        # add milestone button
        self.milestonestr = ""
        self.milestone_arr = list()
        CTkButton(self.milestoneBody, text="Add Milestone", command=self.add_milestone).grid(column=0, row=5, columnspan=2, sticky="news", pady=10)

        self.milestoneLabel = CTkLabel(self.milestoneBody, text="Added Milestones", justify="left", text_color=self.textcolor)
        self.milestoneLabel.grid(column=0, row=6, columnspan=2, sticky="wn")


    def add_milestone(self) :      
        lst = self.milestone.get("1.0","end-1c").split("\n")
        s = ""
        for ele in lst :
            s += f"\t\u25cf\t{ele}\n"
        self.milestone_arr.append([s,self.date.get(),self.professional_mile.get(),self.adminstrative_mile.get()])
        self.milestonestr += f"Milestone\n"
        self.milestonestr += f"Date: {self.date.get()}; Professional Fee: {self.professional}; Govt. Fee {self.adminstrative.get()} \n"
        for ele in lst :
            self.milestonestr += f"\t\u25cf\t{ele}\n"
        self.milestone.delete("1.0","end-1c")
        self.date.set_date(datetime.today().date())
        self.professional_mile.delete(0,END)
        self.adminstrative_mile.delete(0,END)
        self.milestoneLabel.configure(text= self.milestonestr)
    
    def clientResponsibilities(self) :
        self.responsibeBody.destroy()
        self.responsibeBody = CTkFrame(self.responsibeBody_info, fg_color=self.bgcolor)        

        self.responsible = self.row[1].split("\n")
        self.responsibecheck = list()
        
        for i in range(len(self.responsible)-1) :
            check = BooleanVar()
            check.set(True)
            self.responsibecheck.append(check)
            CTkCheckBox(self.responsibeBody, text=self.responsible[i], variable=check).pack()
                
        CTkLabel(self.responsibeBody, text="Add more", font=self.font(18), text_color=self.textcolor).pack()
        self.add = CTkEntry(self.responsibeBody, fg_color=self.entrycolor, text_color=self.entrytextcolor)
        self.add.pack()
        self.add_arr = list()
        CTkButton(self.responsibeBody, text="Add", command=self.add_responsibilities).pack()
        
        self.add_str_data = "Extra Responsibilities\n"
        self.addLabel = CTkLabel(self.responsibeBody, text=self.add_str_data, justify="left", text_color=self.textcolor)
        self.addLabel.pack()

        for child in self.responsibeBody.winfo_children() :
            child.pack_configure(padx=5, pady=3, fill="x", anchor="w")

        self.responsibeBody.pack(fill="x")
    
    def add_responsibilities(self) :
        self.add_arr.append(self.add.get())
        self.add_str_data += f"{self.add.get()}\n"
        self.addLabel.configure(text=self.add_str_data)
        self.add.delete(0, END) 

    def billing(self) :
        CTkLabel(self.mainFrame, text="Billing Info", font=self.font(30), text_color=self.headlinecolor).pack(pady=(40,20))
        self.billing_body = CTkFrame(self.mainFrame, fg_color=self.bgcolor)

        CTkLabel(self.billing_body, text="Administrative Fees", text_color=self.textcolor).grid(column=0, row=1)
        
        self.adminstrative = CTkEntry(self.billing_body, width=300, validate="key", validatecommand=self.vcmd, fg_color=self.entrycolor, text_color=self.entrytextcolor)
        self.adminstrative.grid(column=1, row=1)

        for child in self.billing_body.winfo_children():
            child.configure(font=self.font(18))
            child.grid_configure(padx=5 , pady=5 , sticky="w")

        self.billing_body.pack()
        
        self.reference_body = CTkFrame(self.mainFrame, height=0, fg_color=self.bgcolor)
        self.refer = CTkCheckBox(self.reference_body, text="Reference", command=self.reference, font=self.font(30), text_color=self.headlinecolor)
        self.refer.pack(pady=(20,20))
        self.reference_body.pack()
        self.refer_info_body = CTkFrame(self.reference_body, height=0, fg_color=self.bgcolor)
        self.refer_info_body.pack()

        button_frame = CTkFrame(self.mainFrame, fg_color=self.bgcolor)
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)
        CTkButton(button_frame, text="Create Service Agreement Docx", command=self.create_serviceagreement).grid(row=0, column=0, sticky="news", padx=(0,3))
        self.servicepdf = CTkButton(button_frame, text="Create Service Agreement Pdf", command=self.create_serviceagreementpdf, state="disabled")
        self.servicepdf.grid(row=0, column=1, sticky="news", padx=(3,0))
        button_frame.pack(fill="x", pady=10)

    def reference(self) :
        if self.refer.get() == True :
            self.refer_info_body = CTkFrame(self.reference_body, fg_color=self.bgcolor)

            CTkLabel(self.refer_info_body, text="Name", text_color=self.textcolor).grid(row=1, column=0)
            CTkLabel(self.refer_info_body, text="Email", text_color=self.textcolor).grid(row=2, column=0)
            CTkLabel(self.refer_info_body, text="Date of Birth", text_color=self.textcolor).grid(row=3, column=0)
            CTkLabel(self.refer_info_body, text="Address", text_color=self.textcolor).grid(row=4, column=0)

            self.refername = CTkEntry(self.refer_info_body, width=300, fg_color=self.entrycolor, text_color=self.entrytextcolor)
            self.refername.grid(row=1, column=1)
            self.referemail = CTkEntry(self.refer_info_body, width=300, fg_color=self.entrycolor, text_color=self.entrytextcolor)
            self.referemail.grid(row=2, column=1)
            self.referdob = DateEntry(self.refer_info_body, width=32, font=self.font(20), state="readonly", fg_color=self.entrycolor, text_color=self.entrytextcolor)
            self.referdob.grid(row=3, column=1)
            self.referaddress = CTkEntry(self.refer_info_body, width=300, fg_color=self.entrycolor, text_color=self.entrytextcolor)
            self.referaddress.grid(row=4, column=1)

            for child in self.refer_info_body.winfo_children():
                child.configure(font=self.font(18))
                child.grid_configure(padx=5 , pady=5 , sticky="w")
        
            self.refer_info_body.pack()
        else :
            self.refer_info_body.destroy()

    def clientResponsibilities_data(self) :
        self.responsibilities_arr = list()
        self.responsibilities_str = ""

        for i in range(len(self.responsibecheck)) :
            if self.responsibecheck[i].get() == True :
                self.responsibilities_arr.append(self.responsible[i])
        
        self.responsibilities_arr += self.add_arr
        
        for ele in self.responsibilities_arr :
            self.responsibilities_str += f"{ele}\n"
        for ele in self.add_arr :
            self.responsibilities_str += f"{ele}\n"
        
    def client_file_number(self) :
        number = ""
        number += self.name.get()[-3:-1]        
        number += self.passNo.get()[-2:]
        number +=  datetime.today().strftime("%d%y")
        number = number.upper()
        return number

    def create_initialconsultation(self) :
        try :
            self.client_number = self.client_file_number()
            if self.new_client :
                self.excel_arr = [self.client_number,self.name.get(),self.fname.get(),self.nationality.get(),self.dob.  get_date().strftime("%d/%m/%Y"),self.pN.get(),self.email.get(),self.passNo.get(),self.address.get(),self.type.get(),0,0,datetime.today()]
            else :
                self.excel_arr = [self.client_combobox.get(),self.name.get(),self.fname.get(),self.nationality.get(),self.dob.  get_date().strftime("%d/%m/%Y"),self.pN.get(),self.email.get(),self.passNo.get(),self.address.get(),self.type.get(),0,0,datetime.today()]
                print(self.client_combobox.get())

            self.initial_consultation_arr = [self.client_file_number(),self.name.get(),self.fname.get(),self.nationality.get(),self.dob.get_date().strftime("%d/%m/%Y"),self.pN.get(),self.email.get(),self.passNo.get(),self.address.get(),self.type.get(),self.time.get(),self.pay.get()]
            
            from Helper.Supabase import Excel_Book
            self.excel = Excel_Book()
            if self.new_client :
                self.excel.add_info(self.excel_arr)
            else :
                self.excel.update_info(self.excel_arr)

            from Helper.Initial_Consultation_Agreement import Initial_Consultation
            self.initialclass = Initial_Consultation(self.initial_consultation_arr, self.current_path)
            messagebox.showinfo("Information", f"Your Initial Consultation Agreement docx document created.\nFile name {self.client_file_number()}.")
            self.initalpdf.configure(state="normal")
        except Exception as e:
            messagebox.showerror("Error", f"Some error happend try again({e})")
            print(e)

    def create_initialconsultationpdf(self) :
        try :                     
            self.initialclass.save_pdf()
            messagebox.showinfo("Information", f"Your Initial Consultation Agreement pdf document created.\nFile name {self.client_file_number()}.")
        except Exception as e:
            messagebox.showerror("Error", f"Some error happend try again({e})")
            print(e)

    def create_serviceagreement(self) :
        self.default_milestoneinfo()
        arr = self.dmlst + self.milestone_arr
        self.professional = 0
        self.govt = 0
        for ele in arr :
            self.professional += int(ele[-2])
        
        print(self.professional)

        self.clientResponsibilities_data()
        self.client_number = self.client_file_number()
        if self.new_client :
            self.excel_arr = [self.client_number,self.name.get(),self.fname.get(),self.nationality.get(),self.dob.  get_date().strftime("%d/%m/%Y"),self.pN.get(),self.email.get(),self.passNo.get(),self.address.get(),self.type.get(),self.professional,self.adminstrative.get(),datetime.today()]
        else :
            self.excel_arr = [self.client_combobox.get(),self.name.get(),self.fname.get(),self.nationality.get(),self.dob.  get_date().strftime("%d/%m/%Y"),self.pN.get(),self.email.get(),self.passNo.get(),self.address.get(),self.type.get(),self.professional,self.adminstrative.get(),datetime.today()]
            print(self.professional)

        if self.refer.get() :
            self.service_agreement_arr = [self.client_number,self.name.get(),self.fname.get(),self.nationality.get(),   self.dob.get_date().strftime("%d/%m/%Y"),self.pN.get(),self.email.get(),self.passNo.get(),self.address.get(),self.type.get(),self.professional,self.adminstrative.get(),self.responsibilities_arr,arr,self.refer.get(),self.refername.get(),self.referemail.get(),self.referdob.get_date().strftime("%d/%m/%Y"),self.referaddress.get()]
        else :
            self.service_agreement_arr = [self.client_number,self.name.get(),self.fname.get(),self.nationality.get(),   self.dob.get_date().strftime("%d/%m/%Y"),self.pN.get(),self.email.get(),self.passNo.get(),self.address.get(),self.type.get(),   self.professional,self.adminstrative.get(),self.responsibilities_arr,arr,self.refer.get()]
        try :
            from Helper.Supabase import Excel_Book
            self.excel = Excel_Book()
            if self.new_client :
                self.excel.add_info(self.excel_arr)
            else :
                self.excel.update_info(self.excel_arr)
            from Helper.Service_Agreement import Service_Agreement
            self.serviceclass = Service_Agreement(self.service_agreement_arr, self.current_path)            
            messagebox.showinfo("Information", f"Your input entered in Excel Sheet, Service Agreement docx document created.\nFile name {self.client_number}.")
            self.servicepdf.configure(state = "normal")
        except Exception as e:
            messagebox.showerror("Error", f"Some error happend try again({e})")
            print(e)
        
    def create_serviceagreementpdf(self) :
        try :
            self.serviceclass.save_pdf()
            messagebox.showinfo("Information", f"Service Agreement pdf document created.\nFile name {self.client_number}.")
        except Exception as e :            
            messagebox.showerror("Error", f"Some error happend try again({e})")
            print(e)

    def secondScreen(self) :        
        self.mainFrame.destroy()
        from Screen import SecondScreen
        SecondScreen.Second(self.window , self.current_path)
