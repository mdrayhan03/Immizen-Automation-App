from customtkinter import *
from tkinter import messagebox

class Customization :
    def __init__(self, win, current_dir) :
        self.window = win
        self.bgcolor = "#FCFBFA"
        self.headlinecolor = "#B08A6C"
        self.textcolor = "#151515"
        self.entrycolor = "#A79E87"
        self.entrytextcolor = "#151515"
        self.current_dir = current_dir
        self.main = CTkFrame(self.window, fg_color=self.bgcolor)
        self.mainframe = CTkScrollableFrame(self.main, width=850, height=550, fg_color=self.bgcolor)
        self.vcmd = (self.window.register(self.validate_input), '%P')
        self.info = self.type_name()
        self.first_part()
        self.second_part()
        self.button_part()        
        self.mainframe.pack()
        self.main.pack()

    def font(self, n) :
        return ("Arial", n)
    
    def validate_input(self, value_if_allowed):
        if value_if_allowed == "" or value_if_allowed.isdigit():
            return True
        else:
            return False
        
    def type_name(self) :
        from Helper.Excel_control import Milestone_Book
        self.ml = Milestone_Book(self.current_dir)

        return self.ml.select()        

    def first_part(self) :       
        self.first_body = CTkFrame(self.mainframe, fg_color=self.bgcolor)

        CTkLabel(self.first_body, text="Add New Milestone",font=self.font(30), text_color=self.headlinecolor).grid(row=0, column=0, columnspan=2)

        CTkLabel(self.first_body, text="Type", text_color=self.textcolor).grid(row=1, column=0, sticky="w")
        CTkLabel(self.first_body, text="Milestone", text_color=self.textcolor).grid(row=3, column=0, sticky="wn")

        type = [""]
        for ele in self.info :
            type.append(ele[0])

        self.type = CTkComboBox(self.first_body, width=300, values=type, command=self.type_command, fg_color=self.entrycolor, text_color=self.entrytextcolor)
        self.type.grid(row=1, column=1)
        self.milestonearr = list()
        self.milestone = CTkTextbox(self.first_body, width=300, border_width=1, fg_color=self.entrycolor, text_color=self.entrytextcolor)
        self.milestone.grid(row=3, column=1)
        CTkButton(self.first_body, text="Add Milestone", command=self.milestone_command).grid(row=4, column=1, sticky="news")

        self.milestonebody = CTkFrame(self.first_body, height=0, fg_color=self.bgcolor)
        self.milestonebody.grid(row=5, column=0, columnspan=2, sticky="w")

        for child in self.first_body.winfo_children() :
            child.grid_configure (padx=10, pady=5)

        self.first_body.pack()

    def type_command(self, type) :
        if type == "" :
            self.milestonearr = list()
            self.milestonebody.destroy()
            self.milestonebody = CTkFrame(self.first_body, height=0, fg_color=self.bgcolor)
            self.milestonebody.grid(row=5, column=0, columnspan=2, sticky="w")

            self.responsibilityarr = list()
            self.responsibility_body.destroy()
            self.responsibility_body = CTkFrame(self.second_body, height=0, fg_color=self.bgcolor)
            self.responsibility_body.pack()

            return
        
        row = list()
        for ele in self.info :
            if ele[0] == type :
                row = ele
        
        arr = row[1].split("\n")
        self.responsibilityarr = arr[:-1]
        
        self.milestonearr = list()
        for i in range(2, len(row)) :
            self.milestonearr.append(row[i])

        self.milestone_show()
        self.responsibility_show()

    def milestone_command(self) :
        s = ""        
        lst = self.milestone.get("1.0","end-1c").split("\n")
        for ele in lst :
            s += f"\t\u25cf\t{ele}\n"

        self.milestonearr.append(s)
        self.milestone.delete("1.0", "end")
        self.milestone_show()
    
    def milestone_show(self) :
        self.milestonebody.destroy()
        self.milestonebody = CTkFrame(self.first_body, height=0, fg_color=self.bgcolor)
        self.milestonebody.grid(row=5, column=0, columnspan=2, sticky="w")

        self.milestonecheckbox = list()

        for i in range(len(self.milestonearr)) :
            s = BooleanVar()
            s.set(True)
            CTkCheckBox(self.milestonebody, text=f"Milestone-{i+1}", variable=s).pack(anchor="w")
            CTkLabel(self.milestonebody, text=self.milestonearr[i], wraplength=400, justify="left", text_color=self.textcolor).pack(anchor="w")

            self.milestonecheckbox.append(s)

    def second_part(self) :
        self.second_body = CTkFrame(self.mainframe, fg_color=self.bgcolor)

        CTkLabel(self.second_body, text="Add Client Responsibility", font=self.font(30), text_color=self.headlinecolor).pack(pady=5, fill="x")

        self.responsibilityarr = list()
        self.responsibility = CTkEntry(self.second_body, fg_color=self.entrycolor, text_color=self.entrytextcolor) 
        self.responsibility.pack(pady=5, fill="x")

        CTkButton(self.second_body, text="Add Responsibility", command=self.responsibility_command).pack(pady=5, fill="x")

        self.responsibility_body = CTkFrame(self.second_body, height=0, fg_color=self.bgcolor)
        self.responsibility_body.pack()

        self.second_body.pack(fill="x")

    def responsibility_command(self) :
        self.responsibilityarr.append(self.responsibility.get())
        self.responsibility.delete(0, END)
        self.responsibility_show()
    
    def responsibility_show(self) :
        self.responsibility_body.destroy()
        self.responsibility_body = CTkFrame(self.second_body, fg_color=self.bgcolor)
        self.responsibility_body.pack(fill="x")
        self.responsibilitycheckbox = list()

        for i in range(len(self.responsibilityarr)) :
            s = BooleanVar()
            s.set(True)
            CTkCheckBox(self.responsibility_body, text="", variable=s).grid(row=i, column=0, sticky="en")
            CTkLabel(self.responsibility_body, text=self.responsibilityarr[i], wraplength=700, justify="left", text_color=self.textcolor).grid(row=i, column=1, sticky="w")

            for child in self.responsibility_body.winfo_children() :
                child.grid_configure(pady=3)

            self.responsibilitycheckbox.append(s)

    def button_part(self) :
        CTkButton(self.mainframe, text="Save Milestone & Client Responsibility", command=self.save_command).pack(fill="x")

    def save_command(self) :
        milestonearr = list()
        for i in range(len(self.milestonearr)) :
            if self.milestonecheckbox[i].get() == True :
                milestonearr.append(self.milestonearr[i])

        responsibilitystr = ""
        for i in range(len(self.responsibilityarr)) :
            if self.responsibilitycheckbox[i].get() == True :
                responsibilitystr += f"{self.responsibilityarr[i]}\n"        
        try :
            if (any(self.type.get() in row for row in self.info)) :
                arr = [self.type.get(), responsibilitystr, milestonearr]
                self.ml.update(arr)
                messagebox.showinfo("Save Successful", "Milestone update & save successfully")
            else :
                arr = [self.type.get(), responsibilitystr, milestonearr]
                self.ml.insert(arr)
                messagebox.showinfo("Save Successful", "New milestone save successfully")
            self.main.destroy()
            Customization(self.window, self.current_dir)
            
        except Exception as e :
            messagebox.showerror("Error", f"{e}")
            print(e)
