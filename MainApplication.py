from os import path, makedirs
from customtkinter import CTk, CTkFrame, CTkScrollableFrame, CTkImage, CTkLabel, CTkButton, CTkToplevel, filedialog
from PIL import Image
from Screen import FirstScreen, Customize
import requests

class Application:
    def __init__(self):
        # Get the current directory (where the script/exe is located)
        self.current_dir = ""
        print(self.current_dir)
        
        # Create the window
        self.window = CTk()
        self.width = 850
        self.height = 550
        self.gapx = 200
        self.gapy = 100
        # self.width = self.window.winfo_screenwidth()
        # self.height = self.window.winfo_screenheight()        
        # self.gapx = -20
        # self.gapy = -5
        self.bgcolor = "#FCFBFA"
        self.bgcolor2 = "#FCFBFA"
        self.window.geometry(f"{self.width}x{self.height}+{self.gapx}+{self.gapy}")
        self.window.resizable(False, False)
        # self.window.state("zoomed")
        # # self.window.update_idletasks()
        self.window.configure(fg_color=self.bgcolor2)
        self.window.title("Integration")
        self.window.iconbitmap(f"assets/icon.ico")

        if not path.exists("path_file.txt") :
            top = CTkToplevel(self.window)
            self.f = filedialog.askdirectory()
            CTkLabel(top, text=self.f).pack()
            CTkButton(top, text="Path", command=self.path_file).pack(fill="x")
            
            
        else :
            f = open("path_file.txt", "rt")
            l = f.readline()
            print(l)
            self.folder_existance()
            self.header()
            self.mainFrame = CTkScrollableFrame(self.window, width=self.width, height=self.height, fg_color=self.bgcolor2)
            self.mainFrame.pack()
            
            if self.is_connected() :
                FirstScreen.First(self.mainFrame,self.current_dir)
            
            else :
                CTkLabel(self.mainFrame, text="Unable to connect with server.").pack()
        # SecondScreen.Second(self.mainFrame,self.current_dir)
        self.window.mainloop()

    # making the header part with logo , refresh and customization button
    def is_connected(self) :
        try:        
            requests.get("https://www.google.com", timeout=5)
            return True
        except requests.ConnectionError:
            return False
        
    def header(self):
        self.header_body = CTkFrame(self.window, width=850, height=50, fg_color=self.bgcolor, corner_radius=0)
        # Use path.join to construct paths
        image_path = path.join( 'assets', 'logo.png')    
        img = Image.open(image_path)
        img = CTkImage(img, size=(300, 115))
        
        CTkLabel(self.header_body, text="", image=img).pack(side="left")

        # customization part
        settings_path = path.join( 'assets', 'settings.png')
        simg = Image.open(settings_path)
        simg = CTkImage(simg, size=(30,30))

        setting = CTkLabel(self.header_body, text="", image=simg, cursor="hand2")
        setting.pack(padx=(0,30), side="right")
        setting.bind("<Button-1>", self.customization_command)

        # reload part
        reset_path = path.join( 'assets', 'reload.png')
        rimg = Image.open(reset_path)
        rimg = CTkImage(rimg, size=(30,30))

        reset = CTkLabel(self.header_body, text="", image=rimg, cursor="hand2")
        reset.pack(padx=(0,30), side="right")
        reset.bind("<Button-1>", self.reset_command)

        self.header_body.pack(padx=10, pady=10, ipadx=850)

    # reload command
    def reset_command(self, event) :
        self.window.destroy()
        Application()
    
    # customization command
    def customization_command(self, event) :
        top = CTkToplevel(self.window)
        top.geometry(f"{self.width}x{self.height}+{self.gapx}+{self.gapy}")
        # top.wm_attributes("-topmost", 1)
        Customize.Customization(top, self.current_dir)
    
    # DocFile and necessary files in there if not exist will create automatically
    def folder_existance(self) :
        if path.exists("path_file.txt") :
            f = open("path_file.txt", "rt")
            p = f.readline()

        docfile = path.join( p, 'DocFile')
        if not path.exists(docfile) :
            makedirs(docfile)
        
        initial = path.join( p, 'DocFile', 'Initial Consultation Agreement')
        if not path.exists(initial) :
            makedirs(initial)
        
        invoice = path.join( p, 'DocFile', 'Invoice')
        if not path.exists(invoice) :
            makedirs(invoice)

        money = path.join( p, 'DocFile', 'Money Receipt')
        if not path.exists(money) :
            makedirs(money)

        service = path.join( p, 'DocFile', 'Service Agreement')
        if not path.exists(service) :
            makedirs(service)
        
        sheet = path.join( p, 'DocFile', 'Sheet')
        if not path.exists(sheet) :
            makedirs(sheet)
    
    def path_file(self) :
        try: 
            f = open("path_file.txt", "wt")
            f.write(self.f)
            f.close()
        except Exception as e:
            self.window.destroy()
        
        self.window.destroy()



# calling the main App
if __name__ == "__main__" :
    Application()
