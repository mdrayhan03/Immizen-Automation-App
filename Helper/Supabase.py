from supabase import create_client, Client

class Excel_Book :
    def __init__(self) -> None:
        url = "https://aidjhvfvbyudzrzduwag.supabase.co"
        key = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImFpZGpodmZ2Ynl1ZHpyemR1d2FnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3MjM1NTQ4NDAsImV4cCI6MjAzOTEzMDg0MH0.A7otTKYyjRPA7bVxgcFotw7UswxmWJytnzddYcZyLk4"

        self.base : Client = create_client(url, key)
        # self.base.auth.sign_in_anonymously()

    def add_info(self, arr) :
        data = {
            "client_file_number" : arr[0] ,
            "name" : arr[1] ,
            "family_name" : arr[2] ,
            "nationality" : arr[3] ,
            "dob" : arr[4] ,
            "pN" : arr[5] ,
            "email" : arr[6] ,
            "passport" : arr[7] ,
            "address" : arr[8] ,
            "application_type" : arr[9] ,
            "professional_fee" : arr[10] ,
            "adminstrative_fee" : arr[11] ,
            "initial_date" : arr[12].isoformat()
        }
        response = self.base.table("tb_immigration").insert(data).execute()
        return response

    def select_info(self) :
        response = self.base.table("tb_immigration").select("client_file_number").execute()
        arr = list()
        for ele in response.data :
            arr.append(ele["client_file_number"])
        
        return arr

    def add_invoice(self, arr) :
        data = {
            "serial_no" : arr[0] ,
            "client_file_number" : arr[1] ,
            "professional_fee" : arr[2] ,
            "govt_fee" : arr[3] ,
            "date" : arr[4].isoformat()
        }
        response = self.base.table("tb_invoice").insert(data).execute()
        return response

    def select_invoice(self) :
        response = self.base.table("tb_invoice").select("serial_no").execute()
        arr = list()
        for ele in response.data :
            arr.append(ele["serial_no"])
        
        return arr
    
    def add_receipt(self, arr) :
        data = {
            "serial_no" : arr[0] ,
            "client_file_number" : arr[1] ,
            "consultation_fee" : arr[2] ,
            "application_fee" : arr[3] ,
            "method" : arr[4] ,
            "date" : arr[5].isoformat()
        }
        response = self.base.table("tb_receipt").insert(data).execute()
        return response

    def select_receipt(self) :
        response = self.base.table("tb_receipt").select("serial_no").execute()
        arr = list()
        for ele in response.data :
            arr.append(ele["serial_no"])
        
        return arr

ex = Excel_Book()
arr = ["None", "None", "None", "None", "None", "None", "None", "None", "None", "None", 1, 1]
# ex.add_info(arr)
ex.select_invoice()