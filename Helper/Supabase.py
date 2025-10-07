from supabase import create_client, Client
from datetime import datetime
from dotenv import load_dotenv
import os

load_dotenv()

class Excel_Book :
    def __init__(self) -> None:
        url = os.getenv("API_URL")
        key = os.getenv("API_KEY")

        self.base : Client = create_client(url, key)
        # self.base.auth.sign_in_anonymously()

    def add_info(self, arr) :
        print("add info")
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

    def update_info(self, arr) :
        data = {                        
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
            # "initial_date" : arr[12].isoformat()
        }
        print(arr[10], arr[11])
        response = self.base.table("tb_immigration").update(data).eq("client_file_number" , arr[0]).execute()
        print(response)
        return response
    

    def select_info(self) :
        response = self.base.table("tb_immigration").select("client_file_number").execute()
        arr = list()
        for ele in response.data :
            arr.append(ele["client_file_number"])
        
        return arr

    def select_client_info(self) :
        response = self.base.table("tb_immigration").select("*").execute()
        info = list()
        for ele in response.data :
            arr = list()                    
            arr.append(ele["client_file_number"])
            arr.append(ele["name"])
            arr.append(ele["family_name"])
            arr.append(ele["nationality"])
            arr.append(ele["dob"])
            arr.append(ele["pN"])
            arr.append(ele["email"])
            arr.append(ele["passport"])
            arr.append(ele["address"])
            arr.append(ele["application_type"])
            info.append(arr)
        return info
        

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
