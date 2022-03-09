import re
import subprocess
import time
import os
from tkinter import *

try:
    import pandas as pd
except:
    subprocess.call('pip3 install pandas', shell=True)
    import pandas as pd


class Trafic():

    def __init__(self):

        self.first_date = ''
        self.second_date = ''
        self.dataframe = ''
        self.root = Tk()

    def entry_date(self):

        
        self.root.title("Search excel")
        self.root.geometry("400x300")

        start = StringVar()
        end = StringVar()

        start.set("2018/10/10 23:59")
        end.set("2019/01/09 23:59")

        h1 = Label(self.root, text="   Start Date ej: ")
        h1.place(x=10, y=10)
        input1 =Entry(self.root, textvariable= start)
        input1.place(x=130, y=10)
                
        h2 = Label(self.root, text="   Finish Date ej: ")
        h2.place(x=10, y=40)
        input2 = Entry(self.root, textvariable= end)
        input2.place(x=130, y=40)

        press = Button(self.root, text="Show excel", command = self.root.quit )
        press.place(x=20, y=100)

        self.root.mainloop()

        self.first_date = start.get()
        self.second_date = end.get()
       
        
    def get_date(self):

        while True:

            try:

                self.entry_date()

                validate_first= bool(re.search('^20[0-9]{2}/(0\d|1[0-2])/(0\d|1[0-9]|2[0-9]|3[0-1])(\s)([0-1][1-9]|[2][0-3])(:)([0-5][0-9])$', self.first_date))
                validate_second=bool(re.search('^20[0-9]{2}/(0\d|1[0-2])/(0\d|1[0-9]|2[0-9]|3[0-1])(\s)([0-1][1-9]|[2][0-3])(:)([0-5][0-9])$',self.second_date))

                date1 = time.strptime(self.first_date, "%Y/%m/%d %H:%M")
                date2 = time.strptime(self.second_date,  "%Y/%m/%d %H:%M")      

                if validate_first and validate_second and date1 < date2:
                    return True

            except ValueError as v:
                print('Error in get_date function: ', v)       

    def dataframe_handler(self):

        try:
            self.dataframe = pd.read_csv("trafic.txt", sep=";", header= 0)
        
            self.dataframe["Inicio de Conexion"] = pd.to_datetime(self.dataframe["Inicio de Conexion"])
            self.dataframe["Fin de Conexio"] = pd.to_datetime(self.dataframe["Fin de Conexio"], errors ='coerce')
        
            pd.set_option("expand_frame_repr", True)
            
            print("Convirtiendo archivo a xlsx...")

            self.dataframe = self.dataframe[(self.dataframe["Inicio de Conexion"]>= self.first_date) & (self.dataframe["Fin de Conexio"]<= self.second_date)]
            self.dataframe = self.dataframe.sort_values('Session Time', ascending=False) 
            self.dataframe.to_excel('trafic.xlsx')

            print(self.dataframe)

            os.system("trafic.xlsx")

        except Exception as e:

            print('Error in dataframe_handler function: ', e)    
            exit()

    def main(self):
       
        if self.get_date() == True:
            self.dataframe_handler()
           
      
p = Trafic()
p.main()

    



    
        
