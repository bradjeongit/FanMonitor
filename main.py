import os
import select
import socket
import struct
import threading
import time
import tkinter as tk
from queue import Queue
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import datetime
import sys

sys.setrecursionlimit(10**7)

class Load_excel():

    def __init__(self):
        #엑셀 파일 위치
        self.excel_filename = '\layout.xlsx'
        self.excel_location = os.getcwd() + self.excel_filename


    def read_layout(self):
        #엑셀 파일 읽어오기 #Dataframe으로 읽어옴
        # df: 전체 id 목록, df_status: 상태, df_show: 화면 layout, df_replace: 교체된 fan id
        self.df = pd.read_excel(self.excel_location)
        self.df_replaced = pd.read_excel(self.excel_location, sheet_name='replace_id')
        self.df_show = pd.read_excel(self.excel_location, sheet_name='show')


class Show_table():
    def __init__(self, df_speed):
        #Treeview 세팅
        self.window = tk.Tk()
        self.window.title("SUNINUS Fan")
        self.window.geometry("1200x750")  # set the window dimensions
        self.window.pack_propagate(False)  # tells the window to not let the widgets inside it determine its size.
        self.window.resizable(0, 0)  # makes the window fixed in size.

        # Frame for TreeView
        self.treeFrame = ttk.LabelFrame(self.window, text="")
        self.treeFrame.place(height=750, width=1200)

        # Style for TreeView
        self.style = ttk.Style()
        self.style.theme_use("default")
        self.style.configure("Treeview.Heading", font=('Calibri',12,'bold'))
        self.style.configure('Treeview',
                        background="white",
                        foreground="black",
                        rowheight=21,
                        fieldbackground="white")
        self.style.map('Treeview', background=[('selected', 'blue')])

        ## Treeview Widget
        self.treeView = ttk.Treeview(self.treeFrame)
        self.treeView.place(relheight=1, relwidth=1)  # set the height and width of the widget to 100% of its container (treeFrame).
        self.treeView.delete(*self.treeView.get_children())
        self.treeView["column"] = list(df_speed.columns)
        self.treeView["show"] = "headings"
        
        self.treeView.tag_configure('firstrow', background="#F7F18A",font=('Calibri',13,'bold','italic')) #lightgray
        self.treeView.tag_configure('reference', background="#FAF6AA", font=('Calibri',8)) #fan ID yellow
        self.treeView.tag_configure('reference_value', background="#FAF6AA", font=('Calibri',11)) #fan_speed yellow
        self.treeView.tag_configure('oddrow', background="white",font=('Calibri',11,'bold')) #fan speed
        self.treeView.tag_configure('evenrow', background="#EEEFFE",font=('Calibri',8)) #fan ID lightblue


    def update_treeview(self, df_speed):

        self.now = datetime.datetime.now()
        self.utime = time.mktime(self.now.timetuple())
        self.min = datetime.datetime.fromtimestamp(self.utime)

        self.window.title("SUNINUS Fan  "+str(self.min))

        self.treeView.delete(*self.treeView.get_children())
        self.treeView["column"] = list(df_speed.columns)
        self.treeView["show"] = "headings"

        for column in self.treeView["columns"]:
            self.treeView.heading(column, text=column)  # let the column heading = column name
            self.treeView.column(column, width=75, anchor="center")

        df_rows = df_speed.to_numpy().tolist()  # turns the dataframe into a list of lists

        count = 0
        for row in df_rows:
            if count == 0:
                self.treeView.insert("", "end", values=row, tags =("firstrow"))
            elif count == 1:
                self.treeView.insert("", "end", values=row, tags=("reference"))
            elif count == 2:
                self.treeView.insert("", "end", values=row, tags=("reference_value"))
            elif count % 2 == 0:
                self.treeView.insert("", "end", values=row, tags =("oddrow")) #ID
            else:
                self.treeView.insert("", "end", values=row, tags=("evenrow")) #Speed
            count += 1


class Fan():

    def __init__(self):
        # parameters
        self.port = 502 #485 통신 포트
        self.read_ref = 4 #기존 팬 param 읽어오는 커멘드
        self.read_cntr = 3 #신규 팬 param 읽어오는 커멘드
        self.write_cntr = 6 #신규 팬 param 써넣는 커멘드

        # ip list
        self.reference_fan = [[0 for j in range(16)] for i in range(2)]
        self.replaced_fan = [0 for i in range(16)]
        

    def set_sock(self, fan_ip):
        self.ip = '172.27.0.' + str(fan_ip)
        self.fan_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.fan_socket.connect((self.ip, self.port))
        self.fan_socket.settimeout(2.0)
        
        return self.fan_socket


    def read_speed(self, df, df_replaced):

        self.index_id_range = [3, 5, 7, 9, 11, 13, 15, 17, 19, 21, 23, 25, 27, 29, 31]
        self.df_show = df.copy()

        for i in range(0,16):
            for j in self.index_id_range:
                self.fan_type = df.iloc[j+1][i]
                self.fan_id = df.iloc[j][i]
                self.fan_ip = df.iloc[0][i]

                self.df_show.iloc[j][i] = str(self.fan_id)
                self.df_show.iloc[j + 1][i] = str('-')

                if self.fan_type == 'G': # Fan type G:Good, R:Replaced, B:Broken
                    self.check_speed = struct.pack('>HHHBBHH',
                                    0, 0,  # TID/ PID
                                    6, self.fan_id,  # Length / Unit ID
                                    self.read_ref, 14, 1)  # Function Code / Data (Reg addr, no of Reg)
                    self.fan_socket = self.set_sock(self.fan_ip)
                    self.fan_socket.send(self.check_speed)

                    try:
                        self.rsp_normal = self.fan_socket.recv(60)
                        self.speed_normal = struct.unpack_from('>HHHBBBH', self.rsp_normal, offset=0)
                        self.df_show.iloc[j][i] = str(self.fan_id)
                        self.df_show.iloc[j + 1][i] = str(self.speed_normal[6])
                    except OSError as msg:
                        print (msg,"norm",self.fan_socket, "  ",self.fan_id)
                elif self.fan_type == 'R':
                    self.fan_id = df_replaced.iloc[j][i]
                    self.fan_ip = df_replaced.iloc[0][i]

                    self.df_show.iloc[j][i] = str(self.fan_id)
                    
                    self.check_speed = struct.pack('>HHHBBHH',
                                    0, 0,  # TID/ PID
                                    6, self.fan_id,  # Length / Unit ID
                                    self.read_cntr, 0, 9)  # Function Code / Data (Reg addr, no of Reg)
                    
                    self.fan_socket = self.set_sock(self.fan_ip)
                    self.fan_socket.send(self.check_speed)

                    try:
                        self.rsp_replaced = self.fan_socket.recv(60)
                        self.speed_replaced = struct.unpack_from('>BHHHHHHHHHHHHH', self.rsp_replaced, offset=0)
                        self.df_show.iloc[j][i] = '['+str(self.fan_id)+']'
                        self.df_show.iloc[j + 1][i] = '['+str(self.speed_replaced[10])+']'
                    except OSError as msg:
                        print (msg,"norm",self.fan_socket, "  ",self.fan_id)
 
        return (self.df_show)


    def read_reference(self, df_show):

        for i in range(0,16):
            self.fan_ip = df_show.iloc[0][i]
            self.fan_id = df_show.iloc[1][i]
            if self.fan_id != 0:
                self.check_ref = struct.pack('>HHHBBHH',
                                            0, 0,  # TID/ PID
                                            6, self.fan_id,  # Length / Unit ID
                                            self.read_ref, 14, 1)  # Function Code / Data (Reg addr, no of Reg)
                self.fan_socket = self.set_sock(self.fan_ip)
                self.fan_socket.send(self.check_ref)

                self.fan_socket.send(self.check_ref)

                try:
                    self.rsp_ref = self.fan_socket.recv(60)

                    self.speed_ref = struct.unpack_from('>HHHBBBH', self.rsp_ref, offset=0)
#                    df.iloc[2][i] = self.speed_ref[6]
                    df_show.iloc[2][i] = str(self.speed_ref[6])
                except OSError as msg:
                    print(msg,"ref")
            else:
                df_show.iloc[2][i] = str('')
        return (df_show)


class FanMonitor():

    def __init__(self):
        self.table = Load_excel()
        self.table.read_layout()
        self.f = Fan()
        self.draw_table = Show_table(self.table.df)

    def update(self):

        sys.setrecursionlimit(10**7)
#        print(datetime.now(), '---------------------------------')

        self.df_show = self.f.read_speed(self.table.df, self.table.df_replaced) #팬 속도 읽어오기 교체 Fan 포함
        self.df_show = self.f.read_reference(self.df_show) #참고할 팬 속도 읽어오기
        self.draw_table.update_treeview(self.df_show) #화면 표시 테이블 구성

        self.draw_table.window.after(60000, self.update)
        self.draw_table.window.mainloop()


fan_monitor = FanMonitor()
fan_monitor.update()
