import select
import socket
import struct
import threading
import time
import tkinter as tk
from queue import Queue
from tkinter import filedialog, messagebox, ttk
import pandas as pd

# parameters
port = 502
read_ref = 4
read_cntr = 3
write_cntr = 6

ip_ref = "172.27.0.82"  # reference
ip_cntr = "172.27.0.89"  # control_target
id_ref = 2
id_cntr = 2

reference_fan_list = [[0 for j in range(16)] for i in range(2)]
sock_list = [0 for i in range(16)]


class Go_Fan():
    excel_filename = "./layout.xlsx"
    df = pd.read_excel(excel_filename)
    df_type = pd.read_excel(excel_filename, sheet_name='type')
    df_show = pd.read_excel(excel_filename, sheet_name='show')
    index_id_range = [3, 5, 7, 9, 11, 13, 15, 17, 19, 21, 23, 25, 27, 29, 31]

    for i in range(0, 16):
        reference_fan_list[0][i] = '172.27.0.' + str(df.iloc[0][i])  # reference_ip address
        reference_fan_list[1][i] = df.iloc[1][i]  # reference_ID number

    root = tk.Tk()

    root.title("Fan Control")
    root.geometry("1200x750")  # set the root dimensions
    root.pack_propagate(False)  # tells the root to not let the widgets inside it determine its size.
    root.resizable(0, 0)  # makes the root window fixed in size.

    # Frame for TreeView
    treeFrame = ttk.LabelFrame(root, text="")
    treeFrame.place(height=750, width=1200)

    # Style for TreeView
    style = ttk.Style()
    style.theme_use("default")
    style.configure("Treeview.Heading", font=('Calibri', 12, 'bold'))
    style.configure('Treeview',
                    background="white",
                    foreground="black",
                    rowheight=21,
                    fieldbackground="white")
    style.map('Treeview', background=[('selected', 'blue')])

    ## Treeview Widget
    treeView = ttk.Treeview(treeFrame)
    treeView.place(relheight=1,
                   relwidth=1)  # set the height and width of the widget to 100% of its container (treeFrame).
    treeView.delete(*treeView.get_children())
    treeView["column"] = list(df_show.columns)
    treeView["show"] = "headings"

    for i in range(0, 16):
        sock_list[i] = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock_list[i].connect((reference_fan_list[0][i], port))
        #        sock_list[i].setblocking(False)
        sock_list[i].settimeout(1.0)

    def read_normal(self):

        for i in range(0, 16):
            for j in self.index_id_range:
                fan_type = self.df_type.iloc[j][i]
                fan_id = self.df.iloc[j][i]
                if fan_type == 'O':
                    check_normal_speed = struct.pack('>HHHBBHH',
                                                     0, 0,  # TID/ PID
                                                     6, fan_id,  # Length / Unit ID
                                                     read_ref, 14, 1)  # Function Code / Data (Reg addr, no of Reg)
                    sock_list[i].send(check_normal_speed)

                    try:
                        rsp_normal = sock_list[i].recv(60)
                        speed_normal = struct.unpack_from('>HHHBBBH', rsp_normal, offset=0)
                        self.df_show.iloc[j][i] = str(fan_id)
                        self.df_show.iloc[j + 1][i] = str(speed_normal[6])
                    except OSError as msg:
                        print(msg)

    def read_reference(self):

        for i in range(0, 16):
            check_ref = struct.pack('>HHHBBHH',
                                    0, 0,  # TID/ PID
                                    6, reference_fan_list[1][i],  # Length / Unit ID
                                    read_ref, 14, 1)  # Function Code / Data (Reg addr, no of Reg)

            sock_list[i].send(check_ref)

            rsp_ref = sock_list[i].recv(60)

            speed_ref = struct.unpack_from('>HHHBBBH', rsp_ref, offset=0)
            self.df_show.iloc[2][i] = str(speed_ref[6])

    def patrol(self):

        self.ip_ref = "172.27.0.82"
        self.ip_cntr = "172.27.0.89"
        self.id_ref = 2
        self.id_cntr = 2

        self.read_normal()
        self.read_reference()
        self.operate_fan()

        self.root.after(1000, self.patrol)

    def operate_fan(self):

        sock_ref = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock_cntr = socket.socket(socket.AF_INET, socket.SOCK_STREAM)

        sock_ref.connect((self.ip_ref, port))
        sock_cntr.connect((self.ip_cntr, port))

        check_ref = struct.pack('>HHHBBHH',
                                0, 0,  # TID/ PID
                                6, self.id_ref,  # Length / Unit ID
                                read_ref, 14, 1)  # Function Code / Data (Reg addr, no of Reg)

        onoff_cntr = struct.pack('>HHHBBHH',
                                 0, 0,  # TID/ PID
                                 6, 2,  # Length / Unit ID
                                 write_cntr, 20, 1)  # Function Code / Data (Reg addr, value)

        sock_ref.send(check_ref)
        rsp_ref = sock_ref.recv(60)

        speed_ref = struct.unpack_from('>HHHBBBH', rsp_ref, offset=0)

        self.df_show.iloc[4][1] = str(speed_ref[6])

        check_cntr = struct.pack('>HHHBBHH',
                                 0, 0,  # TID/ PID
                                 6, self.id_cntr,  # Length / Unit ID
                                 read_cntr, 19, 7)  # Function Code / Data (Reg addr, no of Reg)

        sock_cntr.send(check_cntr)
        rsp_cntr = sock_cntr.recv(60)

        status_cntr = struct.unpack_from('>HHHHBHHHH', rsp_cntr, offset=0)

        if (status_cntr[6] == 0):
            print('Fan is off now [', self.ip_cntr, ' ID:', self.id_cntr, ']')
            sock_cntr.send(onoff_cntr)

            rsp_cntr = sock_cntr.recv(60)

        set_speed_cntr = struct.pack('>HHHBBHH',
                                     0, 0,  # TID/ PID
                                     6, self.id_cntr,  # Length / Unit ID
                                     write_cntr, 22, speed_ref[6])  # Function Code / Data (Reg addr, value)

        self.update_treeview()

    def update_treeview(self):

        self.treeView.tag_configure('firstrow', background="lightgray", font=('Calibri', 13, 'bold', 'italic'))
        self.treeView.tag_configure('reference', background="yellow", font=('Calibri', 11, 'bold'))  # fan ID
        self.treeView.tag_configure('reference_value', background="yellow", font=('Calibri', 9))  # fan_speed
        self.treeView.tag_configure('oddrow', background="white", font=('Calibri', 9))  # fan speed
        self.treeView.tag_configure('evenrow', background="lightblue", font=('Calibri', 11, 'bold'))  # fan ID

        self.treeView.delete(*Go_Fan.treeView.get_children())
        self.treeView["column"] = list(Go_Fan.df_show.columns)
        self.treeView["show"] = "headings"

        for column in self.treeView["columns"]:
            self.treeView.heading(column, text=column)  # let the column heading = column name
            self.treeView.column(column, width=75, anchor="center")

        df_rows = Go_Fan.df_show.to_numpy().tolist()  # turns the dataframe into a list of lists

        count = 0
        for row in df_rows:
            if count == 0:
                self.treeView.insert("", "end", values=row, tags=("firstrow"))
            elif count == 1:
                self.treeView.insert("", "end", values=row, tags=("reference"))
            elif count == 2:
                self.treeView.insert("", "end", values=row, tags=("reference_value"))
            elif count % 2 == 0:
                self.treeView.insert("", "end", values=row, tags=("oddrow"))
            else:
                self.treeView.insert("", "end", values=row, tags=("evenrow"))
            count += 1

    def __init__(self):

        self.patrol()
        self.root.mainloop()


app = Go_Fan()