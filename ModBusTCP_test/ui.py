from datetime import *
from tkinter import *
import pandas as pd

df = pd.read_excel("./data.xlsx")


win = Tk()
win.geometry("1000x500")
win.title("Fan Control")
win.option_add("*Font", "Consolas 9")

btn = Button(win)  # 버튼 생성
btn.config(width=10, height=1)
btn.config(text="Button")


def alert():
    print("버튼눌림")
    dnow = datetime.now()
    btn.config(text=dnow)
    print(df)
    print(df.iloc[2][0])


btn.config(command=alert)
btn.pack()  # 버튼 배치


win.mainloop()
