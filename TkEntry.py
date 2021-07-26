from tkinter import *

top = Tk()
L1 = Label(top, text="送检医院：")
L1.pack(side = LEFT)
E1 = Entry(top, bd =5)
E1.pack(side = RIGHT)

L2 = Label(top, text="门诊/住院号：")
L2.pack( side = RIGHT)
E2 = Entry(top, bd =5)
E2.pack(side = RIGHT)

L3 = Label(top, text="送诊医生：")
L3.pack( side = RIGHT)
E3 = Entry(top, bd =5)
E3.pack(side = RIGHT)

L4 = Label(top, text="医生电话：")
L4.pack( side = RIGHT)
E4 = Entry(top, bd =5)
E4.pack(side = RIGHT)

top.mainloop()
