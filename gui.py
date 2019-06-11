import main
from tkinter import *

sub_list = []

window=Tk()

lbl=Label(window, text="Customer Name:", fg='black', font=("Helvetica", 12))
lbl.place(x=50, y=50)

txtfld1=Entry(window, text="", bd=2)
txtfld1.place(x=200, y=50,width = 300)

lbl_2=Label(window, text="Customer Address:", fg='black', font=("Helvetica", 12))
lbl_2.place(x=50, y=100)

txtfld2=Entry(window, text="", bd=2)
txtfld2.place(x=200, y=100,width = 300)

lbl_3=Label(window, text="Warranty Book Pages:", fg='black', font=("Helvetica", 12))
lbl_3.place(x=50, y=150)

Dream_Electrical = BooleanVar()
cbt1=Checkbutton(window, text="Dream_Electrical", variable=Dream_Electrical)
cbt1.place(x=100, y=200)

Aus_Plumb = BooleanVar()
cbt2=Checkbutton(window, text="Aus Plumbing", variable=Aus_Plumb)
cbt2.place(x=100, y=250)

Can_Plumb = BooleanVar()
cbt3=Checkbutton(window, text="Can Plumbing(Phil Foster)", variable=Can_Plumb)
cbt3.place(x=200, y=250)

Showerscreen = BooleanVar()
cbt4=Checkbutton(window, text="Showerscreen", variable=Showerscreen)
cbt4.place(x=100, y=300)

if Dream_Electrical == 1:
    sub_list.append("Dream Electrical")
if Aus_Plumb == 1:
    sub_list.append("Aus Plumb")
if Can_Plumb == 1:
    sub_list.append("Can Plumb")
if Showerscreen == 1:
    sub_list.append("Showerscreen")

btn=Button(window, text="Generate Warranty Book", fg='black',
           command = lambda:main.main(txtfld1.get(),txtfld2.get(),sub_list
                                      ,'C:/Tom/Warranty_project','C:/Tom/Warranty_project'))
btn.place(x=200, y=350)

# btn.bind("<Button-1>",func.select('data/selection.xlsx'))

window.title('Warranty Book Generator v0.1')
window.geometry("600x500+10+20")
window.mainloop()