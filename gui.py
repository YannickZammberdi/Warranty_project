import funs
from tkinter import *
from os import path
from os import environ
from sys import executable

def main(input_list, data_folder = 'C:/Tom/Warranty_project/', output_path = path.join(environ["HOMEPATH"], "Desktop")):
    # Todo: Tile maintenance pages need to be added, just jump to the product pages.
    sub_list = []
    if Dream_Electrical.get() == True:
        sub_list.append("Dream Electrical")
    if Aus_Plumb.get() == True:
        sub_list.append("Aus Plumb")
    if Can_Plumb.get() == True:
        sub_list.append("Can Plumb")
    if Showerscreen.get() == True:
        sub_list.append("Showerscreen")
    if Beaumont.get() == True:
        sub_list.append("Beaumont")
    if Black.get() == True:
        sub_list.append("Black")

    PC_list = []
    if ADP.get() == True:
        PC_list.append("ADP Reece")
    if AXA.get() == True:
        PC_list.append("AXA Reece")
    if CAROMA.get() == True:
        PC_list.append("CAROMA Reece")
    if KADO.get() == True:
        PC_list.append("KADO Reece")
    if LAUFEN.get() == True:
        PC_list.append("LAUFEN Reece")
    if METHVEN.get() == True:
        PC_list.append("METHVEN Reece")
    if MILLI.get() == True:
        PC_list.append("MILLI Reece")
    if MIZU.get() == True:
        PC_list.append("MIZU Reece")
    if NIKLES.get() == True:
        PC_list.append("NIKLES Reece")
    if PHOENIX.get() == True:
        PC_list.append("PHOENIX Reece")
    if POSH.get() == True:
        PC_list.append("POSH Reece")
    if ROCA.get() == True:
        PC_list.append("ROCA Reece")
    if SONIA.get() == True:
        PC_list.append("SONIA Reece")
    if SUSSEX.get() == True:
        PC_list.append("SUSSEX Reece")

    name = input_list[0]
    if output_path is not None and output_path is not "":
        funs.generate_raw_output(name, data_folder, sub_list,output_path)
        funs.fill_information(input_list,output_path)
        funs.pdf_combiner(data_folder, PC_list, name, output_path)
    else:
        funs.generate_raw_output(name, data_folder, sub_list)
        funs.fill_information(input_list,output_path)
        funs.pdf_combiner(data_folder, PC_list, name, output_path)

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

Beaumont = BooleanVar()
cbt5=Checkbutton(window, text="Beaumont", variable=Beaumont)
cbt5.place(x=100, y=350)

Black = BooleanVar()
cbt6=Checkbutton(window, text="Blackrock", variable=Black)
cbt6.place(x=200, y=350)

txtfld3=Entry(window, text="", bd=2)
txtfld3.place(x=200, y=400,width = 300)

lbl_4=Label(window, text="Tile 1:", fg='black', font=("Helvetica", 12))
lbl_4.place(x=50, y=400)

txtfld4=Entry(window, text="", bd=2)
txtfld4.place(x=200, y=430,width = 300)

lbl_5=Label(window, text="Tile 2:", fg='black', font=("Helvetica", 12))
lbl_5.place(x=50, y=430)

txtfld5=Entry(window, text="", bd=2)
txtfld5.place(x=200, y=460,width = 300)

lbl_6=Label(window, text="Tile 3:", fg='black', font=("Helvetica", 12))
lbl_6.place(x=50, y=460)

txtfld6=Entry(window, text="", bd=2)
txtfld6.place(x=200, y=490,width = 300)

lbl_7=Label(window, text="Tile 4:", fg='black', font=("Helvetica", 12))
lbl_7.place(x=50, y=490)

txtfld7=Entry(window, text="", bd=2)
txtfld7.place(x=200, y=520,width = 300)

lbl_8=Label(window, text="Tile 5:", fg='black', font=("Helvetica", 12))
lbl_8.place(x=50, y=520)

lbl_9=Label(window, text="PC Items Warranty Pages:", fg='black', font=("Helvetica", 12))
lbl_9.place(x=550, y=50)

lbl_10=Label(window, text="Reece:", fg='black', font=("Helvetica", 12))
lbl_10.place(x=550, y=80)

ADP = BooleanVar()
cbt7=Checkbutton(window, text="ADP", variable=ADP)
cbt7.place(x=550, y=110)

AXA = BooleanVar()
cbt8=Checkbutton(window, text="AXA", variable=AXA)
cbt8.place(x=550, y=140)

CAROMA = BooleanVar()
cbt9=Checkbutton(window, text="CAROMA", variable=CAROMA)
cbt9.place(x=550, y=170)

KADO = BooleanVar()
cbt10=Checkbutton(window, text="KADO", variable=KADO)
cbt10.place(x=550, y=200)

LAUFEN = BooleanVar()
cbt11=Checkbutton(window, text="LAUFEN", variable=LAUFEN)
cbt11.place(x=550, y=230)

METHVEN = BooleanVar()
cbt12=Checkbutton(window, text="METHVEN", variable=METHVEN)
cbt12.place(x=550, y=260)

MILLI = BooleanVar()
cbt13=Checkbutton(window, text="MILLI", variable=MILLI)
cbt13.place(x=550, y=290)

MIZU = BooleanVar()
cbt14=Checkbutton(window, text="MIZU", variable=MIZU)
cbt14.place(x=550, y=320)

NIKLES = BooleanVar()
cbt15=Checkbutton(window, text="NIKLES", variable=NIKLES)
cbt15.place(x=550, y=350)

PHOENIX = BooleanVar()
cbt16=Checkbutton(window, text="PHOENIX", variable=PHOENIX)
cbt16.place(x=550, y=380)

POSH = BooleanVar()
cbt17=Checkbutton(window, text="POSH", variable=POSH)
cbt17.place(x=550, y=410)

ROCA = BooleanVar()
cbt18=Checkbutton(window, text="ROCA", variable=ROCA)
cbt18.place(x=550, y=440)

SONIA = BooleanVar()
cbt19=Checkbutton(window, text="SONIA", variable=SONIA)
cbt19.place(x=550, y=470)

SUSSEX = BooleanVar()
cbt20=Checkbutton(window, text="SUSSEX", variable=SUSSEX)
cbt20.place(x=550, y=500)

btn=Button(window, text="Generate Warranty Book", fg='black',
           command = lambda:main([txtfld1.get(),txtfld2.get(),txtfld3.get(),txtfld4.get(),
                                  txtfld5.get(),txtfld6.get(),txtfld7.get()]
                                 ,"C:/warranty_book","C:/warranty_book"))
btn.place(x=400, y=550)

window.title('Warranty Book Generator v0.11')
window.geometry("1000x600+10+20")
window.mainloop()
