import funs
from tkinter import *
from os import path
from os import environ

full_sub_list = ["Dream_Electrical", "Aus_Plumb", "Can_Plumb", "Shower_Screen", "Beaumont_Tile", "Black_Tile"]
full_Reece_list = ["ADP Reece","AXA Reece","CAROMA Reece","KADO Reece","LAUFEN Reece","METHVEN Reece","MILLI Reece",
                "MIZU Reece","NIKLES Reece","PHOENIX Reece","POSH Reece","ROCA Reece","SONIA Reece","SUSSEX Reece"]
full_Southern_list = ["Abey Southern","Argent Villeroy and Boch Southern","Argent Southern","Avenir Southern",
                      "Brodware Southern","Caroma Southern","Decina Southern","Fienza Southern","Hansgrohe Southern",
                      "Marquis Southern","Meir Southern","Methven Southern","Neko Southern","PIETRA BIANCA Southern",
                      "RAM Southern","Stone Bath Southern","Victoria and Albert Southern"]
full_Reece_list = sorted(full_Reece_list)
full_Southern_list = sorted(full_Southern_list)
page_list = full_Reece_list + full_Southern_list
def main(input_list, data_folder = 'C:/Tom/Warranty_project/', output_path = path.join(environ["HOMEPATH"], "Desktop")):

    sub_list = []
    for i in range(len(full_sub_list)):
        if sub_var[i].get()==True:
            sub_list.append(full_sub_list[i])

    PC_list = []
    for i in range(len(page_list)):
        if page_var[i].get()==True:
            PC_list.append((page_list[i]))

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

sub_var = []
sub_cbt = []
for i in range(len(full_sub_list)):
    sub_var.append(BooleanVar())
    sub_cbt.append(Checkbutton(window, text=full_sub_list[i], variable=sub_var[i]))
    if 200 + 30 * i < 351:
        sub_cbt[i].place(x=100, y=200 + 30 * i)
    else:
        sub_cbt[i].place(x=300, y=200 + 30 * (i-6))

page_var = []
page_cbt = []
for i in range(len(page_list)):
    page_var.append(BooleanVar())
    if "Reece" in page_list[i]:
        page_cbt.append(Checkbutton(window, text=page_list[i].replace(" Reece", ""), variable=page_var[i]))
        page_cbt[i].place(x=550, y=110 + 30*i)
    if "Southern" in page_list[i]:
        page_cbt.append(Checkbutton(window, text=page_list[i].replace(" Southern", ""), variable=page_var[i]))
        if i-len(full_Reece_list) <14:
            page_cbt[i].place(x=750, y=110 + 30*(i-len(full_Reece_list)))
        else:
            page_cbt[i].place(x=950, y=110 + 30 * (i - 14- len(full_Reece_list)))


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

lbl_11=Label(window, text="Southern:", fg='black', font=("Helvetica", 12))
lbl_11.place(x=750, y=80)

btn=Button(window, text="Generate Warranty Book", fg='black',
           command = lambda:main([txtfld1.get(),txtfld2.get(),txtfld3.get(),txtfld4.get(),
                                  txtfld5.get(),txtfld6.get(),txtfld7.get()]
                                 ,"C:/warranty_book","C:/warranty_book"))
btn.place(x=400, y=550)

window.title('Warranty Book Generator v0.30')
window.geometry("1200x700+10+20")
window.resizable(0,0)
window.mainloop()
