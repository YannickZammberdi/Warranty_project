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

    if ABEY_SOUTHERN.get() == True:
        PC_list.append("Abey Southern")
    if ARGENT_VILLEROY_AND_BOCH_SOUTHERN.get() == True:
        PC_list.append("Argent Boch Southern")
    if ARGENT_SOUTHERN.get() == True:
        PC_list.append("Argent Southern")
    if AVENIR_SOUTHERN.get() == True:
        PC_list.append("Avenir Southern")
    if BRODWARE_SOUTHERN.get() == True:
        PC_list.append("Brodware Southern")
    if CAROMA_SOUTHERN.get() == True:
        PC_list.append("Caroma Southern")
    if DECINA_SOUTHERN.get() == True:
        PC_list.append("Decina Southern")
    if FIENZA_SOUTHERN.get() == True:
        PC_list.append("Fienza Southern")
    if HANSGROHE_SOUTHERN.get() == True:
        PC_list.append("Hansgrohe Southern")
    if MARQUIS_SOUTHERN.get() == True:
        PC_list.append("Marquis Southern")
    if MEIR_SOUTHERN.get() == True:
        PC_list.append("Meir Southern")
    if METHVEN_SOUTHERN.get() == True:
        PC_list.append("Methven Southern")
    if NEKO_SOUTHERN.get() == True:
        PC_list.append("Neko Southern")
    if PIETRA_BIANCA_SOUTHERN.get() == True:
        PC_list.append("Pietra Southern")
    if RAM_SOUTHERN.get() == True:
        PC_list.append("RAM Southern")
    if STONE_BATH_SOUTHERN.get() == True:
        PC_list.append("Stone Bath Southern")
    if VICTORIA_AND_ALBERT_SOUTHERN.get() == True:
        PC_list.append("Victora Albert Southern")

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

lbl_11=Label(window, text="Southern:", fg='black', font=("Helvetica", 12))
lbl_11.place(x=750, y=80)

ABEY_SOUTHERN = BooleanVar()
cbt21=Checkbutton(window, text="ABEY", variable=ABEY_SOUTHERN)
cbt21.place(x=750, y=110)

ARGENT_VILLEROY_AND_BOCH_SOUTHERN = BooleanVar()
cbt22=Checkbutton(window, text="ARGENT VILLEROY AND BOCH", variable=ARGENT_VILLEROY_AND_BOCH_SOUTHERN)
cbt22.place(x=750, y=140)

ARGENT_SOUTHERN = BooleanVar()
cbt23=Checkbutton(window, text="ARGENT", variable=ARGENT_SOUTHERN)
cbt23.place(x=750, y=170)

AVENIR_SOUTHERN = BooleanVar()
cbt24=Checkbutton(window, text="AVENIR", variable=AVENIR_SOUTHERN)
cbt24.place(x=750, y=200)

BRODWARE_SOUTHERN = BooleanVar()
cbt25=Checkbutton(window, text="BRODWARE", variable=BRODWARE_SOUTHERN)
cbt25.place(x=750, y=230)

CAROMA_SOUTHERN = BooleanVar()
cbt26=Checkbutton(window, text="CAROMA", variable=CAROMA_SOUTHERN)
cbt26.place(x=750, y=260)

DECINA_SOUTHERN = BooleanVar()
cbt27=Checkbutton(window, text="DECINA", variable=DECINA_SOUTHERN)
cbt27.place(x=750, y=290)

FIENZA_SOUTHERN = BooleanVar()
cbt28=Checkbutton(window, text="FIENZA", variable=FIENZA_SOUTHERN)
cbt28.place(x=750, y=320)

HANSGROHE_SOUTHERN = BooleanVar()
cbt29=Checkbutton(window, text="HANSGROHE", variable=HANSGROHE_SOUTHERN)
cbt29.place(x=750, y=350)

MARQUIS_SOUTHERN = BooleanVar()
cbt30=Checkbutton(window, text="MARQUIS", variable=MARQUIS_SOUTHERN)
cbt30.place(x=750, y=380)

MEIR_SOUTHERN = BooleanVar()
cbt31=Checkbutton(window, text="MEIR", variable=MEIR_SOUTHERN)
cbt31.place(x=750, y=410)

METHVEN_SOUTHERN = BooleanVar()
cbt32=Checkbutton(window, text="METHVEN", variable=METHVEN_SOUTHERN)
cbt32.place(x=750, y=440)

NEKO_SOUTHERN = BooleanVar()
cbt33=Checkbutton(window, text="NEKO", variable=NEKO_SOUTHERN)
cbt33.place(x=750, y=470)

PIETRA_BIANCA_SOUTHERN = BooleanVar()
cbt34=Checkbutton(window, text="PIETRA BIANCA", variable=PIETRA_BIANCA_SOUTHERN)
cbt34.place(x=750, y=500)

RAM_SOUTHERN = BooleanVar()
cbt35=Checkbutton(window, text="RAM", variable=RAM_SOUTHERN)
cbt35.place(x=850, y=110)

STONE_BATH_SOUTHERN = BooleanVar()
cbt36=Checkbutton(window, text="STONE BATH", variable=STONE_BATH_SOUTHERN)
cbt36.place(x=850, y=170)

VICTORIA_AND_ALBERT_SOUTHERN = BooleanVar()
cbt37=Checkbutton(window, text="VICTORIA AND ALBERT", variable=VICTORIA_AND_ALBERT_SOUTHERN)
cbt37.place(x=850, y=200)

btn=Button(window, text="Generate Warranty Book", fg='black',
           command = lambda:main([txtfld1.get(),txtfld2.get(),txtfld3.get(),txtfld4.get(),
                                  txtfld5.get(),txtfld6.get(),txtfld7.get()]
                                 ,"C:/warranty_book","C:/warranty_book"))
btn.place(x=400, y=550)

window.title('Warranty Book Generator v0.20')
window.geometry("1000x600+10+20")
window.mainloop()
