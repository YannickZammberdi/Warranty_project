import funs
from os import path
from os import environ
# name = "Mark Sheppard" # Input in Gui
# address = "8 Regent Street, Queanbeyan East" # Input in Gui
# data_folder = 'C:/Tom/Warranty_project' # set in Gui
# output_path = 'C:/Tom/Warranty_project'
# sub_list = ["Dream Electrical","Aus Plumb","Showerscreen"]

def main(name, address, sub_list, data_folder = 'C:/Tom/Warranty_project/', output_path = path.join(environ["HOMEPATH"], "Desktop/output.docx")):
    if output_path is not None and output_path is not "":
        funs.generate_raw_output(data_folder,sub_list,output_path)
        funs.fill_information(name, address,output_path)
    else:
        funs.generate_raw_output(data_folder, sub_list)
        funs.fill_information(name, address)

# main(name,address,sub_list,data_folder,output_path)