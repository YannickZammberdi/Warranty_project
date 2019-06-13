import docx
from docx.oxml.ns import qn
from re import finditer
from docx.shared import Pt
from docx.shared import RGBColor
from os import path
from os import environ
import win32com.client
from PyPDF2 import PdfFileMerger

def generate_raw_output(name, data_folder, sub_list, output_path = path.join(environ["HOMEPATH"], "Desktop/output.docx")):

	word = win32com.client.gencache.EnsureDispatch('Word.Application')
	word.Visible = False
	output = word.Documents.Add()
	files = [data_folder+'/data/Title_Page.docx'] # Will move to the server in the future
	if "Dream Electrical" in sub_list:
		files.append(data_folder+'/data/Dream_Electrical.docx')
	if "Aus Plumb" in sub_list:
		files.append(data_folder+'/data/Aus_Plumb.docx')
	if "Can Plumb" in sub_list:
		files.append(data_folder+'/data/Can_Plumb.docx')
	if "Showerscreen" in sub_list:
		files.append(data_folder+'/data/Shower_Screen.docx')
	if "Beaumont" in sub_list:
		files.append(data_folder+'/data/Beaumont_Tile.docx')
	if "Black" in sub_list:
		files.append(data_folder+'/data/Black_Tile.docx')

	for i in range(len(files)):
		output.Application.Selection.Range.InsertFile(files[len(files)-i-1])

	output.Range(output.Content.Start, output.Content.End)

	output.SaveAs(output_path+"/"+name+" Warranty Book.docx")
	output.Close()

def fill_information(input_list, output_path = path.join(environ["HOMEPATH"], "Desktop/output.docx")):
	name = input_list[0]
	address = input_list[1]
	Tile1 = input_list[2]
	Tile2 = input_list[3]
	Tile3 = input_list[4]
	Tile4 = input_list[5]
	Tile5 = input_list[6]
	document = docx.Document(output_path+"/"+name+" Warranty Book.docx")
	for paragraph in document.paragraphs:
		if 'Customer_Address' in paragraph.text:
			paragraph.text = paragraph.text.replace("[Customer_Address]",address)
			paragraph_reserve = paragraph.text
			paragraph.text = ""
			try:
				place = [m.start() for m in finditer(address, paragraph_reserve)][0]
			except:
				place = 0
			for i, ch in enumerate(paragraph_reserve):
				if place != 0:
					run = paragraph.add_run(ch)
					font = run.font
					font.name = u'Calibri (Body)'
					run._element.rPr.rFonts.set(qn('w:eastAsia'), u'Calibri (Body)')
					bold = 0
					if i >= place - 1 and i <= place - 1 + len(address):
						bold = 1
					font.bold = bold
				else:
					run = paragraph.add_run(ch)
					font = run.font
					font.name = u'Arial'
					font.size = Pt(24)
					color = font.color
					color.rgb = RGBColor(255,255,255)
					run._element.rPr.rFonts.set(qn('w:eastAsia'), u'Arial')
		if 'Customer_Name' in paragraph.text:
			paragraph.text = paragraph.text.replace("[Customer_Name]",name)
			paragraph_reserve = paragraph.text
			paragraph.text = ""
			try:
				place = [m.start() for m in finditer(name, paragraph_reserve)][0]
			except:
				place = 0
			for i, ch in enumerate(paragraph_reserve):
				run = paragraph.add_run(ch)
				font = run.font
				font.name = u'Arial'
				font.size = Pt(24)
				color = font.color
				color.rgb = RGBColor(255, 255, 255)
				run._element.rPr.rFonts.set(qn('w:eastAsia'), u'Arial')
		if 'Tile_1' in paragraph.text:
			paragraph.text = paragraph.text.replace("[Tile_1]",Tile1)
			paragraph_reserve = paragraph.text
			paragraph.text = ""
			try:
				place = [m.start() for m in finditer(Tile1, paragraph_reserve)][0]
			except:
				place = 0
			for i, ch in enumerate(paragraph_reserve):
				run = paragraph.add_run(ch)
				font = run.font
				font.name = u'Arial'
				font.size = Pt(11)
				color = font.color
				color.rgb = RGBColor(0,0,0)
				run._element.rPr.rFonts.set(qn('w:eastAsia'), u'Arial')
		if 'Tile_2' in paragraph.text:
			paragraph.text = paragraph.text.replace("[Tile_2]",Tile2)
			paragraph_reserve = paragraph.text
			paragraph.text = ""
			try:
				place = [m.start() for m in finditer(Tile2, paragraph_reserve)][0]
			except:
				place = 0
			for i, ch in enumerate(paragraph_reserve):
				run = paragraph.add_run(ch)
				font = run.font
				font.name = u'Arial'
				font.size = Pt(11)
				color = font.color
				color.rgb = RGBColor(0,0,0)
				run._element.rPr.rFonts.set(qn('w:eastAsia'), u'Arial')
		if 'Tile_3' in paragraph.text:
			paragraph.text = paragraph.text.replace("[Tile_3]",Tile3)
			paragraph_reserve = paragraph.text
			paragraph.text = ""
			try:
				place = [m.start() for m in finditer(Tile3, paragraph_reserve)][0]
			except:
				place = 0
			for i, ch in enumerate(paragraph_reserve):
				run = paragraph.add_run(ch)
				font = run.font
				font.name = u'Arial'
				font.size = Pt(11)
				color = font.color
				color.rgb = RGBColor(0,0,0)
				run._element.rPr.rFonts.set(qn('w:eastAsia'), u'Arial')
		if 'Tile_4' in paragraph.text:
			paragraph.text = paragraph.text.replace("[Tile_4]",Tile4)
			paragraph_reserve = paragraph.text
			paragraph.text = ""
			try:
				place = [m.start() for m in finditer(Tile4, paragraph_reserve)][0]
			except:
				place = 0
			for i, ch in enumerate(paragraph_reserve):
				run = paragraph.add_run(ch)
				font = run.font
				font.name = u'Arial'
				font.size = Pt(11)
				color = font.color
				color.rgb = RGBColor(0,0,0)
				run._element.rPr.rFonts.set(qn('w:eastAsia'), u'Arial')
		if 'Tile_5' in paragraph.text:
			paragraph.text = paragraph.text.replace("[Tile_5]",Tile5)
			paragraph_reserve = paragraph.text
			paragraph.text = ""
			try:
				place = [m.start() for m in finditer(Tile5, paragraph_reserve)][0]
			except:
				place = 0
			for i, ch in enumerate(paragraph_reserve):
				run = paragraph.add_run(ch)
				font = run.font
				font.name = u'Arial'
				font.size = Pt(11)
				color = font.color
				color.rgb = RGBColor(0,0,0)
				run._element.rPr.rFonts.set(qn('w:eastAsia'), u'Arial')
	document.save(output_path+"/"+name+" Warranty Book.docx")

def pdf_combiner(data_folder, PC_list, name, output_path = path.join(environ["HOMEPATH"], "Desktop/output.docx")): # Combine the files in the pdf
	files = []
	if "ADP Reece" in PC_list:
		files.append(data_folder+'/data/Reece/ADP.pdf')
	if "AXA Reece" in PC_list:
		files.append(data_folder+'/data/Reece/AXA.pdf')
	if "CAROMA Reece" in PC_list:
		files.append(data_folder+'/data/Reece/CAROMA.pdf')
	if "KADO Reece" in PC_list:
		files.append(data_folder+'/data/Reece/KADO.pdf')
	if "LAUFEN Reece" in PC_list:
		files.append(data_folder+'/data/Reece/LAUFEN.pdf')
	if "METHVEN Reece" in PC_list:
		files.append(data_folder+'/data/Reece/METHVEN.pdf')
	if "MILLI Reece" in PC_list:
		files.append(data_folder+'/data/Reece/MILLI.pdf')
	if "MIZU Reece" in PC_list:
		files.append(data_folder+'/data/Reece/MIZU.pdf')
	if "NIKLES Reece" in PC_list:
		files.append(data_folder+'/data/Reece/NIKLES.pdf')
	if "PHOENIX Reece" in PC_list:
		files.append(data_folder+'/data/Reece/PHOENIX.pdf')
	if "POSH Reece" in PC_list:
		files.append(data_folder+'/data/Reece/POSH.pdf')
	if "ROCA Reece" in PC_list:
		files.append(data_folder+'/data/Reece/ROCA.pdf')
	if "SONIA Reece" in PC_list:
		files.append(data_folder+'/data/Reece/SONIA.pdf')
	if "SUSSEX Reece" in PC_list:
		files.append(data_folder+'/data/Reece/SUSSEX.pdf')

	if "Abey Southern" in PC_list:
		files.append(data_folder+'/data/Southern/ABEY.pdf')
	if "Argent Boch Southern" in PC_list:
		files.append(data_folder+'/data/Southern/ARGENT VILLEROY AND BOCH.pdf')
	if "Argent Southern" in PC_list:
		files.append(data_folder+'/data/Southern/ARGENT.pdf')
	if "Avenir Southern" in PC_list:
		files.append(data_folder+'/data/Southern/AVENIR.pdf')
	if "Brodware Southern" in PC_list:
		files.append(data_folder+'/data/Southern/BRODWARE.pdf')
	if "Caroma Southern" in PC_list:
		files.append(data_folder+'/data/Southern/CAROMA.pdf')
	if "Decina Southern" in PC_list:
		files.append(data_folder+'/data/Southern/DECINA.pdf')
	if "Fienza Southern" in PC_list:
		files.append(data_folder+'/data/Southern/FIENZA.pdf')
	if "Hansgrohe Southern" in PC_list:
		files.append(data_folder+'/data/Southern/HANSGROHE.pdf')
	if "Meir Southern" in PC_list:
		files.append(data_folder+'/data/Southern/MARQUIS.pdf')
	if "Meir Southern" in PC_list:
		files.append(data_folder+'/data/Southern/MEIR.pdf')
	if "Methven Southern" in PC_list:
		files.append(data_folder+'/data/Southern/METHVEN.pdf')
	if "Neko Southern" in PC_list:
		files.append(data_folder+'/data/Southern/NEKO.pdf')
	if "Pietra Southern" in PC_list:
		files.append(data_folder+'/data/Southern/PIETRA BIANCA.pdf')
	if "RAM Southern" in PC_list:
		files.append(data_folder+'/data/Southern/RAM.pdf')
	if "Stone Bath Southern" in PC_list:
		files.append(data_folder+'/data/Southern/STONE BATH.pdf')
	if "Victora Albert Southern" in PC_list:
		files.append(data_folder+'/data/Southern/VICTORIA AND ALBERT.pdf')

	merger = PdfFileMerger()
	for i in range(len(files)):
		merger.append(open(files[i], 'rb'))
	with open(output_path + "/" + name + " Warranty Book PC Items.pdf", 'wb') as fout:  # different from func
		# This is the output path
		merger.write(fout)