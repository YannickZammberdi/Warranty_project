import win32com.client as win32
import docx
from docx.oxml.ns import qn
from re import finditer
from docx.shared import Pt
from docx.shared import RGBColor
from os import path
from os import environ
def generate_raw_output(data_folder, sub_list, output_path = path.join(environ["HOMEPATH"], "Desktop/output.docx")):

	word = win32.gencache.EnsureDispatch('Word.Application')
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

	for i in range(len(files)):
		output.Application.Selection.Range.InsertFile(files[len(files)-i-1])

	output.Range(output.Content.Start, output.Content.End)

	output.SaveAs(output_path+"/output.docx")
	output.Close()

def fill_information(name, address, output_path = path.join(environ["HOMEPATH"], "Desktop/output.docx")):
	document = docx.Document(output_path+"/output.docx")
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
	document.save(output_path+"/output.docx")