import win32com.client as win32
import docx
from docx.oxml.ns import qn
from re import finditer
from docx.shared import Pt
from docx.shared import RGBColor
name = "Mark Sheppard"
address = "8 Regent Street, Queanbeyan East"

word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = False
output = word.Documents.Add()
files = ['C:\Tom\Warranty_project\data\Title_Page.docx', # Will move to the server in the future
		 'C:\Tom\Warranty_project\data\Dream_Electrical.docx',
		 'C:\Tom\Warranty_project\data\Aus_Plumb.docx',
		 'C:\Tom\Warranty_project\data\Shower_Screen.docx']
for i in range(len(files)):
	output.Application.Selection.Range.InsertFile(files[len(files)-i-1])

doc = output.Range(output.Content.Start, output.Content.End)

output.SaveAs('C:\Tom\Warranty_project\output.docx')
output.Close()

document = docx.Document('output.docx')
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


document.save('output.docx')