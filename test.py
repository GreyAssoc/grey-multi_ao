from docxtpl import DocxTemplate
import jinja2

doc = DocxTemplate("LoA_BO.docx")
context = {'BO':"Mr. F. Williams"}
doc.render(context)
doc.save("generated_doc.docx")
