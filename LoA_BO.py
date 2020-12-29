from docxtpl import DocxTemplate

doc = DocxTemplate("01-Letter_Of_Appointment_BO.docx")
context = {'BO' : "Mr. F. williams"}
doc.render(context)
doc.save("generated_doc.docx")
