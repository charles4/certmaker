import zipfile

class CertMaker():

	def __init__(self, student, template):
		self.PATH = "C:/tmp/"

		self.template = zipfile.ZipFile(self.PATH+str(template))
		self.newdoc = zipfile.ZipFile(self.PATH+str(student)+".docx", "w")
		self.student = student


	def fill_template(self):

		tmpXMLfile = open(self.template.extract("word/document.xml", "C:/"))
		tmpXMLstr = tmpXMLfile.read()
		tmpXMLfile.close()

		tmpXMLstr = tmpXMLstr.replace("{{student}}", self.student)

		tmpXMLfile = open(self.PATH+"temp.xml", "w+")
		tmpXMLfile.write(tmpXMLstr)
		tmpXMLfile.close()

		for file in self.template.filelist:
			if not file.filename == "word/document.xml":
				self.newdoc.writestr(file.filename, self.template.read(file))


		self.newdoc.write(self.PATH+"temp.xml", "word/document.xml")

		self.template.close()
		self.newdoc.close()

		print "Created: " + self.student + ".docx"


if __name__ == "__main__":
	cm = CertMaker("Billy Jo Jenkins", "cert_template.docx")
	cm.fill_template()