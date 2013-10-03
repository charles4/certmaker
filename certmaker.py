import zipfile

replaceText = {"{{ student }}" : "Joe Bob"}
templateDocx = zipfile.ZipFile("C:/cert_template.docx")
newDocx = zipfile.ZipFile("C:/NewDocument.docx", "a")

tempXmlFile = open(templateDocx.extract("word/document.xml", "C:/"))
tempXmlStr = tempXmlFile.read()
tempXmlFile.close()

for key in replaceText.keys():
	tempXmlStr = tempXmlStr.replace(str(key), str(replaceText.get(key))

tempXmlFile = open("C:/temp.xml", "w+")
tempXmlFile.write(tempXmlStr)
tempXmlFile.close()

for file in templateDocx.filelist:
	if not file.filename == "word/document.xml":
		newDocx.writestr(file.filename, templateDocx.read(file))

newDocx.write("C:/temp.xml", "word/document.xml")

templateDocx.close()
newDocx.close()