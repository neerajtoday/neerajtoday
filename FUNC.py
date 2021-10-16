import time
import re

def Answer(q):
	switcher = {
		"Which brand of TV do you own? ( e.g. LG Sony, etc)": 'micromax',
		"What is your mother's name?": 'prabha',
		"What is your birth place?": 'ranchi',
		"What is your shoe size? ( e.g. 5, 7 etc)": '8',
		"What is your favourite colour?": 'black',
	}
	return switcher.get(q, "Invalid q")
			
def highlight(element,filename):
	"""Highlights (blinks) a Selenium Webdriver element"""
	driver = element._parent
	def apply_style(s):
		driver.execute_script("arguments[0].setAttribute('style', arguments[1]);",element, s)
	original_style = element.get_attribute('style')
	apply_style("background: yellow; border: 2px solid red;")
	time.sleep(.1)
	driver.get_screenshot_as_file('C:\\Python\\Result\\' + filename + '.png')
	time.sleep(.4)
	apply_style(original_style)

def basicHtmlStructure(f):
	f.write("<html>")
	f.write("<head><Title>Detailed Test Execution Report</Title>")
	f.write("</head>")
	f.write("<Table border = 1 width =980 cellpadding=3 frame=border>")
	#f.write("<a align=left href=""><img src=logo.jpg width=80 height=60 align=left /></a>")
	f.write("<br />")
	f.write("<br />")
	f.write("<br />")
	f.write("<br />")
	f.write(
		"<body  bgColor=seashell><form><font color=#000080 size=4 face=calibri><left>Project Name "": VE Monthly Report""</left></font><hr>")
	f.write("<tr align=left bgcolor=#000080>")
	f.write("<th width=250><font color=#FFFFFF>Action </th>")
	f.write("<th width = 100><font color=#FFFFFF >Status</th>")
	f.write("<th width=140><font color=#FFFFFF>Execution Time</th>")
	f.write("</tr>")

def ReportN(file, Saction , Status ,timen):
	file.write("<tr align=left>")
	if Status=="Pass" or Status=="Passed" or Status=="PASS":
		file.write("<td align=left bgColor=papayawhip><font color=Green face=century size=2>"+Saction+"</td>")
		result = " <font color = Green >PASS</font>"
	elif Status=="Fail" or Status=="Failed" or Status=="FAIL":
		file.write("<td align=left bgColor=papayawhip><font color=Red face=century size=2>" + Saction + "</td>")
		result = " <font color = Red >FAIL</font>"
	else:
		file.write("<td align=left bgColor=papayawhip><font color=Blue face=century size=3>" + Saction +"</td>")
		result = " <font color = Blue >DONE</font>"
	file.write("<td align=left bgColor=papayawhip><font color=Green face=century size=2>" + result + "</td>")
	file.write("<td align=left bgColor=papayawhip><font color=black face=century size=2>" + timen[:16] + "</td>")
	file.write("</tr>")

def ReportNwithSS(file, Saction , Status ,timen,var):
	file.write("<tr align=left>")
	if Status=="Pass" or Status=="Passed" or Status=="PASS":
		file.write("<td align=left bgColor=papayawhip><font color=Green face=century size=2>"+Saction+"</td>")
		result = " <font color = Green >PASS</font>"
	elif Status=="Fail" or Status=="Failed" or Status=="FAIL":
		file.write("<td align=left bgColor=papayawhip><font color=Red face=century size=2>" + Saction + "</td>")
		result = " <font color = Red >FAIL</font>"
	else:
		img_file='C:\\Python\\Result\\' + var + '.png'
		if Saction.find('account balance of user') != -1:
			file.write("<td align=left bgColor=papayawhip><font color=Blue face=century size=3>" + Saction + "</td>")
			result = " <font color = BLUE Style='color:blue'><a href=" + img_file + ">DONE</a></font>"
		else:
			file.write("<td align=left bgColor=papayawhip><font color=Blue face=century size=3>" + Saction +"</td>")
			result = " <font color = Blue >DONE</font>"
	file.write("<td align=left bgColor=papayawhip><font color=Green face=century size=2>" + result + "</td>")
	file.write("<td align=left bgColor=papayawhip><font color=black face=century size=2>" + timen[:16] + "</td>")
	file.write("</tr>")


def replace( filePath, text, subs, flags=0 ):
	with open( filePath, "r+" ) as file:
		fileContents = file.read()
		textPattern = re.compile( re.escape( text ), flags )
		fileContents = textPattern.sub( subs, fileContents )
		file.seek( 0 )
		file.truncate()
		file.write( fileContents )
		time.sleep(.5)
		file.close()