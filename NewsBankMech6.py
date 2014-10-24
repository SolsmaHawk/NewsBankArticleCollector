import os, sys
import bs4
import requests
import urllib
import urllib2
from bs4 import BeautifulSoup
import mechanize
import time
import datetime
import xlsxwriter
import datetime

#___________Global Variables_______________

baseUrl="http://infoweb.newsbank.com"
mainSearchPage = 'http://infoweb.newsbank.com/iw-search/we/InfoWeb?p_product=AWNB&p_theme=aggregated5&p_action=explore&d_loc=United%20States&d_place=&f_view=loc&d_selLoc=&d_selSrc='
queries = ['queryname=1','queryname=2','queryname=3','queryname=4','queryname=5','queryname=6','queryname=7','queryname=8','queryname=9']
oldNextButton=None # store link button to compare to - checks to see if the end of articles has been reached
nextButton=None #store next button link once found
linkList=[] #store links and article names
linkStrings=[]
nextSwitch=0
#workbook=None #xlsx workbook
#worksheet=None
row=0
col=0
state=None
topic=None
fileName=None

#___________Browser Setup_______________

br = mechanize.Browser() # create a new browse
br.addheaders = [('User-agent', 'Firefox')] # set headers as Firefox
br.set_handle_robots( False ) # ignore robots.txt


def mechBrowseNewsBank():
	global fileName
	global state
	global topic
	print("Initializing browser. Please wait...")
	while True:
		try:
			br.open(mainSearchPage) # open NewsBank main search page
			break
		except:
			print("Cannot connect to NewsBank. Make sure your are on a network with NewsBank or try again later.")
			print("Will automatically try to reconnect in 10 seconds...")
			time.sleep(10)
	print("Browser initialized: "+str(br.title()))
	br.select_form(nr=1) # select second set of forms
	tempList=[str(i) for i in br.find_control(type="checkbox").items]
	while True:
		state = raw_input("What state would you like to search? ex: 'IA' ")
		if state in tempList:
			print("State found - selected")
			break
		else:
			print("State not found. Make sure format is state abbreviation all in caps.")
	topic = raw_input("What topic would you like to search? - Search box 1 - ")
	for i in range(0, len(br.find_control(type="checkbox").items)): #iterate over possible checkboxes
		if str(br.find_control(type="checkbox").items[i]) == state: #check for selected state - state variable
			#print("State found - selected")
			br.find_control(type="checkbox").items[i].selected =True #check selection box
    	list = [] # store search page controls here
    	for f in br.form.controls: 
		#Add the names of each item in br.formcontrols
		list.append(f.name)
	
	#######
	controlNumbers=0 
	control = br.form.find_control("p_field_base-0") # Select first search field
	options=[]
	optionsPrint=[]
	for item in control.items: # itterate over drop down box choices and append them to both option lists, one for indexing, one for reading by the user
		options.append((str(controlNumbers),item.name))
		optionsPrint.append("("+str(controlNumbers)+","+item.name+")")
		controlNumbers+=1
	optionsPrint[0]="(0,All Text)"
	optionsPrint[1]="(1,Lead Paragraph)"
	optionsPrint[8]="(0,Word Count)"
	optionsPrint[9]="(0,Date)"
	searchChoice=raw_input("Choose a number from this list of search options: "+" ".join(optionsPrint)+"  ")
	for item in control.items:
		if item.name == options[int(searchChoice)][1]:
			item.selected = True
			print(item)
	
	boolQuestion = raw_input("Would you like to search with booleans? (yes/no) ")
	if boolQuestion in ["yes","Yes","YES"]:
		controlNumbers=0
		options[:]=[]
		optionsPrint[:]=[]
		control = br.form.find_control("p_bool_base-1")
		for item in control.items:
			options.append((str(controlNumbers),item.name))
			optionsPrint.append("("+str(controlNumbers)+","+item.name+")")
			controlNumbers+=1
		
		searchChoice=raw_input("Choose from these boolean operators: "+" ".join(optionsPrint)+"  ")
		for item in control.items:
			if item.name == options[int(searchChoice)][1]:
				item.selected = True
				print(item)
				
		topic2 = raw_input("What is the second topic you would like to search? - Search box 2 - ")
		br.form["p_text_base-1"]=topic2
		#options.append((item,controlNumbers))
		#controlNumbers+=1
	#print(options)
	#######
	
	#Select the correct one from the list.
	#print list[1] # index one is the main search text box
	br.form[list[1]]=topic #populate text box with search keyword - topic variable
	response = br.submit()
	#print(br.geturl())
	soup = BeautifulSoup(response)
	return(soup)

	
def browserLoop(soup):
	global nextButton
	global nextSwitch
	global oldNextButton
	global row
	global col
	global state
	#global workbook
	#global worksheet
	### Xlsx creation ###
	workbook = xlsxwriter.Workbook(state+"_"+topic+"_"+'.xlsx') # create xlsx file with state-topic-date as name
	fileName=state+"_"+topic+"_"'.xlsx'
	worksheet = workbook.add_worksheet()
	worksheet.write(row, col, "State Name") # write all headers to excel file
	worksheet.write(row, col+1, "State")
	worksheet.write(row, col+2, "Year")
	worksheet.write(row, col+3, "Newspaper")
	worksheet.write(row, col+4, "Title")
	worksheet.write(row, col+5, "dayofweek")
	worksheet.write(row, col+6, "month")
	worksheet.write(row, col+7, "date")
	worksheet.write(row, col+8, "section")
	worksheet.write(row, col+9, "record")
	worksheet.write(row, col+10, "text_article")
	try:
		print("File "+fileName+" created at: "+str(os.path.dirname(os.path.realpath(__file__))))
	except:
		print("Xlsx file created")
	row+=1
	#################
	
	try:
		for searchCriteria in soup.findAll('div', attrs={'id': 'searchString'}): # prints search Criteria
			print(searchCriteria.get_text())
	except:
		pass
	try:
		for results in soup.findAll('div', attrs={'class': 'jump_results'}): # prints the number of results recovered
			print(results.get_text())
	except:
		pass
	
	# Initialize loop structure - collect all article links, titles, article body, etc, save next button link; after collection, send browser to next button, clear list and initialize html soup
	years = [] # initialize year range
	for i in range(1899,2021):
		years.append(i)
	years.reverse() # reverse years - slight optimization, years more likely to be closer to present
	
	while True:
		for linkString in soup.find_all('a'):
			linkStrings.append(linkString.string)
		##########
		for link in soup.find_all('a'):
			try:
				if link.get('href')[-len('queryname=1'):]in queries: 
					#print(link.string)
					linkList.append([str(baseUrl+link.get('href')),link.string])
					count+=1
				elif link.string == 'Next':
					if oldNextButton==None and nextButton==None:
						nextButton=str(baseUrl+link.get('href')) # save next button
					else:
						oldNextButton=nextButton
						nextButton=str(baseUrl+link.get('href'))
			except:
				pass
		###########
		#column: 1: State Name, 2 State, 3 Year: 2014, 4 Newspaper, 6 Title, 7 Day of the week, 8 Month, 9 Date-number, 10 section, 11 record number, 12 article
		for link in linkList:
			br.open(link[0]) # load link from linkList in MechBrowser
			linkHtml = br.response().read() # read the response and save it in variable
			mechSoup = BeautifulSoup(linkHtml) # soupify html response
			############Article Body##############
			
			for articleBody in mechSoup.findAll('div', attrs={'class': 'mainText'}): #search for article body
				worksheet.write(row, col+10, articleBody.get_text())
					#print(articleBody.get_text()) # remove html tags from article body
				
				
			#articleBodyWrite(mechSoup) #write article body 
			############Article Title###############
			for title in mechSoup.findAll('h3', attrs={'class': 'docCite'}):
				try:
					worksheet.write(row, col+4, title.get_text())
					print(title.get_text())
				except:
					pass
			
			############Article Record Number###############
			#mixedSoup=mechSoup.findAll('span', attrs={'class': 'tagName'})
			mixedSoup2=mechSoup.findAll('div', attrs={'class': 'sourceInfo'})
			#for item in mixedSoup:
				#print(item.get_text())
			startIndex = "Record Number:"
			endIndex = "Copyright"
			startIndexNum = 0
			endIndexNum = 0
			for item in mixedSoup2:
				if startIndex in item.get_text() and endIndex in item.get_text():
					#print(item.get_text())
					startIndexNum = item.get_text().index(startIndex) # find index of record number text in string
					endIndexNum =item.get_text().index(endIndex)
					recordNumberFound = item.get_text()[startIndexNum+15:endIndexNum] # slice notation indexing of string containing record number
					worksheet.write(row, col+9, recordNumberFound)
					print("Record Number: "+recordNumberFound)
					break
			#section
			#for item in mixedSoup2: # not working right
			#	if "Section:" in item.get_text() and "Page:" in item.get_text() or "Section:" in item.get_text() and "Record Number:" in item.get_text():
			#		if "Page:"  not in item.get_text():
			#			startIndexNum = item.get_text().index("Section:")
			#			endIndexNum =item.get_text().index("Record Number:")
			#			sectionFound= item.get_text()[startIndexNum+15:endIndexNum]
			#			print(sectionFound)
			#		else:
			#			startIndexNum = item.get_text().index("Section:")
			#			endIndexNum =item.get_text().index("Page:")
			#			sectionFound= item.get_text()[startIndexNum+9:endIndexNum]
			#			print(sectionFound)
			
			############Article Publisher - Newspaper###############
			for newspaper in mechSoup.findAll('span', attrs={'class': 'pubName'}):
				try:
					worksheet.write(row, col+3, newspaper.get_text())
					print(newspaper.get_text())
				except:
					pass
					
			############Article Date###############
			savedDay=None
			savedMonth=None
			savedYear=None
			savedDate=None
			daysOfWeek = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
			months = ["January","February","March","April","May","June","July","August","September","October","November","December"]
			for date in mechSoup.findAll('h3', attrs={'class': 'docCite'}):
				for item in date.findNextSiblings():
					for day in daysOfWeek:
						if day in item.get_text():
							worksheet.write(row, col+5, day)
							savedDay=day
							#print(item.get_text())
							#try:
							#	if int(item.get_text().split()[-1]) in years:
							#		worksheet.write(row, col+2, int(item.get_text().split()[-1]))
							#		break
							#except:
							for year in years:
								if str(year) in item.get_text():
									#print(year)
									worksheet.write(row, col+2,year)
									savedYear= year
									break
							for thing in item.get_text().split():
								try:
									if int(thing[:-1])<=31 and int(thing[:-1])>=1:
										#print(thing[:-1])
										savedDate=int(thing[:-1])
										worksheet.write(row, col+7,savedDate)
										break
								except:
									pass
									
							break
									
							
							
			for date in mechSoup.findAll('h3', attrs={'class': 'docCite'}):
				for item in date.findNextSiblings():
					for month in months:
						if month in item.get_text():
							worksheet.write(row, col+6, month)
							savedMonth=month
							break
			try:
				print(savedMonth +" "+ str(savedDay) + " " + str(savedDate) + " " + str(savedYear))
			except:
				print("Unable to determine entire date")
				pass
			
			
			#Article Section
			#mixedSoup2=mechSoup.findAll('div', attrs={'class': 'sourceInfo'})
			#mixedSoup3=mechSoup.findAll('span', attrs={'class': 'tagName'})
			#for tagName in mixedSoup3:
			#	print(tagName)
			#	if tagName.get_text() == "Section: ":
			#		print(tagName.next_sibling())
			#for item in mixedSoup2:
				#print(item)
			#for sibling in mechSoup.next_siblings:
			#	if sibling.get_text()=="Section: ":
			#		print(sibling.next_sibling())
			#for tag in mechSoup.find_all("br"):
			#	print(tag.get_text())
				
				

							#print(item.get_text())
					#print(item.get_text()+"    Test")
			worksheet.write(row, col, state)
			worksheet.write(row, col+1, state)
			print("\n")
			row+=1
		print("Loading next page of results...")
		br.open(str(nextButton)) # click next button
		soup=BeautifulSoup(br.response().read()) # reconfigure html soup
		#time.sleep(1)
		try:
			for results in soup.findAll('div', attrs={'class': 'jump_results'}): # prints the number of results recovered
				print(results.get_text()+"\n")
		except:
			pass
		
		if 'Next' not in linkStrings:
			nextSwitch+=1
			if nextSwitch==2:
				print("End of results")
				workbook.close()
				return(1)
				

		linkList[:]=[] # empty lists after each loop
		linkStrings[:]=[]
		
def greeting():
	greeting ="""
***************************
NewsBank Article Collector
         v. 1.0
    John Solsma 2014
***************************
"""
	print(greeting+"\n")
	print(datetime.datetime.now())
	while True:
		go = raw_input("Press return to continue... ")
		if go != None:
			break

		

		
"""_______Implementation_______"""			
#greeting()
while True:
	try:
		greeting()
		browserLoop(mechBrowseNewsBank())
	except Exception,e:
		print(str(e))
		print("Uncaught exception in main program loop. Will attempt to reinitialize program in 10 seconds...")
		time.sleep(10)
		
