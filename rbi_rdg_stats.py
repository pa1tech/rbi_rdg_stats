from selenium.webdriver import Chrome
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import re,io,requests
from pdfreader import SimplePDFViewer

driver = Chrome(service=Service(ChromeDriverManager().install()))
df = pd.DataFrame() 

driver.get("https://rbiretaildirect.org.in/#/about_statistics")
text = driver.page_source
driver.quit()

for stats in re.findall("RD Statistics [0-9]+.pdf",text):

	pdfUrl = "https://rbiretaildirect.org.in/stats/RD%20Statistics%20"+ stats[-12:-4]+".pdf"
	pdfBytes = io.BytesIO(requests.get(pdfUrl).content)

	viewer = SimplePDFViewer(pdfBytes)
	viewer.render()

	pdfStrings = viewer.canvas.strings

	words = []
	numbers = []
	word = ""
	for i in range(len(pdfStrings)):

		if(pdfStrings[i] == ' '):
			words.append(word)
			try:
				numbers.append(float( word.replace(',', '').replace('..', '.')  ))
			except:
				pass
			word = ""
		else:
			word = word + pdfStrings[i]

	date = words[8] + " " + words[9] + " " + words[10][:-1]
	df1 = pd.DataFrame({"Date":[date],
						"Total Accounts #":[words[21]],
						"T-bills Subscriptions (in ₹Cr)":[numbers[6]],
						"T-bills Holdings (in ₹Cr)":[numbers[16]],
						"SGB Holdings (in kg)":[numbers[17]]   })
	df = pd.concat((df,df1))



df['Date'] = pd.to_datetime(df['Date'])
df = df.sort_values(by='Date',ascending=True, ignore_index=True)
df['Date'] = df['Date'].dt.strftime('%d %b, %Y')
df.to_excel("RBI_RDG_Stats.xlsx",sheet_name="Sheet1",index = False)