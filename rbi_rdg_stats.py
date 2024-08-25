from selenium.webdriver import Chrome
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import re,io,requests,os
from pdfreader import SimplePDFViewer

chrome_install = ChromeDriverManager().install()

folder = os.path.dirname(chrome_install)
chromedriver_path = os.path.join(folder, "chromedriver.exe")

driver = Chrome(service=Service(chromedriver_path))

#driver = Chrome(service=Service(ChromeDriverManager().install()))
excel_file = "RBI_RDG_Stats.xlsx"

try:
	df = pd.read_excel(excel_file,sheet_name="Sheet1")
	existing = df["Date Code"].to_numpy(dtype=int)
except:
	existing = []
	df = pd.DataFrame()

driver.get("https://rbiretaildirect.org.in/#/about_statistics")
text = driver.page_source
driver.quit()

for stats in re.findall("RD Statistics [0-9]+.pdf",text):

	if (int(stats[-12:-4]) not in existing):
		print(stats)
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

		#date = words[8] + " " + words[9] + " " + words[10][:-1]
		date = stats[-12:-10] + "/" + stats[-10:-8] + "/" + stats[-8:-4]

		'''df1 = pd.DataFrame({"Date Code":[int(stats[-12:-4])],
							"Date":[date],
							"Total Accounts #":[words[21]],
							"T-bills Subscriptions (in ₹Cr)":[numbers[6]],
							"T-bills Holdings (in ₹Cr)":[numbers[16]],
							"SGB Holdings (in kg)":[numbers[17]]   })'''
		df1 = pd.DataFrame({"Date Code":[int(stats[-12:-4])],
							"Date":[date],
							"Total Accounts #":[words[22]],
							"T-bills Subscriptions (in ₹Cr)":[numbers[5]],
							"T-bills Holdings (in ₹Cr)":[numbers[15]],
							"SGB Holdings (in kg)":[numbers[16]]   })
		df = pd.concat((df,df1))



df['Date'] = pd.to_datetime(df['Date'],dayfirst=True)
df = df.sort_values(by='Date',ascending=True, ignore_index=True)
df['Date'] = df['Date'].dt.strftime('%d %b, %Y')
df.to_excel("RBI_RDG_Stats.xlsx",sheet_name="Sheet1",index = False)