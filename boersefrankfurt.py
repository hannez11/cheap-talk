from requests_html import HTMLSession
import openpyxl
import time
import os
import urllib.request


class Parser:
    def __init__(self, filepath, new_filepath):
        self.filepath = filepath
        self.new_filepath = new_filepath


    def load_xlsx(self):
        print(f"File {self.filepath} opened") #set filename
        self.wb = openpyxl.load_workbook(self.filepath, data_only=True) #parses the actual values and not the formulas
        self.wb.create_sheet("Downloadlinks")
        self.ws = self.wb["Downloadlinks"]
        self.ws.cell(column = 1, row = self.ws.max_row).value = "Firmnumber" #create headers for new worksheet
        self.ws.cell(column = 2, row = self.ws.max_row).value = "ISIN"
        self.ws.cell(column = 3, row = self.ws.max_row).value = "Date"
        self.ws.cell(column = 4, row = self.ws.max_row).value = "Titel"
        self.ws.cell(column = 5, row = self.ws.max_row).value = "Link"


    def save_xlsx(self):
        self.wb.save(self.new_filepath)
        print(f"File {self.new_filepath} saved")


    def companies(self): #get pdf download links
        self.firmnumber = 1
        #loop through column 3
        self.ws = self.wb["file"] #set active sheet to grab url from #self.ws.max_row
        for col in self.ws.iter_cols(min_row=2, max_row=self.wb["file"].max_row, min_col=3, max_col=3): #iterate through isin column - adjust max_row
            for cell in col:
                self.url = cell.value
                print(f"Current URL of {self.firmnumber}: {self.url}")
                self.ws = self.wb["Downloadlinks"]
                self.parse_downloadlinks() #runs method to parse the url for a specific ISIN
                self.ws = self.wb["file"]
                self.save_xlsx()
            print("All done")
                

    def parse_downloadlinks(self): 
        time.sleep(2)
        session = HTMLSession() #has to be recalled for each page, otherwise closed session error
        r = session.get(self.url)
        r.html.render(timeout=0, sleep=1) #sleep needed for table (js/ajax) elements to load

        table = r.html.find("tr") #find all table rows (last table row has no content)

        for rows in range(1,len(table)): #parse date, doctitel and url for all rows
            try:
                current_row = self.ws.max_row + 1 #use the next empty row        
                
                self.ws.cell(column = 1, row = current_row).value = self.firmnumber
                
                date = r.html.find(f'tr.ng-star-inserted:nth-child({rows}) > td:nth-child(1)') #creates a list with 1 entry
                self.ws.cell(column = 3, row = current_row).value = date[0].text 
                #print(date[0].text)

                doctitel = r.html.find(f'tr.ng-star-inserted:nth-child({rows}) > td:nth-child(2)')
                self.ws.cell(column = 4, row = current_row).value = doctitel[0].text
                #print(doctitel[0].text)

                url = r.html.find(f'tr.ng-star-inserted:nth-child({rows}) > td:nth-child(3) > a:nth-child(1)')
                self.ws.cell(column = 5, row = current_row).value = list(url[0].absolute_links)[0]
                #print(list(url[0].absolute_links)[0]) #absolute_links creates a set -> transform to list to access link or [print(rows) for rows in about[0].absolute_links]

                self.ws.cell(column = 2, row = current_row).value = self.ws.cell(column = 5, row = current_row).value.split("=")[1][:-3] #isin column
            except:
                print("Error while going through the table rows")

        print(f"Firm {self.firmnumber} done \n")
        
        self.firmnumber += 1

    
    def downloader(self): #download pdf files
        folder = os.path.join(os.getcwd(), "Desktop", "Jahresabschluss")

        for i in range(1997, 2020): #create 23 folders
            if not os.path.exists(os.path.join(folder, str(i))):
                os.makedirs(os.path.join(folder, str(i)))

        self.wb = openpyxl.load_workbook(self.new_filepath, data_only=True)
        self.ws = self.wb["Jahresabschluss"]
        self.ws.cell(column = 9, row = 1).value = "Status"
        #max_row=self.wb["Jahresabschluss"].max_row min_row=2
        start, end = 290,300
        for col in self.ws.iter_cols(min_row=start, max_row=end, min_col=8, max_col=8): #iterate through link column
            for index, cell in enumerate(col):
                self.link = cell.value
                self.ws.cell(column = 9, row = start+index).value = "downloaded"

                year = str(self.ws.cell(column = 6, row = start+index).value)
                isin = str(self.ws.cell(column = 2, row = start+index).value)
                urllib.request.urlretrieve(self.link, os.path.join(folder, year, f"{isin}_{year}.pdf"))
                print(f"File {isin}_{year}.pdf downloaded\n{index+1} of {end-start+1} downloaded")

            self.save_xlsx()
            print("All done")


doc1 = Parser("C://Users//hannez//Desktop//CDAX.xlsx","C://Users//hannez//Desktop//CDAXoutput.xlsx")
doc1.downloader()
# doc1.load_xlsx()
# doc1.companies()
# doc1.save_xlsx()