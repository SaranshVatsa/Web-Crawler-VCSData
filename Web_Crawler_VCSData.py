
#This will scrape all the company details off the results at vcsdata.com for the industry mentioned, page by page.
#Since the main results page of this website is loaded dynamically, the data isn't available in the page source.
#Selenium ais therefore used to launch the website in a browser where the page can render and all contents can load.
#Then all the loaded content is scraped off from there.

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import urllib.request
import xlwt
import math

#Defining the xlwt sheet that will be saved as .xls later

book = xlwt.Workbook(encoding="utf-8")

sheet1 = book.add_sheet("Sheet 1")

#Writes the heading for each column in the first row of the sheet
sheet1.write(0, 0, "Index")
sheet1.write(0, 1, "Company Name")
sheet1.write(0, 2, "Industry")
sheet1.write(0, 3, "Sub Industry")
sheet1.write(0, 4, "Company Type")
sheet1.write(0, 5, "Level Of Office")
sheet1.write(0, 6, "Address")
sheet1.write(0, 7, "City")
sheet1.write(0, 8, "State")
sheet1.write(0, 9, "Email")
sheet1.write(0, 10, "Website")
sheet1.write(0, 11, "Location")
sheet1.write(0, 12, "Contact Name")
sheet1.write(0, 13, "Contact Mobile")
sheet1.write(0, 14, "Contact Phone")

lineNumber=1   #Index number for the entries.

#Function to scrape data from a company's page (the link to company is available with every result listing on main results page).
#This function is called later in the main script.

def scrapeCompanyPage(link):

    mainPageLink = "http://www.vcsdata.com/" + str(link) 

    #Since this page is not dynamically loaded unlike the main results page, this can be scraped straight from page source.
    #Hence, selenium and chromedriver are not needed here. a urllib.request is enough to access the content.

    page=urllib.request.urlopen(mainPageLink).read()
    soup = BeautifulSoup(page, "html.parser")

    contentListing = soup.find("div", {"class": "contentBoxListing"}).find("div", {"class": "contentlisting"})

    contentBox = contentListing.find("div", {"class": "innertitlepackinnerpage"})

    address = contentBox.find("div", {"class": "span7"}).find("span", {"itemprop": "address"}).text

    if (address.split('\r\n')[-1][0:2] == "E:"):
        email = address.split('\r\n')[-1]

        address = address.split('\r\n')[0: len(address.split('\r\n'))-2 ]

        addressString = ""

        for line in address:
            addressString += line
            
        #print("\nEmail: "+ email[2:])
        sheet1.write(lineNumber, 9, email[2:])
        #print("\nAddress: "+ addressString)
        sheet1.write(lineNumber, 6, addressString)

    else:
        address = address.split('\r\n')

        addressString = ""

        for line in address:
            addressString += line
            
        #print("\nEmail: ")
        sheet1.write(lineNumber, 9, "")
        #print("\nAddress: "+ addressString)
        sheet1.write(lineNumber, 6, addressString)
        
    contentListingContent = contentListing.find("div", {"class": "content"}).find("div", {"class": "detailbox"}).find("div", {"class": "content"})

    try:
        city = contentListingContent.find("div", {"class": "list"}).find("div", {"class": "span7"})
        sheet1.write(lineNumber, 7, city.text.lstrip())
    except:
        print("\nError in city on page: "+ str(currentPageNumber) + " at Line: " + str(lineNumber) + "\n\n\n")

    try:
        state = contentListingContent.find("div", {"class": "list1"}).find("div", {"class": "span7"})
        sheet1.write(lineNumber, 8, state.text.lstrip())
    except:
        print("\nError in state on page: "+ str(currentPageNumber) + " at Line: " + str(lineNumber) + "\n\n\n")

    try:
        website = contentListingContent.findAll("div", {"class": "list"})[1].find("div", {"class": "span7"})
        sheet1.write(lineNumber, 10, website.text.lstrip())
    except:
        print("\nError in website on page: "+ str(currentPageNumber) + " at Line: " + str(lineNumber) + "\n\n\n")
    #print ("\nCity is: " + city.text.lstrip())
    #print ("\nState is: " + state.text.lstrip())
    #print ("\nWebsite is: " + website.text.lstrip())
    

    rowMarginData = contentBox.findAll("div", {"class": "margintop10"})

    try:
        subInd = rowMarginData[0].find("span", {"class": "pull_right"}).text[13:]

        #print ("Sub Industry is: " + subInd)
        sheet1.write(lineNumber, 3, subInd)
    except:
        print("\nError in Sub Ind on page: "+ str(currentPageNumber) + " at Line: " + str(lineNumber) + "\n\n\n")



    contentArr = soup.find("div", {"class": "TabbedPanelsContentGroup"}).findAll("div", {"class":"TabbedPanelsContent"})


    #repeaters1 = contentArr[1].find("div", {"class": "content"}).findAll("div",{"class":"repeater"})

    #for rep1 in repeaters1:
    #    print(rep1)

    try:
        repeaters2 = contentArr[2].find("div", {"class": "detailbox"}).find("div", {"class":"content"}).findAll("div", {"class":"listRow"})

        try:
            name = repeaters2[0].find("div", {"class": "span8"})
            sheet1.write(lineNumber, 12, name.text)
        except:
            print("\nError in Contact Name on page: "+ str(currentPageNumber) + " at Line: " + str(lineNumber) + "\n\n\n")


        try:
            mobile = repeaters2[1].find("div", {"class": "span8"}).text.split(" -")[0]
            sheet1.write(lineNumber, 13, mobile)
        except:
            print("\nError in Mobile on page: "+ str(currentPageNumber) + " at Line: " + str(lineNumber) + "\n\n\n")

        try:
            phone = repeaters2[2].find("div", {"class": "span8"}).text.split(" -")[0]
            sheet1.write(lineNumber, 14, phone)
        except:
            print("\nError in Phone on page: "+ str(currentPageNumber) + " at Line: " + str(lineNumber) + "\n\n\n")
    except:
            print("\nError in Contact Details on page: "+ str(currentPageNumber) + " at Line: " + str(lineNumber) + "\n\n\n")
            try:
                content = soup.find("div", {"class": "contentBoxListing"}).find("div", {"class":"contentlisting"}).find("div", {"class":"content"})
                phone = content.findAll("div", {"class":"detailbox"})[1].find("div", {"class":"content"}).find("div", {"class":"listRow"}).find("div", {"class":"span8"})
                sheet1.write(lineNumber, 14, phone.text)
            except:
                print("\nAGAIN Error in Contact Details on page: "+ str(currentPageNumber) + " at Line: " + str(lineNumber) + "\n\n\n")
    

    #print("\n\nCONTACT DETAILS: ")

    #print ("\nName: " + name.text)
    #print ("\nMobile: " + mobile)
    #print ("\nPhone: " + phone)

    
    
    

#THE MAIN SCRIPT


#Next line creates an instance of the webdriver for the preferred browser. I have used Chrome, but other broswers like Firebox can be used too.
#The path has to be the absolute path to the driver on your system.

driver = webdriver.Chrome('path/to/chromedriver.exe')  


#THIS IS THE LINK TO THE INDUSTRY SPECIFIC RESULTS PAGE TO SCRAPE. The URL should be changed to that of the industry whose data is needed.
pageLink = "http://www.vcsdata.com/callcentres.html?category=Call%20Centres"

pageLink += "&page="

#This will launch the browser set above and access the URL mentioned in pageLink
driver.get(pageLink+ "1")

#Waits for the page to load and the "pagination" class div to be rendered so that the totl number of pages in the results can be found.
WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CLASS_NAME, "pagination")))
innerHTML = driver.execute_script("return document.body.innerHTML")
soup = BeautifulSoup(innerHTML, "html.parser")

pagination = soup.find("div", {"class": "extrapadding" }).find("div", {"class": "pagination" })

if (pagination):
    allPagesLinks = pagination.findAll("a")

numberOfPages = allPagesLinks[len(allPagesLinks)-2].text   #The total number of pages in results.
    
print("\nNumber of pages: "+ numberOfPages + "\n\n\n")

prettysoup = soup.find("div", {"class": "contentBoxListing" }).findAll("div", {"class": "contentBoxListing" })

currentPageNumber = 1 #The counter for the page number being scraped.

#The following loop will run for all pages in the results one by one.
while(currentPageNumber <= int(numberOfPages)):

    try:

        #The current page is launched on the browser.
        driver.get(pageLink+ str(currentPageNumber))

        #Waits for the page to be rendered completely ("span75" is the class of one of the innermost divs on the results section).
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CLASS_NAME, "span75")))

        innerHTML = driver.execute_script("return document.body.innerHTML")
        soup = BeautifulSoup(innerHTML, "html.parser")

        #Scraping begings. One by one, each of the results unit (usually 8 on a page on this website) are accessed and their data extracted.

        prettysoup = soup.find("div", {"class": "contentBoxListing" }).findAll("div", {"class": "contentBoxListing" })

        for x in prettysoup:
            sheet1.write(lineNumber, 0, lineNumber)
            try:
                contentlisting = x.find("div", {"class": "contentlisting"})
                if (contentlisting):
                    name = contentlisting.find("div", {"class": "innertitle"})
                    linkToCompany = contentlisting.find("a", {"class": "link"}, href=True)
                if (name):
                    #print(name.text)
                    sheet1.write(lineNumber, 1, name.text)
                if (linkToCompany):
                    scrapeCompanyPage(linkToCompany['href'])
                content = contentlisting.find("div", {"class": "content"})
                if (content):
                    if (content.find("div", {"class": "list"})):
                        list0 = content.findAll("div", {"class": "list"})

                        ind = list0[0].find("div", {"class": "span75"})
                        loc = list0[1].find("div", {"class": "span75"})
                        if (ind):
                            #print(ind.text)
                            sheet1.write(lineNumber, 2, ind.text)

                        if (loc):
                            #print(loc.text)
                            sheet1.write(lineNumber, 11, loc.text)

                        list1 = content.findAll("div", {"class": "list1"})

                        compType = list1[0].findAll("div", {"class": "span5"})[1]
                        officeLevel = list1[1].findAll("div", {"class": "span5"})[1]
                        if (compType):
                            #print(compType.text)
                            sheet1.write(lineNumber, 4, compType.text)

                        if (officeLevel):
                            #print(officeLevel.text)
                            sheet1.write(lineNumber, 5, officeLevel.text)
            except:
                print("\nError on page: "+ str(currentPageNumber) + " at Line: " + str(lineNumber) + "\n\n\n")

            lineNumber+=1   #Line Number is incremented after every result, so every result goes on a new line with new index.

        print("\n\nPage number: "+ str(currentPageNumber) + " DONE!")   #A notification on console once a page is done.
        currentPageNumber+=1    #The page number is incremented so the next time loop runs the next page is accessed.
        book.save("absolute/path/to/result.xls")   #Saves to a .xls file after every page.
    except:
        print("\nError on page: "+ str(currentPageNumber))  #Usually comes up for empty results divs on the dynamically loaded page source.

driver.quit()   #Closes the browser and shuts off the driver once all pages are done.

print("\n\nCOMPLETE!!!")    #Notification on console after successfully scraping all the pages.
book.save("absolute/path/to/result.xls")   #A final save of the .xls file. 
