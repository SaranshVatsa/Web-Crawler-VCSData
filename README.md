# Web-Crawler-VCSData

A Web Crawler to scrape all the results' data for the required industry off vcsdata.com

The data on the website results' page is dynamically loaded so you'll need to install Selenium and ChromeDriver (or any web browser driver you prefer).

And you need a scraping library, I have used BeautifulSoup4 so you can use that to avoid too much editing in the script.

You will need the xlwt package to create a spreadsheet that can be written to a .xls file.

You can download the Chrome Driver from: https://sites.google.com/a/chromium.org/chromedriver/downloads

Extract the zip file at your desired location on your system and provide the absolute path of the executable in the script where the driver instance is created.

To install the libraries, first ensure you have pip installed. Then:

	  $ pip install beautifulsoup4
	  $ pip install selenium
    	  $ pip install xlwt
    
Then you can just run the script giving the URL to the industry specific results page (where the comment to do so exists) and you will find your Excel sheet at the location you set for it to be written to (At the end of the script where the comment telling you to do so exists).


DO NOT close the browser, the chromeexecutable prompt or the shell window while the script is executing.

P.S The "error at: " console logs that you get from this at the console while the script is running is because the element the script was trying to access doesn't exist for that particular record. It simply isn't there on the site. So don't fret, nothing is broken :)

P.P.S I hope I have put enough comments throughout the script for you to understand how this works and enable you to edit it according to your need.
